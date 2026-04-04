from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import pdfplumber

from parser_helpers import build_record, parse_signed_balance
from utils import clean_cell, parse_amount

LINE_TOLERANCE = 3.0
DATE_RE = re.compile(r"^\d{2}[-/]\d{2}[-/]\d{2,4}$")
AMOUNT_RE = re.compile(r"^[0-9,]+\.\d{2}$")
BALANCE_RE = re.compile(r"^[0-9,]+\.\d{2}(?:D[Rr]|C[Rr])$")
TABLE_DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")
TABLE_SERIAL_LINE_RE = re.compile(
    r"^(?P<serial>\d{1,4})\s+"
    r"(?P<txn_date>\d{2}-\d{2}-\d{4})"
    r"(?:\s+(?P<value_date>\d{2}-\d{2}-\d{4}))?"
    r"(?:\s+(?P<tail>.*))?$"
)
FOOTER_PREFIXES = (
    "Contact-Us@18005700",
    "*This is computer-generated statement.",
    "Page ",
)
DATE_FORMATS = ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y")


@dataclass(frozen=True)
class WordLayout:
    details_min_left: float
    cheque_min_left: float
    withdrawal_left: float
    deposit_left: float

    @property
    def deposit_threshold(self) -> float:
        return (self.withdrawal_left + self.deposit_left) / 2.0


DEFAULT_WORD_LAYOUT = WordLayout(
    details_min_left=160.0,
    cheque_min_left=365.0,
    withdrawal_left=457.0,
    deposit_left=602.0,
)


@dataclass
class WordLine:
    top: float
    tokens: list[tuple[float, str]]


@dataclass
class PendingRecord:
    date_text: str
    amount_text: str
    amount_left: float
    balance_text: str
    detail_lines: list[str] = field(default_factory=list)
    cheque_no: str = ""


@dataclass
class TablePendingRecord:
    serial_no: int
    date_text: str
    value_date_text: str
    detail_lines: list[str] = field(default_factory=list)
    cheque_no: str = ""
    debit_text: str = ""
    credit_text: str = ""
    balance_text: str = ""


def _normalize_header_token(text: str) -> str:
    return re.sub(r"[^A-Z]+", "", clean_cell(text).upper())


def _match_header_key(text: str) -> str | None:
    if text == "DATE":
        return "date"
    if text in {"PARTICULARS", "NARRATION"}:
        return "details"
    if text.startswith("CHQ"):
        return "cheque"
    if "WITHDRAWAL" in text:
        return "withdrawal"
    if "DEPOSIT" in text:
        return "deposit"
    if text.startswith("BALANCE"):
        return "balance"
    return None


def _detect_word_layout(lines: list[WordLine]) -> WordLayout | None:
    required_headers = {"date", "details", "withdrawal", "deposit", "balance"}

    for line in lines:
        header_positions: dict[str, float] = {}
        for left, text in line.tokens:
            header_key = _match_header_key(_normalize_header_token(text))
            if header_key is None or header_key in header_positions:
                continue
            header_positions[header_key] = left

        if not required_headers.issubset(header_positions):
            continue

        return WordLayout(
            details_min_left=header_positions["details"],
            cheque_min_left=header_positions.get("cheque", header_positions["withdrawal"]),
            withdrawal_left=header_positions["withdrawal"],
            deposit_left=header_positions["deposit"],
        )

    return None


def _extract_lines(page: pdfplumber.page.Page) -> list[WordLine]:
    words = page.extract_words(
        x_tolerance=2,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )
    lines: list[WordLine] = []

    for word in sorted(words or [], key=lambda item: (float(item["top"]), float(item["x0"]))):
        text = clean_cell(word.get("text"))
        if not text:
            continue

        top = float(word["top"])
        left = float(word["x0"])
        if not lines or abs(lines[-1].top - top) > LINE_TOLERANCE:
            lines.append(WordLine(top=top, tokens=[(left, text)]))
        else:
            lines[-1].tokens.append((left, text))

    for line in lines:
        line.tokens.sort(key=lambda item: item[0])
    return lines


def _line_text(line: WordLine) -> str:
    return clean_cell(" ".join(text for _, text in line.tokens))


def _is_row_start(line: WordLine) -> bool:
    date_count = sum(1 for _, text in line.tokens if DATE_RE.match(text))
    has_balance = any(BALANCE_RE.match(text) for _, text in line.tokens)
    return date_count >= 1 and has_balance


def _is_footer_line(text: str) -> bool:
    if any(prefix in text for prefix in FOOTER_PREFIXES):
        return True

    if set(text) == {"-"}:
        return True

    normalized = re.sub(r"[^a-z0-9@]+", "", text.lower())
    if any(
        marker in normalized
        for marker in (
            "contactus@18005700",
            "thisiscomputergeneratedstatement",
            "thisisacomputergeneratedstatement",
            "signatureisrequired",
            "pagetotal",
            "grandtotal",
            "notechequesreceivedininwardclearing",
            "returningonthebasisopeningbalanceinaccount",
            "unlesstheconstituentnotifiesthebankofanydiscrepancyinthisstatement",
            "bankofbarodadate",
            "abbreviationsused",
            "pendingpenalcharges",
            "endofstatement",
        )
    ):
        return True

    if normalized.startswith("acnumber") and "accountopendate" in normalized:
        return True
    if normalized.startswith("statementofaccountfortheperiodof"):
        return True
    if normalized.startswith("dateparticulars") and "balance" in normalized:
        return True
    if normalized.startswith("address") and "street" in normalized:
        return True
    if normalized.startswith("helplineno") and "cybercrimehelpline" in normalized:
        return True
    if normalized.startswith("branchphoneno"):
        return True
    if normalized.startswith("micrcode") and "ifsccode" in normalized:
        return True
    if normalized.startswith("tiruchirapalli") and "time" in normalized:
        return True

    return False


def _is_continuation_line(line: WordLine, layout: WordLayout) -> bool:
    text = _line_text(line)
    if not text or _is_footer_line(text) or _is_row_start(line):
        return False

    content_positions = [
        left
        for left, token in line.tokens
        if not DATE_RE.match(token) and not AMOUNT_RE.match(token) and not BALANCE_RE.match(token)
    ]
    if not content_positions:
        return False

    return min(content_positions) + LINE_TOLERANCE >= layout.details_min_left


def _build_pending(block: list[WordLine], layout: WordLayout) -> PendingRecord | None:
    if not block:
        return None

    first_line = block[0]
    date_tokens = [text for _, text in first_line.tokens if DATE_RE.match(text)]
    if not date_tokens:
        return None

    balance_candidates: list[tuple[float, float, str]] = []
    amount_candidates: list[tuple[float, float, str]] = []
    cheque_parts: list[str] = []

    for line in block:
        raw_line_text = _line_text(line)
        if _is_footer_line(raw_line_text):
            continue

        for left, text in line.tokens:
            if BALANCE_RE.match(text):
                balance_candidates.append((line.top, left, text))
                continue
            if AMOUNT_RE.match(text):
                amount_candidates.append((line.top, left, text))
                continue
            if DATE_RE.match(text):
                continue
            if left < layout.details_min_left:
                continue
            if layout.cheque_min_left <= left < layout.withdrawal_left:
                cheque_parts.append(text)
                continue

    if not balance_candidates or not amount_candidates:
        return None

    amount_top, amount_left, amount_text = max(amount_candidates, key=lambda item: (item[0], item[1]))
    _, _, balance_text = max(balance_candidates, key=lambda item: (item[0], item[1]))

    filtered_detail_lines: list[str] = []
    for line in block:
        raw_line_text = _line_text(line)
        if _is_footer_line(raw_line_text):
            continue

        line_parts: list[str] = []
        for left, text in line.tokens:
            if DATE_RE.match(text) or BALANCE_RE.match(text):
                continue
            if AMOUNT_RE.match(text) and left == amount_left and abs(line.top - amount_top) <= LINE_TOLERANCE:
                continue
            if left < layout.details_min_left:
                continue
            if layout.cheque_min_left <= left < layout.withdrawal_left:
                continue
            if left >= amount_left and abs(line.top - amount_top) <= LINE_TOLERANCE:
                continue
            line_parts.append(text)

        detail_text = clean_cell(" ".join(line_parts))
        if detail_text and not _is_footer_line(detail_text):
            filtered_detail_lines.append(detail_text)

    return PendingRecord(
        date_text=date_tokens[0],
        amount_text=amount_text,
        amount_left=amount_left,
        balance_text=balance_text,
        detail_lines=filtered_detail_lines,
        cheque_no=clean_cell(" ".join(cheque_parts)),
    )


def _finalize_record(pending: PendingRecord, layout: WordLayout) -> dict[str, Any]:
    amount_value = parse_amount(pending.amount_text)
    debit: float | None = None
    credit: float | None = None
    if amount_value is not None:
        if pending.amount_left >= layout.deposit_threshold:
            credit = amount_value
        else:
            debit = amount_value

    return build_record(
        date_text=pending.date_text,
        details=" ".join(pending.detail_lines),
        cheque_no=pending.cheque_no,
        debit=debit,
        credit=credit,
        balance=parse_signed_balance(pending.balance_text),
        date_formats=DATE_FORMATS,
    )


def _parse_table_row(row: list[Any]) -> TablePendingRecord | None:
    if not row:
        return None

    first_cell = "" if row[0] is None else str(row[0])
    lines = [clean_cell(part) for part in first_cell.splitlines() if clean_cell(part)]
    if not lines:
        return None

    serial_index = -1
    serial_match: re.Match[str] | None = None
    for index, line in enumerate(lines):
        match = TABLE_SERIAL_LINE_RE.match(line)
        if match is None:
            continue
        serial_index = index
        serial_match = match
        break

    if serial_match is None:
        return None

    detail_lines = [line for line in lines[:serial_index] if line]
    tail = clean_cell(serial_match.group("tail") or "")
    if tail and tail != "-":
        detail_lines.append(tail)
    detail_lines.extend(line for line in lines[serial_index + 1 :] if line)

    cheque_no = clean_cell(row[1]) if len(row) > 1 else ""
    if cheque_no == "-":
        cheque_no = ""

    debit_text = clean_cell(row[2]) if len(row) > 2 else ""
    credit_text = clean_cell(row[3]) if len(row) > 3 else ""
    balance_text = clean_cell(row[4]) if len(row) > 4 else ""

    if not detail_lines and not balance_text:
        return None

    return TablePendingRecord(
        serial_no=int(serial_match.group("serial")),
        date_text=serial_match.group("txn_date"),
        value_date_text=serial_match.group("value_date") or serial_match.group("txn_date"),
        detail_lines=detail_lines,
        cheque_no=cheque_no,
        debit_text=debit_text,
        credit_text=credit_text,
        balance_text=balance_text,
    )


def _finalize_table_record(pending: TablePendingRecord) -> dict[str, Any]:
    debit_value = parse_amount(pending.debit_text)
    credit_value = parse_amount(pending.credit_text)

    return build_record(
        date_text=pending.date_text,
        details=" ".join(pending.detail_lines),
        cheque_no=pending.cheque_no,
        debit=debit_value,
        credit=credit_value,
        balance=parse_signed_balance(pending.balance_text),
        date_formats=DATE_FORMATS,
    )


def _parse_table_layout(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Bank of Baroda statement with table extraction: %s", pdf_path)

    by_serial: dict[int, dict[str, Any]] = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("Page %s: extracted %s table(s)", page_idx, len(tables))

            for table_idx, table in enumerate(tables, start=1):
                logger.debug("Page %s table %s: rows=%s", page_idx, table_idx, len(table))
                for row in table:
                    pending = _parse_table_row(row)
                    if pending is None:
                        continue
                    if pending.serial_no in by_serial:
                        continue

                    by_serial[pending.serial_no] = _finalize_table_record(pending)
                    if progress_cb is not None:
                        progress_cb(len(by_serial))

    records: list[dict[str, Any]] = []
    for index, serial_no in enumerate(sorted(by_serial), start=1):
        record = by_serial[serial_no]
        record["Sno"] = index
        records.append(record)

    logger.info("Bank of Baroda table parse complete: rows=%s", len(records))
    return records


def _parse_word_layout(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Bank of Baroda statement with word extraction fallback: %s", pdf_path)

    records: list[dict[str, Any]] = []
    active_layout = DEFAULT_WORD_LAYOUT
    pending_block: list[WordLine] = []
    pending_layout = DEFAULT_WORD_LAYOUT

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            lines = _extract_lines(page)
            detected_layout = _detect_word_layout(lines)
            if detected_layout is not None:
                active_layout = detected_layout
            layout = active_layout

            logger.debug("Page %s: extracted %s grouped line(s)", page_idx, len(lines))
            logger.debug(
                "Page %s: word layout details=%s cheque=%s withdrawal=%s deposit=%s",
                page_idx,
                round(layout.details_min_left, 1),
                round(layout.cheque_min_left, 1),
                round(layout.withdrawal_left, 1),
                round(layout.deposit_left, 1),
            )

            start_indexes = [index for index, line in enumerate(lines) if _is_row_start(line)]
            logger.debug("Page %s: identified %s row block(s)", page_idx, len(start_indexes))

            if pending_block:
                leading_end = start_indexes[0] if start_indexes else len(lines)
                continuation_lines = [
                    line for line in lines[:leading_end] if _is_continuation_line(line, pending_layout)
                ]
                pending_block.extend(continuation_lines)
                pending = _build_pending(pending_block, pending_layout)
                if pending is not None:
                    record = _finalize_record(pending, pending_layout)
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))
                pending_block = []

            if not start_indexes:
                continue

            for block_index, start_index in enumerate(start_indexes):
                end_index = start_indexes[block_index + 1] if block_index + 1 < len(start_indexes) else len(lines)
                block = lines[start_index:end_index]

                if block_index + 1 < len(start_indexes):
                    pending = _build_pending(block, layout)
                    if pending is None:
                        continue

                    record = _finalize_record(pending, layout)
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))
                    continue

                pending_block = list(block)
                pending_layout = layout

    if pending_block:
        pending = _build_pending(pending_block, pending_layout)
        if pending is not None:
            record = _finalize_record(pending, pending_layout)
            records.append(record)
            if progress_cb is not None:
                progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("Bank of Baroda word parse complete: rows=%s", len(records))
    return records


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    table_records = _parse_table_layout(pdf_path, logger, progress_cb)
    if table_records:
        return table_records

    logger.info("Bank of Baroda table parser found no rows. Falling back to word parser.")
    records = _parse_word_layout(pdf_path, logger, progress_cb)
    logger.info("Bank of Baroda parse complete: rows=%s", len(records))
    return records
