"""Bank-specific transaction extraction and first-page identity hints."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
import re

import fitz
from PIL import Image
import pytesseract
from typing import Any
import pdfplumber
from src.transform.cheque_returns import is_cheque_return
from src.utils.ocr import find_tesseract, record_ocr_use
from src.utils.statement_utils import clean_cell, clean_detail, normalize_date, parse_amount


# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {
    "name": "HDFC Bank",
    "ifsc": "HDFC",
    "aliases": [
        "HDFC Bank"
    ],
    "unlabelled_left": True,
    # The logo slogan sits above the unlabelled customer/address block.
    "unlabelled_skip_patterns": (r"We\s*understand\s*your\s*world[.!]?",),
}


PDF_STATUS_PROFILE.update({'address_is_branch': True})

# Text and image-only HDFC layouts share this canonical bank module.


_OCR_ZOOM = 3
_DATE_RE = re.compile(r"\d{2}/\d{2}/\d{2}")
_MONEY_RE = re.compile(r"\d[\d,]*\.\d{2}")
_SUMMARY_RE = re.compile(
    r"^\s*([\d,]+\.\d{2})\s+(\d+)\s+(\d+)\s+"
    r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
)


@dataclass(frozen=True)
class _Word:
    text: str
    left: int
    top: int
    width: int
    height: int

    @property
    def right(self) -> int:
        return self.left + self.width


@dataclass
class _ScannedRow:
    page_index: int
    top: int
    height: int
    date: str
    detail_parts: list[str] = field(default_factory=list)
    cheque_parts: list[str] = field(default_factory=list)
    amount_text: str = ""
    balance_text: str = ""
    is_credit: bool = False


@dataclass(frozen=True)
class _StatementSummary:
    opening: Decimal
    debit_count: int
    credit_count: int
    debits: Decimal
    credits: Decimal
    closing: Decimal


def _tesseract_executable() -> str:
    executable = find_tesseract()
    if executable is not None:
        return executable
    raise RuntimeError(
        "Scanned HDFC statements require Tesseract OCR. Install Tesseract "
        "or set TESSERACT_CMD to its executable path."
    )


def _ocr_lines(image: Image.Image) -> list[list[_Word]]:
    record_ocr_use()
    data = pytesseract.image_to_data(
        image, config="--psm 6", output_type=pytesseract.Output.DICT
    )
    grouped: dict[tuple[int, int, int], list[_Word]] = defaultdict(list)
    for index, raw in enumerate(data["text"]):
        value = str(raw).strip()
        if value:
            key = (
                data["block_num"][index],
                data["par_num"][index],
                data["line_num"][index],
            )
            grouped[key].append(
                _Word(
                    value,
                    int(data["left"][index]),
                    int(data["top"][index]),
                    int(data["width"][index]),
                    int(data["height"][index]),
                )
            )
    return sorted(
        (sorted(words, key=lambda word: word.left) for words in grouped.values()),
        key=lambda words: (min(word.top for word in words), words[0].left),
    )


def _field(words: list[_Word]) -> str:
    return " ".join(
        part for word in words if (part := word.text.strip(" |\\[](){}"))
    ).strip()


def _money(raw: str) -> Decimal | None:
    compact = re.sub(r"\s+", "", raw).strip("|[](){}")
    if not _MONEY_RE.fullmatch(compact):
        return None
    try:
        return Decimal(compact.replace(",", ""))
    except InvalidOperation:
        return None


def _summary_from_line(words: list[_Word]) -> _StatementSummary | None:
    match = _SUMMARY_RE.fullmatch(" ".join(word.text for word in words))
    if not match:
        return None
    amounts = [_money(match.group(index)) for index in (1, 4, 5, 6)]
    if any(amount is None for amount in amounts):
        return None
    return _StatementSummary(
        amounts[0], int(match.group(2)), int(match.group(3)),
        amounts[1], amounts[2], amounts[3],
    )


def _column_words(words: list[_Word], width: int, start: float, end: float) -> list[_Word]:
    return [word for word in words if start <= word.left / width < end]


def _append_detail(row: _ScannedRow, words: list[_Word], width: int) -> None:
    detail = _field(_column_words(words, width, 0.085, 0.44))
    cheque = _field(_column_words(words, width, 0.44, 0.565))
    if detail:
        row.detail_parts.append(detail)
    if cheque:
        row.cheque_parts.append(cheque)


def _rows_on_page(
    lines: list[list[_Word]], page_index: int, width: int
) -> tuple[list[_ScannedRow], _StatementSummary | None]:
    rows: list[_ScannedRow] = []
    current: _ScannedRow | None = None
    summary: _StatementSummary | None = None
    in_footer = False
    for words in lines:
        text = " ".join(word.text for word in words)
        found_summary = _summary_from_line(words)
        if found_summary:
            summary = found_summary
        if "STATEMENT SUMMARY" in text.upper() or "HDFC BANK LIMITED" in text.upper():
            in_footer = True
        if in_footer:
            continue
        date_word = next(
            (
                word for word in words
                if word.left / width < 0.085
                and _DATE_RE.fullmatch(word.text.strip(" |\\[](){}"))
            ),
            None,
        )
        if date_word:
            amount_words = _column_words(words, width, 0.635, 0.89)
            balance_words = _column_words(words, width, 0.89, 1.01)
            current = _ScannedRow(
                page_index=page_index,
                top=min(word.top for word in words),
                height=(
                    max(word.top + word.height for word in words)
                    - min(word.top for word in words)
                ),
                date=date_word.text.strip(" |\\[](){}"),
                amount_text=_field(amount_words).replace(" ", ""),
                balance_text=_field(balance_words).replace(" ", ""),
                is_credit=bool(
                    amount_words
                    and max(word.right for word in amount_words) / width > 0.81
                ),
            )
            _append_detail(current, words, width)
            rows.append(current)
        elif current is not None:
            _append_detail(current, words, width)
    return rows, summary


def _reread_money(document: fitz.Document, row: _ScannedRow, kind: str) -> Decimal | None:
    pixmap = document[row.page_index].get_pixmap(
        matrix=fitz.Matrix(_OCR_ZOOM, _OCR_ZOOM), alpha=False
    )
    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    if kind == "balance":
        left, right = 0.89, 1.0
    elif row.is_credit:
        left, right = 0.76, 0.89
    else:
        left, right = 0.635, 0.76
    cell = image.crop(
        (
            int(image.width * left),
            max(0, row.top - 12),
            int(image.width * right),
            min(image.height, row.top + max(42, row.height + 20)),
        )
    )
    record_ocr_use()
    raw = pytesseract.image_to_string(
        cell,
        config="--psm 7 -c tessedit_char_whitelist=0123456789,.",
    )
    return _money(raw.strip())


def _balance_is_close(observed: Decimal | None, expected: Decimal) -> bool:
    return observed is not None and abs(observed - expected) <= Decimal("0.05")


def _partial_balance_matches(raw: str, expected: Decimal) -> bool:
    # OCR can blur or insert one character in the final cents at the right edge.
    # The complete rupee value must still match the running balance.
    compact = re.sub(r"\s+", "", raw).strip("|")
    match = re.fullmatch(r"([\d,]+)\.([\d:;?]{1,3})", compact)
    if not match:
        return False
    expected_integer, expected_cents = f"{expected:,.2f}".split(".")
    if match.group(1) != expected_integer:
        return False
    observed_cents = match.group(2)
    if len(observed_cents) == len(expected_cents):
        return sum(a != b for a, b in zip(observed_cents, expected_cents)) <= 1
    if len(observed_cents) == len(expected_cents) + 1:
        return any(
            observed_cents[:index] + observed_cents[index + 1:] == expected_cents
            for index in range(len(observed_cents))
        )
    if len(observed_cents) + 1 == len(expected_cents):
        return any(
            expected_cents[:index] + expected_cents[index + 1:] == observed_cents
            for index in range(len(expected_cents))
        )
    return False


def parse_scanned_hdfc(pdf_path: str, logger, progress_cb=None) -> list[dict[str, object]]:
    """Read an image-only HDFC statement and require full summary reconciliation."""
    pytesseract.pytesseract.tesseract_cmd = _tesseract_executable()
    scanned_rows: list[_ScannedRow] = []
    summary: _StatementSummary | None = None
    with fitz.open(pdf_path) as document:
        for page_index, page in enumerate(document):
            pixmap = page.get_pixmap(
                matrix=fitz.Matrix(_OCR_ZOOM, _OCR_ZOOM), alpha=False
            )
            image = Image.frombytes(
                "RGB", (pixmap.width, pixmap.height), pixmap.samples
            )
            page_rows, page_summary = _rows_on_page(
                _ocr_lines(image), page_index, pixmap.width
            )
            scanned_rows.extend(page_rows)
            if page_summary is not None:
                summary = page_summary
            logger.info(
                "HDFC OCR page %s/%s: %s transaction rows",
                page_index + 1, len(document), len(page_rows),
            )

        if summary is None:
            raise ValueError(
                "HDFC OCR could not read the statement summary; "
                "transaction totals cannot be verified."
            )
        expected_count = summary.debit_count + summary.credit_count
        if len(scanned_rows) != expected_count:
            raise ValueError(
                f"HDFC OCR found {len(scanned_rows)} transactions; "
                f"statement summary expects {expected_count}."
            )

        records: list[dict[str, object]] = []
        previous_balance = summary.opening
        debit_count = credit_count = 0
        total_debit = total_credit = Decimal("0.00")
        corrected_cents = 0
        for index, row in enumerate(scanned_rows, start=1):
            amount = _money(row.amount_text)
            if amount is None:
                amount = _reread_money(document, row, "amount")
            if amount is None:
                raise ValueError(
                    f"HDFC OCR could not read the amount on page "
                    f"{row.page_index + 1}, transaction {index}."
                )

            expected_balance = (
                previous_balance + amount if row.is_credit
                else previous_balance - amount
            )
            observed_balance = _money(row.balance_text)
            if not _balance_is_close(observed_balance, expected_balance):
                reread_balance = _reread_money(document, row, "balance")
                if reread_balance is not None:
                    observed_balance = reread_balance
            if not _balance_is_close(observed_balance, expected_balance):
                reread_amount = _reread_money(document, row, "amount")
                if reread_amount is not None:
                    candidate_balance = (
                        previous_balance + reread_amount if row.is_credit
                        else previous_balance - reread_amount
                    )
                    if (
                        observed_balance is not None
                        and abs(observed_balance - candidate_balance) <= Decimal("0.05")
                    ):
                        amount = reread_amount
                        expected_balance = candidate_balance
            if (
                not _balance_is_close(observed_balance, expected_balance)
                and _partial_balance_matches(row.balance_text, expected_balance)
            ):
                observed_balance = expected_balance
            if not _balance_is_close(observed_balance, expected_balance):
                raise ValueError(
                    f"HDFC OCR balance mismatch on page {row.page_index + 1}, "
                    f"transaction {index}: previous={previous_balance}, "
                    f"amount={amount}, printed={row.balance_text!r}."
                )
            if observed_balance != expected_balance:
                corrected_cents += 1

            details = re.sub(r"\s+", " ", " ".join(row.detail_parts)).strip()
            cheque = re.sub(r"\s+", " ", " ".join(row.cheque_parts)).strip()
            if not details:
                raise ValueError(
                    f"HDFC OCR found no narration on page "
                    f"{row.page_index + 1}, transaction {index}."
                )
            try:
                transaction_date = datetime.strptime(
                    row.date, "%d/%m/%y"
                ).strftime("%d/%m/%Y")
            except ValueError as exc:
                raise ValueError(
                    f"HDFC OCR could not read the date on page "
                    f"{row.page_index + 1}, transaction {index}."
                ) from exc

            debit = None if row.is_credit else float(amount)
            credit = float(amount) if row.is_credit else None
            if row.is_credit:
                credit_count += 1
                total_credit += amount
            else:
                debit_count += 1
                total_debit += amount
            records.append(
                {
                    "Sno": index,
                    "Date": transaction_date,
                    "Details": details,
                    "Detail_Clean": re.sub(r"[^A-Za-z0-9]", "", details),
                    "Cheque No": cheque,
                    "Debit": debit,
                    "Credit": credit,
                    "Balance": float(expected_balance),
                }
            )
            previous_balance = expected_balance
            if progress_cb is not None:
                progress_cb(index)

    if (
        debit_count != summary.debit_count
        or credit_count != summary.credit_count
        or total_debit != summary.debits
        or total_credit != summary.credits
        or previous_balance != summary.closing
    ):
        raise ValueError(
            "HDFC OCR totals do not match the statement summary: "
            f"debits {debit_count}/{total_debit} vs "
            f"{summary.debit_count}/{summary.debits}; "
            f"credits {credit_count}/{total_credit} vs "
            f"{summary.credit_count}/{summary.credits}; "
            f"closing {previous_balance} vs {summary.closing}."
        )
    logger.info(
        "HDFC OCR reconciled: rows=%s debits=%s credits=%s "
        "closing=%s corrected_balance_cents=%s",
        len(records), total_debit, total_credit,
        previous_balance, corrected_cents,
    )
    return records


TRANSACTION_LINE_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{2})\s+"
    r"(?P<body>.+?)\s+"
    r"(?P<value_date>\d{2}/\d{2}/\d{2})\s+"
    r"(?P<amount>[\-0-9,]+\.\d{2}|-)\s+"
    r"(?P<balance>[\-0-9,]+\.\d{2}|-)$"
)
RETURN_NOTICE_LINE_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{2})\s+(?P<body>.+?)\s+"
    r"(?P<value_date>\d{2}/\d{2}/\d{2})(?:\s+(?P<tail>.*))?$"
)
TABLE_HEADER_TEXT = "Date Narration Chq./Ref.No. ValueDt WithdrawalAmt. DepositAmt. ClosingBalance"
PAGE_HEADER_END_PREFIX = "StatementFrom :"
SUMMARY_LINE_RE = re.compile(
    r"^(?P<opening>[0-9,]+\.\d{2})\s+"
    r"(?P<dr_count>\d+)\s+"
    r"(?P<cr_count>\d+)\s+"
    r"(?P<debits>[0-9,]+\.\d{2})\s+"
    r"(?P<credits>[0-9,]+\.\d{2})\s+"
    r"(?P<closing>[0-9,]+\.\d{2})$"
)
FOOTER_PREFIXES = (
    "HDFCBANKLIMITED",
    "*Closingbalanceincludes",
    "Contentsofthisstatement",
    "thisstatement.",
    "StateaccountbranchGSTN:",
    "HDFCBankGSTINnumberdetails",
    "RegisteredOfficeAddress:",
    "GeneratedOn:",
    "Thisisacomputergeneratedstatementanddoes",
    "notrequiresignature.",
    "STATEMENTSUMMARY :-",
    "OpeningBalance DrCount CrCount Debits Credits ClosingBal",
)
DEBIT_HINTS = ("DEBIT", "DR", "NEFTDR", "CHQPAID", "ATW-", "ATM", "POS", "FEE")
CREDIT_HINTS = ("CREDIT", "CR", "NEFTCR", "CASHDEPOSIT", "SETTL")


@dataclass
class PendingRecord:
    date_text: str
    detail_head: str
    cheque_no: str
    amount_value: float | None
    balance_value: float | None
    continuation_lines: list[str] = field(default_factory=list)


def _should_skip_footer(line: str) -> bool:
    return line.startswith(FOOTER_PREFIXES)


def _split_body(body: str) -> tuple[str, str]:
    parts = clean_cell(body).rsplit(" ", 1)
    if len(parts) == 2 and len(parts[1]) >= 6 and any(ch.isdigit() for ch in parts[1]):
        return clean_cell(parts[0]), clean_cell(parts[1])
    return clean_cell(body), ""


def _classify_amount(
    details: str,
    amount_value: float,
    balance_value: float | None,
    previous_balance: float | None,
) -> tuple[float | None, float | None]:
    abs_amount = abs(amount_value)
    if previous_balance is not None and balance_value is not None:
        delta = round(balance_value - previous_balance, 2)
        if abs(abs(delta) - abs_amount) <= 0.1:
            if amount_value < 0:
                if delta >= 0:
                    return amount_value, None
                return None, amount_value
            if delta >= 0:
                return None, abs_amount
            return abs_amount, None

    upper_details = details.upper()
    if any(token in upper_details for token in CREDIT_HINTS) and not any(
        token in upper_details for token in DEBIT_HINTS
    ):
        return None, amount_value if amount_value < 0 else abs_amount
    if any(token in upper_details for token in DEBIT_HINTS):
        return amount_value if amount_value < 0 else abs_amount, None
    if previous_balance is not None and balance_value is not None and balance_value >= previous_balance:
        return None, amount_value if amount_value < 0 else abs_amount
    return amount_value if amount_value < 0 else abs_amount, None


def _finalize_record(
    pending: PendingRecord,
    previous_balance: float | None,
) -> tuple[dict[str, Any], float | None]:
    details = clean_cell(" ".join([pending.detail_head, *pending.continuation_lines]))
    debit, credit = (None, None) if pending.amount_value is None else _classify_amount(
        details=details, amount_value=pending.amount_value,
        balance_value=pending.balance_value, previous_balance=previous_balance,
    )
    next_balance = (
        previous_balance if pending.balance_value in (None, 0) and debit is None and credit is None
        else pending.balance_value if pending.balance_value is not None else previous_balance
    )
    return (
        {
            "Sno": 0,
            "Date": normalize_date(pending.date_text) or pending.date_text,
            "Details": details,
            "Detail_Clean": clean_detail(details),
            "Cheque No": pending.cheque_no,
            "Debit": debit,
            "Credit": credit,
            "Balance": pending.balance_value,
        },
        next_balance,
    )


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing HDFC statement: %s", pdf_path)

    # Image-only pages need OCR before the text parser touches the scan.
    with fitz.open(pdf_path) as document:
        if document.page_count and all(not page.get_text().strip() for page in document):
            logger.info("HDFC PDF has no text layer; using OCR")
            return parse_scanned_hdfc(pdf_path, logger, progress_cb)

    pending: PendingRecord | None = None
    raw_records: list[PendingRecord] = []
    opening_balance: float | None = None
    summary_counts: tuple[int, int] | None = None
    statement_started = False

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            lines = (page.extract_text() or "").splitlines()
            logger.debug("Page %s: extracted %s text line(s)", page_idx, len(lines))

            page_header_complete = False
            in_footer = False

            for raw_line in lines:
                line = clean_cell(raw_line)
                if not line:
                    continue

                summary_match = SUMMARY_LINE_RE.match(line)
                if summary_match:
                    opening_balance = parse_amount(summary_match.group("opening"))
                    summary_counts = (
                        int(summary_match.group("dr_count")),
                        int(summary_match.group("cr_count")),
                    )
                    continue

                if TABLE_HEADER_TEXT in line:
                    statement_started = True
                    page_header_complete = True
                    continue

                if not page_header_complete:
                    if line.startswith(PAGE_HEADER_END_PREFIX):
                        page_header_complete = True
                        continue
                    if statement_started and (TRANSACTION_LINE_RE.match(line) or RETURN_NOTICE_LINE_RE.match(line)):
                        page_header_complete = True
                    else:
                        continue

                if not statement_started:
                    continue

                if in_footer:
                    continue

                if _should_skip_footer(line):
                    in_footer = True
                    continue

                match = TRANSACTION_LINE_RE.match(line)
                if match:
                    if pending is not None:
                        raw_records.append(pending)
                        pending = None

                    detail_head, cheque_no = _split_body(match.group("body"))
                    amount_value = parse_amount(match.group("amount"))
                    balance_value = parse_amount(match.group("balance"))
                    if (amount_value is None or balance_value is None) and not is_cheque_return(detail_head, cheque_no):
                        continue

                    pending = PendingRecord(
                        date_text=match.group("date"),
                        detail_head=detail_head,
                        cheque_no=cheque_no,
                        amount_value=amount_value,
                        balance_value=balance_value,
                    )
                    continue

                notice_match = RETURN_NOTICE_LINE_RE.match(line)
                if notice_match and is_cheque_return(notice_match.group("body")):
                    if pending is not None:
                        raw_records.append(pending)
                        pending = None
                    detail_head, cheque_no = _split_body(notice_match.group("body"))
                    pending = PendingRecord(
                        date_text=notice_match.group("date"),
                        detail_head=clean_cell(f"{detail_head} {notice_match.group('tail') or ''}"),
                        cheque_no=cheque_no,
                        amount_value=None,
                        balance_value=None,
                    )
                    continue

                if pending is not None:
                    pending.continuation_lines.append(line)

    if pending is not None:
        raw_records.append(pending)

    records: list[dict[str, Any]] = []
    previous_balance = opening_balance

    for pending_record in raw_records:
        record, previous_balance = _finalize_record(pending_record, previous_balance)
        records.append(record)
        if progress_cb is not None:
            progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    if summary_counts is not None:
        expected_count = summary_counts[0] + summary_counts[1]
        if expected_count != len(records):
            logger.warning(
                "HDFC parsed row count mismatch: expected=%s actual=%s",
                expected_count,
                len(records),
            )
        else:
            logger.info("HDFC parsed row count matches summary: %s", len(records))

    logger.info("HDFC parse complete: rows=%s opening_balance=%s", len(records), opening_balance)
    if not records:
        logger.info("HDFC text parser found no rows; trying OCR")
        return parse_scanned_hdfc(pdf_path, logger, progress_cb)
    return records


BANK_CODE = 'hdfc'
BANK_SIGNATURES = (('HDFC BANK', 4), ('HDFC0', 3))
