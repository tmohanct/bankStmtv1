"""Bank-specific transaction extraction and first-page identity hints."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
import fitz
from src.utils.parser_helpers import build_record, parse_signed_balance
from src.utils.statement_utils import clean_cell, parse_amount
from src.transform.normalize import normalize_cheque_number, extract_cheque_number_from_details
from src.transform.cheque_returns import is_cheque_return


# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {
    "name": "Indian Bank",
    "ifsc": "IDIB",
    "aliases": [
        "Indian Bank"
    ],
    "name_after": "STATEMENT OF ACCOUNT",
    "period_from_label": "Statement From",
    "period_to_label": "To",
}


DATE_RE = re.compile(r"^(?:\d{2} [A-Za-z]{3,9} \d{4}|[A-Za-z]{3,9} \d{2} \d{4})$")
DATE_FORMATS = ("%d %b %Y", "%d %B %Y", "%b %d %Y", "%B %d %Y")
POST_DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{2}$")
POST_DATE_FORMATS = ("%d/%m/%y",)
POST_DATE_SUMMARY_RE = re.compile(
    r"Dr\.\s*Count\s*:\s*(\d+)\s+Cr\.\s*Count\s*:\s*(\d+)\s+"
    r"([0-9,]+\.\d{2})\s+([0-9,]+\.\d{2})",
    re.IGNORECASE,
)
HEADER_LINES = {"Date", "Transaction Details", "Debits", "Credits", "Balance"}
DETAIL_SKIP_PREFIXES = ("page ",)
DETAIL_SKIP_LINES = {"Ending Balance", "Total"}
DETAIL_SKIP_PAGE_RE = re.compile(r"^\d+/\d+$")
AMOUNT_TEXT_RE = re.compile(
    r"^(?:-|[+\-]?\s*(?:INR|RS\.?)?\s*[0-9,]+\.\d{2}\s*(?:CR|DR)?\.?)$",
    re.IGNORECASE,
)


@dataclass
class PendingRecord:
    date_text: str
    lines: list[str] = field(default_factory=list)


def _is_amount_or_dash(line: str) -> bool:
    return bool(AMOUNT_TEXT_RE.match(line))


def _parse_transaction_amount(line: str) -> float | None:
    value = parse_amount(line)
    if value is None:
        return None
    return abs(value)


def _has_balance_suffix(line: str) -> bool:
    return bool(re.search(r"(?:CR|DR)\.?$", line.strip(), re.IGNORECASE))


def _parse_debit_credit_marker(line: str) -> float | None:
    if _has_balance_suffix(line):
        return None
    return _parse_transaction_amount(line)


def _is_post_balance_footer(line: str) -> bool:
    if not line:
        return True

    lowered = line.lower()
    if lowered.startswith("cr inr") or lowered.startswith("dr inr"):
        return True
    if lowered.startswith("rupees "):
        return True
    if any(token in line for token in ("NEFT:", "UPI:", "RTGS:", "BBPS:", "IMPS:")):
        return True
    return False


def _should_skip_detail_line(line: str) -> bool:
    if not line:
        return True
    if line in HEADER_LINES or line in DETAIL_SKIP_LINES:
        return True
    if DETAIL_SKIP_PAGE_RE.match(line):
        return True

    lowered = line.lower()
    if any(lowered.startswith(prefix) for prefix in DETAIL_SKIP_PREFIXES):
        return True
    if line.startswith("Indian Bank |"):
        return True
    return False


def _find_footer_start_index(lines: list[str]) -> int:
    for idx, line in enumerate(lines):
        if line in DETAIL_SKIP_LINES or _is_post_balance_footer(line):
            return idx
    return len(lines)


def _select_row_amount_markers(lines: list[str]) -> list[tuple[int, str]]:
    footer_start_idx = _find_footer_start_index(lines)
    amount_markers = [
        (idx, line)
        for idx, line in enumerate(lines[:footer_start_idx])
        if _is_amount_or_dash(line)
    ]

    for marker_pos in range(len(amount_markers) - 1, -1, -1):
        balance_idx, balance_text = amount_markers[marker_pos]
        if balance_text == "-" or parse_signed_balance(balance_text) is None:
            continue

        if marker_pos >= 2:
            debit_marker = amount_markers[marker_pos - 2]
            credit_marker = amount_markers[marker_pos - 1]
            debit_text = debit_marker[1]
            credit_text = credit_marker[1]
            debit_value = _parse_debit_credit_marker(debit_text) if debit_text != "-" else None
            credit_value = _parse_debit_credit_marker(credit_text) if credit_text != "-" else None
            if (debit_value is not None) != (credit_value is not None):
                return [debit_marker, credit_marker, (balance_idx, balance_text)]

        if marker_pos >= 1:
            amount_marker = amount_markers[marker_pos - 1]
            amount_text = amount_marker[1]
            if amount_text != "-" and _parse_debit_credit_marker(amount_text) is not None:
                return [amount_marker, (balance_idx, balance_text)]

    return []


def _classify_amount_from_balance(
    amount_value: float,
    balance_value: float,
    previous_balance: float | None,
) -> tuple[float | None, float | None]:
    abs_amount = abs(amount_value)
    if previous_balance is not None:
        delta = round(balance_value - previous_balance, 2)
        if abs(abs(delta) - abs_amount) <= 0.1:
            if delta >= 0:
                return None, abs_amount
            return abs_amount, None
    return abs_amount, None


def _finalize_record(
    pending: PendingRecord,
    previous_balance: float | None,
) -> tuple[dict[str, Any] | None, float | None]:
    notice = _build_return_notice(
        pending.date_text, pending.lines[:_find_footer_start_index(pending.lines)],
        previous_balance, DATE_FORMATS,
    )
    if notice is not None:
        return notice, previous_balance
    row_markers = _select_row_amount_markers(pending.lines)
    if len(row_markers) < 2:
        if is_cheque_return(" ".join(pending.lines)):
            raise ValueError(
                f"Indian Bank {pending.date_text}: cheque return/rejection amounts could not be read; "
                f"refusing to omit the entry: {pending.lines!r}"
            )
        return None, previous_balance

    row_marker_indexes = {idx for idx, _ in row_markers}
    footer_start_idx = _find_footer_start_index(pending.lines)
    detail_parts = [
        line
        for idx, line in enumerate(pending.lines)
        if idx < footer_start_idx
        and idx not in row_marker_indexes
        and not _is_amount_or_dash(line)
        and not _should_skip_detail_line(line)
        and not _is_post_balance_footer(line)
    ]
    details = clean_cell(" ".join(detail_parts))

    debit: float | None = None
    credit: float | None = None

    if len(row_markers) == 3:
        debit_text = row_markers[0][1]
        credit_text = row_markers[1][1]
        balance_text = row_markers[2][1]
        debit = _parse_debit_credit_marker(debit_text) if debit_text != "-" else None
        credit = _parse_debit_credit_marker(credit_text) if credit_text != "-" else None
        balance_value = parse_signed_balance(balance_text)
    else:
        amount_text = row_markers[0][1]
        balance_text = row_markers[1][1]
        amount_value = _parse_debit_credit_marker(amount_text)
        balance_value = parse_signed_balance(balance_text)
        if amount_value is None or balance_value is None:
            return None, previous_balance
        debit, credit = _classify_amount_from_balance(
            amount_value=amount_value,
            balance_value=balance_value,
            previous_balance=previous_balance,
        )

    if balance_value is None:
        return None, previous_balance

    record = build_record(
        date_text=pending.date_text,
        details=details,
        debit=debit,
        credit=credit,
        balance=balance_value,
        date_formats=DATE_FORMATS,
    )
    return record, balance_value


def _finalize_post_date_record(
    lines: list[str], previous_balance: float, page_number: int,
) -> tuple[dict[str, Any], float]:
    """Read one row from the Post Date / Value Date statement layout."""
    if len(lines) >= 3 and POST_DATE_RE.fullmatch(lines[0]) and POST_DATE_RE.fullmatch(lines[1]):
        notice = _build_return_notice(lines[0], lines[2:], previous_balance, POST_DATE_FORMATS)
        if notice is not None:
            return notice, previous_balance
    if (len(lines) < 5 or not POST_DATE_RE.fullmatch(lines[0])
            or not POST_DATE_RE.fullmatch(lines[1])
            or not _is_amount_or_dash(lines[-2])
            or not _has_balance_suffix(lines[-1])):
        raise ValueError(f"Indian Bank page {page_number}: unrecognized transaction row: {lines!r}")

    amount = _parse_transaction_amount(lines[-2])
    balance = parse_signed_balance(lines[-1])
    if amount is None or balance is None:
        raise ValueError(f"Indian Bank page {page_number}: invalid transaction amount or balance")
    change = round(balance - previous_balance, 2)
    if abs(abs(change) - amount) > 0.01:
        raise ValueError(
            f"Indian Bank page {page_number}, {lines[0]}: amount {amount:.2f} "
            f"does not match balance change {change:.2f}"
        )

    detail_lines = lines[2:-2]
    cheque_no = ""
    if (detail_lines and re.fullmatch(r"\d{3,12}", detail_lines[-1])
            and normalize_cheque_number(detail_lines[-1], " ".join(detail_lines[:-1]))):
        cheque_no = detail_lines.pop()
    record = build_record(
        date_text=lines[0],
        details=" ".join(detail_lines),
        cheque_no=cheque_no,
        debit=amount if change < 0 else None,
        credit=amount if change > 0 else None,
        balance=balance,
        date_formats=POST_DATE_FORMATS,
    )
    return record, balance


def _build_return_notice(date_text, lines, previous_balance, date_formats):
    """Keep return/rejection notices even with blank or zero amount cells."""
    detail_lines = list(lines)
    cells = []
    while detail_lines and _is_amount_or_dash(detail_lines[-1]):
        cells.insert(0, detail_lines.pop())
    details = clean_cell(" ".join(detail_lines))
    cheque_no = extract_cheque_number_from_details(details)
    if not is_cheque_return(details, cheque_no):
        return None
    balance = None
    if cells and cells[-1] != "-":
        balance = parse_signed_balance(cells[-1])
        if balance not in (None, 0) and (
            not _has_balance_suffix(cells[-1]) or previous_balance is None
            or abs(balance - previous_balance) > 0.01
        ):
            return None
        cells = cells[:-1]
    if any(cell != "-" and parse_signed_balance(cell) != 0 for cell in cells):
        return None
    return build_record(
        date_text=date_text, details=details, cheque_no=cheque_no,
        debit=0.0, credit=0.0, balance=balance, date_formats=date_formats,
    )


def _post_date_page_rows(lines: list[str], start: int, end: int, page_number: int) -> list[list[str]]:
    """Split the transaction section into rows without interpreting ledger amounts."""
    page_rows: list[list[str]] = []
    pending: list[str] = []
    for line in lines[start + 2:end]:
        if POST_DATE_RE.fullmatch(line):
            if len(pending) == 1:
                pending.append(line)  # Value Date follows Post Date.
                continue
            if pending:
                page_rows.append(pending)
            pending = [line]
        elif pending:
            pending.append(line)
        else:
            raise ValueError(f"Indian Bank page {page_number}: unexpected text before first row")
    if pending:
        page_rows.append(pending)
    return page_rows


def _parse_post_date_layout(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    previous_balance: float | None = None

    with fitz.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            lines = [clean_cell(line) for line in (page.get_text("text") or "").splitlines()]
            lines = [line for line in lines if line]
            try:
                start = lines.index("Brought Forward")
                end = next(index for index in range(start + 1, len(lines))
                           if lines[index] in ("Carried Forward", "CLOSING BALANCE :"))
            except (ValueError, StopIteration) as exc:
                raise ValueError(f"Indian Bank page {page_number}: missing transaction section") from exc

            opening_balance = parse_signed_balance(lines[start + 1])
            closing_balance = parse_signed_balance(lines[end + 1])
            if opening_balance is None or closing_balance is None:
                raise ValueError(f"Indian Bank page {page_number}: invalid page balance")
            if previous_balance is not None and abs(opening_balance - previous_balance) > 0.01:
                raise ValueError(f"Indian Bank page {page_number}: brought forward balance mismatch")
            previous_balance = opening_balance

            page_rows = _post_date_page_rows(lines, start, end, page_number)

            # Interest entries can be printed after a later day's transfer.
            # Stable sorting preserves the printed order within each date;
            # every transaction and page balance must still reconcile below.
            page_rows.sort(key=lambda row: datetime.strptime(row[0], POST_DATE_FORMATS[0]))
            for row in page_rows:
                record, previous_balance = _finalize_post_date_record(
                    row, previous_balance, page_number,
                )
                records.append(record)
                if progress_cb is not None:
                    progress_cb(len(records))
            if abs(previous_balance - closing_balance) > 0.01:
                raise ValueError(f"Indian Bank page {page_number}: carried forward balance mismatch")

    for index, record in enumerate(records, start=1):
        record["Sno"] = index
    logger.info("Indian Bank post-date parse complete: rows=%s closing_balance=%s",
                len(records), previous_balance)
    return records


def extract_summary_metrics(pdf_path: str, logger) -> dict[str, float]:
    """Read the cumulative totals on this Indian Bank layout's final page."""
    with fitz.open(pdf_path) as pdf:
        if not pdf:
            return {}
        text = pdf[-1].get_text("text") or ""
    if "CLOSING BALANCE :" not in text:
        return {}
    match = POST_DATE_SUMMARY_RE.search(text)
    if match is None:
        logger.warning("Indian Bank final-page summary was not recognized")
        return {}
    debit = parse_amount(match.group(3))
    credit = parse_amount(match.group(4))
    if debit is None or credit is None:
        return {}
    return {
        "transaction_count": float(int(match.group(1)) + int(match.group(2))),
        "total_debit": debit,
        "total_credit": credit,
    }


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Indian Bank statement: %s", pdf_path)

    with fitz.open(pdf_path) as pdf:
        first_page = pdf[0].get_text("text") if pdf else ""
    if "Post Date" in first_page and "Brought Forward" in first_page:
        return _parse_post_date_layout(pdf_path, logger, progress_cb)

    records: list[dict[str, Any]] = []
    pending: PendingRecord | None = None
    in_activity = False
    expect_opening_balance = False
    previous_balance: float | None = None

    with fitz.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            lines = [clean_cell(line) for line in (page.get_text("text") or "").splitlines()]
            logger.debug("Page %s: extracted %s text line(s)", page_idx, len(lines))

            for line in lines:
                if not line:
                    continue

                if expect_opening_balance:
                    opening_balance = parse_signed_balance(line)
                    expect_opening_balance = False
                    if opening_balance is not None:
                        previous_balance = opening_balance
                    continue

                if line == "Opening Balance":
                    expect_opening_balance = True
                    continue

                if line == "ACCOUNT ACTIVITY":
                    in_activity = True
                    continue

                if not in_activity:
                    continue

                if line in HEADER_LINES:
                    continue

                if DATE_RE.match(line):
                    if pending is not None:
                        record, previous_balance = _finalize_record(pending, previous_balance)
                        if record is not None:
                            records.append(record)
                            if progress_cb is not None:
                                progress_cb(len(records))
                        else:
                            logger.warning(
                                "Skipped Indian Bank row with insufficient amount data on page %s: date=%s lines=%s",
                                page_idx,
                                pending.date_text,
                                pending.lines,
                            )
                    pending = PendingRecord(date_text=line)
                    continue

                if pending is not None:
                    pending.lines.append(line)

    if pending is not None:
        record, previous_balance = _finalize_record(pending, previous_balance)
        if record is not None:
            records.append(record)
            if progress_cb is not None:
                progress_cb(len(records))
        else:
            logger.warning(
                "Skipped Indian Bank trailing row with insufficient amount data: date=%s lines=%s",
                pending.date_text,
                pending.lines,
            )

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("Indian Bank parse complete: rows=%s closing_balance=%s", len(records), previous_balance)
    return records


BANK_CODE = 'indian'
BANK_SIGNATURES = (('INDIAN BANK', 5), ('IDIB0', 5))
