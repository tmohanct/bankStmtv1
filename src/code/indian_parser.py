from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import fitz

from parser_helpers import build_record
from utils import clean_cell, parse_amount

DATE_RE = re.compile(r"^\d{2} [A-Za-z]{3} \d{4}$")
DATE_FORMATS = ("%d %b %Y", "%d %B %Y")
HEADER_LINES = {"Date", "Transaction Details", "Debits", "Credits", "Balance"}
DETAIL_SKIP_PREFIXES = ("page ",)
DETAIL_SKIP_LINES = {"Ending Balance", "Total"}
DETAIL_SKIP_PAGE_RE = re.compile(r"^\d+/\d+$")
AMOUNT_TEXT_RE = re.compile(r"^(?:-|\+?\s*INR\s+[0-9,]+\.\d{2})$", re.IGNORECASE)


@dataclass
class PendingRecord:
    date_text: str
    lines: list[str] = field(default_factory=list)


def _is_amount_or_dash(line: str) -> bool:
    return bool(AMOUNT_TEXT_RE.match(line))


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
    amount_markers = [(idx, line) for idx, line in enumerate(pending.lines) if _is_amount_or_dash(line)]
    if len(amount_markers) < 2:
        return None, previous_balance

    row_markers = amount_markers[:3]
    row_marker_indexes = {idx for idx, _ in row_markers}
    detail_parts = [
        line
        for idx, line in enumerate(pending.lines)
        if idx not in row_marker_indexes
        and not _is_amount_or_dash(line)
        and not _should_skip_detail_line(line)
    ]
    details = clean_cell(" ".join(detail_parts))

    debit: float | None = None
    credit: float | None = None

    if len(row_markers) == 3:
        debit_text = row_markers[0][1]
        credit_text = row_markers[1][1]
        balance_text = row_markers[2][1]
        debit = parse_amount(debit_text) if debit_text != "-" else None
        credit = parse_amount(credit_text) if credit_text != "-" else None
        balance_value = parse_amount(balance_text)
    else:
        amount_text = row_markers[0][1]
        balance_text = row_markers[1][1]
        amount_value = parse_amount(amount_text)
        balance_value = parse_amount(balance_text)
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


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Indian Bank statement: %s", pdf_path)

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
                    opening_balance = parse_amount(line)
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
