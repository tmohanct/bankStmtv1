from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import fitz

from parser_helpers import build_record
from utils import clean_cell, parse_amount

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")
AMOUNT_RE = re.compile(r"^-?[0-9,]+\.\d{2}$")
DATE_FORMATS = ("%d-%m-%Y",)
HEADER_LINES = {"Date", "Particulars", "Deposits", "Withdrawals", "Balance"}
DETAIL_SKIP_PREFIXES = (
    "page ",
    "computer output-",
    "------------------------------ end of statement",
    "are you a merchant",
    "e-mail:",
)
DEBIT_HINTS = ("-DR/", " DEBIT", "/DR/", "IMPS SC", "SC ")
CREDIT_HINTS = ("-CR/", " CREDIT", "/CR/", " NEFT CR", " RTGS CR")


@dataclass
class PendingRecord:
    date_text: str
    lines: list[str] = field(default_factory=list)


def _is_amount_line(line: str) -> bool:
    return bool(AMOUNT_RE.match(line))


def _should_skip_detail_line(line: str) -> bool:
    lowered = line.lower()
    if line in HEADER_LINES:
        return True
    return any(lowered.startswith(prefix) for prefix in DETAIL_SKIP_PREFIXES)


def _classify_amount(
    details: str,
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

    upper_details = details.upper()
    if any(token in upper_details for token in CREDIT_HINTS) and not any(
        token in upper_details for token in DEBIT_HINTS
    ):
        return None, abs_amount
    if any(token in upper_details for token in DEBIT_HINTS):
        return abs_amount, None
    if previous_balance is not None and balance_value >= previous_balance:
        return None, abs_amount
    return abs_amount, None


def _finalize_record(
    pending: PendingRecord,
    previous_balance: float | None,
) -> tuple[dict[str, Any] | None, float | None]:
    numeric_lines: list[tuple[int, float]] = []
    for idx, line in enumerate(pending.lines):
        if not _is_amount_line(line):
            continue
        parsed = parse_amount(line)
        if parsed is not None:
            numeric_lines.append((idx, parsed))

    if len(numeric_lines) >= 3:
        prev_idx, prev_value = numeric_lines[-2]
        last_idx, last_value = numeric_lines[-1]
        trailing_text = " ".join(pending.lines[prev_idx + 1 : last_idx + 1]).upper()
        if abs(last_value - prev_value) <= 0.01 and "CLOSING BALANCE" in trailing_text:
            numeric_lines.pop()

    if len(numeric_lines) < 2:
        return None, previous_balance

    amount_idx, amount_value = numeric_lines[-2]
    _, balance_value = numeric_lines[-1]

    cheque_no = ""
    detail_parts: list[str] = []
    for line in pending.lines[:amount_idx]:
        if _should_skip_detail_line(line):
            continue
        if line.upper().startswith("CHQ:"):
            cheque_no = clean_cell(line.split(":", 1)[1] if ":" in line else line[4:])
            continue
        detail_parts.append(line)

    details = clean_cell(" ".join(detail_parts))
    debit, credit = _classify_amount(
        details=details,
        amount_value=amount_value,
        balance_value=balance_value,
        previous_balance=previous_balance,
    )
    record = build_record(
        date_text=pending.date_text,
        details=details,
        cheque_no=cheque_no,
        debit=debit,
        credit=credit,
        balance=balance_value,
        date_formats=DATE_FORMATS,
    )
    return record, balance_value


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Canara statement: %s", pdf_path)

    records: list[dict[str, Any]] = []
    pending: PendingRecord | None = None
    in_transactions = False
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

                if line in HEADER_LINES:
                    if line == "Balance":
                        in_transactions = True
                    continue

                if not in_transactions:
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
                                "Skipped Canara row with insufficient amount data on page %s: date=%s lines=%s",
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
                "Skipped Canara trailing row with insufficient amount data: date=%s lines=%s",
                pending.date_text,
                pending.lines,
            )

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("Canara parse complete: rows=%s closing_balance=%s", len(records), previous_balance)
    return records
