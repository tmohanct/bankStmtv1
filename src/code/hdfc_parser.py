from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import pdfplumber

from utils import clean_cell, clean_detail, normalize_date, parse_amount

TRANSACTION_LINE_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{2})\s+"
    r"(?P<body>.+?)\s+"
    r"(?P<value_date>\d{2}/\d{2}/\d{2})\s+"
    r"(?P<amount>[\-0-9,]+\.\d{2})\s+"
    r"(?P<balance>[\-0-9,]+\.\d{2})$"
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
    amount_value: float
    balance_value: float
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
    balance_value: float,
    previous_balance: float | None,
) -> tuple[float | None, float | None]:
    abs_amount = abs(amount_value)
    if previous_balance is not None:
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
    if previous_balance is not None and balance_value >= previous_balance:
        return None, amount_value if amount_value < 0 else abs_amount
    return amount_value if amount_value < 0 else abs_amount, None


def _finalize_record(
    pending: PendingRecord,
    previous_balance: float | None,
) -> tuple[dict[str, Any], float]:
    details = clean_cell(" ".join([pending.detail_head, *pending.continuation_lines]))
    debit, credit = _classify_amount(
        details=details,
        amount_value=pending.amount_value,
        balance_value=pending.balance_value,
        previous_balance=previous_balance,
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
        pending.balance_value,
    )


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing HDFC statement: %s", pdf_path)

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
                    if statement_started and TRANSACTION_LINE_RE.match(line):
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

                    detail_head, cheque_no = _split_body(match.group("body"))
                    amount_value = parse_amount(match.group("amount"))
                    balance_value = parse_amount(match.group("balance"))
                    if amount_value is None or balance_value is None:
                        continue

                    pending = PendingRecord(
                        date_text=match.group("date"),
                        detail_head=detail_head,
                        cheque_no=cheque_no,
                        amount_value=amount_value,
                        balance_value=balance_value,
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
    return records


