"""Indian Overseas Bank parser implementation."""

from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd
import pdfplumber

from parsers.base_parser import BaseStatementParser
from utils.amount_utils import parse_amount

ACCOUNT_NUMBER_RE = re.compile(r"Account\s+No\s*:\s*([0-9A-Za-z]+)", re.IGNORECASE)
OUTPUT_COLUMNS = [
    "Date",
    "ValueDate",
    "Narration",
    "Debit",
    "Credit",
    "Balance",
    "Currency",
    "Account_Number",
    "Txn_Ref",
    "Page",
]


def _clean_cell(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\r", " ").replace("\n", " ")).strip()


def _parse_date_token(raw_value: str) -> str | None:
    text = _clean_cell(raw_value)
    if not text:
        return None

    for fmt in ("%d-%b-%y", "%d-%b-%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return None


def _extract_dates(raw_value: object) -> tuple[str | None, str | None]:
    if raw_value is None:
        return None, None

    parts = [
        part.strip().strip("()").strip()
        for part in str(raw_value).replace("\r", "\n").splitlines()
        if part and part.strip()
    ]
    if not parts:
        return None, None

    txn_date = _parse_date_token(parts[0])
    value_date = _parse_date_token(parts[1]) if len(parts) > 1 else None
    if txn_date is None:
        return None, None
    return txn_date, value_date or txn_date


def _extract_account_number(pdf: pdfplumber.PDF) -> str | None:
    if not pdf.pages:
        return None

    first_page_text = pdf.pages[0].extract_text() or ""
    match = ACCOUNT_NUMBER_RE.search(first_page_text)
    if match is None:
        return None
    return match.group(1).strip()


def _build_record(row: list[object], page_number: int, account_number: str | None) -> dict[str, object] | None:
    if len(row) < 7:
        return None

    txn_date, value_date = _extract_dates(row[0])
    if txn_date is None:
        return None

    narration = _clean_cell(row[1])
    txn_ref = _clean_cell(row[2])
    debit = parse_amount(row[4])
    credit = parse_amount(row[5])
    balance = parse_amount(row[6])

    return {
        "Date": txn_date,
        "ValueDate": value_date,
        "Narration": narration,
        "Debit": debit,
        "Credit": credit,
        "Balance": balance,
        "Currency": "INR",
        "Account_Number": account_number,
        "Txn_Ref": txn_ref,
        "Page": page_number,
    }


class IOBParser(BaseStatementParser):
    """Indian Overseas Bank statement parser."""

    bank_code = "iob"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        """Parse IOB statement PDF and return raw transaction rows."""
        _ = rules_df

        rows: list[dict[str, object]] = []

        with pdfplumber.open(str(pdf_path)) as pdf:
            account_number = _extract_account_number(pdf)

            for page_number, page in enumerate(pdf.pages, start=1):
                tables = page.extract_tables() or []
                for table in tables:
                    for raw_row in table:
                        record = _build_record(raw_row, page_number, account_number)
                        if record is not None:
                            rows.append(record)

        return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
