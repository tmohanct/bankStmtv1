"""Union Bank parser implementation."""

from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd
import pdfplumber

from parsers.base_parser import BaseStatementParser

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\r", " ").replace("\n", " ")).strip()


def _normalize_output_date(value: str) -> str:
    text = _clean_text(value)
    if not text:
        return ""

    for fmt in ("%d-%m-%Y", "%d-%m-%y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text


def _clean_detail_key(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", _clean_text(value))


def _parse_signed_amount(value: str) -> float | None:
    text = _clean_text(value)
    if not text:
        return None

    upper = text.upper().rstrip(".")
    negative = upper.endswith("(DR)") or text.startswith("-")
    cleaned = re.sub(r"\((CR|DR)\)$", "", upper, flags=re.IGNORECASE).replace(",", "").strip()
    if not cleaned:
        return None

    try:
        amount = float(cleaned)
    except ValueError:
        return None

    if negative and amount > 0:
        amount = -amount
    return amount


def _split_debit_credit(amount_text: str) -> tuple[float | None, float | None]:
    amount = _parse_signed_amount(amount_text)
    if amount is None:
        return None, None
    if amount < 0:
        return abs(amount), None
    return None, abs(amount)


def parse_unionbank_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    logger.info("Parsing Union Bank statement: %s", pdf_path)

    records: list[dict[str, Any]] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("UnionBank page %s: extracted %s table(s)", page_idx, len(tables))

            for table in tables:
                for raw_row in table:
                    row = [_clean_text(cell) for cell in raw_row]
                    if len(row) < 5:
                        continue

                    date_text, transaction_id, remarks, amount_text, balance_text = row[:5]
                    if not DATE_RE.match(date_text):
                        continue

                    debit, credit = _split_debit_credit(amount_text)
                    balance = _parse_signed_amount(balance_text)
                    details = _clean_text(f"{transaction_id} {remarks}")
                    if not details or balance is None or (debit is None and credit is None):
                        continue

                    record = {
                        "Sno": 0,
                        "Date": _normalize_output_date(date_text),
                        "Details": details,
                        "Detail_Clean": _clean_detail_key(details),
                        "Cheque No": "",
                        "Debit": debit,
                        "Credit": credit,
                        "Balance": balance,
                    }
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("Union Bank parse complete: rows=%s", len(records))
    return records


class UnionBankParser(BaseStatementParser):
    """Union Bank statement parser."""

    bank_code = "unionbank"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        _ = rules_df
        logger = logging.getLogger(__name__)
        records = parse_unionbank_records(pdf_path=pdf_path, logger=logger)

        rows = [
            {
                "Date": record["Date"],
                "Value_Date": record["Date"],
                "Description": record["Details"],
                "Debit": record["Debit"],
                "Credit": record["Credit"],
                "Balance": record["Balance"],
                "Reference": record["Cheque No"],
                "Source_Page": None,
            }
            for record in records
        ]

        return pd.DataFrame(
            rows,
            columns=[
                "Date",
                "Value_Date",
                "Description",
                "Debit",
                "Credit",
                "Balance",
                "Reference",
                "Source_Page",
            ],
        )
