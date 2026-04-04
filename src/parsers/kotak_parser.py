"""Kotak Mahindra Bank parser implementation."""

from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd
import pdfplumber

from parsers.base_parser import BaseStatementParser

DATE_RE = re.compile(r"^\d{2}\s+[A-Za-z]{3}\s+\d{4}$")


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\r", " ").replace("\n", " ")).strip()


def _normalize_output_date(value: str) -> str:
    text = _clean_text(value)
    if not text:
        return ""

    try:
        return datetime.strptime(text, "%d %b %Y").strftime("%d/%m/%Y")
    except ValueError:
        return text


def _clean_detail_key(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", _clean_text(value))


def _parse_amount(value: str) -> float | None:
    text = _clean_text(value)
    if not text or text == "-":
        return None

    upper = text.upper().rstrip(".")
    negative = text.startswith("-") or upper.endswith("DR")
    cleaned = re.sub(r"\s*(CR|DR)\.?$", "", text, flags=re.IGNORECASE)
    cleaned = cleaned.replace(",", "").strip()
    if not cleaned:
        return None

    try:
        amount = float(cleaned)
    except ValueError:
        return None

    if negative and amount > 0:
        amount = -amount
    return amount


def parse_kotak_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    logger.info("Parsing Kotak statement: %s", pdf_path)

    records: list[dict[str, Any]] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("Kotak page %s: extracted %s table(s)", page_idx, len(tables))

            for table in tables:
                for raw_row in table:
                    row = [_clean_text(cell) for cell in raw_row]
                    if len(row) < 7:
                        continue

                    _, date_text, description, reference, withdrawal_text, deposit_text, balance_text = row[:7]
                    if not DATE_RE.match(date_text):
                        continue

                    details = _clean_text(description)
                    debit = _parse_amount(withdrawal_text)
                    credit = _parse_amount(deposit_text)
                    balance = _parse_amount(balance_text)
                    if not details or balance is None or (debit is None and credit is None):
                        continue

                    record = {
                        "Sno": 0,
                        "Date": _normalize_output_date(date_text),
                        "Details": details,
                        "Detail_Clean": _clean_detail_key(details),
                        "Cheque No": _clean_text(reference),
                        "Debit": abs(debit) if debit is not None else None,
                        "Credit": abs(credit) if credit is not None else None,
                        "Balance": balance,
                        "Source_Page": page_idx,
                    }
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("Kotak parse complete: rows=%s", len(records))
    return records


class KotakParser(BaseStatementParser):
    """Kotak Mahindra Bank statement parser."""

    bank_code = "kotak"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        _ = rules_df
        logger = logging.getLogger(__name__)
        records = parse_kotak_records(pdf_path=pdf_path, logger=logger)

        rows = [
            {
                "Date": record["Date"],
                "Value_Date": record["Date"],
                "Description": record["Details"],
                "Debit": record["Debit"],
                "Credit": record["Credit"],
                "Balance": record["Balance"],
                "Reference": record["Cheque No"],
                "Source_Page": record["Source_Page"],
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
