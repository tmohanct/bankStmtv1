"""Equitas Small Finance Bank (ESFB) statement parser."""

from __future__ import annotations

import logging
import re
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd
import pdfplumber

from parsers.base_parser import BaseStatementParser

DATE_RE = re.compile(r"^\d{2}-[A-Za-z]{3}-\d{4}$")
AMOUNT_RE = re.compile(r"(?<!\d)(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?(?!\d)")
NARRATION_PREFIX_ARTIFACT_RE = re.compile(
    r"^\d+\s*(?=(?:IMP\s+P2A|UPI\s+REF\s+NO)\b)",
    re.IGNORECASE,
)


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\r", " ").replace("\n", " ")).strip()


def _clean_detail_key(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", _clean_text(value))


def _normalize_date(value: str) -> str:
    text = _clean_text(value)
    try:
        return datetime.strptime(text, "%d-%b-%Y").strftime("%d/%m/%Y")
    except ValueError:
        return text


def _parse_amount_cell(value: Any) -> tuple[float | None, str]:
    """Return the first amount and any non-amount spillover text in a cell."""
    text = _clean_text(value)
    if not text:
        return None, ""

    match = AMOUNT_RE.search(text)
    if match is None:
        return None, text if text != "-" else ""

    amount = float(match.group(0).replace(",", ""))
    overflow = _clean_text(f"{text[:match.start()]} {text[match.end():]}").strip(" -")
    return amount, overflow


def _clean_narration(value: Any, *spillovers: str) -> str:
    narration = NARRATION_PREFIX_ARTIFACT_RE.sub("", _clean_text(value))
    extra_text = " ".join(text for text in spillovers if text)
    return _clean_text(f"{narration} {extra_text}")


def _select_cheque_number(reference: Any, details: str) -> str:
    value = _clean_text(reference)
    if re.fullmatch(r"\d+", value) and re.search(r"\bCHQ\b", details, re.IGNORECASE):
        return value
    return ""


def _parse_transaction_row(row: list[str], source_page: int) -> dict[str, Any] | None:
    if len(row) < 6:
        return None

    date_text = _clean_text(row[0])
    if not DATE_RE.fullmatch(date_text):
        return None

    debit, debit_spillover = _parse_amount_cell(row[3])
    credit, credit_spillover = _parse_amount_cell(row[4])
    balance, balance_spillover = _parse_amount_cell(row[5])
    details = _clean_narration(row[2], debit_spillover, credit_spillover, balance_spillover)

    if balance is None or (debit is None and credit is None):
        return None

    return {
        "Sno": 0,
        "Date": _normalize_date(date_text),
        "Details": details,
        "Detail_Clean": _clean_detail_key(details),
        "Cheque No": _select_cheque_number(row[1], details),
        "Debit": debit,
        "Credit": credit,
        "Balance": balance,
        "Source_Page": source_page,
    }


def _validate_balance_continuity(records: list[dict[str, Any]], logger: logging.Logger) -> None:
    for previous, current in zip(records, records[1:]):
        previous_balance = previous.get("Balance")
        current_balance = current.get("Balance")
        if previous_balance is None or current_balance is None:
            continue

        debit = float(current.get("Debit") or 0.0)
        credit = float(current.get("Credit") or 0.0)
        expected_balance = float(previous_balance) + credit - debit
        if abs(expected_balance - float(current_balance)) > 0.011:
            logger.warning(
                "ESFB balance discontinuity at row %s: expected %.2f, found %.2f",
                current.get("Sno"),
                expected_balance,
                float(current_balance),
            )


def parse_esfb_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    logger.info("Parsing Equitas Small Finance Bank statement: %s", pdf_path)

    records: list[dict[str, Any]] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("ESFB page %s: extracted %s table(s)", page_idx, len(tables))

            for table in tables:
                for raw_row in table:
                    row = [_clean_text(cell) for cell in raw_row]
                    record = _parse_transaction_row(row, page_idx)
                    if record is None:
                        continue

                    records.append(record)
                    record["Sno"] = len(records)
                    if progress_cb is not None:
                        progress_cb(len(records))

    _validate_balance_continuity(records, logger)
    logger.info("ESFB parse complete: rows=%s", len(records))
    return records


class ESFBParser(BaseStatementParser):
    """Equitas Small Finance Bank statement parser."""

    bank_code = "esfb"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        _ = rules_df
        records = parse_esfb_records(pdf_path, logging.getLogger(__name__))
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
