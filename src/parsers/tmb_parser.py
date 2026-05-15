"""Tamilnad Mercantile Bank parser implementation."""

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
TABLE_DATE_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
PAGE_FOOTER_RE = re.compile(r"^Page\s+\d+\s+of\s+\d+$", re.IGNORECASE)

LINE_Y_TOLERANCE = 3.0

DATE_COLUMN = (15.0, 76.0)
DETAIL_COLUMN = (76.0, 270.0)
CHEQUE_COLUMN = (270.0, 340.0)
WITHDRAWAL_COLUMN = (340.0, 430.0)
DEPOSIT_COLUMN = (430.0, 505.0)
BALANCE_COLUMN = (505.0, 590.0)


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()


def _clean_table_detail(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", "").replace("\n", "")
    return re.sub(r"\s+", " ", text).strip()


def _clean_detail_key(value: Any) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", _clean_text(value))


def _parse_amount(value: Any) -> float | None:
    text = _clean_text(value)
    if not text or text == "-":
        return None

    normalized = text.upper()
    negative = bool(re.search(r"\bDR\.?\b|\bDR\.?$", normalized))
    normalized = re.sub(r"\b(?:CR|DR|INR|RS)\.?", "", normalized)
    normalized = (
        normalized.replace(",", "")
        .replace(" ", "")
        .replace("+", "")
        .replace("(", "-")
        .replace(")", "")
    )
    if normalized in {"", "-", "."}:
        return None

    try:
        amount = float(normalized)
    except ValueError:
        return None

    if negative and amount > 0:
        return -amount
    return amount


def _normalize_date(value: Any) -> str:
    text = _clean_text(value)
    if not text:
        return ""

    for date_format in ("%d-%m-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, date_format).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text


def _normalize_cheque_number(value: Any, details: Any = "") -> str:
    text = _clean_text(value)
    if not text or text == "-":
        return ""
    if not re.fullmatch(r"\d+", text):
        return ""

    details_text = _clean_text(details)
    if re.search(r"\b(?:UPI|IMPS|NEFT|RTGS|MBANK|EBANK|ATM|POS|CARD)\b", details_text, re.IGNORECASE):
        return ""
    return text


def _group_words_by_line(words: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    lines: list[tuple[float, list[dict[str, Any]]]] = []

    for word in sorted(words, key=lambda item: (float(item["top"]), float(item["x0"]))):
        top = float(word["top"])
        if not lines or abs(lines[-1][0] - top) > LINE_Y_TOLERANCE:
            lines.append((top, [word]))
            continue

        previous_top, line_words = lines[-1]
        line_words.append(word)
        lines[-1] = (
            ((previous_top * (len(line_words) - 1)) + top) / len(line_words),
            line_words,
        )

    return [sorted(line_words, key=lambda item: float(item["x0"])) for _, line_words in lines]


def _column_text(words: list[dict[str, Any]], left: float, right: float) -> str:
    return _clean_text(
        " ".join(
            str(word["text"])
            for word in words
            if left <= float(word["x0"]) < right
        )
    )


def _line_text(words: list[dict[str, Any]]) -> str:
    return _clean_text(" ".join(str(word["text"]) for word in words))


def _is_header_or_summary_line(text: str) -> bool:
    upper_text = _clean_text(text).upper()
    if not upper_text:
        return True
    if PAGE_FOOTER_RE.match(text):
        return True
    if "OPENING BALANCE" in upper_text or "CLOSING BALANCE" in upper_text:
        return True
    if "PARTICULARS" in upper_text and "BALANCE(INR)" in upper_text:
        return True
    return False


def _new_record(
    *,
    date_text: str,
    details: str,
    cheque_no: str,
    withdrawal: str,
    deposit: str,
    balance: str,
    source_page: int,
) -> dict[str, Any]:
    clean_details = _clean_text(details)
    date = _normalize_date(date_text)
    return {
        "Sno": 0,
        "Date": date,
        "Value_Date": date,
        "Details": clean_details,
        "Detail_Clean": _clean_detail_key(clean_details),
        "Cheque No": _normalize_cheque_number(cheque_no, clean_details),
        "Debit": _parse_amount(withdrawal),
        "Credit": _parse_amount(deposit),
        "Balance": _parse_amount(balance),
        "Source_Page": source_page,
    }


def _is_table_transaction_row(row: list[str]) -> bool:
    return len(row) >= 7 and bool(TABLE_DATE_RE.match(row[0]))


def _build_table_record(row: list[str], source_page: int) -> dict[str, Any]:
    details = _clean_table_detail(row[3])
    date = _normalize_date(row[0])
    value_date = _normalize_date(row[1]) or date
    cheque_no = _clean_text(row[2])

    return {
        "Sno": 0,
        "Date": date,
        "Value_Date": value_date,
        "Details": details,
        "Detail_Clean": _clean_detail_key(details),
        "Cheque No": _normalize_cheque_number(cheque_no, details),
        "Debit": _parse_amount(row[4]),
        "Credit": _parse_amount(row[5]),
        "Balance": _parse_amount(row[6]),
        "Source_Page": source_page,
    }


def _parse_table_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("TMB page %s: extracted %s table(s)", page_idx, len(tables))

            for table in tables:
                for raw_row in table:
                    row = [_clean_text(cell) for cell in raw_row]
                    if not _is_table_transaction_row(row):
                        continue

                    records.append(_build_table_record(raw_row, page_idx))
                    if progress_cb is not None:
                        progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    return records


def _parse_positioned_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    active_record: dict[str, Any] | None = None

    def flush_active_record() -> None:
        nonlocal active_record
        if active_record is None:
            return

        details = _clean_text(active_record["Details"])
        active_record["Details"] = details
        active_record["Detail_Clean"] = _clean_detail_key(details)
        active_record["Cheque No"] = _normalize_cheque_number(active_record["Cheque No"], details)
        records.append(active_record)
        active_record = None

        if progress_cb is not None:
            progress_cb(len(records))

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            words = page.extract_words(x_tolerance=1, y_tolerance=3, keep_blank_chars=False) or []
            lines = _group_words_by_line(words)
            logger.debug("TMB page %s: extracted %s text line(s)", page_idx, len(lines))

            for line_words in lines:
                date_text = _column_text(line_words, *DATE_COLUMN)
                details = _column_text(line_words, *DETAIL_COLUMN)
                cheque_no = _column_text(line_words, *CHEQUE_COLUMN)
                withdrawal = _column_text(line_words, *WITHDRAWAL_COLUMN)
                deposit = _column_text(line_words, *DEPOSIT_COLUMN)
                balance = _column_text(line_words, *BALANCE_COLUMN)
                full_text = _line_text(line_words)

                if DATE_RE.match(date_text):
                    flush_active_record()
                    active_record = _new_record(
                        date_text=date_text,
                        details=details,
                        cheque_no=cheque_no,
                        withdrawal=withdrawal,
                        deposit=deposit,
                        balance=balance,
                        source_page=page_idx,
                    )
                    continue

                if _is_header_or_summary_line(full_text):
                    flush_active_record()
                    continue

                if active_record is None:
                    continue

                if details:
                    active_record["Details"] = _clean_text(f"{active_record['Details']} {details}")
                    if cheque_no and not active_record["Cheque No"]:
                        active_record["Cheque No"] = cheque_no
                    continue

                if cheque_no or withdrawal or deposit or balance:
                    flush_active_record()

            flush_active_record()

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    return records


def parse_tmb_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    logger.info("Parsing TMB statement: %s", pdf_path)

    records = _parse_table_records(pdf_path=pdf_path, logger=logger, progress_cb=progress_cb)
    if not records:
        records = _parse_positioned_records(pdf_path=pdf_path, logger=logger, progress_cb=progress_cb)

    logger.info("TMB parse complete: rows=%s", len(records))
    return records


class TMBParser(BaseStatementParser):
    """Tamilnad Mercantile Bank statement parser."""

    bank_code = "tmb"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        _ = rules_df
        logger = logging.getLogger(__name__)
        records = parse_tmb_records(pdf_path=pdf_path, logger=logger)

        rows = [
            {
                "Date": record["Date"],
                "Value_Date": record.get("Value_Date", record["Date"]),
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
