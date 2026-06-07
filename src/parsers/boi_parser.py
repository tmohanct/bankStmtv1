"""Bank of India statement parser implementation."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import fitz
import pandas as pd

from parsers.base_parser import BaseStatementParser

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")
DATE_FORMATS = ("%d-%m-%Y",)

SERIAL_COLUMN = (0.0, 45.0)
DATE_COLUMN = (45.0, 105.0)
DETAIL_COLUMN = (100.0, 267.0)
CHEQUE_COLUMN = (267.0, 346.0)
WITHDRAWAL_COLUMN = (346.0, 426.0)
DEPOSIT_COLUMN = (426.0, 501.0)
BALANCE_COLUMN = (501.0, 595.0)

ROW_Y_TOLERANCE = 0.5
NON_CHEQUE_DETAIL_RE = re.compile(
    r"\b(?:UPI|IMPS|NEFT|RTGS|ATM|POS|CARD|NACH|ACH)\b",
    re.IGNORECASE,
)


@dataclass(frozen=True)
class PositionedLine:
    x0: float
    top: float
    text: str


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()


def _clean_detail_key(value: Any) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", _clean_text(value))


def _parse_amount(value: Any) -> float | None:
    text = _clean_text(value)
    if not text:
        return None

    upper = text.upper()
    negative = (
        upper.startswith("-")
        or upper.endswith("-")
        or ("(" in upper and ")" in upper)
        or bool(re.search(r"\bDR\.?$", upper))
    )
    normalized = re.sub(r"\b(?:CR|DR|INR|RS)\.?", "", upper)
    normalized = (
        normalized.replace(",", "")
        .replace(" ", "")
        .replace("+", "")
        .replace("(", "")
        .replace(")", "")
    )
    number_text = re.sub(r"[^0-9.\-]", "", normalized)
    if number_text in {"", "-", "."}:
        return None

    try:
        amount = float(number_text)
    except ValueError:
        return None
    if negative and amount > 0:
        amount = -amount
    return amount


def _normalize_date(value: Any) -> str:
    text = _clean_text(value)
    for date_format in DATE_FORMATS:
        try:
            return datetime.strptime(text, date_format).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text


def _normalize_cheque_number(value: Any, details: Any) -> str:
    text = _clean_text(value)
    if not re.fullmatch(r"\d+", text or "") or set(text) == {"0"}:
        return ""
    if NON_CHEQUE_DETAIL_RE.search(_clean_text(details)):
        return ""
    return text


def _extract_lines(page: fitz.Page) -> list[PositionedLine]:
    grouped_words: dict[tuple[int, int], list[tuple[Any, ...]]] = {}

    for word in page.get_text("words") or []:
        if len(word) < 8:
            continue
        key = (int(word[5]), int(word[6]))
        grouped_words.setdefault(key, []).append(word)

    lines: list[PositionedLine] = []
    for words in grouped_words.values():
        ordered_words = sorted(words, key=lambda item: int(item[7]))
        text = _clean_text(" ".join(str(word[4]) for word in ordered_words))
        if not text:
            continue
        lines.append(
            PositionedLine(
                x0=min(float(word[0]) for word in ordered_words),
                top=min(float(word[1]) for word in ordered_words),
                text=text,
            )
        )

    return sorted(lines, key=lambda line: (line.top, line.x0))


def _column_text(
    lines: list[PositionedLine],
    column: tuple[float, float],
) -> str:
    left, right = column
    return _clean_text(" ".join(line.text for line in lines if left <= line.x0 < right))


def _parse_positive_amount(value: str) -> float | None:
    amount = _parse_amount(value)
    if amount is None:
        return None
    return abs(amount)


def _build_record(
    lines: list[PositionedLine],
    *,
    source_page: int,
    logger: logging.Logger,
) -> tuple[dict[str, Any] | None, str]:
    statement_serial = _column_text(lines, SERIAL_COLUMN)
    date_text = _column_text(lines, DATE_COLUMN)
    details = _column_text(lines, DETAIL_COLUMN)
    cheque_no = _column_text(lines, CHEQUE_COLUMN)
    debit = _parse_positive_amount(_column_text(lines, WITHDRAWAL_COLUMN))
    credit = _parse_positive_amount(_column_text(lines, DEPOSIT_COLUMN))
    balance = _parse_amount(_column_text(lines, BALANCE_COLUMN))

    if not DATE_RE.fullmatch(date_text):
        return None, statement_serial
    if balance is None:
        logger.warning(
            "Skipped Bank of India row without a balance: page=%s serial=%s date=%s",
            source_page,
            statement_serial,
            date_text,
        )
        return None, statement_serial
    if (debit is None) == (credit is None):
        logger.warning(
            "Skipped Bank of India row with ambiguous amount columns: "
            "page=%s serial=%s date=%s debit=%s credit=%s",
            source_page,
            statement_serial,
            date_text,
            debit,
            credit,
        )
        return None, statement_serial

    clean_details = _clean_text(details)
    record = {
        "Sno": 0,
        "Date": _normalize_date(date_text),
        "Details": clean_details,
        "Detail_Clean": _clean_detail_key(clean_details),
        "Cheque No": _normalize_cheque_number(cheque_no, clean_details),
        "Debit": debit,
        "Credit": credit,
        "Balance": balance,
        "Source_Page": source_page,
    }
    return record, statement_serial


def parse_boi_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    logger.info("Parsing Bank of India statement: %s", pdf_path)
    records: list[dict[str, Any]] = []

    with fitz.open(str(pdf_path)) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            lines = _extract_lines(page)
            row_tops = sorted(
                {
                    line.top
                    for line in lines
                    if DATE_COLUMN[0] <= line.x0 < DATE_COLUMN[1]
                    and DATE_RE.fullmatch(line.text)
                }
            )
            logger.debug(
                "Bank of India page %s: extracted %s positioned line(s), %s transaction row(s)",
                page_number,
                len(lines),
                len(row_tops),
            )

            for row_index, row_top in enumerate(row_tops):
                next_row_top = (
                    row_tops[row_index + 1] - ROW_Y_TOLERANCE
                    if row_index + 1 < len(row_tops)
                    else float(page.rect.height)
                )
                row_lines = [
                    line
                    for line in lines
                    if row_top - ROW_Y_TOLERANCE <= line.top < next_row_top
                ]
                record, statement_serial = _build_record(
                    row_lines,
                    source_page=page_number,
                    logger=logger,
                )
                if record is None:
                    continue

                expected_serial = len(records) + 1
                if statement_serial and statement_serial != str(expected_serial):
                    logger.warning(
                        "Bank of India statement serial mismatch: expected=%s actual=%s page=%s",
                        expected_serial,
                        statement_serial,
                        page_number,
                    )

                record["Sno"] = expected_serial
                records.append(record)
                if progress_cb is not None:
                    progress_cb(len(records))

    logger.info("Bank of India parse complete: rows=%s", len(records))
    return records


class BOIParser(BaseStatementParser):
    """Bank of India statement parser."""

    bank_code = "boi"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        _ = rules_df
        records = parse_boi_records(
            pdf_path=pdf_path,
            logger=logging.getLogger(__name__),
        )
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
