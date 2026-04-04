"""South Indian Bank parser implementation."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import fitz
import pandas as pd

from parsers.base_parser import BaseStatementParser

TRANSACTION_DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{2}$")
PAGE_TOTAL_RE = re.compile(r"^Page\s+Total\b", re.IGNORECASE)

DATE_COLUMN_MAX_X = 80
DETAIL_COLUMN_MIN_X = 80
DETAIL_COLUMN_MAX_X = 250
CHEQUE_COLUMN_MIN_X = 250
CHEQUE_COLUMN_MAX_X = 335
WITHDRAWAL_COLUMN_MIN_X = 335
WITHDRAWAL_COLUMN_MAX_X = 445
DEPOSIT_COLUMN_MIN_X = 445
DEPOSIT_COLUMN_MAX_X = 520
LINE_GROUP_TOLERANCE = 2.5


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\r", " ").replace("\n", " ")).strip()


def _normalize_output_date(value: str) -> str:
    text = _clean_text(value)
    if not text:
        return ""

    for fmt in ("%d-%m-%y", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text


def _clean_detail_key(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", _clean_text(value))


def _parse_amount(value: str) -> float | None:
    text = _clean_text(value)
    if not text:
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


@dataclass
class _WordLine:
    y_center: float
    words: list[tuple[float, str]] = field(default_factory=list)

    @property
    def text(self) -> str:
        return _clean_text(" ".join(text for _, text in sorted(self.words)))


@dataclass
class _PendingRecord:
    date_text: str
    detail_lines: list[str] = field(default_factory=list)
    cheque_parts: list[str] = field(default_factory=list)
    debit: float | None = None
    credit: float | None = None
    balance: float | None = None

    def add_line(self, line: _WordLine) -> None:
        detail_words = [text for x, text in line.words if DETAIL_COLUMN_MIN_X <= x < DETAIL_COLUMN_MAX_X]
        cheque_words = [text for x, text in line.words if CHEQUE_COLUMN_MIN_X <= x < CHEQUE_COLUMN_MAX_X]
        debit_words = [text for x, text in line.words if WITHDRAWAL_COLUMN_MIN_X <= x < WITHDRAWAL_COLUMN_MAX_X]
        credit_words = [text for x, text in line.words if DEPOSIT_COLUMN_MIN_X <= x < DEPOSIT_COLUMN_MAX_X]
        balance_words = [text for x, text in line.words if x >= DEPOSIT_COLUMN_MAX_X]

        if detail_words:
            self.detail_lines.append(_clean_text(" ".join(detail_words)))
        if cheque_words:
            self.cheque_parts.append(_clean_text(" ".join(cheque_words)))
        if self.debit is None and debit_words:
            self.debit = _parse_amount(" ".join(debit_words))
        if self.credit is None and credit_words:
            self.credit = _parse_amount(" ".join(credit_words))
        if self.balance is None and balance_words:
            self.balance = _parse_amount(" ".join(balance_words))

    def finalize(self) -> dict[str, Any] | None:
        details = _clean_text(" ".join(part for part in self.detail_lines if part))
        cheque_joined = _clean_text(" ".join(part for part in self.cheque_parts if part))
        cheque_matches = re.findall(r"\d{3,}", cheque_joined)
        cheque_no = cheque_matches[0] if cheque_matches else cheque_joined

        if not details or self.balance is None or (self.debit is None and self.credit is None):
            return None

        return {
            "Sno": 0,
            "Date": _normalize_output_date(self.date_text),
            "Details": details,
            "Detail_Clean": _clean_detail_key(details),
            "Cheque No": cheque_no,
            "Debit": abs(self.debit) if self.debit is not None else None,
            "Credit": abs(self.credit) if self.credit is not None else None,
            "Balance": self.balance,
        }


def _build_word_lines(page: fitz.Page) -> list[_WordLine]:
    words = sorted(page.get_text("words"), key=lambda item: (((item[1] + item[3]) / 2), item[0]))
    lines: list[_WordLine] = []
    current_line: _WordLine | None = None

    for x0, y0, x1, y1, text, *_ in words:
        cleaned = _clean_text(text)
        if not cleaned:
            continue

        y_center = (y0 + y1) / 2
        if current_line is None or abs(y_center - current_line.y_center) > LINE_GROUP_TOLERANCE:
            current_line = _WordLine(y_center=y_center)
            lines.append(current_line)
        current_line.words.append((x0, cleaned))

    return lines


def _parse_page_lines(
    lines: list[_WordLine],
    records: list[dict[str, Any]],
    progress_cb: Callable[[int], None] | None = None,
) -> None:
    in_transaction_table = False
    pending: _PendingRecord | None = None

    for line in lines:
        line_text = line.text
        if not line_text:
            continue

        if "DATE" in line_text and "PARTICULARS" in line_text and "BALANCE" in line_text:
            in_transaction_table = True
            continue

        if not in_transaction_table:
            continue

        if PAGE_TOTAL_RE.match(line_text):
            if pending is not None:
                record = pending.finalize()
                if record is not None:
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))
            break

        if line_text.startswith("Page ") or line_text.startswith("Visit us at"):
            continue

        date_words = [text for x, text in line.words if x < DATE_COLUMN_MAX_X]
        date_text = _clean_text(" ".join(date_words))
        if TRANSACTION_DATE_RE.match(date_text):
            if pending is not None:
                record = pending.finalize()
                if record is not None:
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))

            pending = _PendingRecord(date_text=date_text)
            pending.add_line(line)
            continue

        if pending is not None:
            pending.add_line(line)

    else:
        if pending is not None:
            record = pending.finalize()
            if record is not None:
                records.append(record)
            if progress_cb is not None:
                progress_cb(len(records))


def parse_southind_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    logger.info("Parsing South Indian Bank statement: %s", pdf_path)

    records: list[dict[str, Any]] = []
    with fitz.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            lines = _build_word_lines(page)
            logger.debug("SouthInd page %s: grouped %s word line(s)", page_idx, len(lines))
            _parse_page_lines(lines, records, progress_cb=progress_cb)

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("South Indian Bank parse complete: rows=%s", len(records))
    return records


class SouthIndianParser(BaseStatementParser):
    """South Indian Bank statement parser."""

    bank_code = "southind"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        _ = rules_df
        logger = logging.getLogger(__name__)
        records = parse_southind_records(pdf_path=pdf_path, logger=logger)

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
