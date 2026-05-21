from __future__ import annotations

import re
from typing import Any

import pdfplumber

from parser_helpers import build_record
from utils import clean_cell, parse_amount

DATE_RE = re.compile(r"^\d{2}-[A-Z]{3}-\s*\d{4}$")
DATE_FORMATS = ("%d-%b-%Y",)
OLD_FORMAT_COLUMNS = {
    "date": 0,
    "details": 2,
    "cheque": 4,
    "debit": 5,
    "credit": 6,
    "balance": 7,
}
NEW_FORMAT_COLUMNS = {
    "date": 0,
    "details": 2,
    "cheque": 5,
    "debit": 6,
    "credit": 7,
    "balance": 8,
}


def _normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]", "", clean_cell(value).lower())


def _detect_column_map(row: list[str]) -> dict[str, int] | None:
    column_map: dict[str, int] = {}
    for index, value in enumerate(row):
        token = _normalize_header(value)
        if token == "date":
            column_map["date"] = index
        elif token in {"particulars", "narration", "description"}:
            column_map["details"] = index
        elif token in {"chequedetails", "chequeno", "chequenumber"}:
            column_map["cheque"] = index
        elif token in {"withdrawals", "withdrawal", "debit", "debits"}:
            column_map["debit"] = index
        elif token in {"deposits", "deposit", "credit", "credits"}:
            column_map["credit"] = index
        elif token == "balance":
            column_map["balance"] = index

    required_columns = {"date", "details", "debit", "credit", "balance"}
    if required_columns.issubset(column_map):
        return column_map
    return None


def _fallback_column_map(row: list[str]) -> dict[str, int]:
    if len(row) >= 9:
        return NEW_FORMAT_COLUMNS
    return OLD_FORMAT_COLUMNS


def _pick(row: list[str], column_map: dict[str, int], key: str) -> str:
    index = column_map.get(key, -1)
    if index < 0 or index >= len(row):
        return ""
    return row[index]


def _is_transaction_row(row: list[str]) -> bool:
    return len(row) >= 8 and bool(DATE_RE.match(row[0].upper()))


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Federal statement: %s", pdf_path)

    records: list[dict[str, Any]] = []
    column_map: dict[str, int] | None = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("Page %s: extracted %s table(s)", page_idx, len(tables))

            for table_idx, table in enumerate(tables, start=1):
                logger.debug("Page %s table %s: rows=%s", page_idx, table_idx, len(table))
                for raw_row in table:
                    row = [clean_cell(cell) for cell in raw_row]
                    detected_column_map = _detect_column_map(row)
                    if detected_column_map is not None:
                        column_map = detected_column_map
                        logger.debug(
                            "Page %s table %s: Federal column map=%s",
                            page_idx,
                            table_idx,
                            column_map,
                        )
                        continue

                    if not any(row) or not _is_transaction_row(row):
                        continue

                    row_column_map = column_map or _fallback_column_map(row)
                    record = build_record(
                        date_text=_pick(row, row_column_map, "date"),
                        details=_pick(row, row_column_map, "details"),
                        cheque_no=_pick(row, row_column_map, "cheque"),
                        debit=parse_amount(_pick(row, row_column_map, "debit")),
                        credit=parse_amount(_pick(row, row_column_map, "credit")),
                        balance=parse_amount(_pick(row, row_column_map, "balance")),
                        date_formats=DATE_FORMATS,
                    )
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("Federal parse complete: rows=%s", len(records))
    return records
