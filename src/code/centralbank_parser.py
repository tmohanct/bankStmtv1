from __future__ import annotations

import re
from typing import Any

import pdfplumber

from parser_helpers import build_record, parse_signed_balance
from utils import clean_cell, is_date_token, parse_amount

DATE_FORMATS = ("%d/%m/%Y",)


def _is_header_row(row: list[str]) -> bool:
    normalized = " ".join(row).upper()
    return (
        "POST DATE" in normalized
        and "TRANSACTION DESCRIPTION" in normalized
        and "BALANCE" in normalized
    )


def _is_transaction_row(row: list[str]) -> bool:
    return len(row) >= 8 and is_date_token(row[0]) and is_date_token(row[1])


def _clean_description(value: str) -> str:
    return re.sub(r"\s+", " ", clean_cell(value)).strip()


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Central Bank statement: %s", pdf_path)

    records: list[dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("Page %s: extracted %s table(s)", page_idx, len(tables))

            for table_idx, table in enumerate(tables, start=1):
                logger.debug("Page %s table %s: rows=%s", page_idx, table_idx, len(table))
                for raw_row in table:
                    row = [clean_cell(cell) for cell in raw_row]
                    if not any(row) or _is_header_row(row) or not _is_transaction_row(row):
                        continue

                    record = build_record(
                        date_text=row[0],
                        details=_clean_description(row[4]),
                        cheque_no=row[3],
                        debit=parse_amount(row[5]),
                        credit=parse_amount(row[6]),
                        balance=parse_signed_balance(row[7]),
                        date_formats=DATE_FORMATS,
                    )
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("Central Bank parse complete: rows=%s", len(records))
    return records
