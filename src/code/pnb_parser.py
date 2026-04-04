from __future__ import annotations

import re
from typing import Any

import pdfplumber

from parser_helpers import build_record, parse_signed_balance
from utils import clean_cell, parse_amount

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")
DATE_FORMATS = ("%d-%m-%Y",)


def _is_transaction_row(row: list[str]) -> bool:
    return len(row) >= 8 and bool(DATE_RE.match(row[0]))


def _build_details(row: list[str]) -> str:
    return clean_cell(" ".join(part for part in (row[4], row[6], row[7]) if clean_cell(part)))


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing PNB statement: %s", pdf_path)

    records: list[dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("Page %s: extracted %s table(s)", page_idx, len(tables))

            for table_idx, table in enumerate(tables, start=1):
                logger.debug("Page %s table %s: rows=%s", page_idx, table_idx, len(table))
                for raw_row in table:
                    row = [clean_cell(cell) for cell in raw_row]
                    if not any(row) or row[0] == "Page Total" or not _is_transaction_row(row):
                        continue

                    record = build_record(
                        date_text=row[0],
                        details=_build_details(row),
                        cheque_no=row[5],
                        debit=parse_amount(row[1]),
                        credit=parse_amount(row[2]),
                        balance=parse_signed_balance(row[3]),
                        date_formats=DATE_FORMATS,
                    )
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("PNB parse complete: rows=%s", len(records))
    return records
