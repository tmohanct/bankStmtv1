from __future__ import annotations

import re
from typing import Any

import pdfplumber

from parser_helpers import build_record
from utils import clean_cell, parse_amount

DATE_RE = re.compile(r"^\d{2}-[A-Za-z]{3}-\d{4}$")
DATE_FORMATS = ("%d-%b-%Y",)


def _is_transaction_row(row: list[str]) -> bool:
    return len(row) >= 7 and bool(DATE_RE.match(row[0]))


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing IDFC statement: %s", pdf_path)

    records: list[dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("Page %s: extracted %s table(s)", page_idx, len(tables))

            for table_idx, table in enumerate(tables, start=1):
                logger.debug("Page %s table %s: rows=%s", page_idx, table_idx, len(table))
                for raw_row in table:
                    row = [clean_cell(cell) for cell in raw_row]
                    if not any(row) or not _is_transaction_row(row):
                        continue

                    record = build_record(
                        date_text=row[0],
                        details=row[2],
                        cheque_no=row[3],
                        debit=parse_amount(row[4]),
                        credit=parse_amount(row[5]),
                        balance=parse_amount(row[6]),
                        date_formats=DATE_FORMATS,
                    )
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("IDFC parse complete: rows=%s", len(records))
    return records
