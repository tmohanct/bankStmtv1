from __future__ import annotations

import re
from typing import Any

import pdfplumber

from utils import clean_cell, clean_detail, normalize_date, parse_amount

SERIAL_RE = re.compile(r"^\d+$")


def _is_transaction_row(row: list[str]) -> bool:
    return len(row) >= 8 and bool(SERIAL_RE.match(row[0])) and normalize_date(row[1]) is not None


def _build_details(particulars: Any, channel: Any) -> str:
    details = clean_cell(particulars)
    channel_text = clean_cell(channel)
    if channel_text and channel_text.upper() not in details.upper():
        details = clean_cell(f"{details} {channel_text}")
    return details


def _build_record(row: list[str]) -> dict[str, Any]:
    details = _build_details(row[2], row[7])
    cheque_no = clean_cell(row[3])
    if cheque_no == "-":
        cheque_no = ""

    return {
        "Sno": 0,
        "Date": normalize_date(row[1]) or row[1],
        "Details": details,
        "Detail_Clean": clean_detail(details),
        "Cheque No": cheque_no,
        "Debit": parse_amount(row[4]),
        "Credit": parse_amount(row[5]),
        "Balance": parse_amount(row[6]),
    }


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Bank of Maharashtra statement: %s", pdf_path)

    by_serial: dict[int, dict[str, Any]] = {}
    expected_count: int | None = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("Page %s: extracted %s table(s)", page_idx, len(tables))

            for table_idx, table in enumerate(tables, start=1):
                logger.debug("Page %s table %s: rows=%s", page_idx, table_idx, len(table))
                if not table:
                    continue

                first_row = [clean_cell(cell) for cell in table[0]]
                if len(first_row) > 1 and first_row[0] == "Total Transaction Count":
                    parsed_count = parse_amount(first_row[1])
                    if parsed_count is not None:
                        expected_count = int(parsed_count)
                    continue

                for raw_row in table:
                    row = [clean_cell(cell) for cell in raw_row]
                    if not any(row) or row[0] == "Sr No" or not _is_transaction_row(row):
                        continue

                    serial_no = int(row[0])
                    if serial_no in by_serial:
                        continue

                    by_serial[serial_no] = _build_record(row)
                    if progress_cb is not None:
                        progress_cb(len(by_serial))

    sorted_serials = sorted(by_serial)
    if sorted_serials:
        expected_serials = set(range(sorted_serials[0], sorted_serials[-1] + 1))
        missing = sorted(expected_serials.difference(by_serial))
        if missing:
            logger.warning("BOM parsed rows missing serials: %s", missing[:25])

    records: list[dict[str, Any]] = []
    for index, serial_no in enumerate(sorted_serials, start=1):
        record = by_serial[serial_no]
        record["Sno"] = index
        records.append(record)

    if expected_count is not None:
        if expected_count != len(records):
            logger.warning(
                "BOM parsed row count mismatch: expected=%s actual=%s",
                expected_count,
                len(records),
            )
        else:
            logger.info("BOM parsed row count matches summary: %s", len(records))

    logger.info("Bank of Maharashtra parse complete: rows=%s", len(records))
    return records
