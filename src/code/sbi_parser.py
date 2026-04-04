import re
from datetime import datetime
from typing import Any

import pdfplumber

from utils import clean_cell, clean_detail, normalize_date, parse_amount

TEXT_DATE_RE = re.compile(r"^(?P<day>\d{1,2})\s+(?P<month>[A-Za-z]{3,9})\s+(?P<year>\d{4})$")
MONTH_MAP = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12,
}


def _normalize_sbi_date(value: Any) -> str | None:
    text = clean_cell(value)
    if not text:
        return None

    normalized = normalize_date(text)
    if normalized is not None:
        return normalized

    match = TEXT_DATE_RE.match(text)
    if not match:
        return None

    month = MONTH_MAP.get(match.group("month")[:3].lower())
    if month is None:
        return None

    try:
        parsed = datetime(int(match.group("year")), month, int(match.group("day")))
    except ValueError:
        return None

    return parsed.strftime("%d/%m/%Y")


def _is_sbi_date_token(value: Any) -> bool:
    return _normalize_sbi_date(value) is not None


def _build_row(row: list[str]) -> dict[str, Any]:
    details = clean_cell(row[2]) if len(row) > 2 else ""
    cheque_no = clean_cell(row[3]) if len(row) > 3 else ""
    if cheque_no == "-":
        cheque_no = ""

    return {
        "Sno": 0,
        "Date": _normalize_sbi_date(row[0]) or clean_cell(row[0]),
        "Details": details,
        "Detail_Clean": clean_detail(details),
        "Cheque No": cheque_no,
        "Debit": parse_amount(row[4]) if len(row) > 4 else None,
        "Credit": parse_amount(row[5]) if len(row) > 5 else None,
        "Balance": parse_amount(row[6]) if len(row) > 6 else None,
    }


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing SBI statement: %s", pdf_path)

    records: list[dict[str, Any]] = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            tables = page.extract_tables() or []
            logger.debug("Page %s: extracted %s table(s)", page_idx, len(tables))

            for table in tables:
                for raw_row in table:
                    row = [clean_cell(cell) for cell in raw_row]
                    if not any(row):
                        continue

                    date_text = row[0] if row else ""
                    details_text = row[2] if len(row) > 2 else ""

                    if _is_sbi_date_token(date_text):
                        record = _build_row(row)
                        records.append(record)
                        if progress_cb is not None:
                            progress_cb(len(records))
                        continue

                    if records and not date_text and details_text and details_text.upper() != "BALANCE":
                        merged = f"{records[-1]['Details']} {details_text}".strip()
                        records[-1]["Details"] = merged
                        records[-1]["Detail_Clean"] = clean_detail(merged)

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("SBI parse complete: rows=%s", len(records))
    return records
