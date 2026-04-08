from typing import Any

import pdfplumber

from utils import clean_cell, clean_detail, is_date_token, normalize_date, parse_amount

CUB_HEADER_DETAILS = {
    "PARTICULARS",
    "DESCRIPTION",
}


def _is_non_transaction_row(date_text: str, details_text: str) -> bool:
    normalized_date = clean_cell(date_text).upper()
    normalized_details = clean_cell(details_text).upper()

    if normalized_date == "DATE":
        return True
    if normalized_details in CUB_HEADER_DETAILS:
        return True
    if normalized_details.startswith("TOTAL"):
        return True
    if normalized_details in {"END OF REPORT", "AMT BROUGHT FORWARD :"}:
        return True
    return False


def _build_row(row: list[str]) -> dict[str, Any]:
    details = clean_cell(row[1]) if len(row) > 1 else ""
    cheque_no = clean_cell(row[2]) if len(row) > 2 else ""
    if cheque_no == "-":
        cheque_no = ""

    return {
        "Sno": 0,
        "Date": normalize_date(row[0]) or clean_cell(row[0]),
        "Details": details,
        "Detail_Clean": clean_detail(details),
        "Cheque No": cheque_no,
        "Debit": parse_amount(row[3]) if len(row) > 3 else None,
        "Credit": parse_amount(row[4]) if len(row) > 4 else None,
        "Balance": parse_amount(row[5]) if len(row) > 5 else None,
    }


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing CUB statement: %s", pdf_path)

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
                    details_text = row[1] if len(row) > 1 else ""

                    if _is_non_transaction_row(date_text, details_text):
                        continue

                    if is_date_token(date_text):
                        record = _build_row(row)
                        records.append(record)
                        if progress_cb is not None:
                            progress_cb(len(records))
                        continue

                    if records and details_text:
                        merged = f"{records[-1]['Details']} {details_text}".strip()
                        records[-1]["Details"] = merged
                        records[-1]["Detail_Clean"] = clean_detail(merged)

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("CUB parse complete: rows=%s", len(records))
    return records
