from __future__ import annotations

import sys
from pathlib import Path

SRC_ROOT = Path(__file__).resolve().parents[1]
if str(SRC_ROOT) not in sys.path:
    sys.path.append(str(SRC_ROOT))

from parser_helpers import build_record
from parsers.iob_parser import parse_iob_records

DATE_FORMATS = ("%d/%m/%Y",)


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, object]]:
    logger.info("Parsing IOB statement: %s", pdf_path)

    normalized_rows = parse_iob_records(pdf_path, progress_cb=progress_cb)
    records = [
        build_record(
            date_text=str(row.get("Date") or ""),
            details=str(row.get("Narration") or ""),
            cheque_no=str(row.get("Txn_Ref") or ""),
            debit=row.get("Debit"),
            credit=row.get("Credit"),
            balance=row.get("Balance"),
            date_formats=DATE_FORMATS,
        )
        for row in normalized_rows
    ]

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("IOB parse complete: rows=%s", len(records))
    return records
