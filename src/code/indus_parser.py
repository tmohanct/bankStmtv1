from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any

import pdfplumber

from parser_helpers import build_record
from utils import clean_cell, parse_amount

ROW_RE = re.compile(
    r"^(?P<date>\d{2} [A-Za-z]{3} \d{4})\s+"
    r"(?P<body>.+?)\s+"
    r"(?P<debit>-|[0-9,]+\.\d{2})\s+"
    r"(?P<credit>-|[0-9,]+\.\d{2})\s+"
    r"(?P<balance>-?[0-9,]+\.\d{2})$"
)
DATE_FORMATS = ("%d %b %Y",)
FOOTER_PREFIXES = (
    "Page ",
    "This is a computer generated statement",
    "For any queries",
    "IndusInd Bank",
)


@dataclass
class PendingRecord:
    date_text: str
    detail_head: str
    debit_text: str
    credit_text: str
    balance_text: str
    continuation_lines: list[str] = field(default_factory=list)


def _should_skip_line(line: str) -> bool:
    if not line:
        return True
    if line.startswith("Date Type Description Debit Credit Balance"):
        return True
    return any(line.startswith(prefix) for prefix in FOOTER_PREFIXES)


def _finalize_record(pending: PendingRecord) -> dict[str, Any]:
    details = clean_cell(" ".join([pending.detail_head, *pending.continuation_lines]))
    return build_record(
        date_text=pending.date_text,
        details=details,
        debit=parse_amount(pending.debit_text) if pending.debit_text != "-" else None,
        credit=parse_amount(pending.credit_text) if pending.credit_text != "-" else None,
        balance=parse_amount(pending.balance_text),
        date_formats=DATE_FORMATS,
    )


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing IndusInd statement: %s", pdf_path)

    records: list[dict[str, Any]] = []
    pending: PendingRecord | None = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            lines = (page.extract_text() or "").splitlines()
            logger.debug("Page %s: extracted %s text line(s)", page_idx, len(lines))

            for raw_line in lines:
                line = clean_cell(raw_line)
                if _should_skip_line(line):
                    continue

                match = ROW_RE.match(line)
                if match:
                    if pending is not None:
                        records.append(_finalize_record(pending))
                        if progress_cb is not None:
                            progress_cb(len(records))

                    pending = PendingRecord(
                        date_text=match.group("date"),
                        detail_head=match.group("body"),
                        debit_text=match.group("debit"),
                        credit_text=match.group("credit"),
                        balance_text=match.group("balance"),
                    )
                    continue

                if pending is not None:
                    pending.continuation_lines.append(line)

    if pending is not None:
        records.append(_finalize_record(pending))
        if progress_cb is not None:
            progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("IndusInd parse complete: rows=%s", len(records))
    return records
