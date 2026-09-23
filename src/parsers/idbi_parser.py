"""Bank-specific transaction extraction and first-page identity hints."""

from __future__ import annotations

import re
from typing import Any
import pdfplumber
from src.transform.cheque_returns import is_cheque_return
from src.utils.statement_utils import clean_cell, clean_detail, normalize_date, parse_amount


# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {
    "name": "IDBI Bank",
    "ifsc": "IBKL",
    "aliases": [
        "IDBI Bank"
    ]
}


TRANSACTION_LINE_RE = re.compile(
    r"^(?P<srl>\d+)\s+"
    r"(?P<txn_ts>(?:\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[AP]M|\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3}))\s+"
    r"(?P<value_date>\d{1,2}/\d{1,2}/\d{4})\s+"
    r"(?P<body>.+?)\s+"
    r"(?P<drcr>Dr\.|Cr\.)\s+"
    r"INR\s+"
    r"(?P<amount>-?[0-9,]+\.\d{2}|-)\s+"
    r"(?P<balance>-?[0-9,]+\.\d{2}|-)$"
)
RETURN_NOTICE_LINE_RE = re.compile(
    r"^(?P<srl>\d+)\s+"
    r"(?P<txn_ts>(?:\d{1,2}/\d{1,2}/\d{4}\s+\d{1,2}:\d{2}:\d{2}\s+[AP]M|\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3}))\s+"
    r"(?P<value_date>\d{1,2}/\d{1,2}/\d{4})\s+"
    r"(?P<body>.+)$"
)

CHEQUE_TRAILER_RE = re.compile(r"^(?P<details>.+?)\s+(?P<cheque>\d{5,7})$")


def _parse_transaction_line(line: str) -> dict[str, Any] | None:
    text = clean_cell(line)
    if not text:
        return None

    match = TRANSACTION_LINE_RE.match(text)
    if not match:
        match = RETURN_NOTICE_LINE_RE.match(text)
        if not match or not is_cheque_return(match.group("body")):
            return None

    serial_no = int(match.group("srl"))
    value_date_raw = match.group("value_date")
    value_date = normalize_date(value_date_raw) or value_date_raw

    body = clean_cell(match.group("body"))
    cheque_no = ""
    details = body

    cheque_match = CHEQUE_TRAILER_RE.match(body)
    if cheque_match:
        details = clean_cell(cheque_match.group("details"))
        cheque_no = cheque_match.group("cheque")

    amount = parse_amount(match.groupdict().get("amount"))
    balance = parse_amount(match.groupdict().get("balance"))
    if amount is None and not is_cheque_return(details, cheque_no):
        return None

    drcr = (match.groupdict().get("drcr") or "").upper()
    abs_amount = abs(amount) if amount is not None else None

    debit: float | None = None
    credit: float | None = None
    if abs_amount is not None:
        if drcr.startswith("DR"):
            debit = abs_amount
        elif drcr.startswith("CR"):
            credit = abs_amount

    return {
        "serial_no": serial_no,
        "Sno": 0,
        "Date": value_date,
        "Details": details,
        "Detail_Clean": clean_detail(details),
        "Cheque No": cheque_no,
        "Debit": debit,
        "Credit": credit,
        "Balance": balance,
    }


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing IDBI statement: %s", pdf_path)

    by_serial: dict[int, dict[str, Any]] = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            text = page.extract_text() or ""
            if not text:
                continue

            lines = text.splitlines()
            logger.debug("Page %s: extracted %s text line(s)", page_idx, len(lines))

            for line in lines:
                row = _parse_transaction_line(line)
                if row is None:
                    continue

                serial_no = int(row.pop("serial_no"))
                if serial_no in by_serial:
                    continue

                by_serial[serial_no] = row
                if progress_cb is not None:
                    progress_cb(len(by_serial))

    sorted_serials = sorted(by_serial)

    if sorted_serials:
        expected = set(range(sorted_serials[0], sorted_serials[-1] + 1))
        missing = sorted(expected.difference(by_serial))
        if missing:
            logger.warning("IDBI parsed rows missing serials: %s", missing[:25])
        logger.info(
            "IDBI text parse summary: rows=%s serial_range=%s-%s",
            len(sorted_serials),
            sorted_serials[0],
            sorted_serials[-1],
        )

    records: list[dict[str, Any]] = []
    for index, serial_no in enumerate(sorted_serials, start=1):
        row = by_serial[serial_no]
        row["Sno"] = index
        records.append(row)

    return records


BANK_CODE = 'idbi'
BANK_SIGNATURES = (('IDBI BANK', 4), ('IBKL0', 3), ('YOUR A/C STATUS', 1))
