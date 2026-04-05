from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import fitz
import pytesseract
from PIL import Image

from parser_helpers import build_record
from utils import clean_cell, parse_amount

RENDER_ZOOM = 2.0
OCR_ROW_RE = re.compile(
    r"^(?P<date>\d{2}-\d{2}-\d{4})\s+"
    r"(?P<time>\d{2}:\d{2}:\d{2})\s+"
    r"(?P<value_date>\d{2}-\d{2}-\d{4})\s+"
    r"(?P<body>.+?)\s+"
    r"(?P<amount>[0-9,]+\.\d{2})\s+"
    r"(?P<balance>-?[0-9,]+\.\d{2}[\]\)\}]?)$"
)
OPENING_BALANCE_RE = re.compile(r"Opening Balance.*?(-?[0-9,]+\.\d{2})")
OCR_DATE_FORMATS = ("%d-%m-%Y", "%d/%m/%Y")
TEXT_DATE_FORMATS = ("%d/%m/%y",)
TEXT_ROW_RE = re.compile(
    r"^(?P<date>\d{2}/\d{2}/\d{2})\s+"
    r"(?P<value_date>\d{2}/\d{2}/\d{2})\s*"
    r"(?P<rest>.*)$"
)
TEXT_AMOUNT_RE = re.compile(r"-?(?:\d{1,3}(?:,\d{2,3})*|\d+)?\.\d{2}")
TEXT_DETECTION_MIN_ROWS = 3
TOKENIZED_TEXT_ROW_START_RE = re.compile(r"^\d{2}[-/]\d{2}[-/]\d{4}\s+\d{2}:\d{2}(?::\d{2})?$")
TOKENIZED_TEXT_DATE_RE = re.compile(r"^\d{2}[-/]\d{2}[-/]\d{4}$")
TOKENIZED_TEXT_TIME_RE = re.compile(r"^\d{2}:\d{2}(?::\d{2})?$")
TOKENIZED_TEXT_AMOUNT_RE = re.compile(
    r"^-?\(?[0-9,]+\.\d{2}\)?(?:CR|DR|MR|[\]\)\}])?$",
    re.IGNORECASE,
)
OCR_FOOTER_PREFIXES = ("Page No.",)
OCR_HEADER_PREFIXES = (
    "Account Statement",
    "as of ",
    "Account Name ",
    "Account Holder(s) Name",
    "Account Number ",
    "Branch ",
    "Customer Id ",
    "Account Currency ",
    "Opening Balance ",
    "Closing Balance ",
    "Searched by ",
    "From Date ",
    "To Date ",
    "Transaction Date Value Date",
)
CREDIT_HINTS = (
    "CASH DEP",
    "DEPOSIT",
    "NEFT CR",
    "RTGS CR",
    "BY CLG",
    "CR-",
    "CREDIT",
    "REVERSAL",
    "B/F",
)
DEBIT_HINTS = (
    "CHQ PAID",
    "WITHDRAWL",
    "WITHDRAWAL",
    "CHARGES",
    "SMS CHARGES",
    "BILLDESK",
    "DEBIT",
    "IMPS-",
    "SBIEPAY",
    "TO DESIGN",
    "TO DESIG",
    "TR TO ",
    "FT - DR",
    "FT -100",
)
TOKENIZED_TEXT_BREAK_PREFIXES = (
    "Karur Vysya Bank does not ask",
    "Never disclose your passwords",
    "Account Statement",
    "THE KARUR VYSYA BANK LTD.",
    "Acc.No.",
    "Customer ID",
    "Acc.Type",
    "St.Date",
    "St.Period",
    "Mobile No.",
    "Email Id",
    "Account Summary",
    "Opening Balance",
    "Total Credit Amount",
    "Total Debit Amount",
    "Closing Balance",
    "Count of Cr. & Dr. Transactions",
)
KVB_CHEQUE_HINT_RE = re.compile(r"\b(?:CHQ|CHEQ(?:UE)?|CLG|CLEARING|CTS|RETURN)\b", re.IGNORECASE)
KVB_CHEQUE_TRAILER_RE = re.compile(r"^(?P<details>.+?)\s+(?P<cheque>0\d{5,}|\d{6,})$")


@dataclass
class PendingRecord:
    date_text: str
    body_text: str
    amount_text: str
    balance_text: str
    continuation_lines: list[str] = field(default_factory=list)
    cheque_no: str = ""


def _configure_tesseract() -> str:
    candidates: list[Path] = []

    env_value = os.environ.get("TESSERACT_CMD")
    if env_value:
        candidates.append(Path(env_value))

    resolved = shutil.which("tesseract")
    if resolved:
        candidates.append(Path(resolved))

    candidates.extend(
        [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            pytesseract.pytesseract.tesseract_cmd = str(candidate)
            return str(candidate)

    raise FileNotFoundError(
        "Tesseract executable not found. Install Tesseract OCR or set TESSERACT_CMD to the full path."
    )


def _render_page_text(page: fitz.Page) -> str:
    pixmap = page.get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM), alpha=False)
    image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
    return pytesseract.image_to_string(image, config="--psm 6")


def _clean_ocr_line(raw_line: str) -> str:
    return clean_cell(raw_line)


def _should_skip_ocr_line(line: str) -> bool:
    if not line:
        return True
    if any(line.startswith(prefix) for prefix in OCR_FOOTER_PREFIXES):
        return True
    return any(line.startswith(prefix) for prefix in OCR_HEADER_PREFIXES)


def _split_ocr_body(body_text: str) -> tuple[str, str]:
    text = clean_cell(body_text).replace("|", " ")
    parts = [part for part in text.split() if part]
    if parts and parts[0].isdigit() and len(parts[0]) == 4:
        parts.pop(0)

    cheque_no = ""
    if parts and re.fullmatch(r"[0-9$]{3,}", parts[0]):
        cheque_no = parts.pop(0)

    details_text = clean_cell(" ".join(parts))
    if not cheque_no:
        details_text, cheque_no = _extract_trailing_cheque_no(details_text)

    return details_text, cheque_no


def _extract_trailing_cheque_no(body_text: str) -> tuple[str, str]:
    text = clean_cell(body_text)
    if not text or not KVB_CHEQUE_HINT_RE.search(text):
        return text, ""

    match = KVB_CHEQUE_TRAILER_RE.match(text)
    if match is None:
        return text, ""

    cheque_no = match.group("cheque")
    if not cheque_no.strip("0"):
        return text, ""

    return clean_cell(match.group("details")), cheque_no


def _classify_amount(
    details: str,
    amount_value: float,
    balance_value: float,
    previous_balance: float | None,
) -> tuple[float | None, float | None]:
    if previous_balance is not None:
        if abs((previous_balance + amount_value) - balance_value) <= 1.0:
            return None, amount_value
        if abs((previous_balance - amount_value) - balance_value) <= 1.0:
            return amount_value, None

    upper_details = details.upper()
    if any(token in upper_details for token in CREDIT_HINTS) and not any(
        token in upper_details for token in DEBIT_HINTS
    ):
        return None, amount_value
    return amount_value, None


def _finalize_record(
    pending: PendingRecord,
    previous_balance: float | None,
    date_formats: tuple[str, ...],
) -> tuple[dict[str, Any], float | None]:
    details_text = clean_cell(" ".join([pending.body_text, *pending.continuation_lines]))
    amount_value = parse_amount(pending.amount_text)
    balance_value = parse_amount(pending.balance_text)

    debit: float | None = None
    credit: float | None = None
    if amount_value is not None and balance_value is not None:
        debit, credit = _classify_amount(details_text, amount_value, balance_value, previous_balance)

    record = build_record(
        date_text=pending.date_text,
        details=details_text,
        cheque_no=pending.cheque_no,
        debit=debit,
        credit=credit,
        balance=balance_value,
        date_formats=date_formats,
    )
    next_balance = balance_value if balance_value is not None else previous_balance
    return record, next_balance


def _finalize_tokenized_pending_record(
    current_row_lines: list[str],
    records: list[dict[str, Any]],
    previous_balance: float | None,
    progress_cb=None,
) -> tuple[list[str], float | None]:
    if not current_row_lines:
        return [], previous_balance

    parsed = _parse_tokenized_text_row(current_row_lines)
    if parsed is None:
        return [], previous_balance

    record, next_balance = _finalize_record(parsed, previous_balance, OCR_DATE_FORMATS)
    records.append(record)
    if progress_cb is not None:
        progress_cb(len(records))
    return [], next_balance


def _looks_like_tokenized_row_start(line: str, next_line: str = "") -> bool:
    if TOKENIZED_TEXT_ROW_START_RE.match(line):
        return True
    if not TOKENIZED_TEXT_DATE_RE.match(line):
        return False
    return bool(TOKENIZED_TEXT_TIME_RE.match(next_line) or TOKENIZED_TEXT_DATE_RE.match(next_line))


def _is_tokenized_text_break_line(line: str) -> bool:
    if not line:
        return True
    if line.startswith("Page No.") or line.lower().startswith("page :"):
        return True
    if re.fullmatch(r"Page \d+ of \d+", line, flags=re.IGNORECASE):
        return True
    return any(line.startswith(prefix) for prefix in TOKENIZED_TEXT_BREAK_PREFIXES)


def _parse_ocr_statement(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    tesseract_cmd = _configure_tesseract()
    logger.info("Parsing KVB statement with OCR layout: %s | tesseract=%s", pdf_path, tesseract_cmd)

    inline_records: list[dict[str, Any]] = []
    inline_previous_balance: float | None = None
    inline_pending: PendingRecord | None = None

    tokenized_records: list[dict[str, Any]] = []
    tokenized_previous_balance: float | None = None
    current_row_lines: list[str] = []

    with fitz.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            text = _render_page_text(page)
            lines = text.splitlines()
            logger.debug("Page %s: OCR extracted %s line(s)", page_idx, len(lines))

            cleaned_lines = [_clean_ocr_line(raw_line) for raw_line in lines]
            for line in cleaned_lines:
                if not line:
                    continue

                opening_match = OPENING_BALANCE_RE.search(line)
                if opening_match:
                    opening_balance = parse_amount(opening_match.group(1))
                    if inline_previous_balance is None:
                        inline_previous_balance = opening_balance
                    if tokenized_previous_balance is None:
                        tokenized_previous_balance = opening_balance

            usable_lines = [line for line in cleaned_lines if line and not _should_skip_ocr_line(line)]

            for idx, line in enumerate(usable_lines):
                next_line = usable_lines[idx + 1] if idx + 1 < len(usable_lines) else ""
                if _looks_like_tokenized_row_start(line, next_line):
                    current_row_lines, tokenized_previous_balance = _finalize_tokenized_pending_record(
                        current_row_lines,
                        tokenized_records,
                        tokenized_previous_balance,
                        progress_cb=progress_cb,
                    )
                    current_row_lines = [line]
                    continue

                if current_row_lines:
                    current_row_lines.append(line)

                match = OCR_ROW_RE.match(line)
                if match:
                    if inline_pending is not None:
                        record, inline_previous_balance = _finalize_record(
                            inline_pending,
                            inline_previous_balance,
                            OCR_DATE_FORMATS,
                        )
                        inline_records.append(record)
                        if progress_cb is not None:
                            progress_cb(len(inline_records))

                    details, cheque_no = _split_ocr_body(match.group("body"))
                    inline_pending = PendingRecord(
                        date_text=match.group("date"),
                        body_text=details,
                        amount_text=match.group("amount"),
                        balance_text=match.group("balance"),
                        cheque_no=cheque_no,
                    )
                    continue

                if inline_pending is not None:
                    inline_pending.continuation_lines.append(line)

    if inline_pending is not None:
        record, inline_previous_balance = _finalize_record(inline_pending, inline_previous_balance, OCR_DATE_FORMATS)
        inline_records.append(record)
        if progress_cb is not None:
            progress_cb(len(inline_records))

    current_row_lines, tokenized_previous_balance = _finalize_tokenized_pending_record(
        current_row_lines,
        tokenized_records,
        tokenized_previous_balance,
        progress_cb=progress_cb,
    )

    selected_records = inline_records
    selected_balance = inline_previous_balance
    selected_mode = "inline"
    if len(tokenized_records) > len(inline_records):
        selected_records = tokenized_records
        selected_balance = tokenized_previous_balance
        selected_mode = "tokenized"

    logger.info(
        "KVB OCR parse candidates: inline_rows=%s tokenized_rows=%s selected=%s",
        len(inline_records),
        len(tokenized_records),
        selected_mode,
    )

    for index, record in enumerate(selected_records, start=1):
        record["Sno"] = index

    logger.info("KVB OCR parse complete: rows=%s ending_balance=%s", len(selected_records), selected_balance)
    return selected_records




def _is_separator_line(line: str) -> bool:
    return bool(line) and set(line) == {"-"}


def _should_skip_text_line(
    line: str,
    in_header_block: bool,
    saw_table_header: bool,
    in_summary_block: bool,
) -> tuple[bool, bool, bool, bool]:
    if not line:
        return True, in_header_block, saw_table_header, in_summary_block

    if in_summary_block:
        return True, in_header_block, saw_table_header, in_summary_block

    if line.startswith("Opening Balance"):
        return True, False, False, True

    if line.startswith("THE KARUR VYSYA BANK LTD."):
        return True, True, False, False

    if in_header_block:
        if line.startswith("TXN DT"):
            return True, True, True, False
        if saw_table_header and _is_separator_line(line):
            return True, False, False, False
        return True, True, saw_table_header, False


    if _is_separator_line(line):
        return True, in_header_block, saw_table_header, in_summary_block

    if line.lower().startswith("page :"):
        return True, in_header_block, saw_table_header, in_summary_block

    if re.fullmatch(r"Page \d+ of \d+", line, flags=re.IGNORECASE):
        return True, in_header_block, saw_table_header, in_summary_block

    if re.fullmatch(r"\d{2}/\d{2}/\d{4}", line):
        return True, in_header_block, saw_table_header, in_summary_block

    if line.startswith("https://"):
        return True, in_header_block, saw_table_header, in_summary_block

    return False, in_header_block, saw_table_header, in_summary_block



def _parse_text_row(line: str) -> dict[str, str] | None:
    match = TEXT_ROW_RE.match(line)
    if not match:
        return None

    rest = clean_cell(match.group("rest"))
    if not rest:
        return None

    parts = rest.split()
    if parts and parts[0].isdigit() and len(parts[0]) <= 4:
        rest = clean_cell(rest[len(parts[0]) :])

    amount_matches = list(TEXT_AMOUNT_RE.finditer(rest))
    if not amount_matches:
        return None

    balance_text = amount_matches[-1].group(0)
    amount_text = amount_matches[-2].group(0) if len(amount_matches) >= 2 else ""
    body_end = amount_matches[-2].start() if len(amount_matches) >= 2 else amount_matches[-1].start()
    body_text = clean_cell(rest[:body_end])


    body_text, cheque_no = _extract_trailing_cheque_no(body_text)

    return {
        "date_text": match.group("date"),
        "body_text": body_text,
        "amount_text": amount_text,
        "balance_text": balance_text,
        "cheque_no": cheque_no,
    }


def _detect_text_layout(pdf_path: str) -> bool:
    row_hits = 0
    with fitz.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            for raw_line in page.get_text("text").splitlines():
                if TEXT_ROW_RE.match(clean_cell(raw_line)):
                    row_hits += 1
                    if row_hits >= TEXT_DETECTION_MIN_ROWS:
                        return True
            if page_idx >= 3:
                break
    return False


def _detect_tokenized_text_layout(pdf_path: str) -> bool:
    row_hits = 0
    with fitz.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            significant_lines = [
                clean_cell(raw_line)
                for raw_line in page.get_text("text").splitlines()
                if clean_cell(raw_line)
            ]
            for idx, line in enumerate(significant_lines):
                next_line = significant_lines[idx + 1] if idx + 1 < len(significant_lines) else ""
                if _looks_like_tokenized_row_start(line, next_line):
                    row_hits += 1
                    if row_hits >= TEXT_DETECTION_MIN_ROWS:
                        return True
            if page_idx >= 3:
                break
    return False


def _extract_opening_balance_from_text(page_text: str) -> float | None:
    compact = clean_cell(page_text)
    match = OPENING_BALANCE_RE.search(compact)
    if not match:
        return None
    return parse_amount(match.group(1))


def _parse_tokenized_text_row(row_lines: list[str]) -> PendingRecord | None:
    normalized_lines = [clean_cell(line) for line in row_lines if clean_cell(line)]
    if len(normalized_lines) < 4:
        return None

    date_text = ""
    cursor = 0
    if TOKENIZED_TEXT_ROW_START_RE.match(normalized_lines[0]):
        date_text = normalized_lines[0].split()[0]
        cursor = 1
    elif TOKENIZED_TEXT_DATE_RE.match(normalized_lines[0]):
        date_text = normalized_lines[0]
        cursor = 1
        if cursor < len(normalized_lines) and TOKENIZED_TEXT_TIME_RE.match(normalized_lines[cursor]):
            cursor += 1
    else:
        return None

    if cursor >= len(normalized_lines) or not TOKENIZED_TEXT_DATE_RE.match(normalized_lines[cursor]):
        return None
    cursor += 1

    amount_positions = [
        idx
        for idx, value in enumerate(normalized_lines[cursor:], start=cursor)
        if TOKENIZED_TEXT_AMOUNT_RE.match(value)
    ]
    if len(amount_positions) < 2:
        return None

    amount_idx = amount_positions[-2]
    balance_idx = amount_positions[-1]
    if amount_idx >= balance_idx:
        return None

    amount_text = normalized_lines[amount_idx]
    balance_text = normalized_lines[balance_idx]
    body_parts = [part for part in normalized_lines[cursor:amount_idx] if clean_cell(part)]
    if body_parts and body_parts[0].isdigit() and len(body_parts[0]) <= 4:
        body_parts.pop(0)

    cheque_no = ""
    if body_parts and re.fullmatch(r"\d{6,}", body_parts[0]):
        cheque_no = body_parts.pop(0)

    body_text = clean_cell(" ".join(body_parts))
    if not cheque_no:
        body_text, cheque_no = _extract_trailing_cheque_no(body_text)

    return PendingRecord(
        date_text=date_text,
        body_text=body_text,
        amount_text=amount_text,
        balance_text=balance_text,
        cheque_no=cheque_no,
    )


def _parse_tokenized_text_statement(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing KVB statement with tokenized text layout: %s", pdf_path)

    records: list[dict[str, Any]] = []
    previous_balance: float | None = None
    current_row_lines: list[str] = []

    def finalize_current_row() -> None:
        nonlocal current_row_lines, previous_balance
        if not current_row_lines:
            return

        parsed = _parse_tokenized_text_row(current_row_lines)
        current_row_lines = []
        if parsed is None:
            return

        record, previous_balance_local = _finalize_record(parsed, previous_balance, OCR_DATE_FORMATS)
        previous_balance = previous_balance_local
        records.append(record)
        if progress_cb is not None:
            progress_cb(len(records))

    with fitz.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            page_text = page.get_text("text")
            raw_lines = page_text.splitlines()
            logger.debug("Page %s: tokenized text extracted %s line(s)", page_idx, len(raw_lines))

            if previous_balance is None:
                opening_balance = _extract_opening_balance_from_text(page_text)
                if opening_balance is not None:
                    previous_balance = opening_balance

            significant_lines = [clean_cell(raw_line) for raw_line in raw_lines if clean_cell(raw_line)]

            for idx, line in enumerate(significant_lines):
                next_line = significant_lines[idx + 1] if idx + 1 < len(significant_lines) else ""
                if _looks_like_tokenized_row_start(line, next_line):
                    finalize_current_row()
                    current_row_lines = [line]
                    continue

                if _is_tokenized_text_break_line(line):
                    finalize_current_row()
                    continue

                if current_row_lines:
                    current_row_lines.append(line)

            # KVB rows do not continue across pages; closing the row here avoids
            # the next page header/summary being appended to the previous row.
            finalize_current_row()

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("KVB tokenized text parse complete: rows=%s ending_balance=%s", len(records), previous_balance)
    return records


def _parse_text_statement(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing KVB statement with native text layout: %s", pdf_path)

    records: list[dict[str, Any]] = []
    previous_balance: float | None = None
    pending: PendingRecord | None = None
    in_header_block = False
    saw_table_header = False
    in_summary_block = False

    with fitz.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            lines = page.get_text("text").splitlines()
            logger.debug("Page %s: native text extracted %s line(s)", page_idx, len(lines))

            for raw_line in lines:
                line = clean_cell(raw_line)
                should_skip, in_header_block, saw_table_header, in_summary_block = _should_skip_text_line(
                    line,
                    in_header_block,
                    saw_table_header,
                    in_summary_block,
                )
                if should_skip:
                    continue

                parsed = _parse_text_row(line)
                if parsed is not None:
                    if pending is not None:
                        record, previous_balance = _finalize_record(
                            pending,
                            previous_balance,
                            TEXT_DATE_FORMATS,
                        )
                        records.append(record)
                        if progress_cb is not None:
                            progress_cb(len(records))


                    pending = PendingRecord(**parsed)
                    continue

                if pending is not None:
                    pending.continuation_lines.append(line)

    if pending is not None:
        record, previous_balance = _finalize_record(pending, previous_balance, TEXT_DATE_FORMATS)
        records.append(record)
        if progress_cb is not None:
            progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("KVB native text parse complete: rows=%s ending_balance=%s", len(records), previous_balance)
    return records


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    if _detect_tokenized_text_layout(pdf_path):
        logger.info("Detected tokenized text KVB layout for %s", pdf_path)
        return _parse_tokenized_text_statement(pdf_path, logger, progress_cb=progress_cb)

    if _detect_text_layout(pdf_path):
        logger.info("Detected native text KVB layout for %s", pdf_path)
        return _parse_text_statement(pdf_path, logger, progress_cb=progress_cb)

    logger.info("Detected OCR KVB layout for %s", pdf_path)
    return _parse_ocr_statement(pdf_path, logger, progress_cb=progress_cb)
