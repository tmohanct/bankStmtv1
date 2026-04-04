from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass, field
from datetime import datetime
from itertools import product
from pathlib import Path
from typing import Any

import fitz
import pdfplumber
import pytesseract
from PIL import Image

from utils import clean_cell, clean_detail, parse_amount

RENDER_ZOOM = 2.0
DATE_COL_MAX_LEFT = 180
DETAIL_COL_MIN_LEFT = 250
AMOUNT_COL_MIN_LEFT = 700
CREDIT_AMOUNT_MAX_LEFT = 860
BALANCE_COL_MIN_LEFT = 1000
TABLE_HEADER_RE = re.compile(r"^DATE\b.*\bBALANCE$", re.IGNORECASE)
DATE_TOKEN_RE = re.compile(r"^[0-9A-Za-z]{2}[-/][0-9A-Za-z]{2}[-/][0-9A-Za-z]{4}$")
AMOUNT_TOKEN_RE = re.compile(r"^[0-9OIlSBQDo,\.]+$")
NEXT_PREFIX_RE = re.compile(
    r"(UPI|UPV|NEFT|IMPS|RTGS|ATM|NACH|ACH|POS|CHEQ|CHQ|PAY|BANK|UNION|VYSA)",
    re.IGNORECASE,
)
TEXT_LINE_TOLERANCE = 3.0
TEXT_SERIAL_RE = re.compile(r"^\d+$")
TEXT_AMOUNT_RE = re.compile(r"^-?[0-9,]+\.\d{2}$")
TEXT_PAGE_FOOTER_RE = re.compile(r"^PAGE\s+\d+\s+OF\s+\d+$", re.IGNORECASE)
TEXT_FOOTER_PREFIXES = (
    "WWW.ICICI.BANK.IN",
    "PLEASE CALL FROM YOUR REGISTERED MOBILE NUMBER",
    "NEVER SHARE YOUR OTP",
    "SINCERLY,",
    "TEAM ICICI BANK",
    "LEGENDS FOR TRANSACTIONS IN YOUR ACCOUNT STATEMENT",
)
TEXT_DETAILED_STOP_PREFIXES = (
    "PAGE TOTAL",
    "OPENING BAL",
    "WITHDRAWLS:",
    "WITHDRAWALS:",
    "DEPOSITS:",
    "CLOSING BAL",
    "LEGENDS USED IN ACCOUNT STATEMENT",
    "----------END OF STATEMENT----------",
)
TEXT_DETAILED_DATE_MIN_LEFT = 170.0
TEXT_DETAILED_DATE_MAX_LEFT = 270.0
TEXT_DETAILED_DETAIL_MIN_LEFT = 300.0
TEXT_DETAILED_AMOUNT_MIN_LEFT = 390.0
TEXT_DETAILED_AMOUNT_MAX_LEFT = 470.0
TEXT_DETAILED_BALANCE_MIN_LEFT = 470.0
TEXT_DETAILED_CREDIT_MAX_LEFT = 425.0
TEXT_SUMMARY_BALANCE_PATTERNS = {
    "opening": re.compile(r"Opening\s+Bal:\s*-?\s*([0-9,]+\.\d{2})", re.IGNORECASE),
    "closing": re.compile(r"Closing\s+Bal:\s*-?\s*([0-9,]+\.\d{2})", re.IGNORECASE),
}
ICICI_CLG_CHEQUE_RE = re.compile(r"^CLG/(?:[^/]+/)+(?P<cheque>\d{4,7})(?:/|$)", re.IGNORECASE)
ICICI_REJECT_CHEQUE_RE = re.compile(r"^REJECT:(?P<cheque>\d{1,7})(?::|$)", re.IGNORECASE)
ICICI_RTN_CHG_CHEQUE_RE = re.compile(r"^RTN\s+CHG-\s*(?P<cheque>\d{1,7})(?:/|$)", re.IGNORECASE)
ICICI_GENERIC_CHQ_RE = re.compile(r"\b(?:CHQ|CHEQUE)\s*(?:NO\.?|NUMBER)?[:\-/ ]*(?P<cheque>\d{4,7})\b", re.IGNORECASE)
DATE_CHAR_OPTIONS = {
    "0": ("0",),
    "1": ("1",),
    "2": ("2",),
    "3": ("3",),
    "4": ("4", "1"),
    "5": ("5",),
    "6": ("6", "8"),
    "7": ("7", "1"),
    "8": ("8", "3"),
    "9": ("9", "3"),
    "O": ("0",),
    "Q": ("0",),
    "D": ("0",),
    "I": ("1",),
    "L": ("1",),
    "S": ("5",),
    "B": ("8",),
}


@dataclass
class OcrLine:
    top: int
    text: str
    tokens: list[tuple[int, str]]


@dataclass
class PendingRecord:
    raw_date: str
    amount_text: str | None
    amount_left: int | None
    balance_text: str | None
    detail_parts: list[str] = field(default_factory=list)


@dataclass
class TextLine:
    top: float
    tokens: list[tuple[float, str]]


@dataclass
class TextRecordSeed:
    top: float
    raw_date: str
    amount_text: str
    amount_left: float
    balance_text: str
    inline_detail: str
    prefix_parts: list[str] = field(default_factory=list)
    suffix_parts: list[str] = field(default_factory=list)


@dataclass
class DetailedTextRecordSeed:
    raw_date: str
    amount_text: str
    amount_left: float
    balance_text: str
    detail_text: str


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


def _render_page_image(page: fitz.Page) -> Image.Image:
    pixmap = page.get_pixmap(matrix=fitz.Matrix(RENDER_ZOOM, RENDER_ZOOM), alpha=False)
    return Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)


def _extract_page_lines(page: fitz.Page, page_number: int, logger) -> list[OcrLine]:
    image = _render_page_image(page)
    data = pytesseract.image_to_data(
        image,
        output_type=pytesseract.Output.DATAFRAME,
        config="--psm 6",
    )
    data = data.dropna(subset=["text"])
    data = data[data["text"].astype(str).str.strip() != ""].copy()

    lines: list[OcrLine] = []
    for _, group in data.groupby(["block_num", "par_num", "line_num"]):
        group = group.sort_values("left")
        tokens = [(int(row.left), clean_cell(row.text)) for row in group.itertuples(index=False)]
        tokens = [(left, text) for left, text in tokens if text]
        if not tokens:
            continue
        lines.append(
            OcrLine(
                top=int(group["top"].min()),
                text=" ".join(text for _, text in tokens),
                tokens=tokens,
            )
        )

    logger.debug("ICICI OCR page %s: extracted %s OCR line(s)", page_number, len(lines))
    return sorted(lines, key=lambda line: line.top)


def _extract_text_lines(page: pdfplumber.page.Page, page_number: int, logger) -> list[TextLine]:
    words = page.extract_words(
        x_tolerance=2,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )

    lines: list[TextLine] = []
    for word in sorted(words or [], key=lambda item: (float(item["top"]), float(item["x0"]))):
        text = clean_cell(word.get("text"))
        if not text:
            continue

        top = float(word["top"])
        left = float(word["x0"])
        if not lines or abs(lines[-1].top - top) > TEXT_LINE_TOLERANCE:
            lines.append(TextLine(top=top, tokens=[(left, text)]))
        else:
            lines[-1].tokens.append((left, text))

    logger.debug("ICICI text page %s: extracted %s text line(s)", page_number, len(lines))
    return lines


def _text_line_text(line: TextLine) -> str:
    return clean_cell(" ".join(text for _, text in line.tokens))


def _is_text_table_header(line: TextLine) -> bool:
    upper = _text_line_text(line).upper()
    return (
        "TRANSACTION" in upper
        and "BALANCE" in upper
        and "DEPOSIT" in upper
        and ("WITHDRAWAL" in upper or "WITHDRAWL" in upper)
    )


def _is_text_column_header(line: TextLine) -> bool:
    upper = _text_line_text(line).upper()
    return (
        upper.startswith("S NO")
        or upper.startswith("DATE AMOUNT")
        or upper == "DATE"
        or upper.startswith("NO ID DATE")
        or "REFNO" in upper
        or "REMARKS" in upper
        or "(DR)" in upper
        or "(CR)" in upper
    )


def _is_text_footer_line(line: TextLine) -> bool:
    text = _text_line_text(line)
    upper = text.upper()

    # Some statements place the page number on a standalone far-right token,
    # while detailed transaction rows can contain numeric-only continuation lines.
    if len(line.tokens) == 1 and upper.isdigit() and line.tokens[0][0] > 500:
        return True

    if TEXT_PAGE_FOOTER_RE.match(upper):
        return True

    return any(upper.startswith(prefix) for prefix in TEXT_FOOTER_PREFIXES)


def _is_detailed_stop_line(line: TextLine) -> bool:
    upper = _text_line_text(line).upper()
    return TEXT_PAGE_FOOTER_RE.match(upper) is not None or any(
        upper.startswith(prefix) for prefix in TEXT_DETAILED_STOP_PREFIXES
    )


def _detect_text_amount_layout(lines: list[TextLine]) -> tuple[float, float]:
    for line in lines:
        if not _is_text_table_header(line):
            continue

        withdrawal_left = next(
            (
                left
                for left, text in line.tokens
                if text.upper().startswith("WITHDRAWAL") or text.upper().startswith("WITHDRAWL")
            ),
            None,
        )
        deposit_left = next(
            (left for left, text in line.tokens if text.upper().startswith("DEPOSIT")),
            None,
        )
        if withdrawal_left is not None and deposit_left is not None:
            return withdrawal_left, (withdrawal_left + deposit_left) / 2.0

    return 395.0, 455.0


def _parse_text_date_token(raw_value: str) -> datetime | None:
    cleaned = clean_cell(raw_value)
    for fmt in ("%d.%m.%Y", "%d-%b-%Y", "%d-%B-%Y", "%d/%m/%Y"):
        try:
            return datetime.strptime(cleaned, fmt)
        except ValueError:
            continue
    return None


def _is_text_date_token(raw_value: str) -> bool:
    return _parse_text_date_token(raw_value) is not None

def _extract_text_row_seed(line: TextLine, detail_max_left: float) -> TextRecordSeed | None:
    tokens = [text for _, text in line.tokens]
    if len(tokens) < 4:
        return None
    if not TEXT_SERIAL_RE.match(tokens[0]):
        return None

    date_positions = [idx for idx, token in enumerate(tokens) if _is_text_date_token(token)]
    if not date_positions:
        return None
    row_date_idx = date_positions[1] if len(date_positions) > 1 else date_positions[0]

    amount_candidates = [
        (idx, left, text)
        for idx, (left, text) in enumerate(line.tokens[row_date_idx + 1 :], start=row_date_idx + 1)
        if TEXT_AMOUNT_RE.match(text)
    ]
    if len(amount_candidates) < 2:
        return None

    balance_idx, balance_left, balance_text = amount_candidates[-1]
    amount_idx, amount_left, amount_text = amount_candidates[-2]
    if amount_left >= balance_left:
        return None

    inline_detail = clean_cell(
        " ".join(
            text
            for left, text in line.tokens[row_date_idx + 1 : amount_idx]
            if left < detail_max_left and text.upper() != "NA"
        )
    )
    return TextRecordSeed(
        top=line.top,
        raw_date=tokens[row_date_idx],
        amount_text=amount_text,
        amount_left=amount_left,
        balance_text=balance_text,
        inline_detail=inline_detail,
    )


def _normalize_text_date(raw_value: str) -> str:
    parsed = _parse_text_date_token(raw_value)
    if parsed is not None:
        return parsed.strftime("%d/%m/%Y")
    return raw_value.replace(".", "/")


def _extract_icici_cheque_no(detail_text: str) -> str:
    text = clean_cell(detail_text)
    if not text:
        return ""

    for pattern in (
        ICICI_CLG_CHEQUE_RE,
        ICICI_REJECT_CHEQUE_RE,
        ICICI_RTN_CHG_CHEQUE_RE,
        ICICI_GENERIC_CHQ_RE,
    ):
        match = pattern.search(text)
        if match:
            return clean_cell(match.group("cheque"))

    return ""


def _parse_positive_amount(value: str | None) -> float | None:
    parsed = parse_amount(value)
    if parsed is None:
        return None
    return abs(parsed)


def _extract_statement_balance_summary(pdf_path: str, logger) -> dict[str, float]:
    with fitz.open(pdf_path) as pdf:
        text = "\n".join(page.get_text("text") or "" for page in pdf)

    balances: dict[str, float] = {}
    for key, pattern in TEXT_SUMMARY_BALANCE_PATTERNS.items():
        match = pattern.search(text)
        if not match:
            continue

        parsed = _parse_positive_amount(match.group(1))
        if parsed is not None:
            balances[key] = parsed

    logger.debug("ICICI statement balance summary: %s", balances)
    return balances


def _is_detailed_text_block_start(line: TextLine) -> bool:
    if not line.tokens:
        return False

    first_left, first_text = line.tokens[0]
    if first_left > 100 or not TEXT_SERIAL_RE.match(first_text):
        return False

    has_date = any(
        TEXT_DETAILED_DATE_MIN_LEFT <= left < TEXT_DETAILED_DATE_MAX_LEFT and _is_text_date_token(text)
        for left, text in line.tokens
    )
    has_amount = any(
        TEXT_DETAILED_AMOUNT_MIN_LEFT <= left < TEXT_DETAILED_AMOUNT_MAX_LEFT and _looks_like_amount_token(text)
        for left, text in line.tokens
    )
    return has_date and has_amount


def _extract_detailed_text_blocks(lines: list[TextLine]) -> list[list[TextLine]]:
    blocks: list[list[TextLine]] = []
    current_block: list[TextLine] = []

    for line in lines:
        if _is_detailed_stop_line(line):
            if current_block:
                blocks.append(current_block)
                current_block = []
            break

        if _is_text_footer_line(line):
            if current_block:
                blocks.append(current_block)
                current_block = []
            continue

        if _is_detailed_text_block_start(line):
            if current_block:
                blocks.append(current_block)
            current_block = [line]
            continue

        if current_block:
            current_block.append(line)

    if current_block:
        blocks.append(current_block)

    return blocks


def _merge_numeric_column_tokens(tokens: list[str]) -> str:
    return "".join(
        cleaned
        for cleaned in (clean_cell(token) for token in tokens)
        if cleaned and cleaned != "-"
    )


def _extract_detailed_text_record_seed(block_lines: list[TextLine]) -> DetailedTextRecordSeed | None:
    raw_date: str | None = None
    amount_tokens: list[str] = []
    balance_tokens: list[str] = []
    detail_tokens: list[str] = []
    amount_left: float | None = None

    for line in block_lines:
        for left, text in line.tokens:
            if (
                raw_date is None
                and TEXT_DETAILED_DATE_MIN_LEFT <= left < TEXT_DETAILED_DATE_MAX_LEFT
                and _is_text_date_token(text)
            ):
                raw_date = text

            if TEXT_DETAILED_AMOUNT_MIN_LEFT <= left < TEXT_DETAILED_AMOUNT_MAX_LEFT and _looks_like_amount_token(text):
                amount_tokens.append(text)
                amount_left = left if amount_left is None else min(amount_left, left)
                continue

            if left >= TEXT_DETAILED_BALANCE_MIN_LEFT and (_looks_like_amount_token(text) or text == "-"):
                balance_tokens.append(text)
                continue

            if TEXT_DETAILED_DETAIL_MIN_LEFT <= left < TEXT_DETAILED_AMOUNT_MIN_LEFT:
                detail_tokens.append(text)

    amount_text = _merge_numeric_column_tokens(amount_tokens)
    balance_text = _merge_numeric_column_tokens(balance_tokens)
    detail_text = clean_cell(" ".join(detail_tokens))

    if raw_date is None or not amount_text or not balance_text:
        return None

    return DetailedTextRecordSeed(
        raw_date=raw_date,
        amount_text=amount_text,
        amount_left=amount_left or TEXT_DETAILED_AMOUNT_MAX_LEFT,
        balance_text=balance_text,
        detail_text=detail_text,
    )


def _finalize_detailed_text_record(
    seed: DetailedTextRecordSeed,
    previous_balance: float | None,
    logger,
) -> tuple[dict[str, Any], float | None]:
    amount_value = _parse_positive_amount(seed.amount_text)
    balance_value = _parse_positive_amount(seed.balance_text)

    debit: float | None = None
    credit: float | None = None
    if amount_value is not None and balance_value is not None and previous_balance is not None:
        if abs((previous_balance + amount_value) - balance_value) <= 0.05:
            credit = amount_value
        elif abs((previous_balance - amount_value) - balance_value) <= 0.05:
            debit = amount_value

    if amount_value is not None and debit is None and credit is None:
        if seed.amount_left <= TEXT_DETAILED_CREDIT_MAX_LEFT:
            credit = amount_value
        else:
            debit = amount_value

    record = {
        "Sno": 0,
        "Date": _normalize_text_date(seed.raw_date),
        "Details": seed.detail_text,
        "Detail_Clean": clean_detail(seed.detail_text),
        "Cheque No": _extract_icici_cheque_no(seed.detail_text),
        "Debit": debit,
        "Credit": credit,
        "Balance": balance_value,
    }
    next_balance = balance_value if balance_value is not None else previous_balance

    logger.debug(
        "ICICI detailed text row parsed | date=%s debit=%s credit=%s balance=%s details=%s",
        record["Date"],
        debit,
        credit,
        balance_value,
        seed.detail_text,
    )
    return record, next_balance


def _parse_detailed_text(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing ICICI detailed statement with text extraction: %s", pdf_path)

    records: list[dict[str, Any]] = []
    summary_balances = _extract_statement_balance_summary(pdf_path, logger)
    previous_balance: float | None = summary_balances.get("opening")

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            lines = _extract_text_lines(page, page_idx, logger)
            if not lines:
                continue

            blocks = _extract_detailed_text_blocks(lines)
            logger.debug(
                "ICICI detailed text page %s: identified %s transaction block(s)",
                page_idx,
                len(blocks),
            )

            for block in blocks:
                seed = _extract_detailed_text_record_seed(block)
                if seed is None:
                    continue

                record, previous_balance = _finalize_detailed_text_record(
                    seed,
                    previous_balance,
                    logger,
                )
                records.append(record)
                if progress_cb is not None:
                    progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    closing_balance = summary_balances.get("closing")
    if records and closing_balance is not None:
        final_balance = records[-1].get("Balance")
        if isinstance(final_balance, (int, float)):
            if abs(final_balance - closing_balance) <= 0.05:
                logger.info("ICICI detailed text closing balance reconciled: %.2f", final_balance)
            else:
                logger.warning(
                    "ICICI detailed text closing balance mismatch: parsed=%.2f expected=%.2f",
                    final_balance,
                    closing_balance,
                )

    logger.info("ICICI detailed text parse complete: rows=%s", len(records))
    return records


def _finalize_text_record(
    seed: TextRecordSeed,
    previous_balance: float | None,
    credit_threshold: float,
    logger,
) -> tuple[dict[str, Any], float | None]:
    detail_parts = [part for part in seed.prefix_parts if clean_cell(part)]
    if seed.inline_detail:
        detail_parts.append(seed.inline_detail)
    detail_parts.extend(part for part in seed.suffix_parts if clean_cell(part))

    detail_text = clean_cell(" ".join(detail_parts))
    amount_value = parse_amount(seed.amount_text)
    balance_value = parse_amount(seed.balance_text)

    debit: float | None = None
    credit: float | None = None
    if amount_value is not None and balance_value is not None and previous_balance is not None:
        if abs((previous_balance + amount_value) - balance_value) <= 0.05:
            credit = amount_value
        elif abs((previous_balance - amount_value) - balance_value) <= 0.05:
            debit = amount_value

    if amount_value is not None and debit is None and credit is None:
        if seed.amount_left >= credit_threshold:
            credit = amount_value
        else:
            debit = amount_value

    record = {
        "Sno": 0,
        "Date": _normalize_text_date(seed.raw_date),
        "Details": detail_text,
        "Detail_Clean": clean_detail(detail_text),
        "Cheque No": _extract_icici_cheque_no(detail_text),
        "Debit": debit,
        "Credit": credit,
        "Balance": balance_value,
    }
    next_balance = balance_value if balance_value is not None else previous_balance

    logger.debug(
        "ICICI text row parsed | date=%s debit=%s credit=%s balance=%s details=%s",
        record["Date"],
        debit,
        credit,
        balance_value,
        detail_text,
    )
    return record, next_balance


def _parse_text(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing ICICI statement with text extraction: %s", pdf_path)

    records: list[dict[str, Any]] = []
    previous_balance: float | None = None
    detail_max_left = 395.0
    credit_threshold = 455.0

    with pdfplumber.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf.pages, start=1):
            lines = _extract_text_lines(page, page_idx, logger)
            if not lines:
                continue

            header_idx = next((idx for idx, line in enumerate(lines) if _is_text_table_header(line)), -1)
            if header_idx >= 0:
                detail_max_left, credit_threshold = _detect_text_amount_layout(lines)
                candidate_lines = lines[header_idx + 1 :]
            else:
                candidate_lines = lines

            relevant_lines = [
                line
                for line in candidate_lines
                if not _is_text_column_header(line) and not _is_text_footer_line(line)
            ]

            row_entries: list[tuple[int, TextRecordSeed]] = []
            for idx, line in enumerate(relevant_lines):
                seed = _extract_text_row_seed(line, detail_max_left)
                if seed is not None:
                    row_entries.append((idx, seed))

            logger.debug(
                "ICICI text page %s: identified %s transaction row(s)",
                page_idx,
                len(row_entries),
            )
            if not row_entries:
                continue

            first_row_idx, first_seed = row_entries[0]
            for line in relevant_lines[:first_row_idx]:
                text = _text_line_text(line)
                if text:
                    first_seed.prefix_parts.append(text)

            for entry_idx, (row_idx, seed) in enumerate(row_entries):
                if entry_idx + 1 < len(row_entries):
                    next_row_idx, next_seed = row_entries[entry_idx + 1]
                    for line in relevant_lines[row_idx + 1 : next_row_idx]:
                        text = _text_line_text(line)
                        if not text:
                            continue
                        seed.suffix_parts.append(text)
                else:
                    for line in relevant_lines[row_idx + 1 :]:
                        text = _text_line_text(line)
                        if text:
                            seed.suffix_parts.append(text)

                record, previous_balance = _finalize_text_record(
                    seed,
                    previous_balance,
                    credit_threshold,
                    logger,
                )
                records.append(record)
                if progress_cb is not None:
                    progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("ICICI text parse complete: rows=%s", len(records))
    return records

def _looks_like_amount_token(text: str) -> bool:
    return bool(AMOUNT_TOKEN_RE.match(text)) and any(ch.isdigit() for ch in text)


def _clean_ocr_amount_text(text: str) -> str:
    cleaned = text.upper()
    cleaned = (
        cleaned.replace("O", "0")
        .replace("D", "0")
        .replace("Q", "0")
        .replace("I", "1")
        .replace("L", "1")
        .replace("S", "5")
        .replace("B", "8")
    )
    cleaned = re.sub(r"[^0-9,\.]", "", cleaned)

    if cleaned.count(".") > 1:
        parts = cleaned.split(".")
        cleaned = "".join(parts[:-1]) + "." + parts[-1]

    digits_only = re.sub(r"[^0-9]", "", cleaned)
    if not digits_only:
        return ""

    if "." not in cleaned and len(digits_only) >= 3:
        cleaned = f"{digits_only[:-2]}.{digits_only[-2:]}"

    return cleaned

def _parse_ocr_amount(text: str | None) -> float | None:
    if not text:
        return None
    cleaned = _clean_ocr_amount_text(text)
    if not cleaned:
        return None
    return parse_amount(cleaned)


def _is_table_header(text: str) -> bool:
    return bool(TABLE_HEADER_RE.match(text.strip().upper()))


def _is_footer_line(text: str) -> bool:
    upper = text.upper()
    return upper.startswith("PAGE ")


def _normalize_date_token(token: str) -> str:
    return token.strip().lstrip("([{").rstrip(").,:;")


def _is_date_line(line: OcrLine) -> bool:
    if not line.tokens:
        return False
    left, token = line.tokens[0]
    normalized = _normalize_date_token(token)
    return left <= DATE_COL_MAX_LEFT and bool(DATE_TOKEN_RE.match(normalized.replace("/", "-")))

def _looks_like_next_prefix(text: str) -> bool:
    compact = text.strip()
    return len(compact) >= 20 and ("/" in compact or bool(NEXT_PREFIX_RE.search(compact)))


def _candidate_component_values(value: str) -> list[str]:
    option_sets = [DATE_CHAR_OPTIONS.get(ch.upper(), (ch,)) for ch in value]
    return sorted({"".join(chars) for chars in product(*option_sets)})


def _normalize_ocr_date(raw_value: str, previous_date: datetime | None) -> tuple[str, datetime | None]:
    sanitized = _normalize_date_token(raw_value).replace("/", "-")
    try:
        parsed = datetime.strptime(sanitized, "%d-%m-%Y")
        return parsed.strftime("%d/%m/%Y"), parsed
    except ValueError:
        pass

    parts = sanitized.split("-")
    if len(parts) != 3:
        return sanitized, None

    candidates: list[datetime] = []
    for day in _candidate_component_values(parts[0]):
        for month in _candidate_component_values(parts[1]):
            for year in _candidate_component_values(parts[2]):
                candidate_text = f"{day}-{month}-{year}"
                try:
                    candidates.append(datetime.strptime(candidate_text, "%d-%m-%Y"))
                except ValueError:
                    continue

    if not candidates:
        return sanitized, None

    if previous_date is not None:
        forward = [candidate for candidate in candidates if candidate >= previous_date]
        if forward:
            chosen = min(forward, key=lambda candidate: (candidate - previous_date).days)
        else:
            chosen = min(candidates, key=lambda candidate: abs((candidate - previous_date).days))
    else:
        chosen = min(candidates)

    return chosen.strftime("%d/%m/%Y"), chosen

def _extract_line_fields(line: OcrLine) -> tuple[str, str | None, int | None, str | None, str]:
    date_token = _normalize_date_token(line.tokens[0][1])

    numeric_tokens = [(left, text) for left, text in line.tokens[1:] if _looks_like_amount_token(text)]
    balance_left: int | None = None
    balance_text: str | None = None
    amount_left: int | None = None
    amount_text: str | None = None

    balance_candidates = [(left, text) for left, text in numeric_tokens if left >= BALANCE_COL_MIN_LEFT]
    if balance_candidates:
        balance_left, balance_text = max(balance_candidates, key=lambda item: item[0])

    amount_candidates = [
        (left, text)
        for left, text in numeric_tokens
        if left >= AMOUNT_COL_MIN_LEFT and (balance_left is None or left < balance_left)
    ]
    if amount_candidates:
        amount_left, amount_text = max(amount_candidates, key=lambda item: item[0])

    detail_tokens = []
    for left, text in line.tokens[1:]:
        if left < DETAIL_COL_MIN_LEFT:
            continue
        if amount_left is not None and left >= amount_left:
            continue
        if balance_left is not None and left >= balance_left:
            continue
        detail_tokens.append(text)

    detail_tail = clean_cell(" ".join(detail_tokens))
    return date_token, amount_text, amount_left, balance_text, detail_tail

def _finalize_ocr_record(
    pending: PendingRecord,
    previous_balance: float | None,
    previous_date: datetime | None,
    logger,
) -> tuple[dict[str, Any], float | None, datetime | None]:
    normalized_date, parsed_date = _normalize_ocr_date(pending.raw_date, previous_date)

    detail_text = clean_cell(" ".join(part for part in pending.detail_parts if clean_cell(part)))
    amount_value = _parse_ocr_amount(pending.amount_text)
    balance_value = _parse_ocr_amount(pending.balance_text)

    debit: float | None = None
    credit: float | None = None
    if amount_value is not None and pending.amount_left is not None:
        if pending.amount_left < CREDIT_AMOUNT_MAX_LEFT:
            credit = amount_value
        else:
            debit = amount_value

    if previous_balance is not None and amount_value is not None:
        expected_balance = round(previous_balance + (credit or 0.0) - (debit or 0.0), 2)
        if balance_value is None or abs(balance_value - expected_balance) > 0.05:
            logger.debug(
                "ICICI OCR balance corrected | raw=%s parsed=%s expected=%.2f details=%s",
                pending.balance_text,
                balance_value,
                expected_balance,
                detail_text,
            )
            balance_value = expected_balance

    record = {
        "Sno": 0,
        "Date": normalized_date,
        "Details": detail_text,
        "Detail_Clean": clean_detail(detail_text),
        "Cheque No": _extract_icici_cheque_no(detail_text),
        "Debit": debit,
        "Credit": credit,
        "Balance": balance_value,
    }
    next_date = parsed_date or previous_date
    next_balance = balance_value if balance_value is not None else previous_balance
    return record, next_balance, next_date

def _parse_ocr(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    tesseract_cmd = _configure_tesseract()
    logger.info("Parsing ICICI statement with OCR: %s | tesseract=%s", pdf_path, tesseract_cmd)

    records: list[dict[str, Any]] = []
    previous_balance: float | None = None
    previous_date: datetime | None = None

    with fitz.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            lines = _extract_page_lines(page, page_idx, logger)
            in_table = False
            pending_prefix_lines: list[str] = []
            current_record: PendingRecord | None = None

            for line in lines:
                text = clean_cell(line.text)
                if not text:
                    continue

                if not in_table:
                    if _is_table_header(text):
                        in_table = True
                    continue

                if _is_footer_line(text):
                    continue

                if _is_date_line(line):
                    if current_record is not None:
                        record, previous_balance, previous_date = _finalize_ocr_record(
                            current_record,
                            previous_balance,
                            previous_date,
                            logger,
                        )
                        records.append(record)
                        if progress_cb is not None:
                            progress_cb(len(records))

                    raw_date, amount_text, amount_left, balance_text, detail_tail = _extract_line_fields(line)
                    detail_parts = [part for part in pending_prefix_lines if clean_cell(part)]
                    if detail_tail:
                        detail_parts.append(detail_tail)
                    current_record = PendingRecord(
                        raw_date=raw_date,
                        amount_text=amount_text,
                        amount_left=amount_left,
                        balance_text=balance_text,
                        detail_parts=detail_parts,
                    )
                    pending_prefix_lines = []
                    continue

                if current_record is None:
                    pending_prefix_lines.append(text)
                    continue

                if _looks_like_next_prefix(text):
                    pending_prefix_lines.append(text)
                else:
                    current_record.detail_parts.append(text)

            if current_record is not None:
                record, previous_balance, previous_date = _finalize_ocr_record(
                    current_record,
                    previous_balance,
                    previous_date,
                    logger,
                )
                records.append(record)
                if progress_cb is not None:
                    progress_cb(len(records))

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("ICICI OCR parse complete: rows=%s", len(records))
    return records


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    detailed_text_records = _parse_detailed_text(pdf_path, logger, progress_cb)
    if detailed_text_records:
        return detailed_text_records

    text_records = _parse_text(pdf_path, logger, progress_cb)
    if text_records:
        return text_records

    logger.info("ICICI text parser found no rows. Falling back to OCR parser.")
    return _parse_ocr(pdf_path, logger, progress_cb)
