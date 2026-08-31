from __future__ import annotations

import os
import re
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

import fitz
import numpy as np
import pdfplumber
import pytesseract
from PIL import Image, ImageDraw

from utils import (
    clean_cell,
    clean_detail,
    extract_cheque_number_from_details,
    normalize_cheque_number,
    normalize_date,
    parse_amount,
)

TEXT_DATE_RE = re.compile(r"^(?P<day>\d{1,2})\s+(?P<month>[A-Za-z]{3,9})\s+(?P<year>\d{4})$")
OCR_DAY_FIRST_DATE_RE = re.compile(
    r"(?P<day>\d{1,2})\s*(?P<month>[A-Za-z]{3,9})\s*(?P<year>\d{4})",
    re.IGNORECASE,
)
OCR_MONTH_FIRST_DATE_RE = re.compile(
    r"(?P<month>[A-Za-z]{3,9})\s*(?P<day>\d{1,2})\s*(?P<year>\d{4})",
    re.IGNORECASE,
)
OCR_RENDER_ZOOM = 2.0
OCR_DARK_PIXEL_THRESHOLD = 160
OCR_GRID_ERASE_HALF_WIDTH = 5
OCR_HORIZONTAL_SCAN_LEFT_RATIO = 0.05
OCR_HORIZONTAL_SCAN_RIGHT_RATIO = 0.95
OCR_MIN_HORIZONTAL_LINE_RATIO = 0.70
OCR_COLUMN_BOUNDARY_RATIOS = (
    0.0,
    0.1012,
    0.2025,
    0.3539,
    0.5052,
    0.6060,
    0.7270,
    0.8490,
    1.0,
)
OCR_AMOUNT_TOLERANCE = 0.05
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
        compact = re.sub(r"[^0-9A-Za-z]+", " ", text).strip()
        match = OCR_DAY_FIRST_DATE_RE.search(compact)
        if not match:
            match = OCR_MONTH_FIRST_DATE_RE.search(compact)
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
        "This SBI statement is scanned and requires Tesseract OCR. "
        "Install Tesseract or set TESSERACT_CMD to the full executable path."
    )


def _render_page_image(page: fitz.Page) -> Image.Image:
    pixmap = page.get_pixmap(
        matrix=fitz.Matrix(OCR_RENDER_ZOOM, OCR_RENDER_ZOOM),
        alpha=False,
    )
    return Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)


def _group_adjacent_positions(positions: np.ndarray) -> list[list[int]]:
    groups: list[list[int]] = []
    for raw_position in positions:
        position = int(raw_position)
        if not groups or position > groups[-1][-1] + 1:
            groups.append([position])
        else:
            groups[-1].append(position)
    return groups


def _longest_dark_run(row: np.ndarray) -> tuple[int, int] | None:
    dark_positions = np.flatnonzero(row < OCR_DARK_PIXEL_THRESHOLD)
    if dark_positions.size == 0:
        return None

    runs: list[list[int]] = []
    for raw_position in dark_positions:
        position = int(raw_position)
        if not runs or position > runs[-1][-1] + 2:
            runs.append([position])
        else:
            runs[-1].append(position)

    longest = max(runs, key=lambda run: run[-1] - run[0])
    return longest[0], longest[-1]


def _detect_table_grid(image: Image.Image) -> tuple[list[int], list[int]] | None:
    gray = np.asarray(image.convert("L"))
    scan_left = int(image.width * OCR_HORIZONTAL_SCAN_LEFT_RATIO)
    scan_right = int(image.width * OCR_HORIZONTAL_SCAN_RIGHT_RATIO)
    scan_width = max(1, scan_right - scan_left)

    horizontal_counts = (
        gray[:, scan_left:scan_right] < OCR_DARK_PIXEL_THRESHOLD
    ).sum(axis=1)
    candidate_positions = np.flatnonzero(
        horizontal_counts >= scan_width * OCR_MIN_HORIZONTAL_LINE_RATIO
    )
    line_groups = _group_adjacent_positions(candidate_positions)
    horizontal_lines = [
        int(round(sum(group) / len(group)))
        for group in line_groups
        if group
    ]
    if len(horizontal_lines) < 3:
        return None

    spans: list[tuple[int, int]] = []
    for y_position in horizontal_lines:
        span = _longest_dark_run(gray[y_position])
        if span is None:
            continue
        if span[1] - span[0] >= scan_width * OCR_MIN_HORIZONTAL_LINE_RATIO:
            spans.append(span)
    if not spans:
        return None

    table_left = int(round(float(np.median([span[0] for span in spans]))))
    table_right = int(round(float(np.median([span[1] for span in spans]))))
    table_width = table_right - table_left
    if table_width <= 0:
        return None

    column_boundaries = [
        int(round(table_left + (table_width * ratio)))
        for ratio in OCR_COLUMN_BOUNDARY_RATIOS
    ]
    return horizontal_lines, column_boundaries


def _erase_table_grid(
    image: Image.Image,
    horizontal_lines: list[int],
    column_boundaries: list[int],
) -> Image.Image:
    cleaned = image.copy()
    draw = ImageDraw.Draw(cleaned)
    half_width = OCR_GRID_ERASE_HALF_WIDTH
    table_top = horizontal_lines[0]
    table_bottom = horizontal_lines[-1]

    for y_position in horizontal_lines:
        draw.rectangle(
            (0, y_position - half_width, cleaned.width, y_position + half_width),
            fill="white",
        )

    for x_position in column_boundaries:
        draw.rectangle(
            (x_position - half_width, table_top, x_position + half_width, table_bottom),
            fill="white",
        )

    return cleaned


def _extract_ocr_tokens(image: Image.Image) -> list[tuple[float, float, int, int, str]]:
    data = pytesseract.image_to_data(
        image,
        output_type=pytesseract.Output.DICT,
        config="--psm 6",
    )
    tokens: list[tuple[float, float, int, int, str]] = []
    for index, raw_text in enumerate(data.get("text", [])):
        text = clean_cell(raw_text)
        if not text:
            continue

        left = int(data["left"][index])
        top = int(data["top"][index])
        width = int(data["width"][index])
        height = int(data["height"][index])
        tokens.append(
            (
                left + (width / 2),
                top + (height / 2),
                top,
                left,
                text,
            )
        )
    return tokens


def _extract_ocr_row_cells(
    tokens: list[tuple[float, float, int, int, str]],
    row_top: int,
    row_bottom: int,
    column_boundaries: list[int],
) -> list[str]:
    cells: list[str] = []
    for column_left, column_right in zip(column_boundaries, column_boundaries[1:]):
        cell_tokens = [
            token
            for token in tokens
            if row_top < token[1] < row_bottom
            and column_left < token[0] < column_right
        ]
        cell_tokens.sort(key=lambda token: (token[2], token[3]))
        cells.append(clean_cell(" ".join(token[4] for token in cell_tokens)))
    return cells


def _clean_ocr_text(value: Any) -> str:
    text = clean_cell(value)
    text = text.replace("|", " ").replace("{", " ").replace("}", " ")
    text = clean_cell(text)
    compact = re.sub(r"[^A-Za-z0-9]", "", text)
    if re.fullmatch(r"K{3,}", compact, re.IGNORECASE):
        return ""
    return text


def _parse_ocr_amount(value: Any) -> float | None:
    text = clean_cell(value).upper()
    if not text:
        return None

    text = (
        text.replace("O", "0")
        .replace("Q", "0")
        .replace("D", "0")
        .replace("I", "1")
        .replace("L", "1")
        .replace("S", "5")
        .replace("B", "8")
    )
    cleaned = re.sub(r"[^0-9,.()\-]", "", text)
    if cleaned.count(".") > 1:
        parts = cleaned.split(".")
        cleaned = "".join(parts[:-1]) + "." + parts[-1]
    return parse_amount(cleaned)


def _extract_ocr_cheque_no(reference: str, details: str) -> str:
    numeric_reference = re.sub(r"[^0-9]", "", clean_cell(reference))
    cheque_no = normalize_cheque_number(numeric_reference, details)
    if cheque_no:
        return cheque_no
    return extract_cheque_number_from_details(details, clean_detail(details))


def _build_ocr_record(
    cells: list[str],
    previous_balance: float | None,
    logger,
) -> dict[str, Any] | None:
    if len(cells) < 8:
        return None

    normalized_date = _normalize_sbi_date(cells[0])
    if normalized_date is None:
        return None

    details = _clean_ocr_text(cells[2])
    reference = _clean_ocr_text(cells[3])
    debit = _parse_ocr_amount(cells[5])
    credit = _parse_ocr_amount(cells[6])
    balance = _parse_ocr_amount(cells[7])

    if previous_balance is not None and balance is not None:
        delta = round(balance - previous_balance, 2)
        if delta > OCR_AMOUNT_TOLERANCE:
            if credit is None or abs(credit - delta) > OCR_AMOUNT_TOLERANCE or debit is not None:
                logger.debug(
                    "SBI OCR credit corrected from balance movement | raw_debit=%s raw_credit=%s expected=%.2f details=%s",
                    debit,
                    credit,
                    delta,
                    details,
                )
                debit = None
                credit = delta
        elif delta < -OCR_AMOUNT_TOLERANCE:
            expected_debit = abs(delta)
            if debit is None or abs(debit - expected_debit) > OCR_AMOUNT_TOLERANCE or credit is not None:
                logger.debug(
                    "SBI OCR debit corrected from balance movement | raw_debit=%s raw_credit=%s expected=%.2f details=%s",
                    debit,
                    credit,
                    expected_debit,
                    details,
                )
                debit = expected_debit
                credit = None

    if balance is None and previous_balance is not None:
        if debit is not None and credit is None:
            balance = round(previous_balance - debit, 2)
        elif credit is not None and debit is None:
            balance = round(previous_balance + credit, 2)

    return {
        "Sno": 0,
        "Date": normalized_date,
        "Details": details,
        "Detail_Clean": clean_detail(details),
        "Cheque No": _extract_ocr_cheque_no(reference, details),
        "Debit": debit,
        "Credit": credit,
        "Balance": balance,
    }


def _parse_ocr(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    tesseract_cmd = _configure_tesseract()
    logger.info("Parsing scanned SBI statement with OCR: %s | tesseract=%s", pdf_path, tesseract_cmd)

    records: list[dict[str, Any]] = []
    previous_balance: float | None = None

    with fitz.open(pdf_path) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            image = _render_page_image(page)
            grid = _detect_table_grid(image)
            if grid is None:
                logger.debug("SBI OCR page %s: transaction grid not found", page_idx)
                continue

            horizontal_lines, column_boundaries = grid
            cleaned_image = _erase_table_grid(image, horizontal_lines, column_boundaries)
            tokens = _extract_ocr_tokens(cleaned_image)
            page_count_before = len(records)

            for row_top, row_bottom in zip(horizontal_lines, horizontal_lines[1:]):
                cells = _extract_ocr_row_cells(
                    tokens,
                    row_top,
                    row_bottom,
                    column_boundaries,
                )
                record = _build_ocr_record(cells, previous_balance, logger)
                if record is None:
                    continue

                records.append(record)
                if record["Balance"] is not None:
                    previous_balance = float(record["Balance"])
                if progress_cb is not None:
                    progress_cb(len(records))

            logger.debug(
                "SBI OCR page %s: parsed %s transaction row(s)",
                page_idx,
                len(records) - page_count_before,
            )

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("SBI OCR parse complete: rows=%s", len(records))
    return records


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

    if records:
        logger.info("SBI text-table parse complete: rows=%s", len(records))
        return records

    logger.info("SBI text-table parser found no rows. Falling back to OCR parser.")
    return _parse_ocr(pdf_path, logger, progress_cb)
