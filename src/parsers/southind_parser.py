"""South Indian Bank parser implementation."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

import fitz
import pandas as pd

from parsers.base_parser import BaseStatementParser

TRANSACTION_DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{2}$")
SLNO_LAYOUT_DATE_RE = re.compile(r"^\d{2}-[A-Za-z]{3}-\d{4}$")
SLNO_LAYOUT_SERIAL_RE = re.compile(r"^\d{1,6}$")
PAGE_TOTAL_RE = re.compile(r"^Page\s+Total\b", re.IGNORECASE)

DATE_COLUMN_MAX_X = 80
DETAIL_COLUMN_MIN_X = 80
DETAIL_COLUMN_MAX_X = 250
CHEQUE_COLUMN_MIN_X = 250
CHEQUE_COLUMN_MAX_X = 335
WITHDRAWAL_COLUMN_MIN_X = 335
WITHDRAWAL_COLUMN_MAX_X = 445
DEPOSIT_COLUMN_MIN_X = 445
# Legacy statements place balances around x=513, while newer variants use x=520+.
# Keep the boundary below both layouts without overlapping deposits near x=445-465.
DEPOSIT_COLUMN_MAX_X = 500
LINE_GROUP_TOLERANCE = 2.5

SLNO_SERIAL_COLUMN_MAX_X = 35
SLNO_TRANSACTION_DATE_MIN_X = 35
SLNO_TRANSACTION_DATE_MAX_X = 105
SLNO_DETAIL_COLUMN_MIN_X = 175
SLNO_CHEQUE_COLUMN_MIN_X = 280
SLNO_WITHDRAWAL_COLUMN_MIN_X = 340
SLNO_DEPOSIT_COLUMN_MIN_X = 430
SLNO_BALANCE_COLUMN_MIN_X = 520
SLNO_LEADING_CONTINUATION_GAP = 14.0


def _clean_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\r", " ").replace("\n", " ")).strip()


def _normalize_output_date(value: str) -> str:
    text = _clean_text(value)
    if not text:
        return ""

    for fmt in ("%d-%m-%y", "%d-%m-%Y", "%d-%b-%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return text


def _clean_detail_key(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", _clean_text(value))


def _parse_amount(value: str) -> float | None:
    text = _clean_text(value)
    if not text:
        return None

    upper = text.upper().rstrip(".")
    negative = text.startswith("-") or upper.endswith("DR")
    cleaned = re.sub(r"\s*(CR|DR)\.?$", "", text, flags=re.IGNORECASE)
    cleaned = cleaned.replace(",", "").strip()
    if not cleaned:
        return None
    if cleaned.count(".") > 1:
        decimal_index = cleaned.rfind(".")
        cleaned = cleaned[:decimal_index].replace(".", "") + cleaned[decimal_index:]

    try:
        amount = float(cleaned)
    except ValueError:
        return None

    if negative and amount > 0:
        amount = -amount
    return amount


@dataclass
class _WordLine:
    y_center: float
    words: list[tuple[float, str]] = field(default_factory=list)

    @property
    def text(self) -> str:
        return _clean_text(" ".join(text for _, text in sorted(self.words)))


@dataclass
class _PendingRecord:
    date_text: str
    detail_lines: list[str] = field(default_factory=list)
    cheque_parts: list[str] = field(default_factory=list)
    debit: float | None = None
    credit: float | None = None
    balance: float | None = None

    def add_line(self, line: _WordLine) -> None:
        detail_words = [text for x, text in line.words if DETAIL_COLUMN_MIN_X <= x < DETAIL_COLUMN_MAX_X]
        cheque_words = [text for x, text in line.words if CHEQUE_COLUMN_MIN_X <= x < CHEQUE_COLUMN_MAX_X]
        debit_words = [text for x, text in line.words if WITHDRAWAL_COLUMN_MIN_X <= x < WITHDRAWAL_COLUMN_MAX_X]
        credit_words = [text for x, text in line.words if DEPOSIT_COLUMN_MIN_X <= x < DEPOSIT_COLUMN_MAX_X]
        balance_words = [text for x, text in line.words if x >= DEPOSIT_COLUMN_MAX_X]

        if detail_words:
            self.detail_lines.append(_clean_text(" ".join(detail_words)))
        if cheque_words:
            self.cheque_parts.append(_clean_text(" ".join(cheque_words)))
        if self.debit is None and debit_words:
            self.debit = _parse_amount(" ".join(debit_words))
        if self.credit is None and credit_words:
            self.credit = _parse_amount(" ".join(credit_words))
        if self.balance is None and balance_words:
            self.balance = _parse_amount(" ".join(balance_words))

    def finalize(self) -> dict[str, Any] | None:
        details = _clean_text(" ".join(part for part in self.detail_lines if part))
        cheque_joined = _clean_text(" ".join(part for part in self.cheque_parts if part))
        cheque_matches = re.findall(r"\d{3,}", cheque_joined)
        cheque_no = cheque_matches[0] if cheque_matches else cheque_joined

        if not details or self.balance is None or (self.debit is None and self.credit is None):
            return None

        return {
            "Sno": 0,
            "Date": _normalize_output_date(self.date_text),
            "Details": details,
            "Detail_Clean": _clean_detail_key(details),
            "Cheque No": cheque_no,
            "Debit": abs(self.debit) if self.debit is not None else None,
            "Credit": abs(self.credit) if self.credit is not None else None,
            "Balance": self.balance,
        }


def _build_word_lines(page: fitz.Page) -> list[_WordLine]:
    words = sorted(page.get_text("words"), key=lambda item: (((item[1] + item[3]) / 2), item[0]))
    lines: list[_WordLine] = []
    current_line: _WordLine | None = None

    for x0, y0, x1, y1, text, *_ in words:
        cleaned = _clean_text(text)
        if not cleaned:
            continue

        y_center = (y0 + y1) / 2
        if current_line is None or abs(y_center - current_line.y_center) > LINE_GROUP_TOLERANCE:
            current_line = _WordLine(y_center=y_center)
            lines.append(current_line)
        current_line.words.append((x0, cleaned))

    return lines


def _parse_page_lines(
    lines: list[_WordLine],
    records: list[dict[str, Any]],
    progress_cb: Callable[[int], None] | None = None,
) -> None:
    in_transaction_table = False
    pending: _PendingRecord | None = None

    for line in lines:
        line_text = line.text
        if not line_text:
            continue

        if "DATE" in line_text and "PARTICULARS" in line_text and "BALANCE" in line_text:
            in_transaction_table = True
            continue

        if not in_transaction_table:
            continue

        if PAGE_TOTAL_RE.match(line_text):
            if pending is not None:
                record = pending.finalize()
                if record is not None:
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))
            break

        if line_text.startswith("Page ") or line_text.startswith("Visit us at"):
            continue

        date_words = [text for x, text in line.words if x < DATE_COLUMN_MAX_X]
        date_text = _clean_text(" ".join(date_words))
        if TRANSACTION_DATE_RE.match(date_text):
            if pending is not None:
                record = pending.finalize()
                if record is not None:
                    records.append(record)
                    if progress_cb is not None:
                        progress_cb(len(records))

            pending = _PendingRecord(date_text=date_text)
            pending.add_line(line)
            continue

        if pending is not None:
            pending.add_line(line)

    else:
        if pending is not None:
            record = pending.finalize()
            if record is not None:
                records.append(record)
            if progress_cb is not None:
                progress_cb(len(records))


def _is_slno_layout_header(line_text: str) -> bool:
    normalized = line_text.upper()
    return (
        "SLNO" in normalized
        and "TRANSACTION" in normalized
        and "PARTICULARS" in normalized
        and "BALANCE" in normalized
    )


def _find_slno_layout_header_y(lines: list[_WordLine]) -> float | None:
    for line in lines:
        if _is_slno_layout_header(line.text):
            return line.y_center
    return None


def _is_slno_layout_end_marker(line_text: str) -> bool:
    normalized = line_text.upper()
    return "END" in normalized and "STATEMENT" in normalized


def _extract_slno_transaction_date(line: _WordLine) -> str:
    for x, text in sorted(line.words):
        if (
            SLNO_TRANSACTION_DATE_MIN_X <= x < SLNO_TRANSACTION_DATE_MAX_X
            and SLNO_LAYOUT_DATE_RE.match(text)
        ):
            return text
    return ""


def _line_has_slno_anchor(line: _WordLine) -> bool:
    for x, text in sorted(line.words):
        if x < SLNO_SERIAL_COLUMN_MAX_X and SLNO_LAYOUT_SERIAL_RE.match(text):
            return True
        if (
            SLNO_TRANSACTION_DATE_MIN_X <= x < SLNO_DETAIL_COLUMN_MIN_X
            and SLNO_LAYOUT_DATE_RE.match(text)
        ):
            return True
    return False


def _extract_slno_region_date(region_lines: list[_WordLine]) -> str:
    for line in region_lines:
        date_text = _extract_slno_transaction_date(line)
        if date_text:
            return date_text

    for line in region_lines:
        for x, text in sorted(line.words):
            if (
                SLNO_TRANSACTION_DATE_MAX_X <= x < SLNO_DETAIL_COLUMN_MIN_X
                and SLNO_LAYOUT_DATE_RE.match(text)
            ):
                return text
    return ""


def _extract_slno_anchors(lines: list[_WordLine], start_y: float, end_y: float) -> list[tuple[float, int]]:
    anchors: list[tuple[float, int]] = []
    for line_idx, line in enumerate(lines):
        if not (start_y < line.y_center < end_y):
            continue
        for x, text in sorted(line.words):
            if x < SLNO_SERIAL_COLUMN_MAX_X and SLNO_LAYOUT_SERIAL_RE.match(text):
                anchors.append((line.y_center, line_idx))
                break
    return anchors


def _first_amount_in_column(
    region_lines: list[_WordLine],
    min_x: float,
    max_x: float | None = None,
) -> float | None:
    for line in region_lines:
        for x, text in sorted(line.words):
            if x < min_x:
                continue
            if max_x is not None and x >= max_x:
                continue
            amount = _parse_amount(text)
            if amount is not None:
                return amount
    return None


def _extract_slno_details(region_lines: list[_WordLine]) -> str:
    detail_parts: list[str] = []
    for line in region_lines:
        if _is_slno_layout_header(line.text) or _is_slno_layout_end_marker(line.text):
            continue
        for x, text in sorted(line.words):
            if not (SLNO_DETAIL_COLUMN_MIN_X <= x < SLNO_WITHDRAWAL_COLUMN_MIN_X):
                continue
            if x >= SLNO_CHEQUE_COLUMN_MIN_X and re.fullmatch(r"\d{3,}", text):
                continue
            detail_parts.append(text)
    return _clean_text(" ".join(detail_parts))


def _extract_slno_cheque_no(region_lines: list[_WordLine]) -> str:
    cheque_parts: list[str] = []
    for line in region_lines:
        for x, text in sorted(line.words):
            if SLNO_CHEQUE_COLUMN_MIN_X <= x < SLNO_WITHDRAWAL_COLUMN_MIN_X:
                cheque_parts.append(text)

    cheque_joined = _clean_text(" ".join(cheque_parts))
    cheque_matches = re.findall(r"\d{3,}", cheque_joined)
    return cheque_matches[0] if cheque_matches else ""


def _build_slno_layout_record(region_lines: list[_WordLine]) -> dict[str, Any] | None:
    date_text = _extract_slno_region_date(region_lines)
    details = _extract_slno_details(region_lines)
    debit = _first_amount_in_column(
        region_lines,
        SLNO_WITHDRAWAL_COLUMN_MIN_X,
        SLNO_DEPOSIT_COLUMN_MIN_X,
    )
    credit = _first_amount_in_column(
        region_lines,
        SLNO_DEPOSIT_COLUMN_MIN_X,
        SLNO_BALANCE_COLUMN_MIN_X,
    )
    balance = _first_amount_in_column(region_lines, SLNO_BALANCE_COLUMN_MIN_X)

    if not date_text or not details or balance is None or (debit is None and credit is None):
        return None

    return {
        "Sno": 0,
        "Date": _normalize_output_date(date_text),
        "Details": details,
        "Detail_Clean": _clean_detail_key(details),
        "Cheque No": _extract_slno_cheque_no(region_lines),
        "Debit": abs(debit) if debit is not None else None,
        "Credit": abs(credit) if credit is not None else None,
        "Balance": balance,
    }


def _move_leading_continuations_to_previous(regions: list[list[_WordLine]]) -> None:
    for region_idx in range(1, len(regions)):
        region = regions[region_idx]
        if len(region) < 2:
            continue

        first_anchor_idx = next(
            (idx for idx, line in enumerate(region) if _line_has_slno_anchor(line)),
            None,
        )
        if not first_anchor_idx:
            continue

        leading_lines = region[:first_anchor_idx]
        gap_to_anchor = region[first_anchor_idx].y_center - leading_lines[-1].y_center
        if gap_to_anchor <= SLNO_LEADING_CONTINUATION_GAP:
            continue

        regions[region_idx - 1].extend(leading_lines)
        regions[region_idx - 1].sort(key=lambda line: line.y_center)
        del region[:first_anchor_idx]


def _parse_slno_layout_lines(
    lines: list[_WordLine],
    records: list[dict[str, Any]],
    progress_cb: Callable[[int], None] | None = None,
) -> None:
    if not lines:
        return

    header_y = _find_slno_layout_header_y(lines)
    start_y = header_y if header_y is not None else -1.0
    end_y = max(line.y_center for line in lines) + 1.0
    for line in lines:
        if _is_slno_layout_end_marker(line.text):
            end_y = min(end_y, line.y_center)

    anchors = _extract_slno_anchors(lines, start_y=start_y, end_y=end_y)
    if not anchors:
        return

    anchor_ys = [anchor_y for anchor_y, _ in anchors]
    regions: list[list[_WordLine]] = []
    for anchor_idx, (anchor_y, _) in enumerate(anchors):
        if anchor_idx == 0:
            region_start = start_y
        else:
            region_start = (anchor_ys[anchor_idx - 1] + anchor_y) / 2

        if anchor_idx + 1 < len(anchor_ys):
            region_end = (anchor_y + anchor_ys[anchor_idx + 1]) / 2
        else:
            region_end = end_y

        region_lines = [
            line
            for line in lines
            if region_start < line.y_center < region_end
            and not _is_slno_layout_end_marker(line.text)
        ]
        regions.append(region_lines)

    _move_leading_continuations_to_previous(regions)

    for region_lines in regions:
        record = _build_slno_layout_record(region_lines)
        if record is None:
            continue

        records.append(record)
        if progress_cb is not None:
            progress_cb(len(records))


def parse_southind_records(
    pdf_path: str | Path,
    logger: logging.Logger,
    progress_cb: Callable[[int], None] | None = None,
) -> list[dict[str, Any]]:
    logger.info("Parsing South Indian Bank statement: %s", pdf_path)

    records: list[dict[str, Any]] = []
    slno_layout_seen = False
    with fitz.open(str(pdf_path)) as pdf:
        for page_idx, page in enumerate(pdf, start=1):
            lines = _build_word_lines(page)
            logger.debug("SouthInd page %s: grouped %s word line(s)", page_idx, len(lines))
            if _find_slno_layout_header_y(lines) is not None:
                slno_layout_seen = True

            if slno_layout_seen:
                _parse_slno_layout_lines(lines, records, progress_cb=progress_cb)
            else:
                _parse_page_lines(lines, records, progress_cb=progress_cb)

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("South Indian Bank parse complete: rows=%s", len(records))
    return records


class SouthIndianParser(BaseStatementParser):
    """South Indian Bank statement parser."""

    bank_code = "southind"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        _ = rules_df
        logger = logging.getLogger(__name__)
        records = parse_southind_records(pdf_path=pdf_path, logger=logger)

        rows = [
            {
                "Date": record["Date"],
                "Value_Date": record["Date"],
                "Description": record["Details"],
                "Debit": record["Debit"],
                "Credit": record["Credit"],
                "Balance": record["Balance"],
                "Reference": record["Cheque No"],
                "Source_Page": None,
            }
            for record in records
        ]

        return pd.DataFrame(
            rows,
            columns=[
                "Date",
                "Value_Date",
                "Description",
                "Debit",
                "Credit",
                "Balance",
                "Reference",
                "Source_Page",
            ],
        )
