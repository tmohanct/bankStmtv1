"""ICICI-specific statement layout parsers.

This module contains the canonical parser for the ICICI savings-statement
layout whose columns are DATE, MODE, PARTICULARS, DEPOSITS, WITHDRAWALS, and
BALANCE. The compatibility CLI in ``src/code`` delegates this layout here.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pdfplumber

from src.utils.amount_utils import parse_amount
from src.utils.text_utils import clean_cell, clean_detail

DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{4}$")
AMOUNT_RE = re.compile(r"^-?(?:\d{1,3}(?:,\d{2,3})*|\d+)\.\d{2}$")
LINE_TOP_TOLERANCE = 3.0


@dataclass(frozen=True)
class SavingsTextToken:
    left: float
    text: str


@dataclass
class SavingsTextLine:
    top: float
    tokens: list[SavingsTextToken]


@dataclass(frozen=True)
class SavingsHeaderLayout:
    top: float
    date_left: float
    mode_left: float
    detail_left: float
    deposit_left: float
    withdrawal_left: float
    balance_left: float

    @property
    def detail_amount_boundary(self) -> float:
        return (self.detail_left + self.deposit_left) / 2.0

    @property
    def deposit_withdrawal_boundary(self) -> float:
        return (self.deposit_left + self.withdrawal_left) / 2.0

    @property
    def withdrawal_balance_boundary(self) -> float:
        return (self.withdrawal_left + self.balance_left) / 2.0


def _extract_lines(page: pdfplumber.page.Page) -> list[SavingsTextLine]:
    words = page.extract_words(
        x_tolerance=2,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=False,
    )

    lines: list[SavingsTextLine] = []
    for word in sorted(words or [], key=lambda item: (float(item["top"]), float(item["x0"]))):
        text = clean_cell(word.get("text"))
        if not text:
            continue

        top = float(word["top"])
        token = SavingsTextToken(left=float(word["x0"]), text=text)
        if not lines or abs(lines[-1].top - top) > LINE_TOP_TOLERANCE:
            lines.append(SavingsTextLine(top=top, tokens=[token]))
        else:
            lines[-1].tokens.append(token)

    for line in lines:
        line.tokens.sort(key=lambda token: token.left)
    return lines


def _token_left(line: SavingsTextLine, prefix: str) -> float | None:
    upper_prefix = prefix.upper()
    return next(
        (token.left for token in line.tokens if token.text.upper().startswith(upper_prefix)),
        None,
    )


def _find_header(lines: list[SavingsTextLine]) -> tuple[int, SavingsHeaderLayout] | None:
    for index, line in enumerate(lines):
        date_left = _token_left(line, "DATE")
        mode_left = _token_left(line, "MODE")
        detail_left = _token_left(line, "PARTICULARS")
        deposit_left = _token_left(line, "DEPOSITS")
        withdrawal_left = _token_left(line, "WITHDRAWALS")
        balance_left = _token_left(line, "BALANCE")
        positions = (
            date_left,
            mode_left,
            detail_left,
            deposit_left,
            withdrawal_left,
            balance_left,
        )
        if all(position is not None for position in positions):
            assert date_left is not None
            assert mode_left is not None
            assert detail_left is not None
            assert deposit_left is not None
            assert withdrawal_left is not None
            assert balance_left is not None
            return index, SavingsHeaderLayout(
                top=line.top,
                date_left=date_left,
                mode_left=mode_left,
                detail_left=detail_left,
                deposit_left=deposit_left,
                withdrawal_left=withdrawal_left,
                balance_left=balance_left,
            )
    return None


def _find_date(line: SavingsTextLine, layout: SavingsHeaderLayout) -> str | None:
    date_limit = (layout.mode_left + layout.detail_left) / 2.0
    return next(
        (
            token.text
            for token in line.tokens
            if token.left < date_limit and DATE_RE.fullmatch(token.text)
        ),
        None,
    )


def _amount_in_range(line: SavingsTextLine, minimum: float, maximum: float) -> float | None:
    candidates = [
        parse_amount(token.text)
        for token in line.tokens
        if minimum <= token.left < maximum and AMOUNT_RE.fullmatch(token.text)
    ]
    values = [value for value in candidates if value is not None]
    return values[-1] if values else None


def _line_text_in_range(line: SavingsTextLine, minimum: float, maximum: float) -> str:
    return clean_cell(" ".join(token.text for token in line.tokens if minimum <= token.left < maximum))


def _line_text(line: SavingsTextLine) -> str:
    return clean_cell(" ".join(token.text for token in line.tokens))


def _is_table_stop_line(line: SavingsTextLine) -> bool:
    upper = _line_text(line).upper()
    return (upper.startswith("ACCOUNT TYPE") and "ACCOUNT NUMBER" in upper) or upper.startswith(
        "NOMINEE NAME"
    )


def _is_non_transaction_text(text: str) -> bool:
    upper = text.upper()
    return (
        not upper
        or upper.startswith("PAGE ")
        or upper.startswith("YOUR BASE BRANCH")
        or upper.startswith("VISIT WWW.ICICIBANK")
        or upper.startswith("DIAL YOUR BANK")
    )


def _parse_savings_page_lines(
    lines: list[SavingsTextLine],
    previous_balance: float | None,
    logger,
) -> tuple[list[dict[str, Any]], float | None, bool]:
    header_match = _find_header(lines)
    if header_match is None:
        return [], previous_balance, False

    header_index, layout = header_match
    body_lines = lines[header_index + 1 :]
    dated_lines = [
        (line, date_text)
        for line in body_lines
        if (date_text := _find_date(line, layout)) is not None
    ]
    if not dated_lines:
        return [], previous_balance, True

    records: list[dict[str, Any]] = []
    for index, (date_line, raw_date) in enumerate(dated_lines):
        previous_top = dated_lines[index - 1][0].top if index > 0 else layout.top
        next_top = dated_lines[index + 1][0].top if index + 1 < len(dated_lines) else float("inf")
        lower_bound = (previous_top + date_line.top) / 2.0
        upper_bound = (date_line.top + next_top) / 2.0
        block_lines = [line for line in body_lines if lower_bound <= line.top < upper_bound]

        deposit = _amount_in_range(
            date_line,
            layout.detail_amount_boundary,
            layout.deposit_withdrawal_boundary,
        )
        withdrawal = _amount_in_range(
            date_line,
            layout.deposit_withdrawal_boundary,
            layout.withdrawal_balance_boundary,
        )
        balance = _amount_in_range(date_line, layout.withdrawal_balance_boundary, float("inf"))

        mode_text = _line_text_in_range(date_line, layout.mode_left - 3.0, layout.detail_left)
        detail_parts: list[str] = []
        for line in block_lines:
            if _is_table_stop_line(line):
                break
            part = _line_text_in_range(line, layout.detail_left - 3.0, layout.detail_amount_boundary)
            if not _is_non_transaction_text(part):
                detail_parts.append(part)

        details = clean_cell(" ".join([mode_text, *detail_parts]))

        if deposit is None and withdrawal is None:
            if balance is not None and details.upper() in {"B/F", "BF", "B F"}:
                previous_balance = balance
                logger.debug("ICICI savings opening balance: %.2f", balance)
            continue
        if balance is None:
            logger.warning("ICICI savings row skipped without balance: date=%s details=%s", raw_date, details)
            continue

        if previous_balance is not None:
            expected_balance = round(previous_balance + (deposit or 0.0) - (withdrawal or 0.0), 2)
            if abs(expected_balance - balance) > 0.05:
                logger.warning(
                    "ICICI savings balance mismatch: date=%s expected=%.2f parsed=%.2f details=%s",
                    raw_date,
                    expected_balance,
                    balance,
                    details,
                )

        records.append(
            {
                "Sno": 0,
                "Date": raw_date.replace("-", "/"),
                "Details": details,
                "Detail_Clean": clean_detail(details),
                "Cheque No": "",
                "Debit": withdrawal,
                "Credit": deposit,
                "Balance": balance,
            }
        )
        previous_balance = balance

    return records, previous_balance, True


def parse_savings_transaction_layout(
    pdf_path: str | Path,
    logger,
    progress_cb=None,
) -> list[dict[str, Any]]:
    """Parse the ICICI savings layout with explicit deposit/withdrawal columns."""

    records: list[dict[str, Any]] = []
    previous_balance: float | None = None
    layout_seen = False

    with pdfplumber.open(str(pdf_path)) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            page_records, previous_balance, page_layout_seen = _parse_savings_page_lines(
                _extract_lines(page),
                previous_balance,
                logger,
            )
            layout_seen = layout_seen or page_layout_seen
            logger.debug(
                "ICICI savings text page %s: identified %s transaction row(s)",
                page_number,
                len(page_records),
            )
            for record in page_records:
                records.append(record)
                if progress_cb is not None:
                    progress_cb(len(records))

    if not layout_seen:
        return []

    for index, record in enumerate(records, start=1):
        record["Sno"] = index

    logger.info("ICICI savings text parse complete: rows=%s", len(records))
    return records
