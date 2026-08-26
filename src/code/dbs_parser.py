from __future__ import annotations

import logging
from typing import Any

import pdfplumber

from utils import (
    clean_cell,
    clean_detail,
    is_date_token,
    normalize_cheque_number,
    normalize_date,
    parse_amount,
    parse_with_config,
)

HEADER_ALIASES: dict[str, list[str]] = {
    "date": ["date", "transaction date"],
    "value_date": ["value date", "valuedate"],
    "details": ["transaction details", "particulars", "narration", "description"],
    "debit": ["debit", "withdrawal", "withdrawals"],
    "credit": ["credit", "deposit", "deposits"],
    "balance": ["running balance", "balance"],
}

FALLBACK_MAP = {
    "date": 0,
    "details": 2,
    "debit": 3,
    "credit": 4,
    "balance": 5,
}

DETAILS_LEFT = 145.0
DETAILS_RIGHT = 315.0
DEBIT_LEFT = 315.0
DEBIT_RIGHT = 405.0
CREDIT_LEFT = 405.0
CREDIT_RIGHT = 495.0
BALANCE_LEFT = 495.0
BODY_BOTTOM = 760.0
LINE_TOLERANCE = 1.5


def _word_x(word: dict[str, Any]) -> float:
    return float(word.get("x0", 0.0))


def _word_top(word: dict[str, Any]) -> float:
    return float(word.get("top", 0.0))


def _group_words_into_lines(words: list[dict[str, Any]]) -> list[list[dict[str, Any]]]:
    ordered = sorted(words, key=lambda word: (_word_top(word), _word_x(word)))
    grouped: list[list[dict[str, Any]]] = []
    line_tops: list[float] = []

    for word in ordered:
        top = _word_top(word)
        if not grouped or abs(top - line_tops[-1]) > LINE_TOLERANCE:
            grouped.append([word])
            line_tops.append(top)
            continue

        grouped[-1].append(word)
        count = len(grouped[-1])
        line_tops[-1] = ((line_tops[-1] * (count - 1)) + top) / count

    for line in grouped:
        line.sort(key=_word_x)
    return grouped


def _text_between(words: list[dict[str, Any]], left: float, right: float) -> str:
    return clean_cell(
        " ".join(
            str(word.get("text", ""))
            for word in words
            if left <= _word_x(word) < right
        )
    )


def _date_between(words: list[dict[str, Any]], left: float, right: float) -> str:
    for word in words:
        if left <= _word_x(word) < right and is_date_token(word.get("text", "")):
            return clean_cell(word.get("text", ""))
    return ""


def _amount_between(
    words: list[dict[str, Any]], left: float, right: float
) -> float | None:
    for word in words:
        if left <= _word_x(word) < right:
            amount = parse_amount(word.get("text", ""))
            if amount is not None:
                return amount
    return None


def _positioned_transaction_line(
    words: list[dict[str, Any]],
) -> tuple[str, str, str, float | None, float | None, float] | None:
    transaction_date = _date_between(words, 25.0, 90.0)
    value_date = _date_between(words, 90.0, DETAILS_LEFT)
    if not transaction_date or not value_date:
        return None

    balance = _amount_between(words, BALANCE_LEFT, 580.0)
    if balance is None:
        return None

    debit = _amount_between(words, DEBIT_LEFT, DEBIT_RIGHT)
    credit = _amount_between(words, CREDIT_LEFT, CREDIT_RIGHT)
    if debit is None and credit is None:
        return None

    details = _text_between(words, DETAILS_LEFT, DETAILS_RIGHT)
    return transaction_date, value_date, details, debit, credit, balance


def _append_details(record: dict[str, Any], value: str) -> None:
    detail_text = clean_cell(value)
    if not detail_text:
        return

    existing = clean_cell(record.get("Details", ""))
    merged = f"{existing} {detail_text}".strip()
    record["Details"] = merged
    record["Detail_Clean"] = clean_detail(merged)


def _positioned_signature(
    transaction_date: str,
    value_date: str,
    debit: float | None,
    credit: float | None,
    balance: float,
) -> tuple[str, str, float | None, float | None, float]:
    return (
        normalize_date(transaction_date) or transaction_date,
        normalize_date(value_date) or value_date,
        debit,
        credit,
        balance,
    )


def _parse_positioned_dbs(
    pdf_path: str,
    logger: logging.Logger,
    progress_cb=None,
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    last_signature: tuple[str, str, float | None, float | None, float] | None = None

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            extract_words = getattr(page, "extract_words", None)
            if not callable(extract_words):
                continue

            words = extract_words(use_text_flow=True, keep_blank_chars=False) or []
            lines = _group_words_into_lines(words)
            positioned_lines = [
                (line_index, parsed)
                for line_index, line in enumerate(lines)
                if (parsed := _positioned_transaction_line(line)) is not None
            ]
            if not positioned_lines:
                logger.debug("Page %s: no positioned DBS transaction rows", page_number)
                continue

            transaction_line_indexes = {line_index for line_index, _ in positioned_lines}
            first_transaction_line_index = positioned_lines[0][0]
            current_record: dict[str, Any] | None = None

            for line_index, line in enumerate(lines):
                if _word_top(line[0]) >= BODY_BOTTOM:
                    break

                if line_index in transaction_line_indexes:
                    parsed = _positioned_transaction_line(line)
                    if parsed is None:
                        continue
                    transaction_date, value_date, details, debit, credit, balance = parsed
                    signature = _positioned_signature(
                        transaction_date,
                        value_date,
                        debit,
                        credit,
                        balance,
                    )

                    is_page_continuation = (
                        line_index == first_transaction_line_index
                        and records
                        and signature == last_signature
                    )
                    if is_page_continuation:
                        current_record = records[-1]
                        _append_details(current_record, details)
                        logger.debug(
                            "Page %s: merged repeated page-boundary transaction %s",
                            page_number,
                            signature,
                        )
                        continue

                    normalized_transaction_date = normalize_date(transaction_date)
                    current_record = {
                        "Sno": len(records) + 1,
                        "Date": normalized_transaction_date or transaction_date,
                        "Details": clean_cell(details),
                        "Detail_Clean": clean_detail(details),
                        "Cheque No": normalize_cheque_number("", details),
                        "Debit": debit,
                        "Credit": credit,
                        "Balance": balance,
                    }
                    records.append(current_record)
                    last_signature = signature
                    if progress_cb is not None:
                        progress_cb(len(records))
                    continue

                if current_record is None:
                    continue

                full_line = clean_cell(" ".join(str(word.get("text", "")) for word in line))
                if "TOTAL DEBIT COUNT" in full_line.upper() or "TOTAL CREDIT COUNT" in full_line.upper():
                    current_record = None
                    continue

                _append_details(
                    current_record,
                    _text_between(line, DETAILS_LEFT, DETAILS_RIGHT),
                )

            logger.debug(
                "Page %s: positioned DBS rows seen=%s total_records=%s",
                page_number,
                len(positioned_lines),
                len(records),
            )

    for row_number, record in enumerate(records, start=1):
        record["Sno"] = row_number
    return records


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing DBS statement: %s", pdf_path)
    records = parse_with_config(
        pdf_path=pdf_path,
        logger=logger,
        header_aliases=HEADER_ALIASES,
        fallback_map=FALLBACK_MAP,
        progress_cb=progress_cb,
    )
    if records:
        return records

    logger.info("DBS table extraction returned no rows; trying positioned-text layout")
    return _parse_positioned_dbs(
        pdf_path=pdf_path,
        logger=logger,
        progress_cb=progress_cb,
    )
