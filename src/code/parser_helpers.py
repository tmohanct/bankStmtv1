from __future__ import annotations

import re
from datetime import datetime
from typing import Iterable

from utils import clean_cell, clean_detail, normalize_cheque_number, parse_amount


def normalize_date_with_formats(value: str, formats: Iterable[str]) -> str:
    text = clean_cell(value)
    if not text:
        return ""

    normalized = re.sub(r"\s+", " ", text).strip()
    normalized = normalized.replace(" -", "-").replace("- ", "-")
    normalized = normalized.replace(" /", "/").replace("/ ", "/")

    for fmt in formats:
        try:
            parsed = datetime.strptime(normalized, fmt)
            return parsed.strftime("%d/%m/%Y")
        except ValueError:
            continue
    return normalized


def parse_signed_balance(value: str) -> float | None:
    text = clean_cell(value)
    if not text:
        return None

    upper = text.upper().rstrip(".")
    negative = text.startswith("-") or upper.endswith("DR")
    cleaned = re.sub(r"\s*(CR|DR)\.?$", "", text, flags=re.IGNORECASE)
    amount = parse_amount(cleaned)
    if amount is None:
        return None
    if negative and amount > 0:
        return -amount
    return amount


def build_record(
    *,
    date_text: str,
    details: str,
    cheque_no: str = "",
    debit: float | None = None,
    credit: float | None = None,
    balance: float | None = None,
    date_formats: Iterable[str],
) -> dict[str, object]:
    clean_details = clean_cell(details)
    return {
        "Sno": 0,
        "Date": normalize_date_with_formats(date_text, date_formats),
        "Details": clean_details,
        "Detail_Clean": clean_detail(clean_details),
        "Cheque No": normalize_cheque_number(cheque_no, clean_details),
        "Debit": debit,
        "Credit": credit,
        "Balance": balance,
    }
