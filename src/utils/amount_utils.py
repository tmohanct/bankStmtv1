"""Amount parsing helpers."""

from __future__ import annotations

from decimal import Decimal, InvalidOperation


def parse_amount(raw_value: object) -> float | None:
    if raw_value is None:
        return None

    text = str(raw_value).strip()
    if not text:
        return None

    normalized = (
        text.replace(",", "")
        .replace("INR", "")
        .replace("Cr", "")
        .replace("CR", "")
        .replace("Dr", "")
        .replace("DR", "")
        .strip()
    )

    negative = normalized.startswith("(") and normalized.endswith(")")
    if negative:
        normalized = normalized[1:-1]

    try:
        value = Decimal(normalized)
    except InvalidOperation:
        # TODO: Add custom patterns from rules-driven parsing.
        return None

    if negative:
        value = -value

    return float(value)


def split_debit_credit(amount: float | None) -> tuple[float | None, float | None]:
    if amount is None:
        return None, None
    if amount < 0:
        return abs(amount), None
    return None, amount
