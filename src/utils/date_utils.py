"""Date conversion helpers."""

from __future__ import annotations

from datetime import date, datetime


def parse_date(raw_value: object) -> date | None:
    if raw_value is None:
        return None

    text = str(raw_value).strip()
    if not text:
        return None

    formats = ["%d-%m-%Y", "%d/%m/%Y", "%d-%b-%Y", "%Y-%m-%d"]
    for fmt in formats:
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue

    # TODO: Expand for additional bank-specific formats and locale handling.
    return None
