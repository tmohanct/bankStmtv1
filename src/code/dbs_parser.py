from __future__ import annotations

from typing import Any

from utils import parse_with_config

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


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing DBS statement: %s", pdf_path)
    return parse_with_config(
        pdf_path=pdf_path,
        logger=logger,
        header_aliases=HEADER_ALIASES,
        fallback_map=FALLBACK_MAP,
        progress_cb=progress_cb,
    )
