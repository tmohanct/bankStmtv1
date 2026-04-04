from typing import Any

from utils import parse_with_config

HEADER_ALIASES: dict[str, list[str]] = {
    "sno": ["s.no", "sno", "serial", "sr no"],
    "date": ["transaction date", "transaction", "txn date", "date"],
    "details": ["particulars", "narration", "description", "details"],
    "amount": ["amount(inr)", "amount", "txn amount"],
    "drcr": ["debit/credit", "dr/cr", "drcr", "type"],
    "debit": ["debit amount(inr)", "debit amount", "withdrawal amount", "debit"],
    "credit": ["credit amount(inr)", "credit amount", "deposit amount", "credit"],
    "balance": ["balance(inr)", "balance", "closing balance"],
    "cheque": ["cheque number", "cheque no", "cheque", "chq no"],
}

FALLBACK_MAP = {
    "sno": 0,
    "date": 1,
    "details": 3,
    "amount": 4,
    "drcr": 5,
    "balance": 6,
    "cheque": 7,
}


def parse(pdf_path: str, logger, progress_cb=None) -> list[dict[str, Any]]:
    logger.info("Parsing Axis statement: %s", pdf_path)
    return parse_with_config(
        pdf_path=pdf_path,
        logger=logger,
        header_aliases=HEADER_ALIASES,
        fallback_map=FALLBACK_MAP,
        progress_cb=progress_cb,
    )
