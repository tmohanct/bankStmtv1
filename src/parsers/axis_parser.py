"""Axis Bank parser implementation."""

from __future__ import annotations

from pathlib import Path
import pandas as pd
from src.parsers.base_parser import BaseStatementParser
from typing import Any
from src.utils.statement_utils import parse_with_config


class AxisParser(BaseStatementParser):
    """Axis Bank statement parser."""

    bank_code = "axis"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        import logging
        records = parse(str(pdf_path), logging.getLogger(__name__))
        return pd.DataFrame(records).rename(columns={"Details": "Description", "Cheque No": "Reference"})


# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {'name': 'Axis Bank', 'ifsc': 'UTIB', 'aliases': ['Axis Bank'], 'name_after': 'Account Statement Report', 'address_label': 'Joint Holder'}

PDF_STATUS_PROFILE.update({'unlabelled_left': True})


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


BANK_CODE = 'axis'
BANK_SIGNATURES = (('AXIS BANK', 4), ('UTIB0', 3), ('ACCOUNT STATEMENT REPORT', 1))
