"""Axis Bank parser implementation."""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import pdfplumber

from parsers.base_parser import BaseStatementParser


class AxisParser(BaseStatementParser):
    """Axis Bank statement parser."""

    bank_code = "axis"

    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        """Parse Axis statement PDF and return raw transaction rows."""
        # TODO: Use rules_df to drive row extraction, column mapping, and field cleanups.
        # Axis-specific parsing must stay only in this module.
        rows: list[dict[str, object]] = []

        with pdfplumber.open(str(pdf_path)) as pdf:
            for page_number, page in enumerate(pdf.pages, start=1):
                page_text = page.extract_text() or ""
                _ = page_text
                _ = page_number
                # TODO: Parse Axis statement table blocks and append parsed row dicts.

        return pd.DataFrame(
            rows,
            columns=["Date", "Value_Date", "Description", "Debit", "Credit", "Balance", "Reference", "Source_Page"],
        )
