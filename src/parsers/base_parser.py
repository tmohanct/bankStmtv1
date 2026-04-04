"""Base interface for all bank statement parsers."""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path

import pandas as pd


class BaseStatementParser(ABC):
    """Contract for bank-specific statement parsers."""

    bank_code: str = ""

    @abstractmethod
    def parse(self, pdf_path: Path, rules_df: pd.DataFrame) -> pd.DataFrame:
        """Parse a statement PDF and return parser output rows."""
        raise NotImplementedError
