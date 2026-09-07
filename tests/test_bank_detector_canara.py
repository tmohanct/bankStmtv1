from __future__ import annotations

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

from bank_detector import _detect_from_text


class CanaraBankDetectorTests(unittest.TestCase):
    def test_statement_header_outweighs_counterparty_bank_references(self) -> None:
        statement_text = """
        Current & Saving Account Statement
        IFSC Code CNRB0001206
        Txn Date Value Date Cheque No. Description Branch Code Debit Credit Balance
        NEFT Cr-ICIC0000123-ICICI BANK
        NEFT Cr-BARB0000123-BANK OF BARODA
        """

        self.assertEqual(_detect_from_text(statement_text), "canara")


if __name__ == "__main__":
    unittest.main()
