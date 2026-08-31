from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

from utils import OUTPUT_COLUMNS, remove_exact_duplicate_transactions


class TransactionDeduplicationTests(unittest.TestCase):
    def test_removes_identical_transactions_from_different_source_pdfs(self) -> None:
        original = {
            "Date": "31/08/2026",
            "Details": "UPI/PAID TO HARI/123456789012",
            "Detail_Clean": "UPIPAIDTOHARI123456789012",
            "Cheque No": "",
            "Debit": 500.0,
            "Credit": None,
            "Balance": 1500.0,
        }
        rows = [
            {"Sno": 386, "Source": "statement_august.pdf", **original},
            {"Sno": 729, "Source": "statement_september.pdf", **original},
            {
                "Sno": 730,
                "Source": "statement_september.pdf",
                **{**original, "Balance": 1000.0},
            },
        ]

        result = remove_exact_duplicate_transactions(pd.DataFrame(rows, columns=OUTPUT_COLUMNS))

        self.assertEqual(len(result), 2)
        self.assertEqual(result.iloc[0]["Sno"], 386)
        self.assertEqual(result.iloc[0]["Source"], "statement_august.pdf")
        self.assertEqual(result.iloc[1]["Sno"], 730)


if __name__ == "__main__":
    unittest.main()
