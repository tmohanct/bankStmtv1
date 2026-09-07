from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import canara_parser


class _FakePage:
    def __init__(self, text: str) -> None:
        self._text = text

    def get_text(self, kind: str) -> str:
        return self._text


class _FakePdf:
    def __init__(self, pages: list[_FakePage]) -> None:
        self._pages = pages

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        return False

    def __iter__(self):
        return iter(self._pages)


class CanaraParserTests(unittest.TestCase):
    def setUp(self) -> None:
        self.logger = logging.getLogger("tests.canara")
        self.logger.handlers.clear()
        self.logger.addHandler(logging.NullHandler())

    def test_parses_timestamped_current_savings_layout(self) -> None:
        page_one = """
Opening Balance
Rs. -6,00,66,628.13
Txn Date
Value Date
Cheque No.
Description
Branch
Code
Debit
Credit
Balance
01-04-2025 23:25:51
01 Apr 2025
NEFT Cr-HDFCH00155478552-
HDFC0000240-NZ SEASONAL WEAR PRIVATE LIMITED
33
43,126.00
-6,00,23,502.13
Page 1 of 2
"""
        page_two = """
Txn Date
Value Date
Cheque No.
Description
Branch
Code
Debit
Credit
Balance
03-04-2025 13:53:41
03 Apr 2025
851445100
PALANIYANDI P Online Transaction OTH-Rado interest payment
1206
6,00,000.00
-6,06,23,502.13
Page 2 of 2
"""

        with patch.object(
            canara_parser.fitz,
            "open",
            return_value=_FakePdf([_FakePage(page_one), _FakePage(page_two)]),
        ):
            records = canara_parser.parse("dummy.pdf", self.logger)

        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["Date"], "01/04/2025")
        self.assertEqual(
            records[0]["Details"],
            "NEFT Cr-HDFCH00155478552- HDFC0000240-NZ SEASONAL WEAR PRIVATE LIMITED",
        )
        self.assertEqual(records[0]["Credit"], 43126.0)
        self.assertIsNone(records[0]["Debit"])
        self.assertEqual(records[1]["Cheque No"], "851445100")
        self.assertEqual(records[1]["Debit"], 600000.0)
        self.assertIsNone(records[1]["Credit"])

    def test_keeps_legacy_date_only_layout_supported(self) -> None:
        page = """
Opening Balance
1,000.00
Date
Particulars
Deposits
Withdrawals
Balance
01-04-2025
NEFT CR PAYMENT
250.00
1,250.00
"""

        with patch.object(
            canara_parser.fitz,
            "open",
            return_value=_FakePdf([_FakePage(page)]),
        ):
            records = canara_parser.parse("dummy.pdf", self.logger)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Date"], "01/04/2025")
        self.assertEqual(records[0]["Credit"], 250.0)
        self.assertEqual(records[0]["Balance"], 1250.0)


if __name__ == "__main__":
    unittest.main()
