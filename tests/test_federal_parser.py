from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import federal_parser


class _FakePage:
    def __init__(self, tables: list[list[list[str]]]) -> None:
        self._tables = tables

    def extract_tables(self) -> list[list[list[str]]]:
        return self._tables


class _FakePdf:
    def __init__(self, pages: list[_FakePage]) -> None:
        self.pages = pages

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        return False


class FederalParserUnitTests(unittest.TestCase):
    def setUp(self) -> None:
        self.logger = logging.getLogger("tests.federal")
        self.logger.handlers.clear()
        self.logger.addHandler(logging.NullHandler())

    def test_parses_ten_column_layout_and_balance_indicator(self) -> None:
        table = [
            [
                "Date",
                "Value Date",
                "Particulars",
                "Tran Type",
                "Tran ID",
                "Cheque Details",
                "Withdrawals",
                "Deposits",
                "Balance",
                "DR /CR",
            ],
            [
                "05-JAN-2026",
                "05-JAN-2026",
                "UPI IN/615650936299",
                "TFR",
                "S63323013",
                "",
                "",
                "1,000.00",
                "1,250.00",
                "Cr",
            ],
            [
                "05-JAN-2026",
                "05-JAN-2026",
                "CHEQUE PAYMENT",
                "CLG",
                "S66513477",
                "000123",
                "1,500.00",
                "",
                "250.00",
                "Dr",
            ],
        ]

        with patch.object(
            federal_parser.pdfplumber,
            "open",
            return_value=_FakePdf([_FakePage([table])]),
        ):
            records = federal_parser.parse("dummy.pdf", self.logger)

        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["Date"], "05/01/2026")
        self.assertEqual(records[0]["Credit"], 1000.0)
        self.assertEqual(records[0]["Balance"], 1250.0)
        self.assertEqual(records[1]["Cheque No"], "000123")
        self.assertEqual(records[1]["Debit"], 1500.0)
        self.assertEqual(records[1]["Balance"], -250.0)

    def test_legacy_eight_column_layout_remains_supported(self) -> None:
        table = [
            [
                "Date",
                "Value Date",
                "Particulars",
                "Tran Type",
                "Cheque Details",
                "Withdrawals",
                "Deposits",
                "Balance",
            ],
            [
                "20-AUG- 2025",
                "20-AUG-25 01:54:18 PM",
                "CHRG/MIN BAL/JUL25",
                "D",
                "",
                "413.00",
                "",
                "10,739.26",
            ],
        ]

        with patch.object(
            federal_parser.pdfplumber,
            "open",
            return_value=_FakePdf([_FakePage([table])]),
        ):
            records = federal_parser.parse("dummy.pdf", self.logger)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Date"], "20/08/2025")
        self.assertEqual(records[0]["Debit"], 413.0)
        self.assertIsNone(records[0]["Credit"])
        self.assertEqual(records[0]["Balance"], 10739.26)

    def test_balance_suffix_is_used_when_indicator_column_is_absent(self) -> None:
        self.assertEqual(federal_parser._parse_balance("250.00 Dr"), -250.0)
        self.assertEqual(federal_parser._parse_balance("250.00DR"), -250.0)
        self.assertEqual(federal_parser._parse_balance("250.00 Cr"), 250.0)


if __name__ == "__main__":
    unittest.main()
