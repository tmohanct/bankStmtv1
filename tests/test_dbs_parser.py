from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import dbs_parser
import utils


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


class DBSParserUnitTests(unittest.TestCase):
    def setUp(self) -> None:
        self.logger = logging.getLogger("tests.dbs")
        self.logger.handlers.clear()
        self.logger.addHandler(logging.NullHandler())

    def test_parses_dbs_six_column_layout(self) -> None:
        table = [
            ["Date", "Value Date", "Transaction Details", "Debit", "Credit", "Running Balance"],
            [
                "01-Sep-2025",
                "01-Sep-2025",
                "TRANSFER\nIMPS Pay 524412316098",
                "50,000.00",
                "",
                "1,031.07",
            ],
            [
                "02-Sep-2025",
                "02-Sep-2025",
                "TRANSFER\nNEFTIN IOBAN25245070250 NUZHA",
                "",
                "350,000.00",
                "351,031.07",
            ],
        ]

        with patch.object(
            utils.pdfplumber,
            "open",
            return_value=_FakePdf([_FakePage([table])]),
        ):
            records = dbs_parser.parse("dummy.pdf", self.logger)

        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["Date"], "01/09/2025")
        self.assertEqual(records[0]["Details"], "TRANSFER IMPS Pay 524412316098")
        self.assertEqual(records[0]["Debit"], 50000.0)
        self.assertIsNone(records[0]["Credit"])
        self.assertEqual(records[0]["Balance"], 1031.07)
        self.assertEqual(records[1]["Credit"], 350000.0)


if __name__ == "__main__":
    unittest.main()
