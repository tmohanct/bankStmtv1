from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import sbi_parser
import bank_detector


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


class SBIParserUnitTests(unittest.TestCase):
    def setUp(self) -> None:
        self.logger = logging.getLogger("tests.sbi")
        self.logger.handlers.clear()
        self.logger.addHandler(logging.NullHandler())

    def test_normalizes_compact_and_month_first_ocr_dates(self) -> None:
        self.assertEqual(sbi_parser._normalize_sbi_date("8Mar 2026"), "08/03/2026")
        self.assertEqual(sbi_parser._normalize_sbi_date("Mar 10 2026"), "10/03/2026")

    def test_masked_description_ocr_noise_is_removed(self) -> None:
        self.assertEqual(sbi_parser._clean_ocr_text("KKK KKK"), "")

    def test_detector_recognizes_ocr_ifsc_with_letter_o(self) -> None:
        ocr_text = "Account Statement IFS Code :SBINO070009"
        self.assertEqual(bank_detector._detect_from_text(ocr_text), "sbi")

    def test_builds_ocr_debit_row_and_repairs_split_decimal(self) -> None:
        cells = [
            "10Mar 2026",
            "10Mar 2026",
            "TO | TRANSFER- INB TRANSFER-",
            "CIAAMPKAV 0 TRANSFER TO 44561598990",
            "99922",
            "1,57,500.0 0",
            "",
            "10.63",
        ]

        record = sbi_parser._build_ocr_record(cells, 157510.63, self.logger)

        self.assertIsNotNone(record)
        self.assertEqual(record["Date"], "10/03/2026")
        self.assertEqual(record["Details"], "TO TRANSFER- INB TRANSFER-")
        self.assertEqual(record["Debit"], 157500.0)
        self.assertIsNone(record["Credit"])
        self.assertEqual(record["Balance"], 10.63)
        self.assertEqual(record["Cheque No"], "")

    def test_builds_ocr_clearing_row_and_keeps_cheque_number(self) -> None:
        cells = [
            "9 Jul 2026",
            "9 Jul 2026",
            "TO CLEARING- Chq 125032 Sess 2 C RAMESH HUF",
            "/ 125032",
            "10395",
            "50,000.00",
            "",
            "6,458.36",
        ]

        record = sbi_parser._build_ocr_record(cells, 56458.36, self.logger)

        self.assertIsNotNone(record)
        self.assertEqual(record["Debit"], 50000.0)
        self.assertEqual(record["Balance"], 6458.36)
        self.assertEqual(record["Cheque No"], "125032")

    def test_text_table_layout_remains_supported(self) -> None:
        table = [
            ["Txn Date", "Value Date", "Description", "Ref No.", "Debit", "Credit", "Balance"],
            ["8 Mar 2026", "8 Mar 2026", "CSH DEP", "/", "", "14,500.00", "14,500.00"],
        ]

        with (
            patch.object(
                sbi_parser.pdfplumber,
                "open",
                return_value=_FakePdf([_FakePage([table])]),
            ),
            patch.object(sbi_parser, "_parse_ocr") as ocr_parser,
        ):
            records = sbi_parser.parse("dummy.pdf", self.logger)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Date"], "08/03/2026")
        self.assertEqual(records[0]["Credit"], 14500.0)
        ocr_parser.assert_not_called()

    def test_scanned_pdf_falls_back_to_ocr(self) -> None:
        expected = [
            {
                "Sno": 1,
                "Date": "08/03/2026",
                "Details": "CSH DEP",
                "Detail_Clean": "CSHDEP",
                "Cheque No": "",
                "Debit": None,
                "Credit": 14500.0,
                "Balance": 14500.0,
            }
        ]

        with (
            patch.object(
                sbi_parser.pdfplumber,
                "open",
                return_value=_FakePdf([_FakePage([])]),
            ),
            patch.object(sbi_parser, "_parse_ocr", return_value=expected) as ocr_parser,
        ):
            records = sbi_parser.parse("scanned.pdf", self.logger)

        self.assertEqual(records, expected)
        ocr_parser.assert_called_once_with("scanned.pdf", self.logger, None)


if __name__ == "__main__":
    unittest.main()
