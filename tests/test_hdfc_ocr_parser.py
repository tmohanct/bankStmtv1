from __future__ import annotations

import logging
import sys
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(PROJECT_ROOT / "src" / "code"), str(PROJECT_ROOT / "src")]

import hdfc_parser
from parsers import hdfc_parser as hdfc_ocr


def word(text: str, left: int, top: int, width: int = 90) -> hdfc_ocr._Word:
    return hdfc_ocr._Word(text, left, top, width, 24)


class HDFCOCRParserTests(unittest.TestCase):
    def test_reads_scanned_columns_and_continuation_lines(self) -> None:
        lines = [
            [
                word("01/04/26", 63, 681),
                word("|NEFT-CR", 166, 681, 180),
                word("123456", 839, 681, 100),
                word("01/04/26", 1049, 681),
                word("300,000.00", 1498, 681, 107),
                word("1,723,302.02", 1719, 681, 117),
            ],
            [word("CUSTOMER", 166, 710)],
            [
                word("02/04/26", 63, 750),
                word("|CHQ-PAID", 166, 750, 180),
                word("1124", 839, 750),
                word("02/04/26", 1049, 750),
                word("50,000.00", 1265, 750, 106),
                word("1,673,302.02", 1719, 750, 117),
            ],
            [word("STATEMENT", 63, 900), word("SUMMARY", 200, 900)],
            [
                word("1,423,302.02", 63, 940),
                word("1", 300, 940),
                word("1", 370, 940),
                word("50,000.00", 440, 940),
                word("300,000.00", 700, 940),
                word("1,673,302.02", 1000, 940),
            ],
        ]
        rows, summary = hdfc_ocr._rows_on_page(lines, 0, 1836)
        self.assertEqual(len(rows), 2)
        self.assertTrue(rows[0].is_credit)
        self.assertFalse(rows[1].is_credit)
        self.assertEqual(rows[0].detail_parts, ["NEFT-CR", "CUSTOMER"])
        self.assertEqual(rows[0].cheque_parts, ["123456"])
        self.assertEqual(rows[1].amount_text, "50,000.00")
        self.assertEqual(summary.debit_count, 1)
        self.assertEqual(summary.credit_count, 1)

    def test_balance_ocr_repair_requires_exact_rupees(self) -> None:
        expected = Decimal("1605477.02")
        self.assertTrue(hdfc_ocr._partial_balance_matches("1,605,477.032", expected))
        self.assertTrue(hdfc_ocr._partial_balance_matches("1,605,477.0:", expected))
        self.assertFalse(hdfc_ocr._partial_balance_matches("1,605,478.0:", expected))

    def test_image_only_pdf_routes_to_ocr_parser(self) -> None:
        class Page:
            def get_text(self) -> str:
                return ""

        class Document:
            page_count = 1

            def __iter__(self):
                return iter([Page()])

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        with patch.object(hdfc_parser.fitz, "open", return_value=Document()):
            with patch.object(
                hdfc_parser, "parse_scanned_hdfc", return_value=[{"Sno": 1}]
            ) as scanned:
                result = hdfc_parser.parse("scan.pdf", logging.getLogger("test.hdfc"))
        self.assertEqual(result, [{"Sno": 1}])
        scanned.assert_called_once()


if __name__ == "__main__":
    unittest.main()
