from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import axis_parser
from utils import build_output_row, detect_header_map, parse_with_config


class AxisLayoutTests(unittest.TestCase):
    def test_header_detection_prefers_specific_debit_and_credit_aliases(self) -> None:
        rows = [
            [
                "S.NO",
                "Transaction Date (dd/mm/yyyy)",
                "Value Date (dd/mm/yyyy)",
                "Particulars",
                "Debit Amount(INR)",
                "Credit Amount(INR)",
                "Balance(INR)",
                "Cheque Number",
            ]
        ]
        header_aliases = {
            "sno": ["s.no", "sno"],
            "date": ["transaction date", "date"],
            "details": ["particulars"],
            "amount": ["amount(inr)", "amount"],
            "drcr": ["debit/credit", "dr/cr"],
            "debit": ["debit amount(inr)", "debit amount", "debit"],
            "credit": ["credit amount(inr)", "credit amount", "credit"],
            "balance": ["balance(inr)", "balance"],
            "cheque": ["cheque number", "cheque"],
        }

        col_map, _ = detect_header_map(rows, header_aliases)

        self.assertEqual(col_map["debit"], 4)
        self.assertEqual(col_map["credit"], 5)
        self.assertNotEqual(col_map.get("amount"), 4)

    def test_old_axis_amount_and_drcr_layout_still_maps_debit(self) -> None:
        row = [
            "1",
            "04/04/2025",
            "04/04/2025",
            "RTGS sample",
            "6,00,000.00",
            "DR",
            "-8,16,658.90",
            "",
            "(4894)",
        ]
        col_map = {
            "sno": 0,
            "date": 1,
            "details": 3,
            "amount": 4,
            "drcr": 5,
            "balance": 6,
            "cheque": 7,
        }

        result = build_output_row(row, col_map, {})

        self.assertEqual(result["Debit"], 600000.0)
        self.assertIsNone(result["Credit"])

    def test_new_axis_separate_debit_and_credit_columns_map_debit(self) -> None:
        row = [
            "4",
            "01/11/2025",
            "01/11/2025",
            "TRF/1503/KAVIN ENTERPRISES/Kavin Enterprises",
            "1,40,000.00",
            "",
            "-2,97,08,480.71",
            "3538",
            "(1503)",
        ]
        col_map = {
            "sno": 0,
            "date": 1,
            "details": 3,
            "debit": 4,
            "credit": 5,
            "balance": 6,
            "cheque": 7,
        }

        result = build_output_row(row, col_map, {})

        self.assertEqual(result["Debit"], 140000.0)
        self.assertIsNone(result["Credit"])

    def test_new_axis_separate_debit_and_credit_columns_map_credit(self) -> None:
        row = [
            "5",
            "01/11/2025",
            "01/11/2025",
            "RTGS sample",
            "",
            "2,85,883.00",
            "-2,94,22,597.71",
            "",
            "(248)",
        ]
        col_map = {
            "sno": 0,
            "date": 1,
            "details": 3,
            "debit": 4,
            "credit": 5,
            "balance": 6,
            "cheque": 7,
        }

        result = build_output_row(row, col_map, {})

        self.assertIsNone(result["Debit"])
        self.assertEqual(result["Credit"], 285883.0)

    def test_parse_with_config_reuses_previous_header_map_for_headerless_pages(self) -> None:
        logger = logging.getLogger("tests.axis_layouts")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        tables = [
            (
                1,
                [
                    ["Tran Date", "Chq No", "Particulars", "Debit", "Credit", "Balance", "Init.\nBr"],
                    ["01-10-2025", "", "First row", "400.00", "", "15.98", "1739"],
                ],
            ),
            (
                2,
                [
                    ["12-10-2025", "", "Second row", "", "31.00", "46.98", "1739"],
                ],
            ),
        ]

        with patch("utils.extract_pdf_tables", return_value=tables):
            records = parse_with_config(
                pdf_path="dummy.pdf",
                logger=logger,
                header_aliases=axis_parser.HEADER_ALIASES,
                fallback_map=axis_parser.FALLBACK_MAP,
            )

        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["Date"], "01/10/2025")
        self.assertEqual(records[0]["Debit"], 400.0)
        self.assertEqual(records[1]["Date"], "12/10/2025")
        self.assertEqual(records[1]["Details"], "Second row")
        self.assertIsNone(records[1]["Debit"])
        self.assertEqual(records[1]["Credit"], 31.0)
        self.assertEqual(records[1]["Balance"], 46.98)


if __name__ == "__main__":
    unittest.main()
