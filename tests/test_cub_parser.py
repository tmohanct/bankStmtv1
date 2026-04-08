from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import cub_parser
from utils import is_date_token, normalize_date


class CubParserTests(unittest.TestCase):
    def test_month_abbreviation_dates_are_normalized(self) -> None:
        self.assertEqual(normalize_date("01-NOV-2025"), "01/11/2025")
        self.assertTrue(is_date_token("01-NOV-2025"))

    def test_cub_parser_reads_vairakannu_statement(self) -> None:
        pdf_path = PROJECT_ROOT / "input" / "VAIRAKANNU.pdf"
        logger = logging.getLogger("tests.cub_parser")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = cub_parser.parse(str(pdf_path), logger)

        self.assertEqual(len(records), 80)
        self.assertEqual(sum(record["Debit"] is not None for record in records), 11)
        self.assertEqual(sum(record["Credit"] is not None for record in records), 69)
        self.assertAlmostEqual(sum((record["Debit"] or 0.0) for record in records), 3910376.70)
        self.assertAlmostEqual(sum((record["Credit"] or 0.0) for record in records), 3912217.92)
        self.assertTrue(all("Particulars" not in record["Details"] for record in records))

        first = records[0]
        self.assertEqual(first["Sno"], 1)
        self.assertEqual(first["Date"], "01/11/2025")
        self.assertEqual(first["Details"], "BY NEFT TRF:ONE 97 COMMUNICA YESAP53051626677:")
        self.assertIsNone(first["Debit"])
        self.assertEqual(first["Credit"], 3750.0)
        self.assertEqual(first["Balance"], -2483887.14)

        last = records[-1]
        self.assertEqual(last["Sno"], 80)
        self.assertEqual(last["Date"], "30/11/2025")
        self.assertEqual(last["Details"], "TO DEBIT INTEREST:99999")
        self.assertEqual(last["Debit"], 12009.0)
        self.assertIsNone(last["Credit"])
        self.assertEqual(last["Balance"], -2485795.92)


if __name__ == "__main__":
    unittest.main()
