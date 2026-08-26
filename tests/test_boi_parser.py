from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import bank_detector
import boi_parser

SAMPLE_PDF = PROJECT_ROOT / "input" / "IndianBank.pdf"


@unittest.skipUnless(SAMPLE_PDF.is_file(), "Bank of India sample PDF is required for this regression test.")
class BOIParserTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.pdf_path = SAMPLE_PDF
        cls.logger = logging.getLogger("tests.boi_parser")
        cls.logger.handlers.clear()
        cls.logger.addHandler(logging.NullHandler())

    def test_boi_parser_reads_supplied_statement(self) -> None:
        records = boi_parser.parse(str(self.pdf_path), self.logger)

        self.assertEqual(len(records), 1778)
        self.assertEqual(sum(row["Debit"] is not None for row in records), 383)
        self.assertEqual(sum(row["Credit"] is not None for row in records), 1395)

        first = records[0]
        self.assertEqual(first["Sno"], 1)
        self.assertEqual(first["Date"], "01/01/2026")
        self.assertIn("UPI/600186901588", first["Details"])
        self.assertIsNone(first["Debit"])
        self.assertEqual(first["Credit"], 16830.0)
        self.assertEqual(first["Balance"], -274374806.31)

        cheque_row = records[99]
        self.assertEqual(cheque_row["Sno"], 100)
        self.assertEqual(cheque_row["Date"], "02/01/2026")
        self.assertEqual(cheque_row["Details"], "CHETTINAD CEMENTCORP")
        self.assertEqual(cheque_row["Cheque No"], "29721")
        self.assertEqual(cheque_row["Debit"], 400000.0)
        self.assertIsNone(cheque_row["Credit"])
        self.assertEqual(cheque_row["Balance"], -274603090.05)

        last = records[-1]
        self.assertEqual(last["Sno"], 1778)
        self.assertEqual(last["Date"], "31/01/2026")
        self.assertEqual(last["Details"], "IBRTGS/HDFC/DALMIA CEMENT LTD")
        self.assertEqual(last["Debit"], 800000.0)
        self.assertIsNone(last["Credit"])
        self.assertEqual(last["Balance"], -274914935.44)

    def test_every_amount_matches_the_balance_change(self) -> None:
        records = boi_parser.parse(str(self.pdf_path), self.logger)

        for previous, current in zip(records, records[1:]):
            actual_change = round(current["Balance"] - previous["Balance"], 2)
            expected_change = round(
                (current["Credit"] or 0.0) - (current["Debit"] or 0.0),
                2,
            )
            self.assertAlmostEqual(
                actual_change,
                expected_change,
                places=2,
                msg=f"Balance mismatch at transaction {current['Sno']}",
            )

    def test_bank_detector_identifies_boi_statement(self) -> None:
        detected = bank_detector.detect_bank_from_pdf(self.pdf_path, self.logger)
        self.assertEqual(detected, "boi")


if __name__ == "__main__":
    unittest.main()
