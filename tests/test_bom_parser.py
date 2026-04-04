from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import bank_detector
import bom_parser


class BomParserTests(unittest.TestCase):
    def test_bom_parser_reads_sample_statement(self) -> None:
        pdf_path = PROJECT_ROOT / "input" / "bom.pdf"
        logger = logging.getLogger("tests.bom_parser")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = bom_parser.parse(str(pdf_path), logger)

        self.assertEqual(len(records), 446)

        first = records[0]
        self.assertEqual(first["Sno"], 1)
        self.assertEqual(first["Date"], "01/11/2025")
        self.assertEqual(first["Cheque No"], "530519742698")
        self.assertIsNone(first["Debit"])
        self.assertEqual(first["Credit"], 28937.0)
        self.assertEqual(first["Balance"], -468730.6)
        self.assertIn("NEWCFLIDFCB DISBURSA/L", first["Details"])

        branch_row = records[19]
        self.assertEqual(branch_row["Cheque No"], "174")
        self.assertEqual(branch_row["Debit"], 19107.0)
        self.assertIn("CHENNAI SERVICE BRANCH", branch_row["Details"])

        last = records[-1]
        self.assertEqual(last["Sno"], 446)
        self.assertEqual(last["Date"], "01/03/2026")
        self.assertEqual(last["Debit"], 0.9)
        self.assertIsNone(last["Credit"])
        self.assertEqual(last["Balance"], -491695.26)
        self.assertEqual(last["Details"], "GST IMPS")

    def test_bank_detector_can_score_bom_text(self) -> None:
        detected = bank_detector._detect_from_text("IFSC MAHB0002616 bom2616@mahabank.co.in")
        self.assertEqual(detected, "bom")


if __name__ == "__main__":
    unittest.main()
