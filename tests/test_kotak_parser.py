from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from parsers.kotak_parser import parse_kotak_records

SAMPLE_PDF = PROJECT_ROOT / "input" / "BALASUBRAMANIAN.pdf"


@unittest.skipUnless(SAMPLE_PDF.is_file(), "Kotak sample PDF is required for this regression test.")
class KotakParserRegressionTests(unittest.TestCase):
    def test_balasubramanian_pdf_is_parsed(self) -> None:
        logger = logging.getLogger("tests.kotak")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = parse_kotak_records(pdf_path=SAMPLE_PDF, logger=logger)

        self.assertEqual(len(records), 127)

        first = records[0]
        self.assertEqual(first["Date"], "01/03/2026")
        self.assertEqual(first["Details"], "UPI/B MALLIGA/606030001225/UPI")
        self.assertEqual(first["Cheque No"], "UPI-606022401064")
        self.assertIsNone(first["Debit"])
        self.assertEqual(first["Credit"], 7000.0)
        self.assertEqual(first["Balance"], 147291.4)

        last = records[-1]
        self.assertEqual(last["Date"], "20/03/2026")
        self.assertEqual(last["Details"], "UPI/8807930865@ptye/607913274747/UPI")
        self.assertEqual(last["Cheque No"], "UPI-607945105105")
        self.assertEqual(last["Debit"], 1625.0)
        self.assertIsNone(last["Credit"])
        self.assertEqual(last["Balance"], 473435.48)


if __name__ == "__main__":
    unittest.main()
