from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import axis_parser

SAMPLE_PDF = PROJECT_ROOT / "input" / "RAJASEKARAN1.pdf"


@unittest.skipUnless(SAMPLE_PDF.is_file(), "Axis sample PDF is required for this regression test.")
class AxisParserRegressionTests(unittest.TestCase):
    def test_rajasekaran1_pdf_parses_beyond_first_page(self) -> None:
        logger = logging.getLogger("tests.axis_parser_regression")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = axis_parser.parse(str(SAMPLE_PDF), logger)

        self.assertGreater(len(records), 100)
        self.assertEqual(records[0]["Date"], "01/10/2025")
        self.assertTrue(
            any(
                record["Date"] == "21/12/2025"
                and "CAPITALFLOAT" in record["Details"]
                for record in records
            )
        )
        self.assertEqual(records[-1]["Date"], "30/03/2026")
        self.assertNotIn("TRANSACTION TOTAL", records[-1]["Details"])
        self.assertNotIn("CLOSING BALANCE", records[-1]["Details"])


if __name__ == "__main__":
    unittest.main()
