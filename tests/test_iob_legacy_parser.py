from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
CODE_ROOT = str(PROJECT_ROOT / "src" / "code")
sys.path.insert(0, CODE_ROOT)

import iob_parser

if sys.path and sys.path[0] == CODE_ROOT:
    sys.path.pop(0)
for module_name in ("utils", "parser_helpers"):
    sys.modules.pop(module_name, None)

SAMPLE_PDF = PROJECT_ROOT / "input" / "AKILANMANIVANNAN.pdf"


@unittest.skipUnless(SAMPLE_PDF.is_file(), "IOB sample PDF is required for this regression test.")
class IOBLegacyParserTests(unittest.TestCase):
    def test_akilanmanivannan_pdf_populates_cheque_numbers_for_numeric_refs(self) -> None:
        logger = logging.getLogger("tests.iob_legacy_parser")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = iob_parser.parse(str(SAMPLE_PDF), logger)

        self.assertEqual(len(records), 394)

        cheque_row = next(
            record
            for record in records
            if record["Date"] == "13/03/2026" and record["Details"] == "RITA SUDHA WILSON"
        )
        self.assertEqual(cheque_row["Cheque No"], "350947")
        self.assertEqual(cheque_row["Debit"], 200000.0)
        self.assertIsNone(cheque_row["Credit"])
        self.assertEqual(cheque_row["Balance"], 8748.41)

        feb_row = next(
            record
            for record in records
            if record["Date"] == "11/02/2026" and record["Details"] == "RITA SUDHA WILSON"
        )
        self.assertEqual(feb_row["Cheque No"], "350949")


if __name__ == "__main__":
    unittest.main()
