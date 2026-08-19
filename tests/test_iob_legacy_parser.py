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
NEW_LAYOUT_SAMPLE_PDF = PROJECT_ROOT / "input" / "IOI.pdf"
COD_LAYOUT_SAMPLE_PDF = PROJECT_ROOT / "input" / "iob.pdf"


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

    @unittest.skipUnless(NEW_LAYOUT_SAMPLE_PDF.is_file(), "IOI sample PDF is required for this regression test.")
    def test_ioi_pdf_new_layout_is_parsed(self) -> None:
        logger = logging.getLogger("tests.iob_legacy_parser.ioi")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = iob_parser.parse(str(NEW_LAYOUT_SAMPLE_PDF), logger)

        self.assertEqual(len(records), 1377)
        self.assertEqual(records[0]["Date"], "01/07/2026")
        self.assertEqual(records[0]["Details"], "BY CASH SELF")
        self.assertIsNone(records[0]["Debit"])
        self.assertEqual(records[0]["Credit"], 250000.0)
        self.assertEqual(records[0]["Balance"], 689188.48)

        self.assertEqual(records[-1]["Date"], "02/01/2026")
        self.assertEqual(records[-1]["Credit"], 22680.0)
        self.assertEqual(records[-1]["Balance"], 447231.54)

    @unittest.skipUnless(COD_LAYOUT_SAMPLE_PDF.is_file(), "iob sample PDF is required for this regression test.")
    def test_iob_pdf_cod_layout_details_are_parsed(self) -> None:
        logger = logging.getLogger("tests.iob_legacy_parser.cod_layout")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = iob_parser.parse(str(COD_LAYOUT_SAMPLE_PDF), logger)

        self.assertEqual(len(records), 510)
        self.assertEqual(records[0]["Date"], "01/11/2025")
        self.assertEqual(records[0]["Details"], "CHEQUE BOOK ISSUE CHARGES")
        self.assertEqual(records[0]["Debit"], 236.0)
        self.assertIsNone(records[0]["Credit"])
        self.assertEqual(records[0]["Balance"], 4287.7)

        cheque_row = next(record for record in records if record["Details"] == "P WILSON")
        self.assertEqual(cheque_row["Cheque No"], "000042")
        self.assertEqual(cheque_row["Debit"], 50000.0)
        self.assertEqual(cheque_row["Balance"], 49181.81)


if __name__ == "__main__":
    unittest.main()
