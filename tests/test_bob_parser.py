from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import bob_parser

SAMPLE_PDF = PROJECT_ROOT / "input" / "Abdul.pdf"


@unittest.skipUnless(SAMPLE_PDF.is_file(), "Bank of Baroda sample PDF is required for this regression test.")
class BobParserRegressionTests(unittest.TestCase):
    def test_abdul_pdf_is_parsed(self) -> None:
        logger = logging.getLogger("tests.bob")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = bob_parser.parse(str(SAMPLE_PDF), logger)

        self.assertEqual(len(records), 555)

        first = records[0]
        self.assertEqual(first["Date"], "01/02/2026")
        self.assertEqual(
            first["Details"],
            "DIGITA-MUMBAI/ NEFT-YESAP60320062142-PAYTM PAYMENTS SERVICES LIMI",
        )
        self.assertEqual(first["Cheque No"], "")
        self.assertIsNone(first["Debit"])
        self.assertEqual(first["Credit"], 6590.0)
        self.assertEqual(first["Balance"], 1023495.41)

        row_28 = records[27]
        self.assertEqual(row_28["Date"], "02/02/2026")
        self.assertEqual(
            row_28["Details"],
            "UPI/1180804477 UPI/118080447740/23:38:55/UPI/sandeshyadav87052@o",
        )
        self.assertEqual(row_28["Cheque No"], "")
        self.assertEqual(row_28["Debit"], 7000.0)
        self.assertIsNone(row_28["Credit"])
        self.assertEqual(row_28["Balance"], 967425.41)
        self.assertNotIn("BARODA", row_28["Details"])
        self.assertNotIn("PARTICULARS", row_28["Details"])
        self.assertNotIn("HELPLINE", row_28["Details"])

        row_290 = records[289]
        self.assertEqual(row_290["Date"], "27/02/2026")
        self.assertEqual(
            row_290["Details"],
            "UPI/1192153408 UPI/119215340856/09:16:07/UPI/7904558662@ybl/UPI",
        )
        self.assertEqual(row_290["Cheque No"], "")
        self.assertEqual(row_290["Debit"], 10000.0)
        self.assertIsNone(row_290["Credit"])
        self.assertEqual(row_290["Balance"], 416556.06)
        self.assertNotIn("of account", row_290["Details"])

        row_295 = records[294]
        self.assertEqual(row_295["Date"], "27/02/2026")
        self.assertEqual(
            row_295["Details"],
            "RTGS-CIUBR5202 RTGS-CIUBR52026022700413379-PALANI CAPITAL",
        )
        self.assertEqual(row_295["Cheque No"], "")
        self.assertIsNone(row_295["Debit"])
        self.assertEqual(row_295["Credit"], 820000.0)
        self.assertEqual(row_295["Balance"], 1121056.06)

        last = records[-1]
        self.assertEqual(last["Date"], "23/03/2026")
        self.assertEqual(
            last["Details"],
            "UPI/8200163829 UPI/820016382946/23:22:39/UPI/8220965783@ibl/Paym",
        )
        self.assertEqual(last["Cheque No"], "")
        self.assertIsNone(last["Debit"])
        self.assertEqual(last["Credit"], 10000.0)
        self.assertEqual(last["Balance"], 642359.22)


if __name__ == "__main__":
    unittest.main()
