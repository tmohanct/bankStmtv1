from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import unionbank_parser


SAMPLE_PDF = PROJECT_ROOT / "input" / "SUBAM_TRADERS1.pdf"


@unittest.skipUnless(SAMPLE_PDF.is_file(), "Union Bank sample PDF is required for this regression test.")
class UnionBankParserTests(unittest.TestCase):
    def test_unionbank_parser_reads_subam_traders_statement(self) -> None:
        logger = logging.getLogger("tests.unionbank_parser")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        records = unionbank_parser.parse(str(SAMPLE_PDF), logger)

        self.assertEqual(len(records), 915)

        first = records[0]
        self.assertEqual(first["Sno"], 1)
        self.assertEqual(first["Date"], "01/08/2025")
        self.assertIn("UPIAR/109038657437", first["Details"])
        self.assertIn("Y95172730", first["Details"])
        self.assertEqual(first["Cheque No"], "Y95172730")
        self.assertEqual(first["Debit"], 4500.0)
        self.assertIsNone(first["Credit"])
        self.assertEqual(first["Balance"], 1774591.31)

        credit_row = records[16]
        self.assertEqual(credit_row["Date"], "05/08/2025")
        self.assertIn("RTGS:SOUTH INDIAN BANK", credit_row["Details"])
        self.assertEqual(credit_row["Cheque No"], "U73119237")
        self.assertIsNone(credit_row["Debit"])
        self.assertEqual(credit_row["Credit"], 4562304.0)
        self.assertEqual(credit_row["Balance"], 4682889.61)

        last = records[-1]
        self.assertEqual(last["Sno"], 915)
        self.assertEqual(last["Date"], "06/04/2026")
        self.assertIn("RTGS:LOGICAL AGRI BUSINESS", last["Details"])
        self.assertEqual(last["Cheque No"], "U45136202")
        self.assertIsNone(last["Debit"])
        self.assertEqual(last["Credit"], 1310200.0)
        self.assertEqual(last["Balance"], 1336883.87)


if __name__ == "__main__":
    unittest.main()
