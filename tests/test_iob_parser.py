from __future__ import annotations

import sys
import unittest
from pathlib import Path

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))

from parsers.iob_parser import IOBParser
from parsers.parser_registry import list_supported_banks

SAMPLE_PDF = PROJECT_ROOT / "input" / "AKILANMANIVANNAN.pdf"


@unittest.skipUnless(SAMPLE_PDF.is_file(), "IOB sample PDF is required for this regression test.")
class IOBParserRegressionTests(unittest.TestCase):
    def test_akilanmanivannan_pdf_is_parsed(self) -> None:
        parser = IOBParser()

        parsed = parser.parse(pdf_path=SAMPLE_PDF, rules_df=pd.DataFrame())

        self.assertEqual(len(parsed), 394)

        first = parsed.iloc[0]
        self.assertEqual(first["Date"], "27/03/2026")
        self.assertEqual(first["ValueDate"], "27/03/2026")
        self.assertIn("UPI/645275501663/DR/ V SELVAN/YES /UPI", first["Narration"])
        self.assertEqual(first["Debit"], 25.0)
        self.assertTrue(pd.isna(first["Credit"]))
        self.assertEqual(first["Balance"], 217392.35)
        self.assertEqual(first["Txn_Ref"], "S67581822")
        self.assertEqual(first["Page"], 1)
        self.assertEqual(first["Account_Number"], "247701000009845")

        last = parsed.iloc[-1]
        self.assertEqual(last["Date"], "28/12/2025")
        self.assertEqual(last["ValueDate"], "28/12/2025")
        self.assertEqual(last["Narration"], "UPI/102314923670/DR/ HYUNDAI MOTOR/HDF/COLLECT")
        self.assertEqual(last["Debit"], 2499.0)
        self.assertTrue(pd.isna(last["Credit"]))
        self.assertEqual(last["Balance"], 69950.35)
        self.assertEqual(last["Txn_Ref"], "S63465454")
        self.assertEqual(last["Page"], 13)

    def test_iob_is_registered_in_active_parser_registry(self) -> None:
        self.assertIn("iob", list_supported_banks())


if __name__ == "__main__":
    unittest.main()
