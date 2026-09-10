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
NEW_LAYOUT_SAMPLE_PDF = PROJECT_ROOT / "input" / "IOI.pdf"
COD_LAYOUT_SAMPLE_PDF = PROJECT_ROOT / "input" / "iob.pdf"
TEXT_LAYOUT_SAMPLE_PDF = PROJECT_ROOT / "input" / "BARATHI FOOD PRODUCT_IOB.pdf"


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


@unittest.skipUnless(NEW_LAYOUT_SAMPLE_PDF.is_file(), "IOI sample PDF is required for this regression test.")
class IOBNewLayoutParserRegressionTests(unittest.TestCase):
    def test_ioi_pdf_is_parsed(self) -> None:
        parser = IOBParser()

        parsed = parser.parse(pdf_path=NEW_LAYOUT_SAMPLE_PDF, rules_df=pd.DataFrame())

        self.assertEqual(len(parsed), 1377)

        first = parsed.iloc[0]
        self.assertEqual(first["Date"], "01/07/2026")
        self.assertEqual(first["ValueDate"], "01/07/2026")
        self.assertEqual(first["Narration"], "BY CASH SELF")
        self.assertTrue(pd.isna(first["Debit"]))
        self.assertEqual(first["Credit"], 250000.0)
        self.assertEqual(first["Balance"], 689188.48)
        self.assertEqual(first["Account_Number"], "264702000000191")

        last = parsed.iloc[-1]
        self.assertEqual(last["Date"], "02/01/2026")
        self.assertEqual(last["ValueDate"], "02/01/2026")
        self.assertEqual(last["Credit"], 22680.0)
        self.assertEqual(last["Balance"], 447231.54)


@unittest.skipUnless(COD_LAYOUT_SAMPLE_PDF.is_file(), "iob sample PDF is required for this regression test.")
class IOBCodLayoutParserRegressionTests(unittest.TestCase):
    def test_iob_pdf_cod_layout_details_are_parsed(self) -> None:
        parser = IOBParser()

        parsed = parser.parse(pdf_path=COD_LAYOUT_SAMPLE_PDF, rules_df=pd.DataFrame())

        self.assertEqual(len(parsed), 510)

        first = parsed.iloc[0]
        self.assertEqual(first["Date"], "01/11/2025")
        self.assertEqual(first["ValueDate"], "01/11/2025")
        self.assertEqual(first["Narration"], "CHEQUE BOOK ISSUE CHARGES")
        self.assertEqual(first["Debit"], 236.0)
        self.assertTrue(pd.isna(first["Credit"]))
        self.assertEqual(first["Balance"], 4287.7)
        self.assertEqual(first["Account_Number"], "219502000000395")

        cheque_row = parsed[parsed["Narration"].eq("P WILSON")].iloc[0]
        self.assertEqual(cheque_row["Txn_Ref"], "000042")
        self.assertEqual(cheque_row["Debit"], 50000.0)
        self.assertEqual(cheque_row["Balance"], 49181.81)


@unittest.skipUnless(
    TEXT_LAYOUT_SAMPLE_PDF.is_file(),
    "BARATHI FOOD PRODUCT IOB sample PDF is required for this regression test.",
)
class IOBTextLayoutParserRegressionTests(unittest.TestCase):
    def test_bharathi_food_product_text_layout_is_parsed(self) -> None:
        parser = IOBParser()

        parsed = parser.parse(pdf_path=TEXT_LAYOUT_SAMPLE_PDF, rules_df=pd.DataFrame())

        self.assertEqual(len(parsed), 639)

        first = parsed.iloc[0]
        self.assertEqual(first["Date"], "09/09/2026")
        self.assertEqual(first["ValueDate"], "09/09/2026")
        self.assertEqual(
            first["Narration"],
            "UPI/661809626565/DR/Euronet Servic/UTI/UPI",
        )
        self.assertEqual(first["Debit"], 350.9)
        self.assertTrue(pd.isna(first["Credit"]))
        self.assertEqual(first["Balance"], 21214.77)
        self.assertEqual(first["Account_Number"], "280702000000205")
        self.assertEqual(first["Page"], 1)

        credit_row = parsed.iloc[3]
        self.assertTrue(pd.isna(credit_row["Debit"]))
        self.assertEqual(credit_row["Credit"], 100000.0)
        self.assertEqual(credit_row["Balance"], 221916.57)

        last = parsed.iloc[-1]
        self.assertEqual(last["Date"], "10/03/2026")
        self.assertEqual(
            last["Narration"],
            "UPI/606911872091/DR/BHARATHIRAJA ./TMB/UPI",
        )
        self.assertEqual(last["Debit"], 600.0)
        self.assertEqual(last["Balance"], 28.46)
        self.assertEqual(last["Page"], 10)

        for row_index in range(len(parsed) - 1):
            current = parsed.iloc[row_index]
            older = parsed.iloc[row_index + 1]
            debit = 0.0 if pd.isna(current["Debit"]) else current["Debit"]
            credit = 0.0 if pd.isna(current["Credit"]) else current["Credit"]
            expected_balance = older["Balance"] - debit + credit
            self.assertAlmostEqual(
                current["Balance"],
                expected_balance,
                places=2,
                msg=f"Balance continuity failed at parsed row {row_index + 1}.",
            )


if __name__ == "__main__":
    unittest.main()
