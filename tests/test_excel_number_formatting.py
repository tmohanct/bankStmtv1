from __future__ import annotations

import logging
import shutil
import sys
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

from final_excel_builder import INDIAN_NUMBER_FORMAT_NO_DECIMAL, build_final_workbook
from utils import OUTPUT_COLUMNS, clean_detail


def _statement_frame() -> pd.DataFrame:
    rows = [
        {
            "Sno": 1,
            "Date": "01/01/2026",
            "Details": "BRN-OW RTN CLG: REJECT:238951:Funds insufficient",
            "Detail_Clean": clean_detail("BRN-OW RTN CLG: REJECT:238951:Funds insufficient"),
            "Cheque No": "",
            "Debit": 70000.4,
            "Credit": 0.0,
            "Balance": -5498055.96,
            "Source": "axis.pdf",
        },
        {
            "Sno": 2,
            "Date": "02/01/2026",
            "Details": "Transfer from Sudha",
            "Detail_Clean": clean_detail("Transfer from Sudha"),
            "Cheque No": "",
            "Debit": 0.0,
            "Credit": 2720.5,
            "Balance": -5495335.46,
            "Source": "sample.pdf",
        },
    ]
    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


class ExcelNumberFormattingTests(unittest.TestCase):
    def test_final_workbook_rounds_money_cells_and_uses_indian_no_decimal_format(self) -> None:
        statement_df = _statement_frame()
        logger = logging.getLogger("tests.excel_number_formatting")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        temp_root = PROJECT_ROOT / "output" / "_excel_number_formatting_test"
        shutil.rmtree(temp_root, ignore_errors=True)
        temp_root.mkdir(parents=True, exist_ok=True)

        try:
            rules_path = temp_root / "Rules.xlsx"
            output_dir = temp_root / "output"
            pd.DataFrame(columns=["Category", "subCategory", "SheetName"]).to_excel(rules_path, index=False)

            final_path = build_final_workbook(
                statement_df=statement_df,
                rules_path=rules_path,
                output_dir=output_dir,
                pdf_stem="sample",
                logger=logger,
            )

            workbook = load_workbook(final_path)

            statement_ws = workbook["Statement"]
            self.assertEqual(statement_ws["F2"].value, 70000)
            self.assertEqual(statement_ws["F2"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)
            self.assertEqual(statement_ws["H2"].value, -5498056)
            self.assertEqual(statement_ws["H2"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)
            self.assertEqual(statement_ws["G3"].value, 2721)
            self.assertEqual(statement_ws["G3"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)

            ret_rej_ws = workbook["Ret_Rej"]
            self.assertEqual(ret_rej_ws["F2"].value, 70000)
            self.assertEqual(ret_rej_ws["H2"].value, -5498056)
            self.assertEqual(ret_rej_ws["F2"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)

            month_ws = workbook["month_dr_cr"]
            self.assertEqual(month_ws["B2"].value, 70000)
            self.assertEqual(month_ws["C2"].value, 2721)
            self.assertEqual(month_ws["D2"].value, -67280)
            self.assertEqual(month_ws["E2"].value, -5495335)
            self.assertEqual(month_ws["H2"].value, 70000)
            self.assertEqual(month_ws["I2"].value, 2721)
            self.assertEqual(month_ws["B2"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)
            self.assertEqual(month_ws["I2"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)

            workbook.close()
        finally:
            shutil.rmtree(temp_root, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
