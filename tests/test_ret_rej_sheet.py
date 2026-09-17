from __future__ import annotations

import logging
import tempfile
import unittest
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


from src.transform.analysis import (
    _build_repeat_sheet,
    _build_return_reject_sheet,
)
from src.export.final_excel_builder import (
    build_final_workbook,
)
from src.utils.statement_utils import OUTPUT_COLUMNS, clean_detail
from src.utils.pdf_status_reader import Header, identify_bank, positioned_rows


def _statement_frame() -> pd.DataFrame:
    rows = [
        {
            "Sno": 1,
            "Date": "01/01/2026",
            "Details": "Paid to Rita",
            "Detail_Clean": clean_detail("Paid to Rita"),
            "Cheque No": "",
            "Debit": 100.0,
            "Credit": 0.0,
            "Balance": 900.0,
            "Source": "sample.pdf",
        },
        {
            "Sno": 2,
            "Date": "02/01/2026",
            "Details": "BRN-OW RTN CLG: REJECT:238951:Funds insufficient",
            "Detail_Clean": clean_detail("BRN-OW RTN CLG: REJECT:238951:Funds insufficient"),
            "Cheque No": "",
            "Debit": 70000.0,
            "Credit": 0.0,
            "Balance": -5498055.96,
            "Source": "axis.pdf",
        },
        {
            "Sno": 3,
            "Date": "03/01/2026",
            "Details": "NEFT/RETURN/AXODH01008167350/SP 01/S Kamatchi/A PRI",
            "Detail_Clean": clean_detail("NEFT/RETURN/AXODH01008167350/SP 01/S Kamatchi/A PRI"),
            "Cheque No": "",
            "Debit": 0.0,
            "Credit": 2720.0,
            "Balance": -5489277.06,
            "Source": "axis2.pdf",
        },
        {
            "Sno": 4,
            "Date": "04/01/2026",
            "Details": "CHQRETURNCHGSINCLGST141125-CDT25326 37321333",
            "Detail_Clean": clean_detail("CHQRETURNCHGSINCLGST141125-CDT25326 37321333"),
            "Cheque No": "0000000000000111",
            "Debit": 59.0,
            "Credit": 0.0,
            "Balance": 69056.16,
            "Source": "hdfc.pdf",
        },
        {
            "Sno": 5,
            "Date": "05/01/2026",
            "Details": "Cheque return (Issued):500384:Exceeds Arrangement",
            "Detail_Clean": clean_detail("Cheque return (Issued):500384:Exceeds Arrangement"),
            "Cheque No": "500384",
            "Debit": 0.0,
            "Credit": 83000.0,
            "Balance": -3263976.99,
            "Source": "indus.pdf",
        },
    ]
    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


class PdfAccountSummaryTests(unittest.TestCase):
    def test_statement_title_does_not_become_bank_name(self) -> None:
        header = Header(positioned_rows([(0, 0, 160, 12, "ACCOUNT STATEMENT")]), width=600)
        profile, evidence = identify_bank(header)

        self.assertEqual(profile, {})
        self.assertEqual(evidence, "")

class RepeatAmountSheetTests(unittest.TestCase):
    def test_same_amount_sorts_by_cheque_then_date_when_cheque_missing(self) -> None:
        frame = pd.DataFrame(
            [
                {"Sno": 1, "Date": "05/01/2026", "Details": "blank later", "Detail_Clean": "blank later", "Cheque No": "", "Debit": 100.0, "Credit": 0.0, "Balance": 900.0, "Source": "sample.pdf"},
                {"Sno": 2, "Date": "02/01/2026", "Details": "cheque high", "Detail_Clean": "cheque high", "Cheque No": "20", "Debit": 100.0, "Credit": 0.0, "Balance": 800.0, "Source": "sample.pdf"},
                {"Sno": 3, "Date": "01/01/2026", "Details": "blank earlier", "Detail_Clean": "blank earlier", "Cheque No": "", "Debit": 100.0, "Credit": 0.0, "Balance": 700.0, "Source": "sample.pdf"},
                {"Sno": 4, "Date": "03/01/2026", "Details": "cheque low", "Detail_Clean": "cheque low", "Cheque No": "10", "Debit": 100.0, "Credit": 0.0, "Balance": 600.0, "Source": "sample.pdf"},
            ],
            columns=OUTPUT_COLUMNS,
        )

        result = _build_repeat_sheet(frame, "Debit")

        self.assertEqual(result["Sno"].tolist(), [4, 2, 3, 1])

class ReturnRejectSheetTests(unittest.TestCase):
    def test_build_return_reject_sheet_keeps_only_related_rows(self) -> None:
        result = _build_return_reject_sheet(_statement_frame())

        self.assertListEqual(result["Sno"].tolist(), [2, 5])
        self.assertListEqual(
            result["Source"].tolist(),
            ["axis.pdf", "indus.pdf"],
        )

    def test_final_workbook_writes_cheque_only_ret_rej_after_statement(self) -> None:
        statement_df = _statement_frame()
        logger = logging.getLogger("tests.ret_rej_sheet")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        with tempfile.TemporaryDirectory(prefix="ret_rej_") as temp_directory:
            temp_root = Path(temp_directory)
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

            workbook = load_workbook(final_path, data_only=True)
            self.assertListEqual(workbook.sheetnames[:3], ["PDF_Status", "Statement", "Ret_Rej"])
            pdf_status_ws = workbook["PDF_Status"]
            self.assertEqual(pdf_status_ws["A1"].value, "Customer Name:")
            self.assertEqual(pdf_status_ws["C8"].value, "Status")
            self.assertEqual(pdf_status_ws["C9"].value, "WARNING")
            self.assertEqual(pdf_status_ws["C9"].fill.fgColor.rgb, "00FFF2CC")
            workbook.close()

            ret_rej_df = pd.read_excel(final_path, sheet_name="Ret_Rej")
            self.assertListEqual(ret_rej_df["Sno"].tolist(), [2, 5])
            exported_statement = pd.read_excel(final_path, sheet_name="Statement")
            self.assertEqual(exported_statement["Sno"].tolist(), [1, 2, 3, 4, 5])


if __name__ == "__main__":
    unittest.main()
