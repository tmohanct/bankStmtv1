from __future__ import annotations

import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

from final_excel_builder import (
    INDIAN_NUMBER_FORMAT_NO_DECIMAL,
    MONTH_LABEL_FILL,
    MONTH_DR_CR_FOOTNOTE,
    _apply_month_dr_cr_style,
    _build_month_dr_cr_sheet,
    _format_month_dr_cr_chart_label,
    _patch_month_dr_cr_chart_xml,
)


class MonthDrCrSheetTests(unittest.TestCase):
    def test_format_month_dr_cr_chart_label_uses_k_and_l_suffixes(self) -> None:
        self.assertEqual(_format_month_dr_cr_chart_label(3600), "3.6k")
        self.assertEqual(_format_month_dr_cr_chart_label(130000), "1.3L")
        self.assertEqual(_format_month_dr_cr_chart_label(950), "950.0")
        self.assertEqual(_format_month_dr_cr_chart_label(None), "")

    def test_build_month_dr_cr_sheet_transposes_months_to_rows_with_total(self) -> None:
        statement_df = pd.DataFrame(
            [
                {"Sno": 1, "Date": "01/03/2025", "Debit": 10000.0, "Credit": 0.0, "Balance": 50000.0},
                {"Sno": 2, "Date": "15/03/2025", "Debit": 0.0, "Credit": 15000.0, "Balance": 65000.0},
                {"Sno": 3, "Date": "20/03/2025", "Debit": 4000.0, "Credit": 0.0, "Balance": 61000.0},
                {"Sno": 4, "Date": "05/04/2025", "Debit": 7000.0, "Credit": 0.0, "Balance": 54000.0},
                {"Sno": 5, "Date": "10/04/2025", "Debit": 0.0, "Credit": 8000.0, "Balance": 62000.0},
            ]
        )

        result = _build_month_dr_cr_sheet(statement_df)

        self.assertListEqual(
            result.columns.tolist(),
            ["Yr-Month", "Dr", "Cr", "Net", "EOM Balance", "#.Of.Dr", "#.Of.Cr", "Avg.Dr", "Avg.Cr"],
        )
        self.assertEqual(
            result.to_dict(orient="records"),
            [
                {
                    "Yr-Month": "25-Mar",
                    "Dr": 14000.0,
                    "Cr": 15000.0,
                    "Net": 1000.0,
                    "EOM Balance": 61000.0,
                    "#.Of.Dr": 2,
                    "#.Of.Cr": 1,
                    "Avg.Dr": 7000.0,
                    "Avg.Cr": 15000.0,
                },
                {
                    "Yr-Month": "25-Apr",
                    "Dr": 7000.0,
                    "Cr": 8000.0,
                    "Net": 1000.0,
                    "EOM Balance": 62000.0,
                    "#.Of.Dr": 1,
                    "#.Of.Cr": 1,
                    "Avg.Dr": 7000.0,
                    "Avg.Cr": 8000.0,
                },
                {
                    "Yr-Month": "Total",
                    "Dr": 21000.0,
                    "Cr": 23000.0,
                    "Net": 2000.0,
                    "EOM Balance": "",
                    "#.Of.Dr": 3,
                    "#.Of.Cr": 2,
                    "Avg.Dr": 7000.0,
                    "Avg.Cr": 11500.0,
                },
            ],
        )

    def test_apply_month_dr_cr_style_keeps_month_column_fill_and_alternates_value_rows(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "Yr-Month": "25-Mar",
                    "Dr": 14000.0,
                    "Cr": 15000.0,
                    "Net": 1000.0,
                    "EOM Balance": 61000.0,
                    "#.Of.Dr": 1,
                    "#.Of.Cr": 1,
                    "Avg.Dr": 10000.0,
                    "Avg.Cr": 15000.0,
                },
                {
                    "Yr-Month": "25-Apr",
                    "Dr": 7000.0,
                    "Cr": 8000.0,
                    "Net": 1000.0,
                    "EOM Balance": 62000.0,
                    "#.Of.Dr": 1,
                    "#.Of.Cr": 1,
                    "Avg.Dr": 7000.0,
                    "Avg.Cr": 8000.0,
                },
                {
                    "Yr-Month": "Total",
                    "Dr": 21000.0,
                    "Cr": 23000.0,
                    "Net": 2000.0,
                    "EOM Balance": None,
                    "#.Of.Dr": 2,
                    "#.Of.Cr": 2,
                    "Avg.Dr": 8500.0,
                    "Avg.Cr": 11500.0,
                },
            ]
        )
        workbook = Workbook()
        ws = workbook.active
        ws.title = "month_dr_cr"

        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        _apply_month_dr_cr_style(workbook, "month_dr_cr")

        self.assertEqual(ws["A2"].fill.fgColor.rgb, MONTH_LABEL_FILL.fgColor.rgb)
        self.assertEqual(ws["A3"].fill.fgColor.rgb, MONTH_LABEL_FILL.fgColor.rgb)
        self.assertNotEqual(ws["B2"].fill.fgColor.rgb, ws["B3"].fill.fgColor.rgb)
        self.assertNotEqual(ws["A2"].fill.fgColor.rgb, ws["B2"].fill.fgColor.rgb)
        self.assertEqual(ws["A4"].value, "Total")
        self.assertTrue(pd.isna(ws["E4"].value) or ws["E4"].value == "")
        self.assertEqual(ws["I4"].value, 11500.0)
        self.assertEqual(ws["E2"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)
        self.assertEqual(ws["H2"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)
        self.assertEqual(ws["I2"].number_format, INDIAN_NUMBER_FORMAT_NO_DECIMAL)
        self.assertEqual(ws["A6"].value, MONTH_DR_CR_FOOTNOTE)
        self.assertEqual(ws["J2"].value, "14.0k")
        self.assertEqual(ws["K3"].value, "8.0k")
        self.assertTrue(ws.column_dimensions["J"].hidden)
        self.assertTrue(ws.column_dimensions["K"].hidden)
        self.assertEqual(len(ws._charts), 1)
        self.assertEqual(ws._charts[0].anchor, "A11")
        self.assertEqual(ws._charts[0].legend.position, "r")
        self.assertIsNone(ws._charts[0].title)
        self.assertIsNone(ws._charts[0].y_axis.title)
        self.assertEqual(ws._charts[0].x_axis.tickLblPos, "low")
        self.assertFalse(ws._charts[0].x_axis.delete)
        self.assertFalse(ws._charts[0].y_axis.delete)
        self.assertEqual(ws._charts[0].dLbls.numFmt, "0.0")
        self.assertEqual(ws._charts[0].y_axis.majorGridlines.spPr.ln.prstDash, "sysDot")
        self.assertEqual(ws.auto_filter.ref, "A1:I4")

    def test_patch_month_dr_cr_chart_xml_uses_compact_label_ranges_and_right_legend(self) -> None:
        df = pd.DataFrame(
            [
                {
                    "Yr-Month": "25-Nov",
                    "Dr": 733733.0,
                    "Cr": 832762.0,
                    "Net": 99029.0,
                    "EOM Balance": -398638.0,
                    "#.Of.Dr": 24,
                    "#.Of.Cr": 52,
                    "Avg.Dr": 30569.0,
                    "Avg.Cr": 16015.0,
                },
                {
                    "Yr-Month": "25-Dec",
                    "Dr": 694034.0,
                    "Cr": 588511.0,
                    "Net": -105523.0,
                    "EOM Balance": -504161.0,
                    "#.Of.Dr": 25,
                    "#.Of.Cr": 43,
                    "Avg.Dr": 27757.0,
                    "Avg.Cr": 13686.0,
                },
                {
                    "Yr-Month": "Total",
                    "Dr": 1427767.0,
                    "Cr": 1421273.0,
                    "Net": -6494.0,
                    "EOM Balance": None,
                    "#.Of.Dr": 49,
                    "#.Of.Cr": 95,
                    "Avg.Dr": 29138.0,
                    "Avg.Cr": 14961.0,
                },
            ]
        )
        workbook = Workbook()
        ws = workbook.active
        ws.title = "month_dr_cr"

        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)

        _apply_month_dr_cr_style(workbook, "month_dr_cr")

        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as handle:
            temp_path = Path(handle.name)

        try:
            workbook.save(temp_path)
            logger = type("Logger", (), {"info": lambda *args, **kwargs: None, "warning": lambda *args, **kwargs: None})()
            _patch_month_dr_cr_chart_xml(temp_path, "month_dr_cr", ["25-Nov", "25-Dec"], logger)

            with zipfile.ZipFile(temp_path) as archive:
                chart_xml = next(
                    archive.read(name).decode("utf-8")
                    for name in archive.namelist()
                    if name.startswith("xl/charts/chart") and name.endswith(".xml")
                )

            self.assertNotIn("Debit / Credit by Month", chart_xml)
            self.assertNotIn(">Amount<", chart_xml)
            self.assertIn('<legendPos val="r"', chart_xml)
            self.assertIn("showDataLabelsRange", chart_xml)
            self.assertIn("'month_dr_cr'!$J$2:$J$3", chart_xml)
            self.assertIn("'month_dr_cr'!$K$2:$K$3", chart_xml)
            self.assertIn(">25-Nov<", chart_xml)
            self.assertIn(">25-Dec<", chart_xml)
        finally:
            temp_path.unlink(missing_ok=True)


if __name__ == "__main__":
    unittest.main()
