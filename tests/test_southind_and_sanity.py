from __future__ import annotations

import contextlib
import io
import logging
import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

from parsers.southind_parser import _PendingRecord, _WordLine, _parse_amount, _parse_slno_layout_lines
from run import collect_negative_balance_rows, report_negative_balance_rows


class SouthIndRegressionTests(unittest.TestCase):
    def test_new_slno_layout_captures_detail_lines_around_anchor(self) -> None:
        lines = [
            _WordLine(
                y_center=302.1,
                words=[
                    (7.0, "SlNo"),
                    (42.0, "Transaction"),
                    (86.0, "Date"),
                    (182.0, "Particulars"),
                    (349.0, "Withdrawals"),
                    (442.0, "Deposits"),
                    (530.0, "Balance"),
                    (561.1, "Amount"),
                ],
            ),
            _WordLine(y_center=322.8, words=[(182.0, "UPI/PYTM/345780267437/")]),
            _WordLine(y_center=333.0, words=[(182.0, "BSNL"), (205.1, "Landline")]),
            _WordLine(
                y_center=342.8,
                words=[
                    (7.0, "1"),
                    (42.0, "01-Sep-2025"),
                    (112.0, "01-Sep-2025"),
                    (182.0, "Bill/Oid207350/PTM3F94B0"),
                    (349.0, "943.00"),
                    (530.0, "87,469.56"),
                ],
            ),
            _WordLine(y_center=353.1, words=[(182.0, "C856E940BC87142BCA89")]),
            _WordLine(y_center=363.2, words=[(182.0, "8FD560/paybil3066@pay")]),
            _WordLine(y_center=388.0, words=[(182.0, "MOB/309117804897/Bill")]),
            _WordLine(
                y_center=392.8,
                words=[
                    (7.0, "2"),
                    (42.0, "01-Sep-2025"),
                    (112.0, "01-Sep-2025"),
                    (349.0, "50,000.00"),
                    (530.0, "37,469.56"),
                ],
            ),
            _WordLine(y_center=398.1, words=[(182.0, "Payment/IMPS/")]),
        ]
        records: list[dict[str, object]] = []

        _parse_slno_layout_lines(lines, records)

        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["Date"], "01/09/2025")
        self.assertEqual(
            records[0]["Details"],
            "UPI/PYTM/345780267437/ BSNL Landline Bill/Oid207350/PTM3F94B0 C856E940BC87142BCA89 8FD560/paybil3066@pay",
        )
        self.assertEqual(records[0]["Debit"], 943.0)
        self.assertEqual(records[0]["Balance"], 87469.56)
        self.assertEqual(records[1]["Details"], "MOB/309117804897/Bill Payment/IMPS/")
        self.assertEqual(records[1]["Debit"], 50000.0)

    def test_new_slno_layout_handles_value_date_fallback_and_extra_decimal_amounts(self) -> None:
        lines = [
            _WordLine(y_center=755.8, words=[(182.0, "MOB/406614314325/Own")]),
            _WordLine(
                y_center=780.8,
                words=[
                    (7.0, "2653"),
                    (42.0, "099-Mar-2026"),
                    (112.0, "09-Mar-2026"),
                    (349.0, "20.000.00"),
                    (530.0, "1,029.039.11"),
                ],
            ),
            _WordLine(y_center=790.8, words=[(182.0, "Account/IMPS/")]),
        ]
        records: list[dict[str, object]] = []

        _parse_slno_layout_lines(lines, records)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Date"], "09/03/2026")
        self.assertEqual(records[0]["Debit"], 20000.0)
        self.assertEqual(records[0]["Balance"], 1029039.11)
        self.assertEqual(_parse_amount("20.000.00"), 20000.0)

    def test_new_slno_layout_keeps_trailing_detail_with_previous_row(self) -> None:
        lines = [
            _WordLine(y_center=305.9, words=[(182.0, "IMPS/FDRL/409113753248")]),
            _WordLine(y_center=316.0, words=[(182.0, "/DEEPAK"), (218.8, "M/ELOHIM")]),
            _WordLine(
                y_center=319.8,
                words=[
                    (7.0, "3696"),
                    (42.0, "12-Jun-2026"),
                    (112.0, "12-Jun-2026"),
                    (530.0, "7,72,011.74"),
                ],
            ),
            _WordLine(y_center=324.8, words=[(442.0, "4,000.00"), (182.0, "TRANSPORT/RBI234b29a")]),
            _WordLine(y_center=336.1, words=[(182.0, "7e16b43ea8f34cb67cce857")]),
            _WordLine(y_center=346.2, words=[(182.0, "f9#919944128777#D")]),
            _WordLine(
                y_center=369.8,
                words=[
                    (7.0, "3697"),
                    (42.0, "12-Jun-2026"),
                    (112.0, "12-Jun-2026"),
                    (182.0, "MOB/359114508789/salary/"),
                    (349.0, "24,500.00"),
                    (530.0, "7,47,511.74"),
                ],
            ),
            _WordLine(y_center=381.1, words=[(182.0, "IMPS/")]),
        ]
        records: list[dict[str, object]] = []

        _parse_slno_layout_lines(lines, records)

        self.assertEqual(len(records), 2)
        self.assertIn("f9#919944128777#D", records[0]["Details"])
        self.assertNotIn("f9#919944128777#D", records[1]["Details"])
        self.assertEqual(records[1]["Details"], "MOB/359114508789/salary/ IMPS/")

    def test_shifted_balance_credit_row_is_not_dropped(self) -> None:
        pending = _PendingRecord(date_text="19-11-25")
        pending.add_line(
            _WordLine(
                y_center=682.01,
                words=[
                    (23.0, "19-11-25"),
                    (87.0, "NEFT:SAMRUTHI"),
                    (154.1, "FINCREDIT"),
                    (199.0, "PRIVATE"),
                ],
            )
        )
        pending.add_line(
            _WordLine(
                y_center=686.67,
                words=[
                    (446.7, "1,67,853.00"),
                    (520.3, "1,67,879.24Cr"),
                ],
            )
        )
        pending.add_line(_WordLine(y_center=691.33, words=[(87.0, "LIMITED")]))

        record = pending.finalize()

        self.assertIsNotNone(record)
        assert record is not None
        self.assertEqual(record["Date"], "19/11/2025")
        self.assertEqual(record["Details"], "NEFT:SAMRUTHI FINCREDIT PRIVATE LIMITED")
        self.assertIsNone(record["Debit"])
        self.assertEqual(record["Credit"], 167853.0)
        self.assertEqual(record["Balance"], 167879.24)

    def test_shifted_balance_debit_row_with_detail_continuation_is_not_dropped(self) -> None:
        pending = _PendingRecord(date_text="20-11-25")
        pending.add_line(
            _WordLine(
                y_center=426.01,
                words=[
                    (23.0, "20-11-25"),
                    (87.0, "UPI/BARB/RRN-"),
                ],
            )
        )
        pending.add_line(
            _WordLine(
                y_center=430.67,
                words=[
                    (363.4, "10,000.00"),
                    (520.3, "1,45,935.24Cr"),
                ],
            )
        )
        pending.add_line(
            _WordLine(
                y_center=435.33,
                words=[
                    (87.0, "569069524043/RAVIKUMAR/UPI"),
                ],
            )
        )

        record = pending.finalize()

        self.assertIsNotNone(record)
        assert record is not None
        self.assertEqual(record["Date"], "20/11/2025")
        self.assertEqual(record["Details"], "UPI/BARB/RRN- 569069524043/RAVIKUMAR/UPI")
        self.assertEqual(record["Debit"], 10000.0)
        self.assertIsNone(record["Credit"])
        self.assertEqual(record["Balance"], 145935.24)


class BalanceSanityTests(unittest.TestCase):
    def test_collect_negative_balance_rows_returns_only_negative_rows(self) -> None:
        records = [
            {"Sno": 1, "Balance": 100.0},
            {"Sno": 2, "Balance": -50.0},
            {"Sno": 3, "Balance": None},
            {"Sno": 4, "Balance": "-25.5"},
        ]

        negative_rows = collect_negative_balance_rows(records)

        self.assertEqual([row["Sno"] for row in negative_rows], [2, 4])

    def test_report_negative_balance_rows_prints_first_three_and_last_three(self) -> None:
        records = [
            {
                "Sno": index,
                "Date": f"2025-01-{index:02d}",
                "Details": f"row-{index}",
                "Balance": float(-index),
            }
            for index in range(1, 7)
        ]
        logger = logging.getLogger("tests.balance_sanity")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        buffer = io.StringIO()
        with contextlib.redirect_stdout(buffer):
            report_negative_balance_rows(
                records=records,
                file_name="sample.pdf",
                bank_key="axis",
                logger=logger,
            )

        output = buffer.getvalue()
        self.assertIn("**** WARNING ****", output)
        self.assertIn("-ve balance found sample first 3 records", output)
        self.assertIn("-ve balance found sample last 3 records", output)
        self.assertIn("please cross check with pdf", output)
        self.assertIn("Sno=1", output)
        self.assertIn("Sno=3", output)
        self.assertIn("Sno=4", output)
        self.assertIn("Sno=6", output)


if __name__ == "__main__":
    unittest.main()
