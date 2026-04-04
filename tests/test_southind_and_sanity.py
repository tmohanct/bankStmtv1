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

from parsers.southind_parser import _PendingRecord, _WordLine
from run import collect_negative_balance_rows, report_negative_balance_rows


class SouthIndRegressionTests(unittest.TestCase):
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
