from __future__ import annotations

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import indian_parser


class IndianParserTests(unittest.TestCase):
    def test_new_layout_ignores_wrapped_description_dash(self) -> None:
        pending = indian_parser.PendingRecord(
            date_text="Feb 09 2026",
            lines=[
                "SBIN0000862/Preetha K",
                "N/XXXXX84922/shalnikavi.6",
                "-",
                "2@okaxis/UPI/6406225232",
                "13/UPI Payment /BRANCH :",
                "ATM SERVICE BRANCH",
                "-",
                "INR 5,000.00",
                "INR 2,127,904.26 CR",
            ],
        )

        record, balance = indian_parser._finalize_record(pending, previous_balance=2122904.26)

        self.assertIsNotNone(record)
        self.assertEqual(record["Date"], "09/02/2026")
        self.assertIsNone(record["Debit"])
        self.assertEqual(record["Credit"], 5000.0)
        self.assertEqual(record["Balance"], 2127904.26)
        self.assertEqual(balance, 2127904.26)

    def test_final_page_summary_totals_do_not_replace_last_transaction(self) -> None:
        pending = indian_parser.PendingRecord(
            date_text="Jun 19 2026",
            lines=[
                "INW_CLG :ClgInwPr:",
                "DEIVA TRADERS",
                ",ChqNo:452689, /BRANCH",
                ": SERVICE BRANCH",
                "(CHENNAI)",
                "INR 75,000.00",
                "-",
                "INR 6,897,863.30 CR",
                "Ending Balance",
                "INR 6,897,863.30",
                "CR",
                "Total",
                "INR 70,296,242.96",
                "INR 71,882,479.00",
                "INR 6,897,863.30 CR(Rupees Sixty Eight Lakh)",
            ],
        )

        record, balance = indian_parser._finalize_record(pending, previous_balance=6972863.3)

        self.assertIsNotNone(record)
        self.assertEqual(record["Date"], "19/06/2026")
        self.assertEqual(record["Debit"], 75000.0)
        self.assertIsNone(record["Credit"])
        self.assertEqual(record["Balance"], 6897863.3)
        self.assertEqual(balance, 6897863.3)
        self.assertIn("DEIVA TRADERS", record["Details"])
        self.assertNotIn("Total", record["Details"])
        self.assertNotIn("Rupees", record["Details"])

    def test_existing_day_month_year_date_order_still_parses(self) -> None:
        pending = indian_parser.PendingRecord(
            date_text="01 Feb 2026",
            lines=[
                "Sample transfer",
                "-",
                "INR 1,000.00",
                "INR 11,000.00",
            ],
        )

        record, balance = indian_parser._finalize_record(pending, previous_balance=10000.0)

        self.assertIsNotNone(record)
        self.assertEqual(record["Date"], "01/02/2026")
        self.assertIsNone(record["Debit"])
        self.assertEqual(record["Credit"], 1000.0)
        self.assertEqual(record["Balance"], 11000.0)
        self.assertEqual(balance, 11000.0)


if __name__ == "__main__":
    unittest.main()
