from __future__ import annotations

import unittest
import logging
from unittest.mock import MagicMock, patch


from src.parsers import indian_parser


class IndianParserTests(unittest.TestCase):
    def _parse_post_date_page(self, transaction_lines, progress_cb=None):
        page = MagicMock()
        page.get_text.return_value = "\n".join([
            "INDIAN BANK", "Post Date", "Brought Forward", "1000.00cr",
            *transaction_lines, "CLOSING BALANCE :", "900.00Cr",
        ])
        pdf = MagicMock()
        pdf.__enter__.return_value = pdf
        pdf.__bool__.return_value = True
        pdf.__getitem__.return_value = page
        pdf.__iter__.side_effect = lambda: iter((page,))
        with patch.object(indian_parser.fitz, "open", return_value=pdf):
            return indian_parser.parse("statement.pdf", logging.getLogger(__name__), progress_cb)

    def test_post_date_return_notice_and_out_of_order_interest(self) -> None:
        progress = MagicMock()
        rows = self._parse_post_date_page([
            "30/06/26", "30/06/26", "Paid to vendor", "100.00", "900.00Cr",
            "30/06/26", "30/06/26", "ClgInwRet Chq:059010",
            "Amt:100000.00 Rtn:12 Drawer s/", "0.00",
            "01/07/26", "01/07/26", "TRANSFER FROM customer", "50.00", "975.00Cr",
            "30/06/26", "30/06/26", "CREDIT INTEREST", "25.00", "925.00Cr",
            "03/07/26", "03/07/26", "ATM withdrawal", "75.00", "900.00Cr",
        ], progress)

        self.assertEqual(len(rows), 5)
        self.assertEqual([row["Sno"] for row in rows], [1, 2, 3, 4, 5])
        self.assertEqual([row["Date"] for row in rows],
                         ["30/06/2026", "30/06/2026", "30/06/2026", "01/07/2026", "03/07/2026"])
        self.assertEqual([row["Balance"] for row in rows], [900, 0, 925, 975, 900])
        self.assertIn("ClgInwRet", rows[1]["Details"])
        self.assertEqual(rows[2]["Details"], "CREDIT INTEREST")
        self.assertEqual(sum(row["Debit"] or 0 for row in rows), 175)
        self.assertEqual(sum(row["Credit"] or 0 for row in rows), 75)
        self.assertEqual([call.args[0] for call in progress.call_args_list], [1, 2, 3, 4, 5])

    def test_retains_nonposting_return_in_statement(self) -> None:
        page = MagicMock()
        page.get_text.return_value = "\n".join([
            "INDIAN BANK", "Post Date", "Brought Forward", "1000.00Cr",
            "30/06/26", "30/06/26", "ClgInwRet Chq:059010",
            "Amt:100000.00 Rtn:12 Drawer s/", "0.00",
            "30/06/26", "30/06/26", "Paid to vendor", "100.00", "900.00Cr",
            "CLOSING BALANCE :", "900.00Cr",
        ])
        pdf = MagicMock()
        pdf.__enter__.return_value = pdf
        pdf.__bool__.return_value = True
        pdf.__getitem__.return_value = page
        pdf.__iter__.side_effect = lambda: iter((page,))
        with patch.object(indian_parser.fitz, "open", return_value=pdf):
            rows = indian_parser.parse(
                "statement.pdf", logging.getLogger(__name__),
            )
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Date"], "30/06/2026")
        self.assertEqual(rows[0]["Cheque No"], "059010")
        self.assertIn("Amt:100000.00 Rtn:12", rows[0]["Details"])
        self.assertEqual(rows[0]["Sno"], 1)
        self.assertEqual(rows[0]["Debit"], 0)
        self.assertEqual(rows[0]["Credit"], 0)
        self.assertEqual(rows[0]["Balance"], 0)

    def test_blank_and_zero_returns_and_rejections_are_kept(self) -> None:
        for details in ("cLgInWrEt Chq:059010", "CHEQUE REJECTED Chq:059010", "chq rtn 059010"):
            for cells in ([], ["0.00"], ["-", "-"], ["0.00", "0.00", "1000.00Cr"]):
                with self.subTest(details=details, cells=cells):
                    record, balance = indian_parser._finalize_post_date_record(
                        ["30/06/26", "30/06/26", details, *cells], 1000, 1,
                    )
                    self.assertEqual(record["Details"], details)
                    self.assertEqual(record["Debit"], 0)
                    self.assertEqual(record["Credit"], 0)
                    self.assertEqual(balance, 1000)
                    record, balance = indian_parser._finalize_record(
                        indian_parser.PendingRecord("30 Jun 2026", [details, *cells]), 1000,
                    )
                    self.assertIsNotNone(record)
                    self.assertEqual(balance, 1000)

    def test_post_date_malformed_rows_are_not_treated_as_notices(self) -> None:
        for details, amounts in [
            (["Ordinary transfer"], ["0.00"]),
            (["ClgInwRet Chq:059010", "Amt:100000.00 Rtn:12"], ["100.00", "0.00"]),
            (["ClgInwRet Chq:059010", "Amt:100000.00 Rtn:12"], ["12.00"]),
        ]:
            with self.subTest(details=details, amounts=amounts):
                with self.assertRaisesRegex(ValueError, "unrecognized transaction row"):
                    self._parse_post_date_page(["30/06/26", "30/06/26", *details, *amounts])

    def test_unreadable_return_is_not_silently_skipped(self) -> None:
        with self.assertRaisesRegex(ValueError, "refusing to omit the entry"):
            indian_parser._finalize_record(
                indian_parser.PendingRecord("30 Jun 2026", ["CHQ REJECTED 059010", "100.00"]), 1000,
            )

    def test_post_date_posted_return_is_kept(self) -> None:
        rows = self._parse_post_date_page([
            "30/06/26", "30/06/26", "ClgInwRet Chq:059010",
            "Amt:100.00 Rtn:12", "100.00", "900.00Cr",
        ])
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["Debit"], 100)

    def test_post_date_reordering_does_not_hide_balance_errors(self) -> None:
        with self.assertRaisesRegex(ValueError, "does not match balance change"):
            self._parse_post_date_page([
                "01/07/26", "01/07/26", "Transfer", "50.00", "900.00Cr",
                "30/06/26", "30/06/26", "CREDIT INTEREST", "25.00", "1025.00Cr",
            ])

    def test_post_date_layout_parses_pages_and_cheque_number(self) -> None:
        first_page = MagicMock()
        first_page.get_text.return_value = "\n".join([
            "INDIAN BANK", "Post Date", "Value", "Date", "Details", "Debit", "Credit", "Balance",
            "Brought Forward", "17440.29cr",
            "04/11/25", "04/11/25", "Mandate charges", "135.70", "17304.59Cr",
            "05/11/25", "05/11/25", "UPI Payment", "8000.00", "25304.59Cr",
            "Carried Forward", "25304.59Cr", "Statement Summary",
        ])
        second_page = MagicMock()
        second_page.get_text.return_value = "\n".join([
            "INDIAN BANK", "Post Date", "Value", "Date", "Details", "Debit", "Credit", "Balance",
            "Brought Forward", "25304.59cr",
            "06/11/25", "06/11/25", "Paid to vendor", "684847", "25000.00", "304.59Cr",
            "CLOSING BALANCE :", "304.59Cr", "Statement Summary",
        ])
        pdf = MagicMock()
        pdf.__enter__.return_value = pdf
        pdf.__bool__.return_value = True
        pdf.__getitem__.return_value = first_page
        pdf.__iter__.side_effect = lambda: iter((first_page, second_page))

        with patch.object(indian_parser.fitz, "open", return_value=pdf):
            rows = indian_parser.parse("statement.pdf", logging.getLogger(__name__))

        self.assertEqual(len(rows), 3)
        self.assertEqual([row["Sno"] for row in rows], [1, 2, 3])
        self.assertEqual(rows[0]["Date"], "04/11/2025")
        self.assertEqual(rows[0]["Debit"], 135.70)
        self.assertEqual(rows[1]["Credit"], 8000.00)
        self.assertEqual(rows[2]["Cheque No"], "684847")
        self.assertEqual(rows[2]["Balance"], 304.59)

    def test_post_date_layout_rejects_balance_mismatch(self) -> None:
        with self.assertRaisesRegex(ValueError, "does not match balance change"):
            indian_parser._finalize_post_date_record(
                ["04/11/25", "04/11/25", "Mandate charges", "135.70", "17300.00Cr"],
                17440.29,
                1,
            )

    def test_post_date_layout_keeps_non_cheque_reference_in_details(self) -> None:
        record, _ = indian_parser._finalize_post_date_record(
            ["10/11/25", "10/11/25", "TRANSFER FROM customer", "147197", "30000.00", "40000.00Cr"],
            10000.00,
            1,
        )
        self.assertEqual(record["Cheque No"], "")
        self.assertIn("147197", record["Details"])

    def test_post_date_final_summary_totals(self) -> None:
        page = MagicMock()
        page.get_text.return_value = "\n".join([
            "CLOSING BALANCE :", "2214.18Cr", "Statement Summary",
            "Dr. Count:73 Cr. Count:124", "4546801.11    4531575.00",
            "*** END OF STATEMENT ***",
        ])
        pdf = MagicMock()
        pdf.__enter__.return_value = pdf
        pdf.__bool__.return_value = True
        pdf.__getitem__.return_value = page
        with patch.object(indian_parser.fitz, "open", return_value=pdf):
            summary = indian_parser.extract_summary_metrics("statement.pdf", logging.getLogger(__name__))
        self.assertEqual(summary, {
            "transaction_count": 197.0,
            "total_debit": 4546801.11,
            "total_credit": 4531575.0,
        })

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
