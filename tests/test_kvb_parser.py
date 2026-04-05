from __future__ import annotations

import logging
import sys
import unittest
from unittest.mock import patch
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import kvb_parser


class KVBParserUnitTests(unittest.TestCase):
    def test_parse_tokenized_row_with_datetime_start(self) -> None:
        parsed = kvb_parser._parse_tokenized_text_row(
            [
                "01-04-2026 09:15:01",
                "01-04-2026",
                "0001",
                "UPI/ALICE/REF123",
                "5,000.00",
                "15,000.00",
            ]
        )

        self.assertIsNotNone(parsed)
        assert parsed is not None
        self.assertEqual(parsed.date_text, "01-04-2026")
        self.assertEqual(parsed.body_text, "UPI/ALICE/REF123")
        self.assertEqual(parsed.cheque_no, "")
        self.assertEqual(parsed.amount_text, "5,000.00")
        self.assertEqual(parsed.balance_text, "15,000.00")

        record, next_balance = kvb_parser._finalize_record(parsed, 10000.0, kvb_parser.OCR_DATE_FORMATS)
        self.assertEqual(record["Date"], "01/04/2026")
        self.assertIsNone(record["Debit"])
        self.assertEqual(record["Credit"], 5000.0)
        self.assertEqual(record["Balance"], 15000.0)
        self.assertEqual(next_balance, 15000.0)

    def test_parse_tokenized_row_with_split_date_and_time(self) -> None:
        parsed = kvb_parser._parse_tokenized_text_row(
            [
                "02/04/2026",
                "10:22:11",
                "02/04/2026",
                "0002",
                "987654321",
                "IMPS-TRANSFER",
                "2,500.00",
                "12,500.00CR",
            ]
        )

        self.assertIsNotNone(parsed)
        assert parsed is not None
        self.assertEqual(parsed.date_text, "02/04/2026")
        self.assertEqual(parsed.body_text, "IMPS-TRANSFER")
        self.assertEqual(parsed.cheque_no, "987654321")
        self.assertEqual(parsed.amount_text, "2,500.00")
        self.assertEqual(parsed.balance_text, "12,500.00CR")

        record, next_balance = kvb_parser._finalize_record(parsed, 15000.0, kvb_parser.OCR_DATE_FORMATS)
        self.assertEqual(record["Date"], "02/04/2026")
        self.assertEqual(record["Debit"], 2500.0)
        self.assertIsNone(record["Credit"])
        self.assertEqual(record["Balance"], 12500.0)
        self.assertEqual(next_balance, 12500.0)

    def test_tokenized_row_start_detection_accepts_split_rows(self) -> None:
        self.assertTrue(kvb_parser._looks_like_tokenized_row_start("02/04/2026", "10:22:11"))
        self.assertTrue(kvb_parser._looks_like_tokenized_row_start("02/04/2026", "02/04/2026"))
        self.assertTrue(kvb_parser._looks_like_tokenized_row_start("02-04-2026 10:22:11", ""))
        self.assertFalse(kvb_parser._looks_like_tokenized_row_start("IMPS-TRANSFER", "2,500.00"))

    def test_tokenized_parser_drops_page_header_and_summary_carryover(self) -> None:
        class FakePage:
            def __init__(self, text: str) -> None:
                self._text = text

            def get_text(self, mode: str) -> str:
                self.mode = mode
                return self._text

        class FakeDocument(list):
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb) -> bool:
                return False

        page_1 = "\n".join(
            [
                "Account Statement",
                "01/03/2025",
                "09:00:00",
                "01/03/2025",
                "NEFT CR-YESAP50601134495-ONE 97 COMMUNIC",
                "728.00",
                "1534.88",
                "02/03/2025",
                "10:00:00",
                "02/03/2025",
                "IMPS-506113271038-Shafee-CIUB-xxxxxxxxxx",
                "7,000.00",
                "133.88",
                "Karur Vysya Bank does not ask for personal security details like your Internet banking or phone banking passwords on the email, phone or otherwise.",
                "Never disclose your passwords to anyone, even to the bank's staff.",
            ]
        )
        page_2 = "\n".join(
            [
                "Account Statement",
                "Mr SHAFEE VADHUDULLA",
                "Acc.No. : 1170155000336492",
                "Customer ID",
                "Account Summary",
                "806.88",
                "1,16,27,680.01",
                "02/03/2025",
                "10:00:01",
                "02/03/2025",
                "IMPS CHARGES",
                "5.90",
                "127.98",
            ]
        )
        logger = logging.getLogger("tests.kvb")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        with patch.object(
            kvb_parser.fitz,
            "open",
            return_value=FakeDocument([FakePage(page_1), FakePage(page_2)]),
        ):
            records = kvb_parser._parse_tokenized_text_statement("dummy.pdf", logger)

        self.assertEqual(len(records), 3)
        self.assertEqual(records[0]["Details"], "NEFT CR-YESAP50601134495-ONE 97 COMMUNIC")
        self.assertEqual(records[0]["Credit"], 728.0)
        self.assertEqual(records[0]["Balance"], 1534.88)

        self.assertEqual(records[1]["Details"], "IMPS-506113271038-Shafee-CIUB-xxxxxxxxxx")
        self.assertEqual(records[1]["Debit"], 7000.0)
        self.assertEqual(records[1]["Balance"], 133.88)
        self.assertNotIn("Account Statement", records[1]["Details"])
        self.assertNotIn("Account Summary", records[1]["Details"])

        self.assertEqual(records[2]["Details"], "IMPS CHARGES")
        self.assertEqual(records[2]["Debit"], 5.9)
        self.assertEqual(records[2]["Balance"], 127.98)


if __name__ == "__main__":
    unittest.main()
