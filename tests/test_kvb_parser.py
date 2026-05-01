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
    def test_extract_trailing_cheque_number_from_clg_detail(self) -> None:
        details, cheque_no = kvb_parser._extract_trailing_cheque_no("To Clg:A RAJENDRAN - Federal Bank 000116")

        self.assertEqual(details, "To Clg:A RAJENDRAN - Federal Bank")
        self.assertEqual(cheque_no, "000116")

    def test_parse_tokenized_row_extracts_trailing_cheque_number(self) -> None:
        parsed = kvb_parser._parse_tokenized_text_row(
            [
                "08-12-2025 09:15:01",
                "08-12-2025",
                "To Clg:A RAJENDRAN - Federal Bank 000116",
                "7,000.00",
                "133.88",
            ]
        )

        self.assertIsNotNone(parsed)
        assert parsed is not None
        self.assertEqual(parsed.body_text, "To Clg:A RAJENDRAN - Federal Bank")
        self.assertEqual(parsed.cheque_no, "000116")

        record, _ = kvb_parser._finalize_record(parsed, 7133.88, kvb_parser.OCR_DATE_FORMATS)
        self.assertEqual(record["Cheque No"], "000116")

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

    def test_parse_tokenized_row_with_month_date_and_explicit_debit_credit_columns(self) -> None:
        parsed = kvb_parser._parse_tokenized_text_row(
            [
                "29-SEP-2025",
                "06:17:46",
                "29-SEP-2025",
                "IMPS-527206487277-VODAFONE IDEA",
                "LTD-YESB-xxxxxxxxxxx6186-sa",
                "-",
                "1,000.00",
                "0.00",
                "200.00",
            ]
        )

        self.assertIsNotNone(parsed)
        assert parsed is not None
        self.assertEqual(parsed.date_text, "29-SEP-2025")
        self.assertEqual(parsed.body_text, "IMPS-527206487277-VODAFONE IDEA LTD-YESB-xxxxxxxxxxx6186-sa")
        self.assertEqual(parsed.debit_text, "1,000.00")
        self.assertEqual(parsed.credit_text, "0.00")
        self.assertEqual(parsed.balance_text, "200.00")

        record, next_balance = kvb_parser._finalize_record(parsed, 1200.0, kvb_parser.OCR_DATE_FORMATS)
        self.assertEqual(record["Date"], "29/09/2025")
        self.assertEqual(record["Debit"], 1000.0)
        self.assertIsNone(record["Credit"])
        self.assertEqual(record["Balance"], 200.0)
        self.assertEqual(next_balance, 200.0)

    def test_parse_tokenized_row_accepts_leading_decimal_balance(self) -> None:
        parsed = kvb_parser._parse_tokenized_text_row(
            [
                "23-OCT-2025",
                "06:56:34",
                "23-OCT-2025",
                "MB-WITHIN-DR:XXXX0397-",
                "CR:XXXX8895-942508231025794809",
                "-",
                "749.00",
                "0.00",
                ".24",
            ]
        )

        self.assertIsNotNone(parsed)
        assert parsed is not None

        record, next_balance = kvb_parser._finalize_record(parsed, 749.24, kvb_parser.OCR_DATE_FORMATS)
        self.assertEqual(record["Date"], "23/10/2025")
        self.assertEqual(record["Debit"], 749.0)
        self.assertIsNone(record["Credit"])
        self.assertEqual(record["Balance"], 0.24)
        self.assertEqual(next_balance, 0.24)

    def test_tokenized_row_start_detection_accepts_split_rows(self) -> None:
        self.assertTrue(kvb_parser._looks_like_tokenized_row_start("02/04/2026", "10:22:11"))
        self.assertTrue(kvb_parser._looks_like_tokenized_row_start("02/04/2026", "02/04/2026"))
        self.assertTrue(kvb_parser._looks_like_tokenized_row_start("02-04-2026 10:22:11", ""))
        self.assertTrue(kvb_parser._looks_like_tokenized_row_start("29-SEP-2025", "06:15:14"))
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

    def test_tokenized_parser_handles_uppercase_headers_and_brought_forward_rows(self) -> None:
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
                "ACCOUNT STATEMENT",
                "Txn Date",
                "Value Date",
                "Particulars",
                "Ref. No.",
                "Debit",
                "Credit",
                "Balance",
                "29-SEP-2025",
                "29-SEP-2025",
                "B/F...",
                "-",
                "-",
                "-",
                "0.00",
                "29-SEP-2025",
                "06:15:14",
                "29-SEP-2025",
                "UPI-CR-527223189379-ENAYATHULLA",
                "S-KVBL-1708155000018895-UPI",
                "527223189379",
                "0.00",
                "1,000.00",
                "1,000.00",
                "1",
                "ACCOUNT SUMMARY",
                "Current Balance",
                "1,000.00",
                "Note: This is a computer-generated report and does not require signature.",
            ]
        )

        logger = logging.getLogger("tests.kvb.uppercase")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        with patch.object(
            kvb_parser.fitz,
            "open",
            return_value=FakeDocument([FakePage(page_1)]),
        ):
            records = kvb_parser._parse_tokenized_text_statement("dummy.pdf", logger)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Date"], "29/09/2025")
        self.assertEqual(records[0]["Details"], "UPI-CR-527223189379-ENAYATHULLA S-KVBL-1708155000018895-UPI")
        self.assertIsNone(records[0]["Debit"])
        self.assertEqual(records[0]["Credit"], 1000.0)
        self.assertEqual(records[0]["Balance"], 1000.0)

    def test_ocr_parser_splits_rows_with_ocr_punctuation_noise(self) -> None:
        class FakePage:
            def __init__(self, text: str) -> None:
                self._text = text

        class FakeDocument(list):
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb) -> bool:
                return False

        page_112 = "\n".join(
            [
                "30-01-2026 11:20:44 30-01-2026 | 1763 IMPS-603011766230-SKPLANNERS. 50,000.00 2,338.75}",
                "AND ASSOCIATES-INDB-",
                "XXXXXXXX5555-Y ANNAIKKAL CHIT",
                "PAYMENT",
                "Page No. 112",
            ]
        )
        page_113 = "\n".join(
            [
                "30-01-2026 18:51:36 30-01-2026 | 1763 IMPS-603018300866-SAI 1,00,000.00 1,02,338.75}",
                "KARTHICK ENTERPRISES-",
                "XxXXXXxX9161-Repayment",
                "30-01-2026 18:55:57 30-01-2026 | 2101 KVBLH00253837785- DE NOBILI 1,00,000.00) 2,338.75}",
                "HR SEC SCHOOL-408461073-",
                "PAYMENT RETURN",
                "31-01-2026 10:33:47 31-01-2026 | 2101 NEFT CR-HDFC0000001-SAI 4,50,000.00 4,52,338.75|",
                "KARTHICK ENTERPRISES-Sai",
                "Karthick planners and associates-",
                "HDFCH00769268380",
                "31-01-2026 11:31:16 31-01-2026 | 2101 KVBLH00253860868- 1,00,000.00 3,52,338.75|",
                "N.RAMESH-50100271033149-EMI",
                "PURPOSE",
            ]
        )

        logger = logging.getLogger("tests.kvb.ocr")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        with patch.object(kvb_parser, "_configure_tesseract", return_value="tesseract.exe"):
            with patch.object(kvb_parser, "_render_page_text", side_effect=lambda page: page._text):
                with patch.object(
                    kvb_parser.fitz,
                    "open",
                    return_value=FakeDocument([FakePage(page_112), FakePage(page_113)]),
                ):
                    records = kvb_parser._parse_ocr_statement("dummy.pdf", logger)

        self.assertEqual(len(records), 5)
        self.assertEqual(sum(row["Date"] == "30/01/2026" for row in records), 3)

        self.assertEqual(
            records[1]["Details"],
            "IMPS-603018300866-SAI KARTHICK ENTERPRISES- XxXXXXxX9161-Repayment",
        )
        self.assertIsNone(records[1]["Debit"])
        self.assertEqual(records[1]["Credit"], 100000.0)
        self.assertEqual(records[1]["Balance"], 102338.75)

        self.assertEqual(
            records[2]["Details"],
            "KVBLH00253837785- DE NOBILI HR SEC SCHOOL-408461073- PAYMENT RETURN",
        )
        self.assertEqual(records[2]["Debit"], 100000.0)
        self.assertIsNone(records[2]["Credit"])
        self.assertEqual(records[2]["Balance"], 2338.75)

        self.assertEqual(
            records[3]["Details"],
            "NEFT CR-HDFC0000001-SAI KARTHICK ENTERPRISES-Sai Karthick planners and associates- HDFCH00769268380",
        )
        self.assertIsNone(records[3]["Debit"])
        self.assertEqual(records[3]["Credit"], 450000.0)
        self.assertEqual(records[3]["Balance"], 452338.75)

        self.assertEqual(
            records[4]["Details"],
            "KVBLH00253860868- N.RAMESH-50100271033149-EMI PURPOSE",
        )
        self.assertEqual(records[4]["Debit"], 100000.0)
        self.assertIsNone(records[4]["Credit"])
        self.assertEqual(records[4]["Balance"], 352338.75)

    def test_ocr_parser_falls_back_to_value_date_when_txn_date_is_ocr_corrupted(self) -> None:
        class FakePage:
            def __init__(self, text: str) -> None:
                self._text = text

        class FakeDocument(list):
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb) -> bool:
                return False

        page_text = "\n".join(
            [
                "41-11-2025 14:42:41 11-11-2025 | 2101 KVBLH00248205479- 53,000.00 52,477.33]",
                "GURUSAMY-40569019245-DGL",
                "STAFF SALARY",
                "47-11-2025 17:16:03 17-11-2025 | 2101 KVBLH00248621099- 15,000.00 11,261.89]",
                "VETRIVELMURUGAN-27660200000.",
                "CEMENT PAYMENT",
                "18-11-2025 10:26:59 18-11-2025] 1865 000000000000 | CASH DEP-TP-K ANNAMALAI- 20,000.00 31,261.89]",
                "KOVILOOR",
            ]
        )

        logger = logging.getLogger("tests.kvb.ocr-date")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        with patch.object(kvb_parser, "_configure_tesseract", return_value="tesseract.exe"):
            with patch.object(kvb_parser, "_render_page_text", side_effect=lambda page: page._text):
                with patch.object(
                    kvb_parser.fitz,
                    "open",
                    return_value=FakeDocument([FakePage(page_text)]),
                ):
                    records = kvb_parser._parse_ocr_statement("dummy.pdf", logger)

        self.assertEqual(len(records), 3)
        self.assertEqual(records[0]["Date"], "11/11/2025")
        self.assertEqual(records[0]["Details"], "KVBLH00248205479- GURUSAMY-40569019245-DGL STAFF SALARY")
        self.assertEqual(records[0]["Debit"], 53000.0)
        self.assertEqual(records[1]["Date"], "17/11/2025")
        self.assertEqual(records[1]["Details"], "KVBLH00248621099- VETRIVELMURUGAN-27660200000. CEMENT PAYMENT")
        self.assertEqual(records[1]["Debit"], 15000.0)
        self.assertEqual(records[2]["Date"], "18/11/2025")
        self.assertEqual(records[2]["Details"], "CASH DEP-TP-K ANNAMALAI- KOVILOOR")
        self.assertEqual(records[2]["Credit"], 20000.0)

    def test_ocr_parser_handles_scanned_legacy_kvb_layout(self) -> None:
        class FakePage:
            def __init__(self, text: str) -> None:
                self._text = text

        class FakeDocument(list):
            def __enter__(self):
                return self

            def __exit__(self, exc_type, exc, tb) -> bool:
                return False

        page_text = "\n".join(
            [
                "Karur Vysya Bank",
                "THE KARUR VYSYA BANK LTD.",
                "BRANCH : SRIRANGAM-TRICHY",
                "STATEMENT OF ACCOUNT",
                "INDIAN RUPEES",
                "Messrs. SHREE VEEJAY PROPERTY DEVELOPERS PRIVATE LIMITED CA-KVB-ECONOMY",
                "NO 43 SHREE VEEJAY PROPERTY Period from: 01-JAN-2026 .",
                "DEVELOPERS PRIVATE LIMITED NO 43 Period To: 31-MAR-2026",
                "NEHRU SECOND STREET EDAMALAIPATTI Account Number 1276135000007451",
                "CKYC NO: XXXxXXXXXXX1409",
                "TIRUCHIRAPPALLI 620012",
                "JXN DT VALUE_DT BRN DESCRIPTION REFERENCE DEBITS CREDITS BALANCE",
                "01/01/26 01/01/26 B/F... 32,31,589.38",
                "02/01/26 02/01/26 1276 100% To DEV ACNT DR - 127 000000000000 32,21,000.00 10,589.38",
                "6135000005280 - SHREE VEE",
                "JAY PROPERTY DEVELOPERS P",
                "RIVATE :",
                "09/01/26 09/01/26 2101 NEFT CR-SBINO010205-VENKA 31, 43,340.00 31,53,929.38",
                "TASWAMI KRISHNASWAMI VASA",
                "NISHRER VEE TAY PROPERTY-S",
                "BIN 54",
                "23/01/26 23/01/26 2101 NEFT CR-BKID0008299-BANK 90,109.00 32, 44,038.38",
                "OF INDIA-SHREE VEEJAY PRO",
                "PERTY DEVELOPERS -BKIDY260",
                "28/01/26 28/01/26 2101 NEFT CR-SBINO010205-RAMSA 5,75,000.00 38,19,038.38",
                "NKAR S-SREE VEE JAY PROP",
                "ERTY DEVELOPERS-SBIN4260281.",
                "31/01/26 31/01/26 1276 FT _-100% To DEV ACNT DR - 000000000000 =38,00,000.00 19,038.38",
                "1276135000005280 - SHREE",
                "VEEJAY PROPERTY DEVELOPE",
                "RS PRIVATE",
                "05/02/26 05/02/26 2101 NEFT CR-SBIN0010205-SUDHA 46,53,000.00 46,72,038.38",
                "KAR RATHINAM-VEEJAY PROP",
                "ERTY DEVELOPERS-SBINS2603",
                "11/02/26 11/02/26 2101 NEFT CR-SBINO010205-VENKA 5,82,100.00 52,54,138.38",
                "TASWAMI KRISHNASWAMI VASA",
                "N-SHREE VEEJAY PROPERTY D",
                "EVELOPE-SBIN226042599936",
                "02/03/26 02/03/26 1276 100% TRANSFER TO DEVELOPE 000000000000 52,40,000.00 14,138.38",
                "R_ ACCOUNT DR - 1276135000",
                "005280 - SHREE VEEJAY PRO",
                "PERTY DEVELOPERS PRIVATE",
                "28/03/26 28/03/26 1276 SMS Charges for MAR2026 4.13 14,134.25",
                "page: 1",
                "Opening Balance < 32,31,589.38 _",
                "Total Credit Amount 2 90, 43,549.00 Credit Count :5",
                "Total Debit Amount : 1,22,61,004.13 Debit-Count :4",
                "Closing Balance : 14,134.25",
                "Net Available Balance as of 30-MAR-2026 is : 14134.25 .",
                "ACRONYMS DESCRIPTIONS",
                "IFSC Code : KvBL0001276",
                "MICR_Code : 620053005",
                "Regd. Office : Karur Vysya Bank,Central office,Erode Road, Karur-639002(Tamil Nadu) www.kvb.bank.in",
            ]
        )

        logger = logging.getLogger("tests.kvb.ocr-legacy")
        logger.handlers.clear()
        logger.addHandler(logging.NullHandler())

        with patch.object(kvb_parser, "_configure_tesseract", return_value="tesseract.exe"):
            with patch.object(kvb_parser, "_render_page_text", side_effect=lambda page: page._text):
                with patch.object(
                    kvb_parser.fitz,
                    "open",
                    return_value=FakeDocument([FakePage(page_text)]),
                ):
                    records = kvb_parser._parse_ocr_statement("dummy.pdf", logger)

        self.assertEqual(len(records), 9)
        self.assertEqual(sum(row["Debit"] is not None for row in records), 4)
        self.assertEqual(sum(row["Credit"] is not None for row in records), 5)

        self.assertEqual(records[0]["Date"], "02/01/2026")
        self.assertTrue(records[0]["Details"].startswith("100% To DEV ACNT DR - 127 000000000000"))
        self.assertIn("6135000005280 - SHREE VEE", records[0]["Details"])
        self.assertEqual(records[0]["Debit"], 3221000.0)
        self.assertIsNone(records[0]["Credit"])
        self.assertEqual(records[0]["Balance"], 10589.38)

        self.assertEqual(records[1]["Date"], "09/01/2026")
        self.assertTrue(records[1]["Details"].startswith("NEFT CR-SBINO010205-VENKA"))
        self.assertIsNone(records[1]["Debit"])
        self.assertEqual(records[1]["Credit"], 3143340.0)
        self.assertEqual(records[1]["Balance"], 3153929.38)

        self.assertEqual(records[-1]["Date"], "28/03/2026")
        self.assertEqual(records[-1]["Details"], "SMS Charges for MAR2026")
        self.assertEqual(records[-1]["Debit"], 4.13)
        self.assertIsNone(records[-1]["Credit"])
        self.assertEqual(records[-1]["Balance"], 14134.25)
        self.assertNotIn("Opening Balance", records[-1]["Details"])
        self.assertNotIn("ACRONYMS DESCRIPTIONS", records[-1]["Details"])


if __name__ == "__main__":
    unittest.main()
