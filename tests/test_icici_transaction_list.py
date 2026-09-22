from __future__ import annotations

import logging
from pathlib import Path
import unittest

from src.parsers import detector as bank_detector
from src.parsers import icici_parser
from src.utils.pdf_status_reader import read_first_page


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SAMPLE_PDF = PROJECT_ROOT / "input" / "VRS AGENCIES.pdf"


class IciciTransactionListRowTests(unittest.TestCase):
    def test_detects_headers_with_words_wrapped_across_lines(self) -> None:
        header = """
        DETAILED STATEMENT
        Transactions List - -VRS AGENCIES (INR) - 617705500043
        No. Transactio
        n ID Value Date Txn Posted Date ChequeNo. Description Cr/Dr
        Transaction Amount(INR) Available Balance(INR
        )
        """

        self.assertTrue(icici_parser._is_transaction_list_text_layout(header))

    def test_parses_posted_date_explicit_direction_and_wrapped_negative_balance(self) -> None:
        row = [
            "6",
            "S13510830",
            "01/04/2026",
            "02/04/2026 05:42:49 AM",
            "-",
            "617705500043:Int.Coll:02-03-2026 to 01-04-2026",
            "DR",
            "1,88,082.00",
            "-\n2,57,36,692.4\n3",
        ]

        parsed = icici_parser._transaction_list_record_from_row(
            row,
            logging.getLogger("tests.icici_transaction_list.row"),
        )

        self.assertIsNotNone(parsed)
        assert parsed is not None
        serial, record = parsed
        self.assertEqual(serial, 6)
        self.assertEqual(record["Date"], "02/04/2026")
        self.assertEqual(record["Debit"], 188082.0)
        self.assertIsNone(record["Credit"])
        self.assertEqual(record["Balance"], -25736692.43)
        self.assertEqual(record["Cheque No"], "")


@unittest.skipUnless(SAMPLE_PDF.is_file(), "VRS AGENCIES ICICI sample PDF is required.")
class IciciTransactionListRegressionTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.logger = logging.getLogger("tests.icici_transaction_list.regression")
        cls.logger.handlers.clear()
        cls.logger.addHandler(logging.NullHandler())
        cls.records = icici_parser.parse(str(SAMPLE_PDF), cls.logger)

    def test_auto_detects_image_logo_statement_from_native_layout(self) -> None:
        self.assertEqual(bank_detector.detect_bank_from_pdf(SAMPLE_PDF, self.logger), "icici")

    def test_extracts_transaction_list_identity_without_ocr(self) -> None:
        header = read_first_page(SAMPLE_PDF, allow_ocr=False)

        self.assertEqual(header.values["Bank Name"], "ICICI Bank")
        self.assertEqual(header.values["Customer Name"], "VRS AGENCIES")
        self.assertEqual(header.values["Account Number"], "617705500043")
        self.assertEqual(header.values["Address"], "")

    def test_parses_all_rows_and_reconciles_running_balances(self) -> None:
        self.assertEqual(len(self.records), 1244)
        self.assertEqual(self.records[0]["Date"], "01/04/2026")
        self.assertEqual(self.records[0]["Credit"], 6144.0)
        self.assertIsNone(self.records[0]["Debit"])
        self.assertEqual(self.records[0]["Balance"], -25479975.43)

        self.assertEqual(self.records[-1]["Date"], "19/09/2026")
        self.assertEqual(self.records[-1]["Credit"], 160700.0)
        self.assertEqual(self.records[-1]["Balance"], -25620487.73)

        for previous, current in zip(self.records, self.records[1:]):
            expected = round(
                previous["Balance"]
                + (current["Credit"] or 0.0)
                - (current["Debit"] or 0.0),
                2,
            )
            self.assertAlmostEqual(current["Balance"], expected, places=2)


if __name__ == "__main__":
    unittest.main()
