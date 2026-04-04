from __future__ import annotations

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

from parser_helpers import build_record
from utils import normalize_cheque_number, records_to_dataframe


class ChequeNormalizationTests(unittest.TestCase):
    def test_text_value_is_removed(self) -> None:
        self.assertEqual(normalize_cheque_number("Customer Payment", "Customer Payment"), "")

    def test_upi_reference_digits_are_removed(self) -> None:
        details = "UPI/1180804477 UPI/118080447740/23:38:55/UPI/sandeshyadav87052@o"
        self.assertEqual(normalize_cheque_number("80447740", details), "")

    def test_leading_zero_cheque_is_preserved(self) -> None:
        self.assertEqual(normalize_cheque_number("00001234", "CHQ DEPOSIT"), "00001234")

    def test_all_zero_cheque_is_removed(self) -> None:
        self.assertEqual(normalize_cheque_number("0000000000000000", "CASHDEPOSITBY-SELVARAJ-TENKASI"), "")

    def test_integer_like_float_cheque_is_normalized(self) -> None:
        self.assertEqual(normalize_cheque_number("541.0", "Cheque deposit"), "541")

    def test_build_record_uses_shared_cheque_normalization(self) -> None:
        record = build_record(
            date_text="01/04/2026",
            details="UPI/1180804477 UPI/118080447740/23:38:55/UPI/sandeshyadav87052@o",
            cheque_no="80447740",
            debit=7000.0,
            balance=1000.0,
            date_formats=("%d/%m/%Y",),
        )
        self.assertEqual(record["Cheque No"], "")

    def test_records_to_dataframe_sanitizes_direct_parser_rows(self) -> None:
        frame = records_to_dataframe(
            [
                {
                    "Sno": 1,
                    "Date": "01/04/2026",
                    "Details": "UTR 1234567890 FOR TRANSFER",
                    "Detail_Clean": "UTR1234567890FORTRANSFER",
                    "Cheque No": "12345678",
                    "Debit": None,
                    "Credit": 100.0,
                    "Balance": 500.0,
                    "Source": "sample.pdf",
                },
                {
                    "Sno": 2,
                    "Date": "02/04/2026",
                    "Details": "CHQ DEP",
                    "Detail_Clean": "CHQDEP",
                    "Cheque No": "00001234",
                    "Debit": 100.0,
                    "Credit": None,
                    "Balance": 400.0,
                    "Source": "sample.pdf",
                },
                {
                    "Sno": 3,
                    "Date": "03/04/2026",
                    "Details": "CASHDEPOSITBY-SELVARAJ-TENKASI",
                    "Detail_Clean": "CASHDEPOSITBYSELVARAJTENKASI",
                    "Cheque No": "0000000000000000",
                    "Debit": None,
                    "Credit": 7715.0,
                    "Balance": 7943.88,
                    "Source": "hdfc.pdf",
                },
            ]
        )

        self.assertEqual(frame.loc[0, "Cheque No"], "")
        self.assertEqual(frame.loc[1, "Cheque No"], "00001234")
        self.assertEqual(frame.loc[2, "Cheque No"], "")


if __name__ == "__main__":
    unittest.main()
