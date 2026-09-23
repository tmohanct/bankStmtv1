"""Cheque return rows must survive missing or unusual numeric fields."""

import logging
import unittest
from unittest.mock import patch

import pandas as pd

from src.parsers import boi_parser, canara_parser, dbs_parser, esfb_parser
from src.parsers import hdfc_parser, icici_parser, idbi_parser, indus_parser
from src.parsers import iob_parser, kotak_parser, kvb_parser, southind_parser, unionbank_parser
from src.transform.analysis import _build_return_reject_sheet
from src.transform.validate import validate_records


class MissingValueReturnTests(unittest.TestCase):
    def test_shared_export_and_validation_accept_missing_zero_and_negative_values(self):
        rows = [
            {"Date": "01/01/2026", "Details": "RETURN", "Cheque No": number,
             "Debit": amount, "Credit": None, "Balance": None, "Sno": index}
            for index, (number, amount) in enumerate(
                [(None, None), ("", 0), (0, -25), (-12, None)], 1
            )
        ]
        validate_records(rows)
        statement = pd.DataFrame(rows)
        returns = _build_return_reject_sheet(statement)
        self.assertEqual(returns["Sno"].tolist(), [1, 2, 3, 4])
        self.assertEqual(statement["Sno"].tolist(), [1, 2, 3, 4])

    def test_table_parser_gates_keep_notices_without_amounts_or_balance(self):
        union = unionbank_parser._build_record(
            statement_date="01-01-2026", details="RETURN", cheque_number="",
            debit=None, credit=None, balance=None,
        )
        self.assertIsNotNone(union)
        self.assertIsNone(union["Debit"])

        esfb = esfb_parser._parse_transaction_row(
            ["01-Jan-2026", "", "CHQ RETURN", "", "", ""], 1,
        )
        self.assertIsNotNone(esfb)
        self.assertIsNone(esfb["Balance"])
        esfb_negative = esfb_parser._parse_transaction_row(
            ["01-Jan-2026", "", "CHQ RETURN", "-25.00", "", "-25.00"], 1,
        )
        self.assertEqual(esfb_negative["Debit"], -25.0)

        south = southind_parser._PendingRecord("01-01-26", detail_lines=["CHQ RETURN"])
        self.assertIsNotNone(south.finalize())

        boi_lines = [
            boi_parser.PositionedLine(50, 100, "01-01-2026"),
            boi_parser.PositionedLine(110, 100, "CHQ RETURN"),
        ]
        boi, _ = boi_parser._build_record(
            boi_lines, source_page=1, logger=logging.getLogger(__name__),
        )
        self.assertIsNotNone(boi)
        self.assertIsNone(boi["Balance"])
        zero_amount_lines = boi_lines + [
            boi_parser.PositionedLine(350, 100, "0.00"),
            boi_parser.PositionedLine(430, 100, "0.00"),
        ]
        boi_zero, _ = boi_parser._build_record(
            zero_amount_lines, source_page=1, logger=logging.getLogger(__name__),
        )
        self.assertIsNotNone(boi_zero)

    def test_positioned_and_text_layouts_keep_notices_without_amounts(self):
        words = [
            {"text": "01-Jan-2026", "x0": 39.0, "top": 100.0},
            {"text": "01-Jan-2026", "x0": 102.0, "top": 100.0},
            {"text": "CHQ", "x0": 155.0, "top": 100.0},
            {"text": "RETURN", "x0": 190.0, "top": 100.0},
        ]
        dbs = dbs_parser._positioned_transaction_line(words)
        self.assertIsNotNone(dbs)
        self.assertEqual(dbs[2], "CHQ RETURN")
        self.assertEqual(dbs[3:], (None, None, None))

        class IndusPage:
            def extract_text(self):
                return "01 Jan 2026 CHQ RETURN\nReason: insufficient funds\n02 Jan 2026 SALARY - 10.00 110.00"

        class IndusPdf:
            pages = [IndusPage()]

            def __enter__(self):
                return self

            def __exit__(self, *_):
                return False

        with patch.object(indus_parser.pdfplumber, "open", return_value=IndusPdf()):
            indus_rows = indus_parser.parse("dummy.pdf", logging.getLogger(__name__))
        self.assertEqual(len(indus_rows), 2)
        self.assertIn("insufficient funds", indus_rows[0]["Details"])
        self.assertIsNone(indus_rows[0]["Debit"])

        idbi = idbi_parser._parse_transaction_line(
            "1 01/01/2026 10:00:00 AM 01/01/2026 CHQ RETURN"
        )
        self.assertIsNotNone(idbi)
        self.assertIsNone(idbi["Debit"])

        iob = iob_parser._build_text_layout_record("01-01-2026 CHQ RETURN", 1, None)
        self.assertIsNotNone(iob)
        self.assertIsNone(iob["Credit"])

        canara, _ = canara_parser._finalize_record(
            canara_parser.PendingRecord("01-01-2026", ["CHQ RETURN"]), None,
        )
        self.assertIsNotNone(canara)
        self.assertIsNone(canara["Debit"])

    def test_kotak_table_keeps_return_without_amounts(self):
        class Page:
            def extract_tables(self):
                return [[["1", "01 Jan 2026", "CHQ RETURN", "", "", "", ""]]]

        class Pdf:
            pages = [Page()]

            def __enter__(self):
                return self

            def __exit__(self, *_):
                return False

        with patch.object(kotak_parser.pdfplumber, "open", return_value=Pdf()):
            rows = kotak_parser.parse_kotak_records("dummy.pdf", logging.getLogger(__name__))
        self.assertEqual(len(rows), 1)
        self.assertIsNone(rows[0]["Debit"])

    def test_hdfc_icici_and_kvb_keep_notices_without_amounts(self):
        hdfc, next_balance = hdfc_parser._finalize_record(
            hdfc_parser.PendingRecord("01/01/26", "CHQ RETURN", "", None, 0), 100.0,
        )
        self.assertIsNone(hdfc["Debit"])
        self.assertEqual(next_balance, 100.0)

        icici = icici_parser._transaction_list_record_from_row(
            ["1", "", "01/01/2026", "01/01/2026", "", "CHQ RETURN", "", "", ""],
            logging.getLogger(__name__),
        )
        self.assertIsNotNone(icici)
        self.assertIsNone(icici[1]["Debit"])

        kvb = kvb_parser._parse_tokenized_text_row(
            ["02/04/2026", "02/04/2026", "CHQ RETURN"]
        )
        self.assertIsNotNone(kvb)
        self.assertEqual(kvb.body_text, "CHQ RETURN")

    def test_hdfc_does_not_repeat_previous_row_when_a_later_row_is_unreadable(self):
        class TextPage:
            def get_text(self):
                return "selectable text"

        class Document:
            page_count = 1

            def __iter__(self):
                return iter([TextPage()])

            def __enter__(self):
                return self

            def __exit__(self, *_):
                return False

        class TablePage:
            def extract_text(self):
                return "\n".join([
                    hdfc_parser.TABLE_HEADER_TEXT,
                    "01/01/26 CHQ PAID 01/01/26 10.00 90.00",
                    "02/01/26 UNRELATED 02/01/26 - 90.00",
                ])

        class Pdf:
            pages = [TablePage()]

            def __enter__(self):
                return self

            def __exit__(self, *_):
                return False

        with patch.object(hdfc_parser.fitz, "open", return_value=Document()), patch.object(
            hdfc_parser.pdfplumber, "open", return_value=Pdf()
        ):
            records = hdfc_parser.parse("dummy.pdf", logging.getLogger(__name__))
        self.assertEqual(len(records), 1)


if __name__ == "__main__":
    unittest.main()
