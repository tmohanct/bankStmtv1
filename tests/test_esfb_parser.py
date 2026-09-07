from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src"))
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import esfb_parser
import run
from bank_detector import _detect_from_text
from parsers import esfb_parser as esfb_parser_impl


class _FakePage:
    def __init__(self, tables: list[list[list[str]]]) -> None:
        self._tables = tables

    def extract_tables(self) -> list[list[list[str]]]:
        return self._tables


class _FakePdf:
    def __init__(self, pages: list[_FakePage]) -> None:
        self.pages = pages

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        return False


class ESFBParserTests(unittest.TestCase):
    def setUp(self) -> None:
        self.logger = logging.getLogger("tests.esfb")
        self.logger.handlers.clear()
        self.logger.addHandler(logging.NullHandler())

    def test_parses_esfb_table_and_repairs_spillover_text(self) -> None:
        table = [
            ["Date", "Reference No. / Cheque No.", "Narration", "Withdrawal", "Deposit", "ClosingBalance"],
            [
                "26-Apr-2026",
                "IMPSFIN202604262250439430561050",
                "620IMP P2A 611622639093 SANTHOSHKUMARTHIYA",
                "GARAJAN",
                "300000.00",
                "302307.00",
            ],
            [
                "03-Sep-2026",
                "NACH000018987827",
                "ACH DR:7953002059116",
                "37676.00 9,KISETSU03092026",
                "",
                "205247.79",
            ],
            [
                "10-Aug-2026",
                "000000000069",
                "CHQ PAID-IC 1400-CHIDAMBARANATHANK",
                "150000.00 -",
                "",
                "173643.79",
            ],
        ]

        with patch.object(
            esfb_parser_impl.pdfplumber,
            "open",
            return_value=_FakePdf([_FakePage([table])]),
        ):
            records = esfb_parser.parse("dummy.pdf", self.logger)

        self.assertEqual(len(records), 3)
        self.assertEqual(records[0]["Date"], "26/04/2026")
        self.assertEqual(records[0]["Credit"], 300000.0)
        self.assertIn("GARAJAN", records[0]["Details"])
        self.assertFalse(records[0]["Details"].startswith("620"))
        self.assertEqual(records[1]["Debit"], 37676.0)
        self.assertIn("9,KISETSU03092026", records[1]["Details"])
        self.assertEqual(records[2]["Debit"], 150000.0)
        self.assertEqual(records[2]["Cheque No"], "000000000069")

    def test_detects_esfb_without_being_confused_by_other_bank_narration(self) -> None:
        text = """
        Equitas Small Finance Bank Limited
        IFSC Code ESFB0001121
        RTGS DR-UTIB0005036-AXIS BANK
        NEFT CR-HDFC0000058-HDFC BANK
        """
        self.assertEqual(_detect_from_text(text), "esfb")

    def test_cli_accepts_esfb_and_equitas_aliases(self) -> None:
        self.assertIs(run.PARSERS["esfb"], esfb_parser.parse)
        self.assertEqual(run.BANK_ALIASES["equitas"], "esfb")


if __name__ == "__main__":
    unittest.main()
