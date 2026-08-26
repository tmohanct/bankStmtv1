from __future__ import annotations

import logging
import sys
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import dbs_parser
import utils


class _FakePage:
    def __init__(self, words: list[dict[str, object]]) -> None:
        self._words = words

    def extract_tables(self) -> list[list[list[str]]]:
        return []

    def extract_words(self, **kwargs) -> list[dict[str, object]]:
        return self._words


class _FakePdf:
    def __init__(self, pages: list[_FakePage]) -> None:
        self.pages = pages

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb) -> bool:
        return False


def _word(text: str, x0: float, top: float) -> dict[str, object]:
    return {"text": text, "x0": x0, "top": top}


class DBSPositionedParserTests(unittest.TestCase):
    def setUp(self) -> None:
        self.logger = logging.getLogger("tests.dbs.positioned")
        self.logger.handlers.clear()
        self.logger.addHandler(logging.NullHandler())

    def test_parses_positioned_text_and_merges_page_split_transaction(self) -> None:
        first_page_words = [
            _word("01-Jan-2026", 39.0, 680.0),
            _word("01-Jan-2026", 102.0, 680.0),
            _word("TRANSFER", 155.0, 680.0),
            _word("2,000.00", 337.0, 680.0),
            _word("10,000.00", 525.0, 680.0),
            _word("TRANSFER", 155.0, 687.0),
            _word("0107RF0000001", 189.0, 687.0),
        ]
        second_page_words = [
            _word("01-Jan-2026", 39.0, 79.0),
            _word("01-Jan-2026", 102.0, 79.0),
            _word("INR", 155.0, 79.0),
            _word("2000", 167.0, 79.0),
            _word("2,000.00", 337.0, 79.0),
            _word("10,000.00", 525.0, 79.0),
            _word("02-Jan-2026", 39.0, 115.0),
            _word("02-Jan-2026", 102.0, 115.0),
            _word("TRANSFER", 155.0, 115.0),
            _word("5,000.00", 425.0, 115.0),
            _word("15,000.00", 525.0, 115.0),
            _word("IMPS-REFERENCE", 155.0, 122.0),
        ]
        fake_pdf = _FakePdf(
            [_FakePage(first_page_words), _FakePage(second_page_words)]
        )

        with patch.object(utils.pdfplumber, "open", return_value=fake_pdf):
            records = dbs_parser.parse("dummy.pdf", self.logger)

        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["Debit"], 2000.0)
        self.assertIsNone(records[0]["Credit"])
        self.assertEqual(
            records[0]["Details"],
            "TRANSFER TRANSFER 0107RF0000001 INR 2000",
        )
        self.assertEqual(records[1]["Credit"], 5000.0)
        self.assertEqual(records[1]["Details"], "TRANSFER IMPS-REFERENCE")


if __name__ == "__main__":
    unittest.main()
