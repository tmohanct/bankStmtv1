from __future__ import annotations

import logging
import unittest

from src.parsers.icici_parser import (
    SavingsTextLine,
    SavingsTextToken,
    _parse_savings_page_lines,
)


def _line(top: float, *tokens: tuple[float, str]) -> SavingsTextLine:
    return SavingsTextLine(
        top=top,
        tokens=[SavingsTextToken(left=left, text=text) for left, text in tokens],
    )


class IciciSavingsLayoutTests(unittest.TestCase):
    def test_parses_wrapped_debit_and_credit_rows_after_opening_balance(self) -> None:
        lines = [
            _line(
                100,
                (27, "DATE"),
                (75, "MODE**"),
                (158, "PARTICULARS"),
                (387, "DEPOSITS"),
                (444, "WITHDRAWALS"),
                (545, "BALANCE"),
            ),
            _line(120, (27, "01-08-2026"), (158, "B/F"), (541, "4,56,241.26")),
            _line(130, (158, "UPI/customer")),
            _line(
                140,
                (27, "01-08-2026"),
                (158, "REF/one"),
                (474, "5,000.00"),
                (541, "4,51,241.26"),
            ),
            _line(150, (158, "okicici")),
            _line(160, (158, "NEFT/PAYER")),
            _line(
                170,
                (27, "01-08-2026"),
                (75, "MOBILE"),
                (107, "BANKING"),
                (158, "ICIREF"),
                (396, "10,000.00"),
                (541, "4,61,241.26"),
            ),
        ]

        records, final_balance, layout_seen = _parse_savings_page_lines(
            lines,
            previous_balance=None,
            logger=logging.getLogger("tests.icici_savings_layout"),
        )

        self.assertTrue(layout_seen)
        self.assertEqual(len(records), 2)
        self.assertEqual(records[0]["Date"], "01/08/2026")
        self.assertEqual(records[0]["Details"], "UPI/customer REF/one okicici")
        self.assertEqual(records[0]["Debit"], 5000.0)
        self.assertIsNone(records[0]["Credit"])
        self.assertEqual(records[0]["Balance"], 451241.26)
        self.assertEqual(records[1]["Details"], "MOBILE BANKING NEFT/PAYER ICIREF")
        self.assertIsNone(records[1]["Debit"])
        self.assertEqual(records[1]["Credit"], 10000.0)
        self.assertEqual(final_balance, 461241.26)

    def test_ignores_other_icici_layouts(self) -> None:
        records, final_balance, layout_seen = _parse_savings_page_lines(
            [_line(100, (27, "DATE"), (158, "DESCRIPTION"), (545, "BALANCE"))],
            previous_balance=123.45,
            logger=logging.getLogger("tests.icici_savings_layout"),
        )

        self.assertFalse(layout_seen)
        self.assertEqual(records, [])
        self.assertEqual(final_balance, 123.45)


if __name__ == "__main__":
    unittest.main()
