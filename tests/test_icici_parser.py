from __future__ import annotations

import sys
import logging
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

from icici_parser import (
    TextLine,
    _extract_account_statement_block_seed,
    _extract_account_statement_blocks,
    _finalize_text_record,
)


class IciciWrappedDateRegressionTests(unittest.TestCase):
    def test_wrapped_date_and_balance_fragments_form_one_record(self) -> None:
        lines = [
            TextLine(
                top=372.65,
                tokens=[
                    (50.0, "1"),
                    (97.118, "S132531"),
                    (161.96, "01-Mar-"),
                    (296.3, "UPI/109287"),
                    (363.48, "59500.00"),
                    (497.83, "-"),
                ],
            ),
            TextLine(
                top=384.65,
                tokens=[
                    (161.96, "2026"),
                    (296.3, "034369/UPI"),
                    (497.83, "11795198.0"),
                ],
            ),
            TextLine(
                top=396.65,
                tokens=[
                    (296.3, "Pay/shrisaia"),
                    (497.83, "9"),
                ],
            ),
            TextLine(top=408.65, tokens=[(296.3, "grofood//ICI")]),
            TextLine(top=420.65, tokens=[(296.3, "3b48f")]),
        ]

        seed = _extract_account_statement_block_seed(lines)

        self.assertIsNotNone(seed)
        assert seed is not None
        self.assertEqual(seed.raw_date, "01-Mar-2026")
        self.assertEqual(seed.amount_text, "59500.00")
        self.assertEqual(seed.balance_text, "11795198.09")
        self.assertEqual(seed.inline_detail, "UPI/109287 034369/UPI Pay/shrisaia grofood//ICI 3b48f")

        record, next_balance = _finalize_text_record(
            seed,
            previous_balance=None,
            credit_threshold=400.0,
            logger=logging.getLogger("tests.icici_parser"),
        )

        self.assertEqual(record["Date"], "01/03/2026")
        self.assertEqual(record["Debit"], 59500.0)
        self.assertIsNone(record["Credit"])
        self.assertEqual(record["Balance"], 11795198.09)
        self.assertEqual(next_balance, 11795198.09)


    def test_negative_balance_is_retained(self) -> None:
        lines = [
            TextLine(
                top=136.0,
                tokens=[
                    (50.0, "167"),
                    (94.78, "S7465371"),
                    (161.96, "02-Jun-"),
                    (296.3, "601305500"),
                    (363.48, "84545.00"),
                    (497.83, "-78006.14"),
                ],
            ),
            TextLine(
                top=148.0,
                tokens=[
                    (161.96, "2026"),
                    (296.3, "425:Int.Coll:"),
                ],
            ),
        ]

        seed = _extract_account_statement_block_seed(lines)

        self.assertIsNotNone(seed)
        assert seed is not None
        self.assertEqual(seed.balance_text, "-78006.14")

    def test_legends_do_not_leak_into_last_transaction(self) -> None:
        lines = [
            TextLine(
                top=136.0,
                tokens=[
                    (50.0, "268"),
                    (94.78, "S1105185"),
                    (161.96, "16-Aug-"),
                    (296.3, "MMT/IMPS/6"),
                    (430.65, "10000.00"),
                    (497.83, "10785.10"),
                ],
            ),
            TextLine(
                top=148.0,
                tokens=[
                    (161.96, "2026"),
                    (296.3, "228197477"),
                ],
            ),
            TextLine(top=160.0, tokens=[(296.3, "20/Annapur")]),
            TextLine(
                top=216.48,
                tokens=[
                    (40.0, "Legends"),
                    (92.68, "Used"),
                    (125.356, "in"),
                    (139.36, "Account"),
                    (190.696, "Statement"),
                ],
            ),
            TextLine(top=491.07, tokens=[(299.02, "third"), (321.25, "party")]),
        ]

        blocks = _extract_account_statement_blocks(lines)
        seed = _extract_account_statement_block_seed(blocks[0])

        self.assertEqual(len(blocks), 1)
        self.assertIsNotNone(seed)
        assert seed is not None
        self.assertEqual(seed.inline_detail, "MMT/IMPS/6 228197477 20/Annapur")


if __name__ == "__main__":
    unittest.main()
