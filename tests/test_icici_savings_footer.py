from __future__ import annotations

import logging
import unittest

from src.parsers.icici_parser import SavingsTextLine, SavingsTextToken, _parse_savings_page_lines


def _line(top: float, *tokens: tuple[float, str]) -> SavingsTextLine:
    return SavingsTextLine(
        top=top,
        tokens=[SavingsTextToken(left=left, text=text) for left, text in tokens],
    )


class IciciSavingsFooterTests(unittest.TestCase):
    def test_account_summary_after_last_row_does_not_leak_into_details(self) -> None:
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
            _line(120, (158, "UPI/GEETHA")),
            _line(
                130,
                (27, "27-08-2026"),
                (158, "CANARA/BANK"),
                (474, "2,000.00"),
                (541, "2,33,736.09"),
            ),
            _line(140, (158, "customer")),
            _line(170, (27, "ACCOUNT"), (68, "TYPE"), (151, "ACCOUNT"), (193, "NUMBER")),
            _line(180, (151, "612501125416")),
            _line(190, (27, "Nominee"), (54, "name"), (151, "consent"), (181, "customer.")),
        ]

        records, _, layout_seen = _parse_savings_page_lines(
            lines,
            previous_balance=235736.09,
            logger=logging.getLogger("tests.icici_savings_footer"),
        )

        self.assertTrue(layout_seen)
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["Details"], "UPI/GEETHA CANARA/BANK customer")


if __name__ == "__main__":
    unittest.main()
