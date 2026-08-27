from __future__ import annotations

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

from bank_detector import _detect_from_text


class IciciBankDetectorTests(unittest.TestCase):
    def test_icici_statement_footer_outweighs_banks_in_transaction_details(self) -> None:
        statement_text = """
        Statement of Transactions in Savings Account Number: 123456789
        UPI/merchant/AXIS BANK/one
        UPI/merchant/AXIS BANK/two
        UPI/merchant/AXIS BANK/three
        Visit www.icicibank.com
        """

        self.assertEqual(_detect_from_text(statement_text), "icici")


if __name__ == "__main__":
    unittest.main()
