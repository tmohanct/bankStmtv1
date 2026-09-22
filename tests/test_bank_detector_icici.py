from __future__ import annotations

import unittest


from src.parsers.detector import _detect_from_text


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

    def test_transaction_list_layout_outweighs_counterparty_bank(self) -> None:
        statement_text = """
        DETAILED STATEMENT
        Transactions List - -VRS AGENCIES (INR) - 617705500043
        Transaction ID Value Date Txn Posted Date ChequeNo. Description Cr/Dr
        Transaction Amount(INR) Available Balance(INR)
        UPI/customer/INDIAN BANK/reference
        """

        self.assertEqual(_detect_from_text(statement_text), "icici")


if __name__ == "__main__":
    unittest.main()
