from __future__ import annotations

import unittest

import pandas as pd


from src.transform.analysis import (
    _build_return_reject_sheet,
)
from src.transform.cheque_returns import is_cheque_return
from src.utils.statement_utils import OUTPUT_COLUMNS


class ChequeReturnClassificationTests(unittest.TestCase):
    def test_return_abbreviations_and_word_order(self):
        descriptions = [
            "CHQ RETURN 123456", "Cheque return (Issued):500384:Exceeds Arrangement",
            "BRN-OW RTN CLG: REJECT:238951:Funds insufficient", "IW REJ INST 123456",
            "CHEQUE DISHONOURED 123456", "CHECK DISHONORED 123456", "chq-return/123456",
            "CHQ RTN 123456", "CHEQUE RTN 123456", "RTN CHQ 123456",
            "RETURN OF CHEQUE 123456", "CHQ 123456 RETURN", "INWARD CLEARING RETURN 123456",
            "OUTWARD CLG RTN 123456", "CTS RETURN 123456", "CHQ UNPAID 123456",
            "CHEQUE BOUNCED 123456", "CHQ RETD 123456", "RETD CHQ 123456",
            "I/W CHQ RETURN 123456", "O/W CHQ RET 123456", "CHQ REJ 123456",
            "CHQ RETN 123456", "CHQ RTND 123456", "CHQ RJCT 123456",
            "CHQRETURN123456", "IWCHQRET123456", "OWCHQRTN123456",
            "RETURNCHQ123456", "CTSRETURN123456", "IWREJINST123456",
            "CHQRETURNISSUED123456", "CHQ\nRTN:123456",
        ]
        for details in descriptions:
            with self.subTest(details=details):
                self.assertTrue(is_cheque_return(details))

    def test_return_reasons_with_cheque_context(self):
        reasons = [
            "FUNDS INSUFFICIENT", "INSUFFICIENT FUNDS", "INSUFFICIENT BALANCE",
            "FUNDS INSUFF", "EXCEEDS ARRANGEMENT", "PAYMENT STOPPED BY DRAWER",
            "STOP PAYMENT", "SIGNATURE DIFFERS", "SIGNATURE MISMATCH",
            "DRAWERS SIGNATURE REQUIRED", "ACCOUNT CLOSED", "ACCOUNT FROZEN",
            "ACCOUNT BLOCKED", "ALTERATION REQUIRES DRAWERS AUTHENTICATION",
            "AMOUNT IN WORDS AND FIGURES DIFFER", "STALE", "POST DATED",
        ]
        for reason in reasons:
            with self.subTest(reason=reason):
                self.assertTrue(is_cheque_return("CHQ 123456 " + reason))
                self.assertTrue(is_cheque_return(reason, "00123456"))
                self.assertFalse(is_cheque_return(reason))
        for details in ("STALE CHEQUE 123456", "POST DATED CHEQUE 123456", "CHEQUE STALE"):
            with self.subTest(details=details):
                self.assertTrue(is_cheque_return(details))

    def test_number_backed_returns(self):
        for details in ("RETURN 123456", "REJECT:123456:Reason 01", "PAYMENT RETURN", "UNPAID"):
            with self.subTest(details=details):
                self.assertTrue(is_cheque_return(details, 123456.0))
                self.assertFalse(is_cheque_return(details))
        for number in (None, "", "000000", "reference", "123ABC", float("nan"), pd.NA):
            with self.subTest(number=number):
                self.assertFalse(is_cheque_return("RETURN", number))

    def test_electronic_returns_are_excluded_even_with_cheque_numbers(self):
        descriptions = [
            "NEFT/RETURN/REF123", "RTGS RETURN REF123", "IMPS RETURN REF123",
            "UPI REJECTED REF123", "NACH DISHONOUR REF123", "ECS RETURN REF123",
            "ACH RTN CHRG REF123", "NEFTRETURN123456", "UPIREJECTED123456",
            "NEFT RETURN CHEQUE TRADERS", "NACH RETURN FUNDS INSUFFICIENT",
        ]
        for details in descriptions:
            with self.subTest(details=details):
                self.assertFalse(is_cheque_return(details, "123456"))

    def test_charges_and_taxes_are_excluded(self):
        descriptions = [
            "CHQRETURNCHGSINCLGST141125-CDT25326 37321333", "RETURN CHARGES",
            "CHQ RTN CHRG 123456", "CHQ RTN FEE 123456", "RTN CHQ CHGS 123456",
            "RTN CHG-123456/01", "CHQRTNCHRGS123456", "CHQRETURNFEE123456",
            "CHQ DISHONOUR COMMISSION", "CHARGES FOR CHEQUE RETURN",
            "CHGSCHQRETURN123456", "GST ON CHQ RETURN 123456", "CGSTCHQRETURN123456",
            "CHQ RETRIEVAL FEE 123456", "CHQRETRIEVAL123456",
        ]
        for details in descriptions:
            with self.subTest(details=details):
                self.assertFalse(is_cheque_return(details, "123456"))

    def test_unrelated_successful_negated_or_reversed_entries_are_excluded(self):
        descriptions = [
            "GOODS RETURNED", "RETURNED GOODS", "Paid to Rita", "CHQ DEPOSIT 123456",
            "CHQ PAID 123456", "CHEQUE ISSUED 123456", "CHQ RETRIEVAL 123456",
            "CHEQUE PAYMENT TO REJECTIONLESS TRADERS", "NO CHEQUE RETURN",
            "CHQ 123456 NOT RETURNED", "CHQ RETURN REQUEST", "CHQ RETURN PENDING",
            "CHQ STOP PAYMENT REQUEST", "CHQ RETURN REVERSAL", "REVERSAL OF CHQ RETURN",
            "CHQRETURNREVERSAL", "CHQRETURN123456REVERSAL", "REVERSALCHQRETURN",
            "NOCHQRETURN", "CHQRETURNREQUESTED", "CHQRETURNPENDING",
        ]
        for details in descriptions:
            with self.subTest(details=details):
                self.assertFalse(is_cheque_return(details))
        self.assertFalse(is_cheque_return("GOODS RETURNED", "123456"))

    def test_clean_narration_is_only_a_missing_details_fallback(self):
        for missing in (None, "", "  ", float("nan"), pd.NA):
            with self.subTest(missing=missing):
                self.assertTrue(is_cheque_return(missing, "", "CHQRETURN123456"))
                self.assertFalse(is_cheque_return(missing))
        for original in ("NEFT RETURN", "CHQ RETURN CHARGES", "Paid to Rita"):
            with self.subTest(original=original):
                self.assertFalse(is_cheque_return(original, "123456", "CHQRETURN123456"))


class ChequeReturnSheetTests(unittest.TestCase):
    def test_preserves_debits_credits_small_amounts_order_and_input(self):
        frame = pd.DataFrame([
            {"Sno": 4, "Details": "CHQ RTN", "Debit": 20, "Credit": 0},
            {"Sno": 3, "Details": "RTN CHQ", "Debit": 0, "Credit": 1000},
            {"Sno": 2, "Details": "CHQ RETURN CHGS", "Debit": 59, "Credit": 0},
            {"Sno": 1, "Details": "NEFT RETURN", "Debit": 0, "Credit": 1000},
        ], index=[7, 7, 2, 1])
        original = frame.copy(deep=True)
        result = _build_return_reject_sheet(frame)
        self.assertEqual(result["Sno"].tolist(), [4, 3])
        self.assertEqual(result["Debit"].tolist(), [20, 0])
        self.assertEqual(result["Credit"].tolist(), [0, 1000])
        self.assertEqual(result.index.tolist(), [7, 7])
        self.assertEqual(result.columns.tolist(), OUTPUT_COLUMNS)
        pd.testing.assert_frame_equal(frame, original)

    def test_number_and_clean_details_reach_classifier(self):
        frame = pd.DataFrame([
            {"Sno": 1, "Details": "FUNDS INSUFFICIENT", "Cheque No": "123456"},
            {"Sno": 2, "Details": None, "Detail_Clean": "CHQRETURN123456"},
            {"Sno": 3, "Details": "GOODS RETURNED", "Cheque No": "123456"},
        ])
        self.assertEqual(_build_return_reject_sheet(frame)["Sno"].tolist(), [1, 2])
        cleaned_only = pd.DataFrame([{"Detail_Clean": "RTNCHQ123456"}])
        self.assertEqual(len(_build_return_reject_sheet(cleaned_only)), 1)

    def test_empty_or_nonmatching_input_has_output_schema(self):
        for frame in (pd.DataFrame(), pd.DataFrame([{"Details": "Paid to Rita"}]),
                      pd.DataFrame([{"Cheque No": "123456"}])):
            with self.subTest(columns=frame.columns.tolist()):
                result = _build_return_reject_sheet(frame)
                self.assertTrue(result.empty)
                self.assertEqual(result.columns.tolist(), OUTPUT_COLUMNS)


if __name__ == "__main__":
    unittest.main()
