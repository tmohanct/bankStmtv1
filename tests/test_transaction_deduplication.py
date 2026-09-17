from __future__ import annotations
import unittest
import pandas as pd
from src.transform.normalize import remove_exact_duplicate_transactions


def transaction(serial, source="a.pdf", account="axis:123456789", **values):
    row = {"Sno": serial, "Date": "01/01/2026", "Details": "ATM", "Detail_Clean": "ATM",
           "Cheque No": "", "Debit": 100.0, "Credit": None, "Balance": 900.0,
           "Source": source, "_Source_Id": source, "_Account_Key": account}
    row.update(values)
    return row


class TransactionDeduplicationTests(unittest.TestCase):
    def dedup(self, rows):
        return remove_exact_duplicate_transactions(pd.DataFrame(rows))

    def test_preserves_debit_deposit_debit_cycle_with_identical_end_balances(self):
        rows = [transaction(1), transaction(2, Details="CASH", Detail_Clean="CASH", Debit=None, Credit=100., Balance=1000.), transaction(3)]
        self.assertEqual(self.dedup(rows)["Sno"].tolist(), [1, 2, 3])

    def test_preserves_repeated_rows_without_identity_metadata(self):
        frame = pd.DataFrame([transaction(1), transaction(2)]).drop(columns=["_Account_Key", "_Source_Id"])
        self.assertEqual(len(remove_exact_duplicate_transactions(frame)), 2)

    def test_removes_adjacent_overlap_of_same_account(self):
        rows = [transaction(1), transaction(2, Balance=800.), transaction(3, "b.pdf"), transaction(4, "b.pdf", Balance=800.), transaction(5, "b.pdf", Balance=700.)]
        self.assertEqual(self.dedup(rows)["Sno"].tolist(), [1, 2, 5])

    def test_keeps_lone_ambiguous_match_across_files(self):
        self.assertEqual(len(self.dedup([transaction(1), transaction(2, "b.pdf")])), 2)

    def test_keeps_other_accounts_and_unknown_identity(self):
        for account in ("axis:999999999", "sbi:123456789", None, ""):
            with self.subTest(account=account):
                rows = [transaction(1), transaction(2, Balance=800.), transaction(3, "b.pdf", account), transaction(4, "b.pdf", account, Balance=800.)]
                self.assertEqual(len(self.dedup(rows)), 4)

    def test_preserves_occurrence_count_across_overlapping_statements(self):
        rows = [transaction(1), transaction(2), transaction(3, "b.pdf"), transaction(4, "b.pdf"), transaction(5, "b.pdf")]
        self.assertEqual(len(self.dedup(rows)), 3)

    def test_requires_running_balances_for_overlap(self):
        rows = [transaction(1, Balance=None), transaction(2, Balance=None), transaction(3,"b.pdf",Balance=None), transaction(4,"b.pdf",Balance=None)]
        self.assertEqual(len(self.dedup(rows)), 4)

    def test_preserves_input_frame_and_index(self):
        frame = pd.DataFrame([transaction(1), transaction(2)], index=[10, 20])
        original = frame.copy(deep=True)
        result = remove_exact_duplicate_transactions(frame)
        pd.testing.assert_frame_equal(frame, original)
        self.assertEqual(result.index.tolist(), [10, 20])


if __name__ == "__main__":
    unittest.main()
