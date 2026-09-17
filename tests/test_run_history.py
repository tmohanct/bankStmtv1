from __future__ import annotations

from contextlib import closing, redirect_stderr, redirect_stdout
from concurrent.futures import ThreadPoolExecutor
from decimal import Decimal
import io
import json
import os
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import Mock, patch

import fitz

from src import main
from src.storage.run_history import RunHistory, database_path, file_sha256
from src.transform.validate import ReconciliationResult
from src.utils.ocr import record_ocr_use


def record(**changes):
    row = dict(Sno=1, Date="01/01/2026", Details="Payment", Detail_Clean="Payment",
               Debit=100.49, Credit=None, Balance=899.51)
    row.update(changes)
    return row


class RunHistoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.db = self.root / "data" / "history.db"
        (self.root / "input").mkdir()
        self.pdf = self.make_pdf("sample.pdf")
        env = patch.dict(os.environ, {"BANKSTMT_DB_PATH": str(self.db)})
        env.start()
        self.addCleanup(env.stop)

    def make_pdf(self, name):
        path = self.root / "input" / name
        with fitz.open() as document:
            page = document.new_page()
            page.insert_text((30, 30), "Axis Bank\nAccount Number: 1234567890")
            document.save(path)
        return path

    def rows(self):
        with closing(sqlite3.connect(self.db)) as connection:
            connection.row_factory = sqlite3.Row
            return [dict(row) for row in connection.execute("SELECT * FROM run_history ORDER BY run_id")]

    def invoke(self, *, argv=None, parser=None, reconciliation=None, export_error=None):
        parser = parser or Mock(return_value=[record()])
        argv = argv if argv is not None else ["--bank", "axis", "--pdf", str(self.pdf), "--currency", "INR"]
        final_path = self.root / "output" / "sample.xlsx"
        with patch.object(main, "PROJECT_ROOT", self.root), \
                patch.dict(main.PARSERS, {"axis": parser}), \
                patch.object(main, "reconcile", return_value=reconciliation or ReconciliationResult("passed")), \
                patch.object(main, "write_output_excel") as intermediate, \
                patch.object(main, "build_final_workbook", return_value=final_path, side_effect=export_error) as final, \
                redirect_stdout(io.StringIO()), redirect_stderr(io.StringIO()):
            code = main.main(argv)
        return code, intermediate, final

    def test_reruns_create_one_new_row_each_with_exact_money(self):
        for _ in range(2):
            self.assertEqual(self.invoke()[0], 0)
        rows = self.rows()
        self.assertEqual(len(rows), 2)
        self.assertNotEqual(rows[0]["run_id"], rows[1]["run_id"])
        self.assertNotEqual(rows[0]["log_path"], rows[1]["log_path"])
        for row in rows:
            self.assertEqual(row["status"], "success")
            self.assertEqual((row["input_count"], row["processed_count"], row["failed_count"]), (1, 1, 0))
            self.assertEqual((row["transaction_count"], row["duplicates_removed"]), (1, 0))
            self.assertEqual((row["total_debit_minor"], row["total_credit_minor"]), (10049, 0))
            self.assertEqual(row["reconciliation_status"], "passed")
            self.assertGreaterEqual(row["duration_ms"], 0)
            self.assertIsNotNone(row["finished_at"])
            self.assertTrue(row["app_version"].startswith("sha256:"))
            detail = json.loads(row["input_details_json"])[0]
            self.assertEqual(detail["sha256"], file_sha256(self.pdf))
            self.assertEqual(detail["page_count"], 1)
            self.assertEqual(detail["bank"], "axis")
            self.assertEqual(detail["first_transaction_date"], "2026-01-01")
        with closing(sqlite3.connect(self.db)) as connection:
            tables = connection.execute("SELECT name FROM sqlite_master WHERE type='table'").fetchall()
        self.assertEqual(tables, [("run_history",)])

    def test_running_row_is_committed_before_parser_and_counts_ocr_warnings(self):
        def parse(*args, **kwargs):
            row = self.rows()[0]
            self.assertEqual(row["status"], "running")
            self.assertIsNone(row["finished_at"])
            self.assertEqual(json.loads(row["input_details_json"])[0]["status"], "running")
            record_ocr_use()
            record_ocr_use()
            args[1].warning("Test warning")
            return [record()]
        self.assertEqual(self.invoke(parser=parse)[0], 0)
        row = self.rows()[0]
        self.assertEqual(row["ocr_file_count"], 1)
        self.assertEqual(row["warning_count"], 1)
        self.assertTrue(json.loads(row["input_details_json"])[0]["ocr_used"])

    def test_multiple_pdfs_one_row_and_merged_deduplication_totals(self):
        second = self.make_pdf("second.pdf")
        transactions = [record(), record(Sno=2, Date="02/01/2026", Balance=799.02)]
        args = ["--bank", "axis", "--pdf", f"{self.pdf};{second}", "--currency", "INR"]
        with patch.object(main, "account_identity", return_value="axis:1234567890"):
            code = self.invoke(argv=args, parser=Mock(side_effect=[transactions, [dict(r) for r in transactions]]))[0]
        self.assertEqual(code, 0)
        row, = self.rows()
        self.assertEqual((row["input_count"], row["processed_count"]), (2, 2))
        self.assertEqual((row["transaction_count"], row["duplicates_removed"]), (2, 2))
        self.assertEqual(row["total_debit_minor"], 20098)
        self.assertEqual(len(json.loads(row["input_details_json"])), 2)

    def test_mid_batch_failure_preserves_success_and_skips_remaining(self):
        second, third = self.make_pdf("second.pdf"), self.make_pdf("third.pdf")
        args = ["--bank", "axis", "--pdf", f"{self.pdf};{second};{third}"]
        code, intermediate, final = self.invoke(argv=args, parser=Mock(side_effect=[[record()], ValueError("Bad layout")]))
        self.assertEqual(code, 1)
        intermediate.assert_not_called()
        final.assert_not_called()
        row, = self.rows()
        self.assertEqual((row["processed_count"], row["failed_count"], row["skipped_count"]), (1, 1, 1))
        self.assertEqual(row["error_stage"], "parsing")
        self.assertEqual(row["error_message"], "Bad layout")
        self.assertEqual(row["reconciliation_status"], "partial")
        self.assertIsNone(row["transaction_count"])
        self.assertEqual([item["status"] for item in json.loads(row["input_details_json"])],
                         ["success", "failed", "skipped"])

    def test_reconciliation_failure_and_unavailable_are_distinct(self):
        self.assertEqual(self.invoke(reconciliation=ReconciliationResult("failed", ("total mismatch",)))[0], 1)
        self.assertEqual(self.invoke(reconciliation=ReconciliationResult("unavailable"))[0], 0)
        failed, unavailable = self.rows()
        self.assertEqual(failed["reconciliation_status"], "failed")
        self.assertEqual(failed["error_stage"], "reconciliation")
        self.assertEqual(unavailable["reconciliation_status"], "unavailable")
        self.assertEqual(unavailable["status"], "success")

    def test_export_failure_has_parsed_counts_but_no_final_output(self):
        self.assertEqual(self.invoke(export_error=PermissionError("Workbook open in Excel"))[0], 1)
        row, = self.rows()
        self.assertEqual(row["status"], "failed")
        self.assertEqual(row["processed_count"], 1)
        self.assertEqual(row["failed_count"], 0)
        self.assertEqual(row["transaction_count"], 1)
        self.assertEqual(row["error_stage"], "final_export")
        self.assertIsNone(row["output_path"])

    def test_early_failures_are_recorded(self):
        for args in (["--pdf", "absent.pdf"], ["--pdf", str(self.pdf), "--bank", "unknown"], []):
            self.assertEqual(self.invoke(argv=args)[0], 2)
        rows = self.rows()
        self.assertEqual(len(rows), 3)
        self.assertTrue(all(row["status"] == "failed" and row["finished_at"] for row in rows))
        self.assertEqual(rows[0]["failed_count"], 1)
        self.assertEqual(rows[1]["skipped_count"], 1)
        self.assertEqual(rows[2]["error_message"], "Invalid command-line arguments")

    def test_interrupt_is_logged_and_next_invocation_gets_new_row(self):
        self.assertEqual(self.invoke(parser=Mock(side_effect=KeyboardInterrupt()))[0], 130)
        self.assertEqual(self.invoke()[0], 0)
        first, second = self.rows()
        self.assertEqual(first["status"], "interrupted")
        self.assertEqual(json.loads(first["input_details_json"])[0]["status"], "interrupted")
        self.assertEqual(second["status"], "success")

    def test_passwords_redacted_in_database_even_when_filename_auto_resolved(self):
        secret_pdf = self.make_pdf("private$file-secret.pdf")
        args = ["--pdf", "private", "--bank", "axis", "--pwd", "explicit-secret"]
        message = f"Could not parse {secret_pdf}: file-secret explicit-secret; customer's PDF"
        self.assertEqual(self.invoke(argv=args, parser=Mock(side_effect=ValueError(message)))[0], 1)
        with closing(sqlite3.connect(self.db)) as connection:
            dump = "\n".join(connection.iterdump())
        self.assertNotIn("file-secret", dump)
        self.assertNotIn("explicit-secret", dump)
        self.assertIn("[redacted]", dump)
        detail = json.loads(self.rows()[0]["input_details_json"])[0]
        self.assertEqual(detail["filename"], "private.pdf")

    def test_currency_is_not_assumed(self):
        self.assertEqual(self.invoke(argv=["--pdf", str(self.pdf), "--bank", "axis"])[0], 0)
        row, = self.rows()
        self.assertIsNone(row["currency"])
        self.assertIsNone(row["total_debit_minor"])
        self.assertIsNone(row["total_credit_minor"])
        self.assertEqual(Decimal(json.loads(row["input_details_json"])[0]["total_debit"]), Decimal("100.49"))

    def test_help_does_not_create_database(self):
        with redirect_stdout(io.StringIO()), self.assertRaises(SystemExit) as raised:
            main.main(["--help"])
        self.assertEqual(raised.exception.code, 0)
        self.assertFalse(self.db.exists())

    def test_database_unavailable_stops_before_processing(self):
        with patch.dict(os.environ, {"BANKSTMT_DB_PATH": str(self.root)}):
            code, intermediate, final = self.invoke()
        self.assertEqual(code, 1)
        intermediate.assert_not_called()
        final.assert_not_called()

    def test_concurrent_history_writers_do_not_overwrite_or_mark_other_runs_interrupted(self):
        def execute(_):
            history = RunHistory(self.db)
            history.configure(["same.pdf"], currency="INR")
            history.finish("success")
            return history.run_id
        with ThreadPoolExecutor(max_workers=4) as pool:
            ids = list(pool.map(execute, range(8)))
        self.assertEqual(len(set(ids)), 8)
        self.assertEqual(len(self.rows()), 8)
        self.assertTrue(all(row["status"] == "success" for row in self.rows()))

    def test_short_password_does_not_corrupt_structured_values(self):
        history = RunHistory(self.db)
        history.configure(["sample.pdf"], password="s", currency="USD")
        history.finish("success")
        row, = self.rows()
        self.assertEqual(row["status"], "success")
        self.assertEqual(row["currency"], "USD")
        self.assertEqual(json.loads(row["input_details_json"])[0]["status"], "skipped")

    def test_database_path_can_be_overridden(self):
        self.assertEqual(database_path(), self.db)


    def test_masked_account_prefix_is_not_reported_as_last_four(self):
        history = RunHistory(self.db)
        history.configure([str(self.pdf)])
        history.start_input(0)
        for account, expected in (("1234XXXX", None), ("XXXX 7890", "7890"), ("1234567890", "7890")):
            with patch("src.utils.pdf_status_reader.read_first_page", return_value=Mock(values={"Account Number": account})):
                history.inspect_header(self.pdf)
            self.assertEqual(history.active["account_last4"], expected)

    def test_error_redaction_handles_escaped_passwords(self):
        secret = "line-one\nline-two"
        history = RunHistory(self.db)
        history.configure(["sample.pdf"], secret)
        history.error(f"Invalid password {secret!r}")
        history.finish("failed")
        message = self.rows()[0]["error_message"]
        self.assertNotIn("line-one", message)
        self.assertNotIn("line-two", message)

    def test_wrong_pdf_password_is_logged_before_parser_runs(self):
        encrypted = self.root / "input" / "locked.pdf"
        with fitz.open(self.pdf) as document:
            document.save(encrypted, encryption=fitz.PDF_ENCRYPT_AES_256,
                          owner_pw="owner-secret", user_pw="correct-password")
        parser = Mock(return_value=[record()])
        args = ["--bank", "axis", "--pdf", str(encrypted), "--pwd", "incorrect-password"]
        self.assertEqual(self.invoke(argv=args, parser=parser)[0], 1)
        parser.assert_not_called()
        row, = self.rows()
        self.assertEqual(row["error_stage"], "decryption")
        detail, = json.loads(row["input_details_json"])
        self.assertEqual(detail["sha256"], file_sha256(encrypted))
        self.assertTrue(detail["is_encrypted"])
        self.assertEqual(detail["bank"], "axis")
        self.assertNotIn("incorrect-password", row["error_message"])


if __name__ == "__main__":
    unittest.main()
