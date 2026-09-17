"""One durable SQLite row per invocation; individual inputs are JSON objects."""
from __future__ import annotations

from contextlib import closing
from datetime import datetime, timezone
from decimal import Decimal, ROUND_HALF_UP
import hashlib
import json
import logging
import os
from pathlib import Path
import re
import sqlite3
import time

from src.utils.statement_utils import normalize_date, split_pdf_filename_metadata


SCHEMA = """
CREATE TABLE IF NOT EXISTS run_history (
    run_id INTEGER PRIMARY KEY,
    started_at TEXT NOT NULL,
    finished_at TEXT,
    duration_ms INTEGER,
    status TEXT NOT NULL CHECK(status IN ('running','success','failed','interrupted')),
    app_version TEXT,
    rules_sha256 TEXT,
    input_count INTEGER NOT NULL DEFAULT 0,
    processed_count INTEGER NOT NULL DEFAULT 0,
    failed_count INTEGER NOT NULL DEFAULT 0,
    skipped_count INTEGER NOT NULL DEFAULT 0,
    transaction_count INTEGER,
    duplicates_removed INTEGER,
    total_debit_minor INTEGER,
    total_credit_minor INTEGER,
    currency TEXT,
    ocr_file_count INTEGER NOT NULL DEFAULT 0,
    reconciliation_status TEXT NOT NULL DEFAULT 'not_run',
    warning_count INTEGER NOT NULL DEFAULT 0,
    input_details_json TEXT NOT NULL DEFAULT '[]',
    output_path TEXT,
    log_path TEXT,
    error_stage TEXT,
    error_message TEXT
)
"""


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds")


def database_path() -> Path:
    override = os.environ.get("BANKSTMT_DB_PATH")
    if override:
        return Path(override).expanduser().resolve()
    if os.name == "nt":
        base = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
    else:
        base = Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share")
    return base / "bankStmtv1" / "bankstmt.db"


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def source_version(root: Path) -> str:
    """Also identifies uncommitted code and installations without Git."""
    digest = hashlib.sha256()
    for path in sorted((root / "src").rglob("*.py")):
        digest.update(path.relative_to(root).as_posix().encode())
        digest.update(bytes.fromhex(file_sha256(path)))
    return "sha256:" + digest.hexdigest()


def decimal_amount(value) -> Decimal:
    if value is None or str(value).strip().lower() in ("", "nan", "none"):
        return Decimal(0)
    amount = Decimal(str(value))
    if not amount.is_finite():
        raise ValueError("Non-finite monetary value in run statistics")
    return amount


def transaction_metrics(records) -> dict:
    dates = [normalize_date(row.get("Date")) for row in records]
    dates = sorted(datetime.strptime(value, "%d/%m/%Y").date().isoformat() for value in dates if value)
    return {
        "transaction_count": len(records),
        "total_debit": str(sum((decimal_amount(row.get("Debit")) for row in records), Decimal(0))),
        "total_credit": str(sum((decimal_amount(row.get("Credit")) for row in records), Decimal(0))),
        "first_transaction_date": dates[0] if dates else None,
        "last_transaction_date": dates[-1] if dates else None,
        "negative_balance_count": sum(
            row.get("Balance") not in (None, "") and decimal_amount(row["Balance"]) < 0
            for row in records
        ),
    }


class WarningCounter(logging.Handler):
    def __init__(self):
        super().__init__(logging.WARNING)
        self.count = 0

    def emit(self, record):
        if record.levelno == logging.WARNING:
            self.count += 1


class RunHistory:
    def __init__(self, path: Path):
        self.path = path
        self.started = time.monotonic()
        self.stage = "input"
        self.secrets: set[str] = set()
        self.inputs: list[dict] = []
        self.active: dict | None = None
        self.warning_counter = WarningCounter()
        self.values = {"started_at": utc_now(), "status": "running"}
        path.parent.mkdir(parents=True, exist_ok=True)
        with closing(sqlite3.connect(path, timeout=10)) as connection, connection:
            connection.execute(SCHEMA)
            connection.execute("CREATE INDEX IF NOT EXISTS run_history_started_at ON run_history(started_at)")
            connection.execute("CREATE INDEX IF NOT EXISTS run_history_status ON run_history(status)")
            cursor = connection.execute(
                "INSERT INTO run_history(started_at, status) VALUES (?, 'running')",
                (self.values["started_at"],),
            )
            self.run_id = cursor.lastrowid

    def add_secret(self, value):
        if value:
            self.secrets.add(str(value))

    def redact(self, value):
        if isinstance(value, dict):
            return {key: self.redact(item) for key, item in value.items()}
        if isinstance(value, list):
            return [self.redact(item) for item in value]
        if isinstance(value, str):
            variants = set(self.secrets)
            for secret in self.secrets:
                variants.add(repr(secret)[1:-1])
                variants.add(json.dumps(secret, ensure_ascii=False)[1:-1])
            for secret in sorted(variants, key=len, reverse=True):
                value = value.replace(secret, "[redacted]")
        return value

    def configure(self, filenames, password=None, currency=None):
        self.add_secret(password)
        for filename in filenames:
            name = Path(filename).name
            # CLI inputs can omit .pdf; normalize before splitting the password.
            normalized = name if name.lower().endswith(".pdf") else name + ".pdf"
            stem, secret = split_pdf_filename_metadata(normalized)
            self.add_secret(secret)
            self.inputs.append({"filename": stem + ".pdf", "status": "queued", "ocr_used": False,
                                "reconciliation_status": "not_run", "currency": currency})
        self.values["currency"] = currency
        self.save()

    def save(self, **changes):
        self.values.update(changes)
        checks = [item.get("reconciliation_status", "not_run") for item in self.inputs]
        if "failed" in checks:
            reconciliation = "failed"
        elif checks and all(value == "passed" for value in checks):
            reconciliation = "passed"
        elif checks and all(value == "unavailable" for value in checks):
            reconciliation = "unavailable"
        elif any(value != "not_run" for value in checks):
            reconciliation = "partial"
        else:
            reconciliation = "not_run"
        self.values.update(
            input_count=len(self.inputs),
            processed_count=sum(item["status"] == "success" for item in self.inputs),
            failed_count=sum(item["status"] == "failed" for item in self.inputs),
            skipped_count=sum(item["status"] == "skipped" for item in self.inputs),
            ocr_file_count=sum(bool(item.get("ocr_used")) for item in self.inputs),
            reconciliation_status=reconciliation,
            warning_count=self.warning_counter.count,
        )
        # Redact free text only: short passwords must not corrupt timestamps,
        # status values, currency codes, or SHA-256 identifiers.
        data = dict(self.values)
        for key in ("error_message", "output_path", "log_path"):
            if key in data:
                data[key] = self.redact(data[key])
        safe_inputs = []
        for item in self.inputs:
            safe_item = dict(item)
            for key in ("filename", "error_message", "metadata_error", "statement_period"):
                if key in safe_item:
                    safe_item[key] = self.redact(safe_item[key])
            safe_inputs.append(safe_item)
        data["input_details_json"] = json.dumps(safe_inputs, ensure_ascii=False)
        columns = ", ".join(f"{key} = ?" for key in data)
        with closing(sqlite3.connect(self.path, timeout=10)) as connection, connection:
            connection.execute(f"UPDATE run_history SET {columns} WHERE run_id = ?", (*data.values(), self.run_id))

    def start_input(self, index):
        self.active = self.inputs[index]
        self.active.update(status="running", started_at=utc_now())
        self.stage = "input"
        self.save()

    def resolved_input(self, index, path):
        stem, secret = split_pdf_filename_metadata(path)
        self.add_secret(secret)
        self.inputs[index]["filename"] = stem + ".pdf"

    def record_ocr(self):
        if self.active is not None:
            self.active["ocr_used"] = True

    def inspect_source(self, path):
        import fitz

        try:
            self.active.update(file_size_bytes=path.stat().st_size, sha256=file_sha256(path))
            with fitz.open(path) as document:
                self.active.update(is_encrypted=bool(document.needs_pass), page_count=len(document))
        except Exception as exc:
            self.active["metadata_error"] = str(exc)
        self.save()

    def inspect_header(self, readable_path):
        """Optional metadata must never prevent parsing a supported statement."""
        import fitz
        from src.utils.pdf_status_reader import read_first_page

        try:
            with fitz.open(readable_path) as document:
                self.active["page_count"] = len(document)
            header = read_first_page(readable_path, allow_ocr=False)
            account = re.sub(r"[\s-]", "", header.values.get("Account Number") or "")
            self.active["account_last4"] = account[-4:] if re.search(r"\d{4}$", account) else None
            self.active["statement_period"] = header.values.get("Statement Date Between") or None
        except Exception as exc:
            self.active["metadata_error"] = str(exc)

    def complete_input(self):
        self.active.update(status="success", finished_at=utc_now())
        self.save()
        self.active = None

    def merged_metrics(self, records, duplicates_removed):
        metrics = transaction_metrics(records)
        values = dict(transaction_count=metrics["transaction_count"], duplicates_removed=duplicates_removed)
        if self.values.get("currency"):
            for direction in ("debit", "credit"):
                amount = Decimal(metrics[f"total_{direction}"]) * 100
                values[f"total_{direction}_minor"] = int(amount.quantize(Decimal(1), rounding=ROUND_HALF_UP))
        self.save(**values)

    def error(self, message):
        self.values.update(error_stage=self.stage, error_message=str(message))

    def finish(self, status):
        now = utc_now()
        if self.active is not None and self.active["status"] == "running":
            self.active.update(status="interrupted" if status == "interrupted" else "failed",
                               finished_at=now, error_stage=self.values.get("error_stage"),
                               error_message=self.values.get("error_message"))
        for item in self.inputs:
            if item["status"] == "queued":
                item["status"] = "skipped"
        self.save(status=status, finished_at=now, duration_ms=round((time.monotonic() - self.started) * 1000))
