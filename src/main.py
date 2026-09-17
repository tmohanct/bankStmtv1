import argparse
import logging
import re
import shutil
import subprocess
import sys
import threading
import time
import os
from pathlib import Path

# Support both `python -m src.main` and direct script invocation.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from src.parsers import detector as bank_detector
from src.parsers.parser_registry import PARSER_REGISTRY as PARSERS
from src.export.final_excel_builder import build_final_workbook
from src.export.excel_writer import write_output_excel
from src.transform.normalize import records_to_dataframe, remove_exact_duplicate_transactions
from src.transform.validate import reconcile, validate_records
from src.utils.statement_utils import prepare_pdf_for_reading, resolve_pdf_path, split_pdf_filename_metadata
from src.utils.pdf_status_reader import read_first_page
from src.utils.ocr import observe_ocr
from src.storage.run_history import RunHistory, database_path, file_sha256, source_version, transaction_metrics

EXAMPLE_CMD = "python run.py --pdf My Statement.pdf --bank icici"
INVALID_FILENAME_CHARS_RE = re.compile(r'[<>:"/\\\\|?*]+')

SUPPORTED_BANK_CODES = "/".join(sorted(PARSERS))
BANK_ALIASES = {"equitas": "esfb"}


class CliParser(argparse.ArgumentParser):
    def error(self, message):
        self.print_usage(sys.stderr)
        self.exit(2, f"Error: {message}\nExample: {EXAMPLE_CMD}\n")


def parse_args(argv=None):
    parser = CliParser(description="Modular bank statement PDF to Excel parser")
    parser.add_argument(
        "--pdf",
        "--file",
        dest="pdf",
        required=True,
        nargs="+",
        help=(
            "PDF name from input/ (the .pdf extension is optional). Names containing spaces "
            "may be entered without quotes. Separate multiple PDFs with ',' or ';'. "
            "Use $ in a PDF filename to embed a password."
        ),
    )
    parser.add_argument(
        "--bank",
        help=f"Optional bank code: {SUPPORTED_BANK_CODES}. If omitted, bank is auto-detected.",
    )
    parser.add_argument("--pwd", help="Optional password for encrypted PDFs.")
    parser.add_argument(
        "--out",
        help="Optional output Excel name. Defaults to the first PDF filename.",
    )
    parser.add_argument(
        "--currency", choices=("INR", "USD", "EUR", "GBP", "SGD", "AUD", "CAD"),
        help="Confirmed currency shared by ALL input PDFs, for run totals (optional; no currency is assumed).",
    )
    return parser.parse_args(argv)


def setup_logger(log_file: Path) -> logging.Logger:
    logger = logging.getLogger("bank_stmt_parser")
    logger.setLevel(logging.DEBUG)
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()

    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    error_handler = logging.StreamHandler(sys.stderr)
    error_handler.setLevel(logging.ERROR)
    error_handler.setFormatter(formatter)
    logger.addHandler(error_handler)

    logger.propagate = False
    return logger


def progress_printer(row_number: int) -> None:
    print(f"\rProcessing row : {row_number}", end="", flush=True)


class RuntimeStatusTicker:
    def __init__(self, file_name: str, file_index: int, total_files: int, interval_seconds: int = 5):
        self.file_name = file_name
        self.file_index = file_index
        self.total_files = total_files
        self.interval_seconds = interval_seconds
        self._row_count = 0
        self._started_at = time.monotonic()
        self._stop_event = threading.Event()
        self._thread = threading.Thread(target=self._run, name="runtime-status", daemon=True)

    def start(self) -> None:
        self._thread.start()

    def update_rows(self, row_count: int) -> None:
        self._row_count = row_count

    def stop(self) -> None:
        self._stop_event.set()
        self._thread.join(timeout=1)

    def _run(self) -> None:
        while not self._stop_event.wait(self.interval_seconds):
            elapsed_seconds = int(time.monotonic() - self._started_at)
            print(
                (
                    f"\rStatus: running file {self.file_index}/{self.total_files} "
                    f"({self.file_name}) | rows parsed so far: {self._row_count} "
                    f"| elapsed: {elapsed_seconds}s"
                ),
                flush=True,
            )


def split_pdf_args(pdf_arg: str | list[str]) -> list[str]:
    """Return PDF names, treating only commas and semicolons as separators.

    ``argparse`` supplies a list because ``--pdf`` accepts one or more command
    tokens. Rejoining those tokens lets a filename such as ``My Statement.pdf``
    be passed without quotes, while preserving comma and semicolon as the
    explicit multi-file delimiters.
    """
    raw_value = pdf_arg if isinstance(pdf_arg, str) else " ".join(pdf_arg)
    files = [part.strip() for part in re.split(r"[;,]", raw_value) if part.strip()]
    if not files:
        raise ValueError("No input file was provided in --pdf.")
    return files


def build_output_stem(pdf_paths: list[Path], output_name: str | None) -> str:
    candidate = str(output_name or "").strip()
    if not candidate:
        candidate = split_pdf_filename_metadata(pdf_paths[0])[0]

    candidate_name = Path(candidate).name
    if candidate_name.lower().endswith(".xlsx"):
        candidate_name = Path(candidate_name).stem

    sanitized = INVALID_FILENAME_CHARS_RE.sub("_", candidate_name).strip(" .")
    if not sanitized:
        raise ValueError("Output name is empty after sanitization. Provide a valid value for --out.")
    return sanitized


def build_temp_work_dir(output_dir: Path) -> Path:
    timestamp = int(time.time() * 1000)
    return output_dir / f"_tmp_run_{timestamp}"


def cleanup_empty_temp_work_dirs(output_dir: Path) -> None:
    for candidate in output_dir.glob("_tmp_run_*"):
        if not candidate.is_dir():
            continue
        try:
            next(candidate.iterdir())
        except StopIteration:
            try:
                candidate.rmdir()
            except OSError:
                if sys.platform.startswith("win"):
                    windows_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
                    powershell_exe = windows_root / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
                    command = "Remove-Item -LiteralPath '{}' -Recurse -Force".format(
                        str(candidate).replace("'", "''")
                    )
                    subprocess.run(
                        [
                            str(powershell_exe),
                            "-NoProfile",
                            "-Command",
                            command,
                        ],
                        check=False,
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                    )
        except OSError:
            continue


def _coerce_balance_value(value: object) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def collect_negative_balance_rows(records: list[dict[str, object]]) -> list[dict[str, object]]:
    negative_rows: list[dict[str, object]] = []
    for row in records:
        balance_value = _coerce_balance_value(row.get("Balance"))
        if balance_value is None or balance_value >= 0:
            continue
        negative_rows.append(row)
    return negative_rows


def _format_negative_balance_row(row: dict[str, object]) -> str:
    return "Sno={sno} | Date={date} | Balance={balance} | Details={details}".format(
        sno=row.get("Sno", ""),
        date=row.get("Date", ""),
        balance=row.get("Balance", ""),
        details=row.get("Details", ""),
    )


def report_negative_balance_rows(
    *,
    records: list[dict[str, object]],
    file_name: str,
    bank_key: str,
    logger: logging.Logger,
    limit: int = 3,
) -> None:
    negative_rows = collect_negative_balance_rows(records)
    if not negative_rows:
        return

    logger.warning(
        "Negative balance sanity check flagged %s row(s) | file=%s | bank=%s",
        len(negative_rows),
        file_name,
        bank_key,
    )
    first_rows = negative_rows[:limit]
    last_rows = negative_rows[-limit:]

    print("=================", flush=True)
    print("**** WARNING ****", flush=True)
    print("=================", flush=True)
    print("", flush=True)
    print("************ -ve balance found sample first 3 records **********", flush=True)
    for row in first_rows:
        print(_format_negative_balance_row(row), flush=True)
    print("", flush=True)
    print("******************* -ve balance found sample last 3 records **********", flush=True)
    for row in last_rows:
        print(_format_negative_balance_row(row), flush=True)
    print("", flush=True)
    print("******************please cross check with pdf ************", flush=True)


def account_identity(pdf_path: Path, bank_key: str, logger) -> str | None:
    """Only native, unmasked first-page identity authorizes overlap removal."""
    try:
        header = read_first_page(pdf_path, allow_ocr=False)
        account = re.sub(r"[\s-]", "", header.values.get("Account Number") or "")
        if re.fullmatch(r"[0-9]{6,24}", account):
            return f"{bank_key}:{account}"
    except Exception as exc:
        logger.debug("Account identity unavailable for %s: %s", pdf_path.name, exc)
    logger.info("Overlap removal disabled for %s: account identity is unavailable or masked.", pdf_path.name)
    return None


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    # Help does not process statements and should not create a history row.
    if "--help" in argv or "-h" in argv:
        parse_args(argv)
        return 0
    try:
        history = RunHistory(database_path())
    except Exception as exc:
        print(f"Error: Cannot open run-history database ({type(exc).__name__}). "
              "Check BANKSTMT_DB_PATH and directory permissions.", file=sys.stderr)
        return 1
    code = 1
    status = "failed"
    try:
        with observe_ocr(history.record_ocr):
            code = _run(argv, history)
        status = "success" if code == 0 else "failed"
    except SystemExit as exc:
        code = int(exc.code or 0)
        status = "success" if code == 0 else "failed"
        if code:
            history.error("Invalid command-line arguments")
    except KeyboardInterrupt:
        code, status = 130, "interrupted"
        history.error("Processing interrupted by user")
        print("Processing interrupted.", file=sys.stderr)
    except Exception as exc:
        history.error(exc)
        print(f"Error: {history.redact(str(exc))}", file=sys.stderr)
    finally:
        try:
            history.finish(status)
            print(f"Run history: {history.path} | run ID: {history.run_id} | status: {status}")
        except Exception as exc:
            print(f"Error: Could not finalize run history ({type(exc).__name__}); "
                  f"run {history.run_id} may remain marked running.", file=sys.stderr)
            code = 1
    return code


def _run(argv, history: RunHistory) -> int:
    args = parse_args(argv)
    history.add_secret(args.pwd)
    try:
        requested_files = split_pdf_args(args.pdf)
        history.configure(requested_files, args.pwd, args.currency)
    except ValueError as exc:
        history.error(exc)
        print(f"Error: {history.redact(str(exc))}")
        return 2

    requested_bank = args.bank.strip().lower() if args.bank else None
    requested_bank = BANK_ALIASES.get(requested_bank, requested_bank)
    if requested_bank and requested_bank not in PARSERS:
        supported = ", ".join(sorted(PARSERS.keys()))
        history.error(f"Unsupported bank: {args.bank}")
        print(f"Error: Unsupported bank '{args.bank}'. Supported banks: {supported}")
        print(f"Example: {EXAMPLE_CMD}")
        return 2

    for item in history.inputs:
        item["bank"] = requested_bank

    project_root = PROJECT_ROOT
    src_root = project_root / "src"
    input_dir = project_root / "input"
    output_dir = project_root / "output"
    logs_dir = src_root / "logs"

    rules_path = input_dir / "Rules.xlsx"
    history.save(app_version=source_version(project_root),
                 rules_sha256=file_sha256(rules_path) if rules_path.is_file() else None)
    output_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    cleanup_empty_temp_work_dirs(output_dir)

    try:
        pdf_paths = []
        for index, file_arg in enumerate(requested_files):
            try:
                path = resolve_pdf_path(file_arg, src_root)
            except (FileNotFoundError, ValueError):
                history.start_input(index)
                raise
            history.resolved_input(index, path)
            pdf_paths.append(path)
        history.save()
        output_stem = build_output_stem(pdf_paths, args.out)
    except (FileNotFoundError, ValueError) as exc:
        history.error(exc)
        print(f"Error: {history.redact(str(exc))}")
        print(f"Example: {EXAMPLE_CMD}")
        return 2

    log_file = logs_dir / f"{output_stem}_run_{history.run_id}.log"
    logger = setup_logger(log_file)
    logger.addHandler(history.warning_counter)
    logger.info(
        "Run started | requested_bank=%s | files=%s | output=%s",
        requested_bank or "auto",
        [str(path) for path in pdf_paths],
        output_stem,
    )

    try:
        history.save(log_path=str(log_file))
        merged_records: list[dict[str, object]] = []
        source_pdf_passwords: list[str | None] = []
        saw_progress = False

        temp_dir = build_temp_work_dir(output_dir)
        try:
            for index, pdf_path in enumerate(pdf_paths, start=1):
                history.start_input(index - 1)
                history.inspect_source(pdf_path)
                _, filename_password = split_pdf_filename_metadata(pdf_path)
                pdf_password = args.pwd if args.pwd is not None else filename_password
                source_pdf_passwords.append(pdf_password)
                history.add_secret(pdf_password)
                history.stage = "decryption"
                readable_pdf_path = prepare_pdf_for_reading(pdf_path, pdf_password, temp_dir, logger)
                history.inspect_header(readable_pdf_path)
                history.stage = "detection"
                bank_key = requested_bank or bank_detector.detect_bank_from_pdf(readable_pdf_path, logger)
                parser_fn = PARSERS[bank_key]
                history.active["bank"] = bank_key
                parser_module = sys.modules.get(getattr(parser_fn, "__module__", ""))
                parser_file = getattr(parser_module, "__file__", None)
                parser_path = Path(parser_file) if parser_file else None
                if parser_path and parser_path.is_file():
                    history.active["parser_version"] = file_sha256(parser_path)
                history.stage = "parsing"
                history.save()

                print(
                    f"Starting file {index}/{len(pdf_paths)}: {pdf_path.name} | bank: {bank_key}",
                    flush=True,
                )
                logger.info(
                    "Processing file %s/%s: %s | bank=%s",
                    index,
                    len(pdf_paths),
                    pdf_path,
                    bank_key,
                )
                base_count = len(merged_records)
                ticker = RuntimeStatusTicker(
                    file_name=pdf_path.name,
                    file_index=index,
                    total_files=len(pdf_paths),
                )
                ticker.start()
                try:
                    records = parser_fn(
                        str(readable_pdf_path),
                        logger,
                        progress_cb=lambda row_number, offset=base_count, status=ticker: (
                            status.update_rows(offset + row_number),
                            progress_printer(offset + row_number),
                        ),
                    )
                finally:
                    ticker.stop()

                history.stage = "validation"
                history.active["transaction_count"] = len(records)
                validate_records(records, pdf_path.name)
                history.active.update(transaction_metrics(records))

                if records:
                    print()
                    saw_progress = True

                print(
                    f"Completed file {index}/{len(pdf_paths)}: {pdf_path.name} | rows parsed: {len(records)}",
                    flush=True,
                )
                report_negative_balance_rows(
                    records=records,
                    file_name=pdf_path.name,
                    bank_key=bank_key,
                    logger=logger,
                )
                history.stage = "reconciliation"
                result = reconcile(records, str(readable_pdf_path), logger)
                history.active["reconciliation_status"] = result.status
                if result.status == "failed":
                    raise ValueError(f"Reconciliation failed for {pdf_path.name}: " + "; ".join(result.mismatches))
                if result.status == "unavailable":
                    print(f"Reconciliation unavailable for {pdf_path.name}: no printed summary totals found.", flush=True)

                if len(pdf_paths) > 1:
                    account_key = account_identity(readable_pdf_path, bank_key, logger)
                    for row in records:
                        row["Source"] = pdf_path.name
                        row["_Source_Id"] = str(pdf_path.resolve())
                        row["_Account_Key"] = account_key

                merged_records.extend(records)
                history.complete_input()
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
            cleanup_empty_temp_work_dirs(output_dir)

        if not saw_progress:
            print("\rProcessing row : 0")

        logger.info("Total merged rows parsed: %s", len(merged_records))

        history.stage = "normalization"
        statement_df = records_to_dataframe(
            merged_records,
            include_source=len(pdf_paths) > 1,
        )
        rows_before_deduplication = len(statement_df)
        statement_df = remove_exact_duplicate_transactions(statement_df).reset_index(drop=True)
        statement_df["Sno"] = range(1, len(statement_df) + 1)
        duplicates_removed = rows_before_deduplication - len(statement_df)
        if duplicates_removed:
            print(
                f"Removed exact duplicate transactions: {duplicates_removed}",
                flush=True,
            )
            logger.info(
                "Removed %s exact duplicate transaction row(s) from merged statements",
                duplicates_removed,
            )

        history.merged_metrics(statement_df.to_dict("records"), duplicates_removed)
        history.stage = "intermediate_export"
        intermediate_output = output_dir / "output.xlsx"
        print(f"Writing merged intermediate workbook: {intermediate_output}", flush=True)
        write_output_excel(statement_df, intermediate_output)
        logger.info("Intermediate output written: %s", intermediate_output)

        rules_path = input_dir / "Rules.xlsx"
        history.stage = "final_export"
        print("Building final workbook...", flush=True)
        final_output = build_final_workbook(
            statement_df=statement_df,
            rules_path=rules_path,
            output_dir=output_dir,
            pdf_stem=output_stem,
            logger=logger,
            source_pdf_paths=pdf_paths,
            source_pdf_passwords=source_pdf_passwords,
            include_source=len(pdf_paths) > 1,
        )

        history.save(output_path=str(final_output))
        print(f"Intermediate output written: {intermediate_output}")
        print(f"Files processed: {', '.join(path.name for path in pdf_paths)}")
        print(f"Final output written: {final_output}")
        print(f"Log file written: {log_file}")

        logger.info("Run completed successfully")
        return 0

    except Exception as exc:  # noqa: BLE001
        history.error(exc)
        print()
        logger.exception("Run failed")
        print(f"Error: {exc}")
        print(f"Check log file: {log_file}")
        return 1
    finally:
        for handler in list(logger.handlers):
            logger.removeHandler(handler)
            handler.close()


if __name__ == "__main__":
    raise SystemExit(main())
