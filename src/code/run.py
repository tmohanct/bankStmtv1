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

CODE_ROOT = Path(__file__).resolve().parent
SRC_ROOT = CODE_ROOT.parent
for import_root in (str(CODE_ROOT), str(SRC_ROOT)):
    while import_root in sys.path:
        sys.path.remove(import_root)
for index, import_root in enumerate((str(CODE_ROOT), str(SRC_ROOT))):
    sys.path.insert(index, import_root)

import axis_parser
import bank_detector
import bob_parser
import bom_parser
import canara_parser
import centralbank_parser
import cub_parser
import federal_parser
import hdfc_parser
import icici_parser
import idbi_parser
import idfc_parser
import indian_parser
import indus_parser
import iob_parser
import kvb_parser
import kotak_parser
import pnb_parser
import sbi_parser
import southind_parser
import unionbank_parser
from final_excel_builder import build_final_workbook
from utils import (
    prepare_pdf_for_reading,
    reconcile,
    records_to_dataframe,
    resolve_pdf_path,
    split_pdf_filename_metadata,
    write_output_excel,
)

EXAMPLE_CMD = 'python run.py --pdf "file1;file2" --bank icici --pwd mypassword --out outputname'
INVALID_FILENAME_CHARS_RE = re.compile(r'[<>:"/\\\\|?*]+')

PARSERS = {
    "axis": axis_parser.parse,
    "bob": bob_parser.parse,
    "bom": bom_parser.parse,
    "canara": canara_parser.parse,
    "central": centralbank_parser.parse,
    "cub": cub_parser.parse,
    "federal": federal_parser.parse,
    "hdfc": hdfc_parser.parse,
    "icici": icici_parser.parse,
    "idbi": idbi_parser.parse,
    "idfc": idfc_parser.parse,
    "indian": indian_parser.parse,
    "indus": indus_parser.parse,
    "iob": iob_parser.parse,
    "kvb": kvb_parser.parse,
    "kotak": kotak_parser.parse,
    "pnb": pnb_parser.parse,
    "sbi": sbi_parser.parse,
    "southind": southind_parser.parse,
    "unionbank": unionbank_parser.parse,
}
SUPPORTED_BANK_CODES = "/".join(sorted(PARSERS))


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
        help='Semicolon-separated PDF names from input/ without .pdf extension. In PowerShell, wrap the value in quotes. Use $ in the PDF filename to embed a password.',
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
    return parser.parse_args(argv)


def setup_logger(log_file: Path) -> logging.Logger:
    logger = logging.getLogger("bank_stmt_parser")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

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


def split_pdf_args(pdf_arg: str) -> list[str]:
    files = [part.strip() for part in pdf_arg.split(";") if part.strip()]
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


def main(argv=None) -> int:
    args = parse_args(argv)

    requested_bank = args.bank.strip().lower() if args.bank else None
    if requested_bank and requested_bank not in PARSERS:
        supported = ", ".join(sorted(PARSERS.keys()))
        print(f"Error: Unsupported bank '{args.bank}'. Supported banks: {supported}")
        print(f"Example: {EXAMPLE_CMD}")
        return 2

    src_root = Path(__file__).resolve().parents[1]
    project_root = src_root.parent
    input_dir = project_root / "input"
    output_dir = project_root / "output"
    logs_dir = src_root / "logs"

    output_dir.mkdir(parents=True, exist_ok=True)
    logs_dir.mkdir(parents=True, exist_ok=True)
    cleanup_empty_temp_work_dirs(output_dir)

    try:
        requested_files = split_pdf_args(args.pdf)
        pdf_paths = [resolve_pdf_path(file_arg, src_root) for file_arg in requested_files]
        output_stem = build_output_stem(pdf_paths, args.out)
    except (FileNotFoundError, ValueError) as exc:
        print(f"Error: {exc}")
        print(f"Example: {EXAMPLE_CMD}")
        return 2

    log_file = logs_dir / f"{output_stem}.log"
    logger = setup_logger(log_file)
    logger.info(
        "Run started | requested_bank=%s | files=%s | output=%s",
        requested_bank or "auto",
        [str(path) for path in pdf_paths],
        output_stem,
    )

    try:
        merged_records: list[dict[str, object]] = []
        saw_progress = False

        temp_dir = build_temp_work_dir(output_dir)
        try:
            for index, pdf_path in enumerate(pdf_paths, start=1):
                _, filename_password = split_pdf_filename_metadata(pdf_path)
                pdf_password = args.pwd if args.pwd is not None else filename_password
                readable_pdf_path = prepare_pdf_for_reading(pdf_path, pdf_password, temp_dir, logger)
                bank_key = requested_bank or bank_detector.detect_bank_from_pdf(readable_pdf_path, logger)
                parser_fn = PARSERS[bank_key]

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
                reconcile(records, str(readable_pdf_path), logger)

                for row in records:
                    row["Source"] = pdf_path.name

                merged_records.extend(records)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
            cleanup_empty_temp_work_dirs(output_dir)

        if not saw_progress:
            print("\rProcessing row : 0")

        for row_number, row in enumerate(merged_records, start=1):
            row["Sno"] = row_number

        logger.info("Total merged rows parsed: %s", len(merged_records))

        statement_df = records_to_dataframe(merged_records)
        intermediate_output = output_dir / "output.xlsx"
        print(f"Writing merged intermediate workbook: {intermediate_output}", flush=True)
        write_output_excel(statement_df, intermediate_output)
        logger.info("Intermediate output written: %s", intermediate_output)

        rules_path = input_dir / "Rules.xlsx"
        print("Building final workbook...", flush=True)
        final_output = build_final_workbook(
            statement_df=statement_df,
            rules_path=rules_path,
            output_dir=output_dir,
            pdf_stem=output_stem,
            logger=logger,
        )

        print(f"Intermediate output written: {intermediate_output}")
        print(f"Files processed: {', '.join(path.name for path in pdf_paths)}")
        print(f"Final output written: {final_output}")
        print(f"Log file written: {log_file}")

        logger.info("Run completed successfully")
        return 0

    except Exception as exc:  # noqa: BLE001
        print()
        logger.exception("Run failed")
        print(f"Error: {exc}")
        print(f"Check log file: {log_file}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())



