"""CLI entry point for bank statement analysis pipeline."""

from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

from export.excel_writer import write_intermediate_output
from export.final_excel_builder import build_final_workbook
from parsers.detector import get_parser, supported_banks
from transform.normalize import normalize_transactions
from transform.validate import validate_transactions
from utils.file_utils import OUTPUT_DIR, ensure_project_folders, resolve_input_pdf, resolve_rules_file


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Bank statement analysis tool")
    parser.add_argument("--bank", required=True, choices=supported_banks(), help="Bank code")
    parser.add_argument("--pdf", required=True, help="PDF file name inside input/ directory")
    return parser.parse_args()


def load_rules(rules_path: Path) -> pd.DataFrame:
    try:
        return pd.read_excel(rules_path)
    except Exception as exc:  # pragma: no cover - runtime path
        raise RuntimeError(
            f"Unable to read rules workbook at '{rules_path}'. Ensure input/Rules.xlsx is a valid .xlsx file."
        ) from exc


def run_pipeline(bank_code: str, pdf_file_name: str) -> tuple[Path, Path]:
    ensure_project_folders()

    pdf_path = resolve_input_pdf(pdf_file_name)
    rules_path = resolve_rules_file()
    rules_df = load_rules(rules_path)

    parser = get_parser(bank_code)
    parsed_df = parser.parse(pdf_path=pdf_path, rules_df=rules_df)

    normalized_df = normalize_transactions(parsed_df, bank_code=bank_code)
    validated_df = validate_transactions(normalized_df)

    intermediate_path = write_intermediate_output(validated_df, output_dir=OUTPUT_DIR)
    final_path = build_final_workbook(validated_df, pdf_file_name=pdf_file_name, output_dir=OUTPUT_DIR)
    return intermediate_path, final_path


def main() -> None:
    args = parse_args()
    intermediate_path, final_path = run_pipeline(args.bank, args.pdf)
    print(f"Intermediate workbook written to: {intermediate_path}")
    print(f"Final workbook written to: {final_path}")


if __name__ == "__main__":
    main()
