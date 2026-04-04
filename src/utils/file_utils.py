"""Path and file utilities for the statement pipeline."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[2]
INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"


def ensure_project_folders() -> None:
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def resolve_input_pdf(pdf_file_name: str) -> Path:
    pdf_path = INPUT_DIR / pdf_file_name
    if not pdf_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {pdf_path}")
    return pdf_path


def resolve_rules_file(file_name: str = "Rules.xlsx") -> Path:
    rules_path = INPUT_DIR / file_name
    if not rules_path.exists():
        raise FileNotFoundError(f"Rules file not found: {rules_path}")
    return rules_path


def append_timestamp_if_exists(path: Path) -> Path:
    if not path.exists():
        return path

    timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
    candidate = path.with_name(f"{path.stem}_{timestamp}{path.suffix}")

    suffix_counter = 1
    while candidate.exists():
        candidate = path.with_name(f"{path.stem}_{timestamp}_{suffix_counter}{path.suffix}")
        suffix_counter += 1

    return candidate


def build_final_output_path(pdf_file_name: str) -> Path:
    base_name = Path(pdf_file_name).stem
    target_path = OUTPUT_DIR / f"{base_name}.xlsx"
    return append_timestamp_if_exists(target_path)
