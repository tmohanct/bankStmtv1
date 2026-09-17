"""Shared intermediate workbook export."""
from pathlib import Path

import pandas as pd

from src.transform.normalize import sanitize_cheque_column
from src.export.excel_safety import sanitize_excel_frame, force_leading_equals_to_text


def write_output_excel(frame: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    private_columns = [column for column in frame.columns if column.startswith("_")]
    output_frame = frame.drop(columns=private_columns, errors="ignore")
    output_frame = sanitize_excel_frame(sanitize_cheque_column(output_frame))
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        output_frame.to_excel(writer, index=False, sheet_name="Statement")
        force_leading_equals_to_text(writer.book)


def write_intermediate_output(df: pd.DataFrame, output_dir: Path) -> Path:
    path = output_dir / "output.xlsx"
    write_output_excel(df, path)
    return path
