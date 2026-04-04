"""Intermediate Excel writer."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

INTERMEDIATE_FILE_NAME = "output.xlsx"


def write_intermediate_output(df: pd.DataFrame, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_dir / INTERMEDIATE_FILE_NAME
    df.to_excel(output_path, index=False)
    return output_path
