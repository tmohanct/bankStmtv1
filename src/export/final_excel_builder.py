"""Final workbook builder."""

from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd

from utils.file_utils import build_final_output_path


def build_final_workbook(df: pd.DataFrame, pdf_file_name: str, output_dir: Path) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    final_path = build_final_output_path(pdf_file_name=pdf_file_name)

    metadata = pd.DataFrame(
        [
            {
                "source_pdf": pdf_file_name,
                "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "records": len(df),
            }
        ]
    )

    with pd.ExcelWriter(final_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Transactions", index=False)
        metadata.to_excel(writer, sheet_name="Metadata", index=False)

    return final_path
