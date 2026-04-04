"""Validation helpers for normalized transactions."""

from __future__ import annotations

import pandas as pd

REQUIRED_COLUMNS = ["Txn_Date", "Description", "Debit", "Credit", "Balance"]


def validate_transactions(df: pd.DataFrame) -> pd.DataFrame:
    missing = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"Normalized output is missing required columns: {missing}")

    # TODO: Add richer data quality checks (date parsing, amount coercion, duplicate detection).
    both_debit_and_credit = df["Debit"].notna() & df["Credit"].notna()
    if both_debit_and_credit.any():
        raise ValueError("Rows cannot have both Debit and Credit populated.")

    return df
