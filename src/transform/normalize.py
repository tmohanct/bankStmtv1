"""Normalization helpers for parser output."""

from __future__ import annotations

import pandas as pd

TARGET_COLUMNS = [
    "Txn_Date",
    "Value_Date",
    "Description",
    "Debit",
    "Credit",
    "Balance",
    "Currency",
    "Bank",
    "Account_Number",
    "Reference",
    "Source_Page",
]


def normalize_transactions(parsed_df: pd.DataFrame, bank_code: str) -> pd.DataFrame:
    rename_map = {
        "Date": "Txn_Date",
        "ValueDate": "Value_Date",
        "Narration": "Description",
        "Txn_Ref": "Reference",
        "Page": "Source_Page",
    }

    normalized = parsed_df.rename(columns=rename_map).copy()

    if normalized.empty:
        return pd.DataFrame(columns=TARGET_COLUMNS)

    normalized["Bank"] = bank_code.upper()
    normalized["Currency"] = normalized.get("Currency", "INR")

    for column in TARGET_COLUMNS:
        if column not in normalized.columns:
            normalized[column] = pd.NA

    return normalized[TARGET_COLUMNS]
