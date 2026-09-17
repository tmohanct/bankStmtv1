"""Normalize parser records and conservatively merge overlapping statements."""
from __future__ import annotations

import re
from difflib import SequenceMatcher
from typing import Any
import pandas as pd
from src.utils.statement_utils import (OUTPUT_COLUMNS, DEDUPLICATION_EXCLUDED_COLUMNS, clean_cell, clean_detail, CHEQUE_INTEGER_FLOAT_RE, CHEQUE_DIGITS_ONLY_RE, CHEQUE_DETAIL_HINT_RE, NON_CHEQUE_DETAIL_HINT_RE, COMPACT_CHEQUE_NUMBER_DETAIL_PATTERNS, CHEQUE_NUMBER_DETAIL_PATTERNS)

def normalize_cheque_number(value: Any, details: Any = "") -> str:
    text = clean_cell(value)
    if not text or text == "-":
        return ""

    integer_float_match = CHEQUE_INTEGER_FLOAT_RE.fullmatch(text)
    if integer_float_match is not None:
        text = integer_float_match.group("digits")

    if not CHEQUE_DIGITS_ONLY_RE.fullmatch(text):
        return ""
    if set(text) == {"0"}:
        return ""

    detail_text = clean_cell(details)
    if detail_text:
        if CHEQUE_DETAIL_HINT_RE.search(detail_text):
            return text
        if NON_CHEQUE_DETAIL_HINT_RE.search(detail_text):
            return ""

    return text


def _normalize_extracted_cheque_candidate(value: Any, details: Any) -> str:
    compact_value = re.sub(r"[\s-]+", "", clean_cell(value))
    integer_float_match = CHEQUE_INTEGER_FLOAT_RE.fullmatch(compact_value)
    if integer_float_match is not None:
        compact_value = integer_float_match.group("digits")

    if not CHEQUE_DIGITS_ONLY_RE.fullmatch(compact_value):
        return ""
    if set(compact_value) == {"0"}:
        return ""
    return compact_value


def extract_cheque_number_from_details(details: Any, detail_clean: Any = "") -> str:
    detail_text = clean_cell(details)
    compact_detail = clean_detail(detail_clean or detail_text).upper()
    if not detail_text and not compact_detail:
        return ""

    for pattern in COMPACT_CHEQUE_NUMBER_DETAIL_PATTERNS:
        match = pattern.search(compact_detail)
        if match is None:
            continue
        candidate = _normalize_extracted_cheque_candidate(match.group(1), detail_text)
        if candidate:
            return candidate

    for pattern in CHEQUE_NUMBER_DETAIL_PATTERNS:
        for match in pattern.finditer(detail_text):
            candidate = _normalize_extracted_cheque_candidate(match.group(1), detail_text)
            if candidate:
                return candidate
    return ""


def sanitize_cheque_column(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None or frame.empty or "Cheque No" not in frame.columns:
        return frame

    output = frame.copy()
    if "Details" in output.columns:
        detail_series = output["Details"]
    else:
        detail_series = pd.Series([""] * len(output), index=output.index)
    if "Detail_Clean" in output.columns:
        detail_clean_series = output["Detail_Clean"]
    else:
        detail_clean_series = pd.Series([""] * len(output), index=output.index)

    sanitized_values: list[str] = []
    for cheque_value, detail_value, detail_clean_value in zip(
        output["Cheque No"],
        detail_series,
        detail_clean_series,
    ):
        normalized = normalize_cheque_number(cheque_value, detail_value)
        if not normalized:
            normalized = extract_cheque_number_from_details(detail_value, detail_clean_value)
        sanitized_values.append(normalized)

    output["Cheque No"] = sanitized_values
    return output


def records_to_dataframe(
    records: list[dict[str, Any]],
    *,
    include_source: bool = True,
) -> pd.DataFrame:
    output_columns = [
        column for column in OUTPUT_COLUMNS if include_source or column != "Source"
    ]
    if not records:
        return pd.DataFrame(columns=output_columns)

    frame = pd.DataFrame(records)
    for col in output_columns:
        if col not in frame.columns:
            frame[col] = None
    frame = frame[output_columns + [c for c in ("_Account_Key", "_Source_Id") if c in frame.columns]]
    return sanitize_cheque_column(frame)


def remove_exact_duplicate_transactions(frame: pd.DataFrame) -> pd.DataFrame:
    """Remove confirmed sequence overlaps across PDFs of the same account.

    Preserve every occurrence within one PDF. An unmasked bank/account key and
    a stable source ID are required. A lone matching row is ambiguous; require
    at least two adjacent matching rows, including their running balances.
    """
    if frame.empty or not {"_Account_Key", "_Source_Id"}.issubset(frame.columns):
        return frame.copy()
    columns = [c for c in OUTPUT_COLUMNS if c in frame.columns and c not in DEDUPLICATION_EXCLUDED_COLUMNS]
    if "Balance" not in columns:
        return frame.copy()
    keep = [True] * len(frame)
    previous = {}
    work = frame.reset_index(drop=True)
    for (account, source), group in work.groupby(["_Account_Key", "_Source_Id"], sort=False, dropna=False):
        if pd.isna(account) or pd.isna(source) or not str(account).strip() or not str(source).strip():
            continue
        positions = list(group.index)
        keys = [tuple(None if pd.isna(v) else v for v in row) for row in group[columns].itertuples(index=False, name=None)]
        for prior_source, prior_keys in previous.get(account, []):
            if prior_source == source:
                continue
            for block in SequenceMatcher(None, prior_keys, keys, autojunk=False).get_matching_blocks():
                if block.size < 2:
                    continue
                matching_positions = positions[block.b:block.b + block.size]
                if work.loc[matching_positions, "Balance"].isna().any():
                    continue
                for position in matching_positions:
                    keep[position] = False
        previous.setdefault(account, []).append((source, keys))
    return frame.iloc[[i for i, retain in enumerate(keep) if retain]].copy()
