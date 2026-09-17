"""Date-ordered schedules for text-search rule sheets."""

import re

import pandas as pd

from src.transform.analysis import _coerce_excel_date, _ensure_columns, _normalize_rule_text
from src.utils.statement_utils import CHEQUE_DETAIL_HINT_RE
from src.utils.text_utils import clean_cell


SCHEDULE_COLUMNS = [
    "DUE NO", "NAME", "ACTUAL DATE", "DAY", "FREQ", "CHEQUE NO",
    "DR", "CR", "PAID DATE", "OTHER NAME",
]
MODE_PATTERN = re.compile(
    r"(?<![A-Z])(?:IMPS|UPI|RTGS|NEFT|NACH|ACH|ECS|ATM|POS|CASH)"
    r"(?=[^A-Z]|$|[A-Z]{2,6}\d{6})",
    re.IGNORECASE,
)
WEEKDAYS = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")


def transaction_mode(row) -> str:
    """Show a cheque identifier or the evidenced mode, never a transfer reference."""
    detail = clean_cell(row.get("Details")) or clean_cell(row.get("Detail_Clean"))
    cheque = clean_cell(row.get("Cheque No"))
    if cheque and CHEQUE_DETAIL_HINT_RE.search(detail):
        return cheque
    mode = MODE_PATTERN.search(detail)
    if mode:
        return mode.group().upper()
    if cheque:
        return cheque
    if CHEQUE_DETAIL_HINT_RE.search(detail):
        return "CHEQUE"
    return ""


def order_rule_transactions(frame: pd.DataFrame) -> pd.DataFrame:
    """Use one stable transaction order for both sides of a rule worksheet."""
    work = _ensure_columns(frame).copy()
    work["__date"] = pd.to_datetime(work["Date"].map(_coerce_excel_date), errors="coerce")
    work["__sno"] = pd.to_numeric(work["Sno"], errors="coerce")
    return (work.sort_values(["__date", "__sno"], kind="stable", na_position="last")
            .drop(columns=["__date", "__sno"]).reset_index(drop=True))


def build_rule_schedule(
    frame: pd.DataFrame, rules: list[dict], sheet_name: str, *, align_rows: bool = False,
) -> pd.DataFrame:
    """Use only text rules for this destination; first matching rule supplies NAME.

    A transaction matching multiple rules stays one row. Amount-only matches in
    a mixed destination are excluded because they have no matching search name.
    With align_rows, reserve blank schedule rows for those transactions so the
    two tables remain aligned. The first due starts at the first preceding credit;
    later dues measure the interval since the previous debit.
    """
    text_rules = [
        rule for rule in rules
        if rule.get("category", "TEXT") != "AMT"
        and rule["sheet_name"].strip().casefold() == sheet_name.strip().casefold()
    ]
    if not text_rules or frame.empty:
        return pd.DataFrame(columns=SCHEDULE_COLUMNS)

    work = order_rule_transactions(frame)
    def matching_name(detail):
        key = _normalize_rule_text(clean_cell(detail))
        return next((rule["name"] for rule in text_rules if rule["name_clean"] in key), None)

    work["__name"] = work["Detail_Clean"].map(matching_name)
    if not work["__name"].notna().any():
        return pd.DataFrame(columns=SCHEDULE_COLUMNS)
    if not align_rows:
        work = work[work["__name"].notna()].copy()
    work["__date"] = pd.to_datetime(work["Date"].map(_coerce_excel_date), errors="coerce")

    rows = []
    due_no = 0
    previous_debit_date = None
    first_credit_date = None
    for _, row in work.iterrows():
        if pd.isna(row["__name"]):
            rows.append([None] * len(SCHEDULE_COLUMNS))
            continue
        actual_date = row["__date"]
        valid_date = pd.notna(actual_date)
        debit, credit = row["Debit"], row["Credit"]
        is_credit = credit != 0
        if is_credit and valid_date and first_credit_date is None and due_no == 0:
            first_credit_date = actual_date
        if debit > 0:
            due_no += 1
        freq = None
        if debit > 0 and not is_credit:
            baseline = first_credit_date if due_no == 1 else previous_debit_date
            if valid_date and baseline is not None:
                freq = (actual_date.date() - baseline.date()).days
            previous_debit_date = actual_date if valid_date else None
        rows.append([
            due_no if debit > 0 else None, row["__name"],
            actual_date.to_pydatetime() if valid_date else clean_cell(row["Date"]),
            WEEKDAYS[actual_date.weekday()] if valid_date else None, freq,
            transaction_mode(row), debit if debit != 0 else None,
            credit if credit != 0 else None, None, None,
        ])
    return pd.DataFrame(rows, columns=SCHEDULE_COLUMNS)
