"""Rule matching and transaction summaries used by the final workbook."""

import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd

from src.utils.statement_utils import OUTPUT_COLUMNS, compact_detail_key
from src.transform.normalize import sanitize_cheque_column
from src.transform.cheque_returns import is_cheque_return, is_nonposting_cheque_return

DATE_INPUT_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y")


def _coerce_excel_date(value: Any) -> datetime | None:
    if value in (None, ""):
        return None

    if isinstance(value, datetime):
        return value

    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())

    if hasattr(value, "to_pydatetime"):
        try:
            py_dt = value.to_pydatetime()
            if isinstance(py_dt, datetime):
                return py_dt
            if isinstance(py_dt, date):
                return datetime.combine(py_dt, datetime.min.time())
        except Exception:  # noqa: BLE001
            pass

    text = str(value).strip()
    if not text:
        return None

    for fmt in DATE_INPUT_FORMATS:
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _ensure_columns(frame: pd.DataFrame) -> pd.DataFrame:
    if frame is None or frame.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    output = frame.copy()
    for col in OUTPUT_COLUMNS:
        if col not in output.columns:
            output[col] = None
    for amount_col in ("Debit", "Credit"):
        output[amount_col] = pd.to_numeric(output[amount_col], errors="coerce").fillna(0.0)
    output = sanitize_cheque_column(output)
    return output[OUTPUT_COLUMNS]


def _first_present_column(lower_map: dict[str, Any], *keys: str) -> Any | None:
    for key in keys:
        if key in lower_map:
            return lower_map[key]
    return None


def _parse_rule_amount(value: Any) -> float | None:
    if value is None:
        return None

    text = str(value).strip()
    if not text:
        return None

    cleaned = text.replace(",", "").replace(" ", "")
    try:
        return float(cleaned)
    except ValueError:
        return None


def _normalize_rule_text(value: Any) -> str:
    return compact_detail_key(value).upper()


def _load_rules(rules_path: Path, logger) -> list[dict[str, Any]]:
    if not rules_path.is_file():
        logger.warning("Rules file not found: %s", rules_path)
        return []

    rules_df = pd.read_excel(rules_path, sheet_name=0)
    if rules_df.empty:
        logger.info("Rules file is empty: %s", rules_path)
        return []

    lower_map = {str(col).strip().lower(): col for col in rules_df.columns}

    category_col = _first_present_column(lower_map, "category")
    subcategory_col = _first_present_column(
        lower_map,
        "searchstring",
        "subcategory",
        "sub_category",
        "sub category",
        "name",
        "keyword",
        "search_name",
        "searchname",
        "match",
    )
    sheet_col = _first_present_column(lower_map, "sheetname", "sheet_name", "sheet")
    order_col = _first_present_column(lower_map, "sheet_order", "sheetorder", "sheet order", "order")

    if subcategory_col is None or sheet_col is None:
        logger.warning(
            "Rules missing required columns. Found columns: %s",
            list(rules_df.columns),
        )
        return []

    work = rules_df.copy()
    work["__row_order"] = range(len(work))
    work = work.dropna(subset=[subcategory_col, sheet_col])

    if order_col is not None:
        work["__sheet_order"] = pd.to_numeric(work[order_col], errors="coerce")
    else:
        work["__sheet_order"] = work["__row_order"] + 1

    work = work.sort_values(by=["__sheet_order", "__row_order"], na_position="last")

    rules: list[dict[str, Any]] = []
    for _, row in work.iterrows():
        raw_category = str(row[category_col]).strip() if category_col is not None and pd.notna(row[category_col]) else "Text"
        raw_name = str(row[subcategory_col]).strip()
        raw_sheet = str(row[sheet_col]).strip()
        normalized_category = raw_category.upper()
        clean_name = _normalize_rule_text(raw_name)
        if not raw_name or not raw_sheet:
            continue

        rule: dict[str, Any] = {
            "category": normalized_category,
            "name": raw_name,
            "name_clean": clean_name,
            "sheet_name": raw_sheet,
        }

        if normalized_category == "AMT":
            amount_value = _parse_rule_amount(raw_name)
            if amount_value is None:
                logger.warning("Skipping Amt rule with invalid amount '%s' for sheet %s", raw_name, raw_sheet)
                continue
            rule["amount_value"] = amount_value
        elif not clean_name:
            continue

        rules.append(rule)

    logger.info("Loaded %s rule(s) from %s", len(rules), rules_path)
    return rules


def _build_text_rule_sheet(statement_df: pd.DataFrame, rule: dict[str, Any]) -> pd.DataFrame:
    if statement_df.empty or "Detail_Clean" not in statement_df.columns:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    detail_series = statement_df["Detail_Clean"].fillna("").astype(str).map(compact_detail_key).str.upper()
    mask = detail_series.str.contains(rule["name_clean"], na=False)
    matched = statement_df[mask].copy()
    return _ensure_columns(matched)


def _build_amount_rule_sheet(statement_df: pd.DataFrame, rule: dict[str, Any]) -> pd.DataFrame:
    if statement_df.empty:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    target_amount = float(rule["amount_value"])
    work = statement_df.copy()
    work["Debit"] = pd.to_numeric(work["Debit"], errors="coerce")
    work["Credit"] = pd.to_numeric(work["Credit"], errors="coerce")

    debit_match = work["Debit"].notna() & (work["Debit"].sub(target_amount).abs() <= 0.005)
    credit_match = work["Credit"].notna() & (work["Credit"].sub(target_amount).abs() <= 0.005)
    matched = work[debit_match | credit_match].copy()
    if matched.empty:
        return _ensure_columns(matched)

    matched["__amt_group"] = 1
    matched.loc[credit_match.reindex(matched.index, fill_value=False), "__amt_group"] = 2
    matched["__sort_date"] = matched["Date"].apply(
        lambda value: _coerce_excel_date(value) or datetime.max
    )
    matched = matched.sort_values(
        by=["__amt_group", "__sort_date", "Sno"],
        ascending=[True, True, True],
        na_position="last",
    )
    matched = matched.drop(columns=["__amt_group", "__sort_date"])
    return _ensure_columns(matched)


def _merge_rule_sheet_frames(frames: list[pd.DataFrame]) -> pd.DataFrame:
    if not frames:
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    merged = pd.concat(frames, axis=0, ignore_index=False)
    if merged.empty:
        return _ensure_columns(merged)

    if "Sno" in merged.columns:
        merged["__rule_sno"] = pd.to_numeric(merged["Sno"], errors="coerce")
        if merged["__rule_sno"].notna().any():
            merged = merged.sort_values(by=["__rule_sno"], kind="stable", na_position="last")
            merged = merged.drop_duplicates(subset=["__rule_sno"], keep="first")
            merged = merged.drop(columns=["__rule_sno"])
            return _ensure_columns(merged)
        merged = merged.drop(columns=["__rule_sno"])

    merged = merged.drop_duplicates(keep="first")
    return _ensure_columns(merged)


def _build_rule_sheets(statement_df: pd.DataFrame, rules: list[dict[str, Any]], logger):
    if statement_df.empty:
        return []

    sheet_order: list[str] = []
    grouped_sheet_names: dict[str, str] = {}
    grouped_frames: dict[str, list[pd.DataFrame]] = {}
    grouped_rule_names: dict[str, list[str]] = {}

    for rule in rules:
        category = rule.get("category", "TEXT")
        if category == "AMT":
            matched = _build_amount_rule_sheet(statement_df, rule)
        else:
            matched = _build_text_rule_sheet(statement_df, rule)

        if matched.empty:
            continue
        logger.info("Rule matched: category=%s key=%s rows=%s sheet=%s", category, rule["name"], len(matched), rule["sheet_name"])

        sheet_name = str(rule["sheet_name"]).strip()
        sheet_key = sheet_name.casefold()
        if sheet_key not in grouped_frames:
            sheet_order.append(sheet_key)
            grouped_sheet_names[sheet_key] = sheet_name
            grouped_frames[sheet_key] = []
            grouped_rule_names[sheet_key] = []

        grouped_frames[sheet_key].append(_ensure_columns(matched))
        grouped_rule_names[sheet_key].append(str(rule["name"]))

    sheets: list[tuple[str, pd.DataFrame]] = []
    for sheet_key in sheet_order:
        requested_name = grouped_sheet_names[sheet_key]
        merged = _merge_rule_sheet_frames(grouped_frames[sheet_key])
        if merged.empty:
            continue
        logger.info(
            "Merged %s rule(s) into sheet=%s rows=%s keys=%s",
            len(grouped_rule_names[sheet_key]),
            requested_name,
            len(merged),
            grouped_rule_names[sheet_key],
        )
        sheets.append((requested_name, merged))

    return sheets


def _to_numeric(frame: pd.DataFrame, column: str) -> pd.Series:
    if column not in frame.columns:
        return pd.Series([pd.NA] * len(frame), index=frame.index)
    return pd.to_numeric(frame[column], errors="coerce")


def _build_cheque_sheet(statement_df: pd.DataFrame) -> pd.DataFrame:
    if statement_df.empty:
        return _ensure_columns(statement_df)

    work = statement_df.copy()
    cheque_series = work["Cheque No"].fillna("").astype(str).str.strip()
    work = work[cheque_series != ""].copy()
    if work.empty:
        return _ensure_columns(work)

    def cheque_sort_key(value: Any) -> tuple[int, str]:
        text = str(value).strip()
        digits = re.sub(r"\D", "", text)
        if digits:
            return int(digits), text
        return 10**12, text

    work["__cheque_sort"] = work["Cheque No"].apply(cheque_sort_key)
    work = work.sort_values(by=["__cheque_sort", "Cheque No"])
    work = work.drop(columns=["__cheque_sort"])
    return _ensure_columns(work)


def _is_return_reject_detail(value: Any, cheque_number: Any = "", detail_clean: Any = "") -> bool:
    return is_cheque_return(value, cheque_number, detail_clean)


def _build_return_reject_sheet(statement_df: pd.DataFrame) -> pd.DataFrame:
    work = _ensure_columns(statement_df)
    mask = [
        _is_return_reject_detail(details, number, cleaned)
        for details, number, cleaned in zip(work["Details"], work["Cheque No"], work["Detail_Clean"])
    ]
    return work.loc[mask].copy()


def _cheque_sort_value(value: Any) -> tuple[int, str]:
    text = str(value or "").strip()
    digits = re.sub(r"\D", "", text)
    if digits:
        return int(digits), text
    return 10**12, text


def _build_repeat_sheet(statement_df: pd.DataFrame, amount_column: str) -> pd.DataFrame:
    if statement_df.empty:
        return _ensure_columns(statement_df)

    work = statement_df.copy()
    numeric = _to_numeric(work, amount_column)
    work[amount_column] = numeric
    work = work[work[amount_column].notna() & (work[amount_column] > 0)].copy()
    if work.empty:
        return _ensure_columns(work)

    freq = work.groupby(amount_column)[amount_column].transform("size")
    work = work[freq > 2].copy()
    if work.empty:
        return _ensure_columns(work)

    cheque_series = work["Cheque No"].fillna("").astype(str).str.strip()
    work["__has_cheque"] = cheque_series != ""
    work["__cheque_sort"] = cheque_series.apply(_cheque_sort_value)
    work["__date_sort"] = work["Date"].apply(_coerce_excel_date)
    work = work.sort_values(
        by=[amount_column, "__has_cheque", "__cheque_sort", "__date_sort", "Sno"],
        ascending=[False, False, True, True, True],
        kind="stable",
        na_position="last",
    )
    work = work.drop(columns=["__has_cheque", "__cheque_sort", "__date_sort"])
    return _ensure_columns(work)


def _build_top_sheet(statement_df: pd.DataFrame, amount_column: str, top_n: int = 30) -> pd.DataFrame:
    if statement_df.empty:
        return _ensure_columns(statement_df)

    work = statement_df.copy()
    numeric = _to_numeric(work, amount_column)
    work[amount_column] = numeric
    work = work[work[amount_column].notna() & (work[amount_column] > 0)].copy()
    if work.empty:
        return _ensure_columns(work)

    work = work.sort_values(by=[amount_column, "Sno"], ascending=[False, True]).head(top_n)
    return _ensure_columns(work)


def _build_month_dr_cr_sheet(statement_df: pd.DataFrame) -> pd.DataFrame:
    columns = ["Yr-Month", "Dr", "Cr", "Net", "EOM Balance", "#.Of.Dr", "#.Of.Cr", "Avg.Dr", "Avg.Cr"]
    if statement_df.empty:
        return pd.DataFrame(columns=columns)

    work = statement_df.copy()
    work["__month_date"] = work["Date"].apply(_coerce_excel_date)
    work = work[work["__month_date"].notna()].copy()
    if work.empty:
        return pd.DataFrame(columns=columns)

    work["Debit"] = pd.to_numeric(work["Debit"], errors="coerce").fillna(0.0)
    work["Credit"] = pd.to_numeric(work["Credit"], errors="coerce").fillna(0.0)
    work["Balance"] = pd.to_numeric(work["Balance"], errors="coerce")
    # A notice's printed zero is not the account's closing ledger balance.
    notice_mask = [is_nonposting_cheque_return(row) for row in work.to_dict("records")]
    work.loc[notice_mask, "Balance"] = float("nan")
    work["Sno"] = pd.to_numeric(work["Sno"], errors="coerce")
    work["__month_key"] = work["__month_date"].map(lambda value: datetime(value.year, value.month, 1))
    threshold = 30.0

    rows: list[dict[str, Any]] = []
    grouped = work.groupby("__month_key", sort=True)
    for month_key, month_frame in grouped:
        debit_value = round(float(month_frame["Debit"].sum()), 2)
        credit_value = round(float(month_frame["Credit"].sum()), 2)
        month_frame = month_frame.sort_values(by=["__month_date", "Sno"], ascending=[True, True], na_position="last")
        month_end_balance = month_frame["Balance"].dropna()
        debit_over_threshold = month_frame.loc[month_frame["Debit"] > threshold, "Debit"]
        credit_over_threshold = month_frame.loc[month_frame["Credit"] > threshold, "Credit"]
        rows.append(
            {
                "Yr-Month": month_key.strftime("%y-%b"),
                "Dr": debit_value,
                "Cr": credit_value,
                "Net": round(credit_value - debit_value, 2),
                "EOM Balance": round(float(month_end_balance.iloc[-1]), 2) if not month_end_balance.empty else None,
                "#.Of.Dr": int(debit_over_threshold.count()),
                "#.Of.Cr": int(credit_over_threshold.count()),
                "Avg.Dr": round(float(debit_over_threshold.mean()), 2) if not debit_over_threshold.empty else 0.0,
                "Avg.Cr": round(float(credit_over_threshold.mean()), 2) if not credit_over_threshold.empty else 0.0,
            }
        )

    total_debit = round(float(work["Debit"].sum()), 2)
    total_credit = round(float(work["Credit"].sum()), 2)
    total_debit_over_threshold = work.loc[work["Debit"] > threshold, "Debit"]
    total_credit_over_threshold = work.loc[work["Credit"] > threshold, "Credit"]
    rows.append(
        {
            "Yr-Month": "Total",
            "Dr": total_debit,
            "Cr": total_credit,
            "Net": round(total_credit - total_debit, 2),
            "EOM Balance": "",
            "#.Of.Dr": int(total_debit_over_threshold.count()),
            "#.Of.Cr": int(total_credit_over_threshold.count()),
            "Avg.Dr": round(float(total_debit_over_threshold.mean()), 2) if not total_debit_over_threshold.empty else 0.0,
            "Avg.Cr": round(float(total_credit_over_threshold.mean()), 2) if not total_credit_over_threshold.empty else 0.0,
        }
    )
    return pd.DataFrame(rows, columns=columns)
