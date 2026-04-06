import os
import re
import subprocess
import sys
import tempfile
import textwrap
import zipfile
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from datetime import date, datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import pandas as pd
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from PIL import Image as PILImage, ImageDraw, ImageFont

from utils import OUTPUT_COLUMNS, clean_detail, compact_detail_key, sanitize_cheque_column

FONT_NORMAL = Font(name="Aptos", size=10)
FONT_HEADER = Font(name="Aptos", size=10, bold=True)
FONT_FOOTNOTE = Font(name="Aptos", size=10, italic=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)
HEADER_FILL = PatternFill(fill_type="solid", fgColor="D9E1F2")
ALT_ROW_FILLS = [
    PatternFill(fill_type="solid", fgColor="EAF3FB"),
    PatternFill(fill_type="solid", fgColor="F8F1E5"),
]
MONTH_LABEL_FILL = PatternFill(fill_type="solid", fgColor="C9D5EA")
MONTH_VALUE_ROW_FILLS = [
    PatternFill(fill_type="solid", fgColor="F2F2F2"),
    PatternFill(fill_type="solid", fgColor="FFFFFF"),
]
REPEAT_GROUP_FILLS = [
    PatternFill(fill_type="solid", fgColor="FCE4D6"),
    PatternFill(fill_type="solid", fgColor="D9EAD3"),
    PatternFill(fill_type="solid", fgColor="D9E1F2"),
]
INDIAN_NUMBER_FORMAT = "#,##,##0.00"
INDIAN_NUMBER_FORMAT_NO_DECIMAL = "#,##,##0"
DATE_NUMBER_FORMAT = "yyyy-mm-dd"
DATE_INPUT_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y")
AMOUNT_COLUMN_WIDTH = 16
THIN_BORDER = Border(
    left=Side(style="thin", color="BFBFBF"),
    right=Side(style="thin", color="BFBFBF"),
    top=Side(style="thin", color="BFBFBF"),
    bottom=Side(style="thin", color="BFBFBF"),
)
MONTH_DR_CR_FOOTNOTE = "#.OF Dr/Cr & Avg takes only amount Greater than 30. Less than 30 not counted."
MONTH_DR_CR_CHART_IMAGE_SIZE = (1120, 520)
MONTH_DR_CR_DATA_LABEL_FONT_SIZE = 13
MONTH_DR_CR_EXCEL_DATA_LABEL_FONT_SIZE = 12
C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DIRECT_RETURN_REJECT_MARKERS = (
    "WCHQRET",
    "CHQRETURN",
    "CHQRET",
    "CHEQUERETURN",
    "CHEQUERET",
    "IWCHQRETURN",
    "IWCHQRET",
    "BRNOWRTNCLG",
    "RETURNED",
    "REJECT",
    "REJECTED",
    "IWREJINST",
    "DISHONOUR",
    "DISHONOR",
)
RETURN_RELATED_CHARGE_MARKERS = (
    "RETURNCHARGE",
    "RETURNCHARGES",
    "RETURNCHG",
    "RETURNCHGS",
    "CHQRETURNCHG",
    "CHQRETURNCHGS",
    "CHQRTNCHRG",
    "CHQRTNCHRGS",
    "RTNCHQCHGS",
    "ACHRTNCHRG",
    "RTNCHG",
    "RTNCHRGS",
)
ELECTRONIC_RETURN_MARKERS = ("NEFT", "RTGS", "IMPS")


def _sanitize_sheet_name(name: str) -> str:
    safe = re.sub(r"[\\/*?:\[\]]", "_", str(name).strip())
    safe = safe or "Sheet"
    return safe[:31]


def _unique_sheet_name(name: str, used_names: set[str]) -> str:
    base = _sanitize_sheet_name(name)
    candidate = base
    index = 1
    while candidate in used_names:
        suffix = f"_{index}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        index += 1
    used_names.add(candidate)
    return candidate


def _normalize_header(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").strip().lower())


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


def _round_money_for_excel(value: Any) -> int | None:
    if value in (None, ""):
        return None

    try:
        if pd.isna(value):
            return None
    except TypeError:
        pass

    try:
        rounded = Decimal(str(value)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError):
        return None

    return int(rounded)


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


def _is_return_reject_detail(value: Any) -> bool:
    normalized = compact_detail_key(value).upper()
    if not normalized:
        return False

    if any(marker in normalized for marker in DIRECT_RETURN_REJECT_MARKERS):
        return True
    if any(marker in normalized for marker in RETURN_RELATED_CHARGE_MARKERS):
        return True
    if "RETURN" in normalized and any(marker in normalized for marker in ELECTRONIC_RETURN_MARKERS):
        return True
    return False


def _build_return_reject_sheet(statement_df: pd.DataFrame) -> pd.DataFrame:
    if statement_df.empty:
        return _ensure_columns(statement_df)

    work = statement_df.copy()
    detail_series = work["Details"] if "Details" in work.columns else pd.Series([""] * len(work), index=work.index)
    mask = detail_series.fillna("").astype(str).map(_is_return_reject_detail)
    work = work[mask].copy()
    return _ensure_columns(work)


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

    work = work.sort_values(by=[amount_column, "Sno"], ascending=[False, True])
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


def _apply_base_style(workbook) -> None:
    left_headers = {"date", "detail", "details", "detailclean", "cheque", "chequeno", "source"}
    center_headers = {"sno"}
    right_headers = {"debit", "credit", "balance"}
    date_headers = {"date", "txndate", "valuedate"}
    text_headers = {"cheque", "chequeno"}

    for ws in workbook.worksheets:
        if ws.title.lower() == "month_dr_cr":
            continue

        max_row = ws.max_row
        max_col = ws.max_column

        if max_row < 1 or max_col < 1:
            continue

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        header_to_col: dict[str, int] = {}
        for col_idx in range(1, max_col + 1):
            header_value = ws.cell(row=1, column=col_idx).value
            header_to_col[_normalize_header(header_value)] = col_idx

        align_by_col: dict[int, Alignment] = {}
        for normalized_header, col_idx in header_to_col.items():
            if normalized_header in left_headers:
                align_by_col[col_idx] = ALIGN_LEFT
            elif normalized_header in right_headers:
                align_by_col[col_idx] = ALIGN_RIGHT
            elif normalized_header in center_headers:
                align_by_col[col_idx] = ALIGN_CENTER
            else:
                align_by_col[col_idx] = ALIGN_CENTER

        numeric_cols = {
            col_idx
            for normalized_header, col_idx in header_to_col.items()
            if normalized_header in right_headers
        }
        date_cols = {
            col_idx
            for normalized_header, col_idx in header_to_col.items()
            if normalized_header in date_headers
        }
        text_cols = {
            col_idx
            for normalized_header, col_idx in header_to_col.items()
            if normalized_header in text_headers
        }

        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = FONT_HEADER
            cell.alignment = ALIGN_CENTER
            cell.fill = HEADER_FILL

        for row_idx in range(2, max_row + 1):
            fill = ALT_ROW_FILLS[(row_idx - 2) % 2]
            for col_idx in range(1, max_col + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.font = FONT_NORMAL
                cell.alignment = align_by_col.get(col_idx, ALIGN_CENTER)
                if col_idx in numeric_cols:
                    rounded_value = _round_money_for_excel(cell.value)
                    if rounded_value is not None:
                        cell.value = rounded_value
                    cell.number_format = INDIAN_NUMBER_FORMAT_NO_DECIMAL
                if col_idx in date_cols:
                    parsed_date = _coerce_excel_date(cell.value)
                    if parsed_date is not None:
                        cell.value = parsed_date
                    cell.number_format = DATE_NUMBER_FORMAT
                if col_idx in text_cols and cell.value not in (None, ""):
                    cell.value = str(cell.value)
                    cell.number_format = "@"
                cell.fill = fill

        for col in ws.columns:
            max_len = 0
            col_letter = col[0].column_letter
            for cell in col[: min(len(col), 400)]:
                value = "" if cell.value is None else str(cell.value)
                if len(value) > max_len:
                    max_len = len(value)
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 48)

        for col_idx in numeric_cols:
            col_letter = ws.cell(row=1, column=col_idx).column_letter
            ws.column_dimensions[col_letter].width = AMOUNT_COLUMN_WIDTH


def _apply_month_dr_cr_style(workbook, sheet_name: str) -> None:
    if sheet_name not in workbook.sheetnames:
        return

    ws = workbook[sheet_name]
    if ws.max_row < 1 or ws.max_column < 1:
        return

    data_max_row = ws.max_row
    data_max_col = ws.max_column
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.cell(row=1, column=1).coordinate + ":" + ws.cell(row=data_max_row, column=data_max_col).coordinate

    for row_idx in range(1, data_max_row + 1):
        row_label = str(ws.cell(row=row_idx, column=1).value or "").strip()
        is_total_row = row_label.lower() == "total"
        for col_idx in range(1, data_max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = FONT_NORMAL
            cell.border = THIN_BORDER

            if row_idx == 1:
                cell.font = FONT_HEADER
                cell.alignment = ALIGN_LEFT if col_idx == 1 else ALIGN_CENTER
            elif col_idx == 1:
                cell.font = FONT_HEADER
                cell.alignment = ALIGN_LEFT
                cell.fill = MONTH_LABEL_FILL
            else:
                header_value = str(ws.cell(row=1, column=col_idx).value or "").strip()
                if is_total_row:
                    cell.font = FONT_HEADER
                cell.alignment = ALIGN_RIGHT
                if isinstance(cell.value, (int, float)):
                    rounded_value = _round_money_for_excel(cell.value)
                    if rounded_value is not None:
                        cell.value = rounded_value
                    cell.number_format = INDIAN_NUMBER_FORMAT_NO_DECIMAL
                cell.fill = MONTH_VALUE_ROW_FILLS[(row_idx - 2) % len(MONTH_VALUE_ROW_FILLS)]

    ws.column_dimensions["A"].width = 14
    width_map = {
        "Dr": 14,
        "Cr": 14,
        "Net": 14,
        "EOM Balance": 16,
        "#.Of.Dr": 10,
        "#.Of.Cr": 10,
        "Avg.Dr": 14,
        "Avg.Cr": 14,
    }
    for col_idx in range(2, data_max_col + 1):
        header_value = str(ws.cell(row=1, column=col_idx).value or "").strip()
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = width_map.get(header_value, 14)

    footnote_row = data_max_row + 2
    ws.merge_cells(start_row=footnote_row, start_column=1, end_row=footnote_row, end_column=data_max_col)
    footnote_cell = ws.cell(row=footnote_row, column=1)
    footnote_cell.value = MONTH_DR_CR_FOOTNOTE
    footnote_cell.font = FONT_FOOTNOTE
    footnote_cell.alignment = ALIGN_LEFT

    chart_data_end_row = data_max_row
    if chart_data_end_row >= 2:
        last_label = str(ws.cell(row=chart_data_end_row, column=1).value or "").strip().lower()
        if last_label == "total":
            chart_data_end_row -= 1

    if chart_data_end_row >= 2:
        _add_month_dr_cr_chart_image(ws, chart_data_end_row, footnote_row)


def _format_month_dr_cr_chart_label(value: Any) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""

    try:
        numeric = float(value)
    except (TypeError, ValueError):
        return ""

    absolute = abs(numeric)
    if absolute >= 10000000:
        return f"{numeric / 10000000:.2f} Cr"
    if absolute >= 100000:
        return f"{numeric / 100000:.1f} L"
    if absolute >= 1000:
        return f"{numeric / 1000:.1f} k"
    return f"{numeric:.1f}"


def _load_chart_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    font_names = ["arialbd.ttf", "calibrib.ttf", "arial.ttf", "calibri.ttf"] if bold else [
        "arial.ttf",
        "calibri.ttf",
    ]
    for font_name in font_names:
        try:
            return ImageFont.truetype(font_name, size=size)
        except OSError:
            continue
    return ImageFont.load_default()


def _draw_dotted_horizontal_line(draw: ImageDraw.ImageDraw, x1: int, x2: int, y: int, fill: str) -> None:
    dash_length = 4
    gap_length = 6
    x = x1
    while x < x2:
        draw.line((x, y, min(x + dash_length, x2), y), fill=fill, width=1)
        x += dash_length + gap_length


def _format_month_dr_cr_axis_label(value: float) -> str:
    return f"\u20B9{_format_month_dr_cr_chart_label(value)}"


def _nice_axis_step(max_value: float, tick_count: int = 8) -> float:
    if max_value <= 0:
        return 1.0

    rough_step = max_value / max(tick_count, 1)
    exponent = int(f"{rough_step:e}".split("e")[1])
    magnitude = 10 ** exponent
    fraction = rough_step / magnitude

    if fraction <= 1:
        nice_fraction = 1
    elif fraction <= 2:
        nice_fraction = 2
    elif fraction <= 5:
        nice_fraction = 5
    else:
        nice_fraction = 10

    return nice_fraction * magnitude


def _draw_rotated_text(
    image: PILImage.Image,
    text: str,
    position: tuple[int, int],
    font,
    fill: str,
    angle: float,
) -> None:
    if not text:
        return

    measure_draw = ImageDraw.Draw(PILImage.new("RGBA", (1, 1), (0, 0, 0, 0)))
    bbox = measure_draw.textbbox((0, 0), text, font=font)
    text_width = max(1, bbox[2] - bbox[0])
    text_height = max(1, bbox[3] - bbox[1])
    text_layer = PILImage.new("RGBA", (text_width + 8, text_height + 8), (0, 0, 0, 0))
    text_draw = ImageDraw.Draw(text_layer)
    text_draw.text((4, 4), text, font=font, fill=fill)
    rotated = text_layer.rotate(angle, expand=True, resample=PILImage.Resampling.BICUBIC)
    image.alpha_composite(rotated, dest=position)


def _add_month_dr_cr_chart_image(ws, chart_data_end_row: int, footnote_row: int) -> None:
    month_values: list[tuple[str, float, float]] = []
    for row_idx in range(2, chart_data_end_row + 1):
        month_label = str(ws.cell(row=row_idx, column=1).value or "").strip()
        if not month_label:
            continue
        try:
            debit = float(ws.cell(row=row_idx, column=2).value or 0)
            credit = float(ws.cell(row=row_idx, column=3).value or 0)
        except (TypeError, ValueError):
            continue
        month_values.append((month_label, debit, credit))

    if not month_values:
        return

    width, height = MONTH_DR_CR_CHART_IMAGE_SIZE
    image = PILImage.new("RGBA", (width, height), (255, 255, 255, 0))
    draw = ImageDraw.Draw(image)

    panel_margin = 12
    panel_radius = 26
    shadow_offset = 8
    shadow_color = (0, 0, 0, 48)
    panel_fill = "#272624"
    panel_border = "#4A4744"
    grid_color = "#403D39"
    axis_text_color = "#8F8B86"
    label_text_color = "#BFBAB3"
    value_text_color = "#D9D5CF"
    bar_colors = {"Dr": "#355C91", "Cr": "#B0572D"}

    draw.rounded_rectangle(
        (
            panel_margin + shadow_offset,
            panel_margin + shadow_offset,
            width - panel_margin + shadow_offset,
            height - panel_margin + shadow_offset,
        ),
        radius=panel_radius,
        fill=shadow_color,
    )
    draw.rounded_rectangle(
        (panel_margin, panel_margin, width - panel_margin, height - panel_margin),
        radius=panel_radius,
        fill=panel_fill,
        outline=panel_border,
        width=2,
    )

    font_regular = _load_chart_font(12)
    font_small = _load_chart_font(10)
    font_label = _load_chart_font(MONTH_DR_CR_DATA_LABEL_FONT_SIZE)
    font_bold = _load_chart_font(12, bold=True)

    left_margin = panel_margin + 78
    right_margin = panel_margin + 34
    top_margin = panel_margin + 92
    bottom_margin = panel_margin + 88
    plot_left = left_margin
    plot_top = top_margin
    plot_right = width - right_margin
    plot_bottom = height - bottom_margin
    plot_width = plot_right - plot_left
    plot_height = plot_bottom - plot_top

    max_value = max(max(debit, credit) for _, debit, credit in month_values)
    if max_value <= 0:
        max_value = 1.0
    axis_step = _nice_axis_step(max_value, tick_count=8)
    axis_max = axis_step * max(1, int((max_value + axis_step - 1) // axis_step))

    def value_to_y(value: float) -> int:
        scaled = value / axis_max
        return round(plot_bottom - (plot_height * scaled))

    tick_value = 0.0
    while tick_value <= axis_max + (axis_step / 2):
        y = value_to_y(tick_value)
        _draw_dotted_horizontal_line(draw, plot_left, plot_right, y, grid_color)
        tick_value += axis_step

    group_width = plot_width / max(len(month_values), 1)
    bar_width = max(16, min(34, int(group_width * 0.30)))
    series_gap = max(8, int(bar_width * 0.28))

    def draw_centered_text(text: str, center_x: int, top_y: int, font, fill: str) -> None:
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        draw.text((center_x - text_width / 2, top_y), text, font=font, fill=fill)

    label_offsets = {"Dr": -12, "Cr": 10}

    for index, (month_label, debit, credit) in enumerate(month_values):
        group_center = plot_left + group_width * (index + 0.5)

        dr_left = round(group_center - series_gap / 2 - bar_width)
        dr_right = dr_left + bar_width
        dr_top = value_to_y(debit)
        draw.rounded_rectangle((dr_left, dr_top, dr_right, plot_bottom), radius=7, fill=bar_colors["Dr"])

        cr_left = round(group_center + series_gap / 2)
        cr_right = cr_left + bar_width
        cr_top = value_to_y(credit)
        draw.rounded_rectangle((cr_left, cr_top, cr_right, plot_bottom), radius=7, fill=bar_colors["Cr"])

        dr_label = _format_month_dr_cr_chart_label(debit)
        dr_bbox = draw.textbbox((0, 0), dr_label, font=font_label)
        _draw_rotated_text(
            image,
            dr_label,
            (
                int((dr_left + dr_right) // 2 + label_offsets["Dr"] - ((dr_bbox[2] - dr_bbox[0]) * 0.14)),
                int(max(plot_top - 8, dr_top - 42)),
            ),
            font_label,
            label_text_color,
            64,
        )

        cr_label = _format_month_dr_cr_chart_label(credit)
        cr_bbox = draw.textbbox((0, 0), cr_label, font=font_label)
        _draw_rotated_text(
            image,
            cr_label,
            (
                int((cr_left + cr_right) // 2 + label_offsets["Cr"] - ((cr_bbox[2] - cr_bbox[0]) * 0.12)),
                int(max(plot_top - 8, cr_top - 42)),
            ),
            font_label,
            label_text_color,
            64,
        )

        month_bbox = draw.textbbox((0, 0), month_label, font=font_regular)
        _draw_rotated_text(
            image,
            month_label,
            (
                int(round(group_center) - ((month_bbox[2] - month_bbox[0]) * 0.55)),
                int(plot_bottom + 4),
            ),
            font_regular,
            axis_text_color,
            45,
        )

    legend_x = panel_margin + 28
    legend_y = panel_margin + 28
    legend_cursor = legend_x
    for legend_label, legend_color in (("Debit (Dr)", bar_colors["Dr"]), ("Credit (Cr)", bar_colors["Cr"])):
        draw.rounded_rectangle((legend_cursor, legend_y + 4, legend_cursor + 16, legend_y + 20), radius=4, fill=legend_color)
        draw.text((legend_cursor + 22, legend_y), legend_label, font=font_bold, fill=value_text_color)
        label_bbox = draw.textbbox((0, 0), legend_label, font=font_bold)
        legend_cursor += 22 + (label_bbox[2] - label_bbox[0]) + 28

    image_bytes = BytesIO()
    image.save(image_bytes, format="PNG")
    image_bytes.seek(0)

    chart_image = XLImage(image_bytes)
    chart_image.width = width
    chart_image.height = height
    chart_row = footnote_row + 5
    ws.add_image(chart_image, f"A{chart_row}")


def _patch_month_dr_cr_chart_xml(final_path: Path, sheet_name: str, month_labels: list[str], logger) -> None:
    if not month_labels or not final_path.is_file():
        return

    ET.register_namespace("", C_NS)
    ET.register_namespace("a", A_NS)
    ET.register_namespace("r", R_NS)

    formula_sheet_name = sheet_name.replace("'", "''")
    category_formula = f"'{formula_sheet_name}'!$A$2:$A${len(month_labels) + 1}"
    namespaces = {"c": C_NS}

    def build_str_ref(parent: ET.Element) -> None:
        str_ref = ET.SubElement(parent, f"{{{C_NS}}}strRef")
        formula = ET.SubElement(str_ref, f"{{{C_NS}}}f")
        formula.text = category_formula

        cache = ET.SubElement(str_ref, f"{{{C_NS}}}strCache")
        point_count = ET.SubElement(cache, f"{{{C_NS}}}ptCount")
        point_count.set("val", str(len(month_labels)))
        for idx, label in enumerate(month_labels):
            point = ET.SubElement(cache, f"{{{C_NS}}}pt")
            point.set("idx", str(idx))
            value = ET.SubElement(point, f"{{{C_NS}}}v")
            value.text = str(label)

    temp_output: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as temp_file:
            temp_output = Path(temp_file.name)

        with zipfile.ZipFile(final_path, "r") as source_zip, zipfile.ZipFile(temp_output, "w") as target_zip:
            for entry in source_zip.infolist():
                content = source_zip.read(entry.filename)
                if entry.filename.startswith("xl/charts/chart") and entry.filename.endswith(".xml"):
                    root = ET.fromstring(content)
                    changed = False

                    chart_node = root.find("c:chart", namespaces)
                    if chart_node is not None:
                        chart_title = chart_node.find("c:title", namespaces)
                        if chart_title is not None:
                            chart_node.remove(chart_title)
                            changed = True

                    for ser in root.findall(".//c:barChart/c:ser", namespaces):
                        cat = ser.find("c:cat", namespaces)
                        if cat is None:
                            continue
                        for child in list(cat):
                            cat.remove(child)
                        build_str_ref(cat)
                        changed = True

                    cat_axis = root.find(".//c:catAx", namespaces)
                    if cat_axis is not None:
                        delete_node = cat_axis.find("c:delete", namespaces)
                        if delete_node is None:
                            delete_node = ET.SubElement(cat_axis, f"{{{C_NS}}}delete")
                        delete_node.set("val", "0")

                        ax_pos = cat_axis.find("c:axPos", namespaces)
                        if ax_pos is not None:
                            ax_pos.set("val", "b")

                        tick_label_pos = cat_axis.find("c:tickLblPos", namespaces)
                        if tick_label_pos is None:
                            tick_label_pos = ET.SubElement(cat_axis, f"{{{C_NS}}}tickLblPos")
                        tick_label_pos.set("val", "low")
                        changed = True

                    val_axis = root.find(".//c:valAx", namespaces)
                    if val_axis is not None:
                        delete_node = val_axis.find("c:delete", namespaces)
                        if delete_node is None:
                            delete_node = ET.SubElement(val_axis, f"{{{C_NS}}}delete")
                        delete_node.set("val", "0")

                        ax_pos = val_axis.find("c:axPos", namespaces)
                        if ax_pos is not None:
                            ax_pos.set("val", "l")
                        tick_label_pos = val_axis.find("c:tickLblPos", namespaces)
                        if tick_label_pos is None:
                            tick_label_pos = ET.SubElement(val_axis, f"{{{C_NS}}}tickLblPos")
                        tick_label_pos.set("val", "none")
                        axis_title = val_axis.find("c:title", namespaces)
                        if axis_title is not None:
                            val_axis.remove(axis_title)
                        changed = True

                    legend = root.find(".//c:legend", namespaces)
                    if legend is not None:
                        legend_pos = legend.find("c:legendPos", namespaces)
                        if legend_pos is None:
                            legend_pos = ET.SubElement(legend, f"{{{C_NS}}}legendPos")
                        legend_pos.set("val", "r")
                        changed = True

                    if changed:
                        content = ET.tostring(root, encoding="utf-8", xml_declaration=False)

                target_zip.writestr(entry, content)

        temp_output.replace(final_path)
        logger.info("Patched month_dr_cr chart XML with explicit month categories")
    except Exception as exc:  # noqa: BLE001
        logger.warning("Unable to patch month_dr_cr chart XML. Details: %s", exc)
        if temp_output is not None and temp_output.exists():
            try:
                temp_output.unlink()
            except OSError:
                pass


def _try_apply_excel_chart_postprocess(final_path: Path, sheet_name: str, logger) -> None:
    if not sys.platform.startswith("win"):
        return

    system_root = Path(os.environ.get("SystemRoot", r"C:\Windows"))
    powershell_exe = system_root / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
    if not powershell_exe.is_file():
        logger.warning("Skipping Excel chart post-process because PowerShell was not found: %s", powershell_exe)
        return

    script_text = textwrap.dedent(
        r"""
        param(
            [Parameter(Mandatory = $true)][string]$WorkbookPath,
            [Parameter(Mandatory = $true)][string]$SheetName,
            [Parameter(Mandatory = $true)][string]$FootnoteText
        )

        Set-StrictMode -Version Latest
        $ErrorActionPreference = 'Stop'

        function Get-CompactLabel([double]$Value) {
            $absolute = [Math]::Abs($Value)
            if ($absolute -ge 10000000) {
                return ('{0:0.00} Cr' -f ($Value / 10000000.0))
            }
            if ($absolute -ge 100000) {
                return ('{0:0.0} L' -f ($Value / 100000.0))
            }
            if ($absolute -ge 1000) {
                return ('{0:0.0} k' -f ($Value / 1000.0))
            }
            return ('{0:0.0}' -f $Value)
        }

        function Find-RowByValue($Worksheet, [string]$Text, [int]$MaxRow) {
            for ($row = 1; $row -le $MaxRow; $row++) {
                if ([string]$Worksheet.Cells.Item($row, 1).Value2 -eq $Text) {
                    return $row
                }
            }
            return 0
        }

        $xlCategory = 1
        $xlValue = 2
        $xlColumns = 2
        $xlColumnClustered = 51
        $xlLegendPositionRight = -4152
        $xlTickLabelPositionLow = -4134
        $xlTickLabelPositionNone = -4142
        $xlLabelPositionOutsideEnd = 2
        $msoLineRoundDot = 3

        $excel = $null
        $workbook = $null
        $saveChanges = $false
        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            $workbook = $excel.Workbooks.Open($WorkbookPath)
            $worksheet = $workbook.Worksheets.Item($SheetName)

            $usedRows = $worksheet.UsedRange.Rows.Count
            $footnoteRow = Find-RowByValue $worksheet $FootnoteText $usedRows
            if ($footnoteRow -le 0) {
                throw "Unable to locate the month_dr_cr footnote row."
            }

            $chartDataEndRow = $footnoteRow - 2
            if ($chartDataEndRow -lt 2) {
                throw "month_dr_cr sheet does not contain chart data rows."
            }

            if ([string]$worksheet.Cells.Item($chartDataEndRow, 1).Value2 -eq 'Total') {
                $chartDataEndRow -= 1
            }
            if ($chartDataEndRow -lt 2) {
                throw "month_dr_cr chart has no month rows after excluding Total."
            }

            if ($worksheet.ChartObjects().Count -ge 1) {
                $chartObject = $worksheet.ChartObjects(1)
                $existingChart = $true
            } else {
                $existingChart = $false
                $chartRow = $footnoteRow + 5
                $left = $worksheet.Range("A$chartRow").Left
                $top = $worksheet.Range("A$chartRow").Top
                $width = $worksheet.Range("A1:I1").Width
                if ($width -lt 850) {
                    $width = 850
                }
                $height = 380
                $chartObject = $worksheet.ChartObjects().Add($left, $top, $width, $height)
            }

            $chart = $chartObject.Chart
            $chart.ChartType = $xlColumnClustered
            $chart.HasTitle = $false
            $chart.HasLegend = $true
            $chart.Legend.Position = $xlLegendPositionRight
            $chart.Legend.IncludeInLayout = $false
            $chart.HasAxis($xlCategory, 1) = $true
            $chart.HasAxis($xlValue, 1) = $true

            if (-not $existingChart) {
                $sourceRange = $worksheet.Range("A1:C$chartDataEndRow")
                $chart.SetSourceData($sourceRange, $xlColumns)
            }
            $worksheet.Activate() | Out-Null
            $chartObject.Activate()
            $chart = $excel.ActiveChart

            $categoryRange = $worksheet.Range("A2:A$chartDataEndRow")
            $seriesSpecs = @(
                @{ Index = 1; Name = 'Debit (Dr)'; Column = 2; Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(79, 129, 189)) },
                @{ Index = 2; Name = 'Credit (Cr)'; Column = 3; Color = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(237, 125, 49)) }
            )
            foreach ($spec in $seriesSpecs) {
                if ($chart.SeriesCollection().Count -lt $spec.Index) {
                    continue
                }
                $series = $chart.SeriesCollection($spec.Index)
                $series.Name = $spec.Name
                $series.XValues = $categoryRange
                $series.Format.Fill.Visible = $true
                $series.Format.Fill.Solid()
                $series.Format.Fill.ForeColor.RGB = $spec.Color
                $series.Format.Line.Visible = $false
                $series.ApplyDataLabels()

                for ($pointIndex = 1; $pointIndex -le $series.Points().Count; $pointIndex++) {
                    $point = $series.Points($pointIndex)
                    $point.HasDataLabel = $true

                    $value = [double]$worksheet.Cells.Item($pointIndex + 1, $spec.Column).Value2
                    $label = $point.DataLabel
                    $label.ShowValue = $false
                    $label.ShowSeriesName = $false
                    $label.ShowCategoryName = $false
                    $label.AutoText = $false
                    $label.Caption = Get-CompactLabel $value
                    $label.Position = $xlLabelPositionOutsideEnd
                    $label.Font.Size = __MONTH_DR_CR_EXCEL_DATA_LABEL_FONT_SIZE__
                }
            }

            $categoryAxis = $chart.Axes($xlCategory)
            $categoryAxis.TickLabelPosition = $xlTickLabelPositionLow
            $categoryAxis.TickLabelSpacing = 1
            $categoryAxis.TickMarkSpacing = 1
            $categoryAxis.HasTitle = $false

            $valueAxis = $chart.Axes($xlValue)
            $valueAxis.HasTitle = $false
            $valueAxis.TickLabelPosition = $xlTickLabelPositionNone
            $valueAxis.HasMajorGridlines = $true
            $valueAxis.MajorGridlines.Format.Line.Visible = $true
            $valueAxis.MajorGridlines.Format.Line.ForeColor.RGB = [System.Drawing.ColorTranslator]::ToOle([System.Drawing.Color]::FromArgb(217, 217, 217))
            $valueAxis.MajorGridlines.Format.Line.DashStyle = $msoLineRoundDot
            $valueAxis.MajorGridlines.Format.Line.Weight = 0.75

            $chart.Legend.Top = 8
            $chart.Legend.Left = $chart.ChartArea.Width - $chart.Legend.Width - 12
            $chart.PlotArea.InsideTop = 18
            $chart.PlotArea.InsideHeight = 265

            $workbook.Save()
            $saveChanges = $true
        }
        finally {
            if ($workbook -ne $null) {
                $workbook.Close($saveChanges)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
            }
            if ($excel -ne $null) {
                $excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
            }
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
        }
        """
    ).replace(
        "__MONTH_DR_CR_EXCEL_DATA_LABEL_FONT_SIZE__",
        str(MONTH_DR_CR_EXCEL_DATA_LABEL_FONT_SIZE),
    ).strip()

    temp_script_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile("w", suffix=".ps1", delete=False, encoding="utf-8") as temp_script:
            temp_script.write(script_text)
            temp_script_path = Path(temp_script.name)

        completed = subprocess.run(
            [
                str(powershell_exe),
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                str(temp_script_path),
                str(final_path),
                sheet_name,
                MONTH_DR_CR_FOOTNOTE,
            ],
            check=False,
            capture_output=True,
            text=True,
        )
        if completed.returncode != 0:
            stderr = completed.stderr.strip() or completed.stdout.strip()
            logger.warning("Excel chart post-process failed; keeping openpyxl chart. Details: %s", stderr)
        else:
            logger.info("Applied Excel chart post-process for %s", sheet_name)
    except OSError as exc:
        logger.warning("Skipping Excel chart post-process because PowerShell could not be started: %s", exc)
    finally:
        if temp_script_path is not None:
            try:
                temp_script_path.unlink()
            except OSError:
                pass


def _apply_repeat_group_colors(workbook, sheet_name: str, amount_column: str) -> None:
    if sheet_name not in workbook.sheetnames:
        return

    ws = workbook[sheet_name]
    if ws.max_row <= 1:
        return

    header_to_col: dict[str, int] = {}
    for col_idx in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col_idx).value
        header_to_col[str(value).strip()] = col_idx

    amount_col_idx = header_to_col.get(amount_column)
    if amount_col_idx is None:
        return

    color_map: dict[str, PatternFill] = {}
    color_index = 0

    for row_idx in range(2, ws.max_row + 1):
        raw_value = ws.cell(row=row_idx, column=amount_col_idx).value
        if raw_value in (None, ""):
            continue

        rounded_value = _round_money_for_excel(raw_value)
        if rounded_value is not None:
            key = str(rounded_value)
        else:
            key = str(raw_value)

        if key not in color_map:
            color_map[key] = REPEAT_GROUP_FILLS[color_index % len(REPEAT_GROUP_FILLS)]
            color_index += 1

        fill = color_map[key]
        for col_idx in range(1, ws.max_column + 1):
            ws.cell(row=row_idx, column=col_idx).fill = fill


def _force_leading_equals_to_text(workbook) -> None:
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.data_type = "s"


def _next_final_path(output_dir: Path, pdf_stem: str) -> Path:
    target = output_dir / f"{pdf_stem}.xlsx"
    if target.exists():
        timestamp = datetime.now().strftime("%y%m%d_%H%M%S")
        return output_dir / f"{pdf_stem}_{timestamp}.xlsx"
    return target


def build_final_workbook(
    statement_df: pd.DataFrame,
    rules_path: Path,
    output_dir: Path,
    pdf_stem: str,
    logger,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    statement_df = _ensure_columns(statement_df)

    rules = _load_rules(rules_path, logger)
    rule_sheets = _build_rule_sheets(statement_df, rules, logger)

    return_reject_df = _build_return_reject_sheet(statement_df)
    cheque_df = _build_cheque_sheet(statement_df)
    repeat_credit_df = _build_repeat_sheet(statement_df, "Credit")
    repeat_debit_df = _build_repeat_sheet(statement_df, "Debit")
    top30_debit_df = _build_top_sheet(statement_df, "Debit", top_n=30)
    top30_credit_df = _build_top_sheet(statement_df, "Credit", top_n=30)
    month_dr_cr_df = _build_month_dr_cr_sheet(statement_df)

    planned_sheets: list[tuple[str, pd.DataFrame]] = [
        ("Statement", statement_df),
        ("Ret/Rej", return_reject_df),
    ]
    planned_sheets.extend(rule_sheets)
    planned_sheets.extend(
        [
            ("Cheque_Transactions", cheque_df),
            ("Repeat_Credit_Amount", repeat_credit_df),
            ("Repeat_Debit_Amount", repeat_debit_df),
            ("Top30_Debit", top30_debit_df),
            ("Top30_Credit", top30_credit_df),
            ("month_dr_cr", month_dr_cr_df),
        ]
    )

    final_path = _next_final_path(output_dir, pdf_stem)

    used_names: set[str] = set()
    normalized_sheet_names: dict[str, str] = {}

    with pd.ExcelWriter(final_path, engine="openpyxl") as writer:
        for requested_name, frame in planned_sheets:
            safe_name = _unique_sheet_name(requested_name, used_names)
            normalized_sheet_names[requested_name] = safe_name
            if requested_name == "month_dr_cr":
                frame.to_excel(writer, sheet_name=safe_name, index=False)
            else:
                _ensure_columns(frame).to_excel(writer, sheet_name=safe_name, index=False)

    workbook = load_workbook(final_path)
    _apply_base_style(workbook)
    _apply_repeat_group_colors(
        workbook,
        normalized_sheet_names.get("Repeat_Credit_Amount", "Repeat_Credit_Amount"),
        "Credit",
    )
    _apply_repeat_group_colors(
        workbook,
        normalized_sheet_names.get("Repeat_Debit_Amount", "Repeat_Debit_Amount"),
        "Debit",
    )
    _apply_month_dr_cr_style(
        workbook,
        normalized_sheet_names.get("month_dr_cr", "month_dr_cr"),
    )
    _force_leading_equals_to_text(workbook)
    workbook.save(final_path)

    logger.info("Final workbook created: %s", final_path)
    return final_path
