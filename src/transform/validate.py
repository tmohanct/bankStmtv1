"""Validate extracted transactions before either workbook is written."""
from dataclasses import dataclass
from decimal import Decimal
import math

from src.utils.statement_utils import extract_summary_metrics, normalize_date


@dataclass(frozen=True)
class ReconciliationResult:
    status: str
    mismatches: tuple[str, ...] = ()


def validate_records(records, source="statement"):
    if not records:
        raise ValueError(f"No transactions were parsed from {source}. Check the bank, PDF layout and OCR support.")
    for index, row in enumerate(records, 1):
        if normalize_date(row.get("Date")) is None:
            raise ValueError(f"{source}: row {index} has an invalid transaction date")
        for column in ("Debit", "Credit", "Balance"):
            value = row.get(column)
            if value is None or value == "":
                continue
            try:
                valid = math.isfinite(float(value))
            except (ValueError, TypeError):
                valid = False
            if not valid:
                raise ValueError(f"{source}: row {index} has invalid {column}")
        if row.get("Debit") and row.get("Credit"):
            raise ValueError(f"{source}: row {index} has both debit and credit amounts")
        if row.get("Debit") in (None, "") and row.get("Credit") in (None, ""):
            raise ValueError(f"{source}: row {index} has no debit or credit amount")


def reconcile(records, pdf_path, logger) -> ReconciliationResult:
    actual = {
        "transaction_count": Decimal(len(records)),
        "total_debit": sum((Decimal(str(row.get("Debit") or 0)) for row in records), Decimal(0)),
        "total_credit": sum((Decimal(str(row.get("Credit") or 0)) for row in records), Decimal(0)),
    }
    logger.info("Parsed transactions summary: %s", actual)
    summary = extract_summary_metrics(pdf_path, logger)
    if not summary:
        logger.info("Reconciliation unavailable: no printed summary totals were found.")
        return ReconciliationResult("unavailable")
    mismatches = []
    for key, expected in summary.items():
        tolerance = Decimal(0) if key == "transaction_count" else Decimal("0.01")
        if abs(actual[key] - Decimal(str(expected))) > tolerance:
            mismatches.append(f"{key}: expected={expected}, parsed={actual[key]}")
    if mismatches:
        for message in mismatches:
            logger.warning("Reconciliation mismatch: %s", message)
        return ReconciliationResult("failed", tuple(mismatches))
    logger.info("Reconciliation passed for all available printed totals.")
    return ReconciliationResult("passed")
