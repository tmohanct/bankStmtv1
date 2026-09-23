"""Validate extracted transactions before either workbook is written."""
from dataclasses import dataclass
from decimal import Decimal
import math

from src.utils.statement_utils import extract_summary_metrics, normalize_date
from src.transform.cheque_returns import is_nonposting_cheque_return


@dataclass(frozen=True)
class ReconciliationResult:
    status: str
    mismatches: tuple[str, ...] = ()


@dataclass(frozen=True)
class BalanceCheckResult:
    status: str
    direction: str
    compared: int
    missing: int
    mismatches: tuple[str, ...] = ()


def check_running_balances(records) -> BalanceCheckResult:
    """Check adjacent ledger balances without assuming every bank prints them.

    Statements may list transactions newest first. A mismatch is audit evidence,
    not a reason to discard otherwise useful parsed transactions.
    """
    indexed_records = [(index, row) for index, row in enumerate(records, 1)
                       if not is_nonposting_cheque_return(row)]
    records = [row for _, row in indexed_records]
    if len(records) < 2:
        return BalanceCheckResult("unavailable", "unknown", 0, 0)
    dates = [normalize_date(row.get("Date")) for row in records]
    trends = [(later > earlier) - (later < earlier)
              for earlier, later in zip(dates, dates[1:]) if earlier and later]
    direction = "descending" if trends.count(-1) > trends.count(1) else "ascending"
    compared, missing, mismatches = 0, 0, []
    for (older_index, older), (newer_index, newer) in zip(indexed_records, indexed_records[1:]):
        try:
            first = Decimal(str(older.get("Balance")))
            second = Decimal(str(newer.get("Balance")))
            transaction = older if direction == "descending" else newer
            debit = Decimal(str(transaction.get("Debit") or 0))
            credit = Decimal(str(transaction.get("Credit") or 0))
            expected = first + debit - credit if direction == "descending" else first - debit + credit
        except (TypeError, ValueError, ArithmeticError):
            missing += 1
            continue
        compared += 1
        if abs(second - expected) > Decimal("0.01"):
            mismatches.append(f"Rows {older_index}-{newer_index}: expected {expected:.2f}, printed {second:.2f}")
    status = "mismatch" if mismatches else "unavailable" if not compared else "partial" if missing else "passed"
    return BalanceCheckResult(status, direction, compared, missing, tuple(mismatches))


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
        if (row.get("Debit") in (None, "") and row.get("Credit") in (None, "")
                and not is_nonposting_cheque_return(row)):
            raise ValueError(f"{source}: row {index} has no debit or credit amount")


def reconcile(records, pdf_path, logger, summary_extractor=None) -> ReconciliationResult:
    notice_count = sum(is_nonposting_cheque_return(row) for row in records)
    actual = {
        "transaction_count": Decimal(len(records)),
        "total_debit": sum((Decimal(str(row.get("Debit") or 0)) for row in records), Decimal(0)),
        "total_credit": sum((Decimal(str(row.get("Credit") or 0)) for row in records), Decimal(0)),
    }
    logger.info("Parsed transactions summary: %s", actual)
    if notice_count:
        logger.info("Retained %s cheque-return/rejection notice(s) with no debit or credit movement", notice_count)
    summary = extract_summary_metrics(pdf_path, logger)
    if not summary and summary_extractor is not None:
        summary = summary_extractor(pdf_path, logger)
    if not summary:
        logger.info("Reconciliation unavailable: no printed summary totals were found.")
        return ReconciliationResult("unavailable")
    mismatches = []
    for key, expected in summary.items():
        # Some banks omit notices from their Dr/Cr counts; others include them.
        if key == "transaction_count" and notice_count and Decimal(str(expected)) in (
            actual[key], actual[key] - notice_count,
        ):
            continue
        tolerance = Decimal(0) if key == "transaction_count" else Decimal("0.01")
        if abs(actual[key] - Decimal(str(expected))) > tolerance:
            mismatches.append(f"{key}: expected={expected}, parsed={actual[key]}")
    if mismatches:
        for message in mismatches:
            logger.warning("Reconciliation mismatch: %s", message)
        return ReconciliationResult("failed", tuple(mismatches))
    logger.info("Reconciliation passed for all available printed totals.")
    return ReconciliationResult("passed")
