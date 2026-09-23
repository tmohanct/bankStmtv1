"""Classify cheque return transactions, excluding fees and electronic returns.

These are narration heuristics, not a bank return-code catalogue. Cheque or
clearing context is helpful but a missing or invalid cheque number must not
hide a return/rejection transaction.
"""
from __future__ import annotations

import re
from typing import Any


def _token(pattern: str) -> re.Pattern[str]:
    # Digits can touch a bank abbreviation; other letters cannot. In particular,
    # RET must not match RETRIEVAL, and REJECT must not match REJECTIONLESS.
    return re.compile(r"(?<![A-Z])(?:" + pattern + r")(?![A-Z])")


CHEQUE = r"CHQ|CHEQUE|CHECK"
CONTEXT_RE = _token(rf"{CHEQUE}|CLG|CLEARING|CTS")
EVENT = (
    r"RETURN(?:ED)?|REJECT(?:ED|ION)?|DISHONOU?R(?:ED)?|BOUNC(?:E|ED)|UNPAID"
    r"|RTND|RETN|RETD|RTN|RET|REJ(?:TD|T)?|RJCT"
)
EVENT_RE = _token(EVENT)
ELECTRONIC_RE = _token(r"NEFT|RTGS|IMPS|UPI|NACH|ACH|ECS")
ELECTRONIC_PREFIX_RE = re.compile(r"^(?:NEFT|RTGS|IMPS|UPI|NACH|ACH|ECS)")
FEE_RE = _token(r"CHARGES?|CHGS?|CHRGS?|FEES?|COMMISSION|RETRIEVAL")
COMPACT_FEE_RE = re.compile(
    r"(?:RETURN(?:ED)?|REJECT(?:ED|ION)?|DISHONOU?R(?:ED)?|BOUNC(?:E|ED)|RTN|RETD|RET)"
    rf"(?:{CHEQUE})?(?:CHARGES?|CHGS?|CHRGS?|FEES?|COMMISSION)"
    r"|^(?:CHARGES?|CHGS?|CHRGS?|FEES?|COMMISSION)(?:FOR|ON|OF)?"
    rf"(?:{CHEQUE}|RETURN|RTN|CLG|CTS)"
    r"|^(?:CGST|SGST|IGST|GST|TAX)(?:ON|FOR)?(?:CHQ|CHEQUE|RETURN|RTN)"
)
# Full return words may be joined to the next narration field. Short RET/RTN
# forms require a boundary or an explicit neighbouring clearing/cheque token.
COMPACT_EVENT_RE = re.compile(
    rf"(?:{CHEQUE}|CLG|CLEARING|CTS)(?:RETURN(?:ED)?|REJECT(?:ED|ION)?|DISHONOU?R(?:ED)?|BOUNC(?:E|ED)|UNPAID)"
    rf"|(?:{EVENT})(?:OF)?(?:{CHEQUE}|CLG|CLEARING|CTS)"
    rf"|(?:{CHEQUE}|CLG|CLEARING|CTS)(?:RTND|RETN|RETD|RTN|RET|REJ(?:TD|T)?|RJCT)(?![A-Z])"
    r"|(?:IW|OW|INW|OUTW)REJINST"
    rf"|CLG(?:INW|OUTW)(?:RET|REJ)(?={CHEQUE}|[0-9]|$)"
)
REASON_RE = re.compile(
    r"(?:FUNDS?|BALANCE)(?:ARE)?INSUFF(?:ICIENT)?"
    r"|INSUFF(?:ICIENT)?(?:FUNDS?|BALANCE)"
    r"|EXCEEDSARRANGEMENT"
    r"|PAYMENTSTOPPED|STOPPEDBY(?:THE)?DRAWER|STOPPAYMENT"
    r"|SIGNATURE(?:S)?(?:DIFFERS?|MISMATCH|MISSING|NOTASPER|REQUIRED)"
    r"|(?:ACCOUNT|ACCT|AC)(?:IS)?(?:CLOSED|FROZEN|BLOCKED)"
    rf"|STALE(?:DATED)?(?:{CHEQUE})|(?:{CHEQUE})(?:IS)?STALE"
    rf"|POSTDATED(?:{CHEQUE})|(?:{CHEQUE})(?:IS)?POSTDATED"
    r"|ALTERATIONREQUIRES(?:DRAWERS)?AUTHENTICATION"
    r"|AMOUNTINWORDSAND(?:IN)?FIGURES(?:DIFFER|DIFFERS|MISMATCH)"
)
NEGATED_RE = re.compile(
    rf"\bNO\s+(?:(?:{CHEQUE})\s+)?(?:RETURN|REJECTION|DISHONOUR)\b"
    rf"|\b(?:{CHEQUE})\s+(?:\d+\s+)?NOT\s+(?:RETURNED|REJECTED|DISHONOURED|UNPAID)\b"
    r"|\b(?:RETURN|REJECTION)\s+(?:REQUEST|REQUESTED|PENDING)\b"
    r"|\b(?:STOP\s+PAYMENT|PAYMENT\s+STOP)\s+(?:REQUEST|INSTRUCTION)\b"
)
REVERSAL_RE = _token(r"REVERSAL|REVERSED|RVSL")
COMPACT_NON_EVENT_RE = re.compile(
    rf"NO(?:{CHEQUE})?(?:RETURN|REJECTION|DISHONOUR)"
    rf"|(?:{CHEQUE})[0-9]*NOT(?:RETURNED|REJECTED|DISHONOURED|UNPAID)"
    r"|(?:RETURN|REJECTION|STOPPAYMENT)(?:REQUEST(?:ED)?|PENDING|INSTRUCTION)"
    rf"|(?:REVERSAL|REVERSED|RVSL)(?:OF)?(?:{CHEQUE}|{EVENT})"
    rf"|(?:{EVENT})[0-9]*(?:REVERSAL|REVERSED|RVSL)"
)
DATED_REASON_RE = _token(r"STALE|POST\s*DATED")
GOODS_RETURN_RE = re.compile(
    r"\b(?:(?:RETURN(?:ED)?|REJECT(?:ED|ION)?)\s+(?:GOODS?|PRODUCTS?|MERCHANDISE|ORDERS?)"
    r"|(?:GOODS?|PRODUCTS?|MERCHANDISE|ORDERS?)\s+(?:RETURN(?:ED)?|REJECT(?:ED|ION)?))\b"
)


def _text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    return "" if text in {"NAN", "NAT", "<NA>", "NONE"} else text


def is_cheque_return(details: Any, cheque_number: Any = "", detail_clean: Any = "") -> bool:
    """Return whether a row describes a cheque return rather than its fee.

    A cleaned narration is a fallback only: it must not override an explicit
    electronic transaction or fee in the original narration. Debit and credit
    are deliberately not used, as issued and deposited cheques can both return.
    """
    text = re.sub(r"[^A-Z0-9]+", " ", _text(details) or _text(detail_clean)).strip()
    if not text:
        return False
    compact = text.replace(" ", "")
    if ELECTRONIC_RE.search(text) or ELECTRONIC_PREFIX_RE.search(compact):
        return False
    if FEE_RE.search(text) or COMPACT_FEE_RE.search(compact):
        return False
    if NEGATED_RE.search(text) or REVERSAL_RE.search(text) or COMPACT_NON_EVENT_RE.search(compact):
        return False
    if GOODS_RETURN_RE.search(text) and not CONTEXT_RE.search(text):
        return False

    compact_event = bool(COMPACT_EVENT_RE.search(compact))
    event = bool(EVENT_RE.search(text)) or compact_event
    reason = bool(REASON_RE.search(compact) or DATED_REASON_RE.search(text))
    return event or reason


def is_nonposting_cheque_return(row) -> bool:
    """Recognize retained cheque events with no movement in either amount column."""
    for column in ("Debit", "Credit"):
        value = row.get(column)
        if value is None or value == "":
            continue
        try:
            if float(value) != 0:
                return False
        except (TypeError, ValueError):
            return False
    return is_cheque_return(row.get("Details"), row.get("Cheque No"), row.get("Detail_Clean"))
