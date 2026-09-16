"""First-page header profile used only by PDF_Status."""

# PDF_Status only. Transaction parsing does not use this profile.
PDF_STATUS_PROFILE = {
    "name": "HDFC Bank",
    "ifsc": "HDFC",
    "aliases": [
        "HDFC Bank"
    ],
    "unlabelled_left": True
}



PDF_STATUS_PROFILE.update({'address_is_branch': True})

# Image-only HDFC transaction statements are parsed here. The older text
# statement parser in src/code/hdfc_parser.py delegates to this path when the
# PDF has no usable text layer.
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
import os
import re
import shutil

import fitz
from PIL import Image
import pytesseract


_OCR_ZOOM = 3
_DATE_RE = re.compile(r"\d{2}/\d{2}/\d{2}")
_MONEY_RE = re.compile(r"\d[\d,]*\.\d{2}")
_SUMMARY_RE = re.compile(
    r"^\s*([\d,]+\.\d{2})\s+(\d+)\s+(\d+)\s+"
    r"([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s+([\d,]+\.\d{2})\s*$"
)


@dataclass(frozen=True)
class _Word:
    text: str
    left: int
    top: int
    width: int
    height: int

    @property
    def right(self) -> int:
        return self.left + self.width


@dataclass
class _ScannedRow:
    page_index: int
    top: int
    height: int
    date: str
    detail_parts: list[str] = field(default_factory=list)
    cheque_parts: list[str] = field(default_factory=list)
    amount_text: str = ""
    balance_text: str = ""
    is_credit: bool = False


@dataclass(frozen=True)
class _StatementSummary:
    opening: Decimal
    debit_count: int
    credit_count: int
    debits: Decimal
    credits: Decimal
    closing: Decimal


def _tesseract_executable() -> str:
    candidates = [
        os.environ.get("TESSERACT_CMD"),
        shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(candidate)
    raise RuntimeError(
        "Scanned HDFC statements require Tesseract OCR. Install Tesseract "
        "or set TESSERACT_CMD to its executable path."
    )


def _ocr_lines(image: Image.Image) -> list[list[_Word]]:
    data = pytesseract.image_to_data(
        image, config="--psm 6", output_type=pytesseract.Output.DICT
    )
    grouped: dict[tuple[int, int, int], list[_Word]] = defaultdict(list)
    for index, raw in enumerate(data["text"]):
        value = str(raw).strip()
        if value:
            key = (
                data["block_num"][index],
                data["par_num"][index],
                data["line_num"][index],
            )
            grouped[key].append(
                _Word(
                    value,
                    int(data["left"][index]),
                    int(data["top"][index]),
                    int(data["width"][index]),
                    int(data["height"][index]),
                )
            )
    return sorted(
        (sorted(words, key=lambda word: word.left) for words in grouped.values()),
        key=lambda words: (min(word.top for word in words), words[0].left),
    )


def _field(words: list[_Word]) -> str:
    return " ".join(
        part for word in words if (part := word.text.strip(" |\\[](){}"))
    ).strip()


def _money(raw: str) -> Decimal | None:
    compact = re.sub(r"\s+", "", raw).strip("|[](){}")
    if not _MONEY_RE.fullmatch(compact):
        return None
    try:
        return Decimal(compact.replace(",", ""))
    except InvalidOperation:
        return None


def _summary_from_line(words: list[_Word]) -> _StatementSummary | None:
    match = _SUMMARY_RE.fullmatch(" ".join(word.text for word in words))
    if not match:
        return None
    amounts = [_money(match.group(index)) for index in (1, 4, 5, 6)]
    if any(amount is None for amount in amounts):
        return None
    return _StatementSummary(
        amounts[0], int(match.group(2)), int(match.group(3)),
        amounts[1], amounts[2], amounts[3],
    )


def _column_words(words: list[_Word], width: int, start: float, end: float) -> list[_Word]:
    return [word for word in words if start <= word.left / width < end]


def _append_detail(row: _ScannedRow, words: list[_Word], width: int) -> None:
    detail = _field(_column_words(words, width, 0.085, 0.44))
    cheque = _field(_column_words(words, width, 0.44, 0.565))
    if detail:
        row.detail_parts.append(detail)
    if cheque:
        row.cheque_parts.append(cheque)


def _rows_on_page(
    lines: list[list[_Word]], page_index: int, width: int
) -> tuple[list[_ScannedRow], _StatementSummary | None]:
    rows: list[_ScannedRow] = []
    current: _ScannedRow | None = None
    summary: _StatementSummary | None = None
    in_footer = False
    for words in lines:
        text = " ".join(word.text for word in words)
        found_summary = _summary_from_line(words)
        if found_summary:
            summary = found_summary
        if "STATEMENT SUMMARY" in text.upper() or "HDFC BANK LIMITED" in text.upper():
            in_footer = True
        if in_footer:
            continue
        date_word = next(
            (
                word for word in words
                if word.left / width < 0.085
                and _DATE_RE.fullmatch(word.text.strip(" |\\[](){}"))
            ),
            None,
        )
        if date_word:
            amount_words = _column_words(words, width, 0.635, 0.89)
            balance_words = _column_words(words, width, 0.89, 1.01)
            current = _ScannedRow(
                page_index=page_index,
                top=min(word.top for word in words),
                height=(
                    max(word.top + word.height for word in words)
                    - min(word.top for word in words)
                ),
                date=date_word.text.strip(" |\\[](){}"),
                amount_text=_field(amount_words).replace(" ", ""),
                balance_text=_field(balance_words).replace(" ", ""),
                is_credit=bool(
                    amount_words
                    and max(word.right for word in amount_words) / width > 0.81
                ),
            )
            _append_detail(current, words, width)
            rows.append(current)
        elif current is not None:
            _append_detail(current, words, width)
    return rows, summary


def _reread_money(document: fitz.Document, row: _ScannedRow, kind: str) -> Decimal | None:
    pixmap = document[row.page_index].get_pixmap(
        matrix=fitz.Matrix(_OCR_ZOOM, _OCR_ZOOM), alpha=False
    )
    image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
    if kind == "balance":
        left, right = 0.89, 1.0
    elif row.is_credit:
        left, right = 0.76, 0.89
    else:
        left, right = 0.635, 0.76
    cell = image.crop(
        (
            int(image.width * left),
            max(0, row.top - 12),
            int(image.width * right),
            min(image.height, row.top + max(42, row.height + 20)),
        )
    )
    raw = pytesseract.image_to_string(
        cell,
        config="--psm 7 -c tessedit_char_whitelist=0123456789,.",
    )
    return _money(raw.strip())



def _balance_is_close(observed: Decimal | None, expected: Decimal) -> bool:
    return observed is not None and abs(observed - expected) <= Decimal("0.05")


def _partial_balance_matches(raw: str, expected: Decimal) -> bool:
    # OCR can blur or insert one character in the final cents at the right edge.
    # The complete rupee value must still match the running balance.
    compact = re.sub(r"\s+", "", raw).strip("|")
    match = re.fullmatch(r"([\d,]+)\.([\d:;?]{1,3})", compact)
    if not match:
        return False
    expected_integer, expected_cents = f"{expected:,.2f}".split(".")
    if match.group(1) != expected_integer:
        return False
    observed_cents = match.group(2)
    if len(observed_cents) == len(expected_cents):
        return sum(a != b for a, b in zip(observed_cents, expected_cents)) <= 1
    if len(observed_cents) == len(expected_cents) + 1:
        return any(
            observed_cents[:index] + observed_cents[index + 1:] == expected_cents
            for index in range(len(observed_cents))
        )
    if len(observed_cents) + 1 == len(expected_cents):
        return any(
            expected_cents[:index] + expected_cents[index + 1:] == observed_cents
            for index in range(len(expected_cents))
        )
    return False


def parse_scanned_hdfc(pdf_path: str, logger, progress_cb=None) -> list[dict[str, object]]:
    """Read an image-only HDFC statement and require full summary reconciliation."""
    pytesseract.pytesseract.tesseract_cmd = _tesseract_executable()
    scanned_rows: list[_ScannedRow] = []
    summary: _StatementSummary | None = None
    with fitz.open(pdf_path) as document:
        for page_index, page in enumerate(document):
            pixmap = page.get_pixmap(
                matrix=fitz.Matrix(_OCR_ZOOM, _OCR_ZOOM), alpha=False
            )
            image = Image.frombytes(
                "RGB", (pixmap.width, pixmap.height), pixmap.samples
            )
            page_rows, page_summary = _rows_on_page(
                _ocr_lines(image), page_index, pixmap.width
            )
            scanned_rows.extend(page_rows)
            if page_summary is not None:
                summary = page_summary
            logger.info(
                "HDFC OCR page %s/%s: %s transaction rows",
                page_index + 1, len(document), len(page_rows),
            )

        if summary is None:
            raise ValueError(
                "HDFC OCR could not read the statement summary; "
                "transaction totals cannot be verified."
            )
        expected_count = summary.debit_count + summary.credit_count
        if len(scanned_rows) != expected_count:
            raise ValueError(
                f"HDFC OCR found {len(scanned_rows)} transactions; "
                f"statement summary expects {expected_count}."
            )

        records: list[dict[str, object]] = []
        previous_balance = summary.opening
        debit_count = credit_count = 0
        total_debit = total_credit = Decimal("0.00")
        corrected_cents = 0
        for index, row in enumerate(scanned_rows, start=1):
            amount = _money(row.amount_text)
            if amount is None:
                amount = _reread_money(document, row, "amount")
            if amount is None:
                raise ValueError(
                    f"HDFC OCR could not read the amount on page "
                    f"{row.page_index + 1}, transaction {index}."
                )

            expected_balance = (
                previous_balance + amount if row.is_credit
                else previous_balance - amount
            )
            observed_balance = _money(row.balance_text)
            if not _balance_is_close(observed_balance, expected_balance):
                reread_balance = _reread_money(document, row, "balance")
                if reread_balance is not None:
                    observed_balance = reread_balance
            if not _balance_is_close(observed_balance, expected_balance):
                reread_amount = _reread_money(document, row, "amount")
                if reread_amount is not None:
                    candidate_balance = (
                        previous_balance + reread_amount if row.is_credit
                        else previous_balance - reread_amount
                    )
                    if (
                        observed_balance is not None
                        and abs(observed_balance - candidate_balance) <= Decimal("0.05")
                    ):
                        amount = reread_amount
                        expected_balance = candidate_balance
            if (
                not _balance_is_close(observed_balance, expected_balance)
                and _partial_balance_matches(row.balance_text, expected_balance)
            ):
                observed_balance = expected_balance
            if not _balance_is_close(observed_balance, expected_balance):
                raise ValueError(
                    f"HDFC OCR balance mismatch on page {row.page_index + 1}, "
                    f"transaction {index}: previous={previous_balance}, "
                    f"amount={amount}, printed={row.balance_text!r}."
                )
            if observed_balance != expected_balance:
                corrected_cents += 1

            details = re.sub(r"\s+", " ", " ".join(row.detail_parts)).strip()
            cheque = re.sub(r"\s+", " ", " ".join(row.cheque_parts)).strip()
            if not details:
                raise ValueError(
                    f"HDFC OCR found no narration on page "
                    f"{row.page_index + 1}, transaction {index}."
                )
            try:
                transaction_date = datetime.strptime(
                    row.date, "%d/%m/%y"
                ).strftime("%d/%m/%Y")
            except ValueError as exc:
                raise ValueError(
                    f"HDFC OCR could not read the date on page "
                    f"{row.page_index + 1}, transaction {index}."
                ) from exc

            debit = None if row.is_credit else float(amount)
            credit = float(amount) if row.is_credit else None
            if row.is_credit:
                credit_count += 1
                total_credit += amount
            else:
                debit_count += 1
                total_debit += amount
            records.append(
                {
                    "Sno": index,
                    "Date": transaction_date,
                    "Details": details,
                    "Detail_Clean": re.sub(r"[^A-Za-z0-9]", "", details),
                    "Cheque No": cheque,
                    "Debit": debit,
                    "Credit": credit,
                    "Balance": float(expected_balance),
                }
            )
            previous_balance = expected_balance
            if progress_cb is not None:
                progress_cb(index)

    if (
        debit_count != summary.debit_count
        or credit_count != summary.credit_count
        or total_debit != summary.debits
        or total_credit != summary.credits
        or previous_balance != summary.closing
    ):
        raise ValueError(
            "HDFC OCR totals do not match the statement summary: "
            f"debits {debit_count}/{total_debit} vs "
            f"{summary.debit_count}/{summary.debits}; "
            f"credits {credit_count}/{total_credit} vs "
            f"{summary.credit_count}/{summary.credits}; "
            f"closing {previous_balance} vs {summary.closing}."
        )
    logger.info(
        "HDFC OCR reconciled: rows=%s debits=%s credits=%s "
        "closing=%s corrected_balance_cents=%s",
        len(records), total_debit, total_credit,
        previous_balance, corrected_cents,
    )
    return records
