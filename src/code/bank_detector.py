from __future__ import annotations

import os
import re
import shutil
from pathlib import Path

import fitz
import pdfplumber
import pytesseract
from PIL import Image

BANK_SIGNATURES: dict[str, tuple[tuple[str, int], ...]] = {
    "axis": (("AXIS BANK", 4), ("UTIB0", 3), ("ACCOUNT STATEMENT REPORT", 1)),
    "bob": (("BANK OF BARODA", 4), ("BARB0", 3), ("STATEMENT OF TRANSACTIONS IN CASH CREDIT ACCOUNT", 1)),
    "bom": (("BANK OF MAHARASHTRA", 4), ("MAHB0", 3), ("MAHABANK.CO.IN", 1)),
    "canara": (("CNRB0", 4), ("STATEMENT FOR A/C", 2), ("DEPOSITS WITHDRAWALS BALANCE", 1)),
    "central": (("CENTRAL BANK OF INDIA", 5), ("CBIN0", 3), ("CBIN", 1)),
    "cub": (("CITY UNION BANK", 4), ("CIUB0", 3)),
    "federal": (("FEDERAL BANK", 4), ("FDRL0", 3)),
    "hdfc": (("HDFC BANK", 4), ("HDFC0", 3)),
    "icici": (("ICICI BANK", 4), ("ICIC0", 3)),
    "idbi": (("IDBI BANK", 4), ("IBKL0", 3), ("YOUR A/C STATUS", 1)),
    "idfc": (("IDFC FIRST BANK", 4), ("IDFB0", 3)),
    "indian": (("IDIB0", 5), ("ACCOUNT STATEMENT", 2), ("ACCOUNT ACTIVITY", 2)),
    "indus": (("INDUSIND BANK", 4), ("INDB0", 3)),
    "iob": (("INDIAN OVERSEAS BANK", 4), ("IOBA0", 3)),
    "kvb": (("KARUR VYSYA BANK", 4), ("KVBL0", 3)),
    "kotak": (("KOTAK MAHINDRA BANK", 4), ("KKBK", 3), ("CURRENT ACCOUNT TRANSACTIONS", 1)),
    "pnb": (("PUNJAB NATIONAL BANK", 4), ("PUNB0", 3)),
    "sbi": (("STATE BANK OF INDIA", 4), ("SBIN0", 3)),
    "southind": (("SOUTH INDIAN BANK", 4), ("SIBL", 3), ("STATEMENT OF ACCOUNT", 1)),
    "unionbank": (("UNION BANK", 4), ("UBIN", 3), ("TRANSACTION ID", 1)),
}


def _normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text.upper()).strip()


def _extract_with_pdfplumber(pdf_path: Path, logger) -> str:
    chunks: list[str] = []
    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            for page in pdf.pages[:2]:
                chunks.append(page.extract_text() or "")
    except Exception as exc:  # noqa: BLE001
        logger.debug("pdfplumber detection text extraction failed for %s: %s", pdf_path.name, exc)
    return "\n".join(chunks)


def _extract_with_fitz(pdf_path: Path, logger) -> str:
    chunks: list[str] = []
    try:
        with fitz.open(str(pdf_path)) as pdf:
            for page in list(pdf)[:2]:
                chunks.append(page.get_text() or "")
    except Exception as exc:  # noqa: BLE001
        logger.debug("PyMuPDF detection text extraction failed for %s: %s", pdf_path.name, exc)
    return "\n".join(chunks)


def _configure_tesseract() -> bool:
    candidates: list[Path] = []

    env_value = os.environ.get("TESSERACT_CMD")
    if env_value:
        candidates.append(Path(env_value))

    resolved = shutil.which("tesseract")
    if resolved:
        candidates.append(Path(resolved))

    candidates.extend(
        [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            pytesseract.pytesseract.tesseract_cmd = str(candidate)
            return True

    return False


def _extract_with_ocr(pdf_path: Path, logger) -> str:
    if not _configure_tesseract():
        logger.debug("Tesseract not available for OCR bank detection on %s", pdf_path.name)
        return ""

    try:
        with fitz.open(str(pdf_path)) as pdf:
            if not pdf:
                return ""

            page = pdf[0]
            pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            image = Image.frombytes("RGB", [pixmap.width, pixmap.height], pixmap.samples)
            return pytesseract.image_to_string(image)
    except Exception as exc:  # noqa: BLE001
        logger.debug("OCR bank detection failed for %s: %s", pdf_path.name, exc)
        return ""


def _score_bank_matches(text: str) -> dict[str, int]:
    scores: dict[str, int] = {}
    for bank_code, signatures in BANK_SIGNATURES.items():
        score = 0
        for signature, weight in signatures:
            if signature in text:
                score += weight
        if score > 0:
            scores[bank_code] = score
    return scores


def _detect_from_text(text: str) -> str | None:
    normalized = _normalize_text(text)
    if not normalized:
        return None

    scores = _score_bank_matches(normalized)
    if not scores:
        return None

    return max(scores.items(), key=lambda item: (item[1], item[0]))[0]


def detect_bank_from_pdf(pdf_path: Path, logger) -> str:
    extracted_text = "\n".join(
        [
            _extract_with_pdfplumber(pdf_path, logger),
            _extract_with_fitz(pdf_path, logger),
        ]
    )
    bank_code = _detect_from_text(extracted_text)
    if bank_code:
        logger.info("Auto-detected bank '%s' from extracted PDF text for %s", bank_code, pdf_path.name)
        return bank_code

    ocr_text = _extract_with_ocr(pdf_path, logger)
    bank_code = _detect_from_text(ocr_text)
    if bank_code:
        logger.info("Auto-detected bank '%s' from OCR text for %s", bank_code, pdf_path.name)
        return bank_code

    filename_guess = _detect_from_text(pdf_path.stem)
    if filename_guess:
        logger.warning(
            "Bank auto-detection fell back to filename match for %s: %s",
            pdf_path.name,
            filename_guess,
        )
        return filename_guess

    raise ValueError(
        f"Unable to auto-detect bank from PDF content for '{pdf_path.name}'. Re-run with --bank <code>."
    )
