from __future__ import annotations

import re

from pathlib import Path

import fitz
import pdfplumber
import pytesseract
from PIL import Image

from src.utils.ocr import find_tesseract, record_ocr_use
from src.parsers.parser_registry import bank_signatures

BANK_SIGNATURES = bank_signatures()


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
            for index in range(min(len(pdf), 2)):
                chunks.append(pdf[index].get_text() or "")
    except Exception as exc:  # noqa: BLE001
        logger.debug("PyMuPDF detection text extraction failed for %s: %s", pdf_path.name, exc)
    return "\n".join(chunks)


def _configure_tesseract() -> bool:
    executable = find_tesseract()
    if executable is None:
        return False
    pytesseract.pytesseract.tesseract_cmd = executable
    return True


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
            record_ocr_use()
            return pytesseract.image_to_string(image)
    except Exception as exc:  # noqa: BLE001
        logger.debug("OCR bank detection failed for %s: %s", pdf_path.name, exc)
        return ""


def _matched_signatures(text: str):
    matches = [(code, signature, weight) for code, signatures in BANK_SIGNATURES.items()
               for signature, weight in signatures if signature in text]
    # A bank name nested in another name is not independent evidence:
    # "INDIAN BANK" in "SOUTH INDIAN BANK", or "UNION BANK" in "CITY UNION BANK".
    return [(code, signature, weight) for code, signature, weight in matches
            if not any(other_code != code and signature != other and signature in other
                       and other_weight >= 3 for other_code, other, other_weight in matches)]


def _score_bank_matches(text: str) -> dict[str, int]:
    scores: dict[str, int] = {}
    for code, signature, weight in _matched_signatures(text):
        scores[code] = scores.get(code, 0) + weight
    return scores


def _detect_from_text(text: str) -> str | None:
    normalized = _normalize_text(text)
    if not normalized:
        return None

    scores = _score_bank_matches(normalized)
    if not scores:
        return None

    ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    bank, score = ranked[0]
    # Generic headings and tied evidence are not a bank identity.
    strongest = max(weight for code, signature, weight in _matched_signatures(normalized) if code == bank)
    if strongest < 3 or (len(ranked) > 1 and ranked[1][1] == score):
        return None
    return bank


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
