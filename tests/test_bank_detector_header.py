"""Issuer identity must take priority over banks in transaction narration."""
import logging
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

import fitz

from src.parsers import detector
from src.utils import pdf_status_reader


class HeaderBankDetectorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "statement.pdf"
        self.logger = logging.getLogger(__name__)

    def create_statement(self, raster_logo=False, bank_name="Karur Vysya Bank"):
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((40, 40), bank_name, fontsize=12)
            logo_words = page.get_text("words")
            if raster_logo:
                logo = page.get_pixmap(clip=fitz.Rect(30, 20, 240, 50)).tobytes("png")
                doc.delete_page(0)
                page = doc.new_page()
                page.insert_image(fitz.Rect(30, 20, 240, 50), stream=logo)
            for y, text in (
                (75, "Account Statement"),
                (100, "Account Name: TEST TRADERS"),
                (125, "Account Number: 1234567890123456"),
                (150, "From Date: 01-09-2025 To Date: 27-02-2026"),
                (200, "Transaction Date Value Date Description Debit Credit Balance"),
                (225, "01-09-2025 BY CLG:TEST:ICICI Bank 700000.00 800000.00"),
                (250, "01-09-2025 NEFT ICIC0000123 100000.00 900000.00"),
            ):
                page.insert_text((40, y), text, fontsize=9)
            doc.save(self.path)
        return logo_words

    def test_native_issuer_wins_over_stronger_counterparty_signatures(self):
        self.create_statement()
        with patch.object(pdf_status_reader, "ocr_words") as ocr:
            self.assertEqual(detector.detect_bank_from_pdf(self.path, self.logger), "kvb")
        ocr.assert_not_called()

    def test_image_logo_is_checked_before_accepting_transaction_bank(self):
        logo_words = self.create_statement(raster_logo=True)
        extracted = detector._extract_with_fitz(self.path, self.logger)
        self.assertEqual(detector._detect_from_text(extracted), "icici")
        with patch.object(pdf_status_reader, "ocr_words", return_value=logo_words) as ocr:
            self.assertEqual(detector.detect_bank_from_pdf(self.path, self.logger), "kvb")
        ocr.assert_called_once()

    def test_native_nested_bank_name_uses_the_issuer(self):
        self.create_statement(bank_name="City Union Bank")
        with patch.object(pdf_status_reader, "ocr_words") as ocr:
            self.assertEqual(detector.detect_bank_from_pdf(self.path, self.logger), "cub")
        ocr.assert_not_called()

    def test_existing_text_fallback_when_identity_is_unavailable(self):
        with patch.object(detector, "read_first_page", return_value=pdf_status_reader.FirstPageResult()), \
             patch.object(detector, "_extract_with_pdfplumber", return_value="Visit www.icicibank.com"), \
             patch.object(detector, "_extract_with_fitz", return_value=""), \
             patch.object(detector, "_extract_with_ocr") as ocr:
            self.assertEqual(detector.detect_bank_from_pdf(self.path, self.logger), "icici")
        ocr.assert_not_called()


if __name__ == "__main__":
    unittest.main()
