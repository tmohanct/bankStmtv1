"""Regression coverage for shared OCR lookup and bank-specific failures."""
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from src.parsers import detector, hdfc_parser, icici_parser, kvb_parser, sbi_parser
from src.utils.ocr import find_tesseract


class OcrConfigurationTests(unittest.TestCase):
    def test_explicit_executable_takes_priority_over_path(self):
        with tempfile.TemporaryDirectory() as folder:
            explicit = Path(folder) / "explicit.exe"
            discovered = Path(folder) / "discovered.exe"
            explicit.touch()
            discovered.touch()
            with patch.dict("os.environ", {"TESSERACT_CMD": str(explicit)}), patch(
                "src.utils.ocr.shutil.which", return_value=str(discovered)
            ):
                self.assertEqual(find_tesseract(), str(explicit))

    def test_invalid_override_falls_back_to_path(self):
        with tempfile.TemporaryDirectory() as folder:
            discovered = Path(folder) / "discovered.exe"
            discovered.touch()
            with patch.dict("os.environ", {"TESSERACT_CMD": str(Path(folder) / "missing.exe")}), patch(
                "src.utils.ocr.shutil.which", return_value=str(discovered)
            ):
                self.assertEqual(find_tesseract(), str(discovered))

    def test_unavailable_executable_returns_none(self):
        with patch("src.utils.ocr.Path.is_file", return_value=False):
            self.assertIsNone(find_tesseract())

    def test_missing_ocr_preserves_detection_fallback_and_parser_errors(self):
        with patch.object(detector, "find_tesseract", return_value=None):
            self.assertFalse(detector._configure_tesseract())
        for parser in (icici_parser, kvb_parser, sbi_parser):
            with self.subTest(bank=parser.BANK_CODE), patch.object(parser, "find_tesseract", return_value=None):
                with self.assertRaisesRegex(FileNotFoundError, "Tesseract"):
                    parser._configure_tesseract()
        with patch.object(hdfc_parser, "find_tesseract", return_value=None):
            with self.assertRaisesRegex(RuntimeError, "HDFC"):
                hdfc_parser._tesseract_executable()

    def test_configured_executable_reaches_ocr_library(self):
        for parser in (detector, icici_parser, kvb_parser, sbi_parser):
            with self.subTest(parser=parser.__name__), patch.object(parser, "find_tesseract", return_value="chosen.exe"), patch(
                "pytesseract.pytesseract.tesseract_cmd", "original.exe"
            ):
                self.assertTrue(parser._configure_tesseract())
                self.assertEqual(parser.pytesseract.pytesseract.tesseract_cmd, "chosen.exe")


if __name__ == "__main__":
    unittest.main()
