"""Locate the optional Tesseract executable consistently across parsers."""
import os
from pathlib import Path
import shutil
from contextlib import contextmanager
from contextvars import ContextVar


_observer = ContextVar("ocr_observer", default=None)


@contextmanager
def observe_ocr(callback):
    token = _observer.set(callback)
    try:
        yield
    finally:
        _observer.reset(token)


def record_ocr_use():
    """Record actual OCR invocations, including unsuccessful attempts."""
    callback = _observer.get()
    if callback is not None:
        callback()


def find_tesseract() -> str | None:
    candidates = (
        os.environ.get("TESSERACT_CMD"),
        shutil.which("tesseract"),
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    )
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(candidate)
    return None
