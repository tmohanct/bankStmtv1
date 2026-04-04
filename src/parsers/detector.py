"""Parser selection and future bank detection utilities."""

from __future__ import annotations

from pathlib import Path

from parsers.base_parser import BaseStatementParser
from parsers.parser_registry import get_parser_factory, list_supported_banks


def supported_banks() -> list[str]:
    return list_supported_banks()


def get_parser(bank_code: str) -> BaseStatementParser:
    parser_factory = get_parser_factory(bank_code)
    if parser_factory is None:
        options = ", ".join(supported_banks())
        raise ValueError(f"Unsupported bank '{bank_code}'. Supported banks: {options}")
    return parser_factory()


def detect_bank_from_pdf(pdf_path: Path) -> str:
    # TODO: Implement automatic bank detection heuristics for multi-bank ingestion.
    # Keep this generic so new parsers can be added without changing call sites.
    raise NotImplementedError(f"Bank auto-detection is not implemented for file: {pdf_path}")
