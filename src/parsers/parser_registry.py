"""Central parser registry for all bank statement parsers."""

from __future__ import annotations

from typing import Callable

from parsers.axis_parser import AxisParser
from parsers.base_parser import BaseStatementParser
from parsers.iob_parser import IOBParser
from parsers.kotak_parser import KotakParser
from parsers.southind_parser import SouthIndianParser
from parsers.unionbank_parser import UnionBankParser

ParserFactory = Callable[[], BaseStatementParser]

PARSER_REGISTRY: dict[str, ParserFactory] = {
    "axis": AxisParser,
    "iob": IOBParser,
    "kotak": KotakParser,
    "southind": SouthIndianParser,
    "unionbank": UnionBankParser,
}


def list_supported_banks() -> list[str]:
    return sorted(PARSER_REGISTRY.keys())


def get_parser_factory(bank_code: str) -> ParserFactory | None:
    return PARSER_REGISTRY.get(bank_code.strip().lower())


def register_parser(bank_code: str, parser_factory: ParserFactory) -> None:
    normalized_code = bank_code.strip().lower()
    if not normalized_code:
        raise ValueError("bank_code must not be empty")

    # TODO: Add collision policy if dynamic plugin loading is introduced.
    PARSER_REGISTRY[normalized_code] = parser_factory
