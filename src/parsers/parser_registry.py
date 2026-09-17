"""Discover bank parsers once; all entry points use this registry."""
from importlib import import_module
from pathlib import Path
from typing import Callable


def _modules():
    for path in sorted(Path(__file__).parent.glob("*_parser.py")):
        module = import_module(f"src.parsers.{path.stem}")
        if getattr(module, "BANK_CODE", None) and callable(getattr(module, "parse", None)):
            yield module


_BANK_MODULES = tuple(_modules())
PARSER_REGISTRY = {module.BANK_CODE: module.parse for module in _BANK_MODULES}


def bank_signatures():
    return {module.BANK_CODE: module.BANK_SIGNATURES for module in _BANK_MODULES}


def list_supported_banks() -> list[str]:
    return sorted(PARSER_REGISTRY)


def get_parser(bank_code: str) -> Callable:
    try:
        return PARSER_REGISTRY[bank_code.strip().lower()]
    except KeyError:
        raise ValueError(f"Unsupported bank '{bank_code}'. Supported banks: {', '.join(list_supported_banks())}") from None


def register_parser(bank_code: str, parser: Callable) -> None:
    code = bank_code.strip().lower()
    if not code or not callable(parser):
        raise ValueError("A bank code and callable parser are required")
    if code in PARSER_REGISTRY:
        raise ValueError(f"Parser already registered: {code}")
    PARSER_REGISTRY[code] = parser
