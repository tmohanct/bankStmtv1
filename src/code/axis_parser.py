"""Compatibility import; implementation lives in src.parsers.axis_parser."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from src.parsers import axis_parser as _implementation

sys.modules[__name__] = _implementation
