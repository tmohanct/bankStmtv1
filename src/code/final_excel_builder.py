"""Compatibility import; implementation lives in src.export.final_excel_builder."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from src.export import final_excel_builder as _implementation

sys.modules[__name__] = _implementation
