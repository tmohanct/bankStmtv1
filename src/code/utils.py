"""Compatibility facade for the former shared helpers."""
import sys
from pathlib import Path
ROOT = Path(__file__).resolve().parents[2]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from src.utils.statement_utils import *
from src.utils.statement_utils import _summary_candidate_page_indexes, _resolve_with_dollar_suffix
from src.transform.normalize import *
from src.transform.validate import reconcile
from src.export.excel_writer import write_output_excel
