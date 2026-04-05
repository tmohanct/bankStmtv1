from __future__ import annotations

import sys
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import utils


class _FakePage:
    def __init__(self, text: str) -> None:
        self._text = text

    def get_text(self, mode: str) -> str:
        assert mode == "text"
        return self._text


class _FakeDocument:
    def __init__(self, pages: list[_FakePage]) -> None:
        self._pages = pages

    def __len__(self) -> int:
        return len(self._pages)

    def __getitem__(self, index: int) -> _FakePage:
        return self._pages[index]

    def __enter__(self) -> "_FakeDocument":
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        return None


class _Logger:
    def debug(self, *args, **kwargs) -> None:
        return None


class SummaryMetricsTests(unittest.TestCase):
    def test_summary_candidate_page_indexes_prefers_edge_pages(self) -> None:
        self.assertEqual(utils._summary_candidate_page_indexes(0), [])
        self.assertEqual(utils._summary_candidate_page_indexes(4), [0, 1, 2, 3])
        self.assertEqual(utils._summary_candidate_page_indexes(10), [0, 1, 2, 7, 8, 9])

    def test_extract_summary_metrics_uses_fitz_edge_pages(self) -> None:
        pages = [_FakePage("no summary here") for _ in range(10)]
        pages[9] = _FakePage("Total debit 1,234.00 Total credit 2,345.00 Total transactions 12")

        with patch.object(utils.fitz, "open", return_value=_FakeDocument(pages)):
            metrics = utils.extract_summary_metrics("dummy.pdf", _Logger())

        self.assertEqual(
            metrics,
            {
                "total_debit": 1234.0,
                "total_credit": 2345.0,
                "transaction_count": 12.0,
            },
        )


if __name__ == "__main__":
    unittest.main()
