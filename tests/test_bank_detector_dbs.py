from __future__ import annotations

import sys
import unittest
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "src" / "code"))

import bank_detector


class DBSBankDetectorTests(unittest.TestCase):
    def test_dbs_layout_outscores_sbi_mentions_in_narration(self) -> None:
        text = """
        Account Details
        Earmark Amount
        Trans. Date Value Date Transaction Details Debits Credits Running Balance
        NEFT STATE BANK OF INDIA SBIN0016320
        RTGS STATE BANK OF INDIA SBIN0000796
        """
        self.assertEqual(bank_detector._detect_from_text(text), "dbs")


if __name__ == "__main__":
    unittest.main()
