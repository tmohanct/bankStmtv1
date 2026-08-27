"""Shared text-normalization helpers."""

from __future__ import annotations

import re
from typing import Any


def clean_cell(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\r", " ").replace("\n", " ")
    return re.sub(r"\s+", " ", text).strip()


def clean_detail(value: Any) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", clean_cell(value))
