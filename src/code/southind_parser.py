from __future__ import annotations

from parsers.southind_parser import parse_southind_records


def parse(pdf_path: str, logger, progress_cb=None):
    return parse_southind_records(pdf_path=pdf_path, logger=logger, progress_cb=progress_cb)
