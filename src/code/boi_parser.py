from parsers.boi_parser import parse_boi_records


def parse(pdf_path: str, logger, progress_cb=None):
    return parse_boi_records(pdf_path=pdf_path, logger=logger, progress_cb=progress_cb)


__all__ = ["parse"]
