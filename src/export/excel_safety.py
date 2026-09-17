"""Text safety shared by intermediate and final workbook exports."""
import re
from pathlib import Path
from tempfile import NamedTemporaryFile
from xml.etree import ElementTree
from zipfile import ZipFile

ILLEGAL_CONTROL_CHARACTERS = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")


def sanitize_excel_value(value):
    return ILLEGAL_CONTROL_CHARACTERS.sub("", value) if isinstance(value, str) else value


def sanitize_excel_frame(frame):
    return frame.map(sanitize_excel_value)


def force_leading_equals_to_text(workbook):
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.data_type = "s"


def suppress_number_as_text_warnings(path: Path, sheet_ranges: dict[int, str]) -> None:
    """Suppress only intentional text identifiers after the final openpyxl save.

    openpyxl does not serialize ignoredErrors. Patch the generated sheet XML
    without changing values, styles, or other workbook archive members.
    Keys are one-based worksheet positions in this newly generated workbook.
    """
    if not sheet_ranges:
        return
    members = {f"xl/worksheets/sheet{index}.xml": ref for index, ref in sheet_ranges.items()}
    # These elements follow ignoredErrors in the worksheet schema.
    successor = re.compile(
        rb"<(?:smartTags|drawing|legacyDrawing|legacyDrawingHF|picture|oleObjects|"
        rb"controls|webPublishItems|tableParts|extLst)(?=[\s/>])|</worksheet>"
    )
    with NamedTemporaryFile(dir=path.parent, suffix=".xlsx", delete=False) as temporary:
        temporary_path = Path(temporary.name)
    try:
        with ZipFile(path) as source, ZipFile(temporary_path, "w") as target:
            for member in source.infolist():
                data = source.read(member.filename)
                if member.filename in members:
                    errors = ElementTree.Element("ignoredErrors")
                    ElementTree.SubElement(errors, "ignoredError", {
                        "sqref": members[member.filename], "numberStoredAsText": "1",
                    })
                    insertion = successor.search(data)
                    if insertion is None:
                        raise ValueError(f"Missing worksheet end in {member.filename}")
                    position = insertion.start()
                    data = data[:position] + ElementTree.tostring(errors) + data[position:]
                target.writestr(member, data)
        temporary_path.replace(path)
    finally:
        temporary_path.unlink(missing_ok=True)
