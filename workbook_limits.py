"""Cheap XLSX checks before openpyxl allocates a full editable workbook."""

import posixpath
import re
from xml.etree import ElementTree as ET
from zipfile import BadZipFile, ZipFile

MAX_DATA_ROWS = 10_000
QUICK_ROWS = 1_000
MAX_UPLOAD_BYTES = 25 * 1024 * 1024
MAX_EXPANDED_BYTES = 80 * 1024 * 1024
MAX_WORKBOOK_CELLS = 400_000
MAX_BLOCK_COLUMNS = 64
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def validate_settings(config):
    from alignment import block_width_cols, col_to_num, output_block_extent

    if not 1 <= config.header_first_row <= config.header_last_row < config.start_row <= 101:
        raise ValueError("Choose header rows above the data, with the first data row between 2 and 101.")
    import math
    if not math.isfinite(config.threshold) or config.threshold < 0:
        raise ValueError("Max diff must be a finite number greater than or equal to zero.")
    width = 0
    for side in ("left", "right"):
        start = getattr(config, f"{side}_block_start_col")
        end = getattr(config, f"{side}_block_end_col")
        width += block_width_cols(start, end)
        if col_to_num(end) > 16384 or output_block_extent(
            getattr(config, f"{side}_output_start_col"), start, end
        )[1] > 16384:
            raise ValueError("Column settings must fit within Excel's column range.")
    if width > MAX_BLOCK_COLUMNS:
        raise ValueError(f"Select no more than {MAX_BLOCK_COLUMNS} columns across the two blocks.")
    if config.diff_output_col and col_to_num(config.diff_output_col) > 16384:
        raise ValueError("The difference column must fit within Excel's column range.")


def inspect_workbook(path, config):
    """Count data row spans, ignoring empty styled trailing rows; never truncate data."""
    validate_settings(config)
    if path.stat().st_size > MAX_UPLOAD_BYTES:
        raise ValueError("The upload limit is 25 MB. Save a smaller .xlsx workbook and try again.")
    try:
        with ZipFile(path) as archive:
            if sum(item.file_size for item in archive.infolist()) > MAX_EXPANDED_BYTES:
                raise ValueError("This workbook expands beyond 80 MB. Remove unused sheets or excess formatting.")
            relationships = {
                item.attrib["Id"]: posixpath.normpath(
                    item.attrib["Target"].lstrip("/") if item.attrib["Target"].startswith("/")
                    else posixpath.join("xl", item.attrib["Target"])
                ) for item in ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
                if item.attrib.get("TargetMode") != "External"
            }
            sheets = {
                sheet.attrib["name"]: relationships[sheet.attrib[REL + "id"]]
                for sheet in ET.fromstring(archive.read("xl/workbook.xml")).findall(NS + "sheets/" + NS + "sheet")
            }
            def resolve(name):
                if not name:
                    return next(iter(sheets))
                for title in sheets:
                    if title.casefold() == name.casefold():
                        return title
                raise ValueError(f'Sheet "{name}" was not found. Check the Sheet in setting.')

            selected = {
                resolve(config.left_sheet_name or config.input_sheet_name),
                resolve(config.right_sheet_name or config.left_sheet_name or config.input_sheet_name),
            }
            if config.output_sheet_name in selected:
                raise ValueError("Sheet out must have a different name from the input sheets.")
            row_counts = {}
            total_cells = 0
            # Count cells across retained sheets too: openpyxl loads the whole workbook.
            for name, member in sheets.items():
                last_data_row = 0
                with archive.open(member) as stream:
                    for _, element in ET.iterparse(stream, events=("end",)):
                        if element.tag == NS + "row":
                            cells = element.findall(NS + "c")
                            total_cells += len(cells)
                            if total_cells > MAX_WORKBOOK_CELLS:
                                raise ValueError("This workbook contains too many cells or excess formatting. Keep only the sheets needed for alignment.")
                            if name in selected:
                                for cell in cells:
                                    has_value = any(
                                        node.text is not None
                                        for node in cell.iter()
                                        if node.tag in {NS + "v", NS + "t", NS + "f"}
                                    ) or cell.find(NS + "f") is not None
                                    if has_value:
                                        match = re.search(r"(\d+)$", cell.attrib.get("r", ""))
                                        row = int(match.group(1)) if match else int(element.attrib["r"])
                                        last_data_row = max(last_data_row, row)
                                count = max(0, last_data_row - config.start_row + 1)
                                if count > MAX_DATA_ROWS:
                                    raise ValueError(f'Sheet "{name}" exceeds the 10,000 data-row limit (headers excluded). Split it into smaller workbooks; no rows were processed.')
                            element.clear()
                if name in selected:
                    row_counts[name] = max(0, last_data_row - config.start_row + 1)
            return {"rows": max(row_counts.values(), default=0), "sheet_rows": row_counts}
    except (BadZipFile, ET.ParseError, KeyError, StopIteration, OSError) as exc:
        raise ValueError("This file could not be read as an Excel workbook. Save it as .xlsx in Excel and try again.") from exc
