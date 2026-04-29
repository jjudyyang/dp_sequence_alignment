from copy import copy

from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string


# =========================
# Config
# =========================

INPUT_FILE = "input.xlsx"
OUTPUT_FILE = "output.xlsx"

# Change this if your actual Excel tab name is different.
# Example: "original data" or "sheet1"
INPUT_SHEET_NAME = "sheet1"

OUTPUT_SHEET_NAME = "sheet 1 aligned result"

START_ROW = 2
HEADER_ROW = 1

# Left side: match using D, move A:D together
LEFT_INPUT_COL = "D"
LEFT_BLOCK_START_COL = "A"
LEFT_BLOCK_END_COL = "D"
LEFT_OUTPUT_START_COL = "A"

# Right side: match using F, move F:H together
RIGHT_INPUT_COL = "F"
RIGHT_BLOCK_START_COL = "F"
RIGHT_BLOCK_END_COL = "H"
RIGHT_OUTPUT_START_COL = "F"

# Matching rule
THRESHOLD = 0.5

# DP scoring
MATCH_SCORE = 2
MISMATCH_SCORE = -10
GAP_SCORE = -1

# Highlight unmatched rows orange
UNMATCHED_FILL = PatternFill(fill_type="solid", fgColor="FFA500")


# =========================
# Helpers
# =========================

def to_float(value):
    """
    Convert Excel value to float.
    Empty or invalid cells become None.
    """
    try:
        if value is None or value == "":
            return None
        return float(value)
    except Exception:
        return None


def col_to_num(col_letter):
    """
    Convert Excel column letter to number.

    Example:
        A -> 1
        D -> 4
        F -> 6
    """
    return column_index_from_string(col_letter)


def copy_cell(src_cell, dst_cell):
    """
    Copy cell value and style from one cell to another.
    """
    dst_cell.value = src_cell.value

    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)


def copy_block(src_ws, dst_ws, src_row, dst_row, src_start_col, src_end_col, dst_start_col):
    """
    Copy a horizontal block from one row to another.

    Example:
        Copy A3:D3 into A8:D8
    """
    src_start = col_to_num(src_start_col)
    src_end = col_to_num(src_end_col)
    dst_start = col_to_num(dst_start_col)

    for offset, src_col in enumerate(range(src_start, src_end + 1)):
        src_cell = src_ws.cell(row=src_row, column=src_col)
        dst_cell = dst_ws.cell(row=dst_row, column=dst_start + offset)
        copy_cell(src_cell, dst_cell)


def highlight_block(ws, row, start_col, end_col):
    """
    Highlight a row block orange.
    """
    start = col_to_num(start_col)
    end = col_to_num(end_col)

    for col in range(start, end + 1):
        ws.cell(row=row, column=col).fill = UNMATCHED_FILL


def read_column(ws, col_letter):
    """
    Read one matching column into a list.

    Each item stores:
    - original Excel row
    - numeric value used for matching
    """
    rows = []

    for row in range(START_ROW, ws.max_row + 1):
        value = to_float(ws[f"{col_letter}{row}"].value)

        if value is not None:
            rows.append({
                "row": row,
                "value": value
            })

    return rows


def is_match(left_value, right_value):
    """
    Values match if their absolute difference is within threshold.
    """
    return abs(left_value - right_value) <= THRESHOLD


def pair_score(left_item, right_item):
    """
    Score for aligning one left value with one right value.
    """
    if is_match(left_item["value"], right_item["value"]):
        return MATCH_SCORE

    return MISMATCH_SCORE


# =========================
# Dynamic Programming
# =========================

def align_with_dp(left, right):
    """
    Full DP sequence alignment.

    left  = values from column D
    right = values from column F

    DP choices:
    1. DIAG = align left[i] with right[j]
    2. UP   = left[i] has no matching right value
    3. LEFT = right[j] has no matching left value
    """

    n = len(left)
    m = len(right)

    dp = [[0] * (m + 1) for _ in range(n + 1)]
    trace = [[None] * (m + 1) for _ in range(n + 1)]

    # First column: left values aligned with blanks
    for i in range(1, n + 1):
        dp[i][0] = dp[i - 1][0] + GAP_SCORE
        trace[i][0] = "UP"

    # First row: right values aligned with blanks
    for j in range(1, m + 1):
        dp[0][j] = dp[0][j - 1] + GAP_SCORE
        trace[0][j] = "LEFT"

    # Fill DP table
    for i in range(1, n + 1):
        for j in range(1, m + 1):
            left_item = left[i - 1]
            right_item = right[j - 1]

            diag = dp[i - 1][j - 1] + pair_score(left_item, right_item)
            up = dp[i - 1][j] + GAP_SCORE
            left_gap = dp[i][j - 1] + GAP_SCORE

            best = max(diag, up, left_gap)
            dp[i][j] = best

            # Prefer real matches.
            # Avoid mismatches unless they are truly best.
            if best == diag and is_match(left_item["value"], right_item["value"]):
                trace[i][j] = "DIAG"
            elif best == up:
                trace[i][j] = "UP"
            elif best == left_gap:
                trace[i][j] = "LEFT"
            else:
                trace[i][j] = "DIAG"

    # Backtrack from bottom-right
    alignment = []
    i = n
    j = m

    while i > 0 or j > 0:
        direction = trace[i][j]

        if direction == "DIAG":
            left_item = left[i - 1]
            right_item = right[j - 1]

            alignment.append({
                "left": left_item,
                "right": right_item,
                "matched": is_match(left_item["value"], right_item["value"])
            })

            i -= 1
            j -= 1

        elif direction == "UP":
            left_item = left[i - 1]

            alignment.append({
                "left": left_item,
                "right": None,
                "matched": False
            })

            i -= 1

        elif direction == "LEFT":
            right_item = right[j - 1]

            alignment.append({
                "left": None,
                "right": right_item,
                "matched": False
            })

            j -= 1

        else:
            break

    alignment.reverse()
    return alignment


# =========================
# Output Sheet
# =========================

def create_output_sheet(wb, src_ws):
    """
    Create a new output sheet for aligned results.
    Original sheet is not modified.
    """
    if OUTPUT_SHEET_NAME in wb.sheetnames:
        del wb[OUTPUT_SHEET_NAME]

    out_ws = wb.create_sheet(OUTPUT_SHEET_NAME)

    # Copy left headers A:D
    copy_block(
        src_ws,
        out_ws,
        HEADER_ROW,
        HEADER_ROW,
        LEFT_BLOCK_START_COL,
        LEFT_BLOCK_END_COL,
        LEFT_OUTPUT_START_COL
    )

    # Leave E blank because difference column is ignored
    out_ws["E1"] = ""

    # Copy right headers F:H
    copy_block(
        src_ws,
        out_ws,
        HEADER_ROW,
        HEADER_ROW,
        RIGHT_BLOCK_START_COL,
        RIGHT_BLOCK_END_COL,
        RIGHT_OUTPUT_START_COL
    )

    # Copy column widths from source sheet
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width is not None:
            out_ws.column_dimensions[col_letter].width = dim.width

    return out_ws


def write_alignment(src_ws, out_ws, alignment):
    """
    Write aligned full-row blocks into the output sheet.

    Left block:
        A:D

    Right block:
        F:H

    Column E is ignored.
    """
    output_row = START_ROW

    for step in alignment:
        left_item = step["left"]
        right_item = step["right"]
        matched = step["matched"]

        # Copy left row block A:D
        if left_item is not None:
            copy_block(
                src_ws,
                out_ws,
                left_item["row"],
                output_row,
                LEFT_BLOCK_START_COL,
                LEFT_BLOCK_END_COL,
                LEFT_OUTPUT_START_COL
            )

        # Copy right row block F:H
        if right_item is not None:
            copy_block(
                src_ws,
                out_ws,
                right_item["row"],
                output_row,
                RIGHT_BLOCK_START_COL,
                RIGHT_BLOCK_END_COL,
                RIGHT_OUTPUT_START_COL
            )

        # Highlight unmatched row blocks
        if not matched:
            if left_item is not None:
                highlight_block(
                    out_ws,
                    output_row,
                    LEFT_OUTPUT_START_COL,
                    LEFT_BLOCK_END_COL
                )

            if right_item is not None:
                highlight_block(
                    out_ws,
                    output_row,
                    RIGHT_OUTPUT_START_COL,
                    RIGHT_BLOCK_END_COL
                )

        output_row += 1


# =========================
# Main
# =========================

def main():
    wb = load_workbook(INPUT_FILE)

    if INPUT_SHEET_NAME in wb.sheetnames:
        src_ws = wb[INPUT_SHEET_NAME]
    else:
        src_ws = wb.active
        print(f"Warning: sheet '{INPUT_SHEET_NAME}' was not found.")
        print(f"Using active sheet instead: '{src_ws.title}'")

    left = read_column(src_ws, LEFT_INPUT_COL)
    right = read_column(src_ws, RIGHT_INPUT_COL)

    alignment = align_with_dp(left, right)

    out_ws = create_output_sheet(wb, src_ws)
    write_alignment(src_ws, out_ws, alignment)

    wb.save(OUTPUT_FILE)

    print("Done:", OUTPUT_FILE)
    print("Input sheet:", src_ws.title)
    print("Output sheet:", OUTPUT_SHEET_NAME)
    print("Left values:", len(left))
    print("Right values:", len(right))
    print("Alignment steps:", len(alignment))
    print("Matches written:", sum(1 for step in alignment if step["matched"]))


if __name__ == "__main__":
    main()