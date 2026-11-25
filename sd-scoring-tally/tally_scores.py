#!/usr/bin/env python3
"""
Tally SD Individual Scoring matrices.

Usage:
    python3 tally_scores.py "/path/to/folder/with/xlsx/files"

Reads every .xlsx in the folder (skipping Excel lock files and any previous
output), adds up each score cell in the Important Things grid J5:W12, and writes
SD_Scoring_TOTALS.xlsx into the same folder with three sheets:

  Totals   - the matrix in its original layout; J5:W12 = sum of everyone's scores
             (other sections are left as they are in the first file)
  Averages - same layout; J5:W12 = total / number of files (blanks count as 0)
  Checks   - anything odd: scores other than 0/1/2, blank boxes, mismatched labels
"""
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.utils import column_index_from_string, get_column_letter

# ---- Grid settings (change these if the template layout changes) ----
FIRST_ROW, LAST_ROW = 5, 12
FIRST_COL, LAST_COL = "J", "W"   # Important Things only
HEADER_ROW = 3          # column headings
OUTPUT_NAME = "SD_Scoring_TOTALS.xlsx"
VALID_SCORES = {0, 1, 2}

C1 = column_index_from_string(FIRST_COL)
C2 = column_index_from_string(LAST_COL)


def grid_cells():
    for r in range(FIRST_ROW, LAST_ROW + 1):
        for c in range(C1, C2 + 1):
            yield r, c


def main(folder):
    folder = Path(folder).expanduser()
    files = sorted(
        f for f in folder.glob("*.xlsx")
        if not f.name.startswith("~$") and f.name != OUTPUT_NAME
    )
    if not files:
        sys.exit(f"No .xlsx files found in {folder}")

    totals = {cell: 0 for cell in grid_cells()}
    responses = {cell: 0 for cell in grid_cells()}   # how many people scored the box
    issues = []
    template_headers = None

    for f in files:
        ws = load_workbook(f, data_only=True).worksheets[0]
        headers = [ws.cell(HEADER_ROW, c).value for c in range(C1, C2 + 1)]
        if template_headers is None:
            template_headers = headers
        elif headers != template_headers:
            issues.append((f.name, "-", "Column headings differ from first file - check layout"))

        for (r, c) in grid_cells():
            v = ws.cell(r, c).value
            ref = f"{get_column_letter(c)}{r}"
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            try:
                n = float(v)
                n = int(n) if n.is_integer() else n
            except (TypeError, ValueError):
                issues.append((f.name, ref, f"Not a number: {v!r} (ignored)"))
                continue
            if n not in VALID_SCORES:
                issues.append((f.name, ref, f"Score {n} is not 0/1/2 (included in total)"))
            totals[(r, c)] += n
            responses[(r, c)] += 1

    # Boxes some people filled and others left blank
    for (r, c), cnt in responses.items():
        if 0 < cnt < len(files):
            issues.append(("(various)", f"{get_column_letter(c)}{r}",
                           f"Only {cnt} of {len(files)} people scored this box (blanks counted as 0)"))

    # ---- Output: reuse first file as the template so layout/formatting is kept ----
    out_wb = load_workbook(files[0])
    ws = out_wb.worksheets[0]
    ws.title = "Totals"
    for (r, c) in grid_cells():
        ws.cell(r, c).value = totals[(r, c)]
    ws["B1"] = f"TOTALS from {len(files)} files"
    ws["B1"].font = Font(bold=True)
    for cell in ("O1", "P1"):          # stray values in the template
        ws[cell] = None

    # Averages sheet: same layout, each box = total / number of files (blanks = 0)
    av = out_wb.copy_worksheet(ws)
    av.title = "Averages"
    av["B1"] = f"AVERAGES from {len(files)} files (total / number of people; blanks = 0)"
    av["B1"].font = Font(bold=True)
    for (r, c) in grid_cells():
        cell = av.cell(r, c)
        cell.value = round(totals[(r, c)] / len(files), 2)   # blanks count as 0
        cell.number_format = "0.00"

    bold = Font(bold=True)
    # Checks sheet
    ck = out_wb.create_sheet("Checks")
    for i, h in enumerate(("File", "Cell", "Issue"), start=1):
        ck.cell(1, i, h).font = bold
    if not issues:
        ck["A2"] = "No issues found"
    for i, (f, ref, msg) in enumerate(issues, start=2):
        ck.cell(i, 1, f); ck.cell(i, 2, ref); ck.cell(i, 3, msg)
    ck.column_dimensions["A"].width = 45
    ck.column_dimensions["C"].width = 60

    out_path = folder / OUTPUT_NAME
    out_wb.save(out_path)
    print(f"Tallied {len(files)} files -> {out_path}")
    if issues:
        print(f"{len(issues)} issue(s) - see the Checks sheet:")
        for f, ref, msg in issues:
            print(f"  {f} {ref}: {msg}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
