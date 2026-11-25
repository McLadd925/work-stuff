#!/usr/bin/env python3
"""
Tally SD Individual Scoring matrices.

Usage:
    python3 tally_scores.py "/path/to/folder/with/xlsx/files"

Reads every .xlsx in the folder (skipping Excel lock files and any previous
output), adds up each score cell in the Important Things grid J5:W12, and writes
SD_Scoring_TOTALS.xlsx into the same folder with these sheets:

  Totals   - the matrix in its original layout; J5:W12 = sum of everyone's scores
             (other sections are left as they are in the first file)
  Averages - same layout; J5:W12 = total / number of files (blanks count as 0)
  Symbols  - averages shown as symbols: high impact / some impact / support (limits below)
  Alignment- Important Things with no high-impact target, target impact, and disagreements

Scores are impact scores: 2 = high impact, 1 = some impact, 0 = support.
  Checks   - anything odd: scores other than 0/1/2, blank boxes, mismatched labels
"""
import re
import sys
import uuid
import zipfile
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter

# ---- Grid settings (change these if the template layout changes) ----
FIRST_ROW, LAST_ROW = 5, 12
FIRST_COL, LAST_COL = "J", "W"   # Important Things only
HEADER_ROW = 3          # column headings
OUTPUT_NAME = "SD_Scoring_TOTALS.xlsx"
VALID_SCORES = {0, 1, 2}
TARGET_COL = "AE"       # improvement target text for each row

# Symbol bands for averages (change the limits here)
HIGH_MIN = 1.5          # average >= 1.5        -> high impact
SOME_MIN = 0.75         # 0.75 <= average < 1.5 -> some impact, below -> support
SYM_HIGH, SYM_SOME, SYM_SUPPORT = "\u25cf", "\u25d0", "\u25cb"   # ● ◐ ○

C1 = column_index_from_string(FIRST_COL)
C2 = column_index_from_string(LAST_COL)


def grid_cells():
    for r in range(FIRST_ROW, LAST_ROW + 1):
        for c in range(C1, C2 + 1):
            yield r, c


def band(avg):
    if avg >= HIGH_MIN:
        return "high"
    if avg >= SOME_MIN:
        return "some"
    return "support"


SYMBOL = {"high": SYM_HIGH, "some": SYM_SOME, "support": SYM_SUPPORT}


def header_text(ws, col):
    v = ws.cell(HEADER_ROW, col).value
    return str(v).strip() if v else get_column_letter(col)


def target_text(ws, row):
    t = ws[f"{TARGET_COL}{row}"].value
    if not t:
        return f"Row {row}"
    lines = [x.strip() for x in str(t).split("\n") if x.strip()]
    first = lines[0] if lines else f"Row {row}"
    return first + (f"  (+{len(lines) - 1} more)" if len(lines) > 1 else "")


def add_icon_set(xlsx_path, sheet_name, sqref):
    """Add the template's custom icon set (empty circle / half circle / black
    circle) to a sheet. openpyxl can't write custom icon sets, so this edits
    the sheet XML inside the saved file directly."""
    with zipfile.ZipFile(xlsx_path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    # find the sheet's XML file from workbook.xml + its relationships
    wb_xml = parts["xl/workbook.xml"].decode()
    rid = re.search(r'<sheet[^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheet_name), wb_xml).group(1)
    rels = parts["xl/_rels/workbook.xml.rels"].decode()
    target = re.search(r'<Relationship[^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, rels)
    if not target:
        target = re.search(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, rels)
    path = target.group(1).lstrip("/")
    path = path if path.startswith("xl/") else "xl/" + path

    ext = (
        '<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" '
        'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main">'
        '<x14:conditionalFormattings>'
        '<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">'
        f'<x14:cfRule type="iconSet" priority="1" id="{{{str(uuid.uuid4()).upper()}}}">'
        '<x14:iconSet showValue="0" custom="1">'
        '<x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo>'
        f'<x14:cfvo type="num"><xm:f>{SOME_MIN}</xm:f></x14:cfvo>'
        f'<x14:cfvo type="num"><xm:f>{HIGH_MIN}</xm:f></x14:cfvo>'
        '<x14:cfIcon iconSet="5Quarters" iconId="0"/>'
        '<x14:cfIcon iconSet="5Quarters" iconId="2"/>'
        '<x14:cfIcon iconSet="4TrafficLights" iconId="0"/>'
        '</x14:iconSet></x14:cfRule>'
        f'<xm:sqref>{sqref}</xm:sqref>'
        '</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>'
    )
    sheet_xml = parts[path].decode()
    sheet_xml = sheet_xml.replace("</worksheet>", ext + "</worksheet>")
    parts[path] = sheet_xml.encode()

    tmp = Path(str(xlsx_path) + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)
    tmp.replace(xlsx_path)


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
    scores = {cell: [] for cell in grid_cells()}     # each person's score (blank = 0)
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
                scores[(r, c)].append(0)
                continue
            try:
                n = float(v)
                n = int(n) if n.is_integer() else n
            except (TypeError, ValueError):
                issues.append((f.name, ref, f"Not a number: {v!r} (counted as 0)"))
                scores[(r, c)].append(0)
                continue
            if n not in VALID_SCORES:
                issues.append((f.name, ref, f"Score {n} is not 0/1/2 (included in total)"))
            totals[(r, c)] += n
            responses[(r, c)] += 1
            scores[(r, c)].append(n)

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
    for cell in ("O1", "P1", "Q16", "R16", "S16", "T16"):  # stray values / old legend
        ws[cell] = None

    # Averages sheet: same layout, each box = total / number of files (blanks = 0)
    av = out_wb.copy_worksheet(ws)
    av.title = "Averages"
    av["B1"] = f"AVERAGES from {len(files)} files (total / number of people; blanks = 0)"
    av["B1"].font = Font(bold=True)
    avg = {cell: totals[cell] / len(files) for cell in grid_cells()}   # blanks count as 0
    for (r, c) in grid_cells():
        cell = av.cell(r, c)
        cell.value = round(avg[(r, c)], 2)
        cell.number_format = "0.00"

    # Symbols sheet: averages turned back into symbols
    sy = out_wb.copy_worksheet(ws)
    sy.title = "Symbols"
    sy["B1"] = (f"SYMBOLS from averages:  {SYM_HIGH} high impact (>= {HIGH_MIN})   "
                f"{SYM_SOME} some impact ({SOME_MIN} to <{HIGH_MIN})   "
                f"{SYM_SUPPORT} support (< {SOME_MIN})   - adjust via Conditional Formatting > Manage Rules")
    sy["B1"].font = Font(bold=True)
    for (r, c) in grid_cells():
        cell = sy.cell(r, c)
        cell.value = round(avg[(r, c)], 2)   # icons added by add_icon_set() after saving
        cell.number_format = "0.00"
        cell.alignment = Alignment(horizontal="center", vertical="center")

    bold = Font(bold=True)
    title = Font(bold=True, size=13)
    red = PatternFill("solid", fgColor="F8CBAD")
    amber = PatternFill("solid", fgColor="FFE699")
    green = PatternFill("solid", fgColor="C6E0B4")
    rows = range(FIRST_ROW, LAST_ROW + 1)
    cols = range(C1, C2 + 1)
    n_cols = len(cols)

    # Alignment sheet
    al = out_wb.create_sheet("Alignment")
    al["A1"] = f"Alignment check - {len(files)} files, bands: high impact >= {HIGH_MIN}, some impact >= {SOME_MIN}, below = support"
    al["A1"].font = title
    row = 3

    def header(cells):
        nonlocal row
        for i, h in enumerate(cells, start=1):
            al.cell(row, i, h).font = bold
        row += 1

    # 1. Important things coverage
    al.cell(row, 1, "1. Does every Important Thing have at least one high-impact target? (ranked by total impact)").font = title
    row += 1
    header(("Rank", "Important Thing", "Total impact", "High impact", "Some impact", "Status"))
    col_stats = []
    for c in cols:
        bands = [band(avg[(r, c)]) for r in rows]
        col_stats.append((header_text(ws, c), sum(avg[(r, c)] for r in rows),
                          bands.count("high"), bands.count("some")))
    col_stats.sort(key=lambda x: -x[1])
    for rank, (name, tot, st, mo) in enumerate(col_stats, start=1):
        if st:
            status, fill = "Has high-impact target(s)", green
        elif mo:
            status, fill = "Some impact only - no high-impact target", amber
        else:
            status, fill = "Support only - no target has real impact", red
        for i, v in enumerate((rank, name, round(tot, 2), st, mo, status), start=1):
            al.cell(row, i, v)
        al.cell(row, 6).fill = fill
        row += 1

    # 2. Targets
    row += 1
    al.cell(row, 1, "2. How much impact does each improvement target have on the Important Things?").font = title
    row += 1
    header(("Row", "Improvement target (first line)", "Total impact", "High impact", "Some impact", "Status"))
    for r in rows:
        bands = [band(avg[(r, c)]) for c in cols]
        st, mo = bands.count("high"), bands.count("some")
        tot = sum(avg[(r, c)] for c in cols)
        if st == 0 and mo == 0:
            status, fill = "Support only - drop or rewrite?", red
        elif st == 0:
            status, fill = "Some impact only - no high impact", amber
        elif st >= 0.7 * n_cols:
            status, fill = "High impact on almost everything - too broad?", amber
        else:
            status, fill = "OK", green
        for i, v in enumerate((r, target_text(ws, r), round(tot, 2), st, mo, status), start=1):
            al.cell(row, i, v)
        al.cell(row, 6).fill = fill
        row += 1

    # 3. Disagreement
    row += 1
    al.cell(row, 1, "3. Where did people disagree? (someone scored 0 = support and someone scored 2 = high impact)").font = title
    row += 1
    header(("Cell", "Important Thing", "Target row", "Average", "Scores given", ""))
    found = False
    for (r, c) in grid_cells():
        sc = scores[(r, c)]
        if sc and max(sc) - min(sc) >= 2:
            found = True
            for i, v in enumerate((f"{get_column_letter(c)}{r}", header_text(ws, c), r,
                                   round(avg[(r, c)], 2), ", ".join(str(x) for x in sc)), start=1):
                al.cell(row, i, v)
            row += 1
    if not found:
        al.cell(row, 1, "No boxes with a 0-vs-2 split")
        row += 1

    al.column_dimensions["A"].width = 8
    al.column_dimensions["B"].width = 60
    for col in "CDE":
        al.column_dimensions[col].width = 16
    al.column_dimensions["F"].width = 42

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
    add_icon_set(out_path, "Symbols", f"{FIRST_COL}{FIRST_ROW}:{LAST_COL}{LAST_ROW}")
    print(f"Tallied {len(files)} files -> {out_path}")
    if issues:
        print(f"{len(issues)} issue(s) - see the Checks sheet:")
        for f, ref, msg in issues:
            print(f"  {f} {ref}: {msg}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    main(sys.argv[1])
