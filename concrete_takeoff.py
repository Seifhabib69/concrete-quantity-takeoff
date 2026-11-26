#!/usr/bin/env python3
"""
concrete-takeoff: a small, auditable concrete quantity takeoff tool.

Reads an Excel file listing footings, columns (necks), beams (tie/strap beams)
and slabs, and writes a report workbook with live Excel formulas, so every
number can be checked by hand.

Usage:
    python concrete_takeoff.py template my_project.xlsx      # blank input file
    python concrete_takeoff.py run my_project.xlsx report.xlsx

All dimensions are in metres, volumes in cubic metres (m3).
Volumes are GROSS (b x t x h x count): no deductions for overlaps between
members, rebar, or openings. State this in your submittal if it matters.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

__version__ = "0.1.0"

# --------------------------------------------------------------------------
# Input definition
# --------------------------------------------------------------------------
# Every member type is "count x d1 x d2 x d3".
MEMBER_SHEETS = {
    "Footings": ["Mark", "Count", "Length (m)", "Width (m)", "Thickness (m)", "Notes"],
    "Columns": ["Mark", "Count", "b (m)", "t (m)", "Height (m)", "Notes"],
    "Beams": ["Mark", "Count", "b (m)", "Depth (m)", "Length (m)", "Notes"],
    "Slabs": ["Mark", "Count", "Length (m)", "Width (m)", "Thickness (m)", "Notes"],
}

# Order matters: the report formulas point at fixed cells (Settings!B2..B5).
SETTINGS_ORDER = ["waste_pct", "plain_concrete", "pc_thickness_m", "pc_offset_m"]
DEFAULT_SETTINGS = {
    "waste_pct": 0.0,
    "plain_concrete": "no",
    "pc_thickness_m": 0.10,
    "pc_offset_m": 0.10,
}
SETTINGS_HELP = {
    "waste_pct": "Concrete waste / allowance as a fraction (3% -> 0.03). Applied to every line.",
    "plain_concrete": "yes/no: add plain concrete (PC, blinding) under footings.",
    "pc_thickness_m": "PC thickness in metres (used only if plain_concrete = yes).",
    "pc_offset_m": "PC extension beyond each footing edge, metres (only if plain_concrete = yes).",
}

EXAMPLE_ROWS = {
    "Footings": ["F1 (example)", 4, 1.80, 1.80, 0.50, "delete or overwrite me"],
    "Columns": ["C1 (example)", 4, 0.30, 0.60, 3.00, "neck, footing top to ground beam"],
    "Beams": ["TB1 (example)", 6, 0.30, 0.60, 4.20, "clear length between columns"],
    "Slabs": ["S1 (example)", 1, 5.00, 4.00, 0.15, ""],
}


class InputError(ValueError):
    """Raised with a readable list of every problem found in the input file."""


# --------------------------------------------------------------------------
# Core calculations (pure functions, easy to test)
# --------------------------------------------------------------------------
def member_volume(count: float, d1: float, d2: float, d3: float) -> float:
    """Gross volume of `count` identical prismatic members."""
    return count * d1 * d2 * d3


def plain_concrete_volume(count, length, width, thickness, offset) -> float:
    """PC under footings: footing plan enlarged by `offset` on every side."""
    return count * (length + 2 * offset) * (width + 2 * offset) * thickness


def calculate(data: dict, settings: dict) -> dict:
    """Return net volumes per group, in m3, plus waste-adjusted totals."""
    net = {}
    for sheet in MEMBER_SHEETS:
        net[sheet] = sum(
            member_volume(r["count"], r["d1"], r["d2"], r["d3"]) for r in data[sheet]
        )
    net["Reinforced concrete total"] = sum(net[s] for s in MEMBER_SHEETS)

    if settings["plain_concrete"] == "yes":
        net["Plain concrete"] = sum(
            plain_concrete_volume(
                r["count"], r["d1"], r["d2"], settings["pc_thickness_m"],
                settings["pc_offset_m"],
            )
            for r in data["Footings"]
        )
    else:
        net["Plain concrete"] = 0.0

    k = 1 + settings["waste_pct"]
    return {name: {"net": v, "with_waste": v * k} for name, v in net.items()}


# --------------------------------------------------------------------------
# Reading input
# --------------------------------------------------------------------------
def _is_number(x) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def _read_settings(wb) -> dict:
    settings = dict(DEFAULT_SETTINGS)
    errors = []
    if "Settings" in wb.sheetnames:
        for row in wb["Settings"].iter_rows(min_row=2, values_only=True):
            if not row or row[0] not in settings or row[1] is None:
                continue
            settings[row[0]] = row[1]

    pc = str(settings["plain_concrete"]).strip().lower()
    if pc in ("yes", "y", "true", "1"):
        settings["plain_concrete"] = "yes"
    elif pc in ("no", "n", "false", "0"):
        settings["plain_concrete"] = "no"
    else:
        errors.append(f"Settings: plain_concrete must be yes or no (got {pc!r})")

    for key in ("waste_pct", "pc_thickness_m", "pc_offset_m"):
        v = settings[key]
        if not _is_number(v) or v < 0:
            errors.append(f"Settings: {key} must be a number >= 0 (got {v!r})")
    if _is_number(settings["waste_pct"]) and settings["waste_pct"] > 1:
        errors.append(
            "Settings: waste_pct must be a fraction between 0 and 1 "
            f"(3% -> 0.03), got {settings['waste_pct']!r}"
        )
    if errors:
        raise InputError("\n".join(errors))
    return settings


def _read_members(wb, sheet: str) -> list[dict]:
    if sheet not in wb.sheetnames:
        return []
    rows, errors = [], []
    for i, row in enumerate(wb[sheet].iter_rows(min_row=2, values_only=True), start=2):
        row = list(row) + [None] * 6
        mark, count, d1, d2, d3, notes = row[:6]
        if all(v is None for v in (mark, count, d1, d2, d3)):
            continue  # blank line
        row_ok = True
        for label, v in (("Count", count), ("dim 1", d1), ("dim 2", d2), ("dim 3", d3)):
            if not _is_number(v) or v <= 0:
                row_ok = False
                errors.append(
                    f"{sheet} row {i} ({mark or 'no mark'}): {label} must be a "
                    f"positive number (got {v!r})"
                )
        if row_ok:
            rows.append(
                {"mark": str(mark or f"row {i}"), "count": count,
                 "d1": d1, "d2": d2, "d3": d3, "notes": notes or ""}
            )
    if errors:
        raise InputError("\n".join(errors))
    return rows


def read_input(path) -> tuple[dict, dict]:
    wb = load_workbook(path, data_only=True)
    settings = _read_settings(wb)
    data, errors = {}, []
    for sheet in MEMBER_SHEETS:
        try:
            data[sheet] = _read_members(wb, sheet)
        except InputError as e:
            errors.append(str(e))
    if errors:
        raise InputError("\n".join(errors))
    if not any(data.values()):
        raise InputError("No members found. Fill at least one row in Footings/Columns/Beams/Slabs.")
    return data, settings


# --------------------------------------------------------------------------
# Writing Excel files
# --------------------------------------------------------------------------
FONT = "Arial"
F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_INPUT = Font(name=FONT, size=10, color="0000FF")  # blue = typed-in input
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=14, bold=True)
FILL_HEAD = PatternFill("solid", start_color="1F3864")
FILL_TOTAL = PatternFill("solid", start_color="D9E1F2")
FILL_KEY = PatternFill("solid", start_color="FFFF00")
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NF_DIM, NF_VOL, NF_PCT = "0.00", "#,##0.00", "0.0%"


def _header(ws, row, labels):
    for c, text in enumerate(labels, start=1):
        cell = ws.cell(row=row, column=c, value=text)
        cell.font, cell.fill, cell.border = F_HEAD, FILL_HEAD, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_settings_sheet(wb, settings):
    ws = wb.create_sheet("Settings")
    _header(ws, 1, ["Setting", "Value", "Description"])
    for r, key in enumerate(SETTINGS_ORDER, start=2):
        ws.cell(row=r, column=1, value=key).font = F_BASE
        v = ws.cell(row=r, column=2, value=settings[key])
        v.font, v.fill = F_INPUT, FILL_KEY
        v.number_format = NF_PCT if key == "waste_pct" else NF_DIM
        ws.cell(row=r, column=3, value=SETTINGS_HELP[key]).font = F_BASE
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).border = BORDER
    _widths(ws, [18, 12, 80])
    return ws


def make_template(path) -> None:
    """Blank input workbook with a legend and one example row per sheet."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Instructions"
    lines = [
        ("Concrete takeoff: input file", F_TITLE),
        ("", F_BASE),
        ("1. Fill the sheets Footings, Columns, Beams, Slabs (one row per group of identical members).", F_BASE),
        ("2. Blue text on yellow = values you edit. Dimensions in metres.", F_BASE),
        ("3. Delete the (example) rows or overwrite them with your own.", F_BASE),
        ("4. Optional: edit the Settings sheet (waste %, plain concrete under footings).", F_BASE),
        ("5. Run:  python concrete_takeoff.py run this_file.xlsx report.xlsx", F_BASE),
        ("", F_BASE),
        ("Volumes are gross: count x dim1 x dim2 x dim3. No deduction for overlaps or rebar.", F_BASE),
    ]
    for r, (text, font) in enumerate(lines, start=1):
        ws.cell(row=r, column=1, value=text).font = font
    ws.column_dimensions["A"].width = 100

    _write_settings_sheet(wb, DEFAULT_SETTINGS)
    for sheet, headers in MEMBER_SHEETS.items():
        s = wb.create_sheet(sheet)
        _header(s, 1, headers)
        for c, v in enumerate(EXAMPLE_ROWS[sheet], start=1):
            cell = s.cell(row=2, column=c, value=v)
            cell.font, cell.border = F_INPUT, BORDER
            if c >= 3 and c <= 5:
                cell.number_format = NF_DIM
        s.freeze_panes = "A2"
        _widths(s, [18, 8, 13, 13, 14, 40])
    wb.save(path)


def write_report(data: dict, settings: dict, path) -> None:
    wb = Workbook()
    summ = wb.active
    summ.title = "Summary"
    _write_settings_sheet(wb, settings)

    last_row = {}
    for sheet, headers in MEMBER_SHEETS.items():
        ws = wb.create_sheet(sheet)
        cols = headers[:5] + ["Unit volume (m3)", "Total volume (m3)", "Notes"]
        if sheet == "Footings":
            cols.append("Plain concrete (m3)")
        _header(ws, 1, cols)
        for i, r in enumerate(data[sheet], start=2):
            ws.cell(row=i, column=1, value=r["mark"])
            ws.cell(row=i, column=2, value=r["count"])
            for c, key in zip((3, 4, 5), ("d1", "d2", "d3")):
                ws.cell(row=i, column=c, value=r[key]).number_format = NF_DIM
            for c in range(1, 6):
                ws.cell(row=i, column=c).font = F_INPUT
            ws.cell(row=i, column=6, value=f"=C{i}*D{i}*E{i}").number_format = NF_VOL
            ws.cell(row=i, column=7, value=f"=B{i}*F{i}").number_format = NF_VOL
            ws.cell(row=i, column=8, value=r["notes"])
            if sheet == "Footings":
                ws.cell(row=i, column=9, value=(
                    f'=IF(Settings!$B$3="yes",B{i}*(C{i}+2*Settings!$B$5)'
                    f"*(D{i}+2*Settings!$B$5)*Settings!$B$4,0)"
                )).number_format = NF_VOL
            for c in range(1, len(cols) + 1):
                cell = ws.cell(row=i, column=c)
                cell.border = BORDER
                if cell.font != F_INPUT:
                    cell.font = F_BASE
        n = len(data[sheet])
        last_row[sheet] = max(2, n + 1)
        tot = n + 2
        ws.cell(row=tot, column=1, value="Total").font = F_BOLD
        for col in (7, 9) if sheet == "Footings" else (7,):
            L = get_column_letter(col)
            c = ws.cell(row=tot, column=col, value=f"=SUM({L}2:{L}{last_row[sheet]})")
            c.font, c.number_format = F_BOLD, NF_VOL
        for c in range(1, len(cols) + 1):
            ws.cell(row=tot, column=c).fill = FILL_TOTAL
            ws.cell(row=tot, column=c).border = BORDER
        ws.freeze_panes = "A2"
        _widths(ws, [18, 8, 13, 13, 14, 16, 17, 40, 18])

    # ---- Summary sheet ----
    summ["A1"] = "Concrete quantity takeoff: summary"
    summ["A1"].font = F_TITLE
    _header(summ, 3, ["Element", "Net volume (m3)", "Waste", "Volume incl. waste (m3)"])
    for r, sheet in enumerate(MEMBER_SHEETS, start=4):
        summ.cell(row=r, column=1, value=sheet)
        summ.cell(row=r, column=2, value=f"=SUM({sheet}!G2:G{last_row[sheet]})")
    summ.cell(row=8, column=1, value="Reinforced concrete total")
    summ.cell(row=8, column=2, value="=SUM(B4:B7)")
    summ.cell(row=9, column=1, value="Plain concrete under footings")
    summ.cell(row=9, column=2, value=f"=SUM(Footings!I2:I{last_row['Footings']})")
    for r in range(4, 10):
        summ.cell(row=r, column=3, value="=Settings!$B$2").number_format = NF_PCT
        summ.cell(row=r, column=4, value=f"=B{r}*(1+C{r})")
        for c in (2, 4):
            summ.cell(row=r, column=c).number_format = NF_VOL
        for c in range(1, 5):
            cell = summ.cell(row=r, column=c)
            cell.border = BORDER
            cell.font = F_BOLD if r == 8 else F_BASE
            if r == 8:
                cell.fill = FILL_TOTAL
    notes = [
        "Assumptions",
        "- Volume = count x dim1 x dim2 x dim3 for every member (gross, in m3).",
        "- No deduction for overlaps (columns in footings/slabs, beams in columns), rebar or openings.",
        "- Blue cells in the element sheets are inputs copied from your file; black cells are formulas.",
        "- Waste % and plain-concrete settings come from the Settings sheet.",
        "- Plain concrete = footing plan enlarged by the offset on every side x PC thickness.",
    ]
    for i, text in enumerate(notes):
        c = summ.cell(row=11 + i, column=1, value=text)
        c.font = F_BOLD if i == 0 else F_BASE
    _widths(summ, [34, 18, 10, 24])
    wb.save(path)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def print_summary(result: dict) -> None:
    print(f"{'Element':<32}{'Net m3':>12}{'incl. waste m3':>18}")
    print("-" * 62)
    for name, v in result.items():
        print(f"{name:<32}{v['net']:>12.2f}{v['with_waste']:>18.2f}")


def main(argv=None) -> int:
    p = argparse.ArgumentParser(description="Concrete quantity takeoff (Excel in, Excel out)")
    sub = p.add_subparsers(dest="cmd", required=True)
    t = sub.add_parser("template", help="write a blank input workbook")
    t.add_argument("output")
    r = sub.add_parser("run", help="compute quantities and write the report")
    r.add_argument("input")
    r.add_argument("output")
    args = p.parse_args(argv)

    if args.cmd == "template":
        make_template(args.output)
        print(f"Template written to {args.output}")
        return 0
    try:
        data, settings = read_input(args.input)
    except InputError as e:
        print("Input problems found:\n" + str(e), file=sys.stderr)
        return 1
    result = calculate(data, settings)
    write_report(data, settings, args.output)
    print_summary(result)
    print(f"\nReport written to {Path(args.output)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
