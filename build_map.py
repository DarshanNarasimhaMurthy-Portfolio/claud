#!/usr/bin/env python3
"""Build Consumables_Map.xlsx, the consumables floor map for GXO Sherburn.

Run this OFF the work machine (it needs Python + openpyxl + PyYAML):

    python build_map.py                  # normal build from layout.yaml
    python build_map.py --demo           # build with dummy items/pallets for testing
    python build_map.py -l other.yaml -o Test.xlsx

The workbook works without macros (dropdowns on Locations, formula-driven
Map, conditional formatting). Paste FloorMap_VBA.txt in for click-to-move,
Find Space, Rebuild Layout and full repainting (see SETUP.md).

The map geometry here MUST match ComputeLayout in FloorMap_VBA.txt.
"""
import argparse
import math
import re
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as col_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

FONT = "Arial"
TABLE_STYLE = "TableStyleMedium15"  # black header, light grey banding

# Map sheet fixed rows (the VBA uses the same numbers)
ROW_TITLE, ROW_INFO, ROW_BUTTONS, ROW_BAY, ROW_BAY_GAP, ROW_AREA_TOP = 1, 2, 3, 4, 5, 6
TITLE_PX, INFO_PX, BUTTONS_PX = 32, 18, 28
LEGEND_COLS = 10

# Rows where each table header sits
HDR_ROW = 4
ZONE_SLOTS = 11          # zone columns in the automatic dropdown lists on Config
DD_ROWS = 150            # max items per dropdown list
DASH_ITEM_ROWS = 60      # item rows on the Dashboard
DASH_ZONE_ROWS = 10      # zone rows on the Dashboard
DASH_LOW_ROWS = 20       # low-stock list rows on the Dashboard

AISLE_COLS = ["AisleID", "Side", "Order", "PalletsAcross", "PalletsDeep", "Zone", "GapAfter", "Active"]
ZONE_COLS = ["Code", "Name", "Colour", "Covers"]
ITEM_INPUT_COLS = ["ItemCode", "Description", "Zone", "UnitsPerPallet", "OnHandUnits", "ReorderPallets", "Notes"]
ITEM_CALC_COLS = ["PalletsNeeded", "PalletsPlaced", "Mismatch", "LowStock"]
LOC_COLS = ["LocationID", "Aisle", "Depth", "Column", "Zone", "ItemCode", "UnitsOnPallet", "MapRow", "MapCol"]
LOG_COLS = ["Timestamp", "User", "ItemCode", "FromLocation", "ToLocation"]


# --------------------------------------------------------------------------- helpers
def hx(c):
    """'#c8a165' -> 'C8A165'"""
    return str(c).strip().lstrip("#").upper()


def solid(c):
    c = hx(c)
    return PatternFill(fill_type="solid", start_color=c, end_color=c)


def is_dark(c):
    c = hx(c)
    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 140


def colour_distance(a, b):
    a, b = hx(a), hx(b)
    return math.sqrt(sum((int(a[i:i + 2], 16) - int(b[i:i + 2], 16)) ** 2 for i in (0, 2, 4)))


def yes(v):
    return str(v).strip().lower() in ("yes", "y", "true", "1")


def f(size=10, bold=False, color="000000", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=hx(color), italic=italic)


def side(style, color):
    return Side(style=style, color=hx(color))


def tref(table, col):
    """Row-level structured reference, as Excel stores [@col] in the file."""
    return f"{table}[[#This Row],[{col}]]"


def add_name(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def title_bar(ws, text, subtitle, last_col, s):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    c = ws.cell(1, 1, text)
    c.font = f(16, True, s["HeaderText"])
    c.alignment = Alignment(vertical="center", indent=1)
    for col in range(1, last_col + 1):
        ws.cell(1, col).fill = solid(s["HeaderFill"])
        ws.cell(1, col).border = Border(bottom=side("thick", s["BrandAccent"]))
    ws.row_dimensions[1].height = 30
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    c = ws.cell(2, 1, subtitle)
    c.font = f(9, italic=True, color="595959")
    c.alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[2].height = 30
    ws.sheet_view.showGridLines = False


def section_label(ws, row, col, text, s, width=1):
    c = ws.cell(row, col, text)
    c.font = f(11, True, "000000")
    for k in range(width):
        ws.cell(row, col + k).border = Border(bottom=side("medium", s["BrandAccent"]))


def make_table(ws, name, top_row, left_col, headers, rows, calc=None, widths=None):
    """Write a proper Excel Table. calc = {header: formula-without-'='} for calculated columns."""
    calc = calc or {}
    for j, h in enumerate(headers):
        c = ws.cell(top_row, left_col + j, h)
        c.font = f(10, True, "FFFFFF")
        c.fill = solid("000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    n = max(1, len(rows))
    for i in range(n):
        row = rows[i] if i < len(rows) else {}
        for j, h in enumerate(headers):
            c = ws.cell(top_row + 1 + i, left_col + j)
            if h in calc:
                c.value = "=" + calc[h]
            else:
                c.value = row.get(h)
            c.font = f(10)
    ref = f"{col_letter(left_col)}{top_row}:{col_letter(left_col + len(headers) - 1)}{top_row + n}"
    t = Table(displayName=name, ref=ref)
    cols = []
    for j, h in enumerate(headers):
        tc = TableColumn(id=j + 1, name=h)
        if h in calc:
            tc.calculatedColumnFormula = TableFormula(attr_text=calc[h])
        cols.append(tc)
    t.tableColumns = cols
    t.tableStyleInfo = TableStyleInfo(name=TABLE_STYLE, showRowStripes=True)
    ws.add_table(t)
    if widths:
        for j, w in enumerate(widths):
            ws.column_dimensions[col_letter(left_col + j)].width = w
    ws.row_dimensions[top_row].height = 30
    return top_row + n  # last row


# --------------------------------------------------------------------------- geometry
def compute_layout(aisles, zones, walkway):
    """Work out where every pallet position sits on the Map.

    Fire exit at the bottom, loading bay at the top. Each side stacks its
    aisles upwards from the fire exit in Order. An aisle is a label row
    followed by PalletsAcross rows (column letters A, B, C... top to bottom)
    of PalletsDeep squares; depth 01 touches the main walkway.
    """
    act = [a for a in aisles if yes(a.get("Active", "Yes"))]
    left = sorted([a for a in act if str(a["Side"]).strip().lower().startswith("l")], key=lambda a: int(a["Order"]))
    right = sorted([a for a in act if str(a["Side"]).strip().lower().startswith("r")], key=lambda a: int(a["Order"]))
    lw = max([int(a["PalletsDeep"]) for a in left] or [0])
    rw = max([int(a["PalletsDeep"]) for a in right] or [0])

    def side_height(lst):
        return sum(1 + int(a["PalletsAcross"]) + int(a.get("GapAfter") or 0) for a in lst)

    legend_rows = 14 + len(zones)
    h = max(side_height(left), side_height(right), legend_rows)
    bottom = ROW_AREA_TOP + h - 1
    walk_start = 2 + lw
    walk_end = walk_start + walkway - 1
    right_start = walk_end + 1
    map_last_col = max(walk_end, right_start + rw - 1)
    legend_col = map_last_col + 2
    last_col = legend_col + LEGEND_COLS - 1

    positions, labels = [], []
    for side_code, lst in (("L", left), ("R", right)):
        cur = bottom
        for a in lst:
            across, deep = int(a["PalletsAcross"]), int(a["PalletsDeep"])
            top_data = cur - across + 1
            label_row = top_data - 1
            if side_code == "L":
                span = (walk_start - deep, walk_start - 1)
            else:
                span = (right_start, right_start + deep - 1)
            for k in range(across):
                letter = chr(65 + k)
                for d in range(1, deep + 1):
                    c = walk_start - d if side_code == "L" else right_start + d - 1
                    positions.append({
                        "LocationID": f"{a['AisleID']}-{d:02d}-{letter}",
                        "Aisle": str(a["AisleID"]), "Depth": d, "Column": letter,
                        "Zone": str(a.get("Zone") or ""),
                        "MapRow": top_data + k, "MapCol": c,
                    })
            labels.append((label_row, span, a))
            cur = label_row - 1 - int(a.get("GapAfter") or 0)
    # Stable order: aisles as listed, then depth, then column
    return {
        "positions": positions, "labels": labels, "h": h, "bottom": bottom,
        "walk_start": walk_start, "walk_end": walk_end, "right_start": right_start,
        "map_last_col": map_last_col, "legend_col": legend_col, "last_col": last_col,
        "fire_row": bottom + 2, "has_left": bool(left), "has_right": bool(right),
    }


def fit_to_screen(geo, s):
    """Pick the cell size (and zoom if needed) so Map + legend fits one screen."""
    cols, rows = geo["last_col"], geo["fire_row"] - ROW_BAY + 1
    fixed_h = TITLE_PX + INFO_PX + BUTTONS_PX
    size = int(s["CellSizePx"])
    while True:
        w, h = cols * size, fixed_h + rows * size
        if (w <= int(s["ScreenWidthPx"]) and h <= int(s["ScreenHeightPx"])) or size <= int(s["MinCellSizePx"]):
            break
        size -= 1
    w, h = cols * size, fixed_h + rows * size
    zoom = min(100, int(100 * min(int(s["ScreenWidthPx"]) / w, int(s["ScreenHeightPx"]) / h)))
    return size, max(10, zoom)


# --------------------------------------------------------------------------- sheets
def build_config(wb, ws, cfg, s):
    title_bar(ws, "Config", "Everything editable lives here. Change aisles or zones, then press Rebuild Layout "
              "(or Refresh for colours only) on the Map sheet. Do not sort the Settings table.", 24, s)

    section_label(ws, 3, 1, "Aisles (tblAisles)", s, 8)
    aisles = [{k: a.get(k) for k in AISLE_COLS} for a in cfg["aisles"]]
    make_table(ws, "tblAisles", HDR_ROW, 1, AISLE_COLS, aisles, widths=[9, 8, 7, 9, 9, 7, 8, 7])
    last = HDR_ROW + len(aisles)
    dv = DataValidation(type="list", formula1='"Left,Right"', allow_blank=False)
    dv.add(f"B{HDR_ROW + 1}:B{last + 40}")
    dv2 = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
    dv2.add(f"H{HDR_ROW + 1}:H{last + 40}")
    dv3 = DataValidation(type="list", formula1="ZoneCodes", allow_blank=True)
    dv3.add(f"F{HDR_ROW + 1}:F{last + 40}")
    dv4 = DataValidation(type="whole", operator="between", formula1="0", formula2="99", allow_blank=False)
    dv4.add(f"C{HDR_ROW + 1}:E{last + 40}")
    dv4.add(f"G{HDR_ROW + 1}:G{last + 40}")
    for d in (dv, dv2, dv3, dv4):
        ws.add_data_validation(d)
    ws.cell(last + 2, 1, "Side: Left/Right as seen with the fire exit behind you, facing the bay. "
            "Order: 1 = nearest the fire exit.").font = f(8, italic=True, color="595959")
    ws.cell(last + 3, 1, "Depth 01 is next to the main walkway. Column letters run A, B, C... from the "
            "top of each aisle on the Map.").font = f(8, italic=True, color="595959")

    # Zones
    zc = 10  # column J
    section_label(ws, 3, zc, "Zones (tblZones)", s, 4)
    zones = [{k: z.get(k) for k in ZONE_COLS} for z in cfg["zones"]]
    zlast = make_table(ws, "tblZones", HDR_ROW, zc, ZONE_COLS, zones, widths=[7, 20, 10, 46])
    for i, z in enumerate(zones):
        c = ws.cell(HDR_ROW + 1 + i, zc + 2)
        c.fill = solid(z["Colour"])
        c.font = f(10, True, "FFFFFF" if is_dark(z["Colour"]) else "000000")
        c.alignment = Alignment(horizontal="center")
        ws.cell(HDR_ROW + 1 + i, zc + 3).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(zlast + 2, zc, "Add a zone by adding a row. Colour is a hex code like #C8A165; keep it clearly "
            "different from the magenta brand accent.").font = f(8, italic=True, color="595959")

    # Settings
    sc = 15  # column O
    section_label(ws, 3, sc, "Settings (tblSettings)", s, 3)
    srows = [{"Setting": k, "Value": v[0], "Notes": v[1]} for k, v in cfg["settings"].items()]
    make_table(ws, "tblSettings", HDR_ROW, sc, ["Setting", "Value", "Notes"], srows, widths=[16, 11, 60])
    for i, r in enumerate(srows):
        row = HDR_ROW + 1 + i
        add_name(wb, r["Setting"], f"Config!${col_letter(sc + 1)}${row}")
        ws.cell(row, sc + 2).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row, sc + 1).alignment = Alignment(horizontal="center", vertical="top")
        ws.cell(row, sc).alignment = Alignment(vertical="top")
        v = str(r["Value"])
        if v.startswith("#") and len(v) == 7:
            ws.cell(row, sc + 1).fill = solid(v)
            ws.cell(row, sc + 1).font = f(10, True, "FFFFFF" if is_dark(v) else "000000")

    # Automatic dropdown lists (per-zone item lists for Locations[ItemCode])
    dc = 20  # column T
    ws.cell(3, dc, "Automatic dropdown lists (formulas, do not edit)").font = f(9, True, "808080")
    ws.cell(4, dc - 1, "Zone").font = f(8, italic=True, color="808080")
    ws.cell(5, dc - 1, "Count").font = f(8, italic=True, color="808080")
    first, lastr = 6, 6 + DD_ROWS - 1
    for j in range(ZONE_SLOTS + 1):
        col = dc + j
        L = col_letter(col)
        ws.column_dimensions[L].width = 11
        if j == 0:
            ws.cell(4, col, "ALL")
            ws.cell(5, col, '=COUNTIF(tblItems[ItemCode],"?*")')
        else:
            ws.cell(4, col, f'=IFERROR(INDEX(tblZones[Code],{j})&"","")')
            ws.cell(5, col, f'=IF({L}4="",0,COUNTIFS(tblItems[Zone],{L}4,tblItems[ItemCode],"?*"))')
        for r in range(4, 6):
            ws.cell(r, col).font = f(8, True, "808080")
        for r in range(first, lastr + 1):
            cond = 'tblItems[ItemCode]<>""' if j == 0 else f'(tblItems[Zone]={L}$4)*(tblItems[ItemCode]<>"")'
            formula = (f'=IFERROR(INDEX(tblItems[ItemCode],SMALL(IF({cond},ROW(tblItems[ItemCode])'
                       f'-MIN(ROW(tblItems[ItemCode]))+1),ROWS({L}${first}:{L}{r}))),"")')
            ref = f"{L}{r}"
            ws[ref] = ArrayFormula(ref, formula)
            ws[ref].font = f(8, color="808080")
    ws.column_dimensions[col_letter(dc - 1)].width = 7
    add_name(wb, "DDTop", f"Config!${col_letter(dc)}${first}")
    add_name(wb, "DDZones", f"Config!${col_letter(dc + 1)}$4:${col_letter(dc + ZONE_SLOTS)}$4")
    add_name(wb, "DDCounts", f"Config!${col_letter(dc)}$5:${col_letter(dc + ZONE_SLOTS)}$5")
    ws.freeze_panes = "A5"


def build_items(ws, s, items):
    title_bar(ws, "Items", "Item master. Fill in the white columns; the grey columns on the right calculate "
              "themselves. Replace or delete the two EXAMPLE rows. Add items by typing in the row under the table.",
              11, s)
    calc = {
        "PalletsNeeded": (f'IF(OR({tref("tblItems","ItemCode")}="",N({tref("tblItems","UnitsPerPallet")})<=0),"",'
                          f'ROUNDUP(N({tref("tblItems","OnHandUnits")})/{tref("tblItems","UnitsPerPallet")},0))'),
        "PalletsPlaced": (f'IF({tref("tblItems","ItemCode")}="","",'
                          f'COUNTIF(tblLocations[ItemCode],{tref("tblItems","ItemCode")}))'),
        "Mismatch": (f'IF(OR({tref("tblItems","ItemCode")}="",{tref("tblItems","PalletsNeeded")}=""),"",'
                     f'IF({tref("tblItems","PalletsNeeded")}<>{tref("tblItems","PalletsPlaced")},"Yes",""))'),
        "LowStock": (f'IF({tref("tblItems","ItemCode")}="","",IF({tref("tblItems","PalletsPlaced")}<='
                     f'IF({tref("tblItems","ReorderPallets")}="",DefaultLowStock,{tref("tblItems","ReorderPallets")}),'
                     f'"LOW",""))'),
    }
    headers = ITEM_INPUT_COLS + ITEM_CALC_COLS
    last = make_table(ws, "tblItems", HDR_ROW, 1, headers, items, calc,
                      widths=[14, 36, 8, 12, 12, 12, 34, 12, 12, 10, 10])
    for r in range(HDR_ROW + 1, last + 1):
        for j in range(len(ITEM_INPUT_COLS), len(headers)):
            ws.cell(r, j + 1).alignment = Alignment(horizontal="center")
            ws.cell(r, j + 1).font = f(10, True, "404040")
    for j in range(len(ITEM_INPUT_COLS), len(headers)):
        ws.cell(HDR_ROW, j + 1).fill = solid("404040")
    ws.cell(3, 1, "Zone must be a code from Config (BOX, BAG, FLM, LBL...). ReorderPallets blank = use "
            "DefaultLowStock. Mismatch = pallets placed differs from pallets needed for OnHandUnits.") \
        .font = f(8, italic=True, color="595959")
    dv = DataValidation(type="list", formula1="ZoneCodes", allow_blank=True)
    dv.add(f"C{HDR_ROW + 1}:C{HDR_ROW + 500}")
    dv2 = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True)
    dv2.add(f"D{HDR_ROW + 1}:F{HDR_ROW + 500}")
    ws.add_data_validation(dv)
    ws.add_data_validation(dv2)
    red = PatternFill(fill_type="solid", start_color="FFD9D9", end_color="FFD9D9")
    ws.conditional_formatting.add(f"K{HDR_ROW + 1}:K{HDR_ROW + 500}",
                                  FormulaRule(formula=[f'K{HDR_ROW + 1}="LOW"'], fill=red, font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(f"J{HDR_ROW + 1}:J{HDR_ROW + 500}",
                                  FormulaRule(formula=[f'J{HDR_ROW + 1}="Yes"'], fill=red, font=Font(color="C00000", bold=True)))
    ws.freeze_panes = f"B{HDR_ROW + 1}"


def build_locations(ws, s, geo, placements, items_by_code):
    title_bar(ws, "Locations", "One row per pallet position - the source of truth the Map reads from. Pick an "
              "ItemCode from the dropdown (filtered to the position's zone); leave blank for empty. Positions "
              "are created by Rebuild Layout - do not add or delete rows by hand.", 9, s)
    rows = []
    for p in geo["positions"]:
        r = dict(p)
        item, units = placements.get(p["LocationID"], (None, None))
        r["ItemCode"] = item
        r["UnitsOnPallet"] = units
        rows.append(r)
    last = make_table(ws, "tblLocations", HDR_ROW, 1, LOC_COLS, rows, widths=[12, 7, 7, 8, 7, 14, 13, 8, 8])
    for r in range(HDR_ROW + 1, last + 1):
        for c in (2, 3, 4, 5, 8, 9):
            ws.cell(r, c).alignment = Alignment(horizontal="center")
            ws.cell(r, c).font = f(10, color="595959")
        ws.cell(r, 3).number_format = "00"
    ws.cell(3, 1, "UnitsOnPallet blank = a full pallet (the item's UnitsPerPallet). MapRow/MapCol are "
            "written by the layout builder - do not edit.").font = f(8, italic=True, color="595959")
    dv = DataValidation(
        type="list", allow_blank=True, showErrorMessage=True, errorStyle="warning",
        errorTitle="Item not in this zone",
        error="That item is not listed for this position's zone (or is not on the Items sheet). Keep it anyway?",
        formula1=("OFFSET(DDTop,0,IFERROR(MATCH($E5,DDZones,0),0),"
                  "MAX(1,INDEX(DDCounts,1,IFERROR(MATCH($E5,DDZones,0),0)+1)),1)"))
    dv.add(f"F{HDR_ROW + 1}:F{last}")
    ws.add_data_validation(dv)
    dv2 = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True)
    dv2.add(f"G{HDR_ROW + 1}:G{last}")
    ws.add_data_validation(dv2)
    # Wrong zone -> red item code
    ws.conditional_formatting.add(
        f"F{HDR_ROW + 1}:F{last}",
        FormulaRule(formula=[f'AND($F{HDR_ROW + 1}<>"",$E{HDR_ROW + 1}<>"",IFERROR(INDEX(ItemZoneList,'
                             f'MATCH($F{HDR_ROW + 1},ItemCodeList,0)),$E{HDR_ROW + 1})<>$E{HDR_ROW + 1})'],
                    fill=PatternFill(fill_type="solid", start_color="FFD9D9", end_color="FFD9D9"),
                    font=Font(color="C00000", bold=True)))
    ws.freeze_panes = f"B{HDR_ROW + 1}"


def build_movelog(ws, s):
    title_bar(ws, "MoveLog", "Written automatically by the macros: every move, swap, Find Space placement "
              "and anything dropped by Rebuild Layout.", 5, s)
    make_table(ws, "tblMoveLog", HDR_ROW, 1, LOG_COLS, [], widths=[20, 16, 14, 16, 22])
    ws.cell(HDR_ROW + 1, 1).number_format = "dd/mm/yyyy hh:mm:ss"
    ws.freeze_panes = f"A{HDR_ROW + 1}"


def build_dashboard(ws, s, n_zones):
    title_bar(ws, "Dashboard", "All figures are live formulas from Items and Locations - they work with or "
              "without macros.", 20, s)
    widths = {1: 14, 2: 30, 3: 7, 4: 10, 5: 10, 6: 10, 7: 10, 8: 11, 9: 9, 10: 10, 11: 10, 12: 3,
              13: 8, 14: 20, 15: 10, 16: 9, 17: 9, 18: 9, 19: 16, 20: 3}
    for c, w in widths.items():
        ws.column_dimensions[col_letter(c)].width = w

    # KPI tiles row 4-5
    tiles = [
        ("Positions", "=ROWS(tblLocations[LocationID])"),
        ("Pallets placed", '=COUNTIF(tblLocations[ItemCode],"?*")'),
        ("Free positions", '=ROWS(tblLocations[LocationID])-COUNTIF(tblLocations[ItemCode],"?*")'),
        ("% used", '=IFERROR(COUNTIF(tblLocations[ItemCode],"?*")/ROWS(tblLocations[LocationID]),0)'),
        ("Wrong-zone pallets", f"=SUM(J9:J{8 + DASH_ITEM_ROWS})"),
        ("Low-stock items", '=COUNTIF(tblItems[LowStock],"LOW")'),
    ]
    starts = [1, 3, 5, 7, 9, 13]
    spans = [2, 2, 2, 2, 3, 3]
    for (label, formula), c0, span in zip(tiles, starts, spans):
        ws.merge_cells(start_row=4, start_column=c0, end_row=4, end_column=c0 + span - 1)
        ws.merge_cells(start_row=5, start_column=c0, end_row=5, end_column=c0 + span - 1)
        lc = ws.cell(4, c0, label)
        lc.font = f(9, True, "FFFFFF")
        lc.alignment = Alignment(horizontal="center", vertical="center")
        vc = ws.cell(5, c0, formula)
        vc.font = f(20, True, s["BrandAccent"])
        vc.alignment = Alignment(horizontal="center", vertical="center")
        if label == "% used":
            vc.number_format = "0%"
        for k in range(span):
            ws.cell(4, c0 + k).fill = solid("000000")
            ws.cell(5, c0 + k).fill = solid(s["BandGrey"])
            ws.cell(5, c0 + k).border = Border(bottom=side("thick", s["BrandAccent"]))
    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 36

    # Stock by item (A..K), rows 8..
    section_label(ws, 7, 1, "Stock by item", s, 11)
    hdr = ["Item", "Description", "Zone", "On hand (units)", "Units / pallet", "Pallets needed",
           "Pallets placed", "Units on pallets", "In zone", "Wrong zone", "Status"]
    for j, h in enumerate(hdr):
        c = ws.cell(8, 1 + j, h)
        c.font = f(9, True, "FFFFFF")
        c.fill = solid("000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[8].height = 28
    for i in range(DASH_ITEM_ROWS):
        r = 9 + i
        k = i + 1
        ws.cell(r, 1, f'=IFERROR(INDEX(tblItems[ItemCode],{k})&"","")')
        ws.cell(r, 2, f'=IF($A{r}="","",INDEX(tblItems[Description],{k})&"")')
        ws.cell(r, 3, f'=IF($A{r}="","",INDEX(tblItems[Zone],{k})&"")')
        ws.cell(r, 4, f'=IF($A{r}="","",N(INDEX(tblItems[OnHandUnits],{k})))')
        ws.cell(r, 5, f'=IF($A{r}="","",N(INDEX(tblItems[UnitsPerPallet],{k})))')
        ws.cell(r, 6, f'=IF($A{r}="","",INDEX(tblItems[PalletsNeeded],{k}))')
        ws.cell(r, 7, f'=IF($A{r}="","",COUNTIF(tblLocations[ItemCode],$A{r}))')
        ws.cell(r, 8, f'=IF($A{r}="","",SUMIFS(tblLocations[UnitsOnPallet],tblLocations[ItemCode],$A{r})'
                      f'+COUNTIFS(tblLocations[ItemCode],$A{r},tblLocations[UnitsOnPallet],"")*$E{r})')
        ws.cell(r, 9, f'=IF($A{r}="","",COUNTIFS(tblLocations[ItemCode],$A{r},tblLocations[Zone],$C{r}))')
        ws.cell(r, 10, f'=IF($A{r}="","",COUNTIFS(tblLocations[ItemCode],$A{r},tblLocations[Zone],"<>"&$C{r},'
                       f'tblLocations[Zone],"?*"))')
        ws.cell(r, 11, f'=IF($A{r}="","",IF(INDEX(tblItems[LowStock],{k})="LOW","LOW",'
                       f'IF(INDEX(tblItems[Mismatch],{k})="Yes","CHECK",IF($J{r}>0,"WRONG ZONE","OK"))))')
        for c in range(1, 12):
            cell = ws.cell(r, c)
            cell.font = f(9)
            if c != 2:
                cell.alignment = Alignment(horizontal="center")
            if i % 2:
                cell.fill = solid(s["BandGrey"])
        ws.cell(r, 4).number_format = "#,##0"
        ws.cell(r, 8).number_format = "#,##0"
    rng = f"K9:K{8 + DASH_ITEM_ROWS}"
    pink = PatternFill(fill_type="solid", start_color="FFD9D9", end_color="FFD9D9")
    ws.conditional_formatting.add(rng, FormulaRule(formula=['OR(K9="LOW",K9="WRONG ZONE")'], fill=pink,
                                                   font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=['K9="CHECK"'],
                                                   fill=PatternFill(fill_type="solid", start_color="FFF2CC",
                                                                    end_color="FFF2CC"),
                                                   font=Font(color="7F6000", bold=True)))
    ws.cell(9 + DASH_ITEM_ROWS, 1, "CHECK = pallets placed differs from pallets needed for the on-hand "
            "units.  WRONG ZONE = pallets sitting in another zone's space.").font = f(8, italic=True, color="595959")

    # Zones panel (M..S), rows 7..
    section_label(ws, 7, 13, "Space by zone", s, 7)
    zh = ["Zone", "Name", "Capacity", "Used", "Free", "% used", "Next free (nearest bay)"]
    for j, h in enumerate(zh):
        c = ws.cell(8, 13 + j, h)
        c.font = f(9, True, "FFFFFF")
        c.fill = solid("000000")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i in range(DASH_ZONE_ROWS + 1):
        r = 9 + i
        if i < DASH_ZONE_ROWS:
            ws.cell(r, 13, f'=IFERROR(INDEX(tblZones[Code],{i + 1})&"","")')
            ws.cell(r, 14, f'=IFERROR(INDEX(tblZones[Name],{i + 1})&"","")')
            crit = f"$M{r}"
            show = f'$M{r}=""'
        else:
            ws.cell(r, 13, "(none)")
            ws.cell(r, 14, "Outside any zone")
            crit = '""'
            show = "FALSE"
        ws.cell(r, 15, f'=IF({show},"",COUNTIFS(tblLocations[Zone],{crit}))')
        ws.cell(r, 16, f'=IF({show},"",COUNTIFS(tblLocations[Zone],{crit},tblLocations[ItemCode],"?*"))')
        ws.cell(r, 17, f'=IF({show},"",O{r}-P{r})')
        ws.cell(r, 18, f'=IF({show},"",IF(O{r}=0,0,P{r}/O{r}))')
        ws.cell(r, 18).number_format = "0%"
        zone_expr = f"tblLocations[Zone]=$M{r}" if i < DASH_ZONE_ROWS else 'tblLocations[Zone]=""'
        key = "(tblLocations[MapRow]+tblLocations[Depth])*1000+tblLocations[MapRow]"
        cond = f'({zone_expr})*(tblLocations[ItemCode]="")'
        nf = (f'=IF({show},"",IF(O{r}=0,"",IF(Q{r}=0,"Full",INDEX(tblLocations[LocationID],'
              f'MATCH(MIN(IF({cond},{key})),IF({cond},{key}),0)))))')
        ref = f"S{r}"
        ws[ref] = ArrayFormula(ref, nf)
        for c in range(13, 20):
            ws.cell(r, c).font = f(9)
            if c != 14:
                ws.cell(r, c).alignment = Alignment(horizontal="center")
            if i % 2:
                ws.cell(r, c).fill = solid(s["BandGrey"])
    tr = 9 + DASH_ZONE_ROWS + 1
    ws.cell(tr, 13, "Total").font = f(9, True)
    for c in (15, 16, 17):
        L = col_letter(c)
        ws.cell(tr, c, f"=SUM({L}9:{L}{tr - 1})").font = f(9, True)
        ws.cell(tr, c).alignment = Alignment(horizontal="center")
    ws.cell(tr, 18, f"=IF(O{tr}=0,0,P{tr}/O{tr})").number_format = "0%"
    ws.cell(tr, 18).font = f(9, True)
    ws.cell(tr, 18).alignment = Alignment(horizontal="center")
    for c in range(13, 20):
        ws.cell(tr, c).border = Border(top=side("thin", "000000"))
    ws.conditional_formatting.add(f"R9:R{tr}", FormulaRule(formula=['AND(ISNUMBER(R9),R9>=0.9)'],
                                                            font=Font(color="C00000", bold=True)))
    ws.conditional_formatting.add(f"S9:S{tr - 1}", FormulaRule(formula=['S9="Full"'],
                                                                font=Font(color="C00000", bold=True)))

    # Low-stock list
    lr = tr + 3
    section_label(ws, lr, 13, "Low stock (at or below reorder level)", s, 7)
    for j, h in enumerate(["Item", "Description", "Placed", "Reorder at"]):
        c = 13 + j
        ws.cell(lr + 1, c, h)
        ws.cell(lr + 1, c).font = f(9, True, "FFFFFF")
        ws.cell(lr + 1, c).fill = solid("000000")
        ws.cell(lr + 1, c).alignment = Alignment(horizontal="center")
    for i in range(DASH_LOW_ROWS):
        r = lr + 2 + i
        ref = f"M{r}"
        ws[ref] = ArrayFormula(ref, (
            f'=IFERROR(INDEX(tblItems[ItemCode],SMALL(IF(tblItems[LowStock]="LOW",ROW(tblItems[LowStock])'
            f'-MIN(ROW(tblItems[LowStock]))+1),ROWS(M${lr + 2}:M{r}))),"")'))
        ws.cell(r, 14, f'=IF(M{r}="","",INDEX(tblItems[Description],MATCH(M{r},tblItems[ItemCode],0))&"")')
        ws.cell(r, 15, f'=IF(M{r}="","",INDEX(tblItems[PalletsPlaced],MATCH(M{r},tblItems[ItemCode],0)))')
        ws.cell(r, 16, f'=IF(M{r}="","",IF(INDEX(tblItems[ReorderPallets],MATCH(M{r},tblItems[ItemCode],0))="",'
                       f'DefaultLowStock,INDEX(tblItems[ReorderPallets],MATCH(M{r},tblItems[ItemCode],0))))')
        for c in range(13, 17):
            ws.cell(r, c).font = f(9, c == 13, "C00000" if c == 13 else "000000")
            if c != 14:
                ws.cell(r, c).alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A4"


def paint_position(cell, p, zone_colours, item_zone, low_items, s):
    """Fill/border/font for one pallet square. Mirrors PaintLocation in the VBA."""
    zone = p["Zone"]
    item = p.get("ItemCode") or ""
    colour = zone_colours.get(zone)
    white = side("thin", "FFFFFF")
    if colour is None:
        bg = s["NoZoneUsedGrey"] if item else s["NoZoneGrey"]
        cell.fill = solid(bg)
        dark = is_dark(bg)
    elif item:
        cell.fill = solid(colour)
        dark = is_dark(colour)
    else:
        cell.fill = PatternFill(fill_type="lightUp", start_color=hx(colour), end_color="FFFFFF")
        dark = False
    cell.font = Font(name=FONT, size=7, color="FFFFFF" if dark else "000000")
    cell.alignment = Alignment(horizontal="left", vertical="center", shrink_to_fit=False)
    border = Border(left=white, right=white, top=white, bottom=white)
    if item:
        izone = item_zone.get(item)
        if zone and colour is not None and izone and izone != zone:
            red = side("thick", s["AlertRed"])
            border = Border(left=red, right=red, top=red, bottom=red)
        elif item in low_items:
            border = Border(left=white, right=white, top=white, bottom=side("thick", s["AlertRed"]))
        cm = Comment(f"{p['LocationID']}\n{item}\nUnits: {p.get('UnitsOnPallet') or 'full pallet'}"
                     f"\nZone: {zone or '(none)'}", "Floor Map")
        cm.width, cm.height = 180, 80
        cell.comment = cm
    cell.border = border


def build_map(wb, ws, s, geo, zones, positions, item_zone, low_items, size, zoom, show_labels):
    last_col = geo["last_col"]
    width = size / 7.0
    for c in range(1, last_col + 1):
        ws.column_dimensions[col_letter(c)].width = width
    ws.row_dimensions[ROW_TITLE].height = TITLE_PX * 0.75
    ws.row_dimensions[ROW_INFO].height = INFO_PX * 0.75
    ws.row_dimensions[ROW_BUTTONS].height = BUTTONS_PX * 0.75
    for r in range(ROW_BAY, geo["fire_row"] + 1):
        ws.row_dimensions[r].height = size * 0.75

    # Title bar with LogoSpot on the right
    logo_w = min(6, last_col // 4)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col - logo_w)
    c = ws.cell(1, 1, "Consumables Floor Map, Sherburn")
    c.font = f(14, True, s["HeaderText"])
    c.alignment = Alignment(vertical="center", indent=1)
    ws.merge_cells(start_row=1, start_column=last_col - logo_w + 1, end_row=1, end_column=last_col)
    lc = ws.cell(1, last_col - logo_w + 1, "Logo")
    lc.font = f(8, italic=True, color="808080")
    lc.alignment = Alignment(horizontal="center", vertical="center")
    add_name(wb, "LogoSpot", f"Map!${col_letter(last_col - logo_w + 1)}$1")
    for col in range(1, last_col + 1):
        ws.cell(1, col).fill = solid(s["HeaderFill"])
        ws.cell(1, col).border = Border(bottom=side("thick", s["BrandAccent"]))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_col)
    c = ws.cell(2, 1, "Click a pallet, then click where it goes (empty = move, occupied = swap). Click it again "
                      "or press Esc / Cancel Move to cancel. Without macros: edit ItemCode on the Locations sheet.")
    c.font = f(8, italic=True, color="404040")
    c.alignment = Alignment(vertical="center", indent=1)

    map_first, map_last = 2, geo["map_last_col"]
    # Loading bay and fire exit bars
    for row, text, colour in ((ROW_BAY, "LOADING BAY", s["HeaderFill"]),
                              (geo["fire_row"], "FIRE EXIT", s["FireExitGreen"])):
        ws.merge_cells(start_row=row, start_column=map_first, end_row=row, end_column=map_last)
        c = ws.cell(row, map_first, text)
        c.font = f(9, True, "FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(map_first, map_last + 1):
            ws.cell(row, col).fill = solid(colour)

    # Main walkway label
    top, bottom = ROW_AREA_TOP, geo["bottom"]
    ws.merge_cells(start_row=top, start_column=geo["walk_start"], end_row=bottom, end_column=geo["walk_end"])
    c = ws.cell(top, geo["walk_start"], "MAIN WALKWAY")
    c.font = f(8, True, "A6A6A6")
    c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)

    # Aisle labels
    zone_names = {z["Code"]: z["Name"] for z in zones}
    for label_row, (c0, c1), a in geo["labels"]:
        if c1 > c0:
            ws.merge_cells(start_row=label_row, start_column=c0, end_row=label_row, end_column=c1)
        name = zone_names.get(str(a.get("Zone") or ""), "No zone")
        c = ws.cell(label_row, c0, f"{a['AisleID']}  {name}" if c1 - c0 >= 5 else str(a["AisleID"]))
        c.font = f(8, True, "000000")
        c.alignment = Alignment(horizontal="left", vertical="bottom")
        for col in range(c0, c1 + 1):
            ws.cell(label_row, col).border = Border(bottom=side("thin", s["BrandAccent"]))

    # Pallet squares: formula shows the (short) item code live from tblLocations
    zone_colours = {z["Code"]: z["Colour"] for z in zones}
    numfmt = "General" if show_labels else ";;;"
    for p in positions:
        cell = ws.cell(p["MapRow"], p["MapCol"])
        cell.value = (f'=IFERROR(LEFT(INDEX(tblLocations[ItemCode],MATCH("{p["LocationID"]}",'
                      f'tblLocations[LocationID],0))&"",MapLabelChars),"")')
        cell.number_format = numfmt
        paint_position(cell, p, zone_colours, item_zone, low_items, s)

    # No-macro fallback: solid zone colour when occupied, one rule per aisle
    by_aisle = {}
    for p in positions:
        by_aisle.setdefault(p["Aisle"], []).append(p)
    for aisle, ps in by_aisle.items():
        r0 = min(p["MapRow"] for p in ps)
        r1 = max(p["MapRow"] for p in ps)
        c0 = min(p["MapCol"] for p in ps)
        c1 = max(p["MapCol"] for p in ps)
        zone = ps[0]["Zone"]
        colour = zone_colours.get(zone, s["NoZoneUsedGrey"])
        tl = f"{col_letter(c0)}{r0}"
        ws.conditional_formatting.add(
            f"{tl}:{col_letter(c1)}{r1}",
            FormulaRule(formula=[f"LEN({tl})>0"], fill=solid(colour),
                        font=Font(color="FFFFFF" if is_dark(colour) else "000000")))

    # Legend
    lc0 = geo["legend_col"]
    r = ROW_AREA_TOP

    def legend_header(row, text):
        ws.merge_cells(start_row=row, start_column=lc0, end_row=row, end_column=lc0 + LEGEND_COLS - 1)
        c = ws.cell(row, lc0, text)
        c.font = f(8, True, "FFFFFF")
        c.alignment = Alignment(vertical="center", indent=1)
        for k in range(LEGEND_COLS):
            ws.cell(row, lc0 + k).fill = solid(s["HeaderFill"])

    def legend_text(row, text, bold=False):
        ws.merge_cells(start_row=row, start_column=lc0 + 3, end_row=row, end_column=lc0 + LEGEND_COLS - 1)
        c = ws.cell(row, lc0 + 3, text)
        c.font = f(8, bold)
        c.alignment = Alignment(vertical="center")
        return c

    add_name(wb, "MapLegend", f"Map!${col_letter(lc0)}${r}")
    legend_header(r, "ZONES")
    for z in zones:
        r += 1
        ws.cell(r, lc0).fill = solid(z["Colour"])
        ws.cell(r, lc0 + 1).fill = PatternFill(fill_type="lightUp", start_color=hx(z["Colour"]), end_color="FFFFFF")
        legend_text(r, f"{z['Code']}  {z['Name']}")
    r += 2
    legend_header(r, "KEY")
    white = side("thin", "FFFFFF")
    red = side("thick", s["AlertRed"])
    sample = zones[0]["Colour"] if zones else s["NoZoneGrey"]
    keys = [
        ("Pallet in place", solid(sample), None),
        ("Free space kept for that zone", PatternFill(fill_type="lightUp", start_color=hx(sample), end_color="FFFFFF"), None),
        ("Not in any zone", solid(s["NoZoneGrey"]), None),
        ("Wrong zone", solid(sample), Border(left=red, right=red, top=red, bottom=red)),
        ("Selected, being moved", solid(sample), Border(*(side("thick", "000000"),) * 4)),
        ("Low stock item", solid(sample), Border(left=white, right=white, top=white, bottom=red)),
    ]
    for text, fl, bd in keys:
        r += 1
        ws.cell(r, lc0).fill = fl
        if bd is not None:
            ws.cell(r, lc0).border = bd
        legend_text(r, text)
    r += 2
    legend_header(r, "LIVE COUNTS")
    for text, formula in (("Positions", "=ROWS(tblLocations[LocationID])"),
                          ("Pallets placed", '=COUNTIF(tblLocations[ItemCode],"?*")'),
                          ("Free", '=ROWS(tblLocations[LocationID])-COUNTIF(tblLocations[ItemCode],"?*")')):
        r += 1
        ws.merge_cells(start_row=r, start_column=lc0, end_row=r, end_column=lc0 + 2)
        c = ws.cell(r, lc0, formula)
        c.font = f(8, True, s["BrandAccent"])
        c.alignment = Alignment(horizontal="right", vertical="center")
        legend_text(r, text)

    add_name(wb, "MapArea", f"Map!$A$1:${col_letter(last_col)}${geo['fire_row']}")
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False
    ws.sheet_view.zoomScale = zoom
    ws.print_area = f"A1:{col_letter(last_col)}{geo['fire_row']}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"


# --------------------------------------------------------------------------- demo data
DEMO_ITEMS = [
    {"ItemCode": "BX-S", "Description": "Single wall box small 305x229x152", "Zone": "BOX", "UnitsPerPallet": 500, "OnHandUnits": 1800},
    {"ItemCode": "BX-L", "Description": "Double wall box large 610x457x457", "Zone": "BOX", "UnitsPerPallet": 150, "OnHandUnits": 300, "ReorderPallets": 3},
    {"ItemCode": "PW-17", "Description": "Pallet wrap 17mu 500mm", "Zone": "FLM", "UnitsPerPallet": 240, "OnHandUnits": 480},
    {"ItemCode": "LB-A6", "Description": "Thermal labels A6", "Zone": "LBL", "UnitsPerPallet": 96, "OnHandUnits": 96},
    {"ItemCode": "BG-M", "Description": "Poly mailing bag medium", "Zone": "BAG", "UnitsPerPallet": 2000, "OnHandUnits": 9000},
]
DEMO_PLACEMENTS = {
    "R1-01-A": ("BX-S", None), "R1-02-A": ("BX-S", None), "R1-03-A": ("BX-S", None), "R1-04-A": ("BX-S", 300),
    "L1-01-A": ("BX-L", None), "L1-02-A": ("BX-L", None),
    "R3-01-A": ("PW-17", None), "R3-01-B": ("PW-17", None),
    "R4-01-A": ("LB-A6", None),
    "R5-01-A": ("BG-M", None), "R5-02-A": ("BG-M", None), "R5-03-A": ("BG-M", None),
    "R2-05-C": ("BG-M", None),          # deliberately in the wrong zone
}

EXAMPLE_ITEMS = [
    {"ItemCode": "EX-BOX-01", "Description": "EXAMPLE - Single wall box 400x300x300", "Zone": "BOX",
     "UnitsPerPallet": 400, "OnHandUnits": 0, "Notes": "EXAMPLE row - overwrite or delete"},
    {"ItemCode": "EX-FLM-01", "Description": "EXAMPLE - Pallet wrap 23mu 500mm", "Zone": "FLM",
     "UnitsPerPallet": 200, "OnHandUnits": 0, "Notes": "EXAMPLE row - overwrite or delete"},
]


# --------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("-l", "--layout", default="layout.yaml")
    ap.add_argument("-o", "--out", default="Consumables_Map.xlsx")
    ap.add_argument("--demo", action="store_true", help="add 5 dummy items and pallets (testing only)")
    args = ap.parse_args()

    cfg = yaml.safe_load(Path(args.layout).read_text())
    s = {k: v[0] for k, v in cfg["settings"].items()}
    zones = cfg["zones"]
    for z in zones:
        if colour_distance(z["Colour"], s["BrandAccent"]) < 80:
            print(f"WARNING: zone {z['Code']} colour {z['Colour']} is too close to the brand accent {s['BrandAccent']}")
    ids = [str(a["AisleID"]) for a in cfg["aisles"]]
    if len(set(ids)) != len(ids):
        raise SystemExit("Duplicate AisleID in layout")
    for a in cfg["aisles"]:
        if not re.fullmatch(r"[A-Za-z0-9]+", str(a["AisleID"])):
            raise SystemExit(f"AisleID {a['AisleID']!r}: use letters and numbers only")
        if int(a["PalletsAcross"]) > 26:
            raise SystemExit(f"Aisle {a['AisleID']}: PalletsAcross must be 26 or fewer")

    items = [dict(i) for i in EXAMPLE_ITEMS]
    placements = {}
    if args.demo:
        items += [dict(i) for i in DEMO_ITEMS]
        placements = DEMO_PLACEMENTS

    geo = compute_layout(cfg["aisles"], zones, int(s["WalkwayCells"]))
    size, zoom = fit_to_screen(geo, s)
    show = str(s["ShowMapLabels"]).strip().lower()
    show_labels = show == "yes" or (show == "auto" and size >= 16)

    # Python-side view of flags, used only for painting at build time
    item_zone = {i["ItemCode"]: i["Zone"] for i in items}
    placed_count = {}
    for loc, (it, _) in placements.items():
        placed_count[it] = placed_count.get(it, 0) + 1
    low_items = {i["ItemCode"] for i in items
                 if placed_count.get(i["ItemCode"], 0) <= (i.get("ReorderPallets") or s["DefaultLowStock"])}
    positions = []
    for p in geo["positions"]:
        q = dict(p)
        q["ItemCode"], q["UnitsOnPallet"] = placements.get(p["LocationID"], (None, None))
        positions.append(q)
    unknown = set(placements) - {p["LocationID"] for p in positions}
    if unknown:
        raise SystemExit(f"Demo placements not in layout: {sorted(unknown)}")

    wb = Workbook()
    wb._named_styles["Normal"].font = Font(name=FONT, size=10)
    ws_map = wb.active
    ws_map.title = "Map"
    ws_dash = wb.create_sheet("Dashboard")
    ws_items = wb.create_sheet("Items")
    ws_loc = wb.create_sheet("Locations")
    ws_cfg = wb.create_sheet("Config")
    ws_log = wb.create_sheet("MoveLog")
    for ws, colour in ((ws_map, s["BrandAccent"]), (ws_dash, "000000"), (ws_items, "000000"),
                       (ws_loc, "000000"), (ws_cfg, "808080"), (ws_log, "808080")):
        ws.sheet_properties.tabColor = hx(colour)

    add_name(wb, "ZoneCodes", "tblZones[Code]")
    add_name(wb, "ItemCodeList", "tblItems[ItemCode]")
    add_name(wb, "ItemZoneList", "tblItems[Zone]")

    build_config(wb, ws_cfg, cfg, s)
    build_items(ws_items, s, items)
    build_locations(ws_loc, s, geo, placements, item_zone)
    build_movelog(ws_log, s)
    build_dashboard(ws_dash, s, len(zones))
    build_map(wb, ws_map, s, geo, zones, positions, item_zone, low_items, size, zoom, show_labels)

    wb.active = 0
    wb.save(args.out)
    print(f"Wrote {args.out}: {len(positions)} positions, cell {size}px, zoom {zoom}%, "
          f"map {geo['last_col']} cols x {geo['fire_row']} rows")


if __name__ == "__main__":
    main()
