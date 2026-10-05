#!/usr/bin/env python3
"""Build Floor_Map_Simple.xlsx - the easy, no-macro version of the floor map.

    python build_simple.py            # uses layout.yaml
    python build_simple.py --demo     # adds a few test items/pallets

How it works for the user:
  * Map sheet: click a square, pick the item from the little arrow. Delete to empty.
  * Stock sheet: list of items; pallet counts and LOW flags work themselves out.
No macros, no Locations table, nothing to set up.
"""
import argparse
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.worksheet.table import Table, TableColumn, TableFormula, TableStyleInfo

from build_map import (ROW_AREA_TOP, LEGEND_COLS, add_name, blocked, compute_layout, f, hx,
                       is_dark, side, solid)

MAP_HELP = ("PUT AWAY: click a square \u2192 arrow \u2192 pick.     MOVE: double-click a pallet, "
            "then click where it goes.     REMOVE: click it, press Delete.")
CELL_PX = 28          # big squares, easy to click
DD_ROWS = 150         # max items per dropdown list
STOCK_ROWS_HINT = 2   # example rows


def stripes(colour):
    return PatternFill(fill_type="lightUp", start_color=hx(colour), end_color="FFFFFF")


def build(cfg, out, demo):
    s = {k: v[0] for k, v in cfg["settings"].items()}
    zones = cfg["zones"]
    zcol = {z["Code"]: z["Colour"] for z in zones}
    zname = {z["Code"]: z["Name"] for z in zones}
    geo = compute_layout(cfg["aisles"], zones, int(s["WalkwayCells"]))

    items = [
        {"Item code": "EX-BOX", "Description": "EXAMPLE - type over me", "Zone": "BOX"},
        {"Item code": "EX-WRAP", "Description": "EXAMPLE - type over me", "Zone": "FLM"},
    ]
    placed = {}
    if demo:
        items = [
            {"Item code": "BX-S", "Description": "Small box", "Zone": "BOX"},
            {"Item code": "BX-L", "Description": "Large box", "Zone": "BOX", "Low when pallets at or below": 3},
            {"Item code": "WRAP", "Description": "Pallet wrap", "Zone": "FLM"},
            {"Item code": "LBL6", "Description": "A6 labels", "Zone": "LBL"},
            {"Item code": "BAG-M", "Description": "Medium bag", "Zone": "BAG"},
        ]
        placed = {"R4-01-A": "BX-S", "R4-02-A": "BX-S", "R4-03-A": "BX-S", "L1-01-A": "BX-L",
                  "R2-01-A": "WRAP", "R2-01-B": "WRAP", "R1-01-A": "LBL6", "R3-01-A": "BAG-M",
                  "R3-02-A": "BAG-M", "L2-05-C": "BAG-M"}

    wb = Workbook()
    wb._named_styles["Normal"].font = Font(name="Arial", size=10)
    ws = wb.active
    ws.title = "Map"
    st = wb.create_sheet("Stock")
    lay = wb.create_sheet("Layout")
    moves = wb.create_sheet("Moves")
    lists = wb.create_sheet("Lists")
    squares = wb.create_sheet("Squares")
    ws.sheet_properties.tabColor = hx(s["BrandAccent"])
    st.sheet_properties.tabColor = "000000"

    # ------------------------------------------------------------ Stock sheet
    st.sheet_view.showGridLines = False
    st.merge_cells("A1:F1")
    c = st["A1"]
    c.value = "Stock"
    c.font = f(18, True, "FFFFFF")
    c.alignment = Alignment(vertical="center", indent=1)
    for col in range(1, 7):
        st.cell(1, col).fill = solid("000000")
        st.cell(1, col).border = Border(bottom=side("thick", s["BrandAccent"]))
    st.row_dimensions[1].height = 34
    st.merge_cells("A2:F2")
    st["A2"].value = ("Everything is counted in PALLETS. Type your items in the white columns; grey columns fill in by themselves. "
                      "Add a new item on the first empty row.")
    st["A2"].font = f(11, color="404040")
    st["A2"].alignment = Alignment(vertical="center", indent=1, wrap_text=True)
    st.row_dimensions[2].height = 30

    heads = ["Item code", "Description", "Zone", "Low when pallets at or below", "Pallets on map", "Status"]
    # big fixed range so the count still works after "Rebuild map" makes the map bigger
    map_rng = f"Map!$A${ROW_AREA_TOP}:$DZ$300"
    tr = lambda col: f"tblStock[[#This Row],[{col}]]"
    calc = {
        "Pallets on map": f'IF({tr("Item code")}="","",COUNTIF({map_rng},{tr("Item code")}))',
        "Status": f'IF({tr("Item code")}="","",IF({tr("Pallets on map")}<='
                  f'IF({tr("Low when pallets at or below")}="",{int(s["DefaultLowStock"])},'
                  f'{tr("Low when pallets at or below")}),"LOW","OK"))',
    }
    hr = 4
    for j, h in enumerate(heads):
        cc = st.cell(hr, j + 1, h)
        cc.font = f(11, True, "FFFFFF")
        cc.fill = solid("000000" if h not in calc else "404040")
        cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    st.row_dimensions[hr].height = 36
    n = len(items)
    for i, it in enumerate(items):
        r = hr + 1 + i
        for j, h in enumerate(heads):
            cc = st.cell(r, j + 1)
            cc.value = ("=" + calc[h]) if h in calc else it.get(h)
            cc.font = f(12, h in ("Item code", "Status"))
            cc.alignment = Alignment(horizontal="left" if h == "Description" else "center", vertical="center")
        st.row_dimensions[r].height = 22
    t = Table(displayName="tblStock", ref=f"A{hr}:F{hr + n}")
    cols = []
    for j, h in enumerate(heads):
        tc = TableColumn(id=j + 1, name=h)
        if h in calc:
            tc.calculatedColumnFormula = TableFormula(attr_text=calc[h])
        cols.append(tc)
    t.tableColumns = cols
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium15", showRowStripes=True)
    st.add_table(t)
    for j, w in enumerate([14, 36, 9, 16, 12, 11]):
        st.column_dimensions[L(j + 1)].width = w
    dv = DataValidation(type="list", formula1=f'"{",".join(z["Code"] for z in zones)}"', allow_blank=True)
    dv.add(f"C{hr + 1}:C{hr + 300}")
    st.add_data_validation(dv)
    red = PatternFill(fill_type="solid", start_color="FFD9D9", end_color="FFD9D9")
    st.conditional_formatting.add(f"F{hr + 1}:F{hr + 300}", FormulaRule(
        formula=[f'F{hr + 1}="LOW"'], fill=red, font=Font(color="C00000", bold=True)))
    st.conditional_formatting.add(f"F{hr + 1}:F{hr + 300}", FormulaRule(
        formula=[f'F{hr + 1}="OK"'], font=Font(color="00843D", bold=True)))
    st.freeze_panes = f"B{hr + 1}"

    # Zone key + free space, to the right
    # Code (used in the Zone column) | Zone name (editable) | Spaces | Used | Free
    zc = 8
    for j, w in enumerate([8, 24, 9, 9, 9]):
        st.column_dimensions[L(zc + j)].width = w
    for j, h in enumerate(["Code", "Zone name (you can change it)", "Spaces", "Used", "Free"]):
        cc = st.cell(hr, zc + j, h)
        cc.font = f(11, True, "FFFFFF")
        cc.fill = solid("000000")
        cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    edit = side("medium", s["BrandAccent"])
    zrow = {}
    for i, z in enumerate(zones):
        sr = hr + 1 + i
        zrow[z["Code"]] = sr
        cc = st.cell(sr, zc, z["Code"])
        cc.fill = solid(z["Colour"])
        cc.font = f(12, True, "FFFFFF" if is_dark(z["Colour"]) else "000000")
        cc.alignment = Alignment(horizontal="center", vertical="center")
        cc = st.cell(sr, zc + 1, z["Name"])
        cc.font = f(12, True)
        cc.border = Border(left=edit, right=edit, top=edit, bottom=edit)
        cc.alignment = Alignment(vertical="center", indent=1)
        add_name(wb, f"zn_{z['Code']}", f"Stock!${L(zc + 1)}${sr}")
        st.row_dimensions[sr].height = 22
    st.cell(hr + len(zones) + 1, zc, "Type over a zone name (pink box) and the Map changes too. "
            "In the Zone column on the left, use the Code.").font = f(9, False, "595959")

    # ------------------------------------------------------------ Lists sheet (hidden helper)
    lists.sheet_state = "hidden"
    codes = ["ALL"] + [z["Code"] for z in zones]
    for j, code in enumerate(codes):
        col = j + 1
        Lc = L(col)
        lists.cell(1, col, code)
        lists.cell(2, col, '=COUNTIF(tblStock[Item code],"?*")' if code == "ALL" else
                   f'=COUNTIFS(tblStock[Zone],"{code}",tblStock[Item code],"?*")')
        cond = 'tblStock[Item code]<>""' if code == "ALL" else f'(tblStock[Zone]="{code}")*(tblStock[Item code]<>"")'
        for r in range(3, 3 + DD_ROWS):
            ref = f"{Lc}{r}"
            lists[ref] = ArrayFormula(ref, (
                f'=IFERROR(INDEX(tblStock[Item code],SMALL(IF({cond},ROW(tblStock[Item code])'
                f'-MIN(ROW(tblStock[Item code]))+1),ROWS({Lc}$3:{Lc}{r}))),"")'))
        add_name(wb, f"dd_{code}", f"OFFSET(Lists!${Lc}$3,0,0,MAX(1,Lists!${Lc}$2),1)")
    add_name(wb, "ItemCodes", "tblStock[Item code]")
    add_name(wb, "ItemZones", "tblStock[Zone]")

    # ------------------------------------------------------------ Map sheet
    last = geo["last_col"]
    for col in range(1, last + 1):
        ws.column_dimensions[L(col)].width = CELL_PX / 7.0
    ws.row_dimensions[1].height = 30
    ws.row_dimensions[2].height = 22
    ws.row_dimensions[3].height = 24   # room for the macro buttons
    for r in range(4, geo["fire_row"] + 1):
        ws.row_dimensions[r].height = CELL_PX * 0.75
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last)
    c = ws.cell(1, 1, "Consumables Floor Map, Sherburn")
    c.font = f(18, True, "FFFFFF")
    c.alignment = Alignment(vertical="center", indent=1)
    for col in range(1, last + 1):
        ws.cell(1, col).fill = solid("000000")
        ws.cell(1, col).border = Border(bottom=side("thick", s["BrandAccent"]))
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last)
    c = ws.cell(2, 1, MAP_HELP)
    c.font = f(12, True, s["BrandAccent"])
    c.alignment = Alignment(vertical="center", indent=1)

    mf, ml = 2, geo["map_last_col"]
    for row, text, colour in ((4, "LOADING BAY", "000000"), (geo["fire_row"], "FIRE EXIT", s["FireExitGreen"])):
        ws.merge_cells(start_row=row, start_column=mf, end_row=row, end_column=ml)
        c = ws.cell(row, mf, text)
        c.font = f(12, True, "FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(mf, ml + 1):
            ws.cell(row, col).fill = solid(colour)

    ws.merge_cells(start_row=ROW_AREA_TOP, start_column=geo["walk_start"], end_row=geo["bottom"], end_column=geo["walk_end"])
    c = ws.cell(ROW_AREA_TOP, geo["walk_start"], "WALKWAY")
    c.font = f(12, True, "BFBFBF")
    c.alignment = Alignment(horizontal="center", vertical="center", text_rotation=90)

    white = side("thin", "FFFFFF")
    grid = Border(left=white, right=white, top=white, bottom=white)
    zone_counts = {}      # zone -> list of aisle ranges (for Used counts)
    for label_row, (c0, c1), a in geo["labels"]:
        if c1 > c0:
            ws.merge_cells(start_row=label_row, start_column=c0, end_row=label_row, end_column=c1)
        zone = str(a.get("Zone") or "")
        if blocked(a.get("Active")):
            text = f"{a['AisleID']}   NOT OURS"
        elif zone in zname:
            text = f'="{a["AisleID"]}   "&zn_{zone}'
        else:
            text = f"{a['AisleID']}   No zone"
        c = ws.cell(label_row, c0, text)
        c.font = f(11, True)
        c.alignment = Alignment(horizontal="left", vertical="bottom")
    for r0, r1, (c0, c1), a in geo["blocks"]:
        ws.merge_cells(start_row=r0, start_column=c0, end_row=r1, end_column=c1)
        c = ws.cell(r0, c0, "NOT OURS - someone else's space")
        c.font = f(11, True, "808080")
        c.alignment = Alignment(horizontal="center", vertical="center")
        for rr in range(r0, r1 + 1):
            for cc in range(c0, c1 + 1):
                ws.cell(rr, cc).fill = solid(s["NotOursGrey"])

    # Pallet squares, aisle by aisle
    by_aisle = {}
    for p in geo["positions"]:
        by_aisle.setdefault(p["Aisle"], []).append(p)
    for aisle, ps in by_aisle.items():
        zone = ps[0]["Zone"]
        colour = zcol.get(zone)
        r0, r1 = min(p["MapRow"] for p in ps), max(p["MapRow"] for p in ps)
        c0, c1 = min(p["MapCol"] for p in ps), max(p["MapCol"] for p in ps)
        rng = f"{L(c0)}{r0}:{L(c1)}{r1}"
        tl = f"{L(c0)}{r0}"
        zone_counts.setdefault(zone, []).append((rng, len(ps)))
        for p in ps:
            cell = ws.cell(p["MapRow"], p["MapCol"])
            cell.value = placed.get(p["LocationID"])
            cell.fill = stripes(colour) if colour else solid(s["NoZoneGrey"])
            cell.border = grid
            cell.font = Font(name="Arial", size=7, bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        dvz = DataValidation(
            type="list", formula1=f"dd_{zone}" if colour else "dd_ALL", allow_blank=True,
            showErrorMessage=True, errorStyle="warning", errorTitle="Not for this zone",
            error="That item isn't listed for this zone (or isn't on the Stock sheet). Put it here anyway?",
            showInputMessage=True, promptTitle=f"Aisle {aisle}  ({zone or 'no zone'})"[:32],
            prompt="Click the arrow and pick an item. Press Delete to empty.")
        dvz.add(rng)
        ws.add_data_validation(dvz)
        if colour:
            # Wrong zone first: red square
            ws.conditional_formatting.add(rng, FormulaRule(
                formula=[f'AND(LEN({tl})>0,IFERROR(INDEX(ItemZones,MATCH({tl},ItemCodes,0)),"{zone}")<>"{zone}")'],
                fill=solid(s["AlertRed"]), font=Font(color="FFFFFF", bold=True), stopIfTrue=True))
        fill_c = colour or s["NoZoneUsedGrey"]
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f"LEN({tl})>0"], fill=solid(fill_c),
            font=Font(color="FFFFFF" if is_dark(fill_c) else "000000", bold=True)))

    # Legend with live free-space counts
    lc = geo["legend_col"]
    r = ROW_AREA_TOP

    def band(row, text):
        ws.merge_cells(start_row=row, start_column=lc, end_row=row, end_column=lc + LEGEND_COLS - 1)
        cc = ws.cell(row, lc, text)
        cc.font = f(11, True, "FFFFFF")
        cc.alignment = Alignment(vertical="center", indent=1)
        for k in range(LEGEND_COLS):
            ws.cell(row, lc + k).fill = solid("000000")

    def text(row, col0, value, **o):
        ws.merge_cells(start_row=row, start_column=col0, end_row=row, end_column=lc + LEGEND_COLS - 1)
        cc = ws.cell(row, col0, value)
        cc.font = f(o.get("size", 11), o.get("bold", False), o.get("color", "000000"))
        cc.alignment = Alignment(vertical="center")

    band(r, "COLOURS")
    for z in zones:
        r += 1
        ws.cell(r, lc).fill = solid(z["Colour"])
        text(r, lc + 2, f"=zn_{z['Code']}")
    keys = [(stripes(zones[0]["Colour"]), "Stripes = free space"),
            (solid(s["AlertRed"]), "Red = wrong zone, move it"),
            (solid(s["NotOursGrey"]), "Grey = not ours")]
    for fl, t_ in keys:
        r += 1
        ws.cell(r, lc).fill = fl
        text(r, lc + 2, t_)
    r += 2
    band(r, "FREE SPACES")
    for z in zones:
        entries = zone_counts.get(z["Code"], [])
        if not entries:
            continue
        r += 1
        rngs = [x for x, _ in entries]
        total = sum(k for _, k in entries)
        used = "+".join(f"COUNTA({x})" for x in rngs)
        ws.merge_cells(start_row=r, start_column=lc, end_row=r, end_column=lc + 1)
        cc = ws.cell(r, lc, f"={total}-({used})")
        cc.font = f(12, True, s["BrandAccent"])
        cc.alignment = Alignment(horizontal="right", vertical="center")
        text(r, lc + 2, f"=zn_{z['Code']}")
        # Stock sheet zone summary
        sr = zrow[z["Code"]]
        st.cell(sr, zc + 2, total)
        st.cell(sr, zc + 3, "=" + "+".join(f"COUNTA(Map!{x})" for x in rngs))
        st.cell(sr, zc + 4, f"={L(zc + 2)}{sr}-{L(zc + 3)}{sr}")
        for k in (2, 3, 4):
            st.cell(sr, zc + k).font = f(12, k == 4)
            st.cell(sr, zc + k).alignment = Alignment(horizontal="center", vertical="center")
    r += 2
    band(r, "LOW STOCK")
    r += 1
    ws.merge_cells(start_row=r, start_column=lc, end_row=r, end_column=lc + 1)
    cc = ws.cell(r, lc, '=COUNTIF(tblStock[Status],"LOW")')
    cc.font = f(12, True, "C00000")
    cc.alignment = Alignment(horizontal="right", vertical="center")
    text(r, lc + 2, "items LOW - see Stock sheet")

    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False
    total_w = last * CELL_PX
    total_h = 30 / 0.75 + 22 / 0.75 + 8 + (geo["fire_row"] - 3) * CELL_PX
    zoom = min(100, int(100 * min(1860 / total_w, 800 / total_h)))
    ws.sheet_view.zoomScale = zoom
    ws.print_area = f"A1:{L(last)}{geo['fire_row']}"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_view.selection[0].activeCell = "A1"
    ws.sheet_view.selection[0].sqref = "A1"

    # ------------------------------------------------------------ Layout (aisle sizes, used by "Rebuild map")
    lay.sheet_properties.tabColor = "808080"
    lay.sheet_view.showGridLines = False
    lay.merge_cells("A1:G1")
    lay["A1"].value = "Layout"
    lay["A1"].font = f(18, True, "FFFFFF")
    lay["A1"].alignment = Alignment(vertical="center", indent=1)
    for col in range(1, 8):
        lay.cell(1, col).fill = solid("000000")
        lay.cell(1, col).border = Border(bottom=side("thick", s["BrandAccent"]))
    lay.row_dimensions[1].height = 34
    lay.merge_cells("A2:G2")
    lay["A2"].value = ("Change the numbers, then press the REBUILD MAP button. "
                       "Pallets already on the map stay where they are.")
    lay["A2"].font = f(12, True, s["BrandAccent"])
    lay.row_dimensions[2].height = 22
    lheads = ["Aisle", "Side", "Order from fire exit", "Rows", "Pallets deep", "Zone", "Use"]
    for j, (h, w) in enumerate(zip(lheads, [10, 10, 12, 10, 12, 10, 12])):
        cc = lay.cell(5, j + 1, h)
        cc.font = f(11, True, "FFFFFF")
        cc.fill = solid("000000")
        cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        lay.column_dimensions[L(j + 1)].width = w
    lay.row_dimensions[5].height = 32
    for i, a in enumerate(cfg["aisles"]):
        r = 6 + i
        use = "Blocked" if blocked(a.get("Active")) else ("Yes" if str(a.get("Active", "Yes")).lower() in ("yes", "y", "true", "1") else "No")
        vals = [a["AisleID"], a["Side"], int(a["Order"]), int(a["PalletsAcross"]), int(a["PalletsDeep"]),
                a.get("Zone") or "", use]
        for j, v in enumerate(vals):
            cc = lay.cell(r, j + 1, v)
            cc.font = f(12, j == 0)
            cc.alignment = Alignment(horizontal="center", vertical="center")
        lay.row_dimensions[r].height = 22
    lt = Table(displayName="tblLayout", ref=f"A5:G{5 + len(cfg['aisles'])}")
    lt.tableStyleInfo = TableStyleInfo(name="TableStyleMedium15", showRowStripes=True)
    lay.add_table(lt)
    for formula, rng in (('"Left,Right"', "B6:B60"), ('"Yes,Blocked,No"', "G6:G60"),
                         (f'"{",".join(z["Code"] for z in zones)}"', "F6:F60")):
        d = DataValidation(type="list", formula1=formula, allow_blank=True)
        d.add(rng)
        lay.add_data_validation(d)
    d = DataValidation(type="whole", operator="between", formula1="1", formula2="60", allow_blank=False,
                       showErrorMessage=True, error="Use a whole number from 1 to 60.")
    d.add("C6:E60")
    lay.add_data_validation(d)
    notes = [
        "Rows = how many pallets wide the aisle is (1 row = 1 line of pallets on the map).",
        "Pallets deep = how many pallets from the walkway to the wall.",
        "Side: Left or Right, standing with the fire exit behind you. Order: 1 = nearest the fire exit.",
        "Use: Yes = ours,  Blocked = someone else's (grey),  No = hide it.",
        "Zone codes: " + ",  ".join(f"{z['Code']} = {z['Name']}" for z in zones)
        + "   (rename zones on the Stock sheet).",
        "New aisle? Type it on the first empty row under the table.",
    ]
    for i, t_ in enumerate(notes):
        lay.cell(7 + len(cfg["aisles"]) + i, 1, t_).font = f(11, color="404040")

    # ------------------------------------------------------------ Moves log (written by the macros)
    moves.sheet_properties.tabColor = "808080"
    moves.sheet_view.showGridLines = False
    moves.merge_cells("A1:F1")
    moves["A1"].value = "Moves"
    moves["A1"].font = f(18, True, "FFFFFF")
    moves["A1"].alignment = Alignment(vertical="center", indent=1)
    for col in range(1, 7):
        moves.cell(1, col).fill = solid("000000")
        moves.cell(1, col).border = Border(bottom=side("thick", s["BrandAccent"]))
    moves.row_dimensions[1].height = 34
    moves["A2"].value = "Filled in automatically by the macros. Newest at the bottom."
    moves["A2"].font = f(11, color="404040")
    for j, (h, w) in enumerate([("When", 20), ("Who", 16), ("What", 12), ("Item", 14), ("From", 12), ("To", 12)]):
        cc = moves.cell(4, j + 1, h)
        cc.font = f(11, True, "FFFFFF")
        cc.fill = solid("000000")
        cc.alignment = Alignment(horizontal="center")
        moves.column_dimensions[L(j + 1)].width = w
    moves.freeze_panes = "A5"

    # ------------------------------------------------------------ Squares (hidden, used by the macros)
    squares.sheet_state = "hidden"
    for j, h in enumerate(["ID", "Aisle", "Zone", "Row", "Col", "Depth"]):
        squares.cell(1, j + 1, h)
    for i, p in enumerate(geo["positions"]):
        for j, v in enumerate([p["LocationID"], p["Aisle"], p["Zone"], p["MapRow"], p["MapCol"], p["Depth"]]):
            squares.cell(i + 2, j + 1, v)

    wb.active = 0
    wb.save(out)
    print(f"Wrote {out}: {len(geo['positions'])} squares, zoom {zoom}%")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-l", "--layout", default="layout.yaml")
    ap.add_argument("-o", "--out", default="Floor_Map_Simple.xlsx")
    ap.add_argument("--demo", action="store_true")
    a = ap.parse_args()
    build(yaml.safe_load(Path(a.layout).read_text()), a.out, a.demo)


if __name__ == "__main__":
    main()
