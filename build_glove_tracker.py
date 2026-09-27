"""Builds glove_tracker.xlsx — a red-themed glove stock & issue tracker.

Run:  python build_glove_tracker.py
"""
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

OUT = "glove_tracker.xlsx"
FONT = "Arial"

# Red theme palette
DARK_RED = "7F0000"
RED = "C00000"
MID_RED = "E06666"
LIGHT_RED = "F4CCCC"
PALE = "FFF5F5"
WHITE = "FFFFFF"
GREY = "595959"

SIZES = ["Small", "Medium", "Large"]
N_COLLEAGUES = 200   # rows available on Colleagues sheet
N_LOG = 1000         # rows available on Issue Log
N_RECEIPTS = 300     # rows available for stock deliveries
DATE_FMT = "dd-mmm-yyyy"

thin = Side(style="thin", color="D9D9D9")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def fill(c):
    return PatternFill("solid", start_color=c, end_color=c)


def title(ws, text, subtitle, width_cols):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width_cols)
    c = ws.cell(1, 1, text)
    c.font = Font(name=FONT, size=18, bold=True, color=WHITE)
    c.fill = fill(DARK_RED)
    c.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 36
    for col in range(1, width_cols + 1):
        ws.cell(1, col).fill = fill(DARK_RED)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width_cols)
    s = ws.cell(2, 1, subtitle)
    s.font = Font(name=FONT, size=10, italic=True, color=GREY)
    s.alignment = Alignment(indent=1, wrap_text=True, vertical="center")
    ws.row_dimensions[2].height = 30
    ws.sheet_view.showGridLines = False


def header_row(ws, row, headers, start_col=1):
    for i, h in enumerate(headers):
        c = ws.cell(row, start_col + i, h)
        c.font = Font(name=FONT, bold=True, color=WHITE)
        c.fill = fill(RED)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def body_cell(c, input_cell=False, fmt=None, center=False):
    c.font = Font(name=FONT, color="000000")
    c.border = BORDER
    if input_cell:
        c.fill = fill(PALE)
    if fmt:
        c.number_format = fmt
    if center:
        c.alignment = Alignment(horizontal="center")


def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w


wb = Workbook()
dash = wb.active
dash.title = "Dashboard"
col_ws = wb.create_sheet("Colleagues")
log = wb.create_sheet("Issue Log")
stock = wb.create_sheet("Stock")
for ws in wb.worksheets:
    ws.sheet_properties.tabColor = RED

size_dv_formula = '"' + ",".join(SIZES) + '"'

# ---------------------------------------------------------------- Colleagues
C_FIRST, C_LAST = 5, 5 + N_COLLEAGUES - 1
title(col_ws, "COLLEAGUES",
      "Type names in column B and pick each person's glove size in column C (pink cells). "
      "Everything else fills in automatically from the Issue Log. 3rd pair or more = FLAG.", 10)
header_row(col_ws, 4, ["#", "Name", "Glove Size", "Total Pairs Given",
                       "1st Pair Date", "2nd Pair Date", "3rd Pair Date",
                       "Last Issued", "Status", "Flag #"])
L = "'Issue Log'!"
LD = f"{L}$A$5:$A${4 + N_LOG}"
LN = f"{L}$B$5:$B${4 + N_LOG}"
LQ = f"{L}$D$5:$D${4 + N_LOG}"
LR = f"{L}$F$5:$F${4 + N_LOG}"
for r in range(C_FIRST, C_LAST + 1):
    col_ws.cell(r, 1, r - C_FIRST + 1)
    body_cell(col_ws.cell(r, 1), center=True)
    col_ws.cell(r, 1).font = Font(name=FONT, color=GREY)
    body_cell(col_ws.cell(r, 2), input_cell=True)
    body_cell(col_ws.cell(r, 3), input_cell=True, center=True)
    col_ws.cell(r, 4, f'=IF(B{r}="","",SUMIFS({LQ},{LN},B{r}))')
    body_cell(col_ws.cell(r, 4), center=True)
    for n, col in ((1, 5), (2, 6), (3, 7)):
        col_ws.cell(r, col,
                    f'=IF(OR(B{r}="",N(D{r})<{n}),"",_xlfn.MINIFS({LD},{LN},B{r},{LR},">="&{n}))')
        body_cell(col_ws.cell(r, col), fmt=DATE_FMT, center=True)
    col_ws.cell(r, 8, f'=IF(N(D{r})=0,"",_xlfn.MAXIFS({LD},{LN},B{r}))')
    body_cell(col_ws.cell(r, 8), fmt=DATE_FMT, center=True)
    col_ws.cell(r, 9, f'=IF(B{r}="","",IF(D{r}>=3,"FLAG",IF(D{r}=0,"Not issued","OK")))')
    body_cell(col_ws.cell(r, 9), center=True)
    col_ws.cell(r, 10, f'=IF(I{r}="FLAG",COUNTIF($I${C_FIRST}:I{r},"FLAG"),"")')
    body_cell(col_ws.cell(r, 10), center=True)
    col_ws.cell(r, 10).font = Font(name=FONT, color=GREY, size=8)

widths(col_ws, {"A": 5, "B": 28, "C": 12, "D": 12, "E": 14, "F": 14,
                "G": 14, "H": 14, "I": 12, "J": 7})
col_ws.freeze_panes = "C5"
dv = DataValidation(type="list", formula1=size_dv_formula, allow_blank=True)
dv.add(f"C{C_FIRST}:C{C_LAST}")
col_ws.add_data_validation(dv)
rng = f"A{C_FIRST}:J{C_LAST}"
col_ws.conditional_formatting.add(
    rng, FormulaRule(formula=[f'$I{C_FIRST}="FLAG"'], fill=fill(LIGHT_RED),
                     font=Font(name=FONT, color=DARK_RED, bold=True)))
col_ws.conditional_formatting.add(
    f"I{C_FIRST}:I{C_LAST}",
    CellIsRule(operator="equal", formula=['"FLAG"'], fill=fill(RED),
               font=Font(name=FONT, color=WHITE, bold=True)))
col_ws["J4"].comment = Comment("Helper number used by the Dashboard's flagged list.", "Glove Tracker")

# ---------------------------------------------------------------- Issue Log
G_FIRST, G_LAST = 5, 4 + N_LOG
title(log, "ISSUE LOG",
      "One row each time gloves are handed out, entered in date order. Fill the pink cells: "
      "Date, Name (dropdown from Colleagues), Size, Pairs, Reason. Pair # and Flag are automatic.", 7)
header_row(log, 4, ["Date Given", "Name", "Size", "Pairs Given",
                    "Reason / Notes", "Pair # (running total)", "Flag"])
for r in range(G_FIRST, G_LAST + 1):
    body_cell(log.cell(r, 1), input_cell=True, fmt=DATE_FMT, center=True)
    body_cell(log.cell(r, 2), input_cell=True)
    body_cell(log.cell(r, 3), input_cell=True, center=True)
    body_cell(log.cell(r, 4), input_cell=True, center=True)
    body_cell(log.cell(r, 5), input_cell=True)
    log.cell(r, 6, f'=IF(B{r}="","",SUMIFS($D${G_FIRST}:D{r},$B${G_FIRST}:B{r},B{r}))')
    body_cell(log.cell(r, 6), center=True)
    log.cell(r, 7, f'=IF(B{r}="","",IF(F{r}>=3,"FLAG",""))')
    body_cell(log.cell(r, 7), center=True)
widths(log, {"A": 14, "B": 28, "C": 11, "D": 11, "E": 34, "F": 14, "G": 10})
log.freeze_panes = "A5"

dv_name = DataValidation(type="list", formula1=f"=Colleagues!$B${C_FIRST}:$B${C_LAST}",
                         allow_blank=True, showErrorMessage=True,
                         errorTitle="Unknown name",
                         error="Add this person on the Colleagues sheet first.")
dv_name.add(f"B{G_FIRST}:B{G_LAST}")
dv_size = DataValidation(type="list", formula1=size_dv_formula, allow_blank=True)
dv_size.add(f"C{G_FIRST}:C{G_LAST}")
dv_qty = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1",
                        allow_blank=True, showErrorMessage=True, error="Enter a whole number of pairs (1 or more).")
dv_qty.add(f"D{G_FIRST}:D{G_LAST}")
dv_date = DataValidation(type="date", operator="greaterThan", formula1="1", allow_blank=True,
                         showErrorMessage=True, error="Enter a valid date.")
dv_date.add(f"A{G_FIRST}:A{G_LAST}")
for d in (dv_name, dv_size, dv_qty, dv_date):
    log.add_data_validation(d)
log.conditional_formatting.add(
    f"A{G_FIRST}:G{G_LAST}", FormulaRule(formula=[f'$G{G_FIRST}="FLAG"'], fill=fill(LIGHT_RED),
                                         font=Font(name=FONT, color=DARK_RED, bold=True)))
log.conditional_formatting.add(
    f"G{G_FIRST}:G{G_LAST}",
    CellIsRule(operator="equal", formula=['"FLAG"'], fill=fill(RED),
               font=Font(name=FONT, color=WHITE, bold=True)))

# ---------------------------------------------------------------- Stock
title(stock, "STOCK",
      "Record every delivery of gloves in the Deliveries table (pink cells). "
      "On Hand = Received − Issued, worked out automatically.", 6)
header_row(stock, 4, ["Size", "Received", "Issued", "On Hand", "Status"])
R_FIRST = 13
R_LAST = R_FIRST + N_RECEIPTS - 1
RD = f"$B${R_FIRST}:$B${R_LAST}"
RQ = f"$C${R_FIRST}:$C${R_LAST}"
LS = f"{L}$C$5:$C${4 + N_LOG}"
for i, size in enumerate(SIZES):
    r = 5 + i
    stock.cell(r, 1, size)
    stock.cell(r, 2, f"=SUMIFS({RQ},{RD},A{r})")
    stock.cell(r, 3, f"=SUMIFS({LQ},{LS},A{r})")
    stock.cell(r, 4, f"=B{r}-C{r}")
    stock.cell(r, 5, f'=IF(D{r}<=0,"OUT OF STOCK",IF(D{r}<$B$10,"LOW","OK"))')
    for c in range(1, 6):
        body_cell(stock.cell(r, c), center=c > 1)
    stock.cell(r, 1).font = Font(name=FONT, bold=True)
stock.cell(8, 1, "Total")
for c, col in ((2, "B"), (3, "C"), (4, "D")):
    stock.cell(8, c, f"=SUM({col}5:{col}7)")
for c in range(1, 6):
    body_cell(stock.cell(8, c), center=c > 1)
    stock.cell(8, c).font = Font(name=FONT, bold=True)
    stock.cell(8, c).fill = fill(LIGHT_RED)

stock.cell(10, 1, "Low-stock alert below")
stock.cell(10, 1).font = Font(name=FONT, bold=True)
stock.cell(10, 2, 10)
body_cell(stock.cell(10, 2), input_cell=True, center=True)
stock.cell(10, 3, "pairs per size (agreed default — change as needed)")
stock.cell(10, 3).font = Font(name=FONT, italic=True, color=GREY)

stock.cell(R_FIRST - 1, 1, "DELIVERIES RECEIVED")
header_row(stock, R_FIRST - 1, ["Date Received", "Size", "Pairs Received", "Supplier / Notes"])
for r in range(R_FIRST, R_LAST + 1):
    body_cell(stock.cell(r, 1), input_cell=True, fmt=DATE_FMT, center=True)
    body_cell(stock.cell(r, 2), input_cell=True, center=True)
    body_cell(stock.cell(r, 3), input_cell=True, center=True)
    body_cell(stock.cell(r, 4), input_cell=True)
dv_s2 = DataValidation(type="list", formula1=size_dv_formula, allow_blank=True)
dv_s2.add(f"B{R_FIRST}:B{R_LAST}")
dv_q2 = DataValidation(type="whole", operator="greaterThanOrEqual", formula1="1", allow_blank=True)
dv_q2.add(f"C{R_FIRST}:C{R_LAST}")
stock.add_data_validation(dv_s2)
stock.add_data_validation(dv_q2)
widths(stock, {"A": 22, "B": 14, "C": 16, "D": 30, "E": 16})
for cond, color, fcolor in (('"OUT OF STOCK"', DARK_RED, WHITE), ('"LOW"', MID_RED, WHITE)):
    stock.conditional_formatting.add(
        "E5:E7", CellIsRule(operator="equal", formula=[cond], fill=fill(color),
                            font=Font(name=FONT, color=fcolor, bold=True)))

# ---------------------------------------------------------------- Dashboard
title(dash, "GLOVE TRACKER — DASHBOARD",
      "Live summary. Add names on Colleagues, deliveries on Stock, and hand-outs on Issue Log — "
      "this page updates by itself. Anyone on their 3rd pair or more is FLAGGED.", 12)
widths(dash, {c: 12 for c in "ABCDEFGHIJKL"})
dash.column_dimensions["A"].width = 3
dash.column_dimensions["F"].width = 15


def tile(row, col, label, formula, span=2, big_color=RED):
    dash.merge_cells(start_row=row, start_column=col, end_row=row, end_column=col + span - 1)
    dash.merge_cells(start_row=row + 1, start_column=col, end_row=row + 2, end_column=col + span - 1)
    lab = dash.cell(row, col, label)
    lab.font = Font(name=FONT, bold=True, color=WHITE, size=10)
    lab.alignment = Alignment(horizontal="center", vertical="center")
    val = dash.cell(row + 1, col, formula)
    val.font = Font(name=FONT, bold=True, size=24, color=big_color)
    val.alignment = Alignment(horizontal="center", vertical="center")
    for r in range(row, row + 3):
        for c in range(col, col + span):
            dash.cell(r, c).fill = fill(RED if r == row else PALE)
            dash.cell(r, c).border = BORDER
    return val


dash.row_dimensions[4].height = 22
tile(4, 2, "SMALL ON HAND", "=Stock!D5")
tile(4, 4, "MEDIUM ON HAND", "=Stock!D6")
tile(4, 6, "LARGE ON HAND", "=Stock!D7")
tile(4, 8, "TOTAL PAIRS ISSUED", "=Stock!C8", big_color=DARK_RED)
tile(4, 10, "FLAGGED PEOPLE", f'=COUNTIF(Colleagues!I{C_FIRST}:I{C_LAST},"FLAG")', big_color=DARK_RED)

tile(8, 2, "COLLEAGUES LISTED", f'=COUNTA(Colleagues!B{C_FIRST}:B{C_LAST})', big_color=GREY)
tile(8, 4, "GIVEN 1 PAIR", f'=COUNTIF(Colleagues!D{C_FIRST}:D{C_LAST},1)', big_color=GREY)
tile(8, 6, "GIVEN 2 PAIRS", f'=COUNTIF(Colleagues!D{C_FIRST}:D{C_LAST},2)', big_color=GREY)
tile(8, 8, "NOT YET ISSUED", f'=COUNTIF(Colleagues!I{C_FIRST}:I{C_LAST},"Not issued")', big_color=GREY)
tile(8, 10, "FLAGGED HAND-OUTS", f'=COUNTIF({L}G5:G{4 + N_LOG},"FLAG")', big_color=DARK_RED)
for r in (5, 6, 9, 10):
    dash.row_dimensions[r].height = 22

# Stock status table
header_row(dash, 12, ["Size", "Received", "Issued", "On Hand", "Status"], start_col=2)
for i in range(3):
    r = 13 + i
    for j, col in enumerate("ABCDE"):
        c = dash.cell(r, 2 + j, f"=Stock!{col}{5 + i}")
        body_cell(c, center=j > 0)
dash.cell(16, 2, "Total")
for j, col in enumerate("BCD"):
    dash.cell(16, 3 + j, f"=Stock!{col}8")
for c in range(2, 7):
    body_cell(dash.cell(16, c), center=c > 2)
    dash.cell(16, c).font = Font(name=FONT, bold=True)
    dash.cell(16, c).fill = fill(LIGHT_RED)
for cond, color in (('"OUT OF STOCK"', DARK_RED), ('"LOW"', MID_RED)):
    dash.conditional_formatting.add(
        "F13:F15", CellIsRule(operator="equal", formula=[cond], fill=fill(color),
                              font=Font(name=FONT, color=WHITE, bold=True)))
dash.cell(17, 2, '="Low-stock alert below "&Stock!B10&" pairs per size"')
dash.cell(17, 2).font = Font(name=FONT, italic=True, color=GREY, size=9)

# Chart: on hand vs issued by size
chart = BarChart()
chart.type = "col"
chart.title = "Stock by Size"
chart.y_axis.title = "Pairs"
chart.height, chart.width = 7.5, 13
data = Reference(stock, min_col=3, max_col=4, min_row=4, max_row=7)
cats = Reference(stock, min_col=1, min_row=5, max_row=7)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.series[0].graphicalProperties.solidFill = MID_RED
chart.series[1].graphicalProperties.solidFill = DARK_RED
chart.legend.position = "b"
dash.add_chart(chart, "H12")

# Flagged list
FL = 29
dash.merge_cells(start_row=FL - 1, start_column=2, end_row=FL - 1, end_column=11)
h = dash.cell(FL - 1, 2, "⚑ FLAGGED — 3rd PAIR OR MORE")
h.font = Font(name=FONT, bold=True, color=WHITE, size=12)
h.fill = fill(DARK_RED)
for c in range(2, 12):
    dash.cell(FL - 1, c).fill = fill(DARK_RED)
flag_headers = ["#", "Name", "Size", "Total Pairs", "1st Pair", "2nd Pair", "3rd Pair", "Last Issued"]
dash.row_dimensions[FL].height = 30
dash.merge_cells(start_row=FL, start_column=3, end_row=FL, end_column=4)
# columns: B #, C:D Name (merged), E Size, F Total, G 1st, H 2nd, I 3rd, J Last
N_FLAG_ROWS = 25
src_cols = {"C": "B", "E": "C", "F": "D", "G": "E", "H": "F", "I": "G", "J": "H"}
for col, text in zip("BCEFGHIJ", flag_headers):
    dash[f"{col}{FL}"] = text
    dash[f"{col}{FL}"].font = Font(name=FONT, bold=True, color=WHITE)
    dash[f"{col}{FL}"].fill = fill(RED)
    dash[f"{col}{FL}"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for k in range(1, N_FLAG_ROWS + 1):
    r = FL + k
    dash.cell(r, 2, f'=IF({k}>$J$5,"",{k})')
    body_cell(dash.cell(r, 2), center=True)
    dash.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    for dcol, scol in src_cols.items():
        f = (f'=IF({k}>$J$5,"",INDEX(Colleagues!${scol}${C_FIRST}:${scol}${C_LAST},'
             f'MATCH({k},Colleagues!$J${C_FIRST}:$J${C_LAST},0)))')
        dash[f"{dcol}{r}"] = f
        fmt = DATE_FMT if dcol in "GHIJ" else None
        body_cell(dash[f"{dcol}{r}"], fmt=fmt, center=dcol != "C")
    if k % 2 == 0:
        for c in "BCDEFGHIJ":
            dash[f"{c}{r}"].fill = fill(PALE)
    dash[f"I{r}"].font = Font(name=FONT, bold=True, color=RED)
dash.cell(FL + N_FLAG_ROWS + 1, 2,
          f'=IF($J$5>{N_FLAG_ROWS},"+ "&($J$5-{N_FLAG_ROWS})&" more — see Colleagues sheet (filter Status = FLAG)",'
          f'IF($J$5=0,"No one flagged yet.",""))')
dash.cell(FL + N_FLAG_ROWS + 1, 2).font = Font(name=FONT, italic=True, color=GREY)

# How-to legend
dash.cell(19, 2, "HOW TO USE").font = Font(name=FONT, bold=True, color=DARK_RED)
steps = [
    "1. Colleagues — type each name and pick their glove size (pink cells).",
    "2. Stock — log every delivery: date, size, pairs received.",
    "3. Issue Log — each hand-out: date, name (dropdown), size, pairs, reason.",
    "4. Counts never reset. 1st & 2nd pair = OK; 3rd pair or more = FLAG (red).",
    "   Pink cells = type here. White cells = automatic, don't overwrite.",
]
for i, s in enumerate(steps):
    c = dash.cell(20 + i, 2, s)
    c.font = Font(name=FONT, size=9, color=GREY)

for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 100
dash.sheet_view.zoomScale = 90
wb.active = 0
wb.save(OUT)
print("saved", OUT)
