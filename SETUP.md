# Consumables Floor Map: setup guide

You receive two files:

| File | What it is |
|---|---|
| `Consumables_Map.xlsx` | The workbook. It works straight away without macros. |
| `FloorMap_VBA.txt` | The macro code, as plain text, that you paste in yourself. |

Nothing needs installing on the work machine. You only need Excel (Microsoft 365, Windows).

---

## 1. Using it without macros

If macros are blocked, the workbook still works:

- **Placing a pallet:** open the **Locations** sheet, find the position (for example `R2-03-B`) and pick an item from the **ItemCode** dropdown. The dropdown only lists items for that position's zone. To empty a position, delete the ItemCode.
- **The Map** updates by itself. A pallet shows as a solid square in its zone colour with the start of its item code. Free space shows as stripes.
- **Dashboard** and **Items** totals, pallet counts, mismatch and low-stock flags all work.
- Some things need the macros: click-to-move, Find Space, Rebuild Layout, the wrong-zone and low-stock borders on the Map, and the pop-up notes.

---

## 2. Adding the macros (one-off, about 5 minutes)

1. Open `Consumables_Map.xlsx`.
2. Go to **File > Save As**. Under *Save as type* choose **Excel Macro-Enabled Workbook (\*.xlsm)** and save it as `Consumables_Map.xlsm`. Work in the `.xlsm` from now on.
3. Press **Alt + F11**. This opens the VBA editor.
4. Open `FloorMap_VBA.txt` in Notepad. It has four parts, each marked with a `PART n START` banner and a `PART n END` banner. Copy only the lines **between** the banners.

**Part 1: the main module**

5. In the VBA editor, click **Insert > Module**.
6. Paste **PART 1** into the empty window.
7. (Optional) Press **F4** and change *(Name)* from `Module1` to `FloorMap`.

**Part 2: the Map sheet**

8. On the left, under *Microsoft Excel Objects*, double-click the sheet that says **(Map)**.
9. Paste **PART 2** into that window. If the window already says `Option Explicit` at the top, don't paste a second copy of that line.

**Part 3: the Locations sheet**

10. Double-click the sheet that says **(Locations)** and paste **PART 3**.

**Part 4: ThisWorkbook (optional)**

11. Double-click **ThisWorkbook** and paste **PART 4**. With this in place, the workbook opens on the Map, repainted and fitted to the screen.

**Draw the buttons**

12. Click **Debug > Compile VBAProject**. If nothing happens, all is well.
13. Close the VBA editor. Press **Alt + F8**, choose **SetupButtons** and click **Run**. Five magenta buttons appear on the Map.
14. Press **Alt + F8** again and run **RefreshMap**.
15. Save (**Ctrl + S**).

Next time you open the file, click **Enable Content** on the yellow bar if Excel shows one.

> If Excel says macros are blocked because the file came from the internet, close it, right-click the file in File Explorer, choose **Properties**, tick **Unblock**, click **OK**, then open it again. If your IT policy blocks all macros, use the no-macro method in section 1.

---

## 3. Everyday use (with macros)

| To do this | Do this |
|---|---|
| **Move a pallet** | Click the pallet (it gets a thick black border), then click an empty square. |
| **Swap two pallets** | Click one pallet, then click the other. |
| **Cancel** | Click the same pallet again, press **Esc**, or press **Cancel Move**. |
| **Find space for a delivery** | Press **Find Space**, type the item code and the number of units. It works out the pallets needed and lists free positions in that item's zone, nearest the bay first. If the zone is full, it offers space in other zones, marked OUT OF ZONE. Answer **Yes** to place the pallets. |
| **Repaint everything** | Press **Refresh** (for example, after changing Items or zone colours). |
| **Fit the map on screen** | Press **Fit to Screen**. |

- If you move a pallet into a different zone, Excel asks you to confirm. You can say Yes.
- Every move, swap and Find Space placement is written to **MoveLog** with the time and your Windows username.
- Hover over a pallet to see its full detail in the note.
- On the Map:
  - a **thick red border** means the pallet is in the wrong zone
  - a **red line along the bottom** means that item is low on stock

---

## 4. Setting it up for real

### Items

On the **Items** sheet, type over or delete the two `EXAMPLE` rows, then add your items in the row just under the table (the table grows by itself).

| Column | What to enter |
|---|---|
| ItemCode | A short code. The Map shows the first 4 characters. |
| Zone | BOX, BAG, FLM or LBL (from the dropdown). |
| UnitsPerPallet | Units on a full pallet. |
| OnHandUnits | Current stock in units. |
| ReorderPallets | Leave blank to use the default low-stock level on Config. |

### Aisle sizes

Once you have measured the aisles:

1. Open **Config**.
2. In the **Aisles** table, change **PalletsDeep** (and **PalletsAcross** if needed).
3. Go to the Map and press **Rebuild Layout**.

Pallets stay where their LocationID (for example `R2-03-B`) still exists. Any that no longer fit are listed in a message and written to MoveLog, so you can put them somewhere else.

In the Aisles table:

| Column | Meaning |
|---|---|
| Side | Left or Right, as seen standing with the fire exit behind you, facing the loading bay. |
| Order | 1 is the aisle nearest the fire exit. |
| GapAfter | Blank walkway rows between this aisle and the next one towards the bay. |
| Active | **Yes** = in use. **No** = hidden without deleting it. **Blocked** = someone else's space: drawn as a grey "NOT OURS" block that can't be clicked or filled. |

Other things to know:

- Each aisle is drawn as *PalletsAcross* rows by *PalletsDeep* squares.
- Depth `01` is next to the main walkway.
- Rows are lettered A, B, C... from the top.
- To add an aisle, add a row to the table.

> **Please check the starting layout.** It follows your floor sketch:
>
> - **Left side**, from the fire exit: XL (someone else's, Blocked), then L1, L2, L3. Each is 4 across x 16 deep.
> - **Right side**, from the fire exit: XR (someone else's, Blocked), then R1 (the narrow one), R2, R3, R4.
>
> These are **guesses** to check and correct on Config:
>
> - R1 is 1 across (confirmed).
> - R2 to R4 are 4 x 16, like the left.
> - The Blocked areas are 6 across (left) and 4 across (right).
> - The zones: L1 to L3 and R4 are Boxes, R1 is Labels and Packing, R2 is Films and Wrap, R3 is Bags.

### Zones and colours

- To add a zone, add a row to the **Zones** table with a code, a name and a hex colour (for example `#7A5CC7`).
- Keep zone colours clearly different from the GXO magenta. **Refresh** warns you if a colour is too close.
- After adding a zone, press **Rebuild Layout** so the legend has room for it.

### Settings

These are in the **Settings** table on Config. Don't sort the Settings table.

| Setting | What it does |
|---|---|
| DefaultLowStock | Pallet count at or below which an item is flagged LOW. |
| CellSizePx / MinCellSizePx | Size of a pallet square. It shrinks automatically, down to the minimum, to fit a 1920x1080 screen. |
| MapLabelChars / ShowMapLabels | How many characters of the item code show on each square, and whether codes show at all. |
| BrandAccent | GXO magenta. The value is approximate, so correct it if you have the official code. |
| Other colours | Header, banding, no-zone greys, alert red, fire exit green. |

### Logo

Paste the official GXO logo on the Map, in the black title bar at the top right. That spot is the named cell **LogoSpot**. Rebuild Layout keeps a picture in the title row lined up with that spot.

---

## 5. Rebuilding the workbook from scratch (off the work machine only)

This needs a PC where you can run Python 3 with `openpyxl` and `PyYAML`, **not** the GXO machine.

```
pip install openpyxl pyyaml
python build_map.py                 # writes Consumables_Map.xlsx from layout.yaml
python build_map.py --demo -o Test.xlsx   # same, with 5 dummy items for testing
```

1. Edit the aisles, zones and settings in `layout.yaml` first.
2. Build the file.
3. Repeat section 2 to add the macros.

A rebuilt workbook starts empty. To keep your data, copy the **Items** table and the Locations **ItemCode** / **UnitsOnPallet** columns across from the old file. For normal changes, the **Rebuild Layout** button inside Excel is easier and keeps your data.

---

## 6. Troubleshooting

| Problem | Fix |
|---|---|
| "Compile error" when compiling | Check that each part went in the right place, and that no banner lines were copied. |
| Clicking the map does nothing | Macros aren't enabled, or PART 2 isn't in the Map sheet's code window. |
| Colours look out of date | Press **Refresh**. |
| Map too big or small | Press **Fit to Screen**, or change CellSizePx on Config and press **Rebuild Layout**. |
| A pallet disappeared after Rebuild Layout | Look at the last rows of **MoveLog**, marked "(removed by Rebuild Layout)". |
