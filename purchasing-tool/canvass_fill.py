"""
canvass_fill.py
Fills TEMP.xlsx canvass template with supplied JSON data.
Copies the template, inserts item rows, fills summary fields, saves new file.

Usage: python3 canvass_fill.py <data.json> <template.xlsx> <output.xlsx>
"""

import sys, json, shutil, copy
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.styles.borders import Border, Side

def thin():
    s = Side(style="thin")
    return Border(left=s, right=s, top=s, bottom=s)

def copy_row_style(ws, src_row, dst_row):
    """Copy cell styles from src_row to dst_row."""
    for col in range(1, ws.max_column + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        if src.has_style:
            dst.font      = copy.copy(src.font)
            dst.fill      = copy.copy(src.fill)
            dst.border    = copy.copy(src.border)
            dst.alignment = copy.copy(src.alignment)
            dst.number_format = src.number_format
    # Also copy row height, so inserted item rows match the template row
    # instead of falling back to Excel's default (this is part of why rows
    # looked randomly too small/tall after export)
    if src_row in ws.row_dimensions:
        ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def shift_merges_and_insert_rows(ws, insert_at, n_rows):
    """
    openpyxl's insert_rows() moves cell VALUES down correctly but does NOT
    move MergedCellRange definitions. Any merge at or below `insert_at` has
    to be manually unmerged, the rows inserted, then re-merged at its new
    shifted position — otherwise the old merge boundaries end up overlapping
    the newly written rows, which is what was causing Excel's "repaired /
    removed merge cells" warning and the missing row-15 fields.
    """
    if n_rows <= 0:
        return

    ranges_to_shift = []
    ranges_to_keep  = []
    for mc in list(ws.merged_cells.ranges):
        coords = (mc.min_row, mc.min_col, mc.max_row, mc.max_col)
        if mc.min_row >= insert_at:
            ranges_to_shift.append(coords)
        else:
            ranges_to_keep.append(coords)

    # Unmerge everything first so insert_rows() isn't juggling stale merges
    for mc in list(ws.merged_cells.ranges):
        ws.unmerge_cells(start_row=mc.min_row, start_column=mc.min_col,
                          end_row=mc.max_row,   end_column=mc.max_col)

    ws.insert_rows(insert_at, amount=n_rows)

    # Re-apply merges that were above the insertion point — unchanged
    for (r1, c1, r2, c2) in ranges_to_keep:
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    # Re-apply merges that were at/below the insertion point — shifted down
    for (r1, c1, r2, c2) in ranges_to_shift:
        ws.merge_cells(start_row=r1 + n_rows, start_column=c1,
                        end_row=r2 + n_rows,   end_column=c2)


def to_number(val):
    """Best-effort parse into a real float. Returns None if it isn't a number
    (blank, or text like 'FREE') — used to keep numeric cells actually numeric
    instead of text, and to avoid #VALUE! errors when doing arithmetic."""
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return val
    s = str(val).replace(",", "").strip()
    if s == "":
        return None
    try:
        return float(s)
    except ValueError:
        return None

def main():
    data_path, tmpl_path, out_path = sys.argv[1], sys.argv[2], sys.argv[3]

    with open(data_path) as f:
        d = json.load(f)

    # ── Unpack payload ────────────────────────────────────────────────────────
    title      = d.get("title", "")
    ro         = d.get("ro", "")
    date_val   = d.get("date", "")
    remarks    = d.get("remarks", "")
    items      = d.get("items", [])   # [{desc, qty, unit, p1, p2, p3}, ...]

    suppliers  = d.get("suppliers", [{}, {}, {}])  # [{name,loc,contact,num}, ...]
    while len(suppliers) < 3:
        suppliers.append({})

    dl         = d.get("delivery", ["", "", ""])
    discount   = d.get("discount", ["", "", ""])
    credit     = d.get("credit",   ["", "", ""])
    vat        = d.get("vat",      ["", "", ""])
    avail      = d.get("avail",    ["", "", ""])
    da         = d.get("da",       ["", "", ""])

    # ── Copy template ─────────────────────────────────────────────────────────
    shutil.copy2(tmpl_path, out_path)
    wb = load_workbook(out_path)
    ws = wb.active  # Sheet1

    # ── Fill header fields ────────────────────────────────────────────────────
    # Row 2: item/title
    ws["K2"] = title or "-ITEM NAME OR SUPPLIER TITLE HERE-"

    # Row 13: title + RO (A13:F13 merged)
    title_ro = title
    if ro: title_ro += f"\nRO: {ro}"
    else:  title_ro += "\nRO: IF NO RO LEAVE BLANK"
    ws["A13"] = title_ro

    for i, (col, sup) in enumerate([(11, suppliers[0]), (14, suppliers[1]), (17, suppliers[2])]):
        # Name already appears in the row-11 header directly above this
        # block — don't repeat it here, just location / contact / number.
        info = "\n".join(filter(None, [
            sup.get("loc", ""),
            sup.get("contact", ""),
            sup.get("num", ""),
        ])) or "Location\nContact person\nTheir Number"
        ws.cell(row=12, column=col).value = info

    # Row 11: supplier name headers
    ws["K11"] = suppliers[0].get("name", "") or "SUPPLIER 1"
    ws["N11"] = suppliers[1].get("name", "") or "SUPPLIER 2"
    # Supplier 3 — template only has 2 by default; we'll use col Q if present
    # (Template has 2 suppliers; we keep that layout)

    # ── Insert item rows above row 15 (NOTHING FOLLOWS) ──────────────────────
    # Template has 1 sample row at row 14. We'll overwrite it with real data
    # and insert extras as needed.

    n_items = len(items)
    if n_items == 0:
        items = [{"desc": "", "qty": 1, "unit": "pce", "p1": "", "p2": "", "p3": ""}]
        n_items = 1

    # Insert (n_items - 1) extra rows after row 14 by shifting rows down
    if n_items > 1:
        shift_merges_and_insert_rows(ws, 15, n_items - 1)
        # Copy style from row 14 to the new rows
        for extra in range(1, n_items):
            copy_row_style(ws, 14, 14 + extra)
            # This B:D merge is new (row 14 is the only one that had it
            # originally) so it's created fresh rather than shifted
            ws.merge_cells(
                start_row=14 + extra, start_column=2,
                end_row=14 + extra,   end_column=4
            )

    # ── Fill item rows ────────────────────────────────────────────────────────
    s1_totals = []
    s2_totals = []

    for idx, item in enumerate(items):
        r = 14 + idx
        desc  = item.get("desc", "")
        qty   = item.get("qty", 1)
        unit  = item.get("unit", "pce")
        p1    = item.get("p1", "")
        p2    = item.get("p2", "")
        p3    = item.get("p3", "")

        qty_num = to_number(qty)
        p1_num  = to_number(p1)
        p2_num  = to_number(p2)

        ws.cell(row=r, column=1).value  = idx + 1                              # Item No
        ws.cell(row=r, column=2).value  = desc                                 # Description (B, merged to D)
        ws.cell(row=r, column=5).value  = qty_num if qty_num is not None else qty  # QTY as real number
        ws.cell(row=r, column=6).value  = unit                                 # UNIT
        ws.cell(row=r, column=11).value = p1_num                               # Supplier 1 price (real number)
        if p1_num is not None:
            ws.cell(row=r, column=12).value = f"=K{r}*E{r}"
            ws.cell(row=r, column=13).value = f"=K{r}*E{r}"
        ws.cell(row=r, column=14).value = p2_num                               # Supplier 2 price (real number)
        if p2_num is not None:
            ws.cell(row=r, column=15).value = f"=N{r}*E{r}"
            ws.cell(row=r, column=16).value = f"=N{r}*E{r}"

        s1_totals.append((p1_num or 0) * (qty_num or 0))
        s2_totals.append((p2_num or 0) * (qty_num or 0))

        # Apply borders
        for col in range(1, 17):
            ws.cell(row=r, column=col).border = thin()

    # After inserting rows, NOTHING FOLLOWS is at row 14 + n_items
    nf_row = 14 + n_items
    # Summary rows shift accordingly — fill them
    # Row offsets from NOTHING FOLLOWS (nf_row):
    #   nf_row    = NOTHING FOLLOWS
    #   nf_row+1  = Remarks / DELIVERY CHARGE
    #   nf_row+2  = TOTAL
    #   nf_row+3  = DISCOUNT
    #   nf_row+4  = GRAND TOTAL
    #   nf_row+5  = CREDIT TERMS
    #   nf_row+6  = VAT
    #   nf_row+7  = AVAILABILITY
    #   nf_row+8  = DELIVERY ADDRESS

    def sfill(row_offset, col, value):
        ws.cell(row=nf_row + row_offset, column=col).value = value

    delivery1 = to_number(dl[0])
    delivery2 = to_number(dl[1])
    discount1 = to_number(discount[0])
    discount2 = to_number(discount[1])

    total1 = sum(s1_totals)
    total2 = sum(s2_totals)

    grand1 = total1 + (delivery1 or 0) - (discount1 or 0)
    grand2 = total2 + (delivery2 or 0) - (discount2 or 0)

    # Delivery charge — real number if one was given, otherwise the "FREE" label
    sfill(1, 11, delivery1 if delivery1 is not None else (dl[0] or "FREE"))
    sfill(1, 14, delivery2 if delivery2 is not None else (dl[1] or "FREE"))

    # Total — plain computed number (calculated in Python, not a live SUM formula)
    sfill(2, 11, total1)
    sfill(2, 14, total2)

    # Discount
    sfill(3, 11, discount1 if discount1 is not None else (discount[0] or ""))
    sfill(3, 14, discount2 if discount2 is not None else (discount[1] or ""))

    # Grand total — plain computed number. This is what was throwing #VALUE!
    # before: it used to be a formula adding a numeric TOTAL cell to a
    # delivery-charge cell that often just contained the text "FREE".
    sfill(4, 11, grand1)
    sfill(4, 14, grand2)

    # Credit terms
    sfill(5, 11, credit[0] or "")
    sfill(5, 14, credit[1] or "")

    # VAT
    sfill(6, 11, vat[0] or "VAT INCLUSIVE")
    sfill(6, 14, vat[1] or "VAT INCLUSIVE")

    # Availability
    sfill(7, 11, avail[0] or "")
    sfill(7, 14, avail[1] or "")

    # Delivery address
    sfill(8, 11, da[0] or "")
    sfill(8, 14, da[1] or "")

    # Remarks (A16:C23 merged area)
    sfill(1, 1, remarks or "")

    wb.save(out_path)
    print(f"Saved: {out_path}")

if __name__ == "__main__":
    main()
