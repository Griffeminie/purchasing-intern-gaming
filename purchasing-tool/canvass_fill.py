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

    # Row 12: supplier company info (K12:M12 and N12:P12)
    for i, (col, sup) in enumerate([(11, suppliers[0]), (14, suppliers[1]), (17, suppliers[2])]):
        info = "\n".join(filter(None, [
            sup.get("name", ""),
            sup.get("loc", ""),
            sup.get("contact", ""),
            sup.get("num", ""),
        ])) or "COMPANY NAME\nLocation\nContact person\nTheir Number"
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
        ws.insert_rows(15, amount=n_items - 1)
        # Copy style from row 14 to the new rows
        for extra in range(1, n_items):
            copy_row_style(ws, 14, 14 + extra)
            # Re-apply B:D merge on each new row
            ws.merge_cells(
                start_row=14 + extra, start_column=2,
                end_row=14 + extra,   end_column=4
            )

    # ── Fill item rows ────────────────────────────────────────────────────────
    for idx, item in enumerate(items):
        r = 14 + idx
        desc  = item.get("desc", "")
        qty   = item.get("qty", 1)
        unit  = item.get("unit", "pce")
        p1    = item.get("p1", "")
        p2    = item.get("p2", "")
        p3    = item.get("p3", "")

        ws.cell(row=r, column=1).value  = idx + 1          # Item No
        ws.cell(row=r, column=2).value  = desc              # Description (B, merged to D)
        ws.cell(row=r, column=5).value  = qty               # QTY
        ws.cell(row=r, column=6).value  = unit              # UNIT
        ws.cell(row=r, column=11).value = p1 if p1 != "" else None  # Supplier 1 price
        # Total S1 = price * qty
        if p1 != "" and p1 is not None:
            try:
                ws.cell(row=r, column=12).value = f"=K{r}*E{r}"
                ws.cell(row=r, column=13).value = f"=K{r}*E{r}"
            except: pass
        ws.cell(row=r, column=14).value = p2 if p2 != "" else None  # Supplier 2 price
        if p2 != "" and p2 is not None:
            try:
                ws.cell(row=r, column=15).value = f"=N{r}*E{r}"
                ws.cell(row=r, column=16).value = f"=N{r}*E{r}"
            except: pass

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

    # Delivery charge
    sfill(1, 11, dl[0] or "FREE")
    sfill(1, 14, dl[1] or "FREE")

    # Total (formulas referencing price cols)
    first_item_row = 14
    last_item_row  = 14 + n_items - 1
    sfill(2, 11, f"=SUM(L{first_item_row}:L{last_item_row})")
    sfill(2, 14, f"=SUM(O{first_item_row}:O{last_item_row})")

    # Discount
    sfill(3, 11, discount[0] or "")
    sfill(3, 14, discount[1] or "")

    # Grand total (total + delivery - discount — leave as text formula ref)
    sfill(4, 11, f"={get_column_letter(11)}{nf_row+2}+{get_column_letter(11)}{nf_row+1}")
    sfill(4, 14, f"={get_column_letter(14)}{nf_row+2}+{get_column_letter(14)}{nf_row+1}")

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
