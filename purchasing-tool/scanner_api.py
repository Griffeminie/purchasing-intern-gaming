#!/usr/bin/env python3
"""
scanner_api.py  —  Purchasing Tool PDF extractor bridge
Reads PDF paths from stdin as JSON, returns extracted rows as JSON to stdout.

Usage:
    echo '["path/to/po.pdf"]' | python scanner_api.py
"""

import sys
import json
import re
from pathlib import Path

try:
    import fitz
except ImportError:
    print(json.dumps({"success": False, "error": "PyMuPDF not installed. Run: pip install pymupdf"}))
    sys.exit(1)

try:
    import pytesseract
    from PIL import Image
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

MIN_NATIVE_CHARS = 40

# ── Text extraction ───────────────────────────────────────────────────────────

def _ocr_page(page):
    if not OCR_AVAILABLE:
        return ""
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return pytesseract.image_to_string(img, config="--psm 6 --oem 3")

def extract_pdf_text(pdf_path):
    doc  = fitz.open(pdf_path)
    parts  = []
    method = "native"
    for page in doc:
        text = page.get_text("text", sort=True)
        if len(text.strip()) < MIN_NATIVE_CHARS:
            text = _ocr_page(page)
            method = "ocr"
        parts.append(text)
    doc.close()
    return "\n".join(parts), method

# ── Fallback: rebuild reading order from word bounding boxes ──────────────────
# Some PDFs (typically report-engine templates like the one used here) store
# each TABLE COLUMN as its own text block — every "ITEM NO" value stacked
# together, then every "QTY", then every "DESCRIPTION", etc. PyMuPDF's normal
# `sort=True` extraction orders text block-by-block, so it dumps each column
# as a run of consecutive lines instead of interleaving them row by row. The
# regex-based item parser expects one row per line, so on documents like this
# it silently matches nothing for the numbered items (while a simple two-value
# "label amount amount" charge line still happens to survive intact) — meaning
# only the charges show up and every numbered item goes missing.
#
# This rebuilds text by clustering words into visual rows using their actual
# Y position on the page, then sorting each row's words left-to-right by X —
# i.e. reconstructing the row PyMuPDF's block order threw away.
def _extract_words_as_lines(page, y_tol=3):
    words = page.get_text("words")  # (x0, y0, x1, y1, text, block_no, line_no, word_no)
    if not words:
        return ""
    words = sorted(words, key=lambda w: (w[1], w[0]))
    lines, bucket, bucket_top = [], [], None
    for w in words:
        y0 = w[1]
        if bucket and (y0 - bucket_top) > y_tol:
            bucket.sort(key=lambda x: x[0])
            lines.append(" ".join(x[4] for x in bucket))
            bucket, bucket_top = [], None
        bucket.append(w)
        bucket_top = y0 if bucket_top is None else min(bucket_top, y0)
    if bucket:
        bucket.sort(key=lambda x: x[0])
        lines.append(" ".join(x[4] for x in bucket))
    return "\n".join(lines)

def _amount_sum(items):
    total = 0.0
    for it in items:
        try:
            total += float((it.get("amount") or "0").replace(",", ""))
        except ValueError:
            pass
    return total

# Heuristic: does this extraction look like it missed the numbered item rows?
def _looks_incomplete(text, header, items):
    regular = [it for it in items if not it.get("is_charge")]
    # Signal 1 — numbered item markers ("001", "002", …) appear in the raw
    # text, meaning the table exists, but none of them were parsed into rows.
    has_item_markers = bool(re.search(r"(?m)^0\d{2}\s*$", text))
    if has_item_markers and not regular:
        return True
    # Signal 2 — extracted line total is far off from the PO's printed total.
    stated = header.get("TOTAL_AMOUNT")
    if stated:
        try:
            stated_val = float(stated.replace(",", ""))
            if stated_val > 0 and abs(_amount_sum(items) - stated_val) / stated_val > 0.05:
                return True
        except ValueError:
            pass
    return False

# Did the reflowed re-parse actually do better than the original?
def _is_better(retry_items, orig_items, header):
    stated = header.get("TOTAL_AMOUNT")
    if stated:
        try:
            stated_val = float(stated.replace(",", ""))
            return abs(_amount_sum(retry_items) - stated_val) < abs(_amount_sum(orig_items) - stated_val)
        except ValueError:
            pass
    retry_count = len([it for it in retry_items if not it.get("is_charge")])
    orig_count  = len([it for it in orig_items  if not it.get("is_charge")])
    return retry_count > orig_count

# ── Header parsing ────────────────────────────────────────────────────────────

def _clean(s):
    return re.sub(r"\s+", " ", s or "").strip()

def _parse_amount(s):
    """Convert '1,298.00' or '-1,298.00' to float string, preserving sign."""
    s = s.replace(",", "").strip()
    try:
        return str(float(s))
    except ValueError:
        return s

def extract_header(text):
    h = {}

    # Supplier + PO Number on same line as PURCHASE ORDER
    m = re.search(r"^(.*?)\s+PURCHASE\s+ORDER\s+([A-Z0-9\-]{6,20})\s*$", text, re.M | re.I)
    if m:
        supplier = _clean(m.group(1))
        if not re.search(r"\bDII\b|Tel\.|Fax|CALIFORNIA CLOTHING", supplier, re.I):
            h["SUPPLIER"] = supplier
        h["PO_NUMBER"] = m.group(2)

    # PO Date
    m = re.search(r"\bDATE:\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})", text)
    if m: h["PO_DATE"] = m.group(1)

    # RO NO / JOB ORDER
    m = re.search(r"RO NO:\s*JOB ORDER NO:\s*\n\s*([^\n]+)", text, re.I)
    if m:
        h["RO_NO"] = _clean(m.group(1))
    else:
        m = re.search(r"RO NO:\s*\n\s*([^\n]+)", text, re.I)
        if m: h["RO_NO"] = _clean(m.group(1))

    # Cost center
    m = re.search(r"COST CENTER:\s*ACCOUNT NO:\s*\n([^\n]*)", text, re.I)
    if not m: m = re.search(r"COST CENTER:\s*\n([^\n]*)", text, re.I)
    if m:
        cc = re.search(r"(\d{3,8}(?:[;,]\s*\d{3,8})*;?)\s*$", m.group(1))
        if cc: h["COST_CENTER"] = cc.group(1)

    # Deliver not later than
    m = re.search(r"DELIVER NOT LATER THAN[\s\S]{0,80}?([0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4})", text, re.I)
    if m: h["DELIVER_BY"] = m.group(1)

    # Terms of payment
    m = re.search(r"Terms of payment:?\s*([^\n\*]+)", text, re.I)
    if m: h["TERMS"] = _clean(m.group(1))

    # Total Amount (from PDF footer — for reference only)
    m = re.search(r"Total Amount:?\s*(?:PHP)?\s*([\d,]+\.\d{2})", text, re.I)
    if m: h["TOTAL_AMOUNT"] = m.group(1)

    return h

# ── Line item parsing ─────────────────────────────────────────────────────────

# Standard item row: "001  1  UNI  DESCRIPTION  26,998.00  26,998.00"
_ITEM_FIRST_LINE = re.compile(
    r"^(0\d{2})\s+"
    r"([\d,]+(?:\.\d+)?)\s+"
    r"([A-Z]{2,6})\s+"
    r"(.+?)\s+"
    r"([\d,]+\.\d{2,4})\s+"
    r"([\d,]+\.\d{2})\s*$"
)

# Charge line: matches any text label followed by two amount values
# e.g. "Service Charge - Distributed   698.00   698.00"
# e.g. "Discount Amount   -1,298.00   -1,298.00"
_CHARGE_LINE = re.compile(
    r"^(.+?)\s+([\-\d,]+\.\d{2})\s+([\-\d,]+\.\d{2})\s*$"
)


def parse_items(text):
    items = []

    # ── 1. Standard numbered items ────────────────────────────────────────
    blocks = re.split(r"(?=^0\d{2}\s)", text, flags=re.M)
    for block in blocks:
        block = block.strip("\n")
        if not block.strip(): continue
        lines = block.split("\n")
        m = _ITEM_FIRST_LINE.match(lines[0].strip())
        if not m: continue

        item_no, qty, unit, desc_and_code, unit_price, amount = m.groups()

        code_m = re.match(r"([A-Z]{2}-[A-Z0-9]+-\d+)\s*(.*)", desc_and_code)
        if code_m:
            code = code_m.group(1)
            desc = code_m.group(2).strip(" -")
        else:
            code = ""
            desc = desc_and_code.strip()

        # Continuation lines — stop at Charges/Discount section
        extra = []
        for ln in lines[1:]:
            ln = ln.strip()
            if not ln: continue
            if re.match(r"^(Charges:|Payment Instruction|Delivery Instruction|"
                        r"Terms of payment|Total Amount|PLEASE NOTIFY|"
                        r"The acceptance)", ln, re.I):
                break
            if re.match(r"^\d{7,}", ln): continue
            extra.append(ln)

        if extra:
            desc = _clean(desc + " " + " ".join(extra))

        items.append({
            "no":         item_no,
            "qty":        qty.replace(",", ""),
            "unit":       unit,
            "code":       code,
            "desc":       _clean(desc),
            "unit_price": unit_price.replace(",", ""),
            "amount":     amount.replace(",", ""),
            "is_charge":  False,
        })

    # ── 2. Charges section (after the numbered items) ─────────────────────
    # Find the "Charges:" section in the text
    charges_match = re.search(r"Charges:\s*\n([\s\S]*?)(?:^Discount:\s*$|Total Amount:|Terms of payment:)",
                               text, re.I | re.M)

    if charges_match:
        charges_block = charges_match.group(1)
        for line in charges_block.split("\n"):
            line = line.strip()
            if not line: continue
            m = _CHARGE_LINE.match(line)
            if m:
                label      = _clean(m.group(1))
                unit_price = _parse_amount(m.group(2))
                amount     = _parse_amount(m.group(3))
                items.append({
                    "no":         "",
                    "qty":        "1",
                    "unit":       "LOT",
                    "code":       "",
                    "desc":       label,
                    "unit_price": unit_price,
                    "amount":     amount,
                    "is_charge":  True,
                })

    # ── 3. Discount line ──────────────────────────────────────────────────
    # Look for discount anywhere in the text — usually near the total
    # Pattern: "Discount Amount   -1,298.00   -1,298.00"
    # or       "Discount:  -1,298.00"
    discount_section = re.search(
        r"(Discount\s*(?:Amount)?[:\-]?\s*)([\-\d,]+\.\d{2})(?:\s+([\-\d,]+\.\d{2}))?",
        text, re.I
    )
    if discount_section:
        raw_price  = discount_section.group(2).replace(",", "")
        raw_amount = (discount_section.group(3) or discount_section.group(2)).replace(",", "")

        # Only add if it's actually negative (a real discount)
        # Skip if already captured as a charge line
        already_have_discount = any(
            "discount" in item["desc"].lower() for item in items if item.get("is_charge")
        )
        if not already_have_discount:
            try:
                price_val = float(raw_price)
                # Make negative if not already (discount should reduce total)
                if price_val > 0:
                    price_val  = -price_val
                    raw_price  = str(price_val)
                    raw_amount = str(price_val)
            except ValueError:
                pass

            items.append({
                "no":         "",
                "qty":        "1",
                "unit":       "LOT",
                "code":       "",
                "desc":       "Discount",
                "unit_price": raw_price,
                "amount":     raw_amount,
                "is_charge":  True,
            })

    return items

# ── Build monitoring rows ─────────────────────────────────────────────────────

def build_rows(header, items):
    rows = []
    n = len(items) if items else 1
    for i in range(n):
        it = items[i] if items else {}
        row = {
            "PR DATE":               "",
            "PR DATE RECEIVED":      "",
            "PR NO.":                header.get("RO_NO", ""),
            "REQUESTING DEPT.":      "",
            "PO DATE":               header.get("PO_DATE", ""),
            "PO NUMBER":             header.get("PO_NUMBER", ""),
            "END USER/S":            "",
            "SUPPLIER'S NAME":       header.get("SUPPLIER", ""),
            "ITEM CODE":             it.get("code", ""),
            "ITEM DESCRIPTION":      it.get("desc", ""),
            "SPECIFICATIONS":        header.get("COST_CENTER", "") if i == 0 else "",
            "QTY":                   it.get("qty", ""),
            "UoM":                   it.get("unit", ""),
            "UNIT PRICE":            it.get("unit_price", ""),
            "AMOUNT":                "",
            "TOTAL AMOUNT":          "",
            "PAYMENT TERMS":         header.get("TERMS", "")      if i == 0 else "",
            "PR REQUIRED DATE":      header.get("DELIVER_BY", "") if i == 0 else "",
            "DATE DELIVERED":        "",
            "REMARKS":               "",
            "PURCHASE ORDER STATUS": "",
            "ITEMS/SERVICES":        "",
            "ORIGINAL PRICE":        "",
            "TOTAL COST SAVINGS":    "",
        }
        rows.append(row)
    return rows

# ── Main ──────────────────────────────────────────────────────────────────────

def process_pdf(pdf_path_str):
    import os
    DEBUG = os.environ.get("SCANNER_DEBUG", "")
    pdf_path    = Path(pdf_path_str)
    text, method = extract_pdf_text(pdf_path)
    if DEBUG:
        import sys
        print("=== RAW TEXT START ===", file=sys.stderr)
        print(text, file=sys.stderr)
        print("=== RAW TEXT END ===", file=sys.stderr)
    header      = extract_header(text)
    items       = parse_items(text)

    # Some report-engine PDFs store each table column as its own text block,
    # which scrambles the normal block-ordered extraction into column-major
    # order and drops the numbered item rows entirely (see _extract_words_as_lines
    # above). If that looks like what happened here, rebuild the text from word
    # positions instead and re-parse — only swapping in the retry if it's better.
    if method == "native" and _looks_incomplete(text, header, items):
        try:
            doc = fitz.open(pdf_path)
            reflow_text = "\n".join(_extract_words_as_lines(p) for p in doc)
            doc.close()
            reflow_items = parse_items(reflow_text)
            if _is_better(reflow_items, items, header):
                if DEBUG:
                    import sys
                    print("=== REFLOW TEXT (used) ===", file=sys.stderr)
                    print(reflow_text, file=sys.stderr)
                text, items, method = reflow_text, reflow_items, "native+reflow"
        except Exception as e:
            if DEBUG:
                import sys
                print(f"=== REFLOW FAILED: {e} ===", file=sys.stderr)

    rows        = build_rows(header, items)

    # Separate regular items vs charges for reporting
    regular_items = [it for it in items if not it.get("is_charge")]
    charges       = [it for it in items if it.get("is_charge")]

    return {
        "filename":    pdf_path.name,
        "method":      method,
        "po_number":   header.get("PO_NUMBER", ""),
        "supplier":    header.get("SUPPLIER", ""),
        "total_amount": header.get("TOTAL_AMOUNT", ""),
        "items_count": len(regular_items),
        "charges_count": len(charges),
        "rows":        rows,
    }

def main():
    raw       = sys.stdin.read().strip()
    pdf_paths = json.loads(raw)
    results   = []
    errors    = []
    for path in pdf_paths:
        try:
            results.append(process_pdf(path))
        except Exception as e:
            errors.append({"file": path, "error": str(e)})
    print(json.dumps({"success": True, "results": results, "errors": errors}))

if __name__ == "__main__":
    main()