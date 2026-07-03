"""
clean_monitoring_excel.py

Operates directly on an uploaded monitoring.xlsx workbook (not the JSON cache).

Modes:
  inspect <path>            -> prints JSON stats about merged cells / phantom rows
  clean <path> <out_path>   -> unmerges all merged cells (filling values into every
                                cell in the range), trims phantom empty rows past the
                                real data, saves to out_path, prints JSON summary

Usage:
    python clean_monitoring_excel.py inspect data/uploads/pending_123.xlsx
    python clean_monitoring_excel.py clean data/uploads/pending_123.xlsx data/monitoring.xlsx
"""

import sys
import json
import argparse
import openpyxl


def scan_real_last_row(ws, scan_limit=2000):
    """Find the last row that actually has content, scanning up to scan_limit rows."""
    real_last = 0
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=scan_limit, values_only=True), start=1):
        if any(c is not None and str(c).strip() != "" for c in row):
            real_last = i
    return real_last


def inspect_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    sheets = []
    total_merged = 0
    any_phantom = False

    # IMPORTANT: capture ws.max_row for every sheet FIRST, before calling
    # scan_real_last_row() on any sheet. openpyxl's iter_rows(max_row=N)
    # allocates cell objects up to N as a side effect, which inflates
    # ws.max_row afterwards if it hadn't been read yet. Reading it upfront
    # avoids that trap.
    reported_max_rows = {name: wb[name].max_row for name in wb.sheetnames}

    for name in wb.sheetnames:
        ws = wb[name]
        merged_count = len(ws.merged_cells.ranges)
        total_merged += merged_count

        reported = reported_max_rows[name]
        real_last = scan_real_last_row(ws)
        phantom = reported > max(real_last * 3, real_last + 50) and reported > 500

        if phantom:
            any_phantom = True

        sheets.append({
            "name": name,
            "reportedMaxRow": reported,
            "realLastRow": real_last,
            "mergedCells": merged_count,
            "phantomRows": phantom,
        })

    dirty = total_merged > 0 or any_phantom

    return {
        "dirty": dirty,
        "totalMergedCells": total_merged,
        "sheetCount": len(wb.sheetnames),
        "sheets": sheets,
    }


def clean_workbook(path, out_path, max_data_rows=200):
    wb = openpyxl.load_workbook(path)

    per_sheet_summary = []

    for name in wb.sheetnames:
        ws = wb[name]

        # 1. Unmerge every merged range, filling the top-left value into all
        #    cells that were part of the merge (instead of leaving them blank).
        merged_ranges = list(ws.merged_cells.ranges)
        for merged_range in merged_ranges:
            min_col, min_row = merged_range.min_col, merged_range.min_row
            max_col, max_row = merged_range.max_col, merged_range.max_row
            top_left_value = ws.cell(row=min_row, column=min_col).value
            ws.unmerge_cells(str(merged_range))
            for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                for cell in row:
                    cell.value = top_left_value

        # 2. Trim phantom empty rows (Excel formatting applied to entire
        #    columns makes max_row report ~1,048,576 even with a few hundred
        #    real rows of data). Capture max_row BEFORE scanning - the scan
        #    itself can inflate max_row as an openpyxl side effect.
        original_max_row = ws.max_row
        real_last = scan_real_last_row(ws)
        trimmed = 0
        if ws.max_row > real_last:
            trimmed = original_max_row - real_last
            ws.delete_rows(real_last + 1, ws.max_row - real_last)

        per_sheet_summary.append({
            "name": name,
            "mergedCellsRemoved": len(merged_ranges),
            "phantomRowsTrimmed": trimmed,
            "finalRowCount": ws.max_row,
        })

    wb.save(out_path)

    return {
        "ok": True,
        "outPath": out_path,
        "sheets": per_sheet_summary,
    }


def main():
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="mode", required=True)

    p_inspect = sub.add_parser("inspect")
    p_inspect.add_argument("path")

    p_clean = sub.add_parser("clean")
    p_clean.add_argument("path")
    p_clean.add_argument("out_path")

    args = parser.parse_args()

    try:
        if args.mode == "inspect":
            result = inspect_workbook(args.path)
        else:
            result = clean_workbook(args.path, args.out_path)
    except Exception as e:
        print(json.dumps({"ok": False, "error": str(e)}))
        sys.exit(1)

    print(json.dumps(result))
    sys.exit(0)


if __name__ == "__main__":
    main()