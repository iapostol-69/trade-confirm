"""
build_reconciliation.py
-----------------------
Helper script for the trade-confirm skill.
Builds a formatted Excel reconciliation file from matched PO vs OC data.

Usage:
    python build_reconciliation.py <data.json> <output.xlsx>

Input JSON format:
{
  "order_ref": "2025-1933",
  "mismatches": [
    {
      "order_code": "ART.001",
      "order_desc": "Stainless pipe 50mm",
      "order_qty": 10,
      "order_price": 25.50,
      "conf_code": "ART.001",
      "conf_desc": "SS Pipe 50mm",
      "conf_qty": 8,
      "conf_price": 25.50,
      "match_type": "Code"   // "Code", "Description", or "Unmatched"
    }
  ],
  "matches": [
    {
      "order_code": "ART.002",
      ...same fields...
    }
  ],
  "extra_in_confirmation": [
    {
      "conf_code": "ART.999",
      "conf_desc": "Extra item not in order",
      "conf_qty": 5,
      "conf_price": 12.00
    }
  ]
}
"""

import sys
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------- Colors ----------
RED_HEADER_FILL   = PatternFill("solid", fgColor="CC3333")
GREEN_HEADER_FILL = PatternFill("solid", fgColor="2E7D32")
ROW_MISMATCH_FILL = PatternFill("solid", fgColor="FFCCCC")
ROW_MATCH_FILL    = PatternFill("solid", fgColor="CCFFCC")
CELL_DIFF_FILL    = PatternFill("solid", fgColor="FF6666")  # specific differing cell
UNMATCHED_FILL    = PatternFill("solid", fgColor="FFF0CC")
SUMMARY_FILL      = PatternFill("solid", fgColor="EEF2FF")
EXTRA_FILL        = PatternFill("solid", fgColor="FFE4B5")

WHITE_BOLD = Font(name="Arial", size=10, bold=True, color="FFFFFF")
NORMAL     = Font(name="Arial", size=10)
BOLD       = Font(name="Arial", size=10, bold=True)
SUMMARY_FONT = Font(name="Arial", size=10, bold=True, color="1A237E")

HEADERS = [
    "Code (Order)", "Description (Order)", "Qty (Order)", "Price (Order)",
    "Code (Confirm.)", "Description (Confirm.)", "Qty (Confirm.)", "Price (Confirm.)",
    "Match Type"
]

THIN = Side(style="thin", color="CCCCCC")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _write_header(ws, row, fill):
    for col, h in enumerate(HEADERS, start=1):
        cell = ws.cell(row=row, column=col, value=h)
        cell.font = WHITE_BOLD
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER


def _write_data_row(ws, row, item, row_fill, highlight_qty=False, highlight_price=False, extra=False):
    values = [
        item.get("order_code", ""),
        item.get("order_desc", ""),
        item.get("order_qty", ""),
        item.get("order_price", ""),
        item.get("conf_code", ""),
        item.get("conf_desc", ""),
        item.get("conf_qty", ""),
        item.get("conf_price", ""),
        item.get("match_type", ""),
    ]
    for col, val in enumerate(values, start=1):
        cell = ws.cell(row=row, column=col, value=val)
        cell.font = NORMAL
        cell.fill = row_fill
        cell.border = THIN_BORDER
        cell.alignment = Alignment(vertical="center")
        # Highlight the specific differing cells
        if highlight_qty and col in (3, 7):   # Qty columns
            cell.fill = CELL_DIFF_FILL
            cell.font = Font(name="Arial", size=10, bold=True, color="880000")
        if highlight_price and col in (4, 8): # Price columns
            cell.fill = CELL_DIFF_FILL
            cell.font = Font(name="Arial", size=10, bold=True, color="880000")

    # Number formatting for qty and price columns
    for col in (3, 7):
        ws.cell(row=row, column=col).number_format = "#,##0.##"
    for col in (4, 8):
        ws.cell(row=row, column=col).number_format = "#,##0.00"


def _auto_fit_columns(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_len = max(max_len, len(str(cell.value)))
            except Exception:
                pass
        adjusted = min(max(max_len + 2, 12), 50)
        ws.column_dimensions[col_letter].width = adjusted


def build(data: dict, output_path: str):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reconciliation"
    ws.sheet_view.showGridLines = True

    mismatches = data.get("mismatches", [])
    matches    = data.get("matches", [])
    extras     = data.get("extra_in_confirmation", [])
    order_ref  = data.get("order_ref", "")

    total_order_items = len(mismatches) + len(matches)
    n_mismatches  = len(mismatches)
    n_matches     = len(matches)
    n_unmatched   = sum(1 for m in mismatches if m.get("match_type") == "Unmatched")
    n_extra       = len(extras)

    # ---- Row 1: Summary ----
    summary_text = (
        f"Order ref: {order_ref}   |   "
        f"Order items: {total_order_items}   |   "
        f"Mismatches: {n_mismatches - n_unmatched}   |   "
        f"Unmatched: {n_unmatched}   |   "
        f"Matches: {n_matches}"
        + (f"   |   Extra in confirmation: {n_extra}" if n_extra else "")
    )
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(HEADERS))
    summary_cell = ws.cell(row=1, column=1, value=summary_text)
    summary_cell.font = SUMMARY_FONT
    summary_cell.fill = SUMMARY_FILL
    summary_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    current_row = 2

    # ---- Section 1: Mismatches ----
    if mismatches:
        _write_header(ws, current_row, RED_HEADER_FILL)
        ws.row_dimensions[current_row].height = 28
        current_row += 1

        for item in mismatches:
            is_unmatched = item.get("match_type") == "Unmatched"
            row_fill = UNMATCHED_FILL if is_unmatched else ROW_MISMATCH_FILL
            # Detect which fields differ
            try:
                qty_diff   = float(item.get("order_qty") or 0) != float(item.get("conf_qty") or 0)
            except (TypeError, ValueError):
                qty_diff = item.get("order_qty") != item.get("conf_qty")
            try:
                price_diff = abs(float(item.get("order_price") or 0) - float(item.get("conf_price") or 0)) >= 0.01
            except (TypeError, ValueError):
                price_diff = item.get("order_price") != item.get("conf_price")

            _write_data_row(ws, current_row, item, row_fill,
                            highlight_qty=qty_diff and not is_unmatched,
                            highlight_price=price_diff and not is_unmatched)
            ws.row_dimensions[current_row].height = 18
            current_row += 1

    # ---- Blank separator ----
    current_row += 1

    # ---- Section 2: Matches ----
    if matches:
        _write_header(ws, current_row, GREEN_HEADER_FILL)
        ws.row_dimensions[current_row].height = 28
        current_row += 1

        for item in matches:
            _write_data_row(ws, current_row, item, ROW_MATCH_FILL)
            ws.row_dimensions[current_row].height = 18
            current_row += 1

    # ---- Section 3: Extra items in confirmation (not in order) ----
    if extras:
        current_row += 1
        extra_header_fill = PatternFill("solid", fgColor="E65100")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=len(HEADERS))
        label = ws.cell(row=current_row, column=1,
                        value="Items in Confirmation NOT found in the Order")
        label.font = WHITE_BOLD
        label.fill = extra_header_fill
        label.alignment = Alignment(horizontal="center", vertical="center")
        current_row += 1

        _write_header(ws, current_row, extra_header_fill)
        current_row += 1
        for item in extras:
            extra_item = {
                "order_code": "",
                "order_desc": "",
                "order_qty": "",
                "order_price": "",
                "conf_code":  item.get("conf_code", ""),
                "conf_desc":  item.get("conf_desc", ""),
                "conf_qty":   item.get("conf_qty", ""),
                "conf_price": item.get("conf_price", ""),
                "match_type": "Extra",
            }
            _write_data_row(ws, current_row, extra_item, EXTRA_FILL)
            current_row += 1

    # ---- Auto-fit and freeze ----
    _auto_fit_columns(ws)
    ws.freeze_panes = "A3"  # Keep summary + first header visible

    wb.save(output_path)
    print(f"Saved: {output_path}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python build_reconciliation.py <data.json> <output.xlsx>")
        sys.exit(1)
    with open(sys.argv[1]) as f:
        data = json.load(f)
    build(data, sys.argv[2])
