#!/usr/bin/env python3
"""
pdf_tables_to_excel.py
Extract tables from a PDF and write them to a single Excel sheet.

Uses layout-preserving text extraction: columns are identified by two or more
consecutive spaces in the fixed-width layout, which works for both bordered
and borderless tables.

Usage: python3 pdf_tables_to_excel.py <input.pdf> [output.xlsx]
"""
import sys
import os
import re
import pdfplumber
import openpyxl


def extract_tables_to_excel(pdf_path, excel_path=None):
    if excel_path is None:
        base = os.path.splitext(pdf_path)[0]
        excel_path = base + "_tables.xlsx"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Tables"
    current_row = 1

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(layout=True)
            if not text:
                continue

            lines = text.split("\n")
            page_has_content = False

            for line in lines:
                if not line.strip():
                    continue
                # Split on 2+ spaces to separate columns
                cols = re.split(r" {2,}", line.strip())
                for col_idx, cell in enumerate(cols, start=1):
                    ws.cell(row=current_row, column=col_idx, value=cell if cell else None)
                current_row += 1
                page_has_content = True

            if page_has_content:
                current_row += 1  # blank row between pages

    wb.save(excel_path)
    print(f"Saved {current_row - 1} rows to: {excel_path}")
    return excel_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <input.pdf> [output.xlsx]")
        sys.exit(1)
    pdf_path = sys.argv[1]
    excel_path = sys.argv[2] if len(sys.argv) > 2 else None
    extract_tables_to_excel(pdf_path, excel_path)
