---
name: trade-confirm
description: "Reconcile a purchase order against a supplier order confirmation, matching line items and highlighting discrepancies. Use this skill whenever the user wants to compare an order they sent to a supplier with the confirmation received back — even if they phrase it as 'check the supplier confirmation', 'does the OC match our order', 'reconcile the PO and OC', or 'compare the order and the confirmation'. The skill accepts two PDF (or Excel) files as input and produces a formatted Excel file with two sections: mismatches first (highlighted), then matching items. Trigger for any procurement reconciliation, PO vs OC comparison, supplier order check, or trade confirmation task."
---

# Trade Confirm – Purchase Order vs. Order Confirmation Reconciliation

You are acting as a procurement expert. Your job is to compare a purchase order (PO) we sent to a supplier with the order confirmation (OC) they sent back, identify discrepancies, and produce a structured Excel reconciliation file.

## Step 1 – Identify the input files

The user will provide two files. Identify which is the order and which is the confirmation. Common naming patterns:
- Order: "OUR ORDER", "PO", "ITM", "ORDER" in the filename
- Confirmation: "OC", "CONFIRMATION", "CONF" in the filename

If you cannot tell from the filename, ask the user.

## Step 2 – Extract line items from each file

Use pdfplumber to extract text from PDFs. Install it if needed: `pip install pdfplumber --break-system-packages -q`

For each file, extract the following per line item:
- **Code**: Product/item code or reference number (e.g., "SKU-1234", "ART.001")
- **Description**: Product name and/or specifications
- **Quantity**: Numeric quantity ordered
- **Unit Price**: Unit price per item

**Parsing tips:**
- Tables in PDFs can be read with `page.extract_tables()` — prefer this over raw text when possible
- If tables are not well-structured, fall back to `page.extract_text()` and parse line-by-line
- Look for common column headers like "Qty", "Quantity", "Pcs", "Price", "Unit Price", "€", "Code", "Ref", "Art."
- Prices may include currency symbols (€, $) — strip these when storing numeric values
- Quantities may include units ("pcs", "kg", "m") — store them separately or include in description
- Line items may span multiple pages — collect all pages

Write a small Python script to do the extraction and print the results as JSON to stdout, so you can review them before proceeding.

## Step 3 – Match line items

Match each item from the Order to the corresponding item in the Confirmation:

**Rule 1 – Match by Code (preferred)**
If both the Order and Confirmation have a product code for an item, match on exact code (case-insensitive, ignoring spaces and dashes).

**Rule 2 – Match by Description (fallback)**
If the Confirmation has no code, or the code does not match anything in the Order, attempt to match by description. Use your expertise to identify the same product even when the description is phrased differently (e.g., "Stainless steel pipe 50mm" vs "PIPE SS 50MM"). Consider:
- Key dimensions and specifications (diameter, length, material, grade)
- Product category and type
- Manufacturer codes embedded in the description

Mark the match confidence as "code" or "description" — this will help the user spot uncertain matches.

**Unmatched items**
- Items in the Order with no corresponding item in the Confirmation → report as unmatched (Confirmation columns empty)
- Items in the Confirmation with no corresponding item in the Order → report at the bottom as extra items

## Step 4 – Identify discrepancies

For each matched pair, compare:
- **Quantity**: Flag if Order qty ≠ Confirmation qty
- **Price**: Flag if Order price ≠ Confirmation price (treat values as floats; ignore minor rounding differences < 0.01)

A row is a **mismatch** if quantity OR price differs, OR if the item is unmatched.
A row is a **match** if both quantity and price agree.

## Step 5 – Create the Excel output

Use `openpyxl` to build the Excel file. Install if needed: `pip install openpyxl --break-system-packages -q`

Use the bundled helper script at `scripts/build_reconciliation.py` to generate the Excel file. Read it first and adapt the call as needed for your data.

### Excel structure

**Sheet name:** "Reconciliation"

**Section 1 – Mismatches** (rows that differ or are unmatched)
Header row (bold, dark background, white text):

| Code (Order) | Description (Order) | Qty (Order) | Price (Order) | Code (Confirm.) | Description (Confirm.) | Qty (Confirm.) | Price (Confirm.) | Match Type |

**Section 2 – Matches** (rows where qty and price agree)
Same columns as Section 1. A blank separator row between the two sections.

### Formatting rules
- **Section 1 header**: Fill `FF4444` (red), white bold text
- **Section 2 header**: Fill `44AA44` (green), white bold text
- **Mismatch rows**: Light red background `FFE0E0` on the differing cells (Qty or Price columns only)
- **Match rows**: Light green background `E0FFE0`
- **Unmatched items** (no confirmation): Orange background `FFF0CC` on entire row
- **Match Type column**: Show "Code" (exact code match), "Description" (fuzzy description match), or "Unmatched"
- Column widths: auto-fit to content (min 12, max 50)
- Freeze the top rows so headers stay visible when scrolling
- Use Arial 10pt font throughout
- Currency values: format as `#,##0.00`
- Add a **Summary row** at the very top (above Section 1 header): "Total items: X | Mismatches: Y | Matches: Z | Unmatched: W"

### Output filename
Save as: `Trade_Confirm_<order_ref>.xlsx` where `<order_ref>` is extracted from the order filename (e.g., "2025-1933" from "2025-1933 OUR ORDER SIDERINOX.pdf").

Save to the same folder as the input files.

## Step 6 – Report back

After generating the file, briefly tell the user:
- How many items were in the order
- How many mismatches were found (and what kind: qty-only, price-only, both)
- How many items matched perfectly
- Any items that appeared in the confirmation but not in the order

Keep the summary short — the Excel file is the main deliverable.
