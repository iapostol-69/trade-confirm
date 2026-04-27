---
name: trade-confirm
description: "Reconcile a purchase order against a supplier order confirmation, matching line items and highlighting discrepancies. Use this skill whenever the user wants to compare an order they sent to a supplier with the confirmation received back — even if they phrase it as 'check the supplier confirmation', 'does the OC match our order', 'reconcile the PO and OC', or 'compare the order and the confirmation'. The skill accepts one order file and one or more confirmation files (PDF or Excel) as input and produces a formatted Excel file with two sections: mismatches first (highlighted), then matching items. Trigger for any procurement reconciliation, PO vs OC comparison, supplier order check, or trade confirmation task."
version: 0.3
---

# Trade Confirm – Purchase Order vs. Order Confirmation Reconciliation

You are acting as a procurement expert. Your job is to compare a purchase order (PO) we sent to a supplier with the order confirmation (OC) they sent back, identify discrepancies, and produce a structured Excel reconciliation file.

## Step 1 – Identify the input files

The user will provide one order file and one or more confirmation files (suppliers sometimes split their confirmation across multiple files). Identify which file belongs to the order and which file(s) belong to the confirmation. Common naming patterns:
- Order: "OUR ORDER", "PO", "ITM", "ORDER" in the filename
- Confirmation: "OC", "CONFIRMATION", "CONF" in the filename

If you cannot tell from the filename, ask the user.

## Step 2 – Extract line items from each file

Use pdfplumber to extract text from PDFs. Install it if needed: `pip install pdfplumber --break-system-packages -q`

For the order file, extract the following pre line item details:
- **Our Code**: Product/item code used internally by us, typically a string with 6 digits (e.g. "071234") - usually the column name would be "Our Code" or simply "Code" - always included in the order file
- **Supplier Code**: Product/item code used by the supplier, often alphanumeric (e.g. "ART.001", "SKU-1234") - usually the column name would be "Supplier Code", "Part Number" - very often this collumn is missing from the order file or empty
- **Description**: Product/item description - this column is always included in the order file
- **Quantity**: Numeric quantity ordered - this column is always included in the order file
- **Total Price**: The total price for the line item (not unit price) — if only unit price is available, compute total = unit price * quantity

If there are multiple confirmation files, extract line items from each one and merge them into a single list before proceeding — treat the merged list as one combined confirmation.

For each file, extract the following per line item details:
- **Our Code**: the Product/item code used internally by us, typically a string with 6 digits (e.g. "071234") - may be missing from the confirmation file 
- **Supplier Code**: the Product/item code used by the supplier, often alphanumeric (e.g. "ART.001", "SKU-1234") - may be missing from the confirmation file
- **Description**: Product name and/or specifications - always provided in the confirmation file
- **Quantity**: Numeric quantity ordered - always provided in the confirmation file
- **Total Price**: Total price for the line item - always provided in the confirmation file

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
If both the Order and Confirmation have our product code for an item, match on exact code. Alternatively if both have the supplier code, match on exact supplier code. If one document has both codes and the other document has only one of them, you can still match if the available code matches (e.g. if the Order has both "Our Code" and "Supplier Code" but the Confirmation only has "Supplier Code", you can match based on the supplier code). This is the most reliable method when codes are present. If you find a match by code, mark the match confidence as "code" — this will help the user spot reliable matches.


**Rule 2 – Match by Description (fallback)**
If the Confirmation has no code, or the code does not match anything in the Order, attempt to match by description. Use your expertise to identify the same product even when the description is phrased differently (e.g., "Stainless steel pipe 50mm" vs "PIPE SS 50MM"). Consider:
- Key dimensions and specifications (diameter, length, material, grade)
- Product category and type
- Manufacturer codes embedded in the description
Mark the match confidence as "description" — this will help the user spot uncertain matches.

**Unmatched items**
- Items in the Order with no corresponding item in the Confirmation → report as unmatched (Confirmation columns empty)
- Items in the Confirmation with no corresponding item in the Order → report at the bottom as extra items

## Step 4 – Identify discrepancies

For each matched pair, compare:
- **Quantity**: Flag if Order qty ≠ Confirmation qty
- **Price**: Flag if Order price ≠ Confirmation price (treat values as floats; ignore minor rounding differences < 0.01)
- **Description** (code-matched items only): When two items were matched by product code, also compare their descriptions using the same semantic/fuzzy methodology described in Step 3 Rule 2. If you determine the descriptions refer to different products or meaningfully different specifications (e.g. different material, grade, size, or product type), flag as a discrepancy and set `desc_mismatch: true` on the item. This catches cases where a supplier reuses a code for a different product. If the descriptions are just stylistic variants of the same item (abbreviations, different word order, language differences), do **not** flag — only flag when a procurement officer would genuinely need to investigate further.

A row is a **mismatch** if quantity OR price differs, OR if the item is unmatched, OR if items were matched by code but their descriptions don't match.
A row is a **match** if both quantity and price agree **and** descriptions are consistent (or were matched by description).

## Step 5 – Create the Excel output

**IMPORTANT: You MUST generate the Excel file by calling `scripts/build_reconciliation.py`. Do NOT write custom openpyxl code — all formatting, column layout, and cell-level highlighting logic lives in that script. Bypassing it will produce an output with no cell-level highlighting.**

Install openpyxl if needed: `pip install openpyxl --break-system-packages -q`

### How to call the script

1. Build your reconciliation data as a Python dict matching the JSON schema below.
2. Write it to a temporary JSON file (e.g. `/tmp/rec_data.json`).
3. Call the script: `python scripts/build_reconciliation.py /tmp/rec_data.json <output.xlsx>`

### JSON schema

```json
{
  "order_ref": "2025-1933",
  "mismatches": [
    {
      "order_our_code":       "071203", // our internal product code from the order file; if missing, set to null
      "order_supplier_code":  "CUST-001", //supplier product code from the order file; if missing, set to null
      "order_desc":           "Stainless pipe 50mm", // the description of the item from the order file
      "order_qty":            10, //quantity from the order file
      "order_price":          25.50, // total price/amount for the item from the order file
      "conf_our_code":        "071203", // our internal product code from the confirmation file; if missing, set to null 
      "conf_supplier_code":   "CUST-001", // supplier product code from the confirmation file; if missing, set to null
      "conf_desc":            "SS Pipe 50mm", // the description of the item from the confirmation file
      "conf_qty":             8, // quantity from the confirmation file
      "conf_price":           27.00, // total price/amount for the item from the confirmation file
      "match_type":           "Our Code", // "Our Code", "Supplier Code", "Description", or "Unmatched"
      "desc_mismatch":        false // true only when matched by code but descriptions suggest a different product (see Step 4)
    }
  ],
  "matches": [
    { "...same fields..." }
  ],
  "extra_in_confirmation": [
    {
      "conf_our_code":  "123544",
      "conf_supplier_code": "ART.999",
      "conf_desc":  "Stainless pipe 50mm",
      "conf_qty":   5,
      "conf_price": 12.00
    }
  ]
}
```

**Field mapping notes:**
- `order_our_code` / `conf_our_code`: map our internal code in the order to the internal code in the confirmation. This is typically the code that best identifies the item across both documents (typically the supplier/manufacturer code). 
- `order_supplier_code`/`conf_supplier_code`: map the supplier's code in the order to the supplier's code in the confirmation. If the order file doesn't have a supplier code column, set `order_supplier_code` to null for all items. If the confirmation file doesn't have a supplier code column, set `conf_supplier_code` to null for all items. 
- `order_price` / `conf_price`: unit price. If the document only shows a line total, divide by qty. If qty is zero or ambiguous, store the total as-is and note it in the description.
- `desc_mismatch`: set `true` only when the AI judges the descriptions refer to a genuinely different product (see Step 4). Do not set for stylistic variants.

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
