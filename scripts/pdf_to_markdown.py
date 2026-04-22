#!/usr/bin/env python3
"""
pdf_to_markdown.py
Convert a PDF to Markdown, rendering bordered tables as proper markdown tables
and all other content as layout-preserving text blocks.

Usage: python3 pdf_to_markdown.py <input.pdf> [output.md]
"""
import sys
import os
import pdfplumber


def table_to_markdown(rows):
    """Render a pdfplumber table (list of lists) as a markdown table."""
    if not rows:
        return ""

    # Normalise cells: strip, replace newlines with space, handle None
    clean = []
    for row in rows:
        if row is None:
            continue
        clean.append([str(cell or "").replace("\n", " ").strip() for cell in row])

    if not clean:
        return ""

    col_count = max(len(r) for r in clean)
    # Pad short rows
    clean = [r + [""] * (col_count - len(r)) for r in clean]

    separator = "| " + " | ".join(["---"] * col_count) + " |"
    lines = ["| " + " | ".join(clean[0]) + " |", separator]
    for row in clean[1:]:
        lines.append("| " + " | ".join(row) + " |")

    return "\n".join(lines)


def page_to_markdown(page):
    """Convert a single pdfplumber page to a markdown string."""
    tables = page.find_tables()

    if not tables:
        text = page.extract_text(layout=True)
        if not text or not text.strip():
            return ""
        return "```\n" + text.rstrip() + "\n```"

    # Build a list of vertical slices: (y_start, y_end, kind, payload)
    slices = []
    prev_y = 0.0
    for t in tables:
        y0, y1 = t.bbox[1], t.bbox[3]
        if y0 > prev_y + 2:
            slices.append(("text", prev_y, y0, None))
        slices.append(("table", y0, y1, t))
        prev_y = y1
    if prev_y < page.height - 2:
        slices.append(("text", prev_y, page.height, None))

    parts = []
    for kind, y0, y1, payload in slices:
        if kind == "table":
            md = table_to_markdown(payload.extract())
            if md:
                parts.append(md)
        else:
            region = page.crop((0, y0, page.width, y1))
            text = region.extract_text(layout=True)
            if text and text.strip():
                parts.append("```\n" + text.strip() + "\n```")

    return "\n\n".join(parts)


def pdf_to_markdown(pdf_path, md_path=None):
    if md_path is None:
        base = os.path.splitext(pdf_path)[0]
        md_path = base + ".md"

    title = os.path.splitext(os.path.basename(pdf_path))[0]
    output = [f"# {title}\n"]

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            content = page_to_markdown(page)
            if content:
                output.append(f"## Page {page_num}\n")
                output.append(content)

    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(output) + "\n")

    print(f"Saved to: {md_path}")
    return md_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"Usage: {sys.argv[0]} <input.pdf> [output.md]")
        sys.exit(1)
    pdf_to_markdown(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
