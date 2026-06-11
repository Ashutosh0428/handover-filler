#!/usr/bin/env python3
"""
extract_pdf.py — dump all text and tables from the handover PDF.

Use this to (a) eyeball what data the PDF actually contains and (b) build the
JSON extraction prompt. Writes pdf_dump.txt next to the script.

Usage:
    python3 extract_pdf.py "Infrastructure Handover (2).pdf"
"""
import sys
import pdfplumber


def main(path: str) -> None:
    out_lines: list[str] = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            out_lines.append(f"\n{'=' * 70}\nPAGE {i}/{len(pdf.pages)}\n{'=' * 70}")
            text = page.extract_text() or ""
            out_lines.append(text)

            for t_idx, table in enumerate(page.extract_tables(), start=1):
                out_lines.append(f"\n--- TABLE {t_idx} on page {i} ---")
                for r in table:
                    cells = [("" if c is None else str(c)) for c in r]
                    out_lines.append(" | ".join(cells))

    dump = "\n".join(out_lines)
    with open("pdf_dump.txt", "w", encoding="utf-8") as fh:
        fh.write(dump)
    print(dump)
    print(f"\n[written to pdf_dump.txt — {len(dump)} chars]")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: python3 extract_pdf.py <handover.pdf>")
    main(sys.argv[1])
