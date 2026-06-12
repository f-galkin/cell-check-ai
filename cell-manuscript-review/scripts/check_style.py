#!/usr/bin/env python3
"""Run-level style checks the text-based analyzer cannot perform.

Reads the .docx directly with python-docx and reports:
  - precise main-text word count (Introduction -> end of Discussion, excluding
    embedded figure legends and tables);
  - italic state of organism binomials (auto) and any gene symbols passed via
    --genes, so the reviewer can confirm Cell's convention (gene symbols
    italic, protein products and organism binomials italic).

Use this in Step 2 of the review, passing the gene symbols you noticed while
reading the manuscript:

  python check_style.py MANUSCRIPT.docx --genes KDM1A,SOCS3,CXCL12,SIRT1

Requires python-docx (`pip install python-docx`).
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import docx
from docx.oxml.text.paragraph import CT_P
from docx.text.paragraph import Paragraph

DEFAULT_ORGANISMS = [
    "Drosophila melanogaster", "Caenorhabditis elegans", "Mus musculus",
    "Homo sapiens", "Saccharomyces cerevisiae", "Escherichia coli",
    "D. melanogaster", "C. elegans", "M. musculus", "S. cerevisiae",
]


def _is_heading(style: str | None) -> bool:
    s = (style or "").lower()
    return s.startswith("heading") or s == "title"


def _wc(s: str) -> int:
    return len([w for w in re.sub(r"\s+", " ", s).split() if any(ch.isalnum() for ch in w)])


def main_text_words(d) -> int | None:
    in_main = in_legend = found = False
    total = 0
    for child in d.element.body.iterchildren():
        if not isinstance(child, CT_P):
            continue
        p = Paragraph(child, d)
        style = p.style.name if p.style else "Normal"
        txt = p.text.strip()
        if _is_heading(style):
            up = txt.upper()
            if up == "INTRODUCTION":
                in_main, in_legend, found = True, False, True
                continue
            if up == "RESOURCE AVAILABILITY":
                in_main = False
        if not in_main or not txt:
            continue
        if re.match(r"^(Figure|Table)\s", txt, re.I):
            in_legend = True
            continue
        if _is_heading(style):
            in_legend = False
            total += _wc(txt)
            continue
        if in_legend:
            continue
        total += _wc(txt)
    return total if found else None


def italic_state(d, term: str) -> list[str]:
    """Return one verdict per occurrence: ITAL / PARTIAL / ROMAN."""
    out: list[str] = []
    for child in d.element.body.iterchildren():
        if not isinstance(child, CT_P):
            continue
        p = Paragraph(child, d)
        if term.lower() not in p.text.lower():
            continue
        pos = 0
        spans = []
        for r in p.runs:
            spans.append((pos, pos + len(r.text), bool(r.italic)))
            pos += len(r.text)
        for m in re.finditer(re.escape(term), p.text, re.I):
            s, e = m.start(), m.end()
            flags = [it for (a, b, it) in spans if a < e and b > s]
            if flags and all(flags):
                out.append("ITAL")
            elif any(flags):
                out.append("PARTIAL")
            else:
                out.append("ROMAN")
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("manuscript", type=Path)
    ap.add_argument("--genes", default="",
                    help="Comma-separated gene symbols to check for italicisation")
    ap.add_argument("--organisms", default="",
                    help="Extra comma-separated organism names to check (binomials)")
    args = ap.parse_args()

    try:
        sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
    except Exception:
        pass

    d = docx.Document(str(args.manuscript))

    mtw = main_text_words(d)
    print("=== WORD COUNT ===")
    if mtw is None:
        print("Introduction heading not found; cannot compute precise main-text count.")
    else:
        flag = "  >>> OVER 7,000" if mtw > 7000 else "  (within 7,000 limit)"
        print(f"Main text (Introduction -> Discussion, excl. figure legends & tables): {mtw}{flag}")

    organisms = DEFAULT_ORGANISMS + [o.strip() for o in args.organisms.split(",") if o.strip()]
    genes = [g.strip() for g in args.genes.split(",") if g.strip()]

    print("\n=== ITALICS AUDIT ===")
    print("(organism binomials and gene symbols should be italic in Cell)")
    any_org = False
    for term in organisms:
        h = italic_state(d, term)
        if h:
            any_org = True
            verdict = "OK" if all(x == "ITAL" for x in h) else "CHECK"
            print(f"[organism] {term:28s} {h}  -> {verdict}")
    if not any_org:
        print("[organism] none of the checked binomials were found")
    for term in genes:
        h = italic_state(d, term)
        if not h:
            print(f"[gene]     {term:12s} not found")
            continue
        verdict = "OK" if all(x == "ITAL" for x in h) else "CHECK (should be italic)"
        print(f"[gene]     {term:12s} {h}  -> {verdict}")


if __name__ == "__main__":
    main()
