#!/usr/bin/env python3
"""Convert a Word .docx to Markdown using python-docx.

Maps Word paragraph styles to Markdown structure:
  - "Heading N" / "Title"  -> '#'*N headings (Title -> level 1)
  - "List Paragraph"       -> '- ' bullets
  - tables                 -> GitHub pipe tables
  - everything else        -> plain paragraphs

Two output modes:
  * plain (default): paragraph text only, hyperlink display text preserved,
    no inline emphasis markers. This is the deterministic shape consumed by
    downstream tooling (e.g. the cell-manuscript-review analyzer), where
    stray '*' or duplicated link URLs would corrupt regex checks.
  * rich (--emphasis / --links): wrap bold/italic runs in '**'/'*' and render
    hyperlinks as [text](url). Nicer for human-facing conversion.

Requires python-docx >= 1.1 (for Paragraph.iter_inner_content and hyperlink
text in Paragraph.text).

Usage:
  python docx_to_md.py INPUT.docx [OUTPUT.md] [--emphasis] [--links]

If OUTPUT.md is omitted, Markdown is written to stdout.
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

import docx
from docx.document import Document as _Doc
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import Table, _Cell
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph
from docx.text.run import Run


def _iter_block_items(parent):
    """Yield Paragraph and Table objects in document order."""
    if isinstance(parent, _Doc):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        parent_elm = parent._element
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def _heading_level(style_name: str) -> int | None:
    m = re.match(r"heading\s+(\d+)", style_name.lower())
    if m:
        return int(m.group(1))
    if style_name.lower() == "title":
        return 1
    return None


def _wrap_emphasis(text: str, bold: bool, italic: bool) -> str:
    """Wrap text in Markdown emphasis, keeping surrounding whitespace outside
    the markers (Markdown does not render '** bold **')."""
    if not text.strip() or not (bold or italic):
        return text
    marker = ("**" if bold else "") + ("*" if italic else "")
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    core = text.strip()
    return f"{lead}{marker}{core}{marker}{trail}"


def _render_runs_rich(para: Paragraph, emphasis: bool, links: bool) -> str:
    """Render a paragraph's inner content (runs + hyperlinks) in order,
    merging consecutive runs that share emphasis so we don't emit '**a****b**'."""
    out: list[str] = []
    buf_text, buf_bold, buf_italic = "", None, None

    def flush():
        nonlocal buf_text, buf_bold, buf_italic
        if buf_text:
            out.append(_wrap_emphasis(buf_text, bool(buf_bold), bool(buf_italic))
                       if emphasis else buf_text)
        buf_text, buf_bold, buf_italic = "", None, None

    for item in para.iter_inner_content():
        if isinstance(item, Run):
            b, i = bool(item.bold), bool(item.italic)
            if buf_text and (b != buf_bold or i != buf_italic):
                flush()
            buf_bold, buf_italic = b, i
            buf_text += item.text
        elif isinstance(item, Hyperlink):
            flush()
            txt = item.text
            if links and item.address:
                out.append(f"[{txt}]({item.address})")
            else:
                out.append(txt)
    flush()
    return "".join(out)


def docx_to_md(path: str, *, inline_emphasis: bool = False, hyperlinks: bool = False) -> str:
    """Convert a .docx file to Markdown text.

    inline_emphasis: wrap bold/italic runs in '**'/'*'.
    hyperlinks: render hyperlinks as [display](url) instead of display text only.
    """
    d = docx.Document(path)
    rich = inline_emphasis or hyperlinks
    lines: list[str] = []
    for block in _iter_block_items(d):
        if isinstance(block, Paragraph):
            style = (block.style.name if block.style else "Normal") or "Normal"
            level = _heading_level(style)
            text = (_render_runs_rich(block, inline_emphasis, hyperlinks)
                    if rich else block.text).rstrip()
            if level is not None and text.strip():
                lines.append("#" * level + " " + text.strip())
            elif style == "List Paragraph":
                lines.append("- " + text.strip())
            else:
                lines.append(text)
        else:  # Table -> pipe rows
            for row in block.rows:
                cells = [c.text.strip().replace("\n", " ") for c in row.cells]
                lines.append("| " + " | ".join(cells) + " |")
            lines.append("")
    md = "\n".join(lines)
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("input", type=Path, help="Input .docx path")
    ap.add_argument("output", type=Path, nargs="?", default=None,
                    help="Output .md path (default: stdout)")
    ap.add_argument("--emphasis", action="store_true",
                    help="Wrap bold/italic runs in Markdown emphasis markers")
    ap.add_argument("--links", action="store_true",
                    help="Render hyperlinks as [text](url)")
    args = ap.parse_args()

    if not args.input.exists():
        print(f"Error: file not found: {args.input}", file=sys.stderr)
        sys.exit(1)

    md = docx_to_md(str(args.input), inline_emphasis=args.emphasis, hyperlinks=args.links)
    if args.output:
        args.output.write_text(md, encoding="utf-8")
        print(f"Wrote {args.output} ({len(md)} chars)", file=sys.stderr)
    else:
        try:
            sys.stdout.reconfigure(encoding="utf-8")  # type: ignore[attr-defined]
        except Exception:
            pass
        print(md)


if __name__ == "__main__":
    main()
