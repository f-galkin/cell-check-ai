---
name: docx-to-markdown
description: "Convert a Microsoft Word .docx file to clean Markdown using python-docx, with no Node.js or external CLI dependency. Maps Word heading styles to Markdown headings, list paragraphs to bullets, and tables to GitHub pipe tables; preserves hyperlink text. Use whenever the user wants to turn a .docx into Markdown, extract the text/structure of a Word document, or needs a portable cross-platform docx text extractor. Also serves as the extraction fallback for the cell-manuscript-review skill when the Anthropic `extract-text` CLI is unavailable (e.g. on Windows). Trigger on phrases like 'convert this Word doc to markdown', 'extract the text from this docx', 'turn this .docx into md', or 'get the structure of this document'."
---

# docx to Markdown

Convert a Word `.docx` to Markdown with `python-docx` only — no Node.js, no Pandoc, no Linux-only CLI. Runs anywhere Python and `python-docx` are installed.

## When to use

- The user asks to convert a `.docx` to Markdown, or to extract the text/structure of a Word document.
- You need a portable text extractor for a `.docx` and the Anthropic `extract-text` CLI is not on the PATH (common on Windows).
- Another skill (e.g. `cell-manuscript-review`) needs to read a `.docx` as Markdown and wants a dependency-light fallback.

## What it produces

The converter walks the document body in order and emits:

- **Headings** — a paragraph styled `Heading N` (or `Title`) becomes a level-N Markdown heading (`# `, `## `, …). This is what lets downstream tooling find sections.
- **Bullets** — paragraphs in the `List Paragraph` style become `- ` items.
- **Tables** — each table becomes GitHub pipe rows (`| a | b |`).
- **Paragraphs** — everything else is emitted as plain text.

Two modes:

- **Plain (default)** — paragraph text only; hyperlink *display text* is preserved but link targets are not duplicated, and no `*` emphasis markers are added. This is the deterministic shape that regex-based consumers expect (stray `*` or a doubled URL would corrupt their checks). The `cell-manuscript-review` analyzer relies on this mode.
- **Rich (`--emphasis` / `--links`)** — wraps bold/italic runs in `**`/`*` and renders hyperlinks as `[text](url)`. Better for human-facing conversion.

## Usage

```bash
# Plain Markdown to stdout
python scripts/docx_to_md.py INPUT.docx

# Write to a file
python scripts/docx_to_md.py INPUT.docx OUTPUT.md

# Human-facing conversion with emphasis and real links
python scripts/docx_to_md.py INPUT.docx OUTPUT.md --emphasis --links
```

As a library:

```python
from docx_to_md import docx_to_md
md = docx_to_md("INPUT.docx")                                  # plain
pretty = docx_to_md("INPUT.docx", inline_emphasis=True, hyperlinks=True)
```

On Windows, prefix the command with `PYTHONUTF8=1` (or set `PYTHONIOENCODING=utf-8`) so characters like `★` or `–` do not crash the console encoder when printing to stdout.

## Requirements

- Python ≥ 3.9
- `python-docx` ≥ 1.1 (`pip install python-docx`). Version 1.1+ is required so that `Paragraph.text` includes hyperlink text and `Paragraph.iter_inner_content()` is available.

## Limitations

- Footnotes, comments, text boxes, equations, and embedded images are not extracted (their captions, if they are normal paragraphs, are). For the highest-fidelity extraction of unusual documents, the Anthropic `extract-text` CLI remains preferable where available.
- Nested tables are flattened to their top-level rows.
- Plain mode intentionally drops inline emphasis; use `--emphasis` if you need it.
