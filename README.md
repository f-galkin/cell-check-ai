# cell-check-ai

Two [Claude Code](https://claude.com/claude-code) skills for preparing and converting scientific manuscripts, built to run **anywhere Python runs** — no Linux-only CLI, no Node.js required.

| Skill | What it does |
|---|---|
| [`cell-manuscript-review`](cell-manuscript-review) | Audits a `.docx` manuscript for **Cell** submission readiness against the Final File Requirements and the STAR Methods template, then produces a severity-coded Word review report. |
| [`docx-to-markdown`](docx-to-markdown) | Converts any Word `.docx` to clean Markdown with `python-docx` only. Used standalone, or as the cross-platform extraction fallback for the review skill. |

## Why this exists

The original review workflow assumed an Anthropic Linux environment: it shelled out to the `extract-text` CLI for parsing and to a Node.js + `docx` builder for the report. Neither is present on a typical Windows/macOS setup, so the workflow failed at the first step. This repo makes both skills self-contained on a plain Python install:

- **Extraction** prefers `extract-text` when available and otherwise falls back to the `docx-to-markdown` converter.
- **The report builder** ships in both Node (`build_report.js`) and Python (`build_report.py`) flavours, sharing one JSON schema.
- **The word-count check** was corrected: it now reports the precise Introduction-through-Discussion count (the figure the 7,000-word limit applies to) rather than a flat pre-References count that swept in figure legends, tables, and back matter.
- **A run-level style helper** (`check_style.py`) was added for checks the text-based analyzer structurally cannot do — gene-symbol / organism-binomial italicisation and the precise word count.

## Install

Copy (or symlink) each skill folder into your Claude skills directory:

```bash
# Linux / macOS
cp -r cell-manuscript-review docx-to-markdown ~/.claude/skills/

# Windows (Git Bash)
cp -r cell-manuscript-review docx-to-markdown "$USERPROFILE/.claude/skills/"
```

Both skills must sit in the **same** parent directory (your `~/.claude/skills/` root, or this repo root) so the review skill can find the `docx-to-markdown` converter as a sibling.

## Requirements

- Python ≥ 3.9 with [`python-docx`](https://python-docx.readthedocs.io/) ≥ 1.1 (`pip install python-docx`) — needed for extraction fallback, the style helper, the precise word count, and `build_report.py`.
- *(Optional)* Node.js + the `docx` npm package — only if you prefer `build_report.js`.
- *(Optional)* the Anthropic `extract-text` CLI — used preferentially when on the PATH.

On Windows, prefix commands with `PYTHONUTF8=1` so characters like `★` (in `STAR★METHODS`) or `–` do not crash the console encoder.

## Quick start

Review a manuscript:

```bash
cd cell-manuscript-review
python scripts/analyze.py MANUSCRIPT.docx > findings_raw.json          # mechanical checks
python scripts/check_style.py MANUSCRIPT.docx --genes KDM1A,SOCS3      # word count + italics
# (read the manuscript, compile a findings.json per references/findings-schema.example.json)
python scripts/build_report.py findings.json Review.docx               # build the report
```

Convert a Word document to Markdown:

```bash
python docx-to-markdown/scripts/docx_to_md.py INPUT.docx OUTPUT.md             # plain
python docx-to-markdown/scripts/docx_to_md.py INPUT.docx OUTPUT.md --emphasis --links
```

## Scope

`cell-manuscript-review` is a **formatting and submission-readiness** review — it does not peer-review the science, edit the manuscript, or write the Highlights / graphical abstract / Declaration of Interests for you. It is specific to the journal *Cell* (not the other Cell Press titles).

## License

[MIT](LICENSE). The `cell-manuscript-review` checklists encode the publicly documented Cell Final File Requirements and STAR Methods article template.
