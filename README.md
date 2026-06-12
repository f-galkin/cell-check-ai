# cell-check-ai

Two [Claude Code](https://claude.com/claude-code) skills for preparing and converting scientific manuscripts, built to run **anywhere Python runs** — no Linux-only CLI, no Node.js required.

| Skill | What it does |
|---|---|
| [`cell-manuscript-review`](cell-manuscript-review) | Audits a `.docx` manuscript for **Cell** submission readiness against the Final File Requirements and the STAR Methods template, then produces a severity-coded Word review report. |
| [`docx-to-markdown`](docx-to-markdown) | Converts any Word `.docx` to clean Markdown with `python-docx` only. Used standalone, or as the cross-platform extraction fallback for the review skill. |

## Features

- **Runs anywhere Python runs** — no Node.js or Linux-only CLI required. The review skill extracts text with the Anthropic `extract-text` CLI when it is present, and otherwise falls back to the bundled `docx-to-markdown` converter.
- **Mechanical checks** — title and summary length, keyword count, in-text-citation ↔ reference cross-check, duplicate DOIs, required sections and their order, and leftover-template residue.
- **Precise main-text word count** — measures Introduction through Discussion (the text Cell's 7,000-word limit applies to), excluding figure legends, tables, and back matter.
- **Run-level style checks** — gene-symbol and organism-binomial italicisation, which a text-only pass cannot see.
- **Word report builder in two flavours** — Python (`build_report.py`, no Node) and Node (`build_report.js`), sharing one JSON schema; the output is a severity-coded `.docx` (critical / major / minor / info).

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

## Templates

[`templates/STAR-Methods-article-template.docx`](templates/STAR-Methods-article-template.docx) is the official Cell Press STAR Methods article template, included as a starting point for authors (source and attribution in [`templates/NOTICE.md`](templates/NOTICE.md)). It is Cell Press property and is **not** covered by this repository's MIT license.

## Official Cell author resources

The skill's checklists distil these pages; consult them directly for the authoritative requirements. Links verified **12 June 2026**:

- **Information for authors (Cell):** https://www.cell.com/cell/information-for-authors
- **Journal policies (Cell):** https://www.cell.com/cell/information-for-authors/journal-policies
- **Article templates:** https://www.cell.com/information-for-authors/article-templates
- **STAR Methods author guide:** https://www.cell.com/information-for-authors/star-authors-guide
- **STAR Methods supplemental information:** https://www.cell.com/information-for-authors/star-supplemental-information
- **Declaration of interests (policy):** https://www.cell.com/declaration-of-interests
- **Declaration of interests form (PDF):** https://www.cell.com/pb/assets/raw/shared/forms/di_form.pdf

## Keeping current with Cell's requirements

> [!IMPORTANT]
> These skills encode Cell's Final File Requirements and STAR Methods template **as of June 2026**. Cell Press revises its author guidelines periodically. If Cell changes its formatting requirements, the skills will not know — update them to match:
> - `cell-manuscript-review/references/ffc-checklist.md` — the section-by-section rule set;
> - `cell-manuscript-review/references/template-structure.md` — the required section order and STAR Methods subsections;
> - `cell-manuscript-review/scripts/analyze.py` — any hard-coded limits (title 145 chars, summary 150 words, main text 7,000 words, highlights 85 chars) and the `REQUIRED_SECTIONS` / `REQUIRED_STAR_SUBSECTIONS` lists;
> - `templates/STAR-Methods-article-template.docx` — replace with the current Cell template.
>
> Always confirm a review against the live pages above before relying on it for an actual submission.

## License

[MIT](LICENSE). The `cell-manuscript-review` checklists encode the publicly documented Cell Final File Requirements and STAR Methods article template.
