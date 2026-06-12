---
name: cell-manuscript-review
description: "Review a manuscript draft for Cell journal submission readiness against the Final File Requirements (CELLFFC) and the STAR Methods article template. Produces a Word-document review report with severity-coded findings — critical (submission-blocking), major (formatting), minor (style), and info (separate-upload items) — covering title and summary length, section order, citation-to-reference cross-checks, Highlights and eTOC blurb compliance, Author Contributions and Declaration of Interests presence, STAR Methods subsection completeness, reference format, and items to prepare separately. Trigger this skill whenever a user uploads a manuscript .docx and asks to review, check, audit, proofread, or prepare it for Cell submission, including phrases like 'is this ready for Cell', 'check against Cell guidelines', 'review my STAR Methods paper', 'submission checklist', or 'final file requirements' — even if Cell is not named explicitly but the user is comparing against a Cell template or mentions STAR Methods."
---

# Cell manuscript review

A submission-readiness review for manuscripts targeting the journal **Cell**. Compares a draft against the Cell Final File Requirements (CELLFFC) and the STAR Methods article template, then produces a Word-document review report the user can share with co-authors.

## When to use

Use this skill whenever the user uploads a manuscript `.docx` and wants it checked for Cell submission. Typical phrasings:
- "Help me review this for Cell submission"
- "Is this manuscript ready for submission?"
- "Check this against Cell's guidelines"
- "Audit my STAR Methods paper"
- "Run a submission checklist over this draft"

The skill is **specific to the journal Cell**. It does not cover other Cell Press journals (Cell Reports, Cell Metabolism, etc.) — those have slightly different requirements and would need a different skill.

## Inputs the user typically provides

- The manuscript `.docx` (required)
- The STAR Methods article template `.docx` (sometimes — confirms which template version was used)
- A link to the Final File Requirements PDF (sometimes — but the requirements are encoded in `references/ffc-checklist.md` so a missing link is not blocking)

If the user provides a template or FFC link, briefly skim them to catch any wording the template encodes that the local checklist might have missed. Otherwise proceed directly from the local references.

## Workflow

The review proceeds in four steps. Don't skip any of them — the mechanical checks find errors the eye misses (e.g., citing reference 38 when the bibliography stops at 33), and the human-judgment checks find errors the script can't see (e.g., placeholder template initials that look like real authors).

### Step 1 — Extract and run mechanical checks

Run the analyzer script on the manuscript. It extracts text, computes word/character counts, cross-checks in-text citations against the reference list, detects duplicate DOIs, and surfaces template residue (placeholder initials, fictional template author names). Output is JSON.

```bash
python scripts/analyze.py <path-to-manuscript.docx> > /tmp/findings.json
```

**Extraction is cross-platform.** The analyzer prefers the Anthropic `extract-text` CLI when it is on the PATH; otherwise it falls back to the bundled `docx-to-markdown` skill's python-docx converter (so the analyzer runs on Windows too). For the fallback to work, install the sibling `docx-to-markdown` skill alongside this one and ensure `python-docx` is available (`pip install python-docx`). On Windows, prefix commands with `PYTHONUTF8=1` so the `★` in `STAR★METHODS` and characters like `–` do not crash the console encoder.

**On word count:** the JSON reports `main_text_words` (the precise Introduction-through-Discussion count, computed from the .docx with python-docx — this is the figure the 7,000-word limit applies to) and `main_body_words_loose` (a flat pre-References count that also includes figure legends, tables, and back matter, so it over-counts). Judge the limit against `main_text_words`; quote `main_body_words_loose` to no one. If python-docx is unavailable, only the loose number exists and the analyzer flags it as INFO for manual verification rather than asserting a breach.

Read `/tmp/findings.json` and inspect every flagged item. Each flagged item has a `severity`, `category`, `location`, `evidence`, and `suggested_action`. Treat these as starting points, not final findings — verify each by opening the relevant location in the extracted text (use the extracted Markdown and `grep` for the location keyword) to confirm the finding is real and the suggested action makes sense in context.

### Step 2 — Read the manuscript and run the judgment checks

The mechanical script can't evaluate things like prose quality, subheading capitalisation consistency, or whether the Declaration of Interests genuinely discloses commercial relationships. Read the manuscript end-to-end and check the items in `references/ffc-checklist.md`. The checklist is organised by manuscript section; work through it in order.

Pay particular attention to these recurring failure modes (each one cost a previous user a desk-rejection-worthy issue):
- **Template residue in Author Contributions** — placeholder initials like S.C.P., S.Y.W., A.B., M.E.V. that come from the Cell template's fictional authors (Rosalind Franklin, Lin Lanying, Katherine Johnson). The analyzer flags these, but verify by reading the section.
- **Swapped Author Contributions and Declaration of Interests** — content under one heading actually belongs under the other.
- **Citations beyond the reference list** — in-text citation 34 when the bibliography stops at 33. The analyzer catches this; manually verify the missing references are actually missing (not just unrecognised by the parser).
- **Sentences mis-styled as headings** — body paragraphs that begin with `#` or carry a Heading style, rendering as oversized blue text. The analyzer catches Level-1 mis-styles; manually check Level-2 cases by reading the document structure.
- **Highlights over the 85-character limit** — Cell-mandated, easy to miss because authors write what they want to say first and don't count.
- **Affiliations missing department/postal code** — Cell requires both; commercial-entity affiliations sometimes only have city/country.
- **Gene symbols and organism binomials not italicised** — Cell italicises gene symbols (not their protein products) and organism binomials. The text-based analyzer cannot see run-level italics, so check this with the helper below.

For the run-level checks the analyzer cannot perform (precise main-text word count and italicisation), run the style helper, passing the gene symbols you noticed while reading:

```bash
python scripts/check_style.py <path-to-manuscript.docx> --genes KDM1A,SOCS3,CXCL12,SIRT1
```

It reports the precise Introduction→Discussion word count and, for each organism binomial (auto) and supplied gene symbol, whether the occurrences are italic (`ITAL`), roman (`ROMAN`), or mixed (`PARTIAL`). Requires `python-docx`.

### Step 3 — Compile findings with severity

Categorise each finding into one of four severity levels. The categorisation matters because the user works through the report by severity, and miscategorising trivial items as critical wastes their attention.

**CRITICAL** — would trigger return-without-review or a desk-revision request from the Cell editorial office. Examples: template placeholder text left in Author Contributions; missing references cited in text; missing Declaration of Interests; main body over 7,000-word limit; Highlights over 85-character limit; required section missing.

**MAJOR** — formatting issues that delay scheduling but don't block initial editorial assessment. Examples: STAR Methods missing the Additional Resources subsection; affiliations incomplete; duplicate DOIs; KRT uses bare reference number; Lead Contact footnote missing from byline.

**MINOR** — style and consistency issues that copy-editing would catch. Examples: subheading capitalisation alternates between title and sentence case; thousands separators applied inconsistently; reference format inconsistency for a single entry; typos; figure-callout spacing.

**INFO** — items the user needs to prepare but that are not part of the main manuscript file. Examples: Highlights and eTOC blurb (separate Word file); graphical abstract; Cell Declaration of Interests PDF form; high-resolution figure files.

### Step 4 — Generate the Word report

Run the report builder, passing it the findings as JSON. The builder produces a Word document with the standard structure (executive summary → at-a-glance table → critical issues → major issues → minor issues → separate-upload items → revised checklist).

Two interchangeable builders accept the **same** JSON schema. Use the Node builder when Node.js and the `docx` npm package are available; otherwise use the python-docx builder (no Node required, runs on Windows):

```bash
# Node (preferred where available)
node scripts/build_report.js <path-to-findings.json> <path-to-output.docx>

# Python fallback (requires python-docx)
python scripts/build_report.py <path-to-findings.json> <path-to-output.docx>
```

Validate the output before presenting. If the docx skill's validator is installed, use it:
```bash
python /mnt/skills/public/docx/scripts/office/validate.py <path-to-output.docx>
```
Otherwise, validate by reopening the file with python-docx (a clean open confirms well-formed OOXML):
```bash
python -c "import docx,sys; d=docx.Document(sys.argv[1]); print('OK:', len(d.paragraphs), 'paragraphs,', len(d.tables), 'tables')" <path-to-output.docx>
```

Then share the report with the user (use `present_files` if available; otherwise give the output path). Also give a brief inline summary of the most critical findings in prose — the user often wants the top-line takeaways before they open the document.

## Iterating on revisions

When the user submits a revised version, follow the same four steps but adapt the report:
- Compare against the previous review (held in conversation history) and explicitly note **what was fixed** in an early section of the report.
- Distinguish **persistent issues** (carried over from previous review) from **new issues** (introduced in the revision).
- Mark items in the checklist with their priority: ⓒ critical, ⓜ major formatting, ⓢ separate upload.

## Reference files

Read these as needed:

- `references/ffc-checklist.md` — Cell Final File Requirements distilled into a section-by-section checklist. Read this for **Step 2**.
- `references/template-structure.md` — Required sections in order, with which ones are mandatory vs optional and what each contains. Read this when checking section order in **Step 1** or when verifying that subsections like "Additional Resources" are present.
- `references/report-format.md` — Details on the review report structure (severity colours, table layout, finding entry format). Read this only if the user asks for a custom output format, or when adapting the report builder for an unusual situation.
- `references/findings-schema.example.json` — Annotated example showing the JSON structure that both report builders accept. Refer to this when assembling the findings JSON in **Step 4**.

## Scripts

- `scripts/analyze.py` — mechanical checks; emits findings JSON. Cross-platform extraction (prefers `extract-text`, falls back to the `docx-to-markdown` skill).
- `scripts/check_style.py` — run-level checks python-docx can see but the text analyzer cannot: precise main-text word count and gene/organism italicisation.
- `scripts/build_report.js` — Node report builder (requires Node + the `docx` package).
- `scripts/build_report.py` — equivalent python-docx report builder (no Node; requires `python-docx`).

## Dependencies

- **python-docx** (`pip install python-docx`, ≥ 1.1) — required for the extraction fallback, `check_style.py`, the precise word count, and `build_report.py`.
- **docx-to-markdown skill** — install alongside this skill (same `skills/` root) to enable extraction without the `extract-text` CLI.
- **Node.js + `docx`** — optional; only for `build_report.js`.

## Official Cell references and freshness

The checklists in `references/` distil Cell's publicly documented requirements. The authoritative sources (verified 12 June 2026) are:

- Information for authors (Cell) — https://www.cell.com/cell/information-for-authors
- Journal policies (Cell) — https://www.cell.com/cell/information-for-authors/journal-policies
- Article templates — https://www.cell.com/information-for-authors/article-templates
- STAR Methods author guide — https://www.cell.com/information-for-authors/star-authors-guide
- STAR Methods supplemental information — https://www.cell.com/information-for-authors/star-supplemental-information
- Declaration of interests (policy + PDF form) — https://www.cell.com/declaration-of-interests

**These requirements can change.** Cell Press revises its guidelines periodically, and this skill will not know when it does. If a review touches a requirement that may have moved (word/character limits, required sections, STAR Methods subsections, reference format), confirm it against the live pages above. When Cell changes its formatting, update `references/ffc-checklist.md`, `references/template-structure.md`, and the hard-coded limits and `REQUIRED_SECTIONS` / `REQUIRED_STAR_SUBSECTIONS` lists in `scripts/analyze.py`. Tell the user when a finding rests on a limit that may have changed since the date above.

## Tone and language

The user is typically a senior researcher preparing for submission. Match that tone:
- Professional prose, no exclamation points, no emoji.
- Lists are appropriate in the report (it's a checklist document); avoid them in inline conversational summaries.
- Be specific about locations ("Results §Domain-specific competence", not "in Results").
- Provide the **action** as a concrete edit, not a vague suggestion. "Reformat reference 13 to: 'Galkin, F., …'" is useful; "Fix reference 13" is not.
- When an issue is potentially debatable (e.g., "new" vs "novel" usage), note that explicitly so the user can decide.

## What this skill does not do

- **It does not edit the manuscript.** The output is a review report, not a marked-up manuscript. If the user wants tracked changes, that's a separate task using the docx skill.
- **It does not write the Highlights, eTOC blurb, or Declaration of Interests for them.** It flags that these are missing or non-compliant; drafting the content is a separate request.
- **It does not generate the graphical abstract.** It flags that one is required.
- **It does not verify scientific claims.** It is a formatting and submission-readiness review, not a peer review.
