# Review report format

The review report is a Word document the user can share with co-authors. It has a fixed structure designed for two reading modes: scan-to-the-table for someone who wants a checklist, and read-through for someone who wants the full justification.

## Structure

The report is built by `scripts/build_report.js` from a JSON input. The structure is:

1. **Title block** — "Cell Submission Readiness Review" (or "— Revision N" for follow-up reviews)
   - Manuscript title (italicised)
   - Reference standards line (CELLFFC, STAR Methods template, Information for Authors)
2. **Executive summary** — 2–3 short paragraphs. First paragraph states what's in good shape. Second paragraph states critical blockers in plain prose with bulleted enumeration. Third paragraph (if revision review) describes what was fixed since the previous review.
3. **What was fixed since the previous review** — only for revision reviews. Bullet list of resolved items from the previous report.
4. **Issues at a glance** — table with three columns: Severity, Issue, Location. One row per finding. Severity cells use the colour code below.
5. **Critical issues** — detailed entries for each CRITICAL finding, in order. Each entry uses the four-block format: Location, Issue, Action.
6. **Major formatting issues** — detailed entries for MAJOR findings.
7. **Minor issues worth addressing** — detailed entries for MINOR findings.
8. **Items still to prepare separately** — bulleted list of INFO findings (Highlights, graphical abstract, etc.).
9. **Submission readiness checklist** — consolidated bullet list with priority markers (ⓒ for critical, ⓜ for major, ⓢ for separate upload).
10. **End-of-review** marker — italic grey text.

## Finding entry format

Every detailed finding (sections 5, 6, 7) uses this four-block structure:

```
[SEVERITY in red/orange/yellow]  [Title in bold]
Location: [italicised section reference]
Issue: [Bold "Issue:" + plain text describing what's wrong, including a verbatim quote of the problematic text when short enough]
Action: [Bold blue "Action:" + plain text with the concrete edit to make]
```

The Action block must be a concrete edit, not a vague suggestion. Compare:

- ❌ "Fix the affiliation."
- ✅ "Expand affiliation 2 to: 'Buck Institute for Research on Aging, 8001 Redwood Boulevard, Novato, CA 94945, USA'."

If the action requires the user to verify something (e.g., the exact street address), say so:

- ✅ "Add postal codes to all three affiliations. Verify the exact street addresses with each institution before finalising."

## Severity colour code

The Word builder uses these hex colours:

| Severity | Hex | Visual |
|---|---|---|
| CRITICAL | `C00000` | red |
| MAJOR | `E97132` | orange |
| MINOR | `BF9000` | dark yellow |
| INFO | `548235` | green |

In the at-a-glance table, the severity cell is filled with the colour (white text). In detailed entries, the severity label is coloured text on a white background.

## JSON input schema

The report builder accepts a JSON file with this structure:

```json
{
  "manuscript_title": "Full title from the manuscript",
  "revision_number": 1,
  "executive_summary": {
    "good_state": "Paragraph describing what's in good shape (mention specific metrics: title length, word count, etc.)",
    "blockers": "Paragraph introducing critical issues",
    "blocker_list": ["short description 1", "short description 2", "short description 3"],
    "revision_note": "Optional paragraph for revision reviews summarising what was fixed"
  },
  "fixed_since_previous": [
    "Bullet text for each fixed item (only used for revision reviews)"
  ],
  "findings": [
    {
      "severity": "CRITICAL",
      "title": "Short title of the finding (1 line)",
      "summary": "One-line summary for the at-a-glance table",
      "location": "Section and subsection reference",
      "issue": "Plain text describing the problem, with verbatim quote when useful",
      "action": "Concrete edit the user should make"
    }
  ],
  "separate_uploads": [
    {
      "title": "Highlights and eTOC blurb",
      "description": "What it is and any specific guidance for this manuscript"
    }
  ],
  "checklist": [
    {
      "priority": "critical" | "major" | "separate",
      "text": "Bullet text"
    }
  ]
}
```

## Tone in the report

The report is written for a senior author, not a junior author who needs encouragement. Match that:

- **Direct, professional prose.** No exclamation points, no emoji, no hedging like "you might want to consider".
- **Concrete locations.** "Reference 13" not "one of the references"; "Results §Domain-specific competence" not "in Results".
- **Acknowledge what's debatable.** When an issue is genuinely a judgment call (e.g., "new" used in a non-claim sense), note explicitly: "This is borderline acceptable since…"
- **Quote the problematic text** when short. "The current text reads: '…'" gives the user something concrete to find and replace.
- **Suggest the replacement** when feasible. Cell editors and most co-authors don't want to be told a problem exists; they want to be told the fix.

## Length expectations

A typical first-pass review report is 8–12 pages. The at-a-glance table is usually 15–25 rows. The detailed findings section is 5–15 critical/major entries. If your draft has more than 25 findings, consolidate similar items (e.g., "ten references with duplicate DOIs" is one finding, not ten).

## Generating the report

```bash
node scripts/build_report.js <findings.json> <output.docx>
```

The script validates the JSON structure, builds the document, and writes it to the specified path. After running, validate with:

```bash
python /mnt/skills/public/docx/scripts/office/validate.py <output.docx>
```

Both should succeed with no errors. If validation fails, inspect the error and either fix the JSON input or adjust the builder script.
