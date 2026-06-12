#!/usr/bin/env python3
"""
Analyze a Cell-bound manuscript .docx and emit a JSON report of mechanical findings.

This script handles checks that can be done programmatically:
  - Title and Summary length
  - Main body word count
  - Keyword count
  - Reference list size and citation-to-reference cross-check
  - Duplicate DOI URLs in references
  - Template placeholder initials in Author Contributions
  - Required section headings present and in order
  - Highlights character count (if present)
  - eTOC blurb word count (if present)
  - Body sentences mis-styled as level-1 headings
  - Supplemental figure/table citation vs. SI listing consistency
  - Precise main-text (Introduction -> Discussion) word count, when python-docx
    is available

Text extraction prefers the Anthropic `extract-text` CLI; when it is not on the
PATH (e.g. on Windows), it falls back to the bundled `docx-to-markdown` skill's
python-docx converter installed alongside this skill.

Usage:
  python analyze.py <path-to-manuscript.docx> [--out findings.json]

If --out is omitted, the JSON is written to stdout.
"""

import argparse
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path

# Template-fictional initials. Presence in Author Contributions indicates the
# template's example block was not replaced with the real authors.
TEMPLATE_INITIALS = {
    "S.C.P.", "S.Y.W.", "A.B.", "M.E.", "A.N.V.", "N.A.V.",
    "M.E.V.", "C.K.B.", "A.A.D.", "N.L.W.", "A.A.",
}

# Template author names. Presence in the manuscript indicates leftover template
# example text.
TEMPLATE_AUTHOR_NAMES = {
    "Rosalind Franklin", "Lin Lanying", "Katherine Johnson",
}

# Required section headings (level 1) in order, as they should appear in
# the body before the references and STAR Methods.
REQUIRED_SECTIONS = [
    "SUMMARY",
    "KEYWORDS",
    "INTRODUCTION",
    "RESULTS",
    "DISCUSSION",
    "RESOURCE AVAILABILITY",
    "ACKNOWLEDGMENTS",
    "AUTHOR CONTRIBUTIONS",
    "DECLARATION OF INTERESTS",
    "REFERENCES",
    "STAR★METHODS",
]

# Required STAR Methods subsections (level 1 within the STAR section in the
# extract-text output, since extract-text emits a flat heading structure)
REQUIRED_STAR_SUBSECTIONS = [
    "KEY RESOURCES TABLE",
    "EXPERIMENTAL MODEL AND STUDY PARTICIPANT DETAILS",
    "METHOD DETAILS",
    "QUANTIFICATION AND STATISTICAL ANALYSIS",
    "ADDITIONAL RESOURCES",
]


def _local_converter():
    """Import docx_to_md from the sibling docx-to-markdown skill.

    Works for both the installed layout (~/.claude/skills/<skill>) and the repo
    layout (<repo>/skills/<skill>), since the docx-to-markdown skill is a sibling
    directory of this one in either case."""
    here = Path(__file__).resolve()
    skills_root = here.parent.parent.parent  # scripts/ -> skill/ -> skills-root/
    scripts_dir = skills_root / "docx-to-markdown" / "scripts"
    if (scripts_dir / "docx_to_md.py").exists():
        sys.path.insert(0, str(scripts_dir))
        from docx_to_md import docx_to_md  # type: ignore
        return docx_to_md
    return None


def extract_text(docx_path: Path) -> str:
    """Return the manuscript as Markdown.

    Prefers the Anthropic `extract-text` CLI for highest-fidelity extraction;
    falls back to the bundled docx-to-markdown converter when it is unavailable
    (e.g. on Windows, where `extract-text` is typically not installed)."""
    if shutil.which("extract-text"):
        try:
            result = subprocess.run(
                ["extract-text", str(docx_path)],
                capture_output=True, text=True, check=True,
            )
            return result.stdout
        except (subprocess.SubprocessError, OSError):
            pass  # fall through to the local converter
    convert = _local_converter()
    if convert is None:
        raise RuntimeError(
            "Could not extract text: the `extract-text` CLI is unavailable and the "
            "docx-to-markdown skill was not found alongside this skill. Install the "
            "docx-to-markdown skill (and `pip install python-docx`) or put `extract-text` "
            "on the PATH."
        )
    return convert(str(docx_path))


def docx_main_text_words(docx_path: Path) -> int | None:
    """Precise Cell main-text word count: Introduction through the end of
    Discussion, excluding embedded figure legends and tables. Returns None if
    python-docx is unavailable or the document has no Introduction heading.

    The flat `extract-text`/Markdown body count over-counts the Cell limit because
    it also includes figure legends, tables, and back matter; this walks the .docx
    directly to count only the narrative the 7,000-word limit applies to."""
    try:
        import docx as _dx
        from docx.oxml.text.paragraph import CT_P
        from docx.text.paragraph import Paragraph as _Para
    except ImportError:
        return None
    try:
        d = _dx.Document(str(docx_path))
    except Exception:
        return None

    def is_heading(style: str | None) -> bool:
        s = (style or "").lower()
        return s.startswith("heading") or s == "title"

    def wc(s: str) -> int:
        return len([w for w in re.sub(r"\s+", " ", s).split() if any(ch.isalnum() for ch in w)])

    in_main = False
    in_legend = False
    total = 0
    found_intro = False
    for child in d.element.body.iterchildren():
        if not isinstance(child, CT_P):
            continue  # skip tables entirely
        p = _Para(child, d)
        style = p.style.name if p.style else "Normal"
        txt = p.text.strip()
        if is_heading(style):
            up = txt.upper()
            if up == "INTRODUCTION":
                in_main, in_legend, found_intro = True, False, True
                continue
            if up == "RESOURCE AVAILABILITY":
                in_main = False
        if not in_main or not txt:
            continue
        if re.match(r"^(Figure|Table)\s", txt, re.I):
            in_legend = True
            continue
        if is_heading(style):  # a real subheading resumes prose
            in_legend = False
            total += wc(txt)
            continue
        if in_legend:
            continue
        total += wc(txt)
    return total if found_intro else None


def get_title(text: str) -> str:
    """The title is the first non-empty line, with bold markers stripped."""
    for line in text.splitlines():
        s = line.strip()
        if s:
            return re.sub(r"\*+", "", s).strip()
    return ""


def section_text(text: str, heading: str, next_headings: list) -> str:
    """Return the text between `heading` and the next heading in `next_headings`,
    or the end of the document. Heading match is case-insensitive on the
    extracted markdown `# HEADING` form."""
    pattern = rf"#\s+{re.escape(heading)}\s*\n+(.*?)(?=\n#\s+(?:{'|'.join(re.escape(h) for h in next_headings)})\b|\Z)"
    m = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
    return m.group(1).strip() if m else ""


def count_words(s: str) -> int:
    """Plain word count, ignoring markdown formatting characters."""
    s = re.sub(r"\*+|#+|\|[^\n]*\|", " ", s)
    s = re.sub(r"https?://\S+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return len(s.split())


def get_summary(text: str) -> str:
    m = re.search(r"#\s+SUMMARY\s*\n+(.*?)\n+#\s+KEYWORDS", text, re.DOTALL | re.IGNORECASE)
    return re.sub(r"\*+", "", m.group(1)).strip() if m else ""


def get_keywords(text: str) -> list:
    m = re.search(r"#\s+KEYWORDS\s*\n+(.*?)\n+#\s+INTRODUCTION", text, re.DOTALL | re.IGNORECASE)
    if not m:
        return []
    raw = re.sub(r"\*+", "", m.group(1)).strip()
    return [k.strip() for k in raw.split(",") if k.strip()]


def get_highlights(text: str) -> list:
    """If a Highlights section is in the file (often misplaced there), return the
    bullets. Each bullet returned with trailing semicolons/periods stripped."""
    m = re.search(r"#\s+Highlights\s*\n+(.*?)\n+#\s+\w", text, re.DOTALL | re.IGNORECASE)
    if not m:
        return []
    bullets = re.findall(r"^-\s+(.+)$", m.group(1), re.MULTILINE)
    out = []
    for b in bullets:
        b_clean = re.sub(r"\*+", "", b).strip().rstrip(";").rstrip(".")
        out.append(b_clean)
    return out


def get_etoc(text: str) -> str:
    m = re.search(r"#\s+E?[Tt][Oo][Cc].*?\n+(.*?)\n+#\s+SUMMARY", text, re.DOTALL)
    return re.sub(r"\*+", "", m.group(1)).strip() if m else ""


def count_references(text: str) -> int:
    m = re.search(r"#\s+REFERENCES\s*\n+(.*?)\n+#\s+STAR", text, re.DOTALL | re.IGNORECASE)
    if not m:
        return 0
    return len(re.findall(r"^\d+\.\s", m.group(1), re.MULTILINE))


def find_citations(text: str) -> set:
    """Find numeric in-text citations using conservative patterns. Cell-style
    superscript citations appear in the extracted text as plain integers.
    To avoid false positives from data values (age ranges, sample sizes, etc.),
    we accept only patterns where the citation context is unambiguous:

      1. Number directly attached to a word with no space ("mortality1.")
      2. "et al. N" pattern (citation right after author name)
      3. Comma-separated list of small integers ("12,13", "14,15,16")

    Range patterns like "12–16" are deliberately excluded because age ranges
    ("55–65"), age decades ("70–79"), and other data ranges produce too many
    false positives. Real range citations are rare in practice — authors
    usually write "12,13,14,15,16" instead.
    """
    body = re.sub(r"#\s+REFERENCES\s*\n+.*?\n+#\s+STAR", "\n\n# STAR", text, flags=re.DOTALL | re.IGNORECASE)
    body = re.sub(r"\*+", "", body)
    cites = set()

    # Pattern 1: number directly attached to a word (no space)
    # Strongest signal: "mortality1.", "patients30,31,32"
    for m in re.finditer(r"\b[a-zA-Z]{3,}(\d+(?:,\d+){0,5})(?=[\s\.\,\;\)\:])", body):
        for part in m.group(1).split(","):
            if part.isdigit() and 1 <= int(part) <= 200:
                cites.add(int(part))

    # Pattern 2: "et al. N" or "et al.,N"
    for m in re.finditer(r"et\s+al\.?\s*,?\s*(\d+)", body):
        n = int(m.group(1))
        if 1 <= n <= 200:
            cites.add(n)

    # Pattern 3: comma-separated integer lists where the integers are small
    # and close together. Excludes thousands-separator patterns like "1,234".
    for m in re.finditer(r"(?<![\d.])\b(\d{1,3})(?:,(\d{1,3})){1,5}\b(?![\d.])", body):
        # Collect all captured integers
        full = m.group(0)
        parts = full.split(",")
        ints = [int(p) for p in parts if p.isdigit()]
        if not ints:
            continue
        # Citation lists are tight: all small, close range
        if max(ints) <= 200 and (max(ints) - min(ints)) < 50:
            # Exclude thousands-separator patterns: thousands separators have
            # exactly 3 digits after each comma, e.g. "1,234" or "12,345"
            tail_parts = parts[1:]
            if all(len(p) == 3 for p in tail_parts):
                continue  # likely thousands separator
            for n in ints:
                cites.add(n)

    return cites


def find_duplicate_dois(text: str) -> list:
    """Return reference numbers (as ints) whose entry contains the DOI URL twice."""
    m = re.search(r"#\s+REFERENCES\s*\n+(.*?)\n+#\s+STAR", text, re.DOTALL | re.IGNORECASE)
    if not m:
        return []
    refs = m.group(1)
    dup = []
    for line in refs.split("\n"):
        if line.strip():
            doi_count = len(re.findall(r"https?://doi\.org", line))
            if doi_count >= 2:
                num_match = re.match(r"^\s*(\d+)\.", line)
                if num_match:
                    dup.append(int(num_match.group(1)))
    return dup


def find_template_residue(text: str) -> dict:
    """Search the Author Contributions section for template-fictional initials,
    and the body (excluding references) for template-fictional author names.

    Template initials like 'M.E.' or 'A.A.' will legitimately appear in
    reference author lists (e.g., Levine, M.E.) and figure legends, so we
    restrict the initials check to the Author Contributions section where
    the residue actually causes harm. Author names are unique enough that
    we can search the whole body."""
    ac_match = re.search(
        r"#\s+AUTHOR CONTRIBUTIONS\s*\n+(.*?)\n+#\s+",
        text, re.DOTALL | re.IGNORECASE
    )
    ac_text = ac_match.group(1) if ac_match else ""

    found_initials = sorted(
        i for i in TEMPLATE_INITIALS
        if re.search(r"(?<![A-Za-z])" + re.escape(i), ac_text)
    )

    # For author names, exclude references section
    body_no_refs = re.sub(
        r"#\s+REFERENCES\s*\n+.*?\n+#\s+STAR",
        "\n\n# STAR", text, flags=re.DOTALL | re.IGNORECASE
    )
    found_names = sorted(n for n in TEMPLATE_AUTHOR_NAMES if n in body_no_refs)

    return {"initials": found_initials, "names": found_names}


def find_sections(text: str) -> list:
    """Return all level-1 headings (lines starting with '# ' but not '## ' or more)."""
    out = []
    for m in re.finditer(r"^#\s+(.+)$", text, re.MULTILINE):
        h = re.sub(r"\*+", "", m.group(1)).strip()
        out.append(h)
    return out


def find_sentence_headings(text: str) -> list:
    """Find suspected body sentences mis-formatted as level-1 headings.
    Heuristic: heading text contains a verb-like word AND ends with sentence
    punctuation OR is unusually long (>20 words)."""
    suspicious = []
    for h in find_sections(text):
        # Skip recognised section names
        if h.upper() in REQUIRED_SECTIONS or h in REQUIRED_STAR_SUBSECTIONS:
            continue
        if h.upper().startswith("DECLARATION OF") or h.upper().startswith("SUPPLEMENTAL"):
            continue
        if h.upper().startswith("FIGURE TITLES") or h.upper().startswith("TABLES AND"):
            continue
        if h.upper().startswith("HIGHLIGHTS") or h.upper().lower().startswith("etoc"):
            continue
        word_count = len(h.split())
        ends_with_period = h.rstrip().endswith(".")
        # A genuine section heading is usually ≤8 words; a sentence is usually ≥10
        if word_count >= 12 or (word_count >= 8 and ends_with_period):
            suspicious.append(h)
    return suspicious


def get_si_listing(text: str) -> dict:
    """Parse the SUPPLEMENTAL INFORMATION listing to find which figures/tables
    are documented as present. Returns a dict with 'figures' and 'tables' sets."""
    m = re.search(r"#\s+SUPPLEMENTAL INFORMATION\s*\n+(.*?)\n+#\s+FIGURE", text, re.DOTALL | re.IGNORECASE)
    if not m:
        return {"figures": set(), "tables": set()}
    listing = re.sub(r"\*+", "", m.group(1))
    # Find figure ranges like "Figures S1–S6" or "Figure S3"
    figures = set()
    for m in re.finditer(r"Figures?\s+S(\d+)\s*[–\-]\s*S?(\d+)", listing):
        for n in range(int(m.group(1)), int(m.group(2)) + 1):
            figures.add(n)
    for m in re.finditer(r"Figure\s+S(\d+)\b", listing):
        figures.add(int(m.group(1)))
    # Find tables
    tables = set()
    for m in re.finditer(r"Tables?\s+S(\d+)", listing):
        tables.add(int(m.group(1)))
    return {"figures": figures, "tables": tables}


def find_si_citations(text: str) -> dict:
    """Find supplemental figures and tables cited in the body. Returns sets."""
    body = text.split("# REFERENCES")[0]
    # Strip markdown asterisks so 'S****3' becomes 'S3'
    flat = re.sub(r"\*+", "", body)
    figs = set()
    for m in re.finditer(r"Figure\s+S(\d+)", flat):
        figs.add(int(m.group(1)))
    for m in re.finditer(r"Figures\s+S(\d+)(?:[,\s]+S?(\d+))?(?:[,\s]+S?(\d+))?", flat):
        for g in m.groups():
            if g and g.isdigit():
                figs.add(int(g))
    tabs = set()
    for m in re.finditer(r"Table\s+S(\d+)", flat):
        tabs.add(int(m.group(1)))
    return {"figures": figs, "tables": tabs}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("manuscript", type=Path, help="Path to manuscript .docx")
    parser.add_argument("--out", type=Path, default=None, help="Output JSON path (default: stdout)")
    args = parser.parse_args()

    if not args.manuscript.exists():
        print(f"Error: file not found: {args.manuscript}", file=sys.stderr)
        sys.exit(1)

    text = extract_text(args.manuscript)

    findings = []

    def add(severity, category, title, location, evidence, action):
        findings.append({
            "severity": severity,
            "category": category,
            "title": title,
            "location": location,
            "evidence": evidence,
            "suggested_action": action,
        })

    # ----- Title -----
    title = get_title(text)
    title_chars = len(title)
    if title_chars > 145:
        add("CRITICAL", "title", "Title exceeds 145-character limit",
            "Title", f"Title is {title_chars} characters: '{title}'",
            "Reduce title to ≤145 characters.")

    # ----- Summary -----
    summary = get_summary(text)
    summary_words = count_words(summary)
    if summary_words > 150:
        add("CRITICAL", "summary", "Summary exceeds 150-word limit",
            "Summary", f"Summary is {summary_words} words.",
            "Trim summary to ≤150 words.")
    elif summary_words == 150:
        add("MINOR", "summary", "Summary at exactly the 150-word ceiling",
            "Summary", f"Summary is exactly 150 words.",
            "Consider trimming to ~140 words to leave editorial headroom.")

    if re.search(r"\d+[\,–\-]\d+", summary) and re.search(r"\(\d", summary):
        # weak citation heuristic
        pass

    # ----- Keywords -----
    keywords = get_keywords(text)
    if len(keywords) > 10:
        add("CRITICAL", "keywords", "More than 10 keywords",
            "Keywords", f"Found {len(keywords)} keywords.",
            "Reduce to ≤10 keywords.")

    # ----- Main body word count -----
    # The flat pre-References count (`body_words`) over-counts the Cell limit: it
    # also sweeps in figure legends, tables, and back matter. When python-docx is
    # available we compute the precise Introduction->Discussion count and judge the
    # 7,000-word limit against that; otherwise we only report the loose number.
    WORD_LIMIT = 7000
    body_words = None
    star_split = text.split("# STAR★METHODS")
    if len(star_split) >= 2:
        body = star_split[0]
        body = re.sub(r"#\s+REFERENCES.*$", "", body, flags=re.DOTALL)
        body_words = count_words(body)

    main_text_words = None
    if args.manuscript.suffix.lower() == ".docx":
        main_text_words = docx_main_text_words(args.manuscript)

    if main_text_words is not None:
        if main_text_words > WORD_LIMIT:
            add("CRITICAL", "wordcount", "Main text exceeds 7,000-word limit",
                "Introduction → end of Discussion",
                f"Introduction through Discussion is approximately {main_text_words} words (limit: 7,000).",
                f"Trim ~{main_text_words - WORD_LIMIT + 100} words to leave editorial headroom.")
    elif body_words is not None and body_words > WORD_LIMIT:
        # No precise count available: the loose number over-counts, so flag for
        # manual verification rather than asserting a breach.
        add("INFO", "wordcount", "Approximate body word count is high",
            "Title page → before References",
            f"Approximate pre-References word count is ~{body_words}. This includes figure "
            f"legends, tables, and back matter, so it over-counts the Cell main-text limit. "
            f"Install python-docx for a precise Introduction→Discussion count.",
            "Verify the Introduction-through-Discussion word count against the 7,000-word limit.")

    # ----- Highlights -----
    highlights = get_highlights(text)
    if highlights:
        # Their presence in the main file is itself an issue
        add("CRITICAL", "highlights", "Highlights placed inside main manuscript file",
            "Top of manuscript after title page",
            f"Found {len(highlights)} Highlights bullets in the main file.",
            "Move Highlights and eTOC blurb to a separate Word file as required by Cell.")
        if len(highlights) < 3 or len(highlights) > 4:
            add("MAJOR", "highlights", f"Number of Highlights ({len(highlights)}) outside 3–4 range",
                "Highlights block",
                f"Cell requires 3–4 Highlights bullets; found {len(highlights)}.",
                "Adjust the number of Highlights to 3 or 4.")
        over_limit = [(i + 1, len(b), b) for i, b in enumerate(highlights) if len(b) > 85]
        if over_limit:
            evidence = "; ".join(f"#{n} is {c} chars" for n, c, _ in over_limit)
            add("CRITICAL", "highlights", "Highlights exceed 85-character limit",
                "Highlights block", evidence,
                "Rewrite each over-limit highlight to ≤85 characters.")

    # ----- eTOC blurb -----
    etoc = get_etoc(text)
    if etoc:
        # heading text check
        if not re.search(r"#\s+eTOC", text):
            add("MINOR", "etoc", "eTOC blurb heading capitalisation",
                "eTOC blurb heading",
                "Heading reads 'Etoc Blurb' or similar variant.",
                "Use 'eTOC blurb' (lowercase 'e', uppercase 'TOC', lowercase 'blurb').")
        etoc_words = count_words(etoc)
        if etoc_words > 80:
            add("MAJOR", "etoc", "eTOC blurb exceeds 80-word target",
                "eTOC blurb",
                f"Blurb is {etoc_words} words.",
                "Trim to ~50–80 words.")

    # ----- References and citations -----
    ref_count = count_references(text)
    cites = find_citations(text)
    missing = sorted([c for c in cites if c > ref_count])
    if missing:
        add("CRITICAL", "references", f"In-text citation(s) exceed reference list length",
            "References list vs. body citations",
            f"Reference list contains {ref_count} entries; in-text citations include {missing}.",
            f"Add the missing reference entries: {missing}. Verify each by searching the text for the surrounding context.")

    dup_dois = find_duplicate_dois(text)
    if dup_dois:
        add("MAJOR", "references", "Duplicate DOI URLs in references",
            "References list",
            f"References with duplicated DOI URLs: {dup_dois}",
            "Remove the duplicated DOI in each affected reference. A regex find-and-replace for '(https://doi\\.org/\\S+) \\1' catches all instances.")

    # ----- Template residue -----
    residue = find_template_residue(text)
    if residue["initials"]:
        add("CRITICAL", "template-residue", "Template placeholder initials present",
            "Likely in Author Contributions",
            f"Found template-fictional initials: {residue['initials']}.",
            "Replace placeholder initials with the manuscript's actual author initials. These are from the Cell STAR Methods template's fictional authors.")
    if residue["names"]:
        add("CRITICAL", "template-residue", "Template author names present",
            "Manuscript text",
            f"Found template author names: {residue['names']}.",
            "Remove template example text; replace with this manuscript's content.")

    # ----- Required sections present -----
    sections = find_sections(text)
    sections_upper = [s.upper() for s in sections]
    for req in REQUIRED_SECTIONS:
        if req not in sections_upper:
            add("CRITICAL", "structure", f"Required section missing: {req}",
                "Manuscript structure", f"Section heading '{req}' not found.",
                f"Add the {req} section. See references/template-structure.md for placement.")
    for req in REQUIRED_STAR_SUBSECTIONS:
        if req not in sections_upper:
            add("MAJOR", "structure", f"STAR Methods subsection missing: {req}",
                "STAR Methods", f"Subsection '{req}' not found.",
                f"Add the {req} subsection at the appropriate position in STAR Methods.")

    # ----- Body sentences mis-styled as headings -----
    sus_headings = find_sentence_headings(text)
    if sus_headings:
        evidence = "; ".join(f"'{h[:80]}...'" for h in sus_headings[:5])
        add("CRITICAL", "formatting", "Body paragraphs styled as level-1 headings",
            "Manuscript body",
            f"{len(sus_headings)} suspected sentences carry Heading 1 style. First few: {evidence}",
            "Select each affected paragraph and apply Normal/body-text style instead of Heading 1.")

    # ----- Supplemental items consistency -----
    si_listed = get_si_listing(text)
    si_cited = find_si_citations(text)
    fig_missing_from_listing = sorted(si_cited["figures"] - si_listed["figures"])
    fig_listed_not_cited = sorted(si_listed["figures"] - si_cited["figures"])
    if fig_missing_from_listing:
        add("MAJOR", "supplemental", "Supplemental figure(s) cited but not in SI listing",
            "Supplemental Information section",
            f"Figures cited in text but not in SI listing: S{', S'.join(str(n) for n in fig_missing_from_listing)}",
            "Either add the missing figure(s) to Document S1 and update the SI listing, or correct the in-text callout.")
    if fig_listed_not_cited:
        add("MINOR", "supplemental", "Supplemental figure(s) listed but never cited",
            "Manuscript body",
            f"Figures in SI listing but not cited in text: S{', S'.join(str(n) for n in fig_listed_not_cited)}",
            "Cite the listed figure(s) in the main text, or remove from the SI listing.")

    # ----- Output -----
    output = {
        "manuscript_path": str(args.manuscript),
        "title": title,
        "title_chars": title_chars,
        "summary_words": summary_words,
        "main_text_words": main_text_words,
        "main_body_words_loose": body_words,
        "keyword_count": len(keywords),
        "reference_count": ref_count,
        "citations_found": sorted(cites),
        "section_headings": sections,
        "findings": findings,
        "finding_counts": {
            "CRITICAL": sum(1 for f in findings if f["severity"] == "CRITICAL"),
            "MAJOR": sum(1 for f in findings if f["severity"] == "MAJOR"),
            "MINOR": sum(1 for f in findings if f["severity"] == "MINOR"),
            "INFO": sum(1 for f in findings if f["severity"] == "INFO"),
        },
    }

    out_json = json.dumps(output, indent=2, ensure_ascii=False)
    if args.out:
        args.out.write_text(out_json, encoding="utf-8")
        print(f"Wrote {args.out} ({len(findings)} findings)", file=sys.stderr)
    else:
        print(out_json)


if __name__ == "__main__":
    main()
