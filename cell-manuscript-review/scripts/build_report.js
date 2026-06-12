#!/usr/bin/env node
/*
 * Build a Cell submission readiness review report (.docx) from a JSON findings file.
 *
 * Usage:
 *   node build_report.js <findings.json> <output.docx>
 *
 * See references/report-format.md for the JSON schema and report structure.
 */

const fs = require('fs');
const path = require('path');

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat,
} = require('docx');

if (process.argv.length < 4) {
  console.error("Usage: node build_report.js <findings.json> <output.docx>");
  process.exit(1);
}

const findingsPath = process.argv[2];
const outputPath = process.argv[3];

if (!fs.existsSync(findingsPath)) {
  console.error(`Error: findings file not found: ${findingsPath}`);
  process.exit(1);
}

const data = JSON.parse(fs.readFileSync(findingsPath, 'utf-8'));

// ---------- styling helpers ----------
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };

const SEV_COLORS = {
  CRITICAL: "C00000",
  MAJOR:    "E97132",
  MINOR:    "BF9000",
  INFO:     "548235",
};

const P = (text, opts = {}) => new Paragraph({
  spacing: { before: 60, after: 60 },
  ...opts,
  children: Array.isArray(text) ? text : [new TextRun({ text })],
});

const H1 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  spacing: { before: 280, after: 140 },
  children: [new TextRun({ text: t, bold: true })],
});

const H2 = (t) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  spacing: { before: 200, after: 100 },
  children: [new TextRun({ text: t, bold: true })],
});

const bullet = (children) => new Paragraph({
  numbering: { reference: "bullets", level: 0 },
  spacing: { before: 40, after: 40 },
  children: Array.isArray(children) ? children : [new TextRun({ text: children })],
});

// Severity-colored cell for the at-a-glance table
const sevCell = (label) => new TableCell({
  borders,
  width: { size: 1200, type: WidthType.DXA },
  shading: { fill: SEV_COLORS[label] || "808080", type: ShadingType.CLEAR },
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: label, bold: true, size: 18, color: "FFFFFF" })],
  })],
});

const txtCell = (text, width) => new TableCell({
  borders,
  width: { size: width, type: WidthType.DXA },
  margins: { top: 60, bottom: 60, left: 100, right: 100 },
  children: [new Paragraph({
    children: [new TextRun({ text: text, size: 20 })],
  })],
});

const sumRow = (severity, summary, location) => new TableRow({
  children: [
    sevCell(severity),
    txtCell(summary, 5160),
    txtCell(location, 2880),
  ],
});

// Detailed finding entry (severity label + title, location, issue, action)
function findingBlocks(f) {
  const color = SEV_COLORS[f.severity] || "404040";
  return [
    new Paragraph({
      spacing: { before: 180, after: 60 },
      children: [
        new TextRun({ text: f.severity + "  ", bold: true, color: color, size: 20 }),
        new TextRun({ text: f.title, bold: true, size: 24 }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [
        new TextRun({ text: "Location: ", bold: true, italics: true, size: 20 }),
        new TextRun({ text: f.location, italics: true, size: 20 }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 40 },
      children: [
        new TextRun({ text: "Issue: ", bold: true, size: 20 }),
        new TextRun({ text: f.issue, size: 20 }),
      ],
    }),
    new Paragraph({
      spacing: { before: 0, after: 60 },
      children: [
        new TextRun({ text: "Action: ", bold: true, color: "1F4E79", size: 20 }),
        new TextRun({ text: f.action, size: 20 }),
      ],
    }),
  ];
}

// ---------- assemble document ----------

const children = [];

// Title
const titleText = data.revision_number && data.revision_number > 1
  ? `Cell Submission Readiness Review — Revision ${data.revision_number}`
  : "Cell Submission Readiness Review";

children.push(new Paragraph({
  alignment: AlignmentType.LEFT,
  spacing: { after: 100 },
  children: [new TextRun({ text: titleText, bold: true, size: 40, color: "1F4E79" })],
}));

children.push(new Paragraph({
  spacing: { after: 240 },
  children: [
    new TextRun({ text: "Manuscript: ", bold: true, size: 22 }),
    new TextRun({ text: data.manuscript_title || "(untitled)", italics: true, size: 22 }),
  ],
}));

children.push(new Paragraph({
  spacing: { after: 240 },
  children: [
    new TextRun({ text: "Reference standards: ", bold: true, size: 20 }),
    new TextRun({ text: "Cell Final File Requirements (CELLFFC.pdf); STAR Methods Article Template; Cell Information for Authors.", size: 20 }),
  ],
}));

// Executive summary
if (data.executive_summary) {
  children.push(H1("Executive summary"));
  if (data.executive_summary.good_state) children.push(P(data.executive_summary.good_state));
  if (data.executive_summary.blockers) children.push(P(data.executive_summary.blockers));
  if (Array.isArray(data.executive_summary.blocker_list)) {
    for (const b of data.executive_summary.blocker_list) {
      children.push(bullet(b));
    }
  }
  if (data.executive_summary.revision_note) children.push(P(data.executive_summary.revision_note));
}

// Fixed since previous (only for revision reviews)
if (Array.isArray(data.fixed_since_previous) && data.fixed_since_previous.length > 0) {
  children.push(H1("What was fixed since the previous review"));
  children.push(P("These items from the previous review are now resolved:"));
  for (const item of data.fixed_since_previous) {
    children.push(bullet(item));
  }
}

// Issues at a glance
const findings = Array.isArray(data.findings) ? data.findings : [];
if (findings.length > 0) {
  children.push(H1("Issues at a glance"));
  const rows = [
    new TableRow({
      tableHeader: true,
      children: [
        new TableCell({
          borders,
          shading: { fill: "1F4E79", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          width: { size: 1200, type: WidthType.DXA },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: "Severity", bold: true, color: "FFFFFF" })],
          })],
        }),
        new TableCell({
          borders,
          shading: { fill: "1F4E79", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          width: { size: 5160, type: WidthType.DXA },
          children: [new Paragraph({
            children: [new TextRun({ text: "Issue", bold: true, color: "FFFFFF" })],
          })],
        }),
        new TableCell({
          borders,
          shading: { fill: "1F4E79", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 100, right: 100 },
          width: { size: 2880, type: WidthType.DXA },
          children: [new Paragraph({
            children: [new TextRun({ text: "Location", bold: true, color: "FFFFFF" })],
          })],
        }),
      ],
    }),
  ];
  // Sort findings: CRITICAL → MAJOR → MINOR → INFO, preserving order within each tier
  const sevOrder = { CRITICAL: 0, MAJOR: 1, MINOR: 2, INFO: 3 };
  const sorted = [...findings].sort((a, b) => (sevOrder[a.severity] ?? 9) - (sevOrder[b.severity] ?? 9));
  for (const f of sorted) {
    rows.push(sumRow(f.severity, f.summary || f.title, f.location));
  }
  children.push(new Table({
    width: { size: 9240, type: WidthType.DXA },
    columnWidths: [1200, 5160, 2880],
    rows,
  }));

  // Detailed sections by severity
  const grouped = { CRITICAL: [], MAJOR: [], MINOR: [], INFO: [] };
  for (const f of findings) {
    if (grouped[f.severity]) grouped[f.severity].push(f);
  }

  if (grouped.CRITICAL.length > 0) {
    children.push(H1("Critical issues"));
    children.push(P("These items should be addressed before the file is uploaded."));
    for (const f of grouped.CRITICAL) {
      children.push(...findingBlocks(f));
    }
  }

  if (grouped.MAJOR.length > 0) {
    children.push(H1("Major formatting and structural issues"));
    for (const f of grouped.MAJOR) {
      children.push(...findingBlocks(f));
    }
  }

  if (grouped.MINOR.length > 0) {
    children.push(H1("Minor issues worth addressing"));
    for (const f of grouped.MINOR) {
      children.push(...findingBlocks(f));
    }
  }
}

// Separate uploads
if (Array.isArray(data.separate_uploads) && data.separate_uploads.length > 0) {
  children.push(H1("Items to prepare and upload separately"));
  children.push(P("These items belong outside the main manuscript file."));
  for (const item of data.separate_uploads) {
    children.push(H2(item.title));
    if (item.description) children.push(P(item.description));
  }
}

// Checklist
if (Array.isArray(data.checklist) && data.checklist.length > 0) {
  children.push(H1("Submission readiness checklist"));
  children.push(P("Priority-ordered to-do list. ⓒ = critical/submission-blocking; ⓜ = major formatting fix; ⓢ = separate upload."));
  const marker = { critical: "ⓒ", major: "ⓜ", separate: "ⓢ" };
  const markerColor = { critical: "C00000", major: "E97132", separate: "548235" };
  for (const item of data.checklist) {
    const m = marker[item.priority] || "•";
    const c = markerColor[item.priority] || "404040";
    children.push(bullet([
      new TextRun({ text: m + "  ", bold: true, color: c }),
      new TextRun({ text: item.text }),
    ]));
  }
}

// End marker
children.push(new Paragraph({
  spacing: { before: 480, after: 120 },
  children: [new TextRun({ text: "End of review.", italics: true, color: "808080" })],
}));

// Build document
const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 30, bold: true, font: "Arial", color: "1F4E79" },
        paragraph: { spacing: { before: 280, after: 140 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 200, after: 100 }, outlineLevel: 1 } },
    ],
  },
  numbering: {
    config: [{
      reference: "bullets",
      levels: [
        { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 540, hanging: 270 } } } },
      ],
    }],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
      },
    },
    children,
  }],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync(outputPath, buf);
  console.error(`Wrote ${outputPath} (${buf.length} bytes, ${findings.length} findings)`);
});
