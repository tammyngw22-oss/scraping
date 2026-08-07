// Shared house style for the Standard Chartered Singapore RAG FAQ documents.
// Both build_faq_docx.js (Help) and build_cc_docx.js (Credit Cards) use this,
// so the two deliverables cannot drift apart in styling or structure.
const {
  Document, Paragraph, TextRun, ExternalHyperlink, HeadingLevel,
  AlignmentType, ShadingType, BorderStyle, LevelFormat, convertInchesToTwip,
  Table, TableRow, TableCell, WidthType, TableLayoutType,
} = require("docx");

// ---------------------------------------------------------------- palette ---
const SC_BLUE = "0061C7";  // primary
const NAVY = "2C3A87";     // category bars, subcategory heading text
const TINT = "D9E7F7";     // subcategory box background
const GREEN = "92E773";    // banner eyebrow
const BLACK = "000000";

const NUM = "bulletList";

// A4 (11906 twips) less the 1080-twip side margins set on the section.
const CONTENT_W = 11906 - 1080 * 2; // 9746
const BANNER_PAD = 260;             // single padding value, all four sides
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "auto" };
const edge = { style: BorderStyle.SINGLE, size: 8, space: 6, color: SC_BLUE };

// Rich-text mini-DSL: a string is plain text, ["label", url] is a hyperlink.
function runs(parts, { bold = false, size = 22, color = BLACK } = {}) {
  return parts.map((p) =>
    Array.isArray(p)
      ? new ExternalHyperlink({
          link: p[1],
          children: [new TextRun({ text: p[0], size, color: SC_BLUE, underline: {} })],
        })
      : new TextRun({ text: p, bold, size, color })
  );
}

// Both banner lines live in ONE table cell whose margins supply the only
// horizontal padding, so the two lines cannot drift apart: neither paragraph
// carries an indent of its own.
function banner(title) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    layout: TableLayoutType.FIXED,
    margins: { top: BANNER_PAD, bottom: BANNER_PAD, left: BANNER_PAD, right: BANNER_PAD },
    borders: {
      top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER,
      insideHorizontal: NO_BORDER, insideVertical: NO_BORDER,
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: CONTENT_W, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, fill: SC_BLUE, color: "auto" },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0, line: 260 },
                indent: { left: 0, right: 0 },
                children: [
                  new TextRun({
                    text: "Standard Chartered Singapore",
                    bold: true, smallCaps: true, size: 20, color: GREEN, characterSpacing: 80,
                  }),
                ],
              }),
              // Heading 1 lives inside the banner so the title is TOC-navigable
              // without being repeated as a separate visible heading below it.
              new Paragraph({
                heading: HeadingLevel.HEADING_1,
                spacing: { before: 60, after: 0, line: 420 },
                indent: { left: 0, right: 0 },
                children: [new TextRun({ text: title, bold: true, size: 40, color: "FFFFFF" })],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function categoryBar(name) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    shading: { type: ShadingType.CLEAR, fill: NAVY, color: "auto" },
    spacing: { before: 360, after: 180, line: 320 },
    indent: { left: 160, right: 160 },
    children: [new TextRun({ text: name, bold: true, size: 28, color: "FFFFFF" })],
  });
}

function subcategoryBox(category, name, url) {
  const shade = { type: ShadingType.CLEAR, fill: TINT, color: "auto" };
  return [
    new Paragraph({
      heading: HeadingLevel.HEADING_3,
      shading: shade,
      border: { top: edge, left: edge, right: edge },
      spacing: { before: 240, after: 0, line: 300 },
      indent: { left: 160, right: 160 },
      children: [new TextRun({ text: `${category} › ${name}`, bold: true, size: 24, color: NAVY })],
    }),
    new Paragraph({
      shading: shade,
      border: { bottom: edge, left: edge, right: edge },
      spacing: { before: 0, after: 200, line: 260 },
      indent: { left: 160, right: 160 },
      children: [
        new TextRun({ text: "Source: ", italics: true, size: 18, color: NAVY }),
        new ExternalHyperlink({
          link: url,
          children: [new TextRun({ text: url, italics: true, size: 18, color: SC_BLUE, underline: {} })],
        }),
      ],
    }),
  ];
}

function qa(n, q, aParts) {
  const out = [
    new Paragraph({
      spacing: { before: 220, after: 60 },
      children: [
        new TextRun({ text: `Q${n}: `, bold: true, size: 22, color: SC_BLUE }),
        ...runs(Array.isArray(q) ? q : [q], { bold: true }),
      ],
    }),
  ];
  aParts.forEach((part, i) => {
    if (part && part.bullets) {
      part.bullets.forEach((b) => {
        out.push(new Paragraph({
          numbering: { reference: NUM, level: 0 },
          spacing: { before: 20, after: 20 },
          children: runs(b), // b is always an array of DSL parts
        }));
      });
    } else {
      out.push(new Paragraph({
        spacing: { before: i === 0 ? 0 : 80, after: 60 },
        children: [
          ...(i === 0 ? [new TextRun({ text: `A${n}: `, bold: true, size: 22, color: SC_BLUE })] : []),
          ...runs(Array.isArray(part) ? part : [part]),
        ],
      }));
    }
  });
  return out;
}

const B = (bullets) => ({ bullets });

// Assembles banner + Category (H2) > Subcategory (H3) > Q&A, numbering the
// Q&A pairs continuously across the WHOLE document — never resetting.
function buildDocument({ title, docTitle, description, categories }) {
  const children = [
    banner(title),
    new Paragraph({ spacing: { before: 0, after: 0, line: 240 }, children: [] }), // spacer below banner
  ];
  let counter = 0;
  for (const cat of categories) {
    children.push(categoryBar(cat.name));
    for (const sub of cat.subs) {
      children.push(...subcategoryBox(cat.name, sub.name, sub.url));
      for (const [q, a] of sub.qas) children.push(...qa(++counter, q, a));
    }
  }

  const doc = new Document({
    title: docTitle,
    description,
    creator: "Standard Chartered Singapore FAQ extract",
    numbering: {
      config: [{
        reference: NUM,
        levels: [{
          level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: convertInchesToTwip(0.4), hanging: convertInchesToTwip(0.2) } } },
        }],
      }],
    },
    styles: {
      default: { document: { run: { font: "Calibri", size: 22, color: BLACK } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Calibri", size: 40, bold: true, color: "FFFFFF" } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Calibri", size: 28, bold: true, color: "FFFFFF" } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: "Calibri", size: 24, bold: true, color: NAVY } },
      ],
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4 — CONTENT_W is derived from this
          margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 },
        },
      },
      children,
    }],
  });
  return { doc, count: counter };
}

module.exports = { buildDocument, B, SC_BLUE, NAVY, TINT, GREEN, BLACK };
