// Shared styling helpers for NETS RAG knowledge-base Word docs.
// Layout: full-width banner header, plain Category/Subcategory headings,
// a blue breadcrumb banner per subcategory with a Source line, an
// Audience/Persona metadata line, then continuously-numbered Q&A pairs
// (Q1/A1, Q2/A2, ... restarting at 1 per document).
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel,
  ShadingType, WidthType, Table, TableRow, TableCell,
} = require("docx");

// NETS brand palette (pulled from nets.com.sg site CSS: --wp--preset--color--*)
const COLOR = {
  blue: "0046AD",
  blueLight: "E6EDF7",
  red: "E70033",
  black: "101820",
  darkGray: "626469",
  white: "FFFFFF",
};
const FONT = "Barlow";

function makeBuilder() {
  let qaCounter = 0;

  function bannerBox({ titleRuns, subtitleRuns, titleSize = 44, padTop = 200, padBottom = 200 }) {
    const rows = [
      new Paragraph({
        spacing: { after: subtitleRuns ? 60 : 0 },
        children: titleRuns.map((r) => new TextRun({ ...r, bold: true, color: COLOR.white, font: FONT, size: titleSize })),
      }),
    ];
    if (subtitleRuns) {
      rows.push(new Paragraph({ children: subtitleRuns.map((r) => new TextRun({ ...r, color: COLOR.white, font: FONT, size: 19 })) }));
    }
    return new Table({
      width: { size: 9350, type: WidthType.DXA },
      columnWidths: [9350],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 9350, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: COLOR.blue, color: "auto" },
              margins: { top: padTop, bottom: padBottom, left: 260, right: 260 },
              children: rows,
            }),
          ],
        }),
      ],
    });
  }

  function noteBox(text) {
    return new Table({
      width: { size: 9350, type: WidthType.DXA },
      columnWidths: [9350],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: 9350, type: WidthType.DXA },
              shading: { type: ShadingType.CLEAR, fill: COLOR.blueLight, color: "auto" },
              margins: { top: 150, bottom: 150, left: 200, right: 200 },
              children: [
                new Paragraph({
                  children: [
                    new TextRun({ text: "AI-AGENT CONTENT NOTE  ", bold: true, color: COLOR.blue, font: FONT, size: 18 }),
                    new TextRun({ text, italics: true, color: COLOR.black, font: FONT, size: 18 }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    });
  }

  function docHeader(title, audience, sourceUrl) {
    return [
      bannerBox({
        titleRuns: [{ text: title }],
        subtitleRuns: [{ text: `${audience} · Source: ${sourceUrl}` }],
        titleSize: 44,
      }),
      new Paragraph({ spacing: { after: 200 }, children: [] }),
      noteBox(
        "This document is intended for ingestion into an AI Agent's RAG knowledge base. Every answer is self-contained — no “click here,” no external contact numbers or emails, no unresolved references."
      ),
    ];
  }

  function categoryHeading(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 400, after: 160 },
      children: [new TextRun({ text, color: COLOR.blue, bold: true, font: FONT, size: 40 })],
    });
  }

  function subcategoryHeading(text) {
    return new Paragraph({
      heading: HeadingLevel.HEADING_2,
      spacing: { before: 320, after: 100 },
      children: [new TextRun({ text, color: COLOR.blue, bold: true, font: FONT, size: 28 })],
    });
  }

  function subcategoryBanner(category, subcategory, sourceUrl) {
    return bannerBox({
      titleRuns: [{ text: `${category} › ${subcategory}` }],
      subtitleRuns: [{ text: `Source: ${sourceUrl}`, italics: true }],
      titleSize: 30,
      padTop: 160,
      padBottom: 160,
    });
  }

  function metaLine(audience, persona) {
    return new Paragraph({
      spacing: { before: 100, after: 160 },
      children: [
        new TextRun({ text: "Audience: ", bold: true, color: COLOR.darkGray, font: FONT, size: 18 }),
        new TextRun({ text: audience, color: COLOR.darkGray, font: FONT, size: 18 }),
        new TextRun({ text: "   |   Persona: ", bold: true, color: COLOR.darkGray, font: FONT, size: 18 }),
        new TextRun({ text: persona, color: COLOR.darkGray, font: FONT, size: 18 }),
      ],
    });
  }

  function bulletParas(bullets) {
    return bullets.map(
      (b) =>
        new Paragraph({
          bullet: { level: 0 },
          spacing: { after: 40 },
          children: [new TextRun({ text: b, color: COLOR.black, font: FONT, size: 21 })],
        })
    );
  }

  // answer: string, or { intro, bullets, outro? }, or { intro, groups: [{heading, bullets}], outro? }
  function qa(question, answer) {
    qaCounter += 1;
    const n = qaCounter;
    const out = [
      new Paragraph({
        spacing: { before: 220, after: 60 },
        children: [
          new TextRun({ text: `Q${n}. `, bold: true, color: COLOR.blue, font: FONT, size: 21 }),
          new TextRun({ text: question, bold: true, color: COLOR.black, font: FONT, size: 21 }),
        ],
      }),
    ];

    if (typeof answer === "string") {
      out.push(
        new Paragraph({
          spacing: { after: 120 },
          children: [
            new TextRun({ text: `A${n}. `, bold: true, color: COLOR.blue, font: FONT, size: 21 }),
            new TextRun({ text: answer, color: COLOR.black, font: FONT, size: 21 }),
          ],
        })
      );
    } else {
      out.push(
        new Paragraph({
          spacing: { after: 60 },
          children: [
            new TextRun({ text: `A${n}. `, bold: true, color: COLOR.blue, font: FONT, size: 21 }),
            new TextRun({ text: answer.intro, color: COLOR.black, font: FONT, size: 21 }),
          ],
        })
      );
      if (answer.groups) {
        for (const g of answer.groups) {
          out.push(
            new Paragraph({
              spacing: { before: 100, after: 40 },
              children: [new TextRun({ text: g.heading, bold: true, color: COLOR.darkGray, font: FONT, size: 21 })],
            })
          );
          out.push(...bulletParas(g.bullets));
        }
      } else {
        out.push(...bulletParas(answer.bullets));
      }
      if (answer.outro) {
        out.push(
          new Paragraph({
            spacing: { before: 60, after: 120 },
            children: [new TextRun({ text: answer.outro, color: COLOR.black, font: FONT, size: 21 })],
          })
        );
      }
    }
    return out;
  }

  return { docHeader, categoryHeading, subcategoryHeading, subcategoryBanner, metaLine, qa, get count() { return qaCounter; } };
}

async function writeDoc(outPath, children) {
  const doc = new Document({
    styles: { default: { document: { run: { font: FONT, size: 21, color: COLOR.black } } } },
    sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } } }, ].map((s) => ({ ...s, children })),
  });
  const buf = await Packer.toBuffer(doc);
  require("fs").writeFileSync(outPath, buf);
}

module.exports = { makeBuilder, writeDoc, COLOR, FONT };
