// Renders qa_items/<Category>.json into a branded Word doc using continuous
// Q#/A# numbering (never reset at a subcategory boundary), no tables — pure
// flowing text, AI-agent-friendly per the web-scraper-kb skill's FAQ rules.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle
} = require('docx');

const OUT = '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab';
const QA_DIR = path.join(OUT, 'qa_items');
const DOCS_DIR = path.join(OUT, 'docs_qa');
if (!fs.existsSync(DOCS_DIR)) fs.mkdirSync(DOCS_DIR, { recursive: true });

const PRIMARY = '00B14F'; // Grab brand green, confirmed with user on the earlier sample
const BODY_COLOR = '1A1A1A';
const CONTENT_WIDTH_DXA = 9360;
const SOURCE_ROOT = 'https://help.grab.com/merchant/en-my/';

// The colored banner is built with a single-cell borderless table purely as
// a layout/fill trick (per the skill's branded template) — this is not
// "table format" content and stays even though we render no data tables.
function filledBanner(titleText, subText) {
  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: [CONTENT_WIDTH_DXA],
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.NONE }, insideVertical: { style: BorderStyle.NONE }
    },
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
      shading: { type: ShadingType.CLEAR, fill: PRIMARY },
      margins: { top: 200, bottom: 200, left: 200, right: 200 },
      children: [
        new Paragraph({ children: [new TextRun({ text: titleText, bold: true, color: 'FFFFFF', size: 32 })] }),
        ...(subText ? [new Paragraph({ children: [new TextRun({ text: subText, italics: true, color: 'FFFFFF', size: 20 })], spacing: { before: 80 } })] : []),
      ],
    })] })],
  });
}

// One flowing answer paragraph: join the lines with real line breaks (not
// separate paragraphs) so each Q/A stays a single self-contained unit, per
// the branded template pattern.
function answerParagraph(prefix, lines) {
  const runs = [new TextRun({ text: prefix, bold: true, color: BODY_COLOR, size: 21 })];
  lines.forEach((line, i) => {
    if (i > 0) runs.push(new TextRun({ break: 1 }));
    runs.push(new TextRun({ text: line, color: BODY_COLOR, size: 21 }));
  });
  return new Paragraph({ children: runs, spacing: { after: 220 } });
}

function buildCategoryDoc(category, subcats) {
  const children = [];
  children.push(filledBanner(
    `Grab Merchant Help Centre (Malaysia) — ${category}`,
    `Merchant · AI-agent knowledge source · Source: ${SOURCE_ROOT}`
  ));
  children.push(new Paragraph({ text: '', spacing: { after: 200 } }));

  let qNum = 0;
  for (const [subcategory, items] of Object.entries(subcats)) {
    children.push(filledBanner(`${category} › ${subcategory}`, null));
    children.push(new Paragraph({ text: '', spacing: { after: 150 } }));

    for (const item of items) {
      qNum += 1;
      children.push(new Paragraph({
        children: [new TextRun({ text: `Q${qNum}. ${item.q}`, bold: true, color: BODY_COLOR, size: 22 })],
        spacing: { before: 160, after: 60 },
      }));
      children.push(answerParagraph(`A${qNum}. `, item.a));
    }
  }

  return new Document({
    sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } }, children }],
  });
}

async function main() {
  const onlyCategory = process.argv[2];
  const files = onlyCategory
    ? [`${onlyCategory.replace(/[^a-zA-Z0-9]+/g, '_')}.json`]
    : fs.readdirSync(QA_DIR).filter(f => f.endsWith('.json'));

  for (const file of files) {
    const category = file.replace(/_/g, ' ').replace(/\.json$/, '');
    const subcats = JSON.parse(fs.readFileSync(path.join(QA_DIR, file), 'utf8'));
    const doc = buildCategoryDoc(category, subcats);
    const outName = `Grab_Merchant_MY_HelpCentre_QA_${file.replace(/\.json$/, '')}.docx`;
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(path.join(DOCS_DIR, outName), buf);
    console.log('wrote', outName);
  }
}

main();
