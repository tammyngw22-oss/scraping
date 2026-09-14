const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle
} = require('docx');

const DOCS_DIR = 'gxbank_ext/docs';
if (!fs.existsSync(DOCS_DIR)) fs.mkdirSync(DOCS_DIR, { recursive: true });

const PRIMARY = '00B14F';
const BODY_COLOR = '1A1A1A';
const CONTENT_WIDTH_DXA = 9360;

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

function answerParagraph(prefix, lines) {
  const runs = [new TextRun({ text: prefix, bold: true, color: BODY_COLOR, size: 21 })];
  lines.forEach((line, i) => {
    if (i > 0) runs.push(new TextRun({ break: 1 }));
    runs.push(new TextRun({ text: line, color: BODY_COLOR, size: 21 }));
  });
  return new Paragraph({ children: runs, spacing: { after: 220 } });
}

function buildTopicDoc(topic, groups) {
  const children = [];
  children.push(filledBanner(
    `Grab Merchant Help Centre (Malaysia) — GXBank: ${topic}`,
    `Merchant · AI-agent knowledge source · Source: https://help.gxbank.my/`
  ));
  children.push(new Paragraph({ text: '', spacing: { after: 200 } }));

  let qNum = 0;
  for (const [group, items] of Object.entries(groups)) {
    children.push(filledBanner(`GXBank › ${topic} › ${group}`, null));
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
  return new Document({ sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } }, children }] });
}

async function main() {
  const byTopic = JSON.parse(fs.readFileSync('gxbank_ext/qa_by_topic.json', 'utf8'));
  for (const [topic, groups] of Object.entries(byTopic)) {
    const doc = buildTopicDoc(topic, groups);
    const fname = `Grab_Merchant_MY_HelpCentre_QA_GXBank_${topic.replace(/[^a-zA-Z0-9]+/g, '_')}.docx`;
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(`${DOCS_DIR}/${fname}`, buf);
    console.log('wrote', fname);
  }
}

main();
