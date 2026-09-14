// Combines ALL GXBank content — the original Grab-hosted "GXBank for
// Merchants" articles, the General FAQ, and the 6 Business topics from
// help.gxbank.my — into ONE file with continuous Q#/A# numbering
// throughout, no tables, no contact info.
const fs = require('fs');
const cheerio = require('cheerio');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle
} = require('docx');

// Build a group-name -> hub-page-URL map straight from the already-fetched
// topic pages, so each subcategory banner can cite exactly where that
// group's content lives on the source site (not just the overall topic).
function loadGroupUrls(htmlFile) {
  const html = fs.readFileSync(htmlFile, 'utf8');
  const $ = cheerio.load(html);
  const map = {};
  $('.mt-sortable-listing').each((i, groupEl) => {
    const a = $(groupEl).find('.mt-listing-detailed-title > a').first();
    const name = a.text().trim();
    const href = a.attr('href');
    if (name && href) map[name] = href;
  });
  return map;
}
const GROUP_URLS = {
  'Business Deposit': loadGroupUrls('gxbank_ext/topic_1.html'),
  'Account Opening': loadGroupUrls('gxbank_ext/topic_2.html'),
  'Payments & Transfers': loadGroupUrls('gxbank_ext/topic_3.html'),
  'Business Loan': loadGroupUrls('gxbank_ext/topic_4.html'),
  'App Security & Settings': loadGroupUrls('gxbank_ext/topic_5.html'),
  'Security & Privacy': loadGroupUrls('gxbank_ext/topic_6.html'),
};
const GENERAL_GROUP_URLS = loadGroupUrls('gxbank_ext/topic_general.html');

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

// --- Load all three sources ---
// 1. Original Grab-hosted "GXBank for Merchants" (from the main pipeline)
const original = JSON.parse(fs.readFileSync(
  '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab/qa_items/GXBank.json',
  'utf8'
)); // { "GXBank for Merchants": [ {q,a,url}, ... ] }

// 2. General FAQ (About GXBank, App Settings)
const general = JSON.parse(fs.readFileSync('gxbank_ext/qa_general.json', 'utf8')); // { group: [items] }

// 3. The 6 Business topics
const byTopic = JSON.parse(fs.readFileSync('gxbank_ext/qa_by_topic.json', 'utf8')); // { topic: { group: [items] } }

// --- Assemble one ordered section list: [{ sectionLabel, sourceUrl, items }] ---
const sections = [];
for (const [group, items] of Object.entries(original)) {
  // This section's own items already each carry their individual Grab
  // article URL; the section as a whole lives under this hub page.
  sections.push({
    label: `GXBank › ${group}`,
    sourceUrl: 'https://help.grab.com/merchant/en-my/20000116-GXBank-for-Merchants',
    items,
  });
}
for (const [group, items] of Object.entries(general)) {
  sections.push({
    label: `GXBank › General FAQ › ${group}`,
    sourceUrl: GENERAL_GROUP_URLS[group] || 'https://help.gxbank.my/general',
    items,
  });
}
for (const [topic, groups] of Object.entries(byTopic)) {
  for (const [group, items] of Object.entries(groups)) {
    sections.push({
      label: `GXBank › ${topic} › ${group}`,
      sourceUrl: (GROUP_URLS[topic] && GROUP_URLS[topic][group]) || null,
      items,
    });
  }
}

const totalItems = sections.reduce((a, s) => a + s.items.length, 0);
console.log('sections:', sections.length, 'total Q&A items:', totalItems);

// --- Build the single combined document ---
const children = [];
children.push(filledBanner(
  'Grab Merchant Help Centre (Malaysia) — GXBank (Complete)',
  'Merchant · AI-agent knowledge source · Sources: help.grab.com + help.gxbank.my'
));
children.push(new Paragraph({ text: '', spacing: { after: 200 } }));

let qNum = 0;
for (const section of sections) {
  const subText = section.sourceUrl ? `Source: ${section.sourceUrl}` : null;
  children.push(filledBanner(section.label, subText));
  children.push(new Paragraph({ text: '', spacing: { after: 150 } }));
  for (const item of section.items) {
    qNum += 1;
    children.push(new Paragraph({
      children: [new TextRun({ text: `Q${qNum}. ${item.q}`, bold: true, color: BODY_COLOR, size: 22 })],
      spacing: { before: 160, after: 60 },
    }));
    children.push(answerParagraph(`A${qNum}. `, item.a));
  }
}

const doc = new Document({ sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } }, children }] });
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('gxbank_ext/docs/GXBank.docx', buf);
  console.log('wrote GXBank.docx —', qNum, 'total Q&A items');
});
