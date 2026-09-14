// Extracts the 7 "General FAQ" articles and builds one combined,
// AI-agent-friendly Q&A docx — same rules as the other GXBank files
// (no tables, contact info stripped, continuous Q#/A# numbering).
const fs = require('fs');
const cheerio = require('cheerio');
const {
  Document, Packer, Paragraph, TextRun,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle
} = require('docx');

const manifest = JSON.parse(fs.readFileSync('gxbank_ext/manifest_general.json', 'utf8'));

const EMAIL_RE = /[\w.+-]+@[\w-]+\.[\w.-]+/i;
const PHONE_RE = /\+\d[\d\s-]{6,}\d/;
const DROP_KEYWORDS = [
  'hotline', 'live chat', 'customer support', 'contact them through',
  'reach out to our', 'contact us', 'email:', 'call us', 'whatsapp',
  'gxbank support', 'contact gxbank', "contact gx bank", '24/7 support',
  'in-app live chat', 'customer service', 'helpful link', 'need more information on',
];
function isContactLine(text) {
  const lower = text.toLowerCase();
  if (EMAIL_RE.test(text)) return true;
  if (PHONE_RE.test(text)) return true;
  return DROP_KEYWORDS.some(k => lower.includes(k));
}
function isEmojiOnly(text) { return !/[a-zA-Z0-9]/.test(text); }
function stripContactSentences(text) {
  const sentences = text.split(/(?<=[.!?])\s+/);
  return sentences.filter(s => !isContactLine(s)).join(' ').trim();
}
const DEAD_LINK_RE = /\b(learn more|find out more|click)\s*(about\s+\S+\s+)?here\b\.?/gi;
function stripDeadLinkPhrases(text) {
  return text.replace(DEAD_LINK_RE, '').replace(/\s{2,}/g, ' ').trim();
}
function cleanText(s) { return (s || '').replace(/\s+/g, ' ').trim(); }

function walk($, el, blocks) {
  const node = el;
  if (node.type === 'text') {
    const text = cleanText(node.data || '');
    if (text) blocks.push({ type: 'paragraph', text });
    return;
  }
  const tag = (node.tagName || '').toLowerCase();
  if (['script', 'style', 'button', 'input', 'label', 'textarea', 'form', 'iframe', 'footer'].includes(tag)) return;
  if (['h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(tag)) {
    const text = cleanText($(node).text());
    if (text) blocks.push({ type: 'heading', text });
    return;
  }
  if (tag === 'p') {
    const text = cleanText($(node).text());
    if (text) blocks.push({ type: 'paragraph', text });
    return;
  }
  if (tag === 'ul' || tag === 'ol') {
    const items = [];
    $(node).children('li').each((i, li) => {
      const text = cleanText($(li).clone().children('ul,ol').remove().end().text());
      if (text) items.push(text);
      $(li).children('ul,ol').each((j, nested) => {
        $(nested).children('li').each((k, nli) => {
          const ntext = cleanText($(nli).text());
          if (ntext) items.push(ntext);
        });
      });
    });
    if (items.length) blocks.push({ type: 'list', items });
    return;
  }
  if (tag === 'img' || tag === 'br' || tag === 'hr' || tag === 'svg') return;
  const hasBlockDescendant = $(node).find('p, ul, ol, h1, h2, h3, h4, h5, h6').length > 0;
  if (!hasBlockDescendant) {
    const text = cleanText($(node).text());
    if (text) blocks.push({ type: 'paragraph', text });
    return;
  }
  $(node).contents().each((i, child) => walk($, child, blocks));
}

function blockToLines(b) {
  const lines = [];
  if (b.type === 'heading') {
    const text = stripDeadLinkPhrases(stripContactSentences(b.text));
    if (text && !isEmojiOnly(text)) lines.push(text.endsWith(':') ? text : text + ':');
  } else if (b.type === 'paragraph') {
    const text = stripDeadLinkPhrases(stripContactSentences(b.text));
    if (text && !isEmojiOnly(text)) lines.push(text);
  } else if (b.type === 'list') {
    b.items
      .map(i => stripDeadLinkPhrases(stripContactSentences(i)))
      .filter(i => i && !isEmojiOnly(i))
      .forEach(item => lines.push(`- ${item}`));
  }
  return lines;
}

function mergeLabelLines(lines) {
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const isShortLabel = line.endsWith(':') && line.length < 40 && !line.startsWith('-');
    const nextIsAlsoLabel = i + 1 < lines.length && lines[i + 1].endsWith(':') && lines[i + 1].length < 40;
    if (isShortLabel && i + 1 < lines.length && !lines[i + 1].startsWith('-') && !nextIsAlsoLabel) {
      out.push(`${line} ${lines[i + 1]}`);
      i++;
    } else {
      out.push(line);
    }
  }
  return out;
}

// --- Extract all 7 articles into groups ---
const byGroup = {};
let empties = 0;
for (let i = 0; i < manifest.length; i++) {
  const m = manifest[i];
  const html = fs.readFileSync(`gxbank_ext/general_${i}.html`, 'utf8');
  const $ = cheerio.load(html);
  const title = cleanText($('h1#title').first().text()) || m.title;
  const contentEl = $('section.mt-content-container').first();
  const blocks = [];
  contentEl.contents().each((i2, child) => {
    const cls = (child.attribs && child.attribs.class) || '';
    if (cls.includes('mt-page-summary') || cls.includes('mt-subpage-listings') || cls.includes('mt-category-container')) return;
    walk($, child, blocks);
  });
  const lines = mergeLabelLines(blocks.flatMap(blockToLines));
  if (lines.length === 0) { empties++; continue; }
  byGroup[m.group] = byGroup[m.group] || [];
  byGroup[m.group].push({ q: title, a: lines, url: m.url });
}
console.log('items kept:', Object.values(byGroup).reduce((a, arr) => a + arr.length, 0), 'empty:', empties);

// --- Build one combined docx ---
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

const children = [];
children.push(filledBanner(
  'Grab Merchant Help Centre (Malaysia) — GXBank: General FAQ',
  'Merchant · AI-agent knowledge source · Source: https://help.gxbank.my/'
));
children.push(new Paragraph({ text: '', spacing: { after: 200 } }));

let qNum = 0;
for (const [group, items] of Object.entries(byGroup)) {
  children.push(filledBanner(`GXBank › General FAQ › ${group}`, null));
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

const doc = new Document({ sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } }, children }] });
Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('gxbank_ext/docs/GXBank_General_FAQ.docx', buf);
  console.log('wrote GXBank_General_FAQ.docx');
});
