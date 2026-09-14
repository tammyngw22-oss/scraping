const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle
} = require('docx');

const OUT = '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab';
const EXTRACTED_DIR = path.join(OUT, 'extracted');
const DOCS_DIR = path.join(OUT, 'docs');
if (!fs.existsSync(DOCS_DIR)) fs.mkdirSync(DOCS_DIR, { recursive: true });

const PRIMARY = process.env.PRIMARY_COLOR || '00B14F';
const BODY_COLOR = '1A1A1A';
const CONTENT_WIDTH_DXA = 9360;

const KNOWN_CATEGORIES = ['GrabFood', 'GrabMart', 'Payment Services', 'GrabExpress', 'Dine Out', 'Hubbo', 'Reservations', 'Marketing', 'Financing', 'GXBank'];
const SOURCE_ROOT = 'https://help.grab.com/merchant/en-my/';

const index = JSON.parse(fs.readFileSync(path.join(OUT, 'extracted_index.json'), 'utf8'));

function noBorderTable(children) {
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
      children,
    })] })],
  });
}

function filledBanner(titleText, subText) {
  return noBorderTable([
    new Paragraph({ children: [new TextRun({ text: titleText, bold: true, color: 'FFFFFF', size: 32 })] }),
    ...(subText ? [new Paragraph({ children: [new TextRun({ text: subText, italics: true, color: 'FFFFFF', size: 20 })], spacing: { before: 80 } })] : []),
  ]);
}

function contentTable(rows) {
  // Rows can be ragged (a source <td rowspan> means later rows have fewer
  // <td>s than the header) — pad to a uniform column count so every row
  // lines up under the right header instead of drifting left.
  const maxCols = Math.max(...rows.map(r => r.length), 1);
  const colWidth = Math.floor(CONTENT_WIDTH_DXA / maxCols);
  const columnWidths = new Array(maxCols).fill(colWidth);

  const trs = rows.map((r, ri) => {
    const padded = r.concat(new Array(Math.max(0, maxCols - r.length)).fill(''));
    return new TableRow({
      children: padded.map(cellText => new TableCell({
        width: { size: colWidth, type: WidthType.DXA },
        shading: ri === 0 ? { type: ShadingType.CLEAR, fill: 'E8F7EE' } : undefined,
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: cellText, bold: ri === 0, color: BODY_COLOR, size: 20 })] })],
      })),
    });
  });
  return new Table({
    width: { size: CONTENT_WIDTH_DXA, type: WidthType.DXA },
    columnWidths,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
      left: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
      right: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'CCCCCC' },
    },
    rows: trs,
  });
}

// Blocks identical to one already emitted earlier in the SAME category file
// are within-file near-duplicates (e.g. a support-contact callout box
// repeated verbatim at the end of every article in a section) — skip
// re-emitting them per the skill's near-duplicate consolidation rule.
// Cross-file repetition (the same policy appearing in a different category's
// file) is left alone; only within-file repeats are collapsed.
function blockSignature(b) {
  if (b.type === 'paragraph') return b.text.length > 60 ? `p:${b.text}` : null;
  if (b.type === 'list') return `l:${b.items.join('|')}`;
  if (b.type === 'table') return `t:${b.rows.map(r => r.join('|')).join('~')}`;
  return null;
}

function renderBlocks(blocks, children, depth, seen) {
  const indent = depth ? { left: 360 } : undefined;
  for (const b of blocks) {
    // Only dedupe at the top level of an article's own body (depth 0) —
    // boilerplate preambles/footers (a disclaimer, a contact-channels list)
    // live there. Never dedupe inside a nested FAQ answer: two different
    // articles can legitimately share the same eligibility list inside an
    // accordion answer, and skipping it there would orphan its heading.
    const sig = depth === 0 ? blockSignature(b) : null;
    if (sig) {
      if (seen.has(sig)) continue;
      seen.add(sig);
    }
    if (b.type === 'heading') {
      const size = b.level <= 2 ? 26 : b.level === 3 ? 24 : 22;
      children.push(new Paragraph({
        children: [new TextRun({ text: b.text, bold: true, color: BODY_COLOR, size })],
        spacing: { before: 200, after: 100 },
        indent,
      }));
    } else if (b.type === 'paragraph') {
      children.push(new Paragraph({
        children: [new TextRun({ text: b.text, bold: !!b.emphasis, color: BODY_COLOR, size: 21 })],
        spacing: { after: b.emphasis ? 40 : 120 },
        indent,
      }));
    } else if (b.type === 'list') {
      b.items.forEach((item, i) => {
        const prefix = b.ordered ? `${i + 1}. ` : '';
        const text = item.startsWith('    - ') ? item.replace('    - ', '') : item;
        const bullet = item.startsWith('    - ') ? undefined : (b.ordered ? undefined : { level: 0 });
        children.push(new Paragraph({
          children: [new TextRun({ text: (b.ordered && !item.startsWith('    - ') ? prefix : '') + text, color: BODY_COLOR, size: 21 })],
          bullet,
          indent: item.startsWith('    - ') ? { left: (indent?.left || 0) + 360 } : indent,
          spacing: { after: 60 },
        }));
      });
    } else if (b.type === 'table') {
      children.push(contentTable(b.rows));
      children.push(new Paragraph({ text: '', spacing: { after: 120 } }));
    } else if (b.type === 'faq') {
      children.push(new Paragraph({
        children: [new TextRun({ text: `Q: ${b.question}`, bold: true, color: BODY_COLOR, size: 21 })],
        spacing: { before: 140, after: 60 },
        indent,
      }));
      renderBlocks(b.blocks, children, (depth || 0) + 1, seen);
    }
  }
}

function buildCategoryDoc(category, articlesBySubcat) {
  const children = [];
  const seen = new Set(); // within-file near-duplicate tracker, scoped to this one category doc
  children.push(filledBanner(
    `Grab Merchant Help Centre (Malaysia) — ${category}`,
    `Merchant · Source: ${SOURCE_ROOT}`
  ));
  children.push(new Paragraph({ text: '', spacing: { after: 200 } }));

  for (const [subcategory, articles] of Object.entries(articlesBySubcat)) {
    children.push(filledBanner(`${category} › ${subcategory}`, null));
    children.push(new Paragraph({ text: '', spacing: { after: 150 } }));

    for (const art of articles) {
      children.push(new Paragraph({
        children: [new TextRun({ text: art.title, bold: true, color: BODY_COLOR, size: 24 })],
        spacing: { before: 160, after: 40 },
      }));
      children.push(new Paragraph({
        children: [new TextRun({ text: `Source: ${art.url}`, italics: true, color: '707070', size: 16 })],
        spacing: { after: 100 },
      }));
      renderBlocks(art.blocks, children, 0, seen);
      children.push(new Paragraph({ text: '', spacing: { after: 100 } }));
    }
  }

  const doc = new Document({
    sections: [{ properties: { page: { size: { width: 12240, height: 15840 } } }, children }],
  });
  return doc;
}

async function main() {
  const onlyCategory = process.argv[2]; // optional: build just one category for preview

  const byCategory = {};
  for (const it of index) {
    if (!KNOWN_CATEGORIES.includes(it.category)) continue;
    if (onlyCategory && it.category !== onlyCategory) continue;
    const art = JSON.parse(fs.readFileSync(path.join(EXTRACTED_DIR, it.file), 'utf8'));
    byCategory[it.category] = byCategory[it.category] || {};
    byCategory[it.category][it.subcategory] = byCategory[it.category][it.subcategory] || [];
    byCategory[it.category][it.subcategory].push(art);
  }

  for (const category of Object.keys(byCategory)) {
    const doc = buildCategoryDoc(category, byCategory[category]);
    const fname = `Grab_Merchant_MY_HelpCentre_${category.replace(/[^a-zA-Z0-9]+/g, '_')}.docx`;
    const outPath = path.join(DOCS_DIR, fname);
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(outPath, buf);
    console.log('wrote', outPath);
  }
}

main();
