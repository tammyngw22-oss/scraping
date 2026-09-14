const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');

const OUT = '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab';
const PAGES_DIR = path.join(OUT, 'pages');
const EXTRACTED_DIR = path.join(OUT, 'extracted');
if (!fs.existsSync(EXTRACTED_DIR)) fs.mkdirSync(EXTRACTED_DIR, { recursive: true });

const manifest = JSON.parse(fs.readFileSync(path.join(OUT, 'manifest.json'), 'utf8'));

function cleanText(s) {
  return (s || '').replace(/\s+/g, ' ').replace(/ /g, ' ').trim();
}

// Recursively walk a cheerio element's children, producing block objects.
function walk($, el, blocks) {
  const node = el;
  if (node.type === 'text') return;
  const tag = (node.tagName || '').toLowerCase();

  if (['script', 'style', 'button', 'input', 'label', 'textarea', 'form', 'iframe'].includes(tag)) return;

  if (['h1', 'h2', 'h3', 'h4', 'h5'].includes(tag)) {
    const text = cleanText($(node).text());
    if (text) blocks.push({ type: 'heading', level: parseInt(tag[1], 10), text });
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
      // nested lists inside li: flatten as sub-items with indent marker
      $(li).children('ul,ol').each((j, nested) => {
        $(nested).children('li').each((k, nli) => {
          const ntext = cleanText($(nli).text());
          if (ntext) items.push('    - ' + ntext);
        });
      });
    });
    if (items.length) blocks.push({ type: 'list', ordered: tag === 'ol', items });
    return;
  }
  if (tag === 'table') {
    const rows = [];
    $(node).find('tr').each((i, tr) => {
      const cells = [];
      $(tr).find('th,td').each((j, cell) => {
        cells.push(cleanText($(cell).text()));
      });
      if (cells.length) rows.push(cells);
    });
    if (rows.length) blocks.push({ type: 'table', rows });
    return;
  }
  if (tag === 'details') {
    const summary = cleanText($(node).children('summary').first().text());
    const innerBlocks = [];
    $(node).contents().each((i, child) => {
      if ((child.tagName || '').toLowerCase() === 'summary') return;
      walk($, child, innerBlocks);
    });
    blocks.push({ type: 'faq', question: summary, blocks: innerBlocks });
    return;
  }
  if (tag === 'blockquote') {
    const text = cleanText($(node).text());
    if (text) blocks.push({ type: 'paragraph', text: text });
    return;
  }
  if (tag === 'img' || tag === 'br' || tag === 'hr' || tag === 'svg') {
    return;
  }
  if (tag === 'strong' || tag === 'b' || tag === 'em' || tag === 'i') {
    // Bare inline emphasis not wrapped in a <p> (e.g. a standalone callout
    // label like "Important Note:") — emit as its own short bold paragraph
    // instead of silently dropping it.
    const text = cleanText($(node).text());
    if (text) blocks.push({ type: 'paragraph', text, emphasis: true });
    return;
  }
  // Generic container (div, section, span, a, etc.) — recurse into children
  $(node).contents().each((i, child) => walk($, child, blocks));
}

let processed = 0, skipped = 0;
const index = [];

for (const [key, entry] of Object.entries(manifest)) {
  if (entry.kind !== 'article' || !entry.file) { skipped++; continue; }
  const htmlPath = path.join(PAGES_DIR, entry.file);
  if (!fs.existsSync(htmlPath)) { skipped++; continue; }
  const html = fs.readFileSync(htmlPath, 'utf8');
  const $ = cheerio.load(html);

  const breadcrumbItems = [];
  $('[data-cy^="breadcrumb-item-"]').each((i, el) => {
    breadcrumbItems.push(cleanText($(el).text()));
  });
  const category = breadcrumbItems[1] || entry.category || 'Uncategorized';
  const subcategory = breadcrumbItems[2] || '';

  const title = cleanText($('[data-cy="article-title"]').first().text());
  if (!title) { skipped++; continue; }

  const bodyEl = $('[data-cy="article-body"]').first();
  const blocks = [];
  bodyEl.contents().each((i, child) => walk($, child, blocks));

  const record = {
    url: entry.finalUrl,
    title,
    category,
    subcategory,
    blocks,
  };
  const fname = entry.file.replace(/\.html$/, '.json');
  fs.writeFileSync(path.join(EXTRACTED_DIR, fname), JSON.stringify(record, null, 2));
  index.push({ file: fname, url: entry.finalUrl, title, category, subcategory, blockCount: blocks.length });
  processed++;
}

fs.writeFileSync(path.join(OUT, 'extracted_index.json'), JSON.stringify(index, null, 2));
console.log('processed:', processed, 'skipped:', skipped);

// Print category/subcategory breakdown
const byCat = {};
for (const it of index) {
  byCat[it.category] = byCat[it.category] || {};
  byCat[it.category][it.subcategory] = (byCat[it.category][it.subcategory] || 0) + 1;
}
for (const cat of Object.keys(byCat)) {
  console.log(`\n${cat}`);
  for (const [sub, n] of Object.entries(byCat[cat])) {
    console.log(`  ${sub}: ${n}`);
  }
}
