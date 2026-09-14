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

// cheerio's .text() walks every descendant text node, including the raw
// CSS/JS inside <style>/<script> — a browser's rendered text skips those,
// cheerio does not. Source pages here sometimes nest a stray <style> tag
// directly inside a <p> (malformed but real), so strip script/style before
// reading text anywhere, not just in one fallback branch.
// Also: a <br> carries no text of its own, so plain .text() glues its
// neighbors together with zero separator (e.g. "<strong>Ads</strong><br>On
// the..." becomes "AdsOn the...") — replace every <br> with a space first.
function safeText($, node) {
  return $(node).clone().find('br').replaceWith(' ').end().find('script, style').remove().end().text();
}

// A <ul>/<table> is "external nav chrome" (a link directory pointing off-site,
// not real content) when most of its text comes from <a> tags whose href
// points outside help.grab.com. These showed up verbatim, repeated across
// nearly every article in a section (e.g. a GXBank article-to-article
// "helpful links to help.gxbank.my" block) and add no unique content.
function isExternalNavBlock($, node) {
  const anchors = $(node).find('a[href]');
  if (anchors.length === 0) return false;
  let linkTextLen = 0;
  let externalCount = 0;
  anchors.each((i, a) => {
    const href = $(a).attr('href') || '';
    const text = cleanText($(a).text());
    linkTextLen += text.length;
    if (/^https?:\/\//i.test(href) && !href.includes('help.grab.com')) externalCount++;
  });
  const totalLen = cleanText(safeText($, node)).length || 1;
  // Either most of the block's text is link text (a plain link list), or it
  // links out to the same external site several times (a table whose other
  // columns are plain descriptive text around those links, e.g. a directory
  // of an unrelated site's FAQ sections) — both patterns are nav chrome.
  return externalCount >= 3 || (externalCount >= 1 && (linkTextLen / totalLen) > 0.6);
}

// Recursively walk a cheerio element's children, producing block objects.
function walk($, el, blocks) {
  const node = el;
  if (node.type === 'text') return;
  const tag = (node.tagName || '').toLowerCase();

  if (['script', 'style', 'button', 'input', 'label', 'textarea', 'form', 'iframe'].includes(tag)) return;

  if (['h1', 'h2', 'h3', 'h4', 'h5'].includes(tag)) {
    const text = cleanText(safeText($, node));
    if (text) blocks.push({ type: 'heading', level: parseInt(tag[1], 10), text });
    return;
  }
  if (tag === 'p') {
    const text = cleanText(safeText($, node));
    if (text) blocks.push({ type: 'paragraph', text });
    return;
  }
  if (tag === 'ul' || tag === 'ol') {
    if (isExternalNavBlock($, node)) return; // drop off-site link directories
    const items = [];
    $(node).children('li').each((i, li) => {
      const text = cleanText(safeText($, $(li).clone().children('ul,ol').remove().end()));
      if (text) items.push(text);
      // nested lists inside li: flatten as sub-items with indent marker
      $(li).children('ul,ol').each((j, nested) => {
        $(nested).children('li').each((k, nli) => {
          const ntext = cleanText(safeText($, nli));
          if (ntext) items.push('    - ' + ntext);
        });
      });
    });
    if (items.length) blocks.push({ type: 'list', ordered: tag === 'ol', items });
    return;
  }
  if (tag === 'table') {
    if (isExternalNavBlock($, node)) return; // drop off-site link directories
    const rows = [];
    $(node).find('tr').each((i, tr) => {
      const cells = [];
      $(tr).find('th,td').each((j, cell) => {
        // A cell can contain a nested list (e.g. several bullet points under
        // one topic) — join those as "; "-separated text instead of a
        // run-on blob so the cell stays readable.
        const liTexts = $(cell).find('li').map((k, li) => cleanText(safeText($, li))).get();
        cells.push(liTexts.length ? liTexts.join('; ') : cleanText(safeText($, cell)));
      });
      if (cells.length) rows.push(cells);
    });
    // A table only has a real header row when the source actually marks one
    // (a <th> cell, or a <thead> wrapper) — some source tables are pure
    // term/definition pairs with every row in plain <td>s (e.g. an order
    // status glossary), and treating that first data row as a header would
    // wrongly prefix it onto every other row.
    const hasHeader = $(node).find('th').length > 0 || $(node).find('thead').length > 0;
    if (rows.length) blocks.push({ type: 'table', rows, hasHeader });
    return;
  }
  if (tag === 'details') {
    const summary = cleanText(safeText($, $(node).children('summary').first()));
    const innerBlocks = [];
    $(node).contents().each((i, child) => {
      if ((child.tagName || '').toLowerCase() === 'summary') return;
      walk($, child, innerBlocks);
    });
    // Skip FAQ entries whose answer is image-only / empty (no extractable
    // text at all) — a bare question with nothing under it isn't useful,
    // per the skill's image-only-content rule.
    if (innerBlocks.length === 0) return;
    blocks.push({ type: 'faq', question: summary, blocks: innerBlocks });
    return;
  }
  if (tag === 'blockquote') {
    const text = cleanText(safeText($, node));
    if (text) blocks.push({ type: 'paragraph', text: text });
    return;
  }
  if (tag === 'img' || tag === 'br' || tag === 'hr' || tag === 'svg') {
    return;
  }
  // Generic container (div, span, strong, a, etc.) not handled above. If it
  // has no block-level descendant anywhere inside, it's just mixed inline
  // content (plain text alongside <strong>/<em>/<a> spans, e.g. a callout
  // like "<strong>Important:</strong> some sentence <strong>with a
  // bolded phrase</strong> in the middle.") that isn't wrapped in a <p> —
  // take its whole text as one paragraph instead of recursing, which would
  // silently drop the bare text nodes between the inline tags and keep only
  // the bold fragments. Only descend into children when a real block
  // element (p/list/table/details/heading) is nested inside, so that
  // structure is preserved.
  const hasBlockDescendant = $(node).find('p, ul, ol, table, details, blockquote, h1, h2, h3, h4, h5').length > 0;
  if (!hasBlockDescendant) {
    const text = cleanText(safeText($, node));
    if (text) blocks.push({ type: 'paragraph', text, emphasis: tag === 'strong' || tag === 'b' });
    return;
  }
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
