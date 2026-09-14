// Extracts + flattens all 245 GXBank help.gxbank.my articles into
// AI-agent-friendly Q&A items, grouped by topic > group, matching the same
// content rules as the main pipeline (no tables, contact info stripped).
const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');

const ARTICLES_DIR = 'gxbank_ext/articles';
const manifest = JSON.parse(fs.readFileSync('gxbank_ext/manifest.json', 'utf8'));

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
function cleanText(s) {
  return (s || '').replace(/\s+/g, ' ').trim();
}

function walk($, el, blocks) {
  const node = el;
  if (node.type === 'text') {
    // Some MindTouch articles leave the actual answer as a bare text node
    // directly under the content container (no <p> wrapper at all, e.g.
    // "<style></style>This facility is conventional.<footer>...") — capture
    // real text instead of silently dropping it.
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

// { topic: { group: [ {q, a, url} ] } }
const byTopic = {};
let itemCount = 0, emptyCount = 0;

for (let i = 0; i < manifest.length; i++) {
  const m = manifest[i];
  const htmlPath = path.join(ARTICLES_DIR, `${i}.html`);
  if (!fs.existsSync(htmlPath)) continue;
  const html = fs.readFileSync(htmlPath, 'utf8');
  const $ = cheerio.load(html);

  const title = cleanText($('h1#title').first().text()) || m.title;
  const contentEl = $('section.mt-content-container').first();
  const blocks = [];
  // Skip the page-summary/thumbnail wrapper and the subpage-listing widget
  // (only present on category pages, harmless to exclude defensively here).
  contentEl.contents().each((i, child) => {
    const cls = (child.attribs && child.attribs.class) || '';
    if (cls.includes('mt-page-summary') || cls.includes('mt-subpage-listings') || cls.includes('mt-category-container')) return;
    walk($, child, blocks);
  });

  const lines = mergeLabelLines(blocks.flatMap(blockToLines));
  if (lines.length === 0) { emptyCount++; continue; }

  byTopic[m.topic] = byTopic[m.topic] || {};
  byTopic[m.topic][m.group] = byTopic[m.topic][m.group] || [];
  byTopic[m.topic][m.group].push({ q: title, a: lines, url: m.url });
  itemCount++;
}

fs.writeFileSync('gxbank_ext/qa_by_topic.json', JSON.stringify(byTopic, null, 2));
console.log('total Q&A items:', itemCount, 'empty/skipped:', emptyCount);
for (const [topic, groups] of Object.entries(byTopic)) {
  const n = Object.values(groups).reduce((a, arr) => a + arr.length, 0);
  console.log(`  ${topic}: ${n} items across ${Object.keys(groups).length} groups`);
}
