// Converts the rich extracted block-structure per article into a flat list
// of {category, subcategory, q, a} Q&A units, AI-agent-friendly:
//   - one Q&A per article's own top-level content, plus one Q&A per nested
//     FAQ-accordion question (each self-contained, no cross-references)
//   - tables/lists/headings flattened into plain prose lines (no table format)
//   - contact channels (phone numbers, emails, live-chat/hotline mentions)
//     stripped, since the agent should resolve things in-conversation
//   - blocks that are ENTIRELY escalation content are dropped outright
const fs = require('fs');
const path = require('path');

const OUT = '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab';
const EXTRACTED_DIR = path.join(OUT, 'extracted');
const QA_DIR = path.join(OUT, 'qa_items');
if (!fs.existsSync(QA_DIR)) fs.mkdirSync(QA_DIR, { recursive: true });

const EMAIL_RE = /[\w.+-]+@[\w-]+\.[\w.-]+/i;
const PHONE_RE = /\+\d[\d\s-]{6,}\d/;
const DROP_KEYWORDS = [
  'hotline', 'live chat', 'customer support team', 'contact them through',
  'reach out to our', 'reach out to us', 'contact us', 'email:', 'call us', 'whatsapp',
  "we're here to help", 'feel free to reach out',
  'grab support team via', 'contact grab support', '24/7 support team',
  'chat with gx buddy', 'contact gx bank', "contact gx bank's",
  'helpful link', 'need more information on', 'frequently asked questions.',
];

function isContactLine(text) {
  // Source text uses curly quotes/apostrophes (’) — normalize to straight
  // ones so keyword matching (written with straight ') doesn't silently
  // miss real matches like "We're here to help!".
  const lower = text.toLowerCase().replace(/[‘’]/g, "'");
  if (EMAIL_RE.test(text)) return true;
  if (PHONE_RE.test(text)) return true;
  return DROP_KEYWORDS.some(k => lower.includes(k));
}

// A line made of nothing but emoji/punctuation/whitespace (a stray "⚠️" or
// "📝" span left over from a callout box) — has no letters or digits at all.
function isEmojiOnly(text) {
  return !/[a-zA-Z0-9]/.test(text);
}

// Strip only the offending sentence(s) out of a paragraph rather than
// dropping the whole thing — a disclaimer paragraph often mixes one genuinely
// useful sentence with one escalation sentence ("please contact support").
function stripContactSentences(text) {
  const sentences = text.split(/(?<=[.!?])\s+/);
  const kept = sentences.filter(s => !isContactLine(s));
  // Some source list items merge a question and its answer into one run-on
  // string with no separator (e.g. "...too much? Yes, discuss with
  // GrabFinance Collections Hotline: +603..."). If the only thing left after
  // stripping is a bare question with nothing resolving it — the whole
  // "answer" was pure escalation — that's as unusable as an empty answer,
  // so drop it entirely rather than leave a dangling unanswered question.
  if (kept.length < sentences.length && kept.length > 0 && /\?\s*$/.test(kept[kept.length - 1])) {
    return '';
  }
  return kept.join(' ').trim();
}

// "Learn more here." / "click here" style cross-references pointed at a link
// we've already stripped out of the content — dead references with nothing
// for the agent to act on, per the skill's no-dangling-cross-reference rule.
const DEAD_LINK_RE = /\b(learn more|find out more|click)\s*(about\s+\S+\s+)?here\b\.?/gi;
function stripDeadLinkPhrases(text) {
  let t = text.replace(DEAD_LINK_RE, '');
  // A dangling "...submit it here." at the very end of the text was a link
  // to a form/page we've already stripped out — drop just that trailing
  // word rather than leave an unresolvable reference (only at the end, so
  // legitimate mid-sentence uses of "here" are left alone).
  t = t.replace(/\s+here\.?\s*$/i, '.');
  return t.replace(/\s{2,}/g, ' ').trim();
}

// Render one block (and its children) as plain prose lines. Returns an array
// of strings (each will become one line in the answer), with contact-info
// lines/items removed.
function blockToLines(b) {
  const lines = [];
  if (b.type === 'heading') {
    const text = stripDeadLinkPhrases(stripContactSentences(b.text));
    if (text && !isEmojiOnly(text)) lines.push(text.endsWith(':') ? text : text + ':');
  } else if (b.type === 'paragraph') {
    const text = stripDeadLinkPhrases(stripContactSentences(b.text));
    if (text && !isEmojiOnly(text)) lines.push(text);
  } else if (b.type === 'list') {
    const kept = b.items
      .map(i => i.startsWith('    - ') ? i.replace('    - ', '') : i)
      .map(i => stripDeadLinkPhrases(stripContactSentences(i)))
      .filter(i => i && !isEmojiOnly(i));
    kept.forEach(item => lines.push(`- ${item}`));
  } else if (b.type === 'table') {
    if (!b.rows.length) return lines;
    const header = b.rows[0];
    const dataRows = b.rows.slice(1).length ? b.rows.slice(1) : [];
    if (dataRows.length === 0) {
      // No header/data split detected — just join the single row.
      const rowText = stripContactSentences(header.join(', '));
      if (rowText) lines.push(rowText);
    } else {
      for (const row of dataRows) {
        const parts = row.map((cell, i) => header[i] ? `${header[i]}: ${cell}` : cell);
        const rowText = stripContactSentences(parts.join(', '));
        if (rowText) lines.push(rowText);
      }
    }
  }
  return lines;
}

// Fold a short label line ("Important Note:") into the sentence that
// follows it, so the answer reads as natural prose rather than a
// label floating above unrelated text.
function mergeLabelLines(lines) {
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    const isShortLabel = line.endsWith(':') && line.length < 40 && !line.startsWith('-');
    const nextIsAlsoLabel = i + 1 < lines.length && lines[i + 1].endsWith(':') && lines[i + 1].length < 60;
    if (isShortLabel && i + 1 < lines.length && !lines[i + 1].startsWith('-') && !nextIsAlsoLabel) {
      out.push(`${line} ${lines[i + 1]}`);
      i++;
    } else {
      out.push(line);
    }
  }
  return out;
}

// A pure table-of-contents ("Section 1: Introduction:", "Section 2: ...:")
// with no real sentence anywhere isn't a usable FAQ answer on its own —
// it's just headers pointing at content that lives in the split-out child
// items. Suppress it in that case rather than emit an empty-feeling Q&A.
function isOutlineOnly(lines) {
  return lines.length > 0 && lines.every(l => l.endsWith(':') && !l.startsWith('-'));
}

// Splits a block list into segments at each heading boundary, so a long
// guide that uses H2/H3 sub-headings to organize several distinct topics
// (e.g. "Who can use CRM" / "Create a campaign" / "Tracking performance")
// becomes several focused Q&As instead of one giant answer covering all of
// them — a RAG query about one sub-topic should retrieve just that part,
// not a wall of text where the specific answer is buried among nine others.
function splitByHeading(blocks) {
  const segments = [{ heading: null, blocks: [] }];
  for (const b of blocks) {
    if (b.type === 'heading') {
      segments.push({ heading: b.text, blocks: [] });
    } else {
      segments[segments.length - 1].blocks.push(b);
    }
  }
  return segments.filter(s => s.heading || s.blocks.length > 0);
}

// Recursively walks a block list, pushing one Q&A item per heading-split
// segment of THIS level's own content (question, or "question: heading"
// when segmented) plus one item per nested faq found at ANY depth — a faq
// inside a faq (e.g. "Setting up X" > "From the Y page" > actual steps)
// still becomes its own item instead of being silently dropped.
function collectItems(blocks, url, question, items) {
  const faqBlocks = [];
  const nonFaqBlocks = [];
  for (const b of blocks) {
    if (b.type === 'faq' && !isContactLine(b.question)) faqBlocks.push(b);
    else if (b.type !== 'faq') nonFaqBlocks.push(b);
  }

  const ownItems = [];
  for (const seg of splitByHeading(nonFaqBlocks)) {
    const lines = mergeLabelLines(seg.blocks.flatMap(blockToLines));
    if (lines.length === 0) continue;
    const q = seg.heading ? `${question}: ${seg.heading}` : question;
    ownItems.push({ q, a: lines, url });
  }
  // A single outline-only own-segment ("About X" -> just section labels)
  // adds nothing once the real content already exists as split-out items.
  const suppressOwn = ownItems.length === 1 && isOutlineOnly(ownItems[0].a) && faqBlocks.length > 0;
  // This level's own intro content (e.g. "How to make payments with debit
  // cards" leading into its accordions) reads first in the source, so push
  // it before recursing into the nested faqs — otherwise a parent's intro
  // ends up listed after all its own children, which reads backwards.
  if (!suppressOwn) items.push(...ownItems);
  for (const fb of faqBlocks) {
    collectItems(fb.blocks, url, fb.question, items);
  }
}

const KNOWN_CATEGORIES = ['GrabFood', 'GrabMart', 'Payment Services', 'GrabExpress', 'Dine Out', 'Hubbo', 'Reservations', 'Marketing', 'Financing', 'GXBank'];

let totalItems = 0;
const byCategory = {};

for (const file of fs.readdirSync(EXTRACTED_DIR)) {
  if (!file.endsWith('.json')) continue;
  const article = JSON.parse(fs.readFileSync(path.join(EXTRACTED_DIR, file), 'utf8'));
  const { category, subcategory, title, url } = article;
  if (!KNOWN_CATEGORIES.includes(category)) continue;

  byCategory[category] = byCategory[category] || {};
  byCategory[category][subcategory] = byCategory[category][subcategory] || [];

  const items = [];
  collectItems(article.blocks, url, title, items);

  for (const it of items) {
    byCategory[category][subcategory].push(it);
    totalItems++;
  }
}

for (const [category, subcats] of Object.entries(byCategory)) {
  fs.writeFileSync(path.join(QA_DIR, `${category.replace(/[^a-zA-Z0-9]+/g, '_')}.json`), JSON.stringify(subcats, null, 2));
}

console.log('total Q&A items:', totalItems);
for (const [category, subcats] of Object.entries(byCategory)) {
  const n = Object.values(subcats).reduce((a, arr) => a + arr.length, 0);
  console.log(`  ${category}: ${n} items across ${Object.keys(subcats).length} subcategories`);
}
