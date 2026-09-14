const fs = require('fs');
const path = require('path');

const manifest = JSON.parse(fs.readFileSync('gxbank_ext/manifest.json', 'utf8'));
const ARTICLES_DIR = 'gxbank_ext/articles';
if (!fs.existsSync(ARTICLES_DIR)) fs.mkdirSync(ARTICLES_DIR, { recursive: true });

const CONCURRENCY = 8;

async function fetchOne(i, m) {
  const outPath = path.join(ARTICLES_DIR, `${i}.html`);
  if (fs.existsSync(outPath) && fs.statSync(outPath).size > 500) return { i, ok: true, cached: true };
  try {
    const res = await fetch(m.url, { redirect: 'follow' });
    const text = await res.text();
    fs.writeFileSync(outPath, text);
    return { i, ok: res.ok, status: res.status };
  } catch (e) {
    return { i, ok: false, error: e.message };
  }
}

async function main() {
  let idx = 0;
  let done = 0;
  const errors = [];
  async function worker() {
    while (idx < manifest.length) {
      const i = idx++;
      const r = await fetchOne(i, manifest[i]);
      done++;
      if (!r.ok) errors.push({ i, url: manifest[i].url, ...r });
      if (done % 20 === 0) console.log(`progress: ${done}/${manifest.length}`);
    }
  }
  await Promise.all(new Array(CONCURRENCY).fill(0).map(worker));
  console.log('total done:', done, 'errors:', errors.length);
  if (errors.length) fs.writeFileSync('gxbank_ext/fetch_errors.json', JSON.stringify(errors, null, 2));
}

main();
