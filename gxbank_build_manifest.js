const fs = require('fs');
const path = require('path');
const cheerio = require('cheerio');

const OUT = '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab/gxbank_ext';
const TOPICS = [
  { file: 'topic_1.html', name: 'Business Deposit' },
  { file: 'topic_2.html', name: 'Account Opening' },
  { file: 'topic_3.html', name: 'Payments & Transfers' },
  { file: 'topic_4.html', name: 'Business Loan' },
  { file: 'topic_5.html', name: 'App Security & Settings' },
  { file: 'topic_6.html', name: 'Security & Privacy' },
];

const manifest = [];
for (const t of TOPICS) {
  const html = fs.readFileSync(path.join('gxbank_ext', t.file), 'utf8');
  const $ = cheerio.load(html);
  $('.mt-sortable-listing').each((i, groupEl) => {
    const groupName = $(groupEl).find('.mt-listing-detailed-title > a').first().text().trim();
    $(groupEl).find('a.mt-listing-detailed-subpage-title').each((j, a) => {
      const url = $(a).attr('href');
      const title = $(a).text().trim();
      manifest.push({ topic: t.name, group: groupName, title, url });
    });
  });
}

fs.writeFileSync('gxbank_ext/manifest.json', JSON.stringify(manifest, null, 2));
console.log('total leaf articles:', manifest.length);
const byTopic = {};
for (const m of manifest) byTopic[m.topic] = (byTopic[m.topic] || 0) + 1;
console.log(byTopic);
