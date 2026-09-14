const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const OUT = '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab';
const PAGES_DIR = path.join(OUT, 'pages');
const MANIFEST = path.join(OUT, 'manifest.json');
const QUEUE = path.join(OUT, 'queue.json');
const BATCH_LIMIT = parseInt(process.argv[2] || '40', 10);

if (!fs.existsSync(PAGES_DIR)) fs.mkdirSync(PAGES_DIR, { recursive: true });

function loadJSON(f, def) {
  if (fs.existsSync(f)) return JSON.parse(fs.readFileSync(f, 'utf8'));
  return def;
}
function saveJSON(f, obj) { fs.writeFileSync(f, JSON.stringify(obj, null, 2)); }

function urlToFilename(url) {
  const u = new URL(url);
  return u.pathname.replace(/^\/+/, '').replace(/\//g, '_') + '.html';
}

const SEED_SECTIONS = [
  ["/merchant/en-my/360004458212-Restaurant-management", "GrabFood"],
  ["/merchant/en-my/20000167-Guides", "GrabFood"],
  ["/merchant/en-my/20000168-Manage-orders", "GrabFood"],
  ["/merchant/en-my/20000169-Manage-menu", "GrabFood"],
  ["/merchant/en-my/20000172-Guides", "GrabMart"],
  ["/merchant/en-my/20000173-Manage-orders", "GrabMart"],
  ["/merchant/en-my/20000174-Manage-catalogue", "GrabMart"],
  ["/merchant/en-my/20000177-Payments", "Payment Services"],
  ["/merchant/en-my/20000178-Payment-issues", "Payment Services"],
  ["/merchant/en-my/20000180-GrabExpress-API", "GrabExpress"],
  ["/merchant/en-my/20000181-Booking-issues", "GrabExpress"],
  ["/merchant/en-my/20000211-Dine-Out-Merchant-Loyalty-Programme", "Dine Out"],
  ["/merchant/en-my/20000182-Dine-Out-Deals", "Dine Out"],
  ["/merchant/en-my/20000183-Scan-to-Order", "Dine Out"],
  ["/merchant/en-my/20000184-Loyalty-Stamp-Card", "Dine Out"],
  ["/merchant/en-my/10773768781465-POS-related-solutions", "Hubbo"],
  ["/merchant/en-my/10700845337753-Business-app-related-solutions", "Hubbo"],
  ["/merchant/en-my/10700831467033-Menu-related-solutions-(Business-App)", "Hubbo"],
  ["/merchant/en-my/11382720581657-Menu-related-solutions-(Merchant-Portal)", "Hubbo"],
  ["/merchant/en-my/11337628034329-Portal-related-solutions", "Hubbo"],
  ["/merchant/en-my/20000201-Diner-management", "Reservations"],
  ["/merchant/en-my/20000206-Marketing-tools-and-campaign-management", "Reservations"],
  ["/merchant/en-my/20000139-Account-management", "Reservations"],
  ["/merchant/en-my/20000140-Reservation-management", "Reservations"],
  ["/merchant/en-my/20000141-Booking-availability-%26-other-settings", "Reservations"],
  ["/merchant/en-my/20000142-Restaurant-listing", "Reservations"],
  ["/merchant/en-my/20000143-Widget-%26-integration", "Reservations"],
  ["/merchant/en-my/22972582715801-Marketing-Manager", "Marketing"],
  ["/merchant/en-my/20000188-Business-Growth-Services", "Marketing"],
  ["/merchant/en-my/20000189-Marketing-Manager", "Marketing"],
  ["/merchant/en-my/20000190-Marketing-Manager-Pro", "Marketing"],
  ["/merchant/en-my/20000192-Grab-Biz-Financing-i", "Financing"],
  ["/merchant/en-my/20000116-GXBank-for-Merchants", "GXBank"],
];

async function main() {
  let manifest = loadJSON(MANIFEST, {});
  let queue = loadJSON(QUEUE, null);
  if (queue === null) {
    queue = SEED_SECTIONS.map(([p, cat]) => ({ url: 'https://help.grab.com' + p, category: cat, kind: 'section' }));
  }

  const browser = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    headless: true,
    proxy: { server: 'http://127.0.0.1:40235' },
    args: [
      '--ignore-certificate-errors',
      '--ssl-version-max=tls1.2',
      '--disable-features=PostQuantumKyber,EncryptedClientHello',
    ],
  });
  const page = await browser.newPage();

  let processed = 0;
  let cookieHandled = false;

  while (queue.length > 0 && processed < BATCH_LIMIT) {
    const item = queue.shift();
    const key = item.url;
    if (manifest[key]) continue; // already done

    try {
      await page.goto(item.url, { waitUntil: 'networkidle', timeout: 45000 });
      await page.waitForTimeout(1000);
      if (!cookieHandled) {
        try { await page.locator('#onetrust-accept-btn-handler').click({ timeout: 2000 }); } catch (e) {}
        try { await page.locator('[data-cy="pri-button"]').first().click({ timeout: 2000 }); } catch (e) {}
        cookieHandled = true;
        await page.waitForTimeout(300);
      }

      const finalUrl = page.url();
      const pageType = await page.evaluate(() => {
        if (document.querySelector('[data-cy="article-page-desktop"]')) return 'article';
        if (document.querySelector('[data-cy="section-page-desktop"]')) return 'section';
        return 'unknown';
      });

      const html = await page.content();
      const fname = urlToFilename(finalUrl);
      fs.writeFileSync(path.join(PAGES_DIR, fname), html);

      let newLinks = [];
      if (pageType === 'section') {
        newLinks = await page.evaluate(() => {
          const sidenav = document.querySelector('[data-cy="side-navigation-panel"]');
          const anchors = Array.from(document.querySelectorAll('a[href]'));
          const results = [];
          for (const a of anchors) {
            if (sidenav && sidenav.contains(a)) continue;
            const href = a.getAttribute('href');
            if (href && href.startsWith('/merchant/en-my/')) {
              results.push({ href, text: a.innerText.trim() });
            }
          }
          return results;
        });
      }

      manifest[key] = {
        requestedUrl: item.url,
        finalUrl,
        category: item.category,
        kind: pageType,
        file: fname,
        title: await page.title(),
      };
      console.log(`[${processed + 1}/${BATCH_LIMIT}] ${pageType.padEnd(8)} ${finalUrl}`);

      for (const l of newLinks) {
        const abs = new URL(l.href, 'https://help.grab.com').toString().split('#')[0];
        if (!manifest[abs] && !queue.find(q => q.url === abs)) {
          queue.push({ url: abs, category: item.category, kind: 'unknown' });
        }
      }

      processed++;
    } catch (e) {
      console.log('ERROR on', item.url, e.message.split('\n')[0]);
      manifest[key] = { requestedUrl: item.url, error: e.message.split('\n')[0], category: item.category };
      processed++;
    }
  }

  saveJSON(MANIFEST, manifest);
  saveJSON(QUEUE, queue);
  console.log(`\nBatch done. Processed ${processed}. Queue remaining: ${queue.length}. Manifest size: ${Object.keys(manifest).length}`);

  await browser.close();
}

main();
