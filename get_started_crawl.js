const { chromium } = require('playwright');
const fs = require('fs');
const OUT = '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab';
const PAGES_DIR = `${OUT}/pages`;

const ARTICLE_URLS = [
  'https://help.grab.com/merchant/en-my/40002285-Understanding-E-Stamping-for-Grab-Agreements',
  'https://help.grab.com/merchant/en-my/360027631692-I%E2%80%99m-interested-in-partnering-with-Grab',
  'https://help.grab.com/merchant/en-my/360027910491-Request-to-register-new-or-additional-outlet',
  "https://help.grab.com/merchant/en-my/360027910651-What-is-my-outlet's-registration-status",
  'https://help.grab.com/merchant/en-my/40001106-How-we-verify-business-locations',
  'https://help.grab.com/merchant/en-my/33152757608857-Self-serve-onboarding-for-merchants',
  'https://help.grab.com/merchant/en-my/40001210-Submission-of-Certification-for-Consumer-Protection-(Electronic-Trade-Transactions)-Regulations-2024',
];

(async () => {
  const browser = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    headless: true,
    proxy: { server: 'http://127.0.0.1:46013' },
    args: ['--ignore-certificate-errors', '--ssl-version-max=tls1.2', '--disable-features=PostQuantumKyber,EncryptedClientHello'],
  });
  const page = await browser.newPage();
  const manifest = {};
  let cookieHandled = false;

  for (const url of ARTICLE_URLS) {
    try {
      await page.goto(url, { waitUntil: 'networkidle', timeout: 45000 });
      await page.waitForTimeout(800);
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
      const fname = new URL(finalUrl).pathname.replace(/^\/+/, '').replace(/\//g, '_') + '.html';
      fs.writeFileSync(`${PAGES_DIR}/${fname}`, html);
      manifest[url] = { requestedUrl: url, finalUrl, category: 'Get Started', kind: pageType, file: fname };
      console.log('crawled', pageType, finalUrl);
    } catch (e) {
      console.log('ERROR', url, e.message.split('\n')[0]);
      manifest[url] = { requestedUrl: url, error: e.message.split('\n')[0], category: 'Get Started' };
    }
  }
  fs.writeFileSync(`${OUT}/get_started_manifest.json`, JSON.stringify(manifest, null, 2));
  console.log('done, manifest size:', Object.keys(manifest).length);

  await browser.close();
})();
