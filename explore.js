const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    headless: true,
  });
  const page = await browser.newPage();
  await page.goto('https://www.singtel.com/personal/products-services/mobile/roaming/faqs', { waitUntil: 'networkidle', timeout: 60000 });
  await page.waitForTimeout(3000);

  // Dump some structure to understand accordion markup
  const html = await page.content();
  require('fs').writeFileSync('/tmp/claude-0/-home-user-scraping/5bdc95b8-e473-5ddb-8499-a1952d7637a0/scratchpad/page.html', html);
  console.log('saved html length', html.length);

  await browser.close();
})();
