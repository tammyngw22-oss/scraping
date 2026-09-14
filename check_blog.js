const { chromium } = require('playwright');
const OUT = '/tmp/claude-0/-home-user-scraping/3f7cc647-13c3-5879-aa12-73bfeb00a642/scratchpad/grab';

(async () => {
  const browser = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    headless: true,
    proxy: { server: 'http://127.0.0.1:46013' },
    args: ['--ignore-certificate-errors', '--ssl-version-max=tls1.2', '--disable-features=PostQuantumKyber,EncryptedClientHello'],
  });
  const page = await browser.newPage();
  await page.goto('https://merchant.grab.com/en-my?grab_external=true', { waitUntil: 'networkidle', timeout: 45000 });
  await page.waitForTimeout(2000);
  console.log('Final URL:', page.url());
  console.log('Title:', await page.title());
  await page.screenshot({ path: `${OUT}/blog_page.png`, fullPage: true });
  const html = await page.content();
  require('fs').writeFileSync(`${OUT}/blog_page.html`, html);
  console.log('html length', html.length);
  await browser.close();
})();
