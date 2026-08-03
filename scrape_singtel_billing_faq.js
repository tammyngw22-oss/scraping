const mockttp = require('mockttp');
const { chromium } = require('playwright');
const fs = require('fs');

const URL = 'https://www.singtel.com/personal/support/account-billing';

(async () => {
  const https = await mockttp.generateCACertificate();
  const mitm = mockttp.getLocal({ https });
  mitm.forUnmatchedRequest().thenPassThrough({
    proxyConfig: { proxyUrl: 'http://127.0.0.1:34765' },
  });
  await mitm.start(8765);

  const browser = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    headless: true,
    proxy: { server: 'http://127.0.0.1:8765' },
    args: ['--ignore-certificate-errors'],
  });

  const page = await browser.newPage();
  await page.goto(URL, { waitUntil: 'networkidle', timeout: 60000 });
  await page.waitForTimeout(2000);

  // Dismiss cookie-consent banner if present
  const cookieSelectors = [
    'button#onetrust-accept-btn-handler',
    'button[aria-label="Accept"]',
    '.cookie-consent button',
    'button:has-text("Accept")',
  ];
  for (const sel of cookieSelectors) {
    try {
      const el = await page.$(sel);
      if (el) {
        await el.click({ timeout: 3000 });
        console.log('dismissed cookie banner via', sel);
        await page.waitForTimeout(500);
        break;
      }
    } catch (e) {
      // ignore
    }
  }

  // Expand every FAQ accordion item so any lazily-rendered content loads
  const expandCount = await page.evaluate(async () => {
    function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
    let clicked = 0;
    for (let pass = 0; pass < 3; pass++) {
      const links = Array.from(document.querySelectorAll('.faq-que-link'));
      for (const a of links) {
        a.click();
        clicked++;
        await sleep(50);
      }
      await sleep(300);
    }
    return clicked;
  });
  console.log('faq items clicked:', expandCount);
  await page.waitForTimeout(1000);

  // Extract subsections + Q&A in document order
  const sections = await page.evaluate(() => {
    function textOf(el) {
      if (!el) return '';
      return el.innerText.replace(/ /g, ' ').replace(/[ \t]+\n/g, '\n').trim();
    }

    const subsectionHeaders = Array.from(document.querySelectorAll('h2.subsection-header'));
    const results = [];

    subsectionHeaders.forEach(h2 => {
      const sectionName = h2.textContent.replace(/ /g, ' ').trim();
      const container = h2.closest('.ux-faq-support') || h2.parentElement;
      const items = Array.from(container.querySelectorAll('li.accordion-navigation'));

      const qa = items.map(li => {
        const qEl = li.querySelector('.faq-que-link');
        const aEl = li.querySelector('.faq-ans-desc');
        const question = textOf(qEl);
        let answer = textOf(aEl);
        // If the answer is just a "Get started" link with no other text, capture the link's URL too
        const link = aEl ? aEl.querySelector('a.registrationButton, a') : null;
        if (link && (!answer || /^get started$/i.test(answer))) {
          answer = `${answer || 'Get started'}: ${link.href}`;
        }
        return { question, answer };
      }).filter(qa => qa.question);

      results.push({ section: sectionName, qa });
    });

    return results;
  });

  const output = {
    source_url: URL,
    scraped_at: new Date().toISOString(),
    sections,
  };

  fs.writeFileSync('singtel_billing_faq.json', JSON.stringify(output, null, 2));

  const totalQa = sections.reduce((sum, s) => sum + s.qa.length, 0);
  const emptySections = sections.filter(s => s.qa.length === 0).map(s => s.section);

  console.log('\n--- Summary ---');
  console.log('Sections found:', sections.length);
  console.log('Total Q&A pairs extracted:', totalQa);
  console.log('Empty sections:', emptySections.length ? emptySections.join(', ') : 'none');
  sections.forEach(s => console.log(`  - ${s.section}: ${s.qa.length} Q&A`));

  await browser.close();
  await mitm.stop();
})().catch(err => {
  console.error('FATAL', err);
  process.exit(1);
});
