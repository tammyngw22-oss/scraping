const mockttp = require('mockttp');
const { chromium } = require('playwright');
const fs = require('fs');

const URL = 'https://www.singtel.com/personal/products-services/mobile/roaming/faqs';
const EXPECTED_SECTIONS = [
  'Before You Travel',
  'When You Are Overseas',
  'Back From Your Trip',
  'Getting Started with Roaming',
  'ReadyRoam Plans',
  'DataRoam Unlimited Daily',
  'AutoReadyRoam (ARR)',
  'DataRoam Monthly Recurring',
  'SMSVoiceRoam 7 Days',
  'Inflight DataRoam (24-hour Unlimited)',
  'Other questions when planning a trip',
  'Usage',
  'Charges',
  'Troubleshoot',
  'Bill Enquiry',
];

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

  // Expand every accordion (top-level and nested) so any lazily-rendered content loads
  const expandCount = await page.evaluate(async () => {
    function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
    let clicked = 0;
    // Run multiple passes since expanding one section can reveal nested accordions
    for (let pass = 0; pass < 4; pass++) {
      const headers = Array.from(document.querySelectorAll('.ux-addon-accordion > .section-header'));
      for (const h of headers) {
        const acc = h.parentElement;
        const main = acc.querySelector(':scope > .main-content');
        const hidden = main && (main.querySelector(':scope > .hide') || main.classList.contains('hide'));
        if (hidden) {
          h.click();
          clicked++;
          await sleep(80);
        }
      }
      await sleep(300);
    }
    return clicked;
  });
  console.log('accordion headers clicked:', expandCount);
  await page.waitForTimeout(1500);

  // Extract sections + Q&A from the DOM in document order
  const sections = await page.evaluate(() => {
    function textOf(el) {
      if (!el) return '';
      return el.innerText.replace(/ /g, ' ').replace(/[ \t]+\n/g, '\n').trim();
    }

    const sectionHeaders = Array.from(document.querySelectorAll('h2.section-header, h2.main-header'));
    const results = [];

    for (let i = 0; i < sectionHeaders.length; i++) {
      const h2 = sectionHeaders[i];
      const sectionName = h2.textContent.trim();
      const nextH2 = sectionHeaders[i + 1];

      // Collect all nodes between this h2 and the next h2 in document order
      const range = document.createRange();
      range.setStartAfter(h2);
      if (nextH2) {
        range.setEndBefore(nextH2);
      } else {
        range.setEndAfter(document.body.lastChild);
      }

      const container = range.cloneContents();
      const accordions = container.querySelectorAll('.ux-addon-accordion');

      const qa = [];
      accordions.forEach(acc => {
        const qEl = acc.querySelector(':scope > .section-header');
        const aEl = acc.querySelector(':scope > .main-content');
        const question = textOf(qEl);
        const answer = textOf(aEl);
        if (question) {
          qa.push({ question, answer });
        }
      });

      results.push({ section: sectionName, qa });
    }

    return results;
  });

  const output = {
    source_url: URL,
    scraped_at: new Date().toISOString(),
    sections,
  };

  fs.writeFileSync('singtel_live_roaming_faq.json', JSON.stringify(output, null, 2));

  const NAV_ONLY_SECTIONS = new Set(['Before You Travel', 'When You Are Overseas', 'Back From Your Trip']);
  const totalQa = sections.reduce((sum, s) => sum + s.qa.length, 0);
  const emptySections = sections.filter(s => s.qa.length === 0 && !NAV_ONLY_SECTIONS.has(s.section)).map(s => s.section);

  console.log('\n--- Summary ---');
  console.log('Sections found:', sections.length);
  console.log('Total Q&A pairs extracted:', totalQa);
  console.log('Unexpectedly empty sections:', emptySections.length ? emptySections.join(', ') : 'none');
  sections.forEach(s => {
    const note = NAV_ONLY_SECTIONS.has(s.section) ? ' (nav/TOC section, no accordions expected)' : '';
    console.log(`  - ${s.section}: ${s.qa.length} Q&A${note}`);
  });

  await browser.close();
  await mitm.stop();
})().catch(err => {
  console.error('FATAL', err);
  process.exit(1);
});
