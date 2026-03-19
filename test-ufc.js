const { chromium } = require('playwright');

async function explore() {
  const browser = await chromium.launch({ headless: false, args: ['--start-maximized'] });
  const context = await browser.newContext({
    viewport: null,
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
  });
  const page = await context.newPage();

  await page.goto('https://www.ufc.com/event/ufc-fight-night-march-21-2026', {
    waitUntil: 'domcontentloaded', timeout: 60000
  });
  await page.waitForTimeout(5000);

  // Extract all fights using data-fmid
  const fights = await page.evaluate(() => {
    return [...document.querySelectorAll('.c-listing-fight[data-fmid]')].map(el => {
      const fightId = el.getAttribute('data-fmid');
      const red  = el.querySelector('.c-listing-fight__corner-name--red')?.innerText.trim();
      const blue = el.querySelector('.c-listing-fight__corner-name--blue')?.innerText.trim();
      const weightClass = el.querySelector('.c-listing-fight__class-text')?.innerText.trim();
      // Get iframe src to find event ID
      const iframe = el.querySelector('iframe');
      const iframeSrc = iframe ? iframe.src : null;
      return { fightId, red, blue, weightClass, iframeSrc };
    });
  });

  console.log('=== All fights ===\n');
  fights.forEach((f, i) => console.log(`${i}: [${f.fightId}] ${f.red} vs ${f.blue} (${f.weightClass})`));

  // Extract event ID from the first iframe src (e.g. /matchup/1301/12623/pre)
  // We need to click expand first to load the iframe, OR derive from previous test (1301)
  // Let's click the first expand button to get the event ID dynamically
  await page.locator('.c-listing-fight__expand-button').first().click();
  await page.waitForTimeout(3000);

  const eventId = await page.evaluate(() => {
    const src = document.querySelector('.details-content__iframe-wrapper iframe')?.src || '';
    const match = src.match(/\/matchup\/(\d+)\//);
    return match ? match[1] : null;
  });
  console.log(`\nEvent ID: ${eventId}`);

  // Now visit each fight's matchup page and read all tabs
  const tabs = ['MATCHUP STATS', 'WIN BY', 'SIGNIFICANT STRIKES', 'GRAPPLING', 'ODDS'];

  for (const fight of fights) {
    const url = `https://www.ufc.com/matchup/${eventId}/${fight.fightId}/pre`;
    console.log(`\n${'='.repeat(60)}`);
    console.log(`FIGHT: ${fight.red} vs ${fight.blue}`);
    console.log(`URL: ${url}`);
    console.log('='.repeat(60));

    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForTimeout(4000);

    for (const tabName of tabs) {
      try {
        const tab = page.locator(`text=${tabName}`).first();
        if (await tab.isVisible({ timeout: 1500 })) {
          await tab.click();
          await page.waitForTimeout(2000);
        } else {
          continue;
        }
      } catch { continue; }

      const text = await page.evaluate(() => document.body.innerText);
      const lines = text.split('\n').map(l => l.trim()).filter(l =>
        l.length > 1 &&
        !l.includes('Cookie Policy') &&
        !l.includes('clicking "OK"') &&
        !l.includes('Reject All') &&
        !l.includes('Cookies Settings') &&
        !l.includes('storing of cookies')
      );

      console.log(`\n  --- ${tabName} ---`);
      lines.forEach((l, i) => console.log(`  ${i}: ${l}`));
    }
  }

  await browser.close();
}

explore().catch(console.error);
