const { chromium } = require('playwright');

const TAPOLOGY_URL = process.argv[2] || 'https://www.tapology.com/fightcenter/events/136856-ufc-fight-night';

async function explore() {
  const browser = await chromium.launch({ headless: false, args: ['--start-maximized'] });
  const context = await browser.newContext({
    viewport: null,
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
  });
  const page = await context.newPage();

  await page.goto(TAPOLOGY_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(3000);

  // Dismiss cookie consent if present
  try {
    const cookieBtn = page.locator('button:has-text("OK"), button:has-text("Accept"), button:has-text("I Agree"), button:has-text("Accept All")').first();
    if (await cookieBtn.isVisible({ timeout: 2000 })) {
      await cookieBtn.click();
      await page.waitForTimeout(1000);
    }
  } catch {}

  // Click Predictions tab
  try {
    const predsTab = page.locator('text=Predictions').first();
    if (await predsTab.isVisible({ timeout: 3000 })) {
      await predsTab.click();
      await page.waitForTimeout(3000);
      console.log('Clicked Predictions tab\n');
    }
  } catch {}

  const text = await page.evaluate(() => document.body.innerText);
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  // Dump everything after the Predictions tab area (last 200 lines)
  console.log('=== Full page text (last 300 lines) ===');
  const start = Math.max(0, lines.length - 300);
  for (let i = start; i < lines.length; i++) {
    console.log(`${i}: ${lines[i]}`);
  }

  await browser.close();
}

explore().catch(console.error);
