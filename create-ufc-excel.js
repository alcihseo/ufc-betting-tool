const ExcelJS = require('exceljs');
const { chromium } = require('playwright-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
chromium.use(StealthPlugin());
const { exec } = require('child_process');
const fs = require('fs');
const path = require('path');

const config = fs.existsSync('./config.json') ? JSON.parse(fs.readFileSync('./config.json', 'utf8')) : {};
const UFC_EVENT_URL      = process.argv[2] || config.ufcEventUrl       || 'https://www.ufc.com/event/ufc-fight-night-march-21-2026';
const TAPOLOGY_EVENT_URL = process.argv[3] || config.tapologyUrl       || null;
const PREDICTIONS_URL    = process.argv[4] || config.mmaManiaPredictionsUrl || null;
const SHERDOG_URLS       = (() => {
  if (process.argv[5]) return [process.argv[5]];
  if (config.sherdogUrls) return config.sherdogUrls;
  if (config.sherdogUrl)  return [config.sherdogUrl];
  return null;
})();
const SHERDOG_BASE       = 'https://www.sherdog.com';
const MMA_JUNKIE_URL     = process.argv[6] || config.mmaJunkieUrl      || null;

function eventUrlToSheetName(url) {
  const slug = url.split('/').pop(); // e.g. "ufc-fight-night-march-21-2026"
  return slug.replace(/-/g, ' ').replace(/\b\w/g, c => c.toUpperCase()).slice(0, 31);
}

// Normalize a fighter name for loose matching (lowercase, letters/spaces only)
function normName(name) {
  return (name || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // strip accents
    .replace(/[\r\n\t]+/g, ' ')  // newlines/tabs → space
    .replace(/[^a-z ]/g, '')     // remove remaining non-letter chars
    .replace(/\s+/g, ' ')        // collapse multiple spaces
    .trim();
}

// Read document.body.innerText, retrying while the page is mid-navigation
// (e.g. ad/consent-driven client-side redirects that briefly null out the body).
// Each attempt is raced against its own short timeout — a page whose JS thread
// is wedged by a heavy ad/tracker script can leave page.evaluate() hanging
// indefinitely, which would otherwise defeat the overall timeoutMs budget below.
async function getBodyText(page, timeoutMs = 15000, attemptTimeoutMs = 3000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const text = await Promise.race([
      page.evaluate(() => document.body ? document.body.innerText : null).catch(() => null),
      new Promise(resolve => setTimeout(() => resolve(null), attemptTimeoutMs)),
    ]);
    if (text) return text;
    await page.waitForTimeout(300);
  }
  throw new Error('Page body never settled (stuck mid-navigation, or page JS thread is wedged)');
}

// Close a throwaway browser without risking the whole script hanging forever.
// browser.close() waits for the underlying Chrome process to acknowledge a
// graceful shutdown, which occasionally never happens on ad-heavy sites —
// seen in practice leaving an orphaned chrome.exe process (and the awaited
// close() promise pending indefinitely) after MMA Mania/MMA Junkie scraping.
// Race it against a timeout and force-kill the process if it doesn't close.
async function closeBrowserSafely(browser, timeoutMs = 10000) {
  try {
    await Promise.race([
      browser.close(),
      new Promise((_, reject) => setTimeout(() => reject(new Error('timed out')), timeoutMs)),
    ]);
  } catch (e) {
    console.log(`  (Browser close did not finish in time — force-killing: ${e.message})`);
    try { browser.process()?.kill('SIGKILL'); } catch {}
  }
}

// Scrape community pick % from a Tapology event page.
// Returns a Map: normalized-fighter-name → pick% string (e.g. "65%")
async function scrapeTapologyPicks(page, url) {
  console.log('\nLoading Tapology event page...');
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(3000);

  // Wait for Cloudflare / human verification to be completed if it appears.
  // Detect via the Turnstile challenge iframe (reliable across sites) plus a
  // broad set of known Cloudflare phrasings — wording varies by site (Tapology
  // says "verifies you are not a bot" rather than UFC.com's "verify you are human").
  const CHALLENGE_PHRASES = [
    'verify you are human',
    'verifies you are not a bot',
    'checking if the site connection is secure',
    'performing security verification',
    'checking your browser',
    'please stand by',
    'attention required',
  ];
  const isChallengePage = (phrases) =>
    !!document.querySelector('iframe[src*="challenges.cloudflare.com"]') ||
    document.title.toLowerCase().includes('just a moment') ||
    phrases.some(p => document.body.innerText.toLowerCase().includes(p));
  const isClearOfChallenge = (phrases) =>
    !document.querySelector('iframe[src*="challenges.cloudflare.com"]') &&
    !document.title.toLowerCase().includes('just a moment') &&
    !phrases.some(p => document.body.innerText.toLowerCase().includes(p));

  try {
    const isChallenge = await page.evaluate(isChallengePage, CHALLENGE_PHRASES);
    if (isChallenge) {
      console.log('\n>>> Tapology is showing a bot verification. Complete it in the browser window, then wait — the script will continue automatically...');
      await page.waitForFunction(isClearOfChallenge, CHALLENGE_PHRASES, { timeout: 120000, polling: 500 });
      console.log('  Verification completed, continuing...');
      await page.waitForTimeout(2000);
    }
  } catch (e) {
    console.log('  (Verification check timed out or failed — continuing anyway)');
  }

  // Dismiss cookie consent if present
  try {
    const cookieBtn = page.locator('button:has-text("OK"), button:has-text("Accept"), button:has-text("I Agree"), button:has-text("Accept All")').first();
    if (await cookieBtn.isVisible({ timeout: 2000 })) {
      await cookieBtn.click();
      await page.waitForTimeout(1000);
    }
  } catch {}

  // Click the Predictions tab — try multiple possible tab labels
  const predTabLabels = ['Predictions', 'Community Picks', 'Picks', 'Results'];
  for (const label of predTabLabels) {
    try {
      const tab = page.locator(`text=${label}`).first();
      if (await tab.isVisible({ timeout: 1500 })) {
        await tab.click();
        await page.waitForTimeout(3000);
        break;
      }
    } catch {}
  }

  const text = await getBodyText(page);
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  // Find the predictions breakdown section (flexible match — text differs pre/post event)
  const startIdx = lines.findIndex(l =>
    l.includes('breakdown of event predictions') ||
    l.includes('breaking down as follows') ||
    l.includes('total users made predictions') ||
    l.includes('community picks') ||
    l.includes('users picked') ||
    l.includes('prediction breakdown')
  );

  const picksMap = new Map();

  // Extract name/pct% pairs from a slice of lines
  const SKIP = new Set(['KO/TKO', 'Submission', 'Decision']);
  function extractPairs(slice) {
    for (let i = 0; i < slice.length - 1; i++) {
      const curr = slice[i];
      const next = slice[i + 1];
      if (SKIP.has(curr) || curr === ' ') continue;
      if (!curr.match(/^\d+%$/) && next.match(/^\d+%$/)) {
        picksMap.set(normName(curr), next);
        i++;
      }
    }
  }

  if (startIdx >= 0) {
    // Collect prediction lines from the section header onward
    const predLines = [];
    for (let i = startIdx + 1; i < lines.length; i++) {
      const l = lines[i];
      if (l.startsWith('Update /') || l.startsWith('Verify your')) break;
      if (SKIP.has(l) || l === ' ') continue;
      predLines.push(l);
    }
    extractPairs(predLines);
  }

  // Fallback: scan the full page for name/XX% pairs if section header wasn't found
  // or if the section-based parse found nothing
  if (picksMap.size === 0) {
    console.log('Tapology: section header not found (or empty) — trying full-page fallback...');
    // Save page text for debugging
    try {
      fs.writeFileSync('./tapology-debug.txt', lines.join('\n'), 'utf8');
      console.log('  Debug page text saved to: tapology-debug.txt');
    } catch {}
    // Filter to lines that look like fighter names (short, no URLs, not nav labels)
    const candidateLines = lines.filter(l =>
      l.length >= 3 && l.length <= 60 &&
      !l.includes('http') &&
      !l.includes('@') &&
      !l.includes('©')
    );
    extractPairs(candidateLines);
  }

  console.log(`Tapology: found picks for ${picksMap.size} fighters`);
  for (const [name, pct] of picksMap) console.log(`  ${name}: ${pct}`);

  return picksMap;
}

// Look up a fighter's Tapology pick%, trying full name then last name fallback
const NAME_SUFFIXES = new Set(['jr', 'sr', 'ii', 'iii', 'iv']);

function levenshtein(a, b) {
  const dp = Array.from({ length: a.length + 1 }, () => new Array(b.length + 1).fill(0));
  for (let i = 0; i <= a.length; i++) dp[i][0] = i;
  for (let j = 0; j <= b.length; j++) dp[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      dp[i][j] = a[i - 1] === b[j - 1]
        ? dp[i - 1][j - 1]
        : 1 + Math.min(dp[i - 1][j - 1], dp[i - 1][j], dp[i][j - 1]);
    }
  }
  return dp[a.length][b.length];
}

// Tolerate a 1-letter spelling difference between sites (e.g. Tapology's
// "Elliott" vs UFC.com's "Elliot") when comparing last names.
function namesMatch(a, b) {
  if (a === b) return true;
  if (a.length < 4 || b.length < 4) return false;
  return levenshtein(a, b) <= 1;
}

function lookupPick(picksMap, fullName) {
  if (!picksMap || picksMap.size === 0) return 'N/A';
  const norm = normName(fullName);
  if (picksMap.has(norm)) return picksMap.get(norm);
  // Last-name fallback (strip suffixes like jr/sr first)
  const parts = norm.split(' ').filter(p => !NAME_SUFFIXES.has(p));
  const lastName = parts[parts.length - 1];
  for (const [key, val] of picksMap) {
    const keyParts = key.split(' ').filter(p => !NAME_SUFFIXES.has(p));
    if (namesMatch(keyParts[keyParts.length - 1], lastName)) return val;
  }
  return 'N/A';
}

// Format a raw pick string into a short summary (e.g. "Evloev by Dec", "Over 2.5 Rds")
function formatPick(text) {
  const clean = text.replace(/\s*\([+-]?\d+\)/g, '').trim(); // strip odds
  const lastWord = s => s.trim().split(/\s+/).pop();
  const method = m => {
    const u = m.toUpperCase();
    if (u.includes('KO') || u.includes('TKO')) return 'KO';
    if (u.includes('SUB')) return 'Sub';
    if (u.includes('DEC')) return 'Dec';
    return m.slice(0, 5);
  };
  // Over/Under X rounds
  const ou = clean.match(/^(over|under)\s+([\d.]+)\s*rounds?/i);
  if (ou) return `${ou[1][0].toUpperCase() + ou[1].slice(1)} ${ou[2]} Rds`;
  // [Name] to win by [method]
  const toWin = clean.match(/^(.+?)\s+to win by\s+(.+)/i);
  if (toWin) return `${lastWord(toWin[1])} by ${method(toWin[2])}`;
  // [Name] by [method]
  const by = clean.match(/^(.+?)\s+by\s+(.+)/i);
  if (by) return `${lastWord(by[1])} by ${method(by[2])}`;
  // [Name] moneyline
  const ml = clean.match(/^(.+?)\s+moneyline/i);
  if (ml) return `${lastWord(ml[1])} ML`;
  return clean.slice(0, 20);
}

// Scrape a predictions article and parse picks via text matching.
// Returns a Map: normalized-fighter-name → short prediction string (e.g. "Evloev by Dec")
async function scrapePredictions(page, url) {
  console.log('\nLoading predictions page...');

  // SB Nation's ad stack occasionally wedges the page indefinitely (getBodyText
  // never sees a stable body) — a transient failure, not a permanent one, so a
  // fresh reload usually clears it. Retry a few times before giving up, since a
  // single failed attempt used to blank out the whole predictions column.
  let text;
  const maxAttempts = 3;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForTimeout(2000);
      text = await getBodyText(page);
      break;
    } catch (e) {
      console.log(`  (MMA Mania load attempt ${attempt}/${maxAttempts} failed: ${e.message})`);
      if (attempt === maxAttempts) throw e;
      await page.waitForTimeout(2000);
    }
  }

  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  const predsMap = new Map();
  let currentFighters = [];

  for (const line of lines) {
    // Check pick line FIRST to avoid "Movsar" being split on "vs"
    const pickMatch = line.match(/(?:best bet|our pick|pick):\s*(.+)/i);
    if (pickMatch) {
      const raw = pickMatch[1].replace(/^[""""]|[""""]$/g, '').trim();
      const summary = formatPick(raw);
      if (summary && currentFighters.length === 2) {
        predsMap.set(normName(currentFighters[0]), summary);
        predsMap.set(normName(currentFighters[1]), summary);
        currentFighters = [];
      }
      continue;
    }
    // Detect matchup line: "Fighter1 (odds) vs. Fighter2 (odds)"
    // Length-gated because the anchored regex below will happily match a whole
    // prose sentence that merely mentions "vs" in passing (e.g. a reference to
    // an unrelated fight), which would overwrite the real pairing with garbage
    // right before the pick line that's supposed to use it.
    if (line.length > 70) continue;
    const vsMatch = line.match(/^(.+?)\s*(?:\([+-]?\d+\))?\s+vs\.?\s+(.+?)(?:\s*\([+-]?\d+\))?$/i);
    if (vsMatch) {
      const f1 = vsMatch[1].trim();
      const f2 = vsMatch[2].trim();
      if (f1 && f2) currentFighters = [f1, f2];
    }
  }

  console.log(`MMA Mania: found picks for ${predsMap.size} fighters`);
  for (const [name, pick] of predsMap) console.log(`  ${name}: ${pick}`);
  return predsMap;
}

// Look up a fighter's prediction, trying full name then last-name fallback
function lookupPrediction(predsMap, fullName) {
  if (!predsMap || predsMap.size === 0) return '';
  const norm = normName(fullName);
  if (predsMap.has(norm)) return predsMap.get(norm);
  const parts = norm.split(' ').filter(p => !NAME_SUFFIXES.has(p));
  const lastName = parts[parts.length - 1];
  for (const [key, val] of predsMap) {
    const keyParts = key.split(' ').filter(p => !NAME_SUFFIXES.has(p));
    if (namesMatch(keyParts[keyParts.length - 1], lastName)) return val;
  }
  return '';
}

// Convert method text to short form
function sherdogMethodShort(m) {
  const u = m.toLowerCase();
  if (u.includes('ko') || u.includes('tko') || u.includes('knockout')) return 'KO';
  if (u.includes('sub')) return 'Sub';
  return 'Dec';
}

// Extract "[name] by [method]" from article lines for a given fight
function extractSherdogPick(lines, f1, f2) {
  const getLastName = name => {
    const parts = normName(name).split(' ').filter(p => !NAME_SUFFIXES.has(p));
    return parts[parts.length - 1];
  };
  const n1 = getLastName(f1), n2 = getLastName(f2);
  // Match "by" or "via" followed by a finish method
  const methodRe = /\b(?:by|via)\s+(?:\S+\s+){0,4}(decision|unanimous\s+decision|split\s+decision|majority\s+decision|ko(?:\/tko)?|tko|knockout|submission)\b/gi;

  // Prioritize lines that mention "pick"/"prediction"/"take" — more likely the final verdict
  const pickLines = lines.filter(l => /\b(?:pick|prediction|taking|i'll go|going with)\b/i.test(l));
  const searchLines = [...pickLines, ...lines];

  for (const line of searchLines) {
    const lineLower = line.toLowerCase();
    // Pattern 1: "[name] ... by/via [method]". A single paragraph often
    // mentions both fighters and hedges with more than one "by [method]"
    // phrase before reaching its actual verdict (e.g. "Keith picked X by KO,
    // but I'm leaning towards decision"), so take the LAST such phrase in the
    // line and attribute it to whichever fighter's name sits closest
    // (immediately) before it — not just whichever name appears anywhere
    // earlier in the paragraph, which previously favored whichever fighter
    // happened to be listed first regardless of relevance.
    const matches = [...line.matchAll(methodRe)];
    if (matches.length > 0) {
      const mMatch = matches[matches.length - 1];
      const byIdx = mMatch.index;
      let bestName = null, bestIdx = -1;
      for (const name of [n1, n2]) {
        const nameIdx = lineLower.lastIndexOf(name, byIdx);
        if (nameIdx !== -1 && nameIdx > bestIdx) { bestIdx = nameIdx; bestName = name; }
      }
      if (bestName) {
        return `${bestName.charAt(0).toUpperCase() + bestName.slice(1)} by ${sherdogMethodShort(mMatch[1])}`;
      }
    }
    // Pattern 2: "[name] wins by/via [method]"
    for (const name of [n1, n2]) {
      const wm = line.match(new RegExp(`\\b${name}\\b\\s+wins?\\s+(?:by|via)\\s+([\\w\\s/]+?)(?:\\.|,|$)`, 'i'));
      if (wm) return `${name.charAt(0).toUpperCase() + name.slice(1)} by ${sherdogMethodShort(wm[1])}`;
    }
  }
  return '';
}

// Extract picks from a multi-fight page (prelims) by splitting on vs. section headers
function processSherdogMultiFightPage(lines, picksMap) {
  const sectionStarts = [];
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (l.length < 80 && /^.{2,}\s+vs\.?\s+.{2,}$/.test(l) && !/http/i.test(l)) {
      const m = l.match(/^(.+?)\s+vs\.?\s+(.+?)(?:\s*\(.*\))?$/i);
      if (m) sectionStarts.push({ idx: i, f1: m[1].trim(), f2: m[2].trim() });
    }
  }
  for (let s = 0; s < sectionStarts.length; s++) {
    const { idx, f1, f2 } = sectionStarts[s];
    const end = s + 1 < sectionStarts.length ? sectionStarts[s + 1].idx : lines.length;
    const pick = extractSherdogPick(lines.slice(idx, end), f1, f2);
    if (pick) {
      picksMap.set(normName(f1), pick);
      picksMap.set(normName(f2), pick);
      console.log(`  ${f1} vs ${f2}: ${pick}`);
    }
  }
}

// Collect Sherdog preview links from a loaded page (strips fragments, deduplicates via visited set)
async function collectSherdogLinks(page, visited) {
  return page.evaluate(({ base, visited: vis }) => {
    return [...document.querySelectorAll('a[href]')]
      .map(a => {
        const href = a.href.startsWith('http') ? a.href : base + a.href;
        return href.split('#')[0];
      })
      .filter(h => /\/news\/articles\//.test(h) && /Preview/i.test(h) && !vis.includes(h));
  }, { base: SHERDOG_BASE, visited: [...visited] });
}

// Scrape all Sherdog fight preview pages starting from a main event URL.
// Follows links from each discovered page (so prelims sub-pages are also crawled).
async function scrapeSherdogPicks(page, urls) {
  const picksMap = new Map();
  console.log('\nLoading Sherdog previews...');

  const visited  = new Set();
  const queue    = [];
  for (const url of urls) {
    const clean = url.split('#')[0];
    if (!visited.has(clean)) { visited.add(clean); queue.push(clean); }
  }

  while (queue.length > 0) {
    const pageUrl = queue.shift();
    console.log(`  Visiting: ${pageUrl}`);

    try {
      await page.goto(pageUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForTimeout(2000);

      // Discover new preview links on this page and add to queue
      const newLinks = await collectSherdogLinks(page, visited);
      for (const link of newLinks) {
        if (!visited.has(link)) {
          visited.add(link);
          queue.push(link);
        }
      }

      // Extract picks: try each vs. header against the page's own body content.
      // Each preview page repeats the SAME "Jump To »" nav menu and "related
      // articles" sidebar (which mention every fighter on the card, including
      // ones this particular page never actually discusses), so searching past
      // that marker risks pairing one fighter's incidental name-drop with a
      // completely unrelated verdict sentence from a different fight. Truncate
      // to the actual article body, which always contains its own clean
      // "Fighter1 (record) vs. Fighter2 (record)" title line to key off of.
      const text  = await getBodyText(page);
      const allLines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      const jumpIdx = allLines.findIndex(l => l.startsWith('Jump To'));
      const lines = jumpIdx >= 0 ? allLines.slice(0, jumpIdx) : allLines;
      const vsHeaders = lines.filter(l => l.length < 100 && /^.{2,}\s+vs\.?\s+.{2,}$/.test(l) && !/http/i.test(l));
      const seenPairs = new Set();

      for (const header of vsHeaders) {
        // Strip ALL parenthetical content (records, weights) before splitting on
        // "vs" — a fight-title line like "Levi Rodrigues (5-0, 1 NC) vs. Felipe
        // Franco (10-2)" would otherwise leave "(5-0, 1 NC)" attached to f1, and
        // since it contains letters ("NC"), normalizing it produces a garbage
        // "last name" of "nc" that spuriously matches all sorts of words.
        const cleanHeader = header.replace(/\([^)]*\)/g, '').trim();
        const hm = cleanHeader.match(/^(.+?)\s+vs\.?\s+(.+?)$/i);
        if (!hm) continue;
        const f1 = hm[1].trim(), f2 = hm[2].trim();
        const key = normName(f1) + '|' + normName(f2);
        if (seenPairs.has(key)) continue;
        seenPairs.add(key);
        if (picksMap.has(normName(f1)) || picksMap.has(normName(f2))) continue;
        const pick = extractSherdogPick(lines, f1, f2);
        if (pick) {
          picksMap.set(normName(f1), pick);
          picksMap.set(normName(f2), pick);
          console.log(`    ${f1} vs ${f2}: ${pick}`);
        }
      }
    } catch (e) {
      console.log(`  (Skipping ${pageUrl} — ${e.message})`);
    }
  }

  console.log(`Sherdog: found picks for ${picksMap.size} fighters`);
  return picksMap;
}

// Scrape MMA Junkie's staff "Junkie pick results" numbers.
// Returns a Map: normalized-fighter-name → vote count string (e.g. "6").
// Main-card fights list results on a separate "Junkie pick results: Name X,
// Name Y" line below a "Fighter1 vs. Fighter2" header (with Records/Division/
// Odds lines in between); prelim fights list "Fighter1 vs. Fighter2: Name X,
// Name Y" (or ": N/A") all on one line. In both cases the two names in the
// result text are last names only and are NOT guaranteed to appear in the
// same order as the header (e.g. "Cannonier vs. Duncan" but "Duncan 11,
// Cannonier 0"), so results are matched to fighters by last name, not position.
async function scrapeMmaJunkiePicks(page, url) {
  console.log('\nLoading MMA Junkie picks page...');
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(2000);
  const text = await getBodyText(page);
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  const picksMap = new Map();
  const lastNameOf = name => {
    const parts = normName(name).split(' ').filter(p => !NAME_SUFFIXES.has(p));
    return parts[parts.length - 1];
  };
  function applyResults(resultsText, fighters) {
    if (!resultsText || /^n\/a$/i.test(resultsText.trim())) return;
    for (const part of resultsText.split(',').map(s => s.trim())) {
      const m = part.match(/^(.+?)\s+(\d+)$/);
      if (!m) continue;
      const nickLast = lastNameOf(m[1]);
      for (const full of fighters) {
        if (namesMatch(lastNameOf(full), nickLast)) picksMap.set(normName(full), m[2]);
      }
    }
  }

  let currentFighters = [];
  for (const line of lines) {
    // Prelims: "Fighter1 vs. Fighter2: results" all on one line
    const combined = line.match(/^(.+?)\s+vs\.?\s+(.+?):\s*(.+)$/i);
    if (combined) {
      applyResults(combined[3], [combined[1].trim(), combined[2].trim()]);
      continue;
    }
    // Main card: apply the most recently seen header to its results line
    const resultsOnly = line.match(/^Junkie pick results:\s*(.+)$/i);
    if (resultsOnly) {
      applyResults(resultsOnly[1], currentFighters);
      continue;
    }
    // Main card header: plain "Fighter1 vs. Fighter2" with no colon/results.
    // Length-gated to avoid matching longer related-article titles that
    // happen to mention "X vs. Y" in passing.
    const header = line.match(/^(.+?)\s+vs\.?\s+(.+)$/i);
    if (header && line.length < 60) {
      currentFighters = [header[1].trim(), header[2].trim()];
    }
  }

  console.log(`MMA Junkie: found picks for ${picksMap.size} fighters`);
  for (const [name, val] of picksMap) console.log(`  ${name}: ${val}`);
  return picksMap;
}

// Lines that are labels/nav — never treated as values
const NON_VALUES = new Set([
  'Skip to main content', 'MATCHUP STATS', 'WIN BY', 'SIGNIFICANT STRIKES',
  'GRAPPLING', 'ODDS', 'OK', 'RECORD', 'LAST FIGHT', 'COUNTRY', 'HEIGHT',
  'WEIGHT', 'REACH', 'LEG REACH', 'KO/TKO', 'SUB', 'DEC', 'AVG FIGHT TIME',
  'KNOCKDOWN AVG', 'PER 15 MIN', 'LANDED PER MIN', 'ABSORBED PER MIN',
  'DEFENSE', 'TAKEDOWN AVG', 'TAKEDOWN ACCURACY', 'TAKEDOWN DEFENSE',
  'SUBMISSION AVG', 'MONEY LINE', 'TOTAL ROUNDS', 'ODDS TO WIN BY KO',
  'ODDS TO WIN BY SUBMISSION', 'ODDS TO WIN BY DECISION',
]);

// ─── Win probability model ────────────────────────────────────────────────────

function toNum(val) {
  if (!val || val === 'N/A') return null;
  return parseFloat(String(val).replace('%', ''));
}

function parseWinPct(record) {
  if (!record || record === 'N/A') return null;
  const m = record.match(/^(\d+)-(\d+)/);
  if (!m) return null;
  const total = parseInt(m[1]) + parseInt(m[2]);
  return total > 0 ? parseInt(m[1]) / total : 0.5;
}

function moneyLineToProb(ml) {
  const n = parseFloat(ml);
  if (!ml || ml === 'N/A' || isNaN(n)) return null;
  // Remove the vig (overround) by returning raw implied probability
  return n < 0 ? Math.abs(n) / (Math.abs(n) + 100) : 100 / (n + 100);
}

// Weighted stat diff — returns 0 if either value is missing
function wd(a, b, weight, scale) {
  return (a !== null && b !== null) ? weight * (a - b) / scale : 0;
}

// Compute model win probability for red fighter (0–1)
function computeModelProb(r, b) {
  const rLanded  = toNum(r.landedPerMin),   bLanded  = toNum(b.landedPerMin);
  const rSigStr  = toNum(r.sigStr),          bSigStr  = toNum(b.sigStr);
  const rAbsorb  = toNum(r.absorbedPerMin),  bAbsorb  = toNum(b.absorbedPerMin);
  const rDef     = toNum(r.sigStrDefense),   bDef     = toNum(b.sigStrDefense);
  const rTdAvg   = toNum(r.tdAvg),           bTdAvg   = toNum(b.tdAvg);
  const rTdAcc   = toNum(r.tdAccuracy),      bTdAcc   = toNum(b.tdAccuracy);
  const rTdDef   = toNum(r.tdDefense),       bTdDef   = toNum(b.tdDefense);
  const rWinPct  = parseWinPct(r.record),    bWinPct  = parseWinPct(b.record);

  // Score from red's perspective — positive = red has the edge
  let score = 0;
  score += wd(rLanded,  bLanded,  1.5, 3.0);   // striking volume (landed/min)
  score += wd(rSigStr,  bSigStr,  1.0, 25);    // striking accuracy (%)
  score += wd(bAbsorb,  rAbsorb,  1.5, 3.0);   // absorbed/min flipped: lower absorbed = better
  score += wd(rDef,     bDef,     1.5, 25);    // strike defense (%)
  score += wd(rTdAvg,   bTdAvg,   1.0, 3.0);   // takedown volume
  score += wd(rTdAcc,   bTdAcc,   0.8, 25);    // takedown accuracy (%)
  score += wd(rTdDef,   bTdDef,   1.0, 25);    // takedown defense (%)
  score += wd(rWinPct,  bWinPct,  0.5, 0.3);   // career win %

  // Sigmoid converts score to a 0–1 probability
  return 1 / (1 + Math.exp(-score * 1.5));
}

// ─────────────────────────────────────────────────────────────────────────────

function cleanLines(text) {
  return text.split('\n').map(l => l.trim()).filter(l =>
    (l.length > 1 || /^\d$/.test(l)) &&
    !l.includes('Cookie') &&
    !l.includes('clicking') &&
    !l.includes('storing of') &&
    !l.includes('enhance site')
  );
}

// Find a stat label in lines (skipping nav at top) and return the red/blue values around it.
// unitLine: if the label is followed by a unit line (e.g. "PER 15 MIN"), blue value is 2 lines after.
function findStat(lines, label, unitLine = null) {
  const idx = lines.indexOf(label, 6); // skip first 6 nav lines
  if (idx < 0) return { red: null, blue: null };

  const redLine  = lines[idx - 1];
  const blueIdx  = (unitLine && lines[idx + 1] === unitLine) ? idx + 2 : idx + 1;
  const blueLine = lines[blueIdx];

  const valid = v => v && !NON_VALUES.has(v) && v !== 'N/A';
  return { red: valid(redLine) ? redLine : null, blue: valid(blueLine) ? blueLine : null };
}

async function scrapeMatchup(page, url) {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(4000);

  const tabNames = ['MATCHUP STATS', 'WIN BY', 'SIGNIFICANT STRIKES', 'GRAPPLING'];
  const tabData  = {};

  for (const name of tabNames) {
    try {
      const tab = page.locator(`text=${name}`).first();
      if (await tab.isVisible({ timeout: 1500 })) {
        await tab.click();
        await page.waitForTimeout(2000);
      }
    } catch {}
    const text = await getBodyText(page);
    tabData[name] = cleanLines(text);
  }

  const m = tabData['MATCHUP STATS'];
  const w = tabData['WIN BY'];
  const s = tabData['SIGNIFICANT STRIKES'];
  const g = tabData['GRAPPLING'];

  // Parse each stat — structure is always: red_value, LABEL, blue_value
  const record        = findStat(m, 'RECORD');
  const height        = findStat(m, 'HEIGHT');
  const reach         = findStat(m, 'REACH');
  const ko            = findStat(w, 'KO/TKO');
  const sub           = findStat(w, 'SUB');
  const dec           = findStat(w, 'DEC');
  const landedPerMin  = findStat(s, 'LANDED PER MIN');
  const sigStr        = findStat(s, 'SIGNIFICANT STRIKES');  // sig strike accuracy %
  const absorbedPerMin= findStat(s, 'ABSORBED PER MIN');
  const sigStrDefense = findStat(s, 'DEFENSE');              // sig strike defense %
  const tdAvg         = findStat(g, 'TAKEDOWN AVG', 'PER 15 MIN');
  const tdAccuracy    = findStat(g, 'TAKEDOWN ACCURACY');
  const tdDefense     = findStat(g, 'TAKEDOWN DEFENSE');

  const build = side => ({
    record:         record[side],
    height:         height[side],
    reach:          reach[side],
    landedPerMin:   landedPerMin[side],
    sigStr:         sigStr[side],
    absorbedPerMin: absorbedPerMin[side],
    sigStrDefense:  sigStrDefense[side],
    ko:             ko[side],
    sub:            sub[side],
    dec:            dec[side],
    tdAvg:          tdAvg[side],
    tdAccuracy:     tdAccuracy[side],
    tdDefense:      tdDefense[side],
  });

  return { red: build('red'), blue: build('blue') };
}

async function createUFCExcel() {
  console.log('UFC URL:      ', UFC_EVENT_URL);
  console.log('Tapology URL: ', TAPOLOGY_EVENT_URL || '(none)');
  console.log('MMA Mania URL:', PREDICTIONS_URL    || '(none)');
  console.log('Sherdog URLs: ', SHERDOG_URLS ? SHERDOG_URLS.join(', ') : '(none)');
  console.log('MMA Junkie URL:', MMA_JUNKIE_URL || '(none)');
  console.log('');

  // Check if the output file is currently open in Excel (Windows creates a lock file)
  const today = new Date();
  const tag = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
  const outputDir  = path.join(__dirname, 'output');
  if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });
  const outputPath = path.join(outputDir, `UFC_FightNight_${tag}.xlsx`);
  const lockFile   = path.join(outputDir, `~$UFC_FightNight_${tag}.xlsx`);

  if (fs.existsSync(lockFile)) {
    console.error(`\nERROR: UFC_FightNight_${tag}.xlsx is currently open in Excel.`);
    console.error('Please close the file first, then run the script again.\n');
    return;
  }

  // Use a persistent profile so Cloudflare's cf_clearance cookie is saved between runs.
  // Uses real installed Chrome (not the bundled Chromium) so the binary passes
  // Cloudflare's integrity checks. Launched via playwright-extra + the puppeteer-extra
  // stealth plugin, which patches navigator.webdriver and the other fingerprints
  // Cloudflare Turnstile checks for — the same combo already proven to work
  // against Cloudflare in the foreclosure-tool project.
  const PROFILE_DIR = path.join(__dirname, '.browser-profile');
  const context = await chromium.launchPersistentContext(PROFILE_DIR, {
    headless: false,
    channel: 'chrome',
    args: ['--start-maximized'],
    viewport: null,
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
  });
  const page = await context.newPage();

  // ── Step 0: Scrape Tapology picks and article predictions (if URLs provided) ─
  const tapologyPicks = TAPOLOGY_EVENT_URL
    ? await scrapeTapologyPicks(page, TAPOLOGY_EVENT_URL)
    : null;
  let predictions = null;
  if (PREDICTIONS_URL) {
    try {
      // Use a throwaway, non-persistent browser for MMA Mania specifically.
      // The shared .browser-profile has accumulated some state (most likely
      // an ad-tracking cookie) that reliably wedges this site's navigation
      // into a permanent about:blank/loading state, while a clean profile
      // loads the exact same URL instantly and reliably every time.
      const freshBrowser = await chromium.launch({ headless: false, channel: 'chrome' });
      try {
        const freshContext = await freshBrowser.newContext({
          viewport: null,
          userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        });
        const freshPage = await freshContext.newPage();
        predictions = await scrapePredictions(freshPage, PREDICTIONS_URL);
      } finally {
        await closeBrowserSafely(freshBrowser);
      }
    } catch (e) {
      console.log(`  (MMA Mania predictions failed — skipping: ${e.message})`);
    }
  }
  let sherdogPicks = null;
  if (SHERDOG_URLS) {
    try {
      sherdogPicks = await scrapeSherdogPicks(page, SHERDOG_URLS);
    } catch (e) {
      console.log(`  (Sherdog previews failed — skipping: ${e.message})`);
    }
  }
  let mmaJunkiePicks = null;
  if (MMA_JUNKIE_URL) {
    try {
      // Same throwaway-browser approach as MMA Mania — no need to risk the
      // shared persistent profile on a site it doesn't need Cloudflare
      // cookies for.
      const freshBrowser = await chromium.launch({ headless: false, channel: 'chrome' });
      try {
        const freshContext = await freshBrowser.newContext({
          viewport: null,
          userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        });
        const freshPage = await freshContext.newPage();
        mmaJunkiePicks = await scrapeMmaJunkiePicks(freshPage, MMA_JUNKIE_URL);
      } finally {
        await closeBrowserSafely(freshBrowser);
      }
    } catch (e) {
      console.log(`  (MMA Junkie picks failed — skipping: ${e.message})`);
    }
  }

  // ── Step 1: Load event page and extract all fight IDs ──────────────────────
  console.log('\nLoading UFC event page...');
  await page.goto(UFC_EVENT_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(5000);

  // Extract all fight data BEFORE clicking expand (clicking hides the name elements)
  const fights = await page.evaluate(() =>
    [...document.querySelectorAll('.c-listing-fight[data-fmid]')].map((el, idx) => {
      // Walk up the DOM to find the card section heading
      let card = 'Prelims';
      let ancestor = el.parentElement;
      while (ancestor) {
        const heading = ancestor.querySelector('h2, .l-listing__group-title, .c-card-event--athlete-results__card-name');
        if (heading) {
          const text = heading.innerText.trim().toUpperCase();
          if (text.includes('MAIN')) { card = 'Main Card'; break; }
          if (text.includes('PRELIM')) { card = 'Prelims'; break; }
        }
        ancestor = ancestor.parentElement;
      }
      // Money line lives on the event listing card itself (.c-listing-fight__odds-row),
      // not on the /matchup/ sub-page — UFC.com removed the ODDS tab from that page.
      const oddsAmounts = [...el.querySelectorAll('.c-listing-fight__odds-amount')].map(sp => sp.innerText.trim());
      return {
        fightId:       el.getAttribute('data-fmid'),
        red:           el.querySelector('.c-listing-fight__corner-name--red')?.innerText.trim(),
        blue:          el.querySelector('.c-listing-fight__corner-name--blue')?.innerText.trim(),
        weightClass:   el.querySelector('.c-listing-fight__class-text')?.innerText.trim(),
        redMoneyLine:  oddsAmounts[0] || null,
        blueMoneyLine: oddsAmounts[1] || null,
        card,
        idx,
      };
    })
  );

  // Now click expand on the first fight to load its iframe and read the event ID
  await page.locator('.c-listing-fight__expand-button').first().click();
  await page.waitForTimeout(3000);

  const eventId = await page.evaluate(() => {
    const src = document.querySelector('.details-content__iframe-wrapper iframe')?.src || '';
    const m = src.match(/\/matchup\/(\d+)\//);
    return m ? m[1] : null;
  });
  console.log(`Event ID: ${eventId}\n`);

  console.log(`Found ${fights.length} fights. Scraping...\n`);

  // ── Step 2: Scrape each fight ───────────────────────────────────────────────
  const results = [];
  for (const fight of fights) {
    const url = `https://www.ufc.com/matchup/${eventId}/${fight.fightId}/pre`;
    console.log(`[${fight.idx + 1}/${fights.length}] ${fight.red} vs ${fight.blue}`);
    const stats = await scrapeMatchup(page, url);
    stats.red.moneyLine  = fight.redMoneyLine;
    stats.blue.moneyLine = fight.blueMoneyLine;
    results.push({
      ...fight,
      card:  fight.card,
      stats,
    });
  }

  await context.close();

  // ── Step 3: Build Excel ─────────────────────────────────────────────────────
  const workbook = new ExcelJS.Workbook();
  const sheet    = workbook.addWorksheet(eventUrlToSheetName(UFC_EVENT_URL));

  sheet.columns = [
    { header: 'Fighter',                    key: 'fighter',         width: 28 },
    { header: 'Money Line',                 key: 'moneyLine',       width: 13 },
    { header: 'Odds Win Prob',              key: 'oddsProb',        width: 14 },
    { header: 'Record (W-L-D)',             key: 'record',          width: 15 },
    { header: 'Height',                     key: 'height',          width: 10 },
    { header: 'Reach',                      key: 'reach',           width: 10 },
    { header: 'KO %',                       key: 'ko',              width: 10 },
    { header: 'Sub %',                      key: 'sub',             width: 10 },
    { header: 'Dec %',                      key: 'dec',             width: 10 },
    { header: 'Sig Strike Landed/Min',      key: 'landedPerMin',    width: 20 },
    { header: 'Sig Strike %',               key: 'sigStr',          width: 13 },
    { header: 'Sig Strike Absorbed/Min',    key: 'absorbedPerMin',  width: 22 },
    { header: 'Sig Strike Defense %',       key: 'sigStrDefense',   width: 20 },
    { header: 'TD Avg/15min',               key: 'tdAvg',           width: 14 },
    { header: 'TD Accuracy %',              key: 'tdAccuracy',      width: 15 },
    { header: 'TD Defense %',               key: 'tdDefense',       width: 15 },
    { header: 'Model Win Prob',             key: 'modelProb',       width: 15 },
    { header: 'Tapology Picks %',           key: 'tapologyPick',    width: 16 },
    { header: 'MMA Mania Betting Picks',    key: 'prediction',      width: 24 },
    { header: 'Sherdog Previews Picks',     key: 'sherdogPick',     width: 22 },
    { header: 'MMA Junkie Picks',           key: 'junkiePick',      width: 16 },
  ];

  const colCount = sheet.columns.length;

  // Header row styling
  sheet.getRow(1).eachCell(cell => {
    cell.font      = { bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
    cell.alignment = { horizontal: 'center' };
  });

  // Dark border style helpers
  const thin   = { style: 'thin',   color: { argb: 'FF000000' } };
  const medium = { style: 'medium', color: { argb: 'FF000000' } };

  const applyFightBorder = (topRow, bottomRow) => {
    for (let col = 1; col <= colCount; col++) {
      const top    = topRow.getCell(col);
      const bottom = bottomRow.getCell(col);
      const isFirst = col === 1;
      const isLast  = col === colCount;

      top.border = {
        top:    medium,
        left:   isFirst ? medium : thin,
        right:  isLast  ? medium : thin,
        bottom: thin,
      };
      bottom.border = {
        top:    thin,
        left:   isFirst ? medium : thin,
        right:  isLast  ? medium : thin,
        bottom: medium,
      };
    }
  };

  const probBg     = 'FFDCE6F1'; // light steel blue for probability columns
  const favoriteBg = 'FFE2EFDA'; // light green for the higher probability fighter

  for (const fight of results) {
    const r = fight.stats.red;
    const b = fight.stats.blue;

    // Model probability
    const redModelRaw  = computeModelProb(r, b);
    const redModelProb  = Math.round(redModelRaw * 100);
    const blueModelProb = 100 - redModelProb;

    // Odds-implied probability (remove vig by normalising the two raw probs)
    const rawRed  = moneyLineToProb(r.moneyLine);
    const rawBlue = moneyLineToProb(b.moneyLine);
    let redOddsProb = null, blueOddsProb = null;
    if (rawRed !== null && rawBlue !== null) {
      const total    = rawRed + rawBlue;
      redOddsProb  = Math.round((rawRed  / total) * 100);
      blueOddsProb = Math.round((rawBlue / total) * 100);
    }

    const addRow = (name, s, modelProb, oddsProb) => {
      const row = sheet.addRow({
        fighter:        name,
        moneyLine:      s.moneyLine      || 'N/A',
        oddsProb:       oddsProb !== null ? `${oddsProb}%` : 'N/A',
        modelProb:      `${modelProb}%`,
        tapologyPick:   lookupPick(tapologyPicks, name),
        prediction:     lookupPrediction(predictions, name),
        sherdogPick:    lookupPrediction(sherdogPicks, name),
        junkiePick:     lookupPick(mmaJunkiePicks, name),
        record:         s.record         || 'N/A',
        height:         s.height         || 'N/A',
        reach:          s.reach          || 'N/A',
        landedPerMin:   s.landedPerMin   || 'N/A',
        sigStr:         s.sigStr         || 'N/A',
        absorbedPerMin: s.absorbedPerMin || 'N/A',
        sigStrDefense:  s.sigStrDefense  || 'N/A',
        ko:             s.ko             || 'N/A',
        sub:            s.sub            || 'N/A',
        dec:            s.dec            || 'N/A',
        tdAvg:          s.tdAvg          || 'N/A',
        tdAccuracy:     s.tdAccuracy     || 'N/A',
        tdDefense:      s.tdDefense      || 'N/A',
      });
      row.eachCell(cell => { cell.alignment = { horizontal: 'center' }; });
      row.getCell('fighter').alignment = { horizontal: 'left' };
      // Subtle background on probability columns
      row.getCell('oddsProb').fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: probBg } };
      row.getCell('modelProb').fill     = { type: 'pattern', pattern: 'solid', fgColor: { argb: probBg } };
      row.getCell('tapologyPick').fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: probBg } };
      return row;
    };

    const redRow  = addRow(fight.red,  r, redModelProb,  redOddsProb);
    const blueRow = addRow(fight.blue, b, blueModelProb, blueOddsProb);

    // Green highlight the model favourite's probability cell
    const favModelRow  = redModelProb  >= blueModelProb  ? redRow  : blueRow;
    const favOddsRow   = (redOddsProb  !== null && redOddsProb  >= blueOddsProb)  ? redRow  : blueRow;
    favModelRow.getCell('modelProb').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: favoriteBg } };
    if (redOddsProb !== null) {
      favOddsRow.getCell('oddsProb').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: favoriteBg } };
    }

    applyFightBorder(redRow, blueRow);
  }

  await workbook.xlsx.writeFile(outputPath);
  console.log(`\nSaved to: ${outputPath}`);
  exec(`start "" "${outputPath}"`);
}

createUFCExcel().catch(console.error);
