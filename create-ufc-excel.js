const ExcelJS = require('exceljs');
const { chromium } = require('playwright');
const { exec } = require('child_process');
const fs = require('fs');
const Anthropic = require('@anthropic-ai/sdk').default ?? require('@anthropic-ai/sdk');

const UFC_EVENT_URL      = process.argv[2] || 'https://www.ufc.com/event/ufc-fight-night-march-21-2026';
const TAPOLOGY_EVENT_URL = process.argv[3] || null;
const PREDICTIONS_URL    = process.argv[4] || null;

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

// Scrape community pick % from a Tapology event page.
// Returns a Map: normalized-fighter-name → pick% string (e.g. "65%")
async function scrapeTapologyPicks(page, url) {
  console.log('\nLoading Tapology event page...');
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(3000);

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

  const text = await page.evaluate(() => document.body.innerText);
  const lines = text.split('\n').map(l => l.trim()).filter(l => l.length > 0);

  // Find the predictions breakdown section (flexible match)
  const startIdx = lines.findIndex(l => l.includes('breakdown of event predictions') || l.includes('breakdown of predictions'));
  if (startIdx < 0) {
    console.log('Tapology: predictions section not found');
    return new Map();
  }

  // Collect prediction lines, skipping table headers and stop at end marker
  const SKIP = new Set(['KO/TKO', 'Submission', 'Decision']);
  const predLines = [];
  for (let i = startIdx + 1; i < lines.length; i++) {
    const l = lines[i];
    if (l.startsWith('Update /') || l.startsWith('Verify your')) break;
    if (SKIP.has(l) || l === ' ') continue;
    predLines.push(l);
  }

  // predLines alternates: name, pct%, name, pct%, ...
  const picksMap = new Map();
  for (let i = 0; i < predLines.length - 1; i++) {
    const curr = predLines[i];
    const next = predLines[i + 1];
    if (!curr.match(/^\d+%$/) && next.match(/^\d+%$/)) {
      picksMap.set(normName(curr), next);
      i++; // skip the percentage line we just consumed
    }
  }

  console.log(`Tapology: found picks for ${picksMap.size} fighters`);
  for (const [name, pct] of picksMap) console.log(`  ${name}: ${pct}`);

  return picksMap;
}

// Look up a fighter's Tapology pick%, trying full name then last name fallback
const NAME_SUFFIXES = new Set(['jr', 'sr', 'ii', 'iii', 'iv']);
function lookupPick(picksMap, fullName) {
  if (!picksMap || picksMap.size === 0) return 'N/A';
  const norm = normName(fullName);
  if (picksMap.has(norm)) return picksMap.get(norm);
  // Last-name fallback (strip suffixes like jr/sr first)
  const parts = norm.split(' ').filter(p => !NAME_SUFFIXES.has(p));
  const lastName = parts[parts.length - 1];
  for (const [key, val] of picksMap) {
    const keyParts = key.split(' ').filter(p => !NAME_SUFFIXES.has(p));
    if (keyParts[keyParts.length - 1] === lastName) return val;
  }
  return 'N/A';
}

// Scrape a predictions article and use Claude to extract per-fighter summaries.
// Returns a Map: normalized-fighter-name → short prediction string (e.g. "Evloev by Dec")
async function scrapePredictions(page, url) {
  console.log('\nLoading predictions page...');
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForTimeout(3000);
  const articleText = await page.evaluate(() => document.body.innerText);

  console.log('Extracting predictions with Claude...');
  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
  const response = await client.messages.create({
    model: 'claude-opus-4-6',
    max_tokens: 1024,
    messages: [{
      role: 'user',
      content: `From this MMA betting article, extract the prediction for every fighter mentioned.
Return ONLY a valid JSON object mapping each fighter's name (as it appears in the article) to a short prediction summary.
- For a straight pick, use: "Riley by Dec", "Duncan by KO", "Baraniewski ML"
- For an over/under prop, use: "Over 2.5 Rds", "Under 1.5 Rds"
- If the same prop applies to both fighters in a fight, include both as separate keys with the same value.
- Keep summaries under 20 characters.

Article:
${articleText.slice(0, 8000)}`,
    }],
  });

  let predsMap = new Map();
  try {
    const text = response.content[0].text.trim();
    const json = JSON.parse(text.replace(/^```json\n?|\n?```$/g, ''));
    for (const [name, pred] of Object.entries(json)) {
      predsMap.set(normName(name), String(pred));
    }
    console.log(`Predictions: extracted ${predsMap.size} entries`);
    for (const [name, pred] of predsMap) console.log(`  ${name}: ${pred}`);
  } catch (e) {
    console.log('Predictions: failed to parse Claude response', e.message);
  }
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
    if (keyParts[keyParts.length - 1] === lastName) return val;
  }
  return '';
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
    l.length > 1 &&
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

  const tabNames = ['MATCHUP STATS', 'WIN BY', 'SIGNIFICANT STRIKES', 'GRAPPLING', 'ODDS'];
  const tabData  = {};

  for (const name of tabNames) {
    try {
      const tab = page.locator(`text=${name}`).first();
      if (await tab.isVisible({ timeout: 1500 })) {
        await tab.click();
        await page.waitForTimeout(2000);
      }
    } catch {}
    const text = await page.evaluate(() => document.body.innerText);
    tabData[name] = cleanLines(text);
  }

  const m = tabData['MATCHUP STATS'];
  const w = tabData['WIN BY'];
  const s = tabData['SIGNIFICANT STRIKES'];
  const g = tabData['GRAPPLING'];
  const o = tabData['ODDS'];

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
  const moneyLine     = findStat(o, 'MONEY LINE');

  const build = side => ({
    moneyLine:      moneyLine[side],
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
  // Check if the output file is currently open in Excel (Windows creates a lock file)
  const today = new Date();
  const tag = `${today.getMonth() + 1}-${today.getDate()}-${today.getFullYear()}`;
  const outputPath = `C:\\Users\\User\\UFC_FightNight_${tag}.xlsx`;
  const lockFile   = `C:\\Users\\User\\~$UFC_FightNight_${tag}.xlsx`;

  if (fs.existsSync(lockFile)) {
    console.error(`\nERROR: UFC_FightNight_${tag}.xlsx is currently open in Excel.`);
    console.error('Please close the file first, then run the script again.\n');
    return;
  }

  const browser = await chromium.launch({ headless: false, args: ['--start-maximized'] });
  const context = await browser.newContext({
    viewport: null,
    userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
  });
  const page = await context.newPage();

  // ── Step 0: Scrape Tapology picks and article predictions (if URLs provided) ─
  const tapologyPicks = TAPOLOGY_EVENT_URL
    ? await scrapeTapologyPicks(page, TAPOLOGY_EVENT_URL)
    : null;
  const predictions = PREDICTIONS_URL
    ? await scrapePredictions(page, PREDICTIONS_URL)
    : null;

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
      return {
        fightId:     el.getAttribute('data-fmid'),
        red:         el.querySelector('.c-listing-fight__corner-name--red')?.innerText.trim(),
        blue:        el.querySelector('.c-listing-fight__corner-name--blue')?.innerText.trim(),
        weightClass: el.querySelector('.c-listing-fight__class-text')?.innerText.trim(),
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
    results.push({
      ...fight,
      card:  fight.card,
      stats,
    });
  }

  await browser.close();

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
