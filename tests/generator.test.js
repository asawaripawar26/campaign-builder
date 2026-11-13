// End-to-end tests for the campaign generator (index.html), run in a real headless browser.
// Run: npm test   (needs: npm install, npx playwright install chromium)
// Each test loads the page fresh, so tests don't depend on each other.
const { test, before, after } = require('node:test');
const assert = require('node:assert/strict');
const { chromium } = require('playwright');
const fs = require('fs'), os = require('os'), path = require('path');

const ROOT = path.resolve(__dirname, '..');
const FX = path.join(__dirname, 'fixtures'), ASSETS = path.join(__dirname, 'assets');
const PLAN = path.join(ROOT, 'plan', 'Campaign_Plan_Dec2026.xlsx');
const SHEETJS = require.resolve('xlsx/dist/xlsx.full.min.js');
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'gen-'));
const assetFiles = () => fs.readdirSync(ASSETS).filter(f => !f.endsWith('.txt')).map(f => path.join(ASSETS, f));

let browser;
before(async () => { browser = await chromium.launch(); });
after(async () => { await browser.close(); });

async function open() {
  const page = await browser.newPage({ acceptDownloads: true });
  const errors = []; page.on('pageerror', e => errors.push(e.message));
  // Serve SheetJS from npm instead of the CDN so tests run offline and on a fixed version.
  await page.route('**/xlsx.full.min.js', r => r.fulfill({ path: SHEETJS }));
  await page.route('https://fonts.**', r => r.abort());
  await page.goto('file://' + path.join(ROOT, 'index.html'));
  await page.waitForFunction(() => window.__gen && window.__gen.state.P);
  page.errors = errors;
  return page;
}
const state = page => page.evaluate(() => { const S = window.__gen.state;
  return { camps: S.P.camps.map(c => ({ key: c.key, plat: c.plat, monthly: c.monthly, daily: c.daily, name: c.name })),
    errors: S.issues.filter(i => i.level === 'error').map(i => i.msg), g: S.g.length, m: S.m.length,
    metaSets: [...new Set(S.m.map(r => r['Ad Set Name']))], s2: S.s2.length,
    disabled: Object.fromEntries(['gBtn', 'mBtn', 's1Btn', 's2Btn'].map(id => [id, document.getElementById(id).disabled])) }; });
async function download(page, button) {
  const [d] = await Promise.all([page.waitForEvent('download'), page.click(button)]);
  const file = path.join(TMP, d.suggestedFilename()); await d.saveAs(file);
  return { name: d.suggestedFilename(), text: fs.readFileSync(file, 'utf8').replace(/^﻿/, '') };
}

test('sample plan matches Task 1: 7 campaigns, $109K committed, inside the $120K budget', async () => {
  const p = await open(); const s = await state(p);
  const by = Object.fromEntries(s.camps.map(c => [c.key, c.monthly]));
  assert.deepEqual(by, { Brand: 9000, NonBrand: 29000, CompTest: 5000, PMax: 12000, AdvPlus: 40000, CreativeTest: 10000, RT14d: 4000 });
  const total = s.camps.reduce((a, c) => a + c.monthly, 0);
  assert.equal(total, 109000);
  assert.ok(total + 11000 <= 120000, 'campaigns plus the $11K Meta reserve must fit the $120K budget');
  assert.deepEqual(p.errors, []); await p.close();
});

test('daily budgets: Google = monthly ÷ 30.4, Meta = monthly ÷ days in month (the bug caught in review)', async () => {
  const p = await open(); const s = await state(p);
  for (const c of s.camps) {
    const expected = Math.round(c.monthly / (c.plat === 'Google' ? 30.4 : 31));
    assert.equal(c.daily, expected, `${c.key}: daily ${c.daily}, expected ${expected}`);
  }
  // A ÷31 Google budget would under-spend: $9,000 ÷ 31 = $290/day → $8,816 a month.
  assert.equal(s.camps.find(c => c.key === 'Brand').daily, 296);
  await p.close();
});

test('names follow PLATFORM_FUNNEL_KEY_GEO_MONTH, except PMax which keeps the existing campaign', async () => {
  const p = await open(); const s = await state(p);
  for (const c of s.camps.filter(c => c.key !== 'PMax')) assert.match(c.name, /^(GOOG|META)_[A-Z]+_[A-Za-z0-9]+_US_202612$/, c.name);
  assert.equal(s.camps.find(c => c.key === 'PMax').name, '{{EXISTING_PMAX_NAME}}', 'PMax is edited in place, never recreated');
  await p.close();
});

test('strict export blocks anything unsafe: TEST creative, missing tracking date, PMax without an account export', async () => {
  const p = await open(); const s = await state(p);
  assert.equal(s.disabled.mBtn, true, 'Meta export must be blocked while files are TEST, not Approved');
  assert.equal(s.m, 0);
  assert.equal(s.disabled.s1Btn, true, 'PMax export must be blocked');
  assert.ok(s.errors.some(e => /tracking verified on/.test(e)), 'tracking gate must fire');
  assert.ok(s.errors.some(e => /Editor export/.test(e)), 'PMax must demand an account export');
  await p.close();
});

test('draft export downloads real CSV files with the right columns and budgets', async () => {
  const p = await open(); await p.check('#draft'); const s = await state(p);
  assert.equal(s.g, 61); assert.equal(s.m, 19);
  assert.equal(s.metaSets.length, 5, 'Task 1 has 5 Meta ad sets');
  const g = await download(p, '#gBtn');
  assert.equal(g.name, 'google_search_202612_DRAFT.csv');
  const [hdr, ...rows] = g.text.trim().split(/\r?\n/); const cols = hdr.split(',');
  for (const col of ['Campaign', 'Budget', 'Ad Group', 'Keyword', 'Criterion Type']) assert.ok(cols.includes(col), col);
  const brand = rows.map(r => r.split(',')).find(r => r[0] === 'GOOG_BRAND_Brand_US_202612' && r[cols.indexOf('Budget')]);
  assert.equal(brand[cols.indexOf('Budget')], '296');
  const m = await download(p, '#mBtn');
  assert.equal(m.name, 'meta_bulk_import_202612_DRAFT.csv');
  assert.ok(m.text.startsWith('Campaign Name,'));
  assert.deepEqual(p.errors, []); await p.close();
});

test('the plan template round-trips: uploading the .xlsx gives the same campaigns as the built-in sample', async () => {
  const p = await open(); const sample = await state(p);
  await p.setInputFiles('#fileIn', PLAN); await p.waitForTimeout(400);
  const uploaded = await state(p);
  assert.deepEqual(uploaded.camps, sample.camps);
  await p.close();
});

test('PMax two-stage gate: Stage 2 stays locked until the checklist is ticked', async () => {
  const p = await open();
  await p.setInputFiles('#fileIn', path.join(FX, 'TEST_plan_with_fictional_account_values.xlsx'));
  await p.setInputFiles('#acctIn', path.join(FX, 'TEST_fictional_editor_export.csv'));
  await p.setInputFiles('#creaIn', assetFiles()); await p.waitForTimeout(1500);
  await p.check('#draft');
  assert.equal((await state(p)).disabled.s2Btn, true);
  for (const id of ['#ck0', '#ck1', '#ck2', '#ck3', '#ck4']) await p.check(id);
  const s = await state(p);
  assert.equal(s.disabled.s1Btn, false); assert.equal(s.disabled.s2Btn, false);
  assert.ok(s.s2 > 0);
  await p.close();
});

test('negative checks: wrong-case campaign name in the export, and a creative file in the wrong shape', async () => {
  const p = await open();
  await p.setInputFiles('#fileIn', path.join(FX, 'TEST_plan_with_fictional_account_values.xlsx'));
  const bad = path.join(TMP, 'bad_export.csv');
  fs.writeFileSync(bad, 'Campaign\tAsset Group\r\npmax | us | trials\t\r\n');
  await p.setInputFiles('#acctIn', bad);
  // A square image renamed as 4:5: the checker must read the real pixels, not trust the name.
  const wrong = path.join(TMP, 'CASE_STUDY_v2_4x5_TEST.png');
  fs.copyFileSync(path.join(ASSETS, 'CASE_STUDY_v2_1x1_TEST.png'), wrong);
  await p.setInputFiles('#creaIn', [wrong]); await p.waitForTimeout(600);
  const issues = await p.evaluate(() => window.__gen.state.issues.map(i => i.where + ' | ' + i.msg));
  assert.ok(issues.some(i => /isn't in the campaign|Editor export/.test(i)), 'wrong campaign name must be caught');
  assert.ok(issues.some(i => /CASE_STUDY_v2_4x5/.test(i)), 'wrong-shape file must be caught');
  await p.close();
});
