// Saves the generator's DRAFT output for the sample plan into examples/, so reviewers can open the files without running anything.
// Run: node scripts/export_examples.js
const { chromium } = require('playwright'); const fs = require('fs'), path = require('path');
const ROOT = path.resolve(__dirname, '..'), OUT = path.join(ROOT, 'examples');
(async () => {
  const b = await chromium.launch(); const p = await b.newPage({ acceptDownloads: true });
  await p.route('**/xlsx.full.min.js', r => r.fulfill({ path: require.resolve('xlsx/dist/xlsx.full.min.js') }));
  await p.route('https://fonts.**', r => r.abort());
  await p.goto('file://' + path.join(ROOT, 'index.html')); await p.check('#draft');
  for (const id of ['#gBtn', '#mBtn', '#mapBtn']) {
    const [d] = await Promise.all([p.waitForEvent('download'), p.click(id)]);
    await d.saveAs(path.join(OUT, d.suggestedFilename())); console.log('saved', d.suggestedFilename());
  }
  await b.close();
})();
