// Запасной рендер через Node-Playwright, если в Python нет модуля playwright.
// node render.cjs <deck.html> <outdir> <width> <height> <check-js>  → JSON-отчёт в stdout
const path = require('path');
const { pathToFileURL } = require('url');
const { chromium } = require('playwright');

(async () => {
  const [html, out, w, h, check] = process.argv.slice(2);
  let browser;
  try { browser = await chromium.launch(); }
  catch (e) { browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' }); }
  const page = await browser.newPage({ viewport: { width: +w, height: +h } });
  await page.goto(pathToFileURL(html).href);
  await page.evaluate(() => document.fonts.ready);
  const report = await page.evaluate(`(${check})()`);
  const slides = await page.$$('.slide');
  for (let i = 0; i < slides.length; i++) {
    await slides[i].screenshot({ path: path.join(out, `slide_${String(i + 1).padStart(2, '0')}.png`) });
  }
  await page.pdf({ path: path.join(out, 'deck.pdf'), width: `${w}px`, height: `${h}px`, printBackground: true });
  await browser.close();
  process.stdout.write(JSON.stringify(report));
})().catch(e => { console.error(e); process.exit(1); });
