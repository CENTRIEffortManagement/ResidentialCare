// Render raw Mermaid charts with an installed Mermaid bundle and local Chromium.
// Dependencies and machine paths are supplied by the caller, never embedded.
const fs = require('node:fs');
const path = require('node:path');
const { chromium } = require('playwright');
(async () => {
  const [bundle, browserPath, ...charts] = process.argv.slice(2);
  if (!bundle || !browserPath || !charts.length) throw new Error('Supply Mermaid sidebar bundle, Chromium executable and chart paths.');
  const browser = await chromium.launch({ executablePath: browserPath, headless: true });
  try {
    const page = await browser.newPage({ viewport: { width: 1600, height: 1200 } });
    await page.setContent('<main id="diagram"></main>');
    await page.addScriptTag({ path: bundle });
    for (const file of charts) {
      const content = fs.readFileSync(file, 'utf8');
      const result = await page.evaluate(async text => {
        const svg = await window.sidebarMermaid.render(text, 'verifiedGantt');
        if (!svg) throw new Error('Mermaid rejected the chart.');
        document.getElementById('diagram').innerHTML = svg;
        return { bars: document.querySelectorAll('rect.task').length, text: document.getElementById('diagram').textContent };
      }, content);
      if (!result.bars) throw new Error('No task bars rendered: ' + file);
      console.log('PASS: ' + path.basename(file) + ' — ' + result.bars + ' bars rendered');
    }
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
