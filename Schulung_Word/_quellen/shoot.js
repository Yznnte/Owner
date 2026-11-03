const { chromium } = require('playwright');
(async () => {
  const dir = process.argv[2];
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1300, height: 900 }, deviceScaleFactor: 2 });
  await p.goto('file://' + dir + '/figures.html');
  await p.waitForTimeout(500);
  const ids = await p.$$eval('.fig', els => els.map(e => e.id));
  for (const id of ids) await (await p.$('#' + id)).screenshot({ path: `${dir}/img/${id}.png`, omitBackground: true });
  console.log(ids.join(' '));
  await b.close();
})();
