// Serves packages/core on localhost and runs claims.mjs in Playwright's Chromium, Firefox and WebKit.
// Playwright comes from this folder's package.json (npm ci), or the module file PLAYWRIGHT names.
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { FIXTURES } from './claims.mjs';

const root = fileURLToPath(new URL('../../', import.meta.url));
const types = { '.mjs': 'text/javascript', '.html': 'text/html' };
const page = `<!doctype html><meta charset="utf-8"><script type="module">
  import * as lib from '/dist/index.mjs';
  import { run } from '/scripts/claims/claims.mjs';
  const fx = {};
  for (const [k, f] of Object.entries(${JSON.stringify(FIXTURES)})) fx[k] = new Uint8Array(await (await fetch('/test/fixtures/' + f)).arrayBuffer());
  window.results = await run(lib, fx);
</script>`;

export default async function browsers(report) {
  const pw = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
  const server = createServer(async (req, res) => {
    const path = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    if (path === '/') return res.writeHead(200, { 'content-type': 'text/html' }).end(page);
    if (!/^\/(dist|scripts\/claims|test\/fixtures)\/[\w.-]+$/.test(path)) return res.writeHead(404).end();
    try { res.writeHead(200, { 'content-type': types[extname(path)] ?? 'application/octet-stream' }).end(await readFile(root + path.slice(1))); }
    catch { res.writeHead(404).end(); }
  }).listen(0, '127.0.0.1');
  await new Promise((r) => server.once('listening', r));
  let ok = true;
  try {
    for (const type of [pw.chromium, pw.firefox, pw.webkit]) {
      let browser;
      try {
        browser = await type.launch();
        const tab = await browser.newPage();
        const errors = [];
        tab.on('pageerror', (e) => errors.push(e.message));
        await tab.goto(`http://127.0.0.1:${server.address().port}/`);
        await tab.waitForFunction(() => window.results, null, { timeout: 300000 }).catch((e) => { throw new Error(errors.join('; ') || e.message); });
        ok = report(`${type.name()} ${browser.version()}`, await tab.evaluate(() => window.results)) && ok;
      } catch (e) {
        console.log(`FAIL ${type.name()}: ${e.message.split('\n')[0]}`);
        ok = false;
      } finally {
        await browser?.close();
      }
    }
  } finally {
    server.close();
  }
  return ok;
}
