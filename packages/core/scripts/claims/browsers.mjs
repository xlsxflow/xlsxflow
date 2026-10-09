// Serves packages/core on localhost and runs claims.mjs and examples/browser-download.mjs in Playwright's
// Chromium, Firefox and WebKit.
// Playwright comes from this folder's package.json (npm ci), or the module file PLAYWRIGHT names.
import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { extname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { FIXTURES } from './claims.mjs';
import { SheetReader, createBlobReader } from '../../dist/index.mjs';

const root = fileURLToPath(new URL('../../', import.meta.url));
const repo = fileURLToPath(new URL('../../../../', import.meta.url));
const types = { '.mjs': 'text/javascript', '.html': 'text/html' };
const page = `<!doctype html><meta charset="utf-8">
<script type="importmap">{ "imports": { "@xlsxflow/core": "/dist/index.mjs" } }</script>
<script type="module">
  import * as lib from '/dist/index.mjs';
  import { run } from '/scripts/claims/claims.mjs';
  const fx = {};
  for (const [k, f] of Object.entries(${JSON.stringify(FIXTURES)})) fx[k] = new Uint8Array(await (await fetch('/test/fixtures/' + f)).arrayBuffer());
  window.downloadXlsx = (await import('/examples/browser-download.mjs')).downloadXlsx;
  window.results = await run(lib, fx);
</script>`;

// The "Export to Excel" recipe must start a download of a file that reads back as written
async function downloadRecipe(tab) {
  const name = 'Recipe: examples/browser-download.mjs starts a download';
  try {
    const [download] = await Promise.all([
      tab.waitForEvent('download'),
      tab.evaluate(() => window.downloadXlsx('people.xlsx', ['Name', 'Age'], [['Ana', 34], ['Bo', 27]])),
    ]);
    const rows = [];
    for await (const r of await new SheetReader().parse(createBlobReader(new Blob([await readFile(await download.path())])))) rows.push(r.cells);
    const got = JSON.stringify([download.suggestedFilename(), rows]);
    const want = JSON.stringify(['people.xlsx', [['Name', 'Age'], ['Ana', 34], ['Bo', 27]]]);
    return got === want ? { name, ok: true } : { name, ok: false, error: `got ${got}` };
  } catch (e) {
    return { name, ok: false, error: e.message };
  }
}

export default async function browsers(report) {
  const pw = await import(process.env.PLAYWRIGHT ? pathToFileURL(process.env.PLAYWRIGHT).href : 'playwright');
  const server = createServer(async (req, res) => {
    const path = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    if (path === '/') return res.writeHead(200, { 'content-type': 'text/html' }).end(page);
    if (!/^\/(dist|scripts\/claims|test\/fixtures|examples)\/[\w.-]+$/.test(path)) return res.writeHead(404).end();
    const file = (path.startsWith('/examples/') ? repo : root) + path.slice(1);
    try { res.writeHead(200, { 'content-type': types[extname(path)] ?? 'application/octet-stream' }).end(await readFile(file)); }
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
        ok = report(`${type.name()} ${browser.version()}`, [...await tab.evaluate(() => window.results), await downloadRecipe(tab)]) && ok;
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
