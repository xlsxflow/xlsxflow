// Runs the recipes in /examples against the built library (pnpm build first; Node 22.13+ for node:sqlite):
//   node scripts/claims/examples.mjs
// The browser recipe runs in browsers.mjs.
import { readFileSync } from 'node:fs';
import { mkdtemp, rm } from 'node:fs/promises';
import { createServer } from 'node:http';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { DatabaseSync } from 'node:sqlite';
import * as lib from '../../dist/index.mjs';

const dist = new URL('../../dist/index.mjs', import.meta.url).href;
// Loads an example with '@xlsxflow/core' pointing at dist, as an installed package would
const example = (name) => import('data:text/javascript,' + encodeURIComponent(
  readFileSync(new URL(`../../../../examples/${name}`, import.meta.url), 'utf8').replaceAll("'@xlsxflow/core'", `'${dist}'`)));
const read = async (blob) => {
  const out = [];
  for await (const r of await new lib.SheetReader().parse(lib.createBlobReader(blob))) out.push(r.cells);
  return out;
};
const eq = (a, b, what) => { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(`${what}: got ${JSON.stringify(a)}, want ${JSON.stringify(b)}`); };
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const tests = {
  async 'web-response.mjs'() {
    const res = await (await example('web-response.mjs')).GET();
    eq([res.headers.get('content-type'), res.headers.get('content-disposition')], [XLSX_TYPE, 'attachment; filename="orders.xlsx"'], 'headers');
    const rows = await read(await res.blob());
    eq([rows.length, rows[0], rows[10_000]], [10_001, ['Order', 'Customer', 'Total', 'Placed'], [10_000, 'Customer 0', 15_000, '2026-05-26T00:00:00.000Z']], 'rows');
  },
  async 'node-http.mjs'() {
    const { sendReport } = await example('node-http.mjs');
    const server = createServer(sendReport).listen(0, '127.0.0.1');
    await new Promise((r) => server.once('listening', r));
    try {
      const res = await fetch(`http://127.0.0.1:${server.address().port}/`);
      eq(res.headers.get('content-type'), XLSX_TYPE, 'content-type');
      const rows = await read(await res.blob());
      eq([rows.length, rows[1]], [366, ['2026-01-01T00:00:00.000Z', 1037]], 'rows');
    } finally {
      server.close();
    }
  },
  async 'upload.mjs'() {
    const { POST } = await example('upload.mjs');
    const post = async (file) => {
      const form = new FormData();
      if (file) form.set('file', file, 'people.xlsx');
      const res = await POST(new Request('http://localhost/upload', { method: 'POST', body: form }));
      return [res.status, await res.json()];
    };
    const xlsx = await new Response(new lib.SheetWriter().addSheet('People', [['Name', 'Age'], ['Ana', 34], ['Bo']]).write()).blob();
    eq(await post(xlsx), [200, { columns: ['Name', 'Age'], records: [{ Name: 'Ana', Age: 34 }, { Name: 'Bo', Age: null }] }], 'upload');
    eq((await post(null))[0], 400, 'no file');
    eq((await post(new Blob(['not a spreadsheet'])))[0], 422, 'not a spreadsheet');
  },
  async 'database-export.mjs'() {
    const { exportTable } = await example('database-export.mjs');
    const db = new DatabaseSync(':memory:');
    db.exec('CREATE TABLE users (id INTEGER PRIMARY KEY, name TEXT, email TEXT, created_at TEXT)');
    const insert = db.prepare('INSERT INTO users VALUES (?, ?, ?, ?)');
    for (let i = 1; i <= 5000; i++) insert.run(i, `User ${i}`, `user${i}@example.com`, '2026-10-09T08:00:00Z');
    const dir = await mkdtemp(join(tmpdir(), 'xlsxflow-examples-'));
    try {
      await exportTable(db, join(dir, 'users.xlsx'));
      const rows = [];
      for await (const r of await lib.XlsxFlow.readFile(join(dir, 'users.xlsx'))) rows.push(r.cells);
      eq([rows.length, rows[5000]], [5001, [5000, 'User 5000', 'user5000@example.com', '2026-10-09T08:00:00.000Z']], 'rows');
    } finally {
      await rm(dir, { recursive: true });
    }
  },
};

let ok = true;
for (const [name, test] of Object.entries(tests)) {
  try { await test(); console.log(`ok   examples/${name}`); }
  catch (e) { console.log(`FAIL examples/${name}\n     ${e.message}`); ok = false; }
}
if (!ok) process.exit(1);
