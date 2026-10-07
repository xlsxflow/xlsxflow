// Read benchmark: each library reads the same realistic file in its own process.
//   pnpm build && npx tsx scripts/benchmark-read-competitors.ts
// The file is written by ExcelJS (shared strings, numbers, dates, booleans), not by XlsxFlow,
// so no library reads its own output. Memory is peak RSS of the worker process, which is fair
// for both streaming and in-memory readers (post-GC heap would hide an in-memory reader's peak).
import * as fs from 'fs';
import * as path from 'path';
import { spawnSync } from 'child_process';
import { createRequire } from 'module';

const ROWS = 100_000;
const COLS = 10;
const SCRIPTS_DIR = path.join(process.cwd(), 'scripts');
const TEST_FILE = path.join(SCRIPTS_DIR, `bench-read-mixed-${ROWS}.xlsx`);
const WORKER = path.join(SCRIPTS_DIR, 'worker-read.mjs');
const LIBRARIES = ['xlsxflow', 'exceljs-stream', 'exceljs', 'xlsx'];
const require = createRequire(path.join(SCRIPTS_DIR, 'competitors', 'package.json'));

async function prepareTestFile() {
  if (fs.existsSync(TEST_FILE)) return;
  console.log(`Generating ${ROWS} x ${COLS} mixed test file with ExcelJS...`);
  const ExcelJS = require('exceljs');
  const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: TEST_FILE, useSharedStrings: true });
  const ws = wb.addWorksheet('Data');
  ws.addRow(Array.from({ length: COLS }, (_, c) => `Col${c}`)).commit();
  for (let r = 0; r < ROWS; r++) {
    ws.addRow([
      r, `customer ${r % 1000}`, r * 1.25, r % 3 === 0, new Date(Date.UTC(2020, 0, 1 + (r % 2000))),
      `sku-${r}`, -r / 7, r % 2 ? 'open' : 'closed', r * 1000, `note ${r % 50}`,
    ]).commit();
  }
  await wb.commit();
}

const WORKER_CODE = `
import { createRequire } from 'module';
const require = createRequire(${JSON.stringify(path.join(SCRIPTS_DIR, 'competitors', 'package.json'))});
const [lib, file] = process.argv.slice(2);
const t0 = performance.now();
let rows = 0, cells = 0;
if (lib === 'xlsxflow') {
  const { SheetReader, createFileReader } = await import(${JSON.stringify(new URL('../dist/index.mjs', 'file:///' + SCRIPTS_DIR.replace(/\\\\/g, '/') + '/').href)});
  for await (const row of await new SheetReader().parse(await createFileReader(file))) { rows++; cells += row.cells.length; }
} else if (lib === 'exceljs-stream') {
  const ExcelJS = require('exceljs');
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(file, { sharedStrings: 'cache', worksheets: 'emit' });
  for await (const ws of reader) for await (const row of ws) { rows++; cells += row.values.length - 1; }
} else if (lib === 'exceljs') {
  const ExcelJS = require('exceljs');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  wb.worksheets[0].eachRow(row => { rows++; cells += row.values.length - 1; });
} else if (lib === 'xlsx') {
  const XLSX = require('xlsx');
  const wb = XLSX.readFile(file, { cellDates: true });
  const data = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
  rows = data.length; for (const r of data) cells += r.length;
}
const ms = performance.now() - t0;
console.log(JSON.stringify({ ms, rows, cells, peakRssMb: process.resourceUsage().maxRSS / 1024 }));
`;

async function run() {
  await prepareTestFile();
  fs.writeFileSync(WORKER, WORKER_CODE);
  const sizeMb = fs.statSync(TEST_FILE).size / 1024 / 1024;
  console.log(`File: ${path.basename(TEST_FILE)} (${sizeMb.toFixed(1)} MB), Node ${process.version}\n`);
  console.log('| Library | Read time | Rows | Cells | Peak RSS |');
  console.log('|---|---|---|---|---|');
  try {
    for (const lib of LIBRARIES) {
      const p = spawnSync(process.execPath, ['--max-old-space-size=4096', WORKER, lib, TEST_FILE], { encoding: 'utf-8' });
      const line = p.stdout.trim().split('\n').pop() ?? '';
      if (p.status !== 0 || !line.startsWith('{')) {
        console.log(`| ${lib} | failed: ${(p.stderr || '').trim().split('\n').pop()} | | | |`);
        continue;
      }
      const r = JSON.parse(line);
      console.log(`| ${lib} | ${Math.round(r.ms).toLocaleString('en-US')} ms | ${r.rows.toLocaleString('en-US')} | ${r.cells.toLocaleString('en-US')} | ${Math.round(r.peakRssMb)} MB |`);
    }
  } finally {
    fs.unlinkSync(WORKER);
  }
}

run().catch(e => { console.error(e); process.exit(1); });
