import * as fs from 'fs';
import * as path from 'path';
import { spawn } from 'child_process';

interface BenchResult {
  name: string;
  rows: number;
  writeMs?: number;
  writeMb?: number;
  readMs?: number;
  memMb?: number;
  error?: string;
  oom?: boolean;
}

const ROWS_FULL = 100_000;
const ROWS_FALLBACK = 10_000;
const COLS = 10;

const LIBRARIES = [
  'xlsxflow',
  'xlsx',
  'exceljs',
  'excel4node',
  'write-excel-file',
  'xlsx-populate',
  'msexcel-builder'
];

const SCRIPTS_DIR = path.join(process.cwd(), 'scripts');
const WORKER_SCRIPT = path.join(SCRIPTS_DIR, 'worker.js');

async function writeWorkerScript() {
  const code = `
import * as fs from 'fs';
import * as path from 'path';
import { createRequire } from 'module';
import { Readable } from 'stream';
const require = createRequire(import.meta.url);

// --- Worker ---
const lib = process.argv[2];
const rows = parseInt(process.argv[3], 10);
const filepath = path.join(process.cwd(), \`bench-\${lib}-\${rows}.xlsx\`);

function generateRow(r) {
  const row = [];
  for (let c = 0; c < 10; c++) row.push(r * 10 + c);
  return row;
}

async function run() {
  let writeMs = -1;
  let writeMb = -1;
  let readMs = -1;
  let memMb = -1;

  if (global.gc) global.gc();
  const startMem = process.memoryUsage().heapUsed;

  const tStart = performance.now();

  if (lib === 'xlsxflow') {
    const { SheetWriter } = await import('../dist/index.mjs');
    const writer = new SheetWriter();
    async function* gen() { for (let r=0; r<rows; r++) yield generateRow(r); }
    const stream = await writer.write(gen());
    const out = fs.createWriteStream(filepath);
    const reader = stream.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      out.write(value);
    }
    out.end();
    await new Promise(r => out.on('finish', () => r()));
  } else if (lib === 'xlsx') {
    const XLSX = require('./competitors/node_modules/xlsx');
    const data = [];
    data.push(Array.from({length: 10}, (_, i) => \`Col\${i}\`));
    for (let r=0; r<rows; r++) data.push(generateRow(r));
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filepath);
  } else if (lib === 'exceljs') {
    const ExcelJS = require('./competitors/node_modules/exceljs');
    const wb = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: filepath });
    const ws = wb.addWorksheet('Sheet1');
    for (let r=0; r<rows; r++) ws.addRow(generateRow(r)).commit();
    await wb.commit();
  } else if (lib === 'excel4node') {
    const xl = require('./competitors/node_modules/excel4node');
    const wb = new xl.Workbook();
    const ws = wb.addWorksheet('Sheet1');
    for (let r=0; r<rows; r++) {
      const row = generateRow(r);
      for (let c=0; c<10; c++) ws.cell(r+1, c+1).number(row[c]);
    }
    await new Promise((resolve, reject) => wb.write(filepath, (err) => err ? reject(err) : resolve(null)));
  } else if (lib === 'write-excel-file') {
    const writeXlsxFile = require('./competitors/node_modules/write-excel-file/node');
    const data = [];
    for (let r=0; r<rows; r++) {
      data.push(generateRow(r).map((v) => ({ type: Number, value: v })));
    }
    await writeXlsxFile(data, { filePath: filepath });
  } else if (lib === 'xlsx-populate') {
    const XlsxPopulate = require('./competitors/node_modules/xlsx-populate');
    const wb = await XlsxPopulate.fromBlankAsync();
    const ws = wb.sheet(0);
    for (let r=0; r<rows; r++) {
      const row = generateRow(r);
      for (let c=0; c<10; c++) ws.cell(r+1, c+1).value(row[c]);
    }
    await wb.toFileAsync(filepath);
  } else if (lib === 'msexcel-builder') {
    const builder = require('./competitors/node_modules/msexcel-builder');
    const wb = builder.createWorkbook(process.cwd(), path.basename(filepath));
    const ws = wb.createSheet('Sheet1', 10, rows);
    for (let r=0; r<rows; r++) {
      const row = generateRow(r);
      for (let c=0; c<10; c++) ws.set(c+1, r+1, row[c]);
    }
    await new Promise((resolve, reject) => wb.save((err) => err ? reject(err) : resolve(null)));
  } else {
    throw new Error('Unsupported lib: ' + lib);
  }

  writeMs = performance.now() - tStart;
  writeMb = fs.statSync(filepath).size / 1024 / 1024;

  if (global.gc) global.gc();
  const peakMem = process.memoryUsage().heapUsed;
  memMb = (peakMem - startMem) / 1024 / 1024;

  console.log(JSON.stringify({ writeMs, writeMb, memMb }));
  process.exit(0);
}
run().catch(e => {
  console.error(e.message);
  process.exit(1);
});
`;
  fs.writeFileSync(WORKER_SCRIPT, code);
}

async function runWorker(lib: string, rows: number): Promise<BenchResult> {
  return new Promise((resolve) => {
    // Increase heap to 4GB to give competitors a fighting chance!
    const p = spawn('node', ['--expose-gc', '--max-old-space-size=4096', WORKER_SCRIPT, lib, rows.toString()], {
      cwd: SCRIPTS_DIR,
      shell: true,
      env: process.env
    });
    
    let out = '';
    let err = '';
    p.stdout.on('data', (d: any) => out += d.toString());
    p.stderr.on('data', (d: any) => err += d.toString());

    p.on('close', (code: number | null) => {
      if (code !== 0 || err.includes('FATAL ERROR') || err.includes('heap out of memory')) {
        if (!err.includes('FATAL ERROR') && !err.includes('heap out of memory')) {
          console.error(`Worker error (${lib}):`, err.slice(0, 500));
        }
        resolve({ name: lib, rows, oom: true, error: 'OOM / Crash' });
        return;
      }
      try {
        const parsed = JSON.parse(out.trim().split('\\n').pop()!);
        resolve({ name: lib, rows, ...parsed });
      } catch (e) {
        resolve({ name: lib, rows, error: 'Parse error: ' + err.slice(0, 100) });
      }
    });
  });
}

async function run() {
  console.log('================================================================');
  console.log('  Library Benchmark Suite (1M Cells & 100k Cells)');
  console.log('================================================================');
  
  await writeWorkerScript();

  const results: BenchResult[] = [];

  for (const lib of LIBRARIES) {
    console.log(`Testing ${lib} (100,000 rows, 1M cells)...`);
    let res = await runWorker(lib, ROWS_FULL);
    
    if (res.oom) {
      console.log(`  -> OOM crashed on 100k rows! Falling back to 10k rows...`);
      const fallback = await runWorker(lib, ROWS_FALLBACK);
      fallback.error = '100k OOM (tested 10k)';
      results.push(fallback);
    } else {
      console.log(`  -> Success: ${res.writeMs?.toFixed(0)} ms`);
      results.push(res);
    }
  }

  console.log('\n================================================================');
  console.log('  Summary Table');
  console.log('  ┌──────────────────────┬─────────┬────────────┬────────────┬──────────┐');
  console.log('  │ Library              │ Rows    │ Write (ms) │ File (MB)  │ Mem (MB) │');
  console.log('  ├──────────────────────┼─────────┼────────────┼────────────┼──────────┤');
  for (const r of results) {
    const n = r.name.padEnd(20);
    const rows = r.rows.toLocaleString().padEnd(7);
    const w = r.error && !r.writeMs ? r.error.padEnd(10) : r.writeMs?.toFixed(0).padStart(10);
    const f = r.writeMb ? r.writeMb.toFixed(1).padStart(10) : '      N/A ';
    const m = r.memMb ? `+${r.memMb.toFixed(0)}`.padStart(8) : '   N/A  ';
    console.log(`  │ ${n} │ ${rows} │ ${w} │ ${f} │ ${m} │`);
  }
  console.log('  └──────────────────────┴─────────┴────────────┴────────────┴──────────┘');
  
  if (fs.existsSync(WORKER_SCRIPT)) fs.unlinkSync(WORKER_SCRIPT);
}

run().catch(console.error);
