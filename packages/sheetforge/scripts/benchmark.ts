/// <reference types="node" />
import { SheetWriter } from '../src/core/writer.js';
import { SheetReader } from '../src/core/index.js';
import * as fs from 'fs';
import * as path from 'path';
import { Readable } from 'stream';

interface ScenarioResult {
  name: string;
  writeMb: number;
  writeMs: number;
  readMs: number;
  memMb: number;
  rows: number;
  cells: number;
}

async function runScenario(
  name: string,
  generateRow: (r: number) => (string | number)[],
  ROWS = 100_000,
  COLS = 10,
): Promise<ScenarioResult> {
  const filepath = path.join(process.cwd(), `benchmark-${name.replace(/\s+/g, '-')}.xlsx`);

  // --- WRITE ---
  const writer = new SheetWriter();
  const rows: (string | number)[][] = [];
  for (let r = 0; r < ROWS; r++) {
    rows.push(generateRow(r));
  }

  if (global.gc) global.gc();
  const startWrite = performance.now();
  const buffer = await writer.write(rows as any);
  fs.writeFileSync(filepath, buffer);
  const writeMs = performance.now() - startWrite;
  const writeMb = buffer.length / 1024 / 1024;

  // --- READ ---
  if (global.gc) global.gc();
  const startMem = process.memoryUsage().heapUsed;

  const reader = new SheetReader();
  const fileStream = fs.createReadStream(filepath);
  const webStream = Readable.toWeb(fileStream);

  let rowCount = 0;
  let cellCount = 0;
  let peakMem = startMem;

  const startRead = performance.now();
  const iterator = await reader.parse(webStream as any);

  try {
    for await (const row of iterator) {
      rowCount++;
      cellCount += row.length;
      const cur = process.memoryUsage().heapUsed;
      if (cur > peakMem) peakMem = cur;
    }
  } catch (e) {
    console.error(`  Error during ${name} iteration:`, e);
  }

  const readMs = performance.now() - startRead;
  const memMb = (peakMem - startMem) / 1024 / 1024;

  fs.unlinkSync(filepath);

  return { name, writeMb, writeMs, readMs, memMb, rows: rowCount, cells: cellCount };
}

function printResult(r: ScenarioResult) {
  console.log(`\n  📊 ${r.name}`);
  console.log(`     Write:  ${r.writeMs.toFixed(0)} ms  |  File: ${r.writeMb.toFixed(1)} MB`);
  console.log(`     Read:   ${r.readMs.toFixed(0)} ms  |  Rows: ${r.rows.toLocaleString()} (${r.cells.toLocaleString()} cells)`);
  console.log(`     Memory: ${r.memMb >= 0 ? '+' : ''}${r.memMb.toFixed(2)} MB peak heap delta`);
}

async function runBenchmark() {
  console.log('================================================================');
  console.log('  SheetForge Benchmark Suite — 1 Million Cells');
  console.log(`  Node: ${process.version}`);
  console.log('================================================================');

  // 1. 1M Numbers
  const numResult = await runScenario(
    '1M Numbers',
    (r) => Array.from({ length: 10 }, (_, c) => r * 10 + c),
  );
  printResult(numResult);

  // 2. 1M Duplicate strings (high SST reuse)
  const dupStrings = ['Apple', 'Banana', 'Cherry', 'Date', 'Elderberry', 'Fig', 'Grape', 'Honeydew', 'Kiwi', 'Lemon'];
  const dupResult = await runScenario(
    '1M Duplicate Strings',
    (_r) => Array.from({ length: 10 }, (_, c) => dupStrings[c % dupStrings.length]),
  );
  printResult(dupResult);

  // 3. 1M Unique strings (worst-case SST)
  const uniqResult = await runScenario(
    '1M Unique Strings',
    (r) => Array.from({ length: 10 }, (_, c) => `val_r${r}_c${c}_${Math.random().toString(36).slice(2, 6)}`),
  );
  printResult(uniqResult);

  console.log('\n================================================================');
  console.log('  Summary Table');
  console.log('  ┌────────────────────────┬────────────┬────────────┬──────────┬──────────┐');
  console.log('  │ Scenario               │ Write (ms) │ File (MB)  │ Read (ms)│ Mem (MB) │');
  console.log('  ├────────────────────────┼────────────┼────────────┼──────────┼──────────┤');
  for (const r of [numResult, dupResult, uniqResult]) {
    const n = r.name.padEnd(22);
    const w = r.writeMs.toFixed(0).padStart(10);
    const f = r.writeMb.toFixed(1).padStart(10);
    const rd = r.readMs.toFixed(0).padStart(8);
    const m = `${r.memMb >= 0 ? '+' : ''}${r.memMb.toFixed(2)}`.padStart(8);
    console.log(`  │ ${n} │ ${w} │ ${f} │ ${rd} │ ${m} │`);
  }
  console.log('  └────────────────────────┴────────────┴────────────┴──────────┴──────────┘');
  console.log('');
}

runBenchmark().catch(console.error);
