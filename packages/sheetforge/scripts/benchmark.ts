/// <reference types="node" />
import { SheetWriter } from '../src/core/writer.js';
import { SheetReader } from '../src/core/index.js';
import * as fs from 'fs';
import * as path from 'path';
import { Readable } from 'stream';

async function runBenchmark() {
  console.log('--- SheetForge Benchmark (1 Million Cells) ---');
  console.log(`Node Environment: ${process.version}`);
  
  const COLS = 10;
  const ROWS = 100_000;
  const filepath = path.join(process.cwd(), 'benchmark-1m.xlsx');
  
  console.log(`\n1. Generating 1M cell workbook... (${ROWS}x${COLS})`);
  
  const writer = new SheetWriter();
  const rows = [];
  for (let r = 0; r < ROWS; r++) {
    const row = [];
    for (let c = 0; c < COLS; c++) {
      // Use numbers to represent standard data (avoiding huge shared string tables)
      row.push(r * c);
    }
    rows.push(row);
  }
  
  const startTimeWrite = performance.now();
  const buffer = writer.write(rows);
  fs.writeFileSync(filepath, buffer);
  const writeTime = performance.now() - startTimeWrite;
  
  console.log(`   Written ${buffer.length / 1024 / 1024 | 0} MB to disk.`);
  console.log(`   Time: ${writeTime.toFixed(0)} ms`);

  // Force GC if exposed, to measure clean parsing footprint
  if (global.gc) global.gc();

  console.log(`\n2. Parsing 1M cell workbook (Streaming)...`);
  
  const reader = new SheetReader();
  const fileStream = fs.createReadStream(filepath);
  const webStream = Readable.toWeb(fileStream);

  const startMem = process.memoryUsage().heapUsed;
  let maxMem = 0;
  
  const startTimeRead = performance.now();
  const iterator = await reader.parse(webStream as any);
  
  let rowCount = 0;
  let cellCount = 0;
  console.log('Starting iteration...');
  try {
    for await (const row of iterator) {
      rowCount++;
      cellCount += row.length;
      if (rowCount % 10000 === 0) console.log(`Read ${rowCount} rows...`);
    }
  } catch (e) {
    console.error('Error during iteration:', e);
  }
  console.log('Iteration finished!');
  
  const readTime = performance.now() - startTimeRead;
  
  console.log(`   Parsed ${rowCount.toLocaleString()} rows (${cellCount.toLocaleString()} cells).`);
  console.log(`   Time: ${readTime.toFixed(0)} ms`);
  console.log(`   Memory Footprint: ${((maxMem - startMem) / 1024 / 1024).toFixed(2)} MB`);

  fs.unlinkSync(filepath);
  console.log('\nDone.');
}

runBenchmark().catch(console.error);
