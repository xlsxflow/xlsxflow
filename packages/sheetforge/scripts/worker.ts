
import * as fs from 'fs';
import * as path from 'path';
import { createRequire } from 'module';
import { Readable } from 'stream';
const require = createRequire(import.meta.url);

// --- Worker ---
const lib = process.argv[2];
const rows = parseInt(process.argv[3], 10);
const filepath = path.join(process.cwd(), `bench-${lib}-${rows}.xlsx`);

function generateRow(r: number) {
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

  if (lib === 'sheetforge') {
    const { SheetWriter } = await import('../src/core/writer.ts');
    const writer = new SheetWriter();
    async function* gen() { for (let r=0; r<rows; r++) yield generateRow(r); }
    const stream = await writer.write(gen() as any);
    const out = fs.createWriteStream(filepath);
    Readable.fromWeb(stream as any).pipe(out);
    await new Promise<void>(r => out.on('finish', () => r()));
  } else if (lib === 'xlsx') {
    const XLSX = require('./competitors/node_modules/xlsx');
    const data = [];
    data.push(Array.from({length: 10}, (_, i) => `Col${i}`));
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
    await new Promise((resolve, reject) => wb.write(filepath, (err: any) => err ? reject(err) : resolve(null)));
  } else if (lib === 'write-excel-file') {
    const writeXlsxFile = require('./competitors/node_modules/write-excel-file/node');
    const data = [];
    for (let r=0; r<rows; r++) {
      data.push(generateRow(r).map((v: any) => ({ type: Number, value: v })));
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
    await new Promise((resolve, reject) => wb.save((err: any) => err ? reject(err) : resolve(null)));
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
