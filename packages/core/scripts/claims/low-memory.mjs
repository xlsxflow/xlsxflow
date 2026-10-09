// README claim: writing and reading stream in near-constant memory. CI runs this with a 32 MB heap, smaller
// than the 34 MB file it writes and reads back, so holding the workbook in memory would crash it (pnpm build first):
//   node --max-old-space-size=32 scripts/claims/low-memory.mjs
import { createWriteStream } from 'node:fs';
import { stat, unlink } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { Writable } from 'node:stream';
import { SheetWriter, XlsxFlow } from '../../dist/index.mjs';

const ROWS = 1_000_000, COLS = 10;
const file = join(tmpdir(), `xlsxflow-low-memory-${process.pid}.xlsx`);
async function* rows() { for (let r = 0; r < ROWS; r++) yield Array.from({ length: COLS }, (_, c) => c === 0 ? `row ${r}` : r * COLS + c); }

let peak = 0;
const sample = setInterval(() => { peak = Math.max(peak, process.memoryUsage().heapUsed); }, 20);
const mb = (n) => (n / 2 ** 20).toFixed(0) + ' MB';
try {
  await new SheetWriter().addSheet('Data', rows()).write().pipeTo(Writable.toWeb(createWriteStream(file)));
  let n = 0, last;
  for await (const r of await XlsxFlow.readFile(file)) { n++; last = r.cells; }
  const want = [`row ${ROWS - 1}`, ...Array.from({ length: COLS - 1 }, (_, c) => (ROWS - 1) * COLS + c + 1)];
  if (n !== ROWS || JSON.stringify(last) !== JSON.stringify(want)) throw new Error(`read back ${n} rows, last ${JSON.stringify(last)}`);
  const { size } = await stat(file);
  console.log(`${(ROWS * COLS).toLocaleString('en')} cells written and read back: file ${mb(size)}, peak heap ${mb(peak)}`);
} catch (e) {
  console.log(`FAIL low memory: ${e.message}`);
  process.exitCode = 1;
} finally {
  clearInterval(sample);
  await unlink(file).catch(() => {});
}
