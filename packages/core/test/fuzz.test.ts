import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { readdirSync, readFileSync, statSync } from 'fs';
import path from 'path';
import { SheetReader } from '../src/core/index';
import { createBlobReader } from '../src/core/random-access';
import { createXmlStreamParser } from '../src/core/xml-stream';
import { parseWorksheet } from '../src/core/worksheet-parser';

// Deterministic PRNG so failures are reproducible from the seed in the test name
function rng(seed: number) {
  return () => ((seed = Math.imul(seed ^ (seed >>> 15), 0x2c1b3c6d) + 0x9e3779b9 | 0) >>> 0) / 2 ** 32;
}

async function drainTokens(bytes: Uint8Array, chunk: number) {
  let n = 0;
  const stream = new ReadableStream<Uint8Array>({
    start(c) { for (let i = 0; i < bytes.length; i += chunk) c.enqueue(bytes.subarray(i, i + chunk)); c.close(); }
  });
  const reader = stream.pipeThrough(createXmlStreamParser()).getReader();
  while (!(await reader.read()).done) n++;
  return n;
}

// Reading a corrupted file must end in rows or an Error: never a hang, crash or non-Error throw.
async function readAll(bytes: Uint8Array): Promise<'ok' | 'error'> {
  try {
    const result = await new SheetReader().parse(createBlobReader(new Blob([bytes])));
    for await (const _ of result) { /* drain */ }
    await result.getMetadata();
    return 'ok';
  } catch (e) {
    if (!(e instanceof Error)) throw new Error(`non-Error thrown: ${String(e)}`);
    return 'error';
  }
}

const withTimeout = <T>(p: Promise<T>, ms: number, what: string) =>
  Promise.race([p, new Promise<never>((_, rej) => setTimeout(() => rej(new Error(`TIMEOUT (${ms} ms): ${what}`)), ms))]);

describe('xml tokenizer is linear on hostile input', () => {
  const MB = 1 << 20;
  const cases: [string, string][] = [
    ['unterminated quoted attribute', '<a b="' + 'x'.repeat(4 * MB)],
    ['giant tag', '<a ' + 'b="1" '.repeat(700_000) + '/>'],
    ['giant text node', '<t>' + 'y'.repeat(4 * MB) + '</t>'],
    ['unterminated CDATA', '<t><![CDATA[' + 'z'.repeat(4 * MB)],
    ['unterminated comment', '<!--' + 'c'.repeat(4 * MB)],
  ];
  for (const [name, xml] of cases) {
    it(name, async () => {
      const t0 = performance.now();
      await withTimeout(drainTokens(new TextEncoder().encode(xml), 4096), 10_000, name);
      // Linear takes well under a second; quadratic would take minutes. The margin covers slow CI and coverage runs
      expect(performance.now() - t0).toBeLessThan(8000);
    }, 15_000);
  }
});

async function parseXml(xml: string): Promise<'ok' | 'error'> {
  const bytes = new TextEncoder().encode(xml);
  const stream = new ReadableStream<Uint8Array>({
    start(c) { for (let i = 0; i < bytes.length; i += 97) c.enqueue(bytes.subarray(i, i + 97)); c.close(); }
  });
  try {
    const result = parseWorksheet(stream.pipeThrough(createXmlStreamParser()), new Map([[0, 's0']]), new Map([[1, 14]]));
    for await (const _ of result) { /* drain */ }
    await result.getMetadata();
    return 'ok';
  } catch (e) {
    if (!(e instanceof Error)) throw new Error(`non-Error thrown: ${String(e)}`);
    return 'error';
  }
}

describe('worksheet parser resource limits', () => {
  it('rejects a cell reference beyond column XFD instead of padding billions of nulls', async () => {
    const xml = '<worksheet><sheetData><row r="1"><c r="ZZZZZZZZZZ1"><v>1</v></c></row></sheetData></worksheet>';
    expect(await withTimeout(parseXml(xml), 2000, 'huge column')).toBe('error');
  });

  it('clamps hidden column ranges to the sheet width', async () => {
    const xml = '<worksheet><cols><col min="1" max="999999999" hidden="1"/></cols><sheetData/></worksheet>';
    const bytes = new TextEncoder().encode(xml);
    const result = parseWorksheet(new Blob([bytes]).stream().pipeThrough(createXmlStreamParser()), new Map(), new Map());
    for await (const _ of result) { /* drain */ }
    expect((await result.getMetadata()).hiddenCols.length).toBe(16384);
  });

  it('places out-of-order cells at their referenced column', async () => {
    const xml = '<worksheet><sheetData><row r="1"><c r="C1"><v>3</v></c><c r="A1"><v>1</v></c></row></sheetData></worksheet>';
    const rows = [];
    for await (const r of parseWorksheet(new Blob([xml]).stream().pipeThrough(createXmlStreamParser()), new Map(), new Map())) rows.push(r.cells);
    expect(rows).toEqual([[1, null, 3]]);
  });
});

describe('worksheet parser survives hostile XML', () => {
  const FRAGMENTS = ['<', '>', '"', "'", '&', '&#x110000;', '&#99999999999;', '&bogus;', '<c r="ZZZZZZZZZZZ1">', '<c r="">', '<c r="1">',
    '<row>', '</row>', '<row r="-5">', '<row r="abc">', '<v>', '</v>', '<v>1e999</v>', '<c t="s"><v>999999</v></c>', '<c s="1"><v>-1e308</v></c>',
    '<![CDATA[', ']]>', '<!--', '-->', '<?pi', '<!DOCTYPE x [<!ENTITY a "b">]>', '<f t="shared" si="x"/>', '</sheetData>', '<sheetData>',
    '<mergeCell ref="A1:ZZZZ99999999"/>', '<col min="-1" max="1e12" hidden="1"/>', '<pane state="frozen" ySplit="NaN"/>', '<x:c xmlns:x="u">', '�', '\u0000'];
  const BASE = '<?xml version="1.0"?><worksheet xmlns="m"><cols><col min="2" max="3" hidden="1"/></cols><sheetData>' +
    Array.from({ length: 30 }, (_, r) => `<row r="${r + 1}"><c r="A${r + 1}" t="s"><v>0</v></c><c r="B${r + 1}" s="1"><v>${45000 + r}</v></c>` +
      `<c r="C${r + 1}" t="inlineStr"><is><t>a &amp; b</t></is></c><c r="D${r + 1}"><f>A1</f><v>${r}</v></c></row>`).join('') +
    '</sheetData><mergeCells><mergeCell ref="A1:B2"/></mergeCells></worksheet>';
  const ITER = Number(process.env.FUZZ_ITER ?? 25) * 8;

  // fast-check generates the edits; a failure is shrunk to the smallest edit list that still breaks the parser
  const edit = fc.oneof(
    fc.record({ at: fc.nat(), insert: fc.oneof(fc.constantFrom(...FRAGMENTS), fc.string({ maxLength: 12 })) }),
    fc.record({ at: fc.nat(), remove: fc.integer({ min: 1, max: 40 }) }),
    fc.record({ at: fc.nat(), cut: fc.constant(true) }),
  );
  const apply = (edits: { at: number; insert?: string; remove?: number; cut?: boolean }[]) => edits.reduce((xml, e) => {
    const at = e.at % (xml.length + 1);
    return e.insert !== undefined ? xml.slice(0, at) + e.insert + xml.slice(at) : e.cut ? xml.slice(0, at) : xml.slice(0, at) + xml.slice(at + e.remove!);
  }, BASE);

  it(`${ITER} mutated worksheets end in rows or an Error`, async () => {
    await fc.assert(fc.asyncProperty(fc.array(edit, { minLength: 1, maxLength: 6 }), async (edits) => {
      const xml = apply(edits);
      await withTimeout(parseXml(xml), 3000, xml.slice(0, 200));
    }), { numRuns: ITER, seed: 1234 + ITER });
  }, 300_000);
});

describe('reader survives corrupted files', () => {
  const FIX = path.join(__dirname, 'fixtures');
  const small = readdirSync(FIX).filter(f => f.endsWith('.xlsx') && statSync(path.join(FIX, f)).size < 64 * 1024);
  const ITER = Number(process.env.FUZZ_ITER ?? 25);

  for (const file of small) {
    it(`${file}: ${ITER} mutations`, async () => {
      const original = new Uint8Array(readFileSync(path.join(FIX, file)));
      const rand = rng(file.length * 7919 + ITER);
      const outcomes = { ok: 0, error: 0 };
      for (let i = 0; i < ITER; i++) {
        const b = original.slice();
        const kind = i % 4;
        if (kind === 0) {
          // flip a few random bytes
          for (let k = 0; k < 1 + rand() * 8; k++) b[Math.floor(rand() * b.length)] ^= 1 << Math.floor(rand() * 8);
        } else if (kind === 1) {
          // corrupt the central directory / EOCD region (sizes, offsets, counts)
          const from = Math.max(0, b.length - 200);
          for (let k = 0; k < 4; k++) b[from + Math.floor(rand() * (b.length - from))] = Math.floor(rand() * 256);
        } else if (kind === 2) {
          // truncate, keeping the EOCD so the archive still "opens"
          const cut = Math.floor(rand() * (b.length - 22));
          const eocd = b.subarray(b.length - 22);
          const t = new Uint8Array(cut + 22); t.set(b.subarray(0, cut)); t.set(eocd, cut);
          outcomes[await withTimeout(readAll(t), 5000, `${file} #${i} truncate@${cut}`)]++;
          continue;
        } else {
          // overwrite a run with 0xFF (huge sizes / offsets / counts)
          const at = Math.floor(rand() * b.length);
          b.fill(0xff, at, at + 4);
        }
        outcomes[await withTimeout(readAll(b), 5000, `${file} #${i} kind=${kind}`)]++;
      }
      expect(outcomes.ok + outcomes.error).toBe(ITER);
    }, 120_000);
  }
});
