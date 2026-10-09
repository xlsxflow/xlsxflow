import { describe, it, expect } from 'vitest';
import fc from 'fast-check';
import { SheetReader, SheetWriter } from '../src/core/index';
import { createBlobReader } from '../src/core/random-access';

// Any grid of strings, numbers, booleans and blanks reads back exactly as it was written.
// fast-check shrinks a failure to the smallest grid that still differs.
const cell = fc.oneof(
  fc.string({ maxLength: 30, unit: 'grapheme' }),
  fc.double({ noNaN: true, noDefaultInfinity: true }),
  fc.integer(),
  fc.boolean(),
  fc.constant(null),
);

async function roundTrip(grid: unknown[][], sharedStrings: boolean) {
  const bytes = new Uint8Array(await new Response(new SheetWriter({ sharedStrings }).addSheet('S', grid as never).write()).arrayBuffer());
  const rows: { rowNumber: number; cells: unknown[] }[] = [];
  for await (const r of await new SheetReader().parse(createBlobReader(new Blob([bytes])))) rows.push(r);
  return rows;
}

describe('write then read', () => {
  it('returns every value as written', async () => {
    await fc.assert(fc.asyncProperty(fc.array(fc.array(cell, { maxLength: 8 }), { maxLength: 12 }), fc.boolean(), async (grid, shared) => {
      const rows = await roundTrip(grid, shared);
      const got = new Map(rows.map((r) => [r.rowNumber, r.cells]));
      grid.forEach((want, i) => {
        const cells = got.get(i + 1) ?? [];
        want.forEach((v, c) => expect(cells[c] ?? null).toEqual(Object.is(v, -0) ? 0 : v));
      });
    }), { numRuns: 200, seed: 42 });
  }, 120_000);
});
