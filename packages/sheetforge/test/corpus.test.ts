import { describe, it, expect } from 'vitest';
import { readdirSync, readFileSync, existsSync } from 'fs';
import path from 'path';
import { SheetReader } from '../src/core/index';
import { createBlobReader } from '../src/core/random-access';

// Real-world corpus: files written by other producers, compared against expected values
// (openpyxl oracle, hand-written, or source data). Regenerate with scripts/fixtures/*.
const FIX = path.join(__dirname, 'fixtures');

interface Expected { source: string; sheets: { name: string; rows: [number, unknown[]][] }[] }

async function readSheet(file: string, sheetName?: string) {
  const blob = new Blob([readFileSync(path.join(FIX, file))]);
  const rows: [number, unknown[]][] = [];
  for await (const r of await new SheetReader().parse(createBlobReader(blob), sheetName ? { sheetName } : undefined)) {
    const cells = [...r.cells];
    while (cells.length && cells[cells.length - 1] === null) cells.pop();
    if (cells.length) rows.push([r.rowNumber, cells]);
  }
  return rows;
}

const files = readdirSync(FIX).filter(f => f.endsWith('.xlsx') && existsSync(path.join(FIX, f.replace(/\.xlsx$/, '.expected.json'))));

describe('reader corpus', () => {
  it('has fixtures', () => expect(files.length).toBeGreaterThan(10));

  for (const file of files) {
    const expected: Expected = JSON.parse(readFileSync(path.join(FIX, file.replace(/\.xlsx$/, '.expected.json')), 'utf-8'));
    describe(file, () => {
      for (const sheet of expected.sheets) {
        it(`sheet "${sheet.name}" matches ${expected.source}`, async () => {
          expect(await readSheet(file, sheet.name)).toEqual(sheet.rows);
        });
      }
      it('defaults to the first tab', async () => {
        expect(await readSheet(file)).toEqual(expected.sheets[0].rows);
      });
    });
  }
});
