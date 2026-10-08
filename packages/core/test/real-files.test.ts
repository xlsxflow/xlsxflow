import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { SheetReader } from '../src/core/index';
import { createBlobReader } from '../src/core/random-access';

// One workbook saved by desktop Excel as .xlsx, .xls and .ods, and one saved by LibreOffice as .ods
// with frozen panes. Regenerate with scripts/fixtures/make-excel-formats.ps1 and make-libreoffice-*.
const open = (name: string) => createBlobReader(new Blob([readFileSync(new URL(`./fixtures/${name}`, import.meta.url))]));

async function load(name: string) {
  const reader = open(name);
  const info = await new SheetReader().readWorkbook(reader);
  const sheets: Record<string, { cells: unknown[][]; meta: object }> = {};
  for (const { name: sheetName } of info.sheets) {
    const result = await new SheetReader().parse(reader, { sheetName });
    const cells = [];
    for await (const row of result) {
      const c = [...row.cells];
      while (c.length && c[c.length - 1] === null) c.pop();
      if (c.length) cells.push([row.rowNumber, ...c]);
    }
    sheets[sheetName] = { cells, meta: await result.getMetadata() };
  }
  return { info, sheets };
}

describe('files saved by Excel and LibreOffice', () => {
  it('reads the same cells from Excel\'s .xlsx, .xls and .ods', async () => {
    const xlsx = await load('excel-made.xlsx');
    expect(xlsx.sheets.Data.cells.length).toBe(3004);
    expect(xlsx.sheets.Data.cells[0].slice(0, 7)).toEqual([1, 'Text', 42, 3.14159, '2026-10-08T00:00:00.000Z', true, '#DIV/0!']);
    expect(xlsx.sheets.Data.meta).toMatchObject({
      mergedCells: ['A3:C3'], hiddenRows: [2], hiddenCols: [19], freezePanes: { row: 2, col: 1, topLeftCell: 'B3' },
    });
    for (const other of ['excel-made.xls', 'excel-made.ods']) {
      const got = await load(other);
      expect(got.info.sheets).toEqual(xlsx.info.sheets);
      for (const name of Object.keys(xlsx.sheets)) expect(got.sheets[name].cells, `${other} ${name}`).toEqual(xlsx.sheets[name].cells);
      expect(got.sheets.Data.meta).toMatchObject({ mergedCells: ['A3:C3'], hiddenRows: [2], hiddenCols: [19] });
    }
    // Excel keeps frozen panes in .xls but drops them from .ods
    expect((await load('excel-made.xls')).sheets.Data.meta).toMatchObject({ freezePanes: { row: 2, col: 1, topLeftCell: 'B3' } });
  });

  it('reads LibreOffice\'s .ods, frozen panes included', async () => {
    const { info, sheets } = await load('libreoffice-made.ods');
    expect(info.sheets.map(s => s.state)).toEqual(['visible', 'visible', 'hidden']);
    expect(info.properties).toMatchObject({ title: 'Fixture title', creator: 'Fixture author' });
    expect(info.definedNames).toEqual([{ name: 'Answer', ref: 'Data!$B$1' }, { name: 'Numbers', ref: 'Data!$C$10:$C$3009' }]);
    expect(sheets.Data.cells.length).toBe(3004);
    expect(sheets.Data.meta).toMatchObject({ freezePanes: { row: 2, col: 0 } });
    expect(Object.values(sheets).find(s => 'freezePanes' in s.meta && s !== sheets.Data)?.meta).toMatchObject({ freezePanes: { row: 1, col: 0 } });
  });
});
