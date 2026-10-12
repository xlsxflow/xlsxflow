import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { SheetReader } from '../src/core/index';
import { createBlobReader, type RandomAccessReader } from '../src/core/random-access';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import type { SheetMetadata } from '../src/core/worksheet-parser';

// One workbook saved by desktop Excel as .xlsx, .xls and .ods, and one saved by LibreOffice as .ods
// with frozen panes. Regenerate with scripts/fixtures/make-excel-formats.ps1 and make-libreoffice-*.
const open = (name: string) => createBlobReader(new Blob([readFileSync(new URL(`./fixtures/${name}`, import.meta.url))]));

async function load(name: string | RandomAccessReader) {
  const reader = typeof name === 'string' ? open(name) : name;
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

    // Columns D, M and P, and the hidden S; Excel's .ods export rounds the widths a little differently
    const widths = async (name: string) => [...((await load(name)).sheets.Data.meta as SheetMetadata).columnWidths.entries()].filter(([, w]) => w !== undefined);
    const excel = [[3, 10.33203125], [12, 18.109375], [15, 9.44140625], [18, 0]];
    expect(await widths('excel-made.xlsx')).toEqual(excel);
    expect(await widths('excel-made.xls')).toEqual(excel);
    const ods = await widths('excel-made.ods');
    expect(ods.map(([c]) => c)).toEqual(excel.map(([c]) => c));
    ods.forEach(([, w], i) => expect(Math.abs(w - excel[i][1])).toBeLessThan(0.5));
  });

  it('reads LibreOffice\'s .ods, frozen panes included', async () => {
    const { info, sheets } = await load('libreoffice-made.ods');
    expect(info.sheets.map(s => s.state)).toEqual(['visible', 'visible', 'hidden']);
    expect(info.properties).toMatchObject({ title: 'Fixture title', creator: 'Fixture author' });
    expect(sheets.Data.meta).toMatchObject({ columnWidths: Object.assign([], { 16: 11.48 }) });
    expect(info.definedNames).toEqual([{ name: 'Answer', ref: 'Data!$B$1' }, { name: 'Numbers', ref: 'Data!$C$10:$C$3009' }]);
    expect(sheets.Data.cells.length).toBe(3004);
    expect(sheets.Data.meta).toMatchObject({ freezePanes: { row: 2, col: 0 } });
    expect(Object.values(sheets).find(s => 'freezePanes' in s.meta && s !== sheets.Data)?.meta).toMatchObject({ freezePanes: { row: 1, col: 0 } });
  });

  it('reads a flat .fods the same as the .ods it was made from', async () => {
    // Flat ODF is the zip's XML parts inside one office:document, indented between elements as LibreOffice saves it
    const zip = new ZipRandomAccessParser(open('libreoffice-made.ods'));
    await zip.parseCentralDirectory();
    const part = async (name: string) => {
      const xml = await new Response(await zip.extractStream(name)).text();
      return xml.slice(xml.indexOf('>', xml.indexOf('<office:document-')) + 1, xml.lastIndexOf('</office:document-'));
    };
    const parts = await Promise.all(['meta.xml', 'settings.xml', 'styles.xml', 'content.xml'].map(part));
    const fods = '<?xml version="1.0" encoding="UTF-8"?>\n<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
      + ' office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">\n'
      + parts.join('\n').replace(/<\/text:p>/g, '</text:p>\n     ').replace(/<table:table-(row|cell)\b/g, '\n    $&')
      + '\n</office:document>\n';
    expect(fods).toContain('</text:p>\n     </table:table-cell>');

    const ods = await load('libreoffice-made.ods');
    expect(await load(createBlobReader(new Blob([fods])))).toEqual(ods);
  });
});
