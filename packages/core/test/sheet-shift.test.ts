import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { SheetWriter } from '../src/core/writer';
import { SheetEditor } from '../src/core/editor';
import { SheetReader } from '../src/core/index';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { createBlobReader } from '../src/core/random-access';
import { AxisMap, mapArea, mapFormula, createShiftTransform } from '../src/core/sheet-shift';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());
const blob = (b: Uint8Array) => createBlobReader(new Blob([b]));
const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='), c => c.charCodeAt(0));

async function read(bytes: Uint8Array, sheetName?: string) {
  const rows = await new SheetReader().parse(blob(bytes), { sheetName, formulas: true });
  const out = new Map<number, { cells: unknown[]; formulas?: (string | undefined)[] }>();
  for await (const row of rows) out.set(row.rowNumber, { cells: row.cells, formulas: row.formulas });
  return out;
}

async function entry(bytes: Uint8Array, name: string) {
  const zip = new ZipRandomAccessParser(blob(bytes));
  await zip.parseCentralDirectory();
  return zip.has(name) ? new Response(await zip.extractStream(name)).text() : '';
}

describe('AxisMap', () => {
  it('moves rows, shrinks and grows ranges', () => {
    const ins = new AxisMap([{ at: 3, count: 2, insert: true }]);
    const insRows = { rows: ins };
    expect([1, 2, 3, 9].map(r => ins.map(r, 'single'))).toEqual([1, 2, 5, 11]);
    expect(mapArea('B2:B6', insRows)).toBe('B2:B8');
    expect(mapArea('A1:A2', insRows)).toBe('A1:A2');

    const del = new AxisMap([{ at: 3, count: 2, insert: false }]);
    const delRows = { rows: del };
    expect([2, 3, 4, 5].map(r => del.map(r, 'single'))).toEqual([2, null, null, 3]);
    expect(mapArea('B2:B6', delRows)).toBe('B2:B4');
    expect(mapArea('B3:B6', delRows)).toBe('B3:B4');
    expect(mapArea('B3:C4', delRows)).toBeNull();
    expect(mapArea('3:4', delRows)).toBeNull();
    expect(mapArea('2:9', delRows)).toBe('2:7');

    const maps = new Map([['data', delRows]]);
    expect(mapFormula('SUM(B2:B6)+B5+B3', 'Data', maps)).toBe('SUM(B2:B4)+B3+#REF!');
    expect(mapFormula("Data!B5+'data'!$B$6+Other!B5", 'Other', maps)).toBe("Data!B3+'data'!$B$4+Other!B5");
    expect(mapFormula('B5', 'Other', maps)).toBe('B5');
    // Conditional formats keep relative references counted from their range's first row
    expect(mapFormula('$A3>0', 'Data', maps, { row: 3, toRow: 3, col: 0, toCol: 0 })).toBe('$A3>0');
  });
});

function sample() {
  const w = new SheetWriter();
  w.addSheet('Data', [
    ['Name', 'Qty', { value: null, formula: 'SUM(B2:B6)' }],
    ['a', 1],
    ['b', 2],
    [{ value: 'c', comment: 'note on c' }, 3],
    ['d', 4],
    [{ value: 'e', hyperlink: 'https://example.com' }, 5],
  ], {
    mergeCells: ['A8:B8'], // below the table: Excel tables cannot hold merged cells
    conditionalFormats: [{ range: 'B2:B6', rule: { type: 'expression', formula: '$B2>2', style: { font: { bold: true } } } }],
    dataValidations: [{ sqref: 'B2:B6', type: 'whole', operator: 'greaterThan', formula1: '0' }],
    tables: [{ name: 'Items', ref: 'A1:B6' }],
    pageSetup: { printArea: 'A1:B6' },
    images: [{ data: PNG, range: 'D4:E6' }],
  });
  w.addSheet('Other', [[{ value: null, formula: 'Data!B4' }, { value: null, formula: 'SUM(Data!B2:B6)' }, { value: null, formula: 'B1' }]]);
  return bytesOf(w.write());
}

describe('SheetEditor.insertRows / deleteRows', () => {
  it('inserts rows and moves everything that points at them', async () => {
    const out = await bytesOf(new SheetEditor().insertRows('Data', 3, 2).edit(blob(await sample())));
    const data = await read(out, 'Data');
    expect([...data.keys()]).toEqual([1, 2, 5, 6, 7, 8]);
    expect(data.get(5)!.cells).toEqual(['b', 2]);
    expect(data.get(1)!.formulas?.[2]).toBe('SUM(B2:B8)');
    const other = await read(out, 'Other');
    expect(other.get(1)!.formulas).toEqual(['Data!B6', 'SUM(Data!B2:B8)', 'B1']);

    const sheet = await entry(out, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<mergeCell ref="A10:B10"/>');
    expect(sheet).toContain('sqref="B2:B8"');
    expect(sheet).toContain('<formula>$B2&gt;2</formula>');
    expect(sheet).toMatch(/<hyperlink ref="A8"/);
    expect(await entry(out, 'xl/comments1.xml')).toContain('ref="A6"');
    expect(await entry(out, 'xl/drawings/vmlDrawing1.vml')).toContain('<x:Row>5</x:Row>');
    expect(await entry(out, 'xl/tables/table1.xml')).toMatch(/<table [^>]*ref="A1:B8"/);
    expect(await entry(out, 'xl/workbook.xml')).toContain("'Data'!$A$1:$B$8");
    // The picture over D4:E6 (keeping its size) moves down two rows: 0-based rows 5 to 8
    const drawing = await entry(out, 'xl/drawings/drawing1.xml');
    expect(drawing).toMatch(/<xdr:from>[\s\S]*?<xdr:row>5<\/xdr:row>[\s\S]*?<xdr:to>[\s\S]*?<xdr:row>8<\/xdr:row>/);
    expect(await entry(out, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');
  });

  it('deletes rows, turning references to them into #REF! and shrinking ranges', async () => {
    const out = await bytesOf(new SheetEditor().deleteRows('Data', 4, 2).edit(blob(await sample())));
    const data = await read(out, 'Data');
    expect([...data.keys()]).toEqual([1, 2, 3, 4]);
    expect(data.get(4)!.cells).toEqual(['e', 5]);
    expect(data.get(1)!.formulas?.[2]).toBe('SUM(B2:B4)');
    expect((await read(out, 'Other')).get(1)!.formulas).toEqual(['Data!#REF!', 'SUM(Data!B2:B4)', 'B1']);

    const sheet = await entry(out, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<mergeCell ref="A6:B6"/>');
    expect(sheet).toContain('sqref="B2:B4"');
    expect(sheet).toMatch(/<hyperlink ref="A4"/);
    expect(await entry(out, 'xl/comments1.xml')).not.toContain('<comment ');
    expect(await entry(out, 'xl/drawings/vmlDrawing1.vml')).not.toContain('<x:Row>');
    expect(await entry(out, 'xl/tables/table1.xml')).toMatch(/<table [^>]*ref="A1:B4"/);
    expect(await entry(out, 'xl/workbook.xml')).toContain("'Data'!$A$1:$B$4");
  });

  it('runs several operations in order and combines with setCells', async () => {
    const out = await bytesOf(new SheetEditor()
      .insertRows('Data', 2)          // rows 2.. move down one
      .deleteRows('Data', 7)          // the old row 6 ('e')
      .setCells('Data', { A2: 'new' }) // addresses after the moves
      .edit(blob(await sample())));
    const data = await read(out, 'Data');
    expect([...data.keys()]).toEqual([1, 2, 3, 4, 5, 6]);
    expect(data.get(2)!.cells).toEqual(['new']);
    expect(data.get(6)!.cells).toEqual(['d', 4]);
    expect(data.get(1)!.formulas?.[2]).toBe('SUM(B3:B6)');
  });

  it('writes out shared formulas when rows move under them', async () => {
    // Report!C3 holds B3*2, shared into C4
    const tpl = new Uint8Array(readFileSync(new URL('./fixtures/openpyxl-template.xlsx', import.meta.url)));
    const out = await bytesOf(new SheetEditor().insertRows('Report', 4).edit(blob(tpl)));
    const rows = await read(out, 'Report');
    expect(rows.get(3)!.formulas?.[2]).toBe('B3*2');
    expect(rows.get(5)!.formulas?.[2]).toBe('B5*2');
    expect(await entry(out, 'xl/worksheets/sheet1.xml')).not.toContain('t="shared"');
  });

  it('refuses to delete a table header or leave a table empty', async () => {
    const src = await sample();
    await expect(bytesOf(new SheetEditor().deleteRows('Data', 1).edit(blob(src)))).rejects.toThrow(/header row of table "Items"/);
    await expect(bytesOf(new SheetEditor().deleteRows('Data', 2, 5).edit(blob(src)))).rejects.toThrow(/without data rows/);
    expect(() => new SheetEditor().insertRows('Data', 0)).toThrow(/Invalid rows/);
    await expect(bytesOf(new SheetEditor().insertRows('Nope', 1).edit(blob(src)))).rejects.toThrow(/not found/);
  });
});

describe('SheetEditor.insertColumns / deleteColumns', () => {
  it('inserts columns, growing the table it lands in', async () => {
    const out = await bytesOf(new SheetEditor().insertColumns('Data', 'B', 2).edit(blob(await sample())));
    const data = await read(out, 'Data');
    expect(data.get(1)!.cells.slice(0, 4)).toEqual(['Name', 'Column1', 'Column2', 'Qty']);
    expect(data.get(2)!.cells).toEqual(['a', null, null, 1]);
    expect(data.get(1)!.formulas?.[4]).toBe('SUM(D2:D6)');
    expect((await read(out, 'Other')).get(1)!.formulas).toEqual(['Data!D4', 'SUM(Data!D2:D6)', 'B1']);

    const sheet = await entry(out, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<mergeCell ref="A8:D8"/>');
    expect(sheet).toContain('sqref="D2:D6"');
    expect(sheet).toContain('<formula>$D2&gt;2</formula>');
    const table = await entry(out, 'xl/tables/table1.xml');
    expect(table).toMatch(/<table [^>]*ref="A1:D6"/);
    expect(table).toMatch(/<tableColumns count="4"><tableColumn id="1" name="Name"\/><tableColumn id="3" name="Column1"\/><tableColumn id="4" name="Column2"\/><tableColumn id="2" name="Qty"\/>/);
    expect(await entry(out, 'xl/drawings/drawing1.xml')).toMatch(/<xdr:from><xdr:col>5<\/xdr:col>/);
    expect(await entry(out, 'xl/workbook.xml')).toContain("'Data'!$A$1:$D$6");
  });

  it('deletes columns, turning references to them into #REF!', async () => {
    const out = await bytesOf(new SheetEditor().deleteColumns('Data', 2).edit(blob(await sample())));
    const data = await read(out, 'Data');
    expect(data.get(2)!.cells).toEqual(['a']);
    expect(data.get(1)!.formulas?.[1]).toBe('SUM(#REF!)');
    expect((await read(out, 'Other')).get(1)!.formulas).toEqual(['Data!#REF!', 'SUM(Data!#REF!)', 'B1']);

    const sheet = await entry(out, 'xl/worksheets/sheet1.xml');
    expect(sheet).not.toContain('mergeCell');       // A8:B8 is down to one cell
    expect(sheet).not.toContain('conditionalFormatting');
    expect(sheet).not.toContain('dataValidation');
    const table = await entry(out, 'xl/tables/table1.xml');
    expect(table).toMatch(/<table [^>]*ref="A1:A6"/);
    expect(table).toContain('<tableColumns count="1"><tableColumn id="1" name="Name"/></tableColumns>');
    expect(await entry(out, 'xl/drawings/drawing1.xml')).toMatch(/<xdr:from><xdr:col>2<\/xdr:col>/);
  });

  it('combines row and column moves and keeps filters on their columns', async () => {
    const w = new SheetWriter();
    w.addSheet('S', [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]], { autoFilter: 'A1:C3', columns: [{ width: 5 }, { width: 10 }, { width: 20 }] });
    const src = await bytesOf(w.write());
    const out = await bytesOf(new SheetEditor().deleteColumns('S', 'A').insertRows('S', 1).edit(blob(src)));
    const rows = await read(out, 'S');
    expect([...rows.keys()]).toEqual([2, 3, 4]);
    expect(rows.get(2)!.cells).toEqual(['b', 'c']);
    const sheet = await entry(out, 'xl/worksheets/sheet1.xml');
    expect(sheet).toMatch(/<autoFilter ref="A2:B4"/);
    expect(sheet).toMatch(/<col min="1" max="1" width="10/);
    expect(sheet).toMatch(/<col min="2" max="2" width="20/);
  });
});
describe('sparklines and data tables', () => {
  const shift = async (xml: string, maps: Map<string, { rows?: AxisMap; cols?: AxisMap }>, sheet = 'S') => {
    const stream = new Response(xml).body!.pipeThrough(new TextDecoderStream()).pipeThrough(createShiftTransform(sheet, maps));
    return new Response(stream.pipeThrough(new TextEncoderStream())).text();
  };
  const sparks = (spark: string) => `<worksheet><sheetData/><extLst><ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}"><x14:sparklineGroups xmlns:xm="m"><x14:sparklineGroup type="line"><x14:sparklines>${spark}</x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst></worksheet>`;
  const line = (f: string, at: string) => `<x14:sparkline><xm:f>${f}</xm:f><xm:sqref>${at}</xm:sqref></x14:sparkline>`;

  it('moves sparkline data and locations, dropping ones whose cell is deleted', async () => {
    const rows = new Map([['s', { rows: new AxisMap([{ at: 2, count: 1, insert: false }]) }]]);
    const out = await shift(sparks(line('S!A1:A10', 'B1') + line('S!A1:A10', 'B2')), rows);
    expect(out).toContain(line('S!A1:A9', 'B1'));
    expect(out).not.toContain('B2</xm:sqref>');
    // The only sparkline gone: the group and the empty extension go too
    expect(await shift(sparks(line('S!A1:A10', 'B2')), rows)).toBe('<worksheet><sheetData/></worksheet>');
    // Another sheet's sparkline over the moved rows follows them
    expect(await shift(sparks(line('S!A1:A10', 'B2')), rows, 'Other')).toContain(line('S!A1:A9', 'B2'));
  });

  it('moves a What-If data table with its input cells', async () => {
    const xml = '<worksheet><sheetData><row r="5"><c r="B5"><f t="dataTable" ref="B5:C9" dt2D="0" dtr="0" r1="A1"/><v>1</v></c></row></sheetData></worksheet>';
    const out = await shift(xml, new Map([['s', { rows: new AxisMap([{ at: 1, count: 2, insert: true }]) }]]));
    expect(out).toContain('<row r="7"><c r="B7"><f t="dataTable" ref="B7:C11" dt2D="0" dtr="0" r1="A3"/><v>1</v></c></row>');
  });
});