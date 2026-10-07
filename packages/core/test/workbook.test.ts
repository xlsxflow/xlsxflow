import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { SheetWriter } from '../src/core/writer';
import { SheetEditor } from '../src/core/editor';
import { SheetReader } from '../src/core/index';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { createBlobReader } from '../src/core/random-access';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());
const blob = (b: Uint8Array) => createBlobReader(new Blob([b]));

async function entry(bytes: Uint8Array, name: string) {
  const zip = new ZipRandomAccessParser(blob(bytes));
  await zip.parseCentralDirectory();
  return zip.has(name) ? new Response(await zip.extractStream(name)).text() : '';
}

describe('workbook properties, names and sheet visibility', () => {
  it('writes them and reads them back', async () => {
    const created = new Date(Date.UTC(2026, 9, 8, 9, 30));
    const w = new SheetWriter({
      properties: { title: 'Q3 <report>', creator: 'Ana', subject: 'Sales', keywords: 'q3 sales', description: 'Draft', category: 'Finance', company: 'ACME & Co', manager: 'Raj', created },
      definedNames: [
        { name: 'TaxRate', ref: '0.18', comment: 'GST' },
        { name: 'Sales', ref: "=Data!$A$1:$B$3" },
        { name: 'Local.Total', ref: "'Other sheet'!$A$1", sheet: 'Other sheet', hidden: true },
      ],
    });
    w.addSheet('Lookup', [['k', 'v']], { state: 'hidden' });
    w.addSheet('Data', [['a', 1], ['b', 2], ['c', 3]], { freezePanes: { row: 1 }, view: { zoom: 125, showGridLines: false, rightToLeft: true } });
    w.addSheet('Other sheet', [[{ value: null, formula: 'SUM(Sales)*TaxRate' }]], { state: 'veryHidden', view: { showHeadings: false } });
    const out = await bytesOf(w.write());

    const info = await new SheetReader().readWorkbook(blob(out));
    expect(info.sheets).toEqual([
      { name: 'Lookup', state: 'hidden' }, { name: 'Data', state: 'visible' }, { name: 'Other sheet', state: 'veryHidden' },
    ]);
    expect(info.definedNames).toEqual([
      { name: 'TaxRate', ref: '0.18', comment: 'GST' },
      { name: 'Sales', ref: 'Data!$A$1:$B$3' },
      { name: 'Local.Total', ref: "'Other sheet'!$A$1", sheet: 'Other sheet', hidden: true },
    ]);
    expect(info.properties).toEqual({ title: 'Q3 <report>', creator: 'Ana', subject: 'Sales', keywords: 'q3 sales', description: 'Draft', category: 'Finance', company: 'ACME & Co', manager: 'Raj', created });

    // Excel opens on the first visible sheet; only views that were asked for are written
    const wb = await entry(out, 'xl/workbook.xml');
    expect(wb).toContain('<bookViews><workbookView firstSheet="1" activeTab="1"/></bookViews>');
    const data = await entry(out, 'xl/worksheets/sheet2.xml');
    expect(data).toContain('<sheetView showGridLines="0" rightToLeft="1" zoomScale="125" workbookViewId="0"><pane ');
    expect(data).not.toContain('tabSelected');
    expect(await entry(out, 'xl/worksheets/sheet3.xml')).toContain('<sheetView showRowColHeaders="0" workbookViewId="0"></sheetView>');
    expect(await entry(out, 'xl/worksheets/sheet1.xml')).not.toContain('<sheetViews>');
    expect(await entry(out, '[Content_Types].xml')).toContain('/docProps/core.xml');
    expect(await entry(out, '_rels/.rels')).toContain('Target="docProps/app.xml"');

    // The reader still opens the first sheet in the file, hidden or not, and names survive an edit
    const edited = await bytesOf(new SheetEditor().setCells('Data', { B1: { value: 10 } }).edit(blob(out)));
    expect((await new SheetReader().readWorkbook(blob(edited))).definedNames).toHaveLength(3);
  });

  it('leaves out everything that was not asked for', async () => {
    const out = await bytesOf(new SheetWriter().addSheet('S', [[1]]).write());
    expect(await entry(out, 'docProps/core.xml')).toBe('');
    expect(await entry(out, 'xl/workbook.xml')).not.toMatch(/bookViews|definedNames/);
    expect(await new SheetReader().readWorkbook(blob(out))).toEqual({ sheets: [{ name: 'S', state: 'visible' }], definedNames: [], properties: {} });
  });

  it('rejects names and settings Excel would refuse', () => {
    const write = (options: ConstructorParameters<typeof SheetWriter>[0], sheet: Parameters<SheetWriter['addSheet']>[2] = {}) =>
      () => new SheetWriter(options).addSheet('S', [[1]], sheet).write();
    for (const name of ['A1', 'xfd1048576', 'R1C1', 'r', 'C', '1st', 'has space', '_xlnm.Print_Area']) {
      expect(write({ definedNames: [{ name, ref: '1' }] }), name).toThrow(/Invalid defined name/);
    }
    for (const name of ['Rate', '_tax', '\\x', 'Total.2026', 'AAAA1']) expect(write({ definedNames: [{ name, ref: '1' }] }), name).not.toThrow();
    expect(write({ definedNames: [{ name: 'X', ref: '1' }, { name: 'x', ref: '2' }] })).toThrow(/Duplicate/);
    expect(write({ definedNames: [{ name: 'X', ref: '1' }, { name: 'x', ref: '2', sheet: 'S' }] })).not.toThrow();
    expect(write({ definedNames: [{ name: 'X', ref: '1', sheet: 'Nope' }] })).toThrow(/not found/);
    expect(write({}, { state: 'hidden' })).toThrow(/visible sheet/);
    expect(write({}, { view: { zoom: 500 } })).toThrow(/Zoom/);
  });

  it('reads names and properties from files other libraries wrote', async () => {
    const info = await new SheetReader().readWorkbook(blob(new Uint8Array(readFileSync(new URL('./fixtures/openpyxl-template.xlsx', import.meta.url)))));
    expect(info.sheets.length).toBeGreaterThan(0);
    expect(info.properties.creator).toBe('openpyxl');
  });
});
