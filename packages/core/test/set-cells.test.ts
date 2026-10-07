import { describe, it, expect } from 'vitest';
import { readFileSync } from 'fs';
import { SheetWriter } from '../src/core/writer';
import { SheetEditor } from '../src/core/editor';
import { SheetReader } from '../src/core/index';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { createBlobReader } from '../src/core/random-access';
import type { Row } from '../src/core/types';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());

async function read(bytes: Uint8Array, sheetName?: string) {
  const rows = await new SheetReader().parse(createBlobReader(new Blob([bytes])), { sheetName, formulas: true, styles: true });
  const out: { r: number; cells: unknown[]; formulas?: (string | undefined)[]; styles?: unknown[] }[] = [];
  for await (const row of rows) out.push({ r: row.rowNumber, cells: row.cells, formulas: row.formulas, styles: row.styles });
  return out;
}

async function entry(bytes: Uint8Array, name: string) {
  const zip = new ZipRandomAccessParser(createBlobReader(new Blob([bytes])));
  await zip.parseCentralDirectory();
  return zip.has(name) ? new Response(await zip.extractStream(name)).text() : '';
}

const bold = { font: { bold: true } };

describe('SheetEditor.setCells', () => {
  it('overwrites, adds and clears cells, keeping their styles', async () => {
    const w = new SheetWriter();
    w.addSheet('Other', [['untouched']]);
    w.addSheet('Data', [
      ['Name', { value: 'Amount', style: bold }],
      ['a', { value: 1, style: { numFmt: '0.00' } }],
      ['b', 2],
      [],
      ['e', 5],
    ]);
    const src = await bytesOf(w.write());
    const editor = new SheetEditor().setCells('Data', {
      B1: 'Total & tax', B2: 10.5, A3: null, C3: { formula: '=B2*2' }, D3: true,
      A4: ' padded ', B7: new Date(Date.UTC(2026, 0, 2)), Z1: 'far',
    });
    const out = await bytesOf(editor.edit(createBlobReader(new Blob([src]))));
    const rows = await read(out, 'Data');

    expect(rows.map(r => r.r)).toEqual([1, 2, 3, 4, 5, 7]);
    expect(rows[0].cells[1]).toBe('Total & tax');
    expect(rows[0].styles?.[1]).toEqual(bold); // style kept
    expect(rows[0].cells[25]).toBe('far');
    expect(rows[1].cells[1]).toBe(10.5);
    expect(rows[1].styles?.[1]).toMatchObject({ numFmt: '0.00' });
    expect(rows[2].cells.slice(0, 4)).toEqual([null, 2, null, true]);
    expect(rows[2].formulas?.[2]).toBe('B2*2');
    expect(rows[3].cells).toEqual([' padded ']);
    expect(rows[4].cells).toEqual(['e', 5]);
    expect(rows[5].cells[1]).toBe(46024); // a date over an unformatted cell is its serial number

    expect(await read(out, 'Other')).toEqual([expect.objectContaining({ cells: ['untouched'] })]);
    expect(await entry(out, 'xl/workbook.xml')).toContain('fullCalcOnLoad="1"');

    // The editor can be reused: its edits are not consumed
    expect((await read(await bytesOf(editor.edit(createBlobReader(new Blob([src])))), 'Data'))[1].cells[1]).toBe(10.5);
  });

  it('fills an empty sheet and combines with appendSheet', async () => {
    const src = await bytesOf(new SheetWriter().addSheet('E', []).write());
    const editor = new SheetEditor().setCells('E', { B2: 'x' });
    editor.appendSheet('E', [['appended']]);
    const rows = await read(await bytesOf(editor.edit(createBlobReader(new Blob([src])))));
    expect(rows).toEqual([expect.objectContaining({ r: 2, cells: [null, 'x'] }), expect.objectContaining({ r: 3, cells: ['appended'] })]);
  });

  it('streams large sheets across chunk boundaries', async () => {
    const data: Row[] = Array.from({ length: 20000 }, (_, i) => [`row ${i + 1}`, i + 1]);
    const src = await bytesOf(new SheetWriter().addSheet('Big', data).write());
    const out = await bytesOf(new SheetEditor().setCells('Big', { B1: -1, B9999: -2, A20000: 'last', A20002: 'after' }).edit(createBlobReader(new Blob([src]))));
    const rows = await read(out);
    expect(rows).toHaveLength(20001);
    expect(rows[0].cells).toEqual(['row 1', -1]);
    expect(rows[9998].cells).toEqual(['row 9999', -2]);
    expect(rows[10000].cells).toEqual(['row 10001', 10001]);
    expect(rows[19999].cells).toEqual(['last', 20000]);
    expect(rows[20000]).toMatchObject({ r: 20002, cells: ['after'] });
  }, 30000); // 20,000 rows: slow when the machine is busy

  it('writes out shared formulas whose anchor was overwritten, and drops the calculation chain', async () => {
    // C3 anchors the shared formula B3*2 that C4 reuses; the file also has a calcChain
    const tpl = new Uint8Array(readFileSync(new URL('./fixtures/openpyxl-template.xlsx', import.meta.url)));
    const out = await bytesOf(new SheetEditor().setCells('Report', { C3: 0 }).edit(createBlobReader(new Blob([tpl]))));
    const rows = await read(out);
    expect(rows.find(r => r.r === 3)?.cells[2]).toBe(0);
    expect(rows.find(r => r.r === 4)?.formulas?.[2]).toBe('B4*2');
    expect(await entry(out, 'xl/worksheets/sheet1.xml')).not.toContain('si="0"/>');
    expect(await entry(out, 'xl/calcChain.xml')).toBe('');
    expect(await entry(out, '[Content_Types].xml')).not.toContain('calcChain');
  });

  it('rejects bad input', async () => {
    expect(() => new SheetEditor().setCells('S', { 'A0': 1 })).toThrow(/Invalid cell reference/);
    expect(() => new SheetEditor().setCells('S', { '1A': 1 })).toThrow(/Invalid cell reference/);
    const src = await bytesOf(new SheetWriter().addSheet('S', [[1]]).write());
    await expect(bytesOf(new SheetEditor().setCells('Nope', { A1: 1 }).edit(createBlobReader(new Blob([src]))))).rejects.toThrow(/Sheet "Nope" not found/);
  });
});
