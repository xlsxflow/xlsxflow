import { describe, it, expect } from 'vitest';
import { SheetWriter } from '../src/core/writer';
import { SheetEditor } from '../src/core/editor';
import { SheetReader } from '../src/core/index';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { createBlobReader } from '../src/core/random-access';
import { parseCsv } from '../src/core/csv';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());
const blob = (b: Uint8Array) => createBlobReader(new Blob([b]));

async function read(bytes: Uint8Array, sheetName?: string) {
  const rows = await new SheetReader().parse(blob(bytes), { sheetName, formulas: true, styles: true });
  const out: { r: number; cells: unknown[]; formulas?: (string | undefined)[]; styles?: any[] }[] = [];
  for await (const row of rows) out.push({ r: row.rowNumber, cells: row.cells, formulas: row.formulas, styles: row.styles });
  return out;
}

async function entry(bytes: Uint8Array, name: string) {
  const zip = new ZipRandomAccessParser(blob(bytes));
  await zip.parseCentralDirectory();
  return zip.has(name) ? new Response(await zip.extractStream(name)).text() : '';
}

describe('SheetEditor styles', () => {
  it('merges style changes into the current format and keeps content', async () => {
    const src = await bytesOf(new SheetWriter().addSheet('S', [
      [{ value: 'head', style: { font: { bold: true, color: 'FF0000FF' }, alignment: { horizontal: 'center' } } }, 7],
      [{ value: 1.5, style: { numFmt: '0.00', border: { top: { style: 'thin' } } } }],
    ]).write());
    const out = await bytesOf(new SheetEditor().setCells('S', {
      A1: { style: { font: { italic: true, bold: false } } },          // keeps text, colour and alignment
      B1: { value: 8, style: { fill: { type: 'solid', fgColor: 'FFFFFF00' } } },
      A2: { style: { border: { bottom: { style: 'double' } }, numFmt: '0.000' } },
      C2: { style: { font: { bold: true } } },                          // a new, empty styled cell
      D2: { formula: 'B1*2', style: { font: { bold: true } } },
    }).edit(blob(src)));
    const rows = await read(out);

    expect(rows[0].cells).toEqual(['head', 8]);
    expect(rows[0].styles?.[0]).toMatchObject({ font: { italic: true, color: 'FF0000FF' }, alignment: { horizontal: 'center' } });
    expect(rows[0].styles?.[0].font.bold).toBeFalsy();
    expect(rows[0].styles?.[1]).toMatchObject({ fill: { type: 'solid', fgColor: 'FFFFFF00' } });
    expect(rows[1].cells[0]).toBe(1.5);
    expect(rows[1].styles?.[0]).toMatchObject({ numFmt: '0.000', border: { top: { style: 'thin' }, bottom: { style: 'double' } } });
    expect(rows[1].styles?.[2]).toMatchObject({ font: { bold: true } });
    expect(rows[1].formulas?.[3]).toBe('B1*2');
    expect(rows[1].styles?.[3]).toMatchObject({ font: { bold: true } });

    // The same change on two cells adds one format
    const styles = await entry(out, 'xl/styles.xml');
    const before = (await entry(src, 'xl/styles.xml')).match(/<xf /g)!.length;
    expect(styles.match(/<xf /g)!.length).toBe(before + 4);
  });
});

describe('SheetEditor sheets', () => {
  it('adds a sheet with values, formulas and styles', async () => {
    const src = await bytesOf(new SheetWriter().addSheet('First', [['a']]).write());
    const out = await bytesOf(new SheetEditor()
      .addSheet('New & Co', [['Name', { value: 2, style: { font: { bold: true } } }], [null, { value: null, formula: '=B1+1' }]])
      .edit(blob(src)));
    const rows = await read(out, 'New & Co');
    expect(rows[0].cells).toEqual(['Name', 2]);
    expect(rows[0].styles?.[1]).toMatchObject({ font: { bold: true } });
    expect(rows[1].formulas?.[1]).toBe('B1+1');
    expect((await read(out, 'First'))[0].cells).toEqual(['a']);
    const wb = await entry(out, 'xl/workbook.xml');
    expect(wb).toContain('name="New &amp; Co" sheetId="2"');
    expect(await entry(out, '[Content_Types].xml')).toContain('/xl/worksheets/sheet2.xml');

    await expect(bytesOf(new SheetEditor().addSheet('first', []).edit(blob(src)))).rejects.toThrow(/Duplicate sheet name/);
    expect(() => new SheetEditor().addSheet('a:b', [])).toThrow(/Invalid sheet name/);
  });

  it('deletes a sheet and fixes the names that refer to sheets', async () => {
    const w = new SheetWriter();
    w.addSheet('One', [[1]], { pageSetup: { printArea: 'A1:B2' } });
    w.addSheet('Two', [[2]], { pageSetup: { printArea: 'A1:C3' } });
    w.addSheet('Three', [[3]], { pageSetup: { printArea: 'A1:D4' } });
    const src = await bytesOf(w.write());
    const out = await bytesOf(new SheetEditor().deleteSheet('Two').edit(blob(src)));

    const wb = await entry(out, 'xl/workbook.xml');
    expect(wb).not.toContain('name="Two"');
    expect(wb).toContain('localSheetId="0">&apos;One&apos;!$A$1:$B$2');
    expect(wb).toContain('localSheetId="1">&apos;Three&apos;!$A$1:$D$4');
    expect(wb).not.toContain('C$3');
    expect(await entry(out, 'xl/worksheets/sheet2.xml')).toBe('');
    expect(await entry(out, '[Content_Types].xml')).not.toContain('sheet2.xml');
    expect(await entry(out, 'xl/_rels/workbook.xml.rels')).not.toContain('sheet2.xml');
    expect((await read(out, 'Three'))[0].cells).toEqual([3]);

    await expect(bytesOf(new SheetEditor().deleteSheet('Nope').edit(blob(src)))).rejects.toThrow(/not found/);
    const single = await bytesOf(new SheetWriter().addSheet('Only', [[1]]).write());
    await expect(bytesOf(new SheetEditor().deleteSheet('Only').edit(blob(single)))).rejects.toThrow(/visible sheet/);
  });

  it('replaces a deleted sheet with a new one of the same name', async () => {
    const src = await bytesOf(new SheetWriter().addSheet('A', [[1]]).addSheet('B', [[2]]).write());
    const out = await bytesOf(new SheetEditor().deleteSheet('B').addSheet('B', [['fresh']]).edit(blob(src)));
    expect((await read(out, 'B'))[0].cells).toEqual(['fresh']);
    expect(await entry(out, 'xl/workbook.xml')).toMatch(/^(?![\s\S]*name="B"[\s\S]*name="B")[\s\S]*<sheet name="B" sheetId="\d" r:id="rId\d"\/>/);
  });
});

describe('notes with rich text', () => {
  it('writes formatted runs and reads the text back', async () => {
    const out = await bytesOf(new SheetWriter().addSheet('S', [[{
      value: 1, comment: { author: 'Ana', text: [{ text: 'Ana:', font: { bold: true } }, { text: ' restated' }] },
    }]]).write());
    expect(await entry(out, 'xl/comments1.xml')).toContain('<r><rPr><b/></rPr><t>Ana:</t></r><r><t xml:space="preserve"> restated</t></r>');
    const rows = await new SheetReader().parse(blob(out));
    expect(await rows.getComments()).toEqual([{ ref: 'A1', text: 'Ana: restated', author: 'Ana' }]);
  });
});

describe('parseCsv', () => {
  const all = async (it: AsyncIterable<unknown[]>) => {
    const rows: unknown[][] = [];
    for await (const row of it) rows.push(row);
    return rows;
  };

  it('parses quoting, line endings and types', async () => {
    const csv = '﻿name,qty,ok,code\r\n"Smith, J",3,true,007\n"He said ""hi""",-1.5e2,FALSE,\r"multi\nline",,x,"12"\n\nlast';
    expect(await all(parseCsv(csv))).toEqual([
      ['name', 'qty', 'ok', 'code'],
      ['Smith, J', 3, true, '007'],
      ['He said "hi"', -150, false, null],
      ['multi\nline', null, 'x', '12'],
      [],
      ['last'],
    ]);
    expect(await all(parseCsv('a;1\n', { delimiter: ';', convert: false }))).toEqual([['a', '1']]);
    await expect(all(parseCsv('"open'))).rejects.toThrow(/quoted field/);
  });

  it('streams across chunk boundaries into a workbook', async () => {
    const csv = 'id,text\n' + Array.from({ length: 5000 }, (_, i) => `${i},"line ""${i}""\r\nnext"`).join('\r\n') + '\r\n';
    // One byte at a time around a split quote pair and \r\n
    const bytes = new TextEncoder().encode(csv);
    const stream = new ReadableStream<Uint8Array>({
      start(c) {
        for (let i = 0; i < bytes.length; i += 7) c.enqueue(bytes.slice(i, i + 7));
        c.close();
      },
    });
    const rows = await all(parseCsv(stream));
    expect(rows).toHaveLength(5001);
    expect(rows[4000]).toEqual([3999, 'line "3999"\r\nnext']);

    const out = await bytesOf(new SheetWriter().addSheet('CSV', parseCsv(csv)).write());
    expect((await read(out))[5000].cells).toEqual([4999, 'line "4999"\r\nnext']);
  });
});
