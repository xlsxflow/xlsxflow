// Regressions for the third round of end-user tests (core 1.1.4)
import { describe, it, expect } from 'vitest';
import path from 'path';
import { SheetWriter } from '../src/core/writer';
import { SheetEditor } from '../src/core/editor';
import { SheetReader } from '../src/core/index';
import { OdsWriter } from '../src/core/ods-writer';
import { FormulaEngine } from '../src/core/formula-engine';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { ZipStreamWriter } from '../src/core/zip-stream-writer';
import { createBlobReader } from '../src/core/random-access';
import { parseCsv } from '../src/core/csv';
import { sheetToJson, streamToCsv } from '../src/core/utils';
import { anchorXml, imageInfo } from '../src/core/image';
import { XlsxFlow } from '../src/index';
import type { Row } from '../src/core/types';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());
const blob = (b: Uint8Array) => createBlobReader(new Blob([b]));
const write = (rows: Row[] | AsyncIterable<Row>, options = {}) => bytesOf(new SheetWriter().addSheet('S', rows, options).write());

async function entries(bytes: Uint8Array): Promise<Record<string, string>> {
  const zip = new ZipRandomAccessParser(blob(bytes));
  await zip.parseCentralDirectory();
  const out: Record<string, string> = {};
  for (const f of zip.getFiles()) out[f] = await new Response(await zip.extractStream(f)).text();
  return out;
}
async function zipOf(files: Record<string, string | Uint8Array>): Promise<Uint8Array> {
  const zip = new ZipStreamWriter();
  const done = bytesOf(zip.stream);
  for (const [name, data] of Object.entries(files)) await zip.addFile(name, new Response(data).body!);
  await zip.close();
  return done;
}
// A one-sheet workbook around the given <sheetData> content
const crafted = (sheetData: string, sheet = 'S') => zipOf({
  '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
  '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  'xl/workbook.xml': `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${sheet}" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  'xl/_rels/workbook.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  'xl/worksheets/sheet1.xml': `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${sheetData}</sheetData></worksheet>`,
});
async function rows(bytes: Uint8Array, options = {}) {
  const out = [];
  for await (const r of await new SheetReader().parse(blob(bytes), options)) out.push(r);
  return out;
}
const cells = async (bytes: Uint8Array, options = {}) => (await rows(bytes, options)).map(r => r.cells);
// Control characters XML 1.0 cannot hold
const hasRawControl = (xml: string) => /[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/.test(xml);

describe('formula engine', () => {
  const cached = async (row: Row) => (await cells(await write([row])))[0];
  const f = (formula: string) => ({ value: null, formula });

  it('follows formulas over formula cells, and stores 0 for a circular reference', async () => {
    expect(await cached([1, f('A1*2'), f('B1+1')])).toEqual([1, 2, 3]);
    expect(await cached([f('A1+1')])).toEqual([0]);
  });

  it('reads scientific literals, reversed ranges and values typed into SUM as Excel does', async () => {
    expect(await cached([f('1E3'), f('2.5E-1*4'), f('SUM(TRUE,1)'), f('SUM("3",1)')])).toEqual([1000, 1, 2, 4]);
    expect((await cells(await write([[1], [2], [3], [f('SUM(A3:A1)')]])))[3]).toEqual([6]);
  });

  it('compares text without case, puts text after numbers, and a blank equals ""', async () => {
    expect(await cached([null, f('"a"="A"'), f('A1=""'), f('"b">"a"'), f('"1"=1'), f('"a">5')])).toEqual([null, true, true, true, false, true]);
  });

  it('turns numbers into text with 15 digits and booleans into TRUE/FALSE', async () => {
    expect(await cached([f('CONCATENATE(1/3)'), f('CONCATENATE("x",0.1+0.2)'), f('CONCATENATE("x",TRUE)')])).toEqual(['0.333333333333333', 'x0.3', 'xTRUE']);
  });

  it('passes errors through, and caches them as error cells; text starting with "#" stays text', async () => {
    const xml = (await entries(await write([['a', f('1/0+1'), f('A1+1'), f('SUM(1/0,1)'), f('IF(1,"#tag","x")')]])))['xl/worksheets/sheet1.xml'];
    expect(xml).toContain('<c r="B1" t="e"><f>1/0+1</f><v>#DIV/0!</v></c>');
    expect(xml).toContain('<c r="C1" t="e"><f>A1+1</f><v>#VALUE!</v></c>');
    expect(xml).toContain('<c r="D1" t="e"><f>SUM(1/0,1)</f><v>#DIV/0!</v></c>');
    expect(xml).toContain('<c r="E1" t="str"><f>IF(1,&quot;#tag&quot;,&quot;x&quot;)</f><v>#tag</v></c>');
  });

  it('gives streamed rows no made-up cached value', async () => {
    async function* streamed(): AsyncGenerator<Row> { yield [1]; yield [2]; yield [f('SUM(A1:A2)')]; }
    expect((await entries(await write(streamed())))['xl/worksheets/sheet1.xml']).toContain('<c r="A3"><f>SUM(A1:A2)</f></c>');
  });

  it('evaluate: IF without an else gives FALSE, and a bare blank reference 0', () => {
    const e = new FormulaEngine();
    e.loadData([[1]], 5);
    expect([e.evaluate('IF(0,1)'), e.evaluate('A1'), e.evaluate('A5*2')]).toEqual([false, 0, 2]);
  });
});

describe('writer', () => {
  it('keeps characters XML cannot hold out of every part', async () => {
    const parts = await entries(await bytesOf(new SheetWriter({ properties: { title: 'a\u0001b' } }).addSheet('S', [
      [{ value: 'x￿', comment: { text: 'n', author: 'a\u0002' } }, { value: null, formula: 'CONCATENATE("a\u0001b")' }],
    ], { pageSetup: { footer: 'f\u0003' }, dataValidations: [{ sqref: 'A1', type: 'list', formula1: '"a,b"', prompt: 'p\u0004' }] }).write()));
    for (const [name, xml] of Object.entries(parts)) expect(hasRawControl(xml), name).toBe(false);
    expect(await cells(await write([['x￿']]))).toEqual([['x￿']]);
    expect(() => new SheetWriter().addSheet('a\u0001', [])).toThrow(/Invalid sheet name/);
  });

  it('refuses overlapping or malformed merges and tables, cell-like table names and long defined names', async () => {
    await expect(write([[1]], { mergeCells: ['A1:B2', 'B2:C3'] })).rejects.toThrow(/overlaps/);
    await expect(write([[1]], { mergeCells: ['A1:XFE1'] })).rejects.toThrow(/Invalid merge range/);
    const data = [['a', 'b', 'c'], [1, 2, 3], [4, 5, 6]];
    await expect(write(data, { tables: [{ name: 'Table1', ref: 'A1:B3' }, { name: 'Table2', ref: 'B1:C3' }] })).rejects.toThrow(/overlaps/);
    for (const name of ['AB12', 'R1C1', 'C']) await expect(write(data, { tables: [{ name, ref: 'A1:C3' }] })).rejects.toThrow(/table name/);
    const long = new SheetWriter({ definedNames: [{ name: 'n'.repeat(256), ref: 'S!A1' }] }).addSheet('S', []);
    expect(() => long.write()).toThrow(/Invalid defined name/);
  });

  it('refuses row options outside the sheet and list literals over 255 characters', async () => {
    await expect(write([[1]], { rows: { 0: { height: 20 } } })).rejects.toThrow(/rows are numbered/);
    const list = `"${'x,'.repeat(150)}"`;
    await expect(write([[1]], { dataValidations: [{ sqref: 'A1', type: 'list', formula1: list }] })).rejects.toThrow(/Excel allows 255/);
  });

  it('writes validation formulas without "=", and names the sheet on every print-area range', async () => {
    const parts = await entries(await write([[1]], {
      dataValidations: [{ sqref: 'A1', type: 'list', formula1: '=$D$1:$D$3' }], pageSetup: { printArea: 'A1:B2,D1:E2' },
    }));
    expect(parts['xl/worksheets/sheet1.xml']).toContain('<formula1>$D$1:$D$3</formula1>');
    expect(parts['xl/workbook.xml']).toContain('&apos;S&apos;!$A$1:$B$2,&apos;S&apos;!$D$1:$E$2');
  });

  it('refuses truncated images and sizes that are not positive numbers', () => {
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13]);
    expect(() => imageInfo(png)).toThrow(/truncated/);
    for (const width of [-100, NaN, Infinity]) expect(() => anchorXml({ at: 'A1', width }, { width: 10, height: 10 }, '')).toThrow(/Invalid size/);
    expect(() => anchorXml(undefined as never, { width: 10, height: 10 }, '')).toThrow(/needs a placement/);
  });

  it('stops pulling rows while nobody reads, and ends the row source when the output is cancelled', async () => {
    let pulled = 0, finalized = false;
    async function* source(): AsyncGenerator<Row> {
      try { for (;;) { if (++pulled % 500 === 0) await new Promise(setImmediate); yield [crypto.randomUUID(), crypto.randomUUID()]; } } finally { finalized = true; }
    }
    const reader = new SheetWriter().addSheet('S', source()).write().getReader();
    await new Promise(r => setTimeout(r, 500));
    expect(pulled).toBeLessThan(100000);
    await reader.cancel();
    await new Promise(r => setTimeout(r, 200));
    expect(finalized).toBe(true);
  });

  it('marks UTF-8 entry names and refuses 65,535 entries', async () => {
    const bytes = await zipOf({ 'données.txt': 'x' });
    expect(new DataView(bytes.buffer).getUint16(6, true) & 0x800).toBe(0x800);
    const zip = new ZipStreamWriter();
    zip.stream.pipeTo(new WritableStream()).catch(() => {});
    (zip as unknown as { cdEntries: unknown[] }).cdEntries.length = 0xffff;
    await expect(zip.close()).rejects.toThrow(/65535 entries/);
  });
});

describe('reader', () => {
  it('reads .ods files, comments and images through XlsxFlow.readFile', async () => {
    const ods = await XlsxFlow.readFile(path.join(__dirname, 'fixtures', 'excel-made.ods'));
    let n = 0;
    for await (const _ of ods) n++;
    expect(n).toBeGreaterThan(0);
    const xlsx = await XlsxFlow.readFile(path.join(__dirname, 'fixtures', 'excel-made.xlsx'));
    for await (const _ of xlsx) break;
    await expect(xlsx.getComments()).resolves.toBeInstanceOf(Array);
    await expect(xlsx.getImages()).resolves.toBeInstanceOf(Array);
  });

  it('refuses entries whose data no longer matches their CRC-32', async () => {
    const good = await crafted('<row r="1"><c r="A1"><v>7</v></c></row>');
    const zip = new ZipStreamWriter();
    const done = bytesOf(zip.stream);
    const parts = await entries(good);
    for (const [name, xml] of Object.entries(parts)) {
      const data = new TextEncoder().encode(name.endsWith('sheet1.xml') ? xml.replace('<v>7<', '<v>8<') : xml);
      // stored with the CRC of the original text
      const { crc32 } = await import('../src/core/zip-stream-writer');
      await zip.addCompressedFile(name, new Response(data).body!, data.length, data.length, crc32(new TextEncoder().encode(xml)), 0);
    }
    await zip.close();
    await expect(cells(await done)).rejects.toThrow(/CRC-32/);
  });

  it('reads past a zip comment holding the end-record signature, and part names in any case', async () => {
    const bytes = await crafted('<row r="1"><c r="A1"><v>1</v></c></row>');
    const comment = new Uint8Array([0x50, 0x4b, 0x05, 0x06, 1, 2, 3, 4]);
    const withComment = new Uint8Array(bytes.length + comment.length);
    withComment.set(bytes);
    withComment.set(comment, bytes.length);
    new DataView(withComment.buffer).setUint16(bytes.length - 2, comment.length, true);
    expect(await cells(withComment)).toEqual([[1]]);

    const parts = await entries(bytes);
    parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace('sheet1.xml', 'Sheet1.xml');
    expect(await cells(await zipOf(parts))).toEqual([[1]]);
  });

  it('reads UTF-16 parts, numbers rows with a bad r after the previous row, and refuses a cut-off sheet', async () => {
    const parts = await entries(await crafted('<row r="1"><c r="A1" t="inlineStr"><is><t>wide é</t></is></c></row>'));
    const xml = parts['xl/worksheets/sheet1.xml'];
    const utf16 = new Uint8Array(2 + xml.length * 2);
    utf16.set([0xff, 0xfe]);
    for (let i = 0; i < xml.length; i++) new DataView(utf16.buffer).setUint16(2 + i * 2, xml.charCodeAt(i), true);
    expect(await cells(await zipOf({ ...parts, 'xl/worksheets/sheet1.xml': utf16 }))).toEqual([['wide é']]);

    const numbers = (await rows(await crafted('<row r="1"><c><v>1</v></c></row><row r="abc"><c><v>2</v></c></row><row r="0"><c><v>3</v></c></row>'))).map(r => r.rowNumber);
    expect(numbers).toEqual([1, 2, 3]);
    const cut = (await entries(await crafted('<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>')));
    cut['xl/worksheets/sheet1.xml'] = cut['xl/worksheets/sheet1.xml'].slice(0, cut['xl/worksheets/sheet1.xml'].indexOf('<row r="2">') + 20);
    await expect(cells(await zipOf(cut))).rejects.toThrow(/cut off/);
  });

  it('gives t="d" cells as UTC ISO strings, formatted with their date format', async () => {
    const bytes = await crafted('<row r="1"><c r="A1" t="d"><v>2026-10-08T14:05:00</v></c><c r="B1" t="d"><v>2026-10-08</v></c></row>');
    expect(await cells(bytes)).toEqual([['2026-10-08T14:05:00.000Z', '2026-10-08T00:00:00.000Z']]);
  });

  it('lists sheet names unescaped once, and opens the first worksheet when a chart sheet comes first', async () => {
    const parts = await entries(await crafted('<row r="1"><c r="A1"><v>42</v></c></row>', 'R&amp;amp;D'));
    const info = await new SheetReader().readWorkbook(blob(await zipOf(parts)));
    expect(info.sheets.map(s => s.name)).toEqual(['R&amp;D']);
    expect(await cells(await zipOf(parts), { sheetName: 'R&amp;D' })).toEqual([[42]]);

    parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace('<sheets>', '<sheets><sheet name="Chart" sheetId="2" r:id="rId2"/>');
    parts['xl/_rels/workbook.xml.rels'] = parts['xl/_rels/workbook.xml.rels'].replace('</Relationships>',
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" Target="chartsheets/sheet1.xml"/></Relationships>');
    parts['xl/chartsheets/sheet1.xml'] = '<chartsheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>';
    expect(await cells(await zipOf(parts))).toEqual([[42]]);
  });

  it('marks error cells with the errors option', async () => {
    const bytes = await crafted('<row r="1"><c r="A1" t="e"><v>#N/A</v></c><c r="B1" t="inlineStr"><is><t>#N/A</t></is></c></row>');
    expect((await rows(bytes, { errors: true }))[0]).toEqual({ rowNumber: 1, cells: ['#N/A', '#N/A'], errors: [true] });
  });

  it('sheetToJson keeps columns right of the header; streamToCsv keeps blank rows', async () => {
    const parse = async (b: Uint8Array) => new SheetReader().parse(blob(b));
    expect(await sheetToJson(await parse(await write([['a', 'b'], [1, 2, 3]])))).toEqual([{ a: 1, b: 2, Column3: 3 }]);
    expect(await streamToCsv(await parse(await crafted('<row r="1"><c r="A1"><v>1</v></c></row><row r="4"><c r="B4"><v>2</v></c></row>')))).toBe('1\n\n\n,2');
  });

  it('parseCsv keeps numbers that overflow a double as text', async () => {
    const out = [];
    for await (const row of parseCsv('1e400,1e308')) out.push(row);
    expect(out).toEqual([['1e400', 1e308]]);
  });
});

describe('editor', () => {
  const edited = async (src: Uint8Array, edit: (e: SheetEditor) => unknown) => {
    const e = new SheetEditor();
    edit(e);
    return bytesOf(e.edit(blob(src)));
  };

  it('appendSheet writes dates, formulas, styles and both batches, in the sheet\'s namespace', async () => {
    const src = await write([['h']]);
    const out = await edited(src, e => {
      e.appendSheet('S', [[new Date(Date.UTC(2026, 1, 3)), { value: null, formula: 'A1&"x"' }, { value: 1, style: { font: { bold: true } } }]]);
      e.appendSheet('S', [['second']]);
    });
    const r = await rows(out, { formulas: true, styles: true });
    expect(r.map(x => x.cells[0])).toEqual(['h', '2026-02-03T00:00:00.000Z', 'second']);
    expect(r[1].formulas?.[1]).toBe('A1&"x"');
    expect(r[1].styles?.[2]?.font?.bold).toBe(true);
    const prefixed = (await entries(await crafted('<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>')));
    prefixed['xl/worksheets/sheet1.xml'] = '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>';
    const xml = (await entries(await edited(await zipOf(prefixed), e => e.appendSheet('S', [[NaN]]))))['xl/worksheets/sheet1.xml'];
    expect(xml).toContain('<x:row r="2"><x:c r="A2" t="e"><x:v>#NUM!</x:v></x:c></x:row>');
  });

  it('appendSheet refuses rows past the last row', async () => {
    const src = await crafted('<row r="1048576"><c r="A1048576"><v>1</v></c></row>');
    await expect(edited(src, e => e.appendSheet('S', [[2]]))).rejects.toThrow(/past Excel's last row/);
  });

  it('setCells edits row 1 of a sheet whose rows have no r attribute', async () => {
    const src = await crafted('<row><c><v>1</v></c></row><row><c><v>2</v></c></row>');
    expect(await cells(await edited(src, e => e.setCells('S', { A1: 'edited' })))).toEqual([['edited'], [2]]);
  });

  it('keeps the style of a date written with one, and reuses formats on a repeated restyle', async () => {
    const out = await edited(await write([[1]]), e => e.setCells('S', { A1: { value: new Date(Date.UTC(2026, 0, 1)), style: { font: { bold: true } } } }));
    const style = (await rows(out, { styles: true }))[0].styles?.[0];
    expect(style?.font?.bold).toBe(true);
    expect(style?.numFmt).toBe('yyyy-mm-dd');

    const restyle = (b: Uint8Array) => edited(b, e => e.setCells('S', { A1: { style: { font: { italic: true } } } }));
    const once = await restyle(await write([[1]]));
    const twice = await restyle(once);
    const xfs = (b: Promise<Record<string, string>>) => b.then(p => /<cellXfs count="(\d+)"/.exec(p['xl/styles.xml'])?.[1]);
    expect(await xfs(entries(twice))).toBe(await xfs(entries(once)));
  });

  it('insertRows: ranges to the last row stay at the edge, external references stay, shared formulas follow', async () => {
    const book = await bytesOf(new SheetWriter().addSheet('S', [[{ value: null, formula: 'SUM(B1:B1048576)' }, { value: null, formula: '[1]S!A5+A6' }]]).write());
    const out = await edited(book, e => e.insertRows('S', 3, 1));
    expect((await rows(out, { formulas: true }))[0].formulas).toEqual(['SUM(B1:B1048576)', '[1]S!A5+A7']);

    const parts = await entries(await bytesOf(new SheetWriter().addSheet('Data', [[1]]).addSheet('Calc', [[1]]).write()));
    parts['xl/worksheets/sheet2.xml'] = parts['xl/worksheets/sheet2.xml'].replace(/<sheetData>[\s\S]*<\/sheetData>/,
      '<sheetData><row r="1"><c r="A1"><f t="shared" ref="A1:A5" si="0">Data!A1</f></c></row><row r="5"><c r="A5"><f t="shared" si="0"/></c></row></sheetData>');
    const shifted = await edited(await zipOf(parts), e => e.insertRows('Data', 3, 2));
    expect((await rows(shifted, { sheetName: 'Calc', formulas: true })).map(r => r.formulas?.[0])).toEqual(['Data!A1', 'Data!A7']);
  });
});

describe('.ods', () => {
  it('keeps milliseconds, converts whole columns, leaves streamed formulas uncached, and reads created as UTC', async () => {
    const created = new Date(Date.UTC(2020, 1, 3, 4, 5, 6));
    const when = new Date(Date.UTC(2026, 9, 8, 14, 5, 6, 789));
    const out = await bytesOf(new OdsWriter({ properties: { created } }).addSheet('S', [[when, { value: null, formula: 'SUM(C:C)' }]]).write());
    const content = (await entries(out))['content.xml'];
    expect(content).toContain('2026-10-08T14:05:06.789');
    expect(content).toContain('of:=SUM([.C:.C])');
    expect((await new SheetReader().readWorkbook(blob(out))).properties.created).toEqual(created);
    async function* streamed(): AsyncGenerator<Row> { yield [1]; yield [{ value: null, formula: 'SUM(A1:A1)' }]; }
    expect(await cells(await bytesOf(new OdsWriter().addSheet('S', streamed()).write()))).toEqual([[1], [null]]);
  });
});
