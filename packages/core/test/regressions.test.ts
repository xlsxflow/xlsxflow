import { describe, it, expect } from 'vitest';
import { SheetWriter } from '../src/core/writer';
import { SheetEditor } from '../src/core/editor';
import { SheetReader } from '../src/core/index';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { ZipStreamWriter } from '../src/core/zip-stream-writer';
import { createBlobReader } from '../src/core/random-access';
import { resolveSheetPaths } from '../src/core/utils';
import { Row } from '../src/core/types';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());

async function entryText(bytes: Uint8Array, name: string): Promise<string> {
  const zip = new ZipRandomAccessParser(createBlobReader(new Blob([bytes])));
  await zip.parseCentralDirectory();
  return new Response(await zip.extractStream(name)).text();
}

async function readRows(bytes: Uint8Array, sheetName?: string) {
  const rows = [];
  for await (const r of await new SheetReader().parse(createBlobReader(new Blob([bytes])), sheetName ? { sheetName } : undefined)) rows.push(r.cells);
  return rows;
}

describe('regressions', () => {
  it('styles point at the custom font/fill, not the defaults', async () => {
    const bytes = await bytesOf(new SheetWriter().write([[{ value: 'x', style: { font: { bold: true }, fill: { type: 'solid', fgColor: 'FFFF0000' } } }]]));
    const styles = await entryText(bytes, 'xl/styles.xml');
    expect(styles).toContain('fontId="1" fillId="2"');
    expect(styles).toContain('<fonts count="2">');
    expect(styles).toContain('<fills count="3">');
  });

  it('styles registered by AsyncIterable rows reach styles.xml', async () => {
    async function* rows(): AsyncGenerator<Row> {
      yield [{ value: 1, style: { font: { italic: true } } }];
    }
    const writer = new SheetWriter();
    writer.addSheet('S', rows());
    const bytes = await bytesOf(writer.write());
    expect(await entryText(bytes, 'xl/styles.xml')).toContain('<i/>');
    expect(await entryText(bytes, 'xl/worksheets/sheet1.xml')).toContain('s="1"');
  });

  it('formulas resolve against their own sheet', async () => {
    const writer = new SheetWriter();
    writer.addSheet('A', [[10], [{ value: null, formula: 'SUM(A1:A1)' }]]);
    writer.addSheet('B', [[99], [{ value: null, formula: 'SUM(A1:A1)' }]]);
    const bytes = await bytesOf(writer.write());
    expect(await entryText(bytes, 'xl/worksheets/sheet1.xml')).toContain('<v>10</v></c>');
    expect(await entryText(bytes, 'xl/worksheets/sheet2.xml')).toContain('<f>SUM(A1:A1)</f><v>99</v>');
  });

  it('writer errors reach the stream consumer instead of hanging', async () => {
    async function* bad(): AsyncGenerator<Row> {
      yield [1];
      throw new Error('boom');
    }
    const writer = new SheetWriter();
    writer.addSheet('S', bad());
    await expect(bytesOf(writer.write())).rejects.toThrow('boom');
  });

  it('ZIP writer waits for the consumer instead of buffering everything', async () => {
    const zip = new ZipStreamWriter(1024);
    let produced = 0;
    const big = new ReadableStream<Uint8Array>({
      pull(c) {
        if (produced >= 64) return c.close();
        produced++;
        c.enqueue(new Uint8Array(16 * 1024));
      }
    });
    // Pass-through path: Node's CompressionStream has its own deep internal queue, which would hide the stall
    const done = zip.addCompressedFile('big.bin', big, 64 * 16 * 1024, 64 * 16 * 1024, 0, 0);
    await new Promise(r => setTimeout(r, 50));
    expect(produced).toBeLessThan(64); // stalled: nobody is reading
    const reader = zip.stream.getReader();
    const drain = (async () => { while (!(await reader.read()).done); })();
    await done;
    await zip.close();
    await drain;
    expect(produced).toBe(64);
  });

  it('resolves absolute targets, attribute order and escaped names', () => {
    const paths = resolveSheetPaths(
      `<x:sheets><x:sheet r:id="rId7" sheetId="1" name="P&amp;L"/></x:sheets>`,
      `<Relationships><Relationship Target="/xl/worksheets/data.xml" Id="rId7" Type="t"/></Relationships>`
    );
    expect(paths.get('P&L')).toBe('xl/worksheets/data.xml');
  });

  it('reads the first tab by default, not sheet1.xml', async () => {
    const writer = new SheetWriter();
    writer.addSheet('First', [['first']]).addSheet('Second', [['second']]);
    const bytes = await bytesOf(writer.write());
    expect(await readRows(bytes)).toEqual([['first']]);
    expect(await readRows(bytes, 'Second')).toEqual([['second']]);
  });

  it('editor appends to an empty <sheetData/> and passes other entries through', async () => {
    const original = await bytesOf(new SheetWriter().write([]));
    expect(await entryText(original, 'xl/worksheets/sheet1.xml')).toMatch(/<sheetData>\s*<\/sheetData>/);

    // Force the self-closing form a real Excel file uses for empty sheets
    const zip = new ZipStreamWriter();
    const src = new ZipRandomAccessParser(createBlobReader(new Blob([original])));
    await src.parseCentralDirectory();
    const building = (async () => {
      for (const f of src.getFiles()) {
        let text = await new Response(await src.extractStream(f)).text();
        if (f === 'xl/worksheets/sheet1.xml') text = text.replace(/<sheetData>\s*<\/sheetData>/, '<sheetData/>');
        await zip.addFile(f, new Response(text).body!);
      }
      await zip.close();
    })();
    const emptySheetFile = await bytesOf(zip.stream);
    await building;

    const editor = new SheetEditor();
    editor.appendSheet('Sheet1', [['a', 1]]);
    const edited = await bytesOf(editor.edit(createBlobReader(new Blob([emptySheetFile]))));
    expect(await readRows(edited)).toEqual([['a', 1]]);
    expect(await entryText(edited, 'xl/styles.xml')).toBe(await entryText(emptySheetFile, 'xl/styles.xml'));
  });

  it('size limit guards in-memory parts; the streamed sheet only when asked', async () => {
    const bytes = await bytesOf(new SheetWriter().write([['a'.repeat(5000)]]));
    const parse = (maxUncompressedBytes: number) =>
      new SheetReader().parse(createBlobReader(new Blob([bytes])), { maxUncompressedBytes });
    await expect(parse(100)).rejects.toThrow(/maximum uncompressed size/); // workbook parts
    expect((await readRows(bytes))[0][0]).toHaveLength(5000); // default: sheet unlimited
    const sheetSize = (await entryText(bytes, 'xl/worksheets/sheet1.xml')).length;
    const rows = await parse(sheetSize - 1);
    await expect((async () => { for await (const _ of rows); })()).rejects.toThrow(/maximum uncompressed size/);
  });

  it('editor rejects unknown sheet names instead of silently skipping', async () => {
    const original = await bytesOf(new SheetWriter().write([[1]]));
    const editor = new SheetEditor();
    editor.appendSheet('Nope', [[2]]);
    await expect(bytesOf(editor.edit(createBlobReader(new Blob([original]))))).rejects.toThrow('Nope');
  });
});

describe('formula cached values', () => {
  it('keep their type', async () => {
    const bytes = await bytesOf(new SheetWriter().write([[1, 0], [
      { value: null, formula: 'A1>0' }, { value: null, formula: 'A1<0' },
      { value: null, formula: 'CONCATENATE("a","b")' }, { value: null, formula: 'A1/B1' },
    ]]));
    expect(await readRows(bytes)).toEqual([[1, 0], [true, false, 'ab', '#DIV/0!']]);
  });
});

// Found by end-user edge-case tests against the published 1.1.1
describe('end-user edge cases', () => {
  it('refuses cells past Excel\'s last column and rows past its last row', async () => {
    const wide = new Array(16385).fill(null);
    await expect(bytesOf(new SheetWriter().write([wide]))).rejects.toThrow(/16384 columns/);
    await expect(bytesOf(new SheetWriter().write([new Array(16384).fill(1)]))).resolves.toBeInstanceOf(Uint8Array);
    async function* rows(): AsyncGenerator<Row> { for (let i = 0; i < 1048577; i++) yield []; }
    const writer = new SheetWriter();
    writer.addSheet('S', rows());
    await expect(bytesOf(writer.write())).rejects.toThrow(/last row \(1048576\)/);
  }, 60000);

  it('OdsWriter stores formula results', async () => {
    const { OdsWriter } = await import('../src/core/ods-writer');
    const bytes = await bytesOf(new OdsWriter().addSheet('S', [[1, 2, { value: null, formula: '=A1+B1' }, { value: null, formula: 'A1/0' }]]).write());
    expect(await readRows(bytes)).toEqual([[1, 2, 3, '#DIV/0!']]);
  });

  it('sheetToJson keeps columns with repeated headers', async () => {
    const { sheetToJson } = await import('../src/core/utils');
    const bytes = await bytesOf(new SheetWriter().write([['Name', 'Name', null, 'Name'], ['a', 'b', 'c', 'd']]));
    const json = await sheetToJson(await new SheetReader().parse(createBlobReader(new Blob([bytes]))));
    expect(json).toEqual([{ Name: 'a', Name_2: 'b', Column3: 'c', Name_3: 'd' }]);
  });

  it('parseCsv gives null for empty unquoted fields with or without convert', async () => {
    const { parseCsv } = await import('../src/core/csv');
    const all = async (it: AsyncIterable<unknown[]>) => { const out = []; for await (const r of it) out.push(r); return out; };
    expect(await all(parseCsv('a,,""\n'))).toEqual([['a', null, '']]);
    expect(await all(parseCsv('a,,""\n', { convert: false }))).toEqual([['a', null, '']]);
  });

  it('SheetEditor gives a date in a General cell a date format, and keeps an existing one', async () => {
    const src = await bytesOf(new SheetWriter().write([['x', { value: 1, style: { numFmt: 'dd-mmm-yyyy' } }]]));
    const d = new Date(Date.UTC(2026, 9, 8));
    const out = await bytesOf(new SheetEditor().setCells('Sheet1', { A1: d, B1: d, C1: new Date(Date.UTC(2026, 9, 8, 14, 30)) }).edit(createBlobReader(new Blob([src]))));
    const rows = [];
    for await (const r of await new SheetReader().parse(createBlobReader(new Blob([out])), { formatted: true })) rows.push(r.formatted);
    expect(rows[0]).toEqual(['2026-10-08', '08-Oct-2026', '2026-10-08 14:30:00']);
  });
});

describe('end-user edge cases, round 2', () => {
  const blob = (b: Uint8Array) => createBlobReader(new Blob([b]));

  it('SheetEditor.setCells refuses addresses past Excel\'s last column and row', async () => {
    for (const ref of ['XFE1', 'A1048577']) expect(() => new SheetEditor().setCells('Sheet1', { [ref]: 1 })).toThrow(/Invalid cell reference/);
    expect(() => new SheetEditor().setCells('Sheet1', { XFD1048576: 1 })).not.toThrow();
  });

  it('refuses text longer than Excel\'s 32,767 characters per cell', async () => {
    const long = 'x'.repeat(32768);
    await expect(bytesOf(new SheetWriter().write([[long]]))).rejects.toThrow(/Cell A1 has 32768 characters/);
    await expect(bytesOf(new SheetWriter({ sharedStrings: true }).write([['x'.repeat(32767)]]))).resolves.toBeInstanceOf(Uint8Array);
    const src = await bytesOf(new SheetWriter().write([[1]]));
    await expect(bytesOf(new SheetEditor().setCells('Sheet1', { B1: long }).edit(blob(src)))).rejects.toThrow(/32,?767/);
  });

  it('checks colours: ARGB and RGB hex pass, other text throws', async () => {
    const bytes = await bytesOf(new SheetWriter().write([[{ value: 1, style: { font: { color: '#ff0000' }, fill: { type: 'solid', fgColor: '00FF00' } } }]]));
    const styles = await entryText(bytes, 'xl/styles.xml');
    expect(styles).toContain('<color rgb="FFFF0000"/>');
    expect(styles).toContain('<fgColor rgb="FF00FF00"/>');
    await expect(bytesOf(new SheetWriter().write([[{ value: 1, style: { font: { color: 'red' } } }]]))).rejects.toThrow(/Invalid colour "red"/);
  });

  it('a shared string index past the table reads as an empty cell', async () => {
    const original = await bytesOf(new SheetWriter({ sharedStrings: true }).write([['a', 'b']]));
    const zip = new ZipStreamWriter();
    const src = new ZipRandomAccessParser(blob(original));
    await src.parseCentralDirectory();
    const building = (async () => {
      for (const f of src.getFiles()) {
        let text = await new Response(await src.extractStream(f)).text();
        if (f === 'xl/worksheets/sheet1.xml') text = text.replace('<v>1</v>', '<v>99</v>');
        await zip.addFile(f, new Response(text).body!);
      }
      await zip.close();
    })();
    const crafted = await bytesOf(zip.stream);
    await building;
    expect(await readRows(crafted)).toEqual([['a', null]]);
  });

  it('drawing anchors refuse cells past the sheet and put a reversed range the right way round', async () => {
    const { anchorXml } = await import('../src/core/image');
    const size = { width: 10, height: 10 };
    expect(() => anchorXml({ at: 'XFE1' }, size, '')).toThrow(/Invalid anchor cell "XFE1"/);
    expect(() => anchorXml({ range: 'A1:A1048577' }, size, '')).toThrow(/Invalid anchor cell/);
    expect(anchorXml({ range: 'F20:A2' }, size, '')).toBe(anchorXml({ range: 'A2:F20' }, size, ''));
  });
});
