import { describe, it, expect } from 'vitest';
import { deflateRawSync, crc32 } from 'node:zlib';
import { SheetReader, SheetEditor, SheetWriter, createBlobReader, sheetToJson } from '../src/index';
import { formatValue } from '../src/core/number-format';

// Hand-built ZIPs, so tests can state sizes that lie and point several names at one entry
function buildZip(entries: Record<string, string>, opts: { alias?: [string, string][]; statedSize?: Record<string, number> } = {}): Uint8Array {
  const parts: Buffer[] = [], cds: Buffer[] = [], info = new Map<string, { crc: number; c: number; u: number; off: number }>();
  let off = 0;
  for (const [name, text] of Object.entries(entries)) {
    const raw = Buffer.from(text), comp = deflateRawSync(raw), nb = Buffer.from(name), crc = crc32(raw);
    const lh = Buffer.alloc(30);
    lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(8, 8);
    lh.writeUInt32LE(crc, 14); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(raw.length, 22); lh.writeUInt16LE(nb.length, 26);
    info.set(name, { crc, c: comp.length, u: opts.statedSize?.[name] ?? raw.length, off });
    parts.push(lh, nb, comp);
    off += 30 + nb.length + comp.length;
  }
  const all = [...Object.keys(entries).map(n => [n, n]), ...(opts.alias ?? [])];
  for (const [name, target] of all) {
    const i = info.get(target)!, nb = Buffer.from(name), cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6); cd.writeUInt16LE(8, 10);
    cd.writeUInt32LE(i.crc, 16); cd.writeUInt32LE(i.c, 20); cd.writeUInt32LE(i.u, 24); cd.writeUInt16LE(nb.length, 28); cd.writeUInt32LE(i.off, 42);
    cds.push(cd, nb);
  }
  const cd = Buffer.concat(cds), eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(all.length, 8); eocd.writeUInt16LE(all.length, 10);
  eocd.writeUInt32LE(cd.length, 12); eocd.writeUInt32LE(off, 16);
  return new Uint8Array(Buffer.concat([...parts, cd, eocd]));
}
const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const book = (parts: Record<string, string>, opts?: Parameters<typeof buildZip>[1]) => buildZip({
  '_rels/.rels': `<Relationships><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  'xl/workbook.xml': '<workbook><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels': `<Relationships><Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
  'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
  ...parts,
}, opts);
const blob = (b: Uint8Array) => createBlobReader(new Blob([b]));
const drain = async (bytes: Uint8Array, options?: Parameters<SheetReader['parse']>[1]) => {
  const result = await new SheetReader().parse(blob(bytes), options);
  for await (const _ of result);
  return result;
};
const fast = async (work: () => Promise<unknown>, ms = 2000) => {
  const t = performance.now();
  await work();
  expect(performance.now() - t).toBeLessThan(ms);
};

describe('hostile files', () => {
  it('rejects a central directory that points outside the file instead of allocating it', async () => {
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(1, 10); eocd.writeUInt32LE(0xfffffffe, 12);
    await expect(new SheetReader().parse(blob(new Uint8Array(eocd)))).rejects.toThrow(/Corrupt ZIP/);
  });

  it('rejects entries that share bytes and entries that inflate past their stated size', async () => {
    const aliased = book({}, { alias: [['xl/worksheets/sheet2.xml', 'xl/worksheets/sheet1.xml']] });
    await expect(new SheetReader().parse(blob(aliased))).rejects.toThrow(/overlap/);
    const lying = book({ 'xl/worksheets/sheet1.xml': `<worksheet><sheetData>${' '.repeat(100_000)}</sheetData></worksheet>` },
      { statedSize: { 'xl/worksheets/sheet1.xml': 100 } });
    await expect(drain(lying)).rejects.toThrow(/stated size/);
  });

  it('stays linear on tags and elements that never close', async () => {
    await fast(() => drain(book({ 'xl/worksheets/sheet1.xml': `<worksheet><sheetData><row r="1"><c ${'a'.repeat(200_000)}/></row></sheetData></worksheet>` })));
    const props = book({
      '_rels/.rels': `<Relationships><Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/>` +
        `<Relationship Id="rId2" Type="${rel}/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
      'docProps/core.xml': '<dc:title>'.repeat(100_000),
      'xl/workbook.xml': '<workbook><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames>' + '<definedName name="a">'.repeat(50_000),
    });
    await fast(() => new SheetReader().readWorkbook(blob(props)));
    const sheet = '<worksheet><sheetData><row r="1">' + '<c r="A1">'.repeat(100_000);
    await fast(() => new Response(new SheetEditor().setCells('S', { A1: 2 }).edit(blob(book({ 'xl/worksheets/sheet1.xml': sheet })))).arrayBuffer()
      .catch(() => undefined));
  });

  it('keeps tiny far-right cells and repeated hidden columns from blowing up memory', async () => {
    const cols = '<cols>' + '<col min="1" max="16384" hidden="1"/>'.repeat(20_000) + '</cols>';
    const result = await drain(book({ 'xl/worksheets/sheet1.xml': `<worksheet>${cols}<sheetData/></worksheet>` }));
    expect((await result.getMetadata()).hiddenCols).toHaveLength(16384);
    const wide = book({ 'xl/worksheets/sheet1.xml': `<worksheet><sheetData>${'<row><c r="XFD1"/></row>'.repeat(2000)}</sheetData></worksheet>` });
    await expect(drain(wide, { maxUncompressedBytes: 1_000_000 })).rejects.toThrow(/padded/);
  });

  it('formats fractions with long denominators without trying every denominator', async () => {
    await fast(async () => expect(formatValue(0.1234567890123, '# ?/??????????')).toMatch(/^ \d+\/\d+ *$/));
  });

  it('caps whole parts the editor reads when asked', async () => {
    const big = book({ 'xl/workbook.xml': '<workbook><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets>' + ' '.repeat(2_000_000) + '</workbook>' });
    await expect(new Response(new SheetEditor().setCells('S', { A1: 2 }).edit(blob(big), { maxUncompressedBytes: 1_000_000 })).arrayBuffer())
      .rejects.toThrow(/over the limit/);
  });

  it('survives out-of-range character references and a __proto__ header', async () => {
    const names = book({ 'xl/workbook.xml': '<workbook><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets><definedNames><definedName name="&#x110000;">1</definedName></definedNames></workbook>' });
    expect((await new SheetReader().readWorkbook(blob(names))).definedNames[0].name).toBe('&#x110000;');
    const out = await new Response(new SheetWriter().addSheet('S', [['__proto__'], ['x']]).write()).arrayBuffer();
    const rows = await sheetToJson(await new SheetReader().parse(blob(new Uint8Array(out))));
    expect(Object.keys(rows[0])).toEqual(['__proto__']);
    expect(({} as Record<string, unknown>).x).toBeUndefined();
  });
});

describe('hostile options', () => {
  it('escapes option values typed as enums, which JSON or JS callers can set to anything', async () => {
    const evil = 'x"/><evil/><y a="' as never;
    const out = await new Response(new SheetWriter().addSheet('S', [[{ value: 1, style: { border: { left: { style: evil } }, alignment: { horizontal: evil } } }]],
      { dataValidations: [{ sqref: 'A1', type: evil, formula1: '1' }] }).write()).arrayBuffer();
    const text = new TextDecoder().decode(new Uint8Array(out));
    expect(text).not.toContain('<evil/>');
  });

  it('encodes control characters in rows appended by the editor', async () => {
    const editor = new SheetEditor();
    editor.appendSheet('S', [['a\u0001b']]);
    const edited = await new Response(editor.edit(blob(book({})))).arrayBuffer();
    const rows: unknown[][] = [];
    for await (const r of await new SheetReader().parse(blob(new Uint8Array(edited)))) rows.push(r.cells);
    expect(rows[1]).toEqual(['a\u0001b']);
  });
});
