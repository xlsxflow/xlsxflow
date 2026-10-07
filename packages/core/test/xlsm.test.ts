import { describe, it, expect } from 'vitest';
import { SheetWriter } from '../src/core/writer';
import { SheetEditor } from '../src/core/editor';
import { SheetReader } from '../src/core/index';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { ZipStreamWriter } from '../src/core/zip-stream-writer';
import { createBlobReader } from '../src/core/random-access';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());
const blob = (b: Uint8Array) => createBlobReader(new Blob([b]));
const VBA = new Uint8Array(Array.from({ length: 3000 }, (_, i) => (i * 37) % 256));
const MACRO_WORKBOOK = 'application/vnd.ms-excel.sheet.macroEnabled.main+xml';

async function entries(bytes: Uint8Array): Promise<Map<string, Uint8Array>> {
  const zip = new ZipRandomAccessParser(blob(bytes));
  await zip.parseCentralDirectory();
  const out = new Map<string, Uint8Array>();
  for (const name of zip.getFiles()) out.set(name, await bytesOf(await zip.extractStream(name)));
  return out;
}

// A workbook from SheetWriter turned into an .xlsm: macro-enabled content type and a VBA project part
async function xlsm(): Promise<Uint8Array> {
  const w = new SheetWriter();
  w.addSheet('Data', [['a', 1], ['b', 2], ['c', 3]]);
  w.addSheet('Other', [[{ value: null, formula: 'Data!B2*2' }]]);
  const parts = await entries(await bytesOf(w.write()));
  const text = (name: string) => new TextDecoder().decode(parts.get(name));
  parts.set('[Content_Types].xml', new TextEncoder().encode(text('[Content_Types].xml')
    .replace(/application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet\.main\+xml/, MACRO_WORKBOOK)
    .replace(/<Default /, '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ')));
  parts.set('xl/_rels/workbook.xml.rels', new TextEncoder().encode(text('xl/_rels/workbook.xml.rels').replace('</Relationships>',
    '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>')));
  parts.set('xl/vbaProject.bin', VBA);
  const zip = new ZipStreamWriter();
  const done = bytesOf(zip.stream);
  for (const [name, data] of parts) await zip.addFile(name, new Blob([data]).stream());
  await zip.close();
  return done;
}

describe('macro-enabled workbooks', () => {
  it('keep their VBA project through every edit', async () => {
    const src = await xlsm();
    const edits: [string, SheetEditor][] = [
      ['setCells', new SheetEditor().setCells('Data', { A1: { value: 'x' }, C1: { value: 5, style: { font: { bold: true } } } })],
      ['insertRows', new SheetEditor().insertRows('Data', 2, 2)],
      ['deleteColumns', new SheetEditor().deleteColumns('Data', 'A')],
      ['addSheet', new SheetEditor().addSheet('New', [[1]])],
      ['deleteSheet', new SheetEditor().deleteSheet('Other')],
    ];
    for (const [name, editor] of edits) {
      const out = await entries(await bytesOf(editor.edit(blob(src))));
      const text = (part: string) => new TextDecoder().decode(out.get(part));
      expect(out.get('xl/vbaProject.bin'), name).toEqual(VBA);
      expect(text('[Content_Types].xml'), name).toContain(MACRO_WORKBOOK);
      expect(text('[Content_Types].xml'), name).toContain('Extension="bin" ContentType="application/vnd.ms-office.vbaProject"');
      expect(text('xl/_rels/workbook.xml.rels'), name).toContain('Target="vbaProject.bin"');
    }
    const edited = await bytesOf(new SheetEditor().setCells('Data', { A1: { value: 'x' } }).edit(blob(src)));
    const rows: unknown[][] = [];
    for await (const row of await new SheetReader().parse(blob(edited))) rows.push(row.cells);
    expect(rows[0]).toEqual(['x', 1]);
  });
});
