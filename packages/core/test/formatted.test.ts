import { describe, it, expect } from 'vitest';
import { SheetWriter } from '../src/core/writer';
import { SheetReader } from '../src/core/index';
import { ZipStreamWriter } from '../src/core/zip-stream-writer';
import { createBlobReader } from '../src/core/random-access';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());
const blob = (b: Uint8Array) => createBlobReader(new Blob([b]));

async function formatted(bytes: Uint8Array) {
  const rows: { cells: unknown[]; formatted?: (string | undefined)[] }[] = [];
  for await (const row of await new SheetReader().parse(blob(bytes), { formatted: true })) rows.push({ cells: row.cells, formatted: row.formatted });
  return rows;
}

describe('formatted cell text', () => {
  it('shows values the way Excel does, next to the raw values', async () => {
    const out = await bytesOf(new SheetWriter().addSheet('S', [
      [1234.5, { value: 1234.5, style: { numFmt: '#,##0.00' } }, { value: 0.256, style: { numFmt: '0.0%' } }, { value: -42, style: { numFmt: '#,##0;(#,##0)' } }],
      [new Date(Date.UTC(2026, 9, 8)), { value: new Date(Date.UTC(2026, 9, 8, 14, 5)), style: { numFmt: 'dd-mmm-yyyy h:mm AM/PM' } }, { value: 1.5, style: { numFmt: '[h]:mm' } }],
      ['text', true, { value: 'id', style: { numFmt: '"#"@' } }, null, { value: 0.75, style: { numFmt: '# ?/?' } }],
    ]).write());
    const rows = await formatted(out);
    expect(rows[0].cells).toEqual([1234.5, 1234.5, 0.256, -42]);
    expect(rows[0].formatted).toEqual(['1234.5', '1,234.50', '25.6%', '(42)']);
    expect(rows[1].cells[0]).toBe('2026-10-08T00:00:00.000Z');
    expect(rows[1].formatted).toEqual(['2026-10-08', '08-Oct-2026 2:05 PM', '36:00']);
    expect(rows[2].formatted).toEqual(['text', 'TRUE', '#id', undefined, ' 3/4']);

    // Without the option nothing extra is reported
    for await (const row of await new SheetReader().parse(blob(out))) expect(row.formatted).toBeUndefined();
  });

  it('uses en-US display for formats stored only by id, and keeps error values', async () => {
    // Excel writes built-in formats as a numFmtId with no format code
    const zip = new ZipStreamWriter();
    const done = bytesOf(zip.stream);
    const add = (name: string, text: string) => zip.addFile(name, new Blob([text]).stream());
    const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
    const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
    await add('_rels/.rels', rels(`<Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/>`));
    await add('xl/workbook.xml', `<workbook ${ns}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`);
    await add('xl/_rels/workbook.xml.rels', rels(`<Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/>`));
    await add('xl/styles.xml', `<styleSheet ${ns}><cellXfs count="5"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="8"/><xf numFmtId="44"/><xf numFmtId="10"/></cellXfs></styleSheet>`);
    await add('xl/worksheets/sheet1.xml', `<worksheet ${ns}><sheetData><row r="1">` +
      '<c r="A1" s="1"><v>46303</v></c><c r="B1" s="2"><v>-1234.5</v></c><c r="C1" s="3"><v>0</v></c><c r="D1" s="4"><v>0.5</v></c>' +
      '<c r="E1" t="e"><v>#DIV/0!</v></c><c r="F1"><v>123456789012</v></c></row></sheetData></worksheet>');
    await zip.close();
    const [row] = await formatted(await done);
    expect(row.formatted).toEqual(['10/8/2026', '($1,234.50)', ' $-   ', '50.00%', '#DIV/0!', '1.23457E+11']);
  });
});
