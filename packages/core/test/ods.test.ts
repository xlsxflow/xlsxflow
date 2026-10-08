import { describe, it, expect } from 'vitest';
import { OdsWriter } from '../src/core/ods-writer';
import { toOpenFormula } from '../src/core/ods-writer';
import { odsFormula } from '../src/core/ods';
import { SheetReader } from '../src/core/index';
import { createBlobReader } from '../src/core/random-access';

const toBlob = async (s: ReadableStream<Uint8Array>) => new Response(s).blob();

async function readAll(blob: Blob, options = {}) {
  const result = await new SheetReader().parse(createBlobReader(blob), options);
  const rows = [];
  for await (const row of result) rows.push(row);
  return { rows, meta: await result.getMetadata(), comments: await result.getComments() };
}

describe('OpenDocument (.ods)', () => {
  it('round-trips values, dates, formulas, merges and sheets through OdsWriter and the reader', async () => {
    async function* lazy() { yield ['lazy', 1]; yield ['rows', 2]; }
    const writer = new OdsWriter({ properties: { title: 'Report', creator: 'Ana' } });
    writer.addSheet('Data', [
      ['Text', 42, -0.5, true, null, 'two  spaces', 'line\nbreak'],
      [new Date(Date.UTC(2026, 9, 8)), new Date(Date.UTC(2026, 9, 8, 14, 4, 30)), { value: 84, formula: '=B1*2' }, 'tab\there'],
      [],
      ['Merged', null, null, ' leading', 'ctrl\x01char'],
    ], { mergeCells: ['A4:C4'], columnWidths: [20, 10], freezePanes: { row: 1, col: 1 } });
    writer.addSheet('Lazy', lazy());
    writer.addSheet('Secret', [['hidden']], { state: 'hidden' });
    const blob = await toBlob(writer.write());

    const { rows, meta } = await readAll(blob, { formulas: true });
    expect(rows.map(r => [r.rowNumber, r.cells])).toEqual([
      [1, ['Text', 42, -0.5, true, null, 'two  spaces', 'line\nbreak']],
      [2, ['2026-10-08T00:00:00.000Z', '2026-10-08T14:04:30.000Z', 84, 'tab\there']],
      [4, ['Merged', null, null, ' leading', 'ctrlchar']],
    ]);
    expect(rows[1].formulas?.[2]).toBe('B1*2');
    expect(meta.mergedCells).toEqual(['A4:C4']);
    expect(meta.freezePanes).toMatchObject({ row: 1, col: 1 });

    const lazyRows = await readAll(blob, { sheetName: 'Lazy' });
    expect(lazyRows.rows.map(r => r.cells)).toEqual([['lazy', 1], ['rows', 2]]);

    const info = await new SheetReader().readWorkbook(createBlobReader(blob));
    expect(info.sheets).toEqual([{ name: 'Data', state: 'visible' }, { name: 'Lazy', state: 'visible' }, { name: 'Secret', state: 'hidden' }]);
    expect(info.properties).toMatchObject({ title: 'Report', creator: 'Ana' });
  });

  it('starts the package with an uncompressed mimetype entry', async () => {
    const bytes = new Uint8Array(await (await toBlob(new OdsWriter().addSheet('A', [[1]]).write())).arrayBuffer());
    const head = new TextDecoder().decode(bytes.subarray(30, 30 + 8 + 46));
    expect(head).toBe('mimetypeapplication/vnd.oasis.opendocument.spreadsheet');
  });

  it('converts formulas between Excel and OpenFormula syntax', () => {
    expect(toOpenFormula('=SUM(A1:B2,Sheet2!C3)')).toBe('of:=SUM([.A1:.B2];[$Sheet2.C3])');
    expect(toOpenFormula("='My Sheet'!$A$1&\"a,b A1\"")).toBe("of:=[$'My Sheet'.$A$1]&\"a,b A1\"");
    expect(toOpenFormula('LOG10(A1)+ATAN2(1,2)')).toBe('of:=LOG10([.A1])+ATAN2(1;2)');
    expect(odsFormula('of:=SUM([.A1:.B2];[$Sheet2.C3])')).toBe('SUM(A1:B2,Sheet2!C3)');
    expect(odsFormula("of:=[$'My Sheet'.$A$1]&\"a;b\"")).toBe("'My Sheet'!$A$1&\"a;b\"");
  });

  it('does not expand empty repeated rows and cells', async () => {
    // LibreOffice pads sheets to their full size with repeated empty rows
    const { ZipStreamWriter } = await import('../src/core/zip-stream-writer');
    const zip = new ZipStreamWriter();
    const text = (s: string) => new ReadableStream({ start(c) { c.enqueue(new TextEncoder().encode(s)); c.close(); } });
    (async () => {
      await zip.addFile('mimetype', text('application/vnd.oasis.opendocument.spreadsheet'));
      await zip.addFile('content.xml', text('<office:document-content xmlns:office="o" xmlns:table="t" xmlns:text="x"><office:body><office:spreadsheet>'
        + '<table:table table:name="S"><table:table-row><table:table-cell table:number-columns-repeated="16384"/></table:table-row>'
        + '<table:table-row table:number-rows-repeated="1048575"><table:table-cell table:number-columns-repeated="16383"/>'
        + '<table:table-cell office:value-type="float" office:value="1"><text:p>1</text:p></table:table-cell></table:table-row>'
        + '</table:table></office:spreadsheet></office:body></office:document-content>'));
      await zip.close();
    })();
    const blob = await toBlob(zip.stream);
    await expect(readAll(blob, { maxUncompressedBytes: 1_000_000 })).rejects.toThrow(/padded/);
  });
});
