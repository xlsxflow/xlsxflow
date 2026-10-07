import { describe, it, expect } from 'vitest';
import { SheetWriter, legacyPasswordHash } from '../src/core/writer';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { createBlobReader } from '../src/core/random-access';
import type { Row } from '../src/core/types';

async function entries(stream: ReadableStream<Uint8Array>) {
  const zip = new ZipRandomAccessParser(createBlobReader(await new Response(stream).blob()));
  await zip.parseCentralDirectory();
  return async (name: string) => (zip.has(name) ? new Response(await zip.extractStream(name)).text() : '');
}

const red = { fill: { type: 'solid' as const, fgColor: 'FFFFC7CE' }, font: { color: 'FF9C0006' } };

describe('conditional formats', () => {
  it('writes highlight rules with differential styles', async () => {
    const w = new SheetWriter();
    w.addSheet('S', [[1], [2], [3]], {
      conditionalFormats: [
        { range: 'A1:A3', rule: { type: 'cellIs', operator: 'greaterThan', formulae: [1], style: red } },
        { range: 'A1:A3', rule: { type: 'cellIs', operator: 'between', formulae: [1, '$B$1'], style: red } },
        { range: 'B2:B9', rule: { type: 'containsText', text: 'a"b', style: { font: { bold: true } } } },
        { range: 'A1:A3', rule: { type: 'expression', formula: '=MOD(A1,2)=0', style: { numFmt: '0.0%' } } },
        { range: 'A1:A3', rule: { type: 'top10', rank: 2, bottom: true, style: red } },
        { range: 'A1:A3', rule: { type: 'aboveAverage', below: true, style: red } },
        { range: 'A1:A3', rule: { type: 'duplicateValues', style: red } },
        { range: 'A1:A3', rule: { type: 'iconSet', iconSet: '3TrafficLights1', showValue: false } },
      ],
    });
    const read = await entries(w.write());
    const sheet = await read('xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule>');
    expect(sheet).toContain('<formula>1</formula><formula>$B$1</formula>');
    expect(sheet).toContain('<cfRule type="containsText" dxfId="1" priority="3" operator="containsText" text="a&quot;b"><formula>NOT(ISERROR(SEARCH(&quot;a&quot;&quot;b&quot;,B2)))</formula>');
    expect(sheet).toContain('<formula>MOD(A1,2)=0</formula>');
    expect(sheet).toContain('type="top10" dxfId="0" priority="5" bottom="1" rank="2"');
    expect(sheet).toContain('type="aboveAverage" dxfId="0" priority="6" aboveAverage="0"');
    expect(sheet).toContain('<iconSet iconSet="3TrafficLights1" showValue="0"><cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="percent" val="67"/></iconSet>');

    const styles = await read('xl/styles.xml');
    expect(styles).toMatch(/<dxfs count="3"><dxf><font><color rgb="FF9C0006"\/><\/font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"\/><bgColor rgb="FFFFC7CE"\/>/);
    expect(styles).toMatch(/<numFmt numFmtId="164" formatCode="0.0%"\/>[\s\S]*<dxf><numFmt numFmtId="164" formatCode="0.0%"\/><\/dxf>/);
  });
});

describe('comments', () => {
  it('writes notes with their VML shapes', async () => {
    const w = new SheetWriter();
    w.addSheet('Plain', [['x']]);
    w.addSheet('Notes', [[{ value: 'a', comment: 'Check & fix' }, { value: 1, comment: { text: 'From Ana', author: 'Ana' } }]]);
    const read = await entries(w.write());
    const comments = await read('xl/comments2.xml');
    expect(comments).toContain('<authors><author></author><author>Ana</author></authors>');
    expect(comments).toContain('<comment ref="A1" authorId="0"><text><r><t>Check &amp; fix</t></r></text></comment>');
    expect(comments).toContain('<comment ref="B1" authorId="1">');
    const vml = await read('xl/drawings/vmlDrawing2.vml');
    expect(vml.match(/ObjectType="Note"/g)).toHaveLength(2);
    expect(vml).toContain('<x:Row>0</x:Row><x:Column>1</x:Column>');
    expect(await read('xl/worksheets/sheet2.xml')).toContain('<legacyDrawing r:id="rIdVml"/>');
    expect(await read('xl/worksheets/_rels/sheet2.xml.rels')).toContain('Target="../comments2.xml"');
    expect(await read('xl/worksheets/_rels/sheet1.xml.rels')).toBe('');
    const types = await read('[Content_Types].xml');
    expect(types).toContain('/xl/comments2.xml');
    expect(types).toContain('Extension="vml"');
  });

  it('collects notes from streamed rows', async () => {
    async function* rows(): AsyncGenerator<Row> { yield [{ value: 1, comment: 'late' }]; }
    const read = await entries(new SheetWriter().addSheet('S', rows()).write());
    expect(await read('xl/comments1.xml')).toContain('late');
    expect(await read('[Content_Types].xml')).toContain('/xl/comments1.xml');
  });
});

describe('sheet layout options', () => {
  it('writes rows, columns, protection, page setup and tab colour in schema order', async () => {
    const w = new SheetWriter();
    w.addSheet('Report', [['h1', 'h2'], [1, 2], [3, 4]], {
      columnWidths: [20],
      columns: [{ outlineLevel: 1 }, { width: 12, hidden: true }],
      rows: { 2: { height: 30, outlineLevel: 1 }, 3: { hidden: true, outlineLevel: 1 }, 6: { height: 40 } },
      protection: { password: 'secret', sort: true },
      pageSetup: {
        orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0, margins: { left: 0.5 },
        header: '&CQuarterly & more', footer: '&RPage &P of &N', printArea: 'A1:B3', printTitleRows: '1', gridLines: true,
      },
      tabColor: 'FF00B050',
    });
    const read = await entries(w.write());
    const sheet = await read('xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<sheetPr><tabColor rgb="FF00B050"/><pageSetUpPr fitToPage="1"/></sheetPr>');
    expect(sheet).toContain('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1" outlineLevelCol="1"/>');
    expect(sheet).toContain('<cols><col min="1" max="1" width="20" customWidth="1" outlineLevel="1"/><col min="2" max="2" width="12" customWidth="1" hidden="1"/></cols>');
    expect(sheet).toContain('<row r="2" ht="30" customHeight="1" outlineLevel="1">');
    expect(sheet).toContain('<row r="3" hidden="1" outlineLevel="1">');
    expect(sheet).toContain('<row r="6" ht="40" customHeight="1"/>');
    expect(sheet).toContain(`<sheetProtection password="${legacyPasswordHash('secret')}" sheet="1" objects="1" scenarios="1" sort="0"/>`);
    expect(sheet).toContain('<pageMargins left="0.5" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');
    expect(sheet).toContain('<pageSetup paperSize="9" fitToWidth="1" fitToHeight="0" orientation="landscape"/>');
    expect(sheet).toContain('<oddHeader>&amp;CQuarterly &amp; more</oddHeader>');
    expect(sheet.indexOf('<sheetProtection')).toBeLessThan(sheet.indexOf('<printOptions'));
    const wb = await read('xl/workbook.xml');
    expect(wb).toContain(`<definedName name="_xlnm.Print_Area" localSheetId="0">&apos;Report&apos;!$A$1:$B$3</definedName>`);
    expect(wb).toContain(`<definedName name="_xlnm.Print_Titles" localSheetId="0">&apos;Report&apos;!$1:$1</definedName>`);
  });

  it('hashes passwords like Excel', () => {
    // Known values of Excel's legacy sheet-protection hash
    expect(legacyPasswordHash('test')).toBe('CBEB'); // openpyxl's hash_password agrees
    expect(legacyPasswordHash('secret')).toBe('DAA7');
    expect(legacyPasswordHash('')).toBe('CE4B');
  });

  it('writes validation operators and messages', async () => {
    const read = await entries(new SheetWriter().addSheet('S', [[1]], {
      dataValidations: [{ sqref: 'A1:A9', type: 'whole', operator: 'between', formula1: '1', formula2: '10', error: 'Whole numbers 1-10', errorTitle: 'Nope', prompt: 'Enter 1-10' }],
    }).write());
    expect(await read('xl/worksheets/sheet1.xml')).toContain(
      '<dataValidation sqref="A1:A9" type="whole" operator="between" showErrorMessage="1" showInputMessage="1" errorTitle="Nope" error="Whole numbers 1-10" prompt="Enter 1-10"><formula1>1</formula1><formula2>10</formula2></dataValidation>');
  });
});

describe('tables', () => {
  it('writes tables with header names from the rows', async () => {
    const w = new SheetWriter();
    w.addSheet('Data', [['Name', 'Amount & tax'], ['a', 1], ['b', 2]], { tables: [{ name: 'Sales', ref: 'A1:B3' }] });
    w.addSheet('More', [['X'], [1]], { tables: [{ name: 'Other', ref: 'A1:A2', style: 'TableStyleLight9', columns: ['X'] }] });
    const read = await entries(w.write());
    const t1 = await read('xl/tables/table1.xml');
    expect(t1).toContain('id="1" name="Sales" displayName="Sales" ref="A1:B3"');
    expect(t1).toContain('<tableColumn id="1" name="Name"/><tableColumn id="2" name="Amount &amp; tax"/>');
    expect(await read('xl/tables/table2.xml')).toContain('name="TableStyleLight9"');
    expect(await read('xl/worksheets/sheet2.xml')).toContain('<tableParts count="1"><tablePart r:id="rIdTable1"/></tableParts>');
    expect(await read('xl/worksheets/_rels/sheet2.xml.rels')).toContain('Target="../tables/table2.xml"');
    expect(await read('[Content_Types].xml')).toContain('/xl/tables/table2.xml');
  });

  it('rejects bad tables', () => {
    const write = (rows: Row[], tables: { name: string; ref: string; columns?: string[] }[]) => () => {
      const w = new SheetWriter();
      w.addSheet('S', rows, { tables });
      return (w as unknown as { planTables(): unknown }).planTables();
    };
    expect(write([['a', 'a']], [{ name: 'T', ref: 'A1:B2' }])).toThrow(/unique, non-empty/);
    expect(write([['a']], [{ name: 'T', ref: 'A1:A2' }, { name: 't', ref: 'A1:A2' }])).toThrow(/duplicate table name/);
    expect(write([['a']], [{ name: 'has space', ref: 'A1:A2' }])).toThrow(/Invalid or duplicate/);
    expect(write([['a']], [{ name: 'T', ref: 'A1:B2', columns: ['x'] }])).toThrow(/2 columns but 1 names/);
  });
});

describe('reading comments', () => {
  it('reads notes back, with authors', async () => {
    const { SheetReader } = await import('../src/core/index');
    const w = new SheetWriter();
    w.addSheet('S', [[{ value: 1, comment: { text: ' two\nlines ', author: 'Ana' } }, { value: 2, comment: 'plain' }]]);
    const rows = await new SheetReader().parse(createBlobReader(await new Response(w.write()).blob()));
    expect(await rows.getComments()).toEqual([{ ref: 'A1', text: ' two\nlines ', author: 'Ana' }, { ref: 'B1', text: 'plain' }]);
  });

  it('joins rich text runs and skips phonetic text', async () => {
    const { parseComments } = await import('../src/core/worksheet-parser');
    const xml = '<comments><authors><author>Bo &amp; Co</author></authors><commentList>' +
      '<comment ref="C3" authorId="0"><text><r><rPr><b/></rPr><t>Bo &amp; Co:</t></r><r><t xml:space="preserve"> fix_x000D_</t></r><rPh sb="0" eb="1"><t>x</t></rPh></text></comment>' +
      '</commentList></comments>';
    expect(parseComments(xml)).toEqual([{ ref: 'C3', text: 'Bo & Co: fix\r', author: 'Bo & Co' }]);
  });
});
