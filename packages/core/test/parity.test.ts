import { describe, it, expect } from 'vitest';
import { SheetWriter } from '../src/core/writer';
import { SheetReader, ParseOptions } from '../src/core/index';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { ZipStreamWriter } from '../src/core/zip-stream-writer';
import { createBlobReader } from '../src/core/random-access';
import { shiftFormula, encodeXString, decodeXString } from '../src/core/utils';
import { colorResolver, parseThemeColors } from '../src/core/style-reader';
import { Row } from '../src/core/types';

const bytesOf = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());

async function entryText(bytes: Uint8Array, name: string): Promise<string> {
  const zip = new ZipRandomAccessParser(createBlobReader(new Blob([bytes])));
  await zip.parseCentralDirectory();
  return zip.has(name) ? new Response(await zip.extractStream(name)).text() : '';
}

async function read(bytes: Uint8Array, options?: ParseOptions) {
  const rows = [];
  for await (const r of await new SheetReader().parse(createBlobReader(new Blob([bytes])), options)) rows.push(r);
  return rows;
}

// Minimal workbook around hand-written sheet XML and optional parts
async function workbook(sheetData: string, parts: { styles?: string; theme?: string; sst?: string; sheetRels?: string; tail?: string } = {}): Promise<Uint8Array> {
  const zip = new ZipStreamWriter();
  const done = bytesOf(zip.stream);
  const add = (name: string, text: string) =>
    zip.addFile(name, new ReadableStream({ start(c) { c.enqueue(new TextEncoder().encode(text)); c.close(); } }));
  const ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const rel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const rels = (body: string) => `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
  await add('_rels/.rels', rels(`<Relationship Id="rId1" Type="${rel}/officeDocument" Target="xl/workbook.xml"/>`));
  await add('xl/workbook.xml', `<workbook ${ns}><sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>`);
  await add('xl/_rels/workbook.xml.rels', rels(
    `<Relationship Id="rId1" Type="${rel}/worksheet" Target="worksheets/sheet1.xml"/>` +
    `<Relationship Id="rId2" Type="${rel}/styles" Target="styles.xml"/>` +
    `<Relationship Id="rId3" Type="${rel}/theme" Target="theme/theme1.xml"/>` +
    `<Relationship Id="rId4" Type="${rel}/sharedStrings" Target="sharedStrings.xml"/>`));
  await add('xl/worksheets/sheet1.xml', `<worksheet ${ns}><sheetData>${sheetData}</sheetData>${parts.tail ?? ''}</worksheet>`);
  if (parts.styles) await add('xl/styles.xml', parts.styles);
  if (parts.theme) await add('xl/theme/theme1.xml', parts.theme);
  if (parts.sst) await add('xl/sharedStrings.xml', `<sst ${ns}>${parts.sst}</sst>`);
  if (parts.sheetRels) await add('xl/worksheets/_rels/sheet1.xml.rels', rels(parts.sheetRels));
  await zip.close();
  return done;
}

// Office 2013+ default theme colours
const THEME = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:clrScheme name="Office">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
  <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
  <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
  <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
  <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
  <a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>`;

describe('shiftFormula', () => {
  it.each([
    ['A1+B2', 1, 0, 'A2+B3'],
    ['$A1+A$1+$A$1', 2, 3, '$A3+D$1+$A$1'],
    ['SUM(A1:B2)*2', 0, 1, 'SUM(B1:C2)*2'],
    ['SUM(A:A)+SUM(1:1)', 1, 1, 'SUM(B:B)+SUM(2:2)'],
    ['LOG10(A1)+ATAN2(B1,C1)', 1, 0, 'LOG10(A2)+ATAN2(B2,C2)'],
    ['"A1"&A1', 1, 0, '"A1"&A2'],
    ["'Sheet A1'!A1+Sheet2!B1", 1, 0, "'Sheet A1'!A2+Sheet2!B2"],
    ['A1-1', -1, 0, '#REF!-1'],
    ['XFD1', 0, 1, '#REF!'],
    ['TRUE+1E5+A1', 1, 0, 'TRUE+1E5+A2'],
  ])('%s moved by (%i, %i)', (f, r, c, out) => {
    expect(shiftFormula(f, r, c)).toBe(out);
  });
});

describe('ST_Xstring encoding', () => {
  it('round-trips escapes, control characters and CR', () => {
    for (const s of ['_x0041_', 'a\rb', 'bell\x07', 'plain', '__x_']) {
      expect(decodeXString(encodeXString(s))).toBe(s);
    }
    expect(encodeXString('_x0041_')).toBe('_x005F_x0041_');
  });
});

describe('writer parity', () => {
  it('writes Date values as dated serials and reads them back', async () => {
    const when = new Date(Date.UTC(2024, 1, 29, 13, 45, 30));
    const day = new Date(Date.UTC(1900, 0, 1));
    const bytes = await bytesOf(new SheetWriter().write([[when, day, { value: when, style: { numFmt: 'dd/mm/yyyy' } }]]));
    const sheet = await entryText(bytes, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<v>45351.57'); // 2024-02-29 = serial 45351
    expect(sheet).toMatch(/<c r="B1" s="\d+"><v>1<\/v>/); // 1900-01-01 = serial 1
    const [row] = await read(bytes);
    expect(row.cells).toEqual([when.toISOString(), day.toISOString(), when.toISOString()]);
    const styles = await entryText(bytes, 'xl/styles.xml');
    expect(styles).toContain('formatCode="yyyy-mm-dd hh:mm:ss"');
    expect(styles).toContain('formatCode="yyyy-mm-dd"');
    expect(styles).toContain('formatCode="dd/mm/yyyy"');
  });

  it('rejects invalid dates', async () => {
    await expect(bytesOf(new SheetWriter().write([[new Date(NaN)]]))).rejects.toThrow(/Invalid Date in cell A1/);
  });

  it('shared strings: stored once, read back, escapes and spaces kept', async () => {
    const rows: Row[] = [['dup', ' lead', '_x0041_'], ['dup', 'a\rb', 'dup']];
    const bytes = await bytesOf(new SheetWriter({ sharedStrings: true }).write(rows));
    const sst = await entryText(bytes, 'xl/sharedStrings.xml');
    expect(sst).toContain('count="6" uniqueCount="4"');
    expect(sst).toContain('<t xml:space="preserve"> lead</t>');
    expect((await read(bytes)).map(r => r.cells)).toEqual(rows);
    // Inline strings get the same treatment
    expect((await read(await bytesOf(new SheetWriter().write(rows)))).map(r => r.cells)).toEqual(rows);
  });

  it('hyperlinks: external via sheet rels, internal via location', async () => {
    const bytes = await bytesOf(new SheetWriter().write([[
      { value: 'site', hyperlink: 'https://example.com/?a=1&b=2' },
      { value: 'jump', hyperlink: "#'Other Sheet'!A1" },
    ]]));
    const sheet = await entryText(bytes, 'xl/worksheets/sheet1.xml');
    expect(sheet).toContain('<hyperlink ref="A1" r:id="rId1"/>');
    expect(sheet).toContain(`<hyperlink ref="B1" location="&apos;Other Sheet&apos;!A1"/>`);
    const rels = await entryText(bytes, 'xl/worksheets/_rels/sheet1.xml.rels');
    expect(rels).toContain('Target="https://example.com/?a=1&amp;b=2" TargetMode="External"');
    expect(await entryText(await bytesOf(new SheetWriter().write([['x']])), 'xl/worksheets/_rels/sheet1.xml.rels')).toBe('');
  });

  it('autoFilter plus the hidden _FilterDatabase name, elements in schema order', async () => {
    const bytes = await bytesOf(new SheetWriter().write([['a', 'b'], [1, 2]], {
      name: "Bob's", autoFilter: 'A1:B2', mergeCells: ['A3:B3'],
      conditionalFormats: [{ range: 'A2:A2', rule: { type: 'dataBar', color: 'FF638EC6' } }],
      dataValidations: [{ sqref: 'B2', type: 'whole' }],
    }));
    const sheet = await entryText(bytes, 'xl/worksheets/sheet1.xml');
    const order = ['</sheetData>', '<autoFilter ref="A1:B2"/>', '<mergeCells', '<conditionalFormatting', '<dataValidations'].map(t => sheet.indexOf(t));
    expect(order.every((at, i) => at > 0 && (i === 0 || at > order[i - 1]))).toBe(true);
    const wb = await entryText(bytes, 'xl/workbook.xml');
    expect(wb).toContain(`<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">&apos;Bob&apos;&apos;s&apos;!$A$1:$B$2</definedName>`);
  });

  it("vertical 'middle' is written as the valid 'center'", async () => {
    const bytes = await bytesOf(new SheetWriter().write([[{ value: 1, style: { alignment: { vertical: 'middle' } } }]]));
    expect(await entryText(bytes, 'xl/styles.xml')).toContain('vertical="center"');
  });
});

describe('reader parity', () => {
  it('reports formulas, expanding shared formulas per cell', async () => {
    const bytes = await workbook(
      '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B3" si="0">A1*2+$A$1</f><v>3</v></c></row>' +
      '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><f t="shared" si="0"/><v>5</v></c><c r="C2"><f>"x"&amp;A2</f><v>x2</v></c></row>' +
      '<row r="3"><c r="B3"><f t="shared" si="0"/><v>1</v></c></row>' +
      '<row r="4"><c r="A4"><v>9</v></c></row>'
    );
    const rows = await read(bytes, { formulas: true });
    expect(rows.map(r => r.formulas)).toEqual([
      [undefined, 'A1*2+$A$1'],
      [undefined, 'A2*2+$A$1', '"x"&A2'],
      [undefined, 'A3*2+$A$1'],
      undefined,
    ]);
    expect(rows[1].cells).toEqual([2, 5, 'x2']); // cached values unchanged
    expect((await read(bytes))[0].formulas).toBeUndefined(); // opt-in
  });

  it('round-trips writer formulas', async () => {
    const bytes = await bytesOf(new SheetWriter().write([[1, 2, { value: null, formula: '=SUM(A1:B1)' }]]));
    const [row] = await read(bytes, { formulas: true });
    expect(row.cells[2]).toBe(3);
    expect(row.formulas?.[2]).toBe('SUM(A1:B1)');
  });

  it('round-trips writer styles', async () => {
    const style = {
      font: { bold: true, italic: true, size: 14, color: 'FFFF0000', name: 'Arial' },
      fill: { type: 'solid' as const, fgColor: 'FF00FF00' },
      border: { top: { style: 'thin' as const, color: 'FF0000FF' }, bottom: { style: 'double' as const } },
      alignment: { horizontal: 'center' as const, vertical: 'top' as const, wrapText: true },
      numFmt: '#,##0.00 "USD"',
    };
    const gradient = { fill: { type: 'gradient' as const, degree: 90, stops: [{ position: 0, color: 'FF000000' }, { position: 1, color: 'FFFFFFFF' }] } };
    const bytes = await bytesOf(new SheetWriter().write([[{ value: 1234.5, style }, { value: 'g', style: gradient }, 'plain']]));
    const [row] = await read(bytes, { styles: true });
    expect(row.cells[0]).toBe(1234.5); // not mistaken for a date
    expect(row.styles).toEqual([style, gradient]);
    expect((await read(bytes))[0].styles).toBeUndefined(); // opt-in
  });

  it('builtin number formats, dxf look-alikes and default styles', async () => {
    const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b val="0"/><u/><sz val="9"/></font></fonts>
      <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
      <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
      <cellStyleXfs count="1"><xf numFmtId="0" fontId="1"/></cellStyleXfs>
      <cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="10" fontId="1"/><xf numFmtId="14"/></cellXfs>
      <dxfs count="1"><dxf><font><b/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/></patternFill></fill></dxf></dxfs>
    </styleSheet>`;
    const bytes = await workbook('<row r="1"><c r="A1" s="0"><v>1</v></c><c r="B1" s="1"><v>0.5</v></c><c r="C1" s="2"><v>1</v></c></row>', { styles });
    const [row] = await read(bytes, { styles: true });
    expect(row.styles).toEqual([undefined, { font: { bold: false, underline: true, size: 9 }, numFmt: '0.00%' }, { numFmt: 'mm-dd-yy' }]);
    expect(row.cells[2]).toBe('1900-01-01T00:00:00.000Z');
  });
});

describe('colours', () => {
  const color = colorResolver(parseThemeColors(THEME));

  it('resolves theme slots (lt/dk swapped), tints, indexed and rgb', () => {
    expect(color({ theme: '0' })).toBe('FFFFFFFF'); // lt1
    expect(color({ theme: '1' })).toBe('FF000000'); // dk1
    expect(color({ theme: '4' })).toBe('FF4472C4'); // accent1
    // Excel's own swatches ("Blue, Accent 1, Lighter 60%" etc.), within 2 per channel
    const near = (got: string | undefined, want: string) =>
      [2, 4, 6].every(i => Math.abs(parseInt(got!.slice(i, i + 2), 16) - parseInt(want.slice(i - 2, i), 16)) <= 2);
    expect(near(color({ theme: '4', tint: '0.59999389629810485' }), 'B4C6E7')).toBe(true);
    expect(near(color({ theme: '4', tint: '0.79998168889431442' }), 'D9E1F2')).toBe(true);
    expect(near(color({ theme: '4', tint: '-0.249977111117893' }), '2F5597')).toBe(true);
    expect(near(color({ theme: '9', tint: '0.39997558519241921' }), 'A9D08E')).toBe(true);
    expect(near(color({ theme: '2', tint: '-0.499984740745262' }), '767171')).toBe(true);
    expect(color({ theme: '0', tint: '-0.14999847407452621' })).toBe('FFD9D9D9');
    expect(color({ indexed: '10' })).toBe('FFFF0000');
    expect(color({ indexed: '64' })).toBe('FF000000');
    expect(color({ rgb: 'ff00ff00' })).toBe('FF00FF00');
    expect(color({ auto: '1' })).toBeUndefined();
    expect(color({ theme: '99' })).toBeUndefined();
  });

  it('styles use the theme and the workbook palette', async () => {
    const styles = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="2"><font><sz val="11"/></font><font><color theme="4" tint="-0.249977111117893"/></font></fonts>
      <fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor indexed="2"/><bgColor indexed="64"/></patternFill></fill></fills>
      <borders count="2"><border/><border><left style="thin"><color auto="1"/></left><top style="thin"><color theme="1"/></top></border></borders>
      <cellXfs count="2"><xf/><xf fontId="1" fillId="2" borderId="1"/></cellXfs>
      <colors><indexedColors><rgbColor rgb="FF000000"/><rgbColor rgb="FFFFFFFF"/><rgbColor rgb="FF123456"/></indexedColors></colors>
    </styleSheet>`;
    const bytes = await workbook('<row r="1"><c r="A1" s="1"><v>1</v></c></row>', { styles, theme: THEME });
    const [row] = await read(bytes, { styles: true });
    expect(row.styles?.[0]).toEqual({
      font: { color: 'FF2F5597' },
      fill: { type: 'solid', fgColor: 'FF123456' }, // custom palette overrides index 2
      border: { left: { style: 'thin' }, top: { style: 'thin', color: 'FF000000' } },
    });
  });
});

describe('hyperlinks', () => {
  it('round-trip through metadata in the writer format', async () => {
    const bytes = await bytesOf(new SheetWriter().write([[
      { value: 'site', hyperlink: 'https://example.com/?a=1&b=2' },
      { value: 'jump', hyperlink: '#Other!B2' },
    ]]));
    const rows = await new SheetReader().parse(createBlobReader(new Blob([bytes])));
    for await (const _ of rows);
    expect((await rows.getMetadata()).hyperlinks).toEqual([
      { ref: 'A1', hyperlink: 'https://example.com/?a=1&b=2' },
      { ref: 'B1', hyperlink: '#Other!B2' },
    ]);
  });

  it('keeps URL fragments and tooltips, skips unknown relationships', async () => {
    const bytes = await workbook('<row r="1"><c r="A1"><v>1</v></c></row>', {
      tail: '<hyperlinks><hyperlink ref="A1:B2" r:id="rId7" location="part" tooltip="Go"/><hyperlink ref="C1" r:id="rIdMissing"/></hyperlinks>',
      sheetRels: '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/doc" TargetMode="External"/>',
    });
    const rows = await new SheetReader().parse(createBlobReader(new Blob([bytes])));
    for await (const _ of rows);
    expect((await rows.getMetadata()).hyperlinks).toEqual([{ ref: 'A1:B2', hyperlink: 'https://example.com/doc#part', tooltip: 'Go' }]);
  });
});

describe('rich text', () => {
  const runs = [
    { text: 'Bold', font: { bold: true, color: 'FFFF0000' } },
    { text: ' and ' },
    { text: 'italic', font: { italic: true, size: 14, name: 'Arial' } },
  ];

  it('round-trips inline runs; cells hold the plain text', async () => {
    const bytes = await bytesOf(new SheetWriter({ sharedStrings: true }).write([[{ value: null, richText: runs }, 'plain']]));
    expect(await entryText(bytes, 'xl/worksheets/sheet1.xml')).toContain('<rPr><b/><color rgb="FFFF0000"/></rPr><t>Bold</t>');
    const [row] = await read(bytes, { richText: true });
    expect(row.cells).toEqual(['Bold and italic', 'plain']);
    expect(row.richText).toEqual([runs]);
    expect((await read(bytes))[0].richText).toBeUndefined(); // opt-in
  });

  it('reads runs from shared strings, with theme colours and without phonetic text', async () => {
    const sst = '<si><t>plain</t></si>' +
      '<si><r><rPr><b/><color theme="4"/><rFont val="Calibri"/></rPr><t>Hi</t></r><r><t xml:space="preserve"> _x0041_</t></r><rPh sb="0" eb="1"><t>\u30CF\u30A4</t></rPh></si>';
    const bytes = await workbook('<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>', { sst, theme: THEME });
    const [row] = await read(bytes, { richText: true });
    expect(row.cells).toEqual(['plain', 'Hi A']);
    expect(row.richText).toEqual([undefined, [{ text: 'Hi', font: { bold: true, color: 'FF4472C4', name: 'Calibri' } }, { text: ' A' }]]);
  });
});
