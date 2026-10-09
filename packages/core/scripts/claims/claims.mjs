// README claims that need only Web APIs. Runs unchanged on Node, Bun, Deno, browsers and Workers:
// run(lib, fixtures) -> [{ name, ok, error }]. fixtures: the FIXTURES files from test/fixtures, as Uint8Array.
export const FIXTURES = { xls: 'excel-made.xls', ods: 'excel-made.ods', xlsx: 'excel-made.xlsx', png: 'red-40x20.png', jpg: 'blue-30x60.jpg', gif: '12x7.gif', encrypted: 'encrypted.xlsx' };

export async function run(lib, fx) {
  const { SheetReader, SheetWriter, SheetEditor, OdsWriter, createBlobReader, parseCsv, sheetToJson, streamToCsv,
    ZipRandomAccessParser, ZipStreamWriter } = lib;
  const results = [];
  const t = async (name, fn) => {
    try { await fn(); results.push({ name, ok: true }); }
    catch (e) { results.push({ name, ok: false, error: String(e?.message ?? e).slice(0, 400) }); }
  };
  const eq = (a, b, what = '') => {
    const x = JSON.stringify(a), y = JSON.stringify(b);
    if (x !== y) throw new Error(`${what} expected ${y}, got ${x}`);
  };
  const ok = (c, what) => { if (!c) throw new Error(what); };
  const throwsLike = async (fn, re, what) => {
    try { await fn(); } catch (e) { if (re && !re.test(String(e?.message ?? e))) throw new Error(`${what}: wrong error ${e?.message}`); return; }
    throw new Error(`${what}: did not throw`);
  };
  // read by hand: Chrome's Response.arrayBuffer() replaces a stream's error with "Failed to fetch"
  const bytes = async (s) => {
    const r = s.getReader(), parts = []; let n = 0;
    for (;;) { const { done, value } = await r.read(); if (done) break; parts.push(value); n += value.length; }
    const out = new Uint8Array(n); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out;
  };
  const reader = (b) => createBlobReader(new Blob([b]));
  const rows = async (b, o) => { const out = []; for await (const r of await new SheetReader().parse(reader(b), o)) out.push(r); return out; };
  const write = (fn, opts) => { const w = new SheetWriter(opts); fn(w); return bytes(w.write()); };
  const parts = async (b) => {
    const z = new ZipRandomAccessParser(reader(b)); await z.parseCentralDirectory();
    const m = new Map(); for (const n of z.getFiles()) m.set(n, await bytes(await z.extractStream(n))); return m;
  };
  const text = async (b, name) => new TextDecoder().decode((await parts(b)).get(name));
  const rezip = async (b, change) => {
    const m = await parts(b); await change(m);
    const z = new ZipStreamWriter(); const done = bytes(z.stream);
    for (const [n, d] of m) await z.addFile(n, new Blob([typeof d === 'string' ? new TextEncoder().encode(d) : d]).stream());
    await z.close(); return done;
  };
  const collect = async (gen) => { const out = []; for await (const r of gen) out.push(r); return out; };

  // ---------- Features ----------
  await t('Streaming: 100,000 rows written from an async generator and read back', async () => {
    async function* many() { for (let i = 0; i < 100000; i++) yield [i, `row ${i}`, i / 2]; }
    const b = await bytes(new SheetWriter().addSheet('Big', many()).write());
    let n = 0, last; for await (const r of await new SheetReader().parse(reader(b))) { n++; last = r.cells; }
    eq([n, last], [100000, [99999, 'row 99999', 49999.5]]);
  });

  // ---------- Quick start: reading ----------
  const sample = await write((w) => {
    w.addSheet('Other', [['x']]);
    w.addSheet('Sheet1', [['ID', 'Name', 'Amount'], [1, 'Ana', new Date(Date.UTC(2026, 9, 8, 14, 5))], [2, 'Raj', 3], [3, 'Li', 4]], {
      mergeCells: ['A5:B5'], freezePanes: { row: 1 }, rows: { 3: { hidden: true } }, columns: [{}, { hidden: true }],
    });
  });
  await t('Quick start: parse by sheetName, first row is the header', async () => {
    const r = await rows(sample, { sheetName: 'Sheet1' });
    eq([r[0].rowNumber, r[0].cells], [1, ['ID', 'Name', 'Amount']]);
  });
  await t('Quick start: omitting sheetName reads the first tab', async () => eq((await rows(sample))[0].cells, ['x']));
  await t('Quick start: getMetadata gives merged cells, hidden rows/cols, freeze panes', async () => {
    const res = await new SheetReader().parse(reader(sample), { sheetName: 'Sheet1' });
    for await (const _ of res);
    const m = await res.getMetadata();
    eq([m.mergedCells, m.hiddenRows, m.hiddenCols, m.freezePanes?.row], [['A5:B5'], [3], [2], 1]);
  });
  await t('Dates come back as ISO-8601 strings', async () => eq((await rows(sample, { sheetName: 'Sheet1' }))[1].cells[2], '2026-10-08T14:05:00.000Z'));

  await t('{ formulas: true } expands shared formulas per cell', async () => {
    const b = await rezip(await write((w) => w.addSheet('S', [[1, 0], [2, 0], [3, 0]])), async (m) => {
      const s = new TextDecoder().decode(m.get('xl/worksheets/sheet1.xml'));
      let i = 0;
      m.set('xl/worksheets/sheet1.xml', s.replace(/<c r="B(\d)"([^>]*)>.*?<\/c>/g, (_, r, a) =>
        `<c r="B${r}"${a.replace(/ t="\w+"/, '')}>${i++ === 0 ? '<f t="shared" ref="B1:B3" si="0">A1*2</f>' : '<f t="shared" si="0"/>'}<v>0</v></c>`));
    });
    eq((await rows(b, { formulas: true })).map((r) => r.formulas[1]), ['A1*2', 'A2*2', 'A3*2']);
  });
  await t('{ styles: true } gives CellStyle in the writer shape; theme colours resolved to ARGB', async () => {
    const style = { font: { bold: true, color: 'FFC00000' }, fill: { type: 'solid', fgColor: 'FF00B050' }, numFmt: '0.00' };
    const b = await write((w) => w.addSheet('S', [[{ value: 1, style }, { value: 2, style: { font: { color: 'FF000000' } } }]]));
    const r = (await rows(b, { styles: true }))[0].styles;
    ok(r[0].font.bold === true && r[0].font.color === 'FFC00000' && r[0].fill.fgColor === 'FF00B050' && r[0].numFmt === '0.00', JSON.stringify(r[0]));
    const themed = await rezip(b, async (m) => {
      const s = new TextDecoder().decode(m.get('xl/styles.xml'));
      ok(s.includes('rgb="FF000000"'), 'no black font to swap');
      m.set('xl/styles.xml', s.replace('rgb="FF000000"', 'theme="4"'));
      const excel = await parts(fx.xlsx);
      m.set('xl/theme/theme1.xml', excel.get('xl/theme/theme1.xml'));
      m.set('xl/_rels/workbook.xml.rels', new TextDecoder().decode(m.get('xl/_rels/workbook.xml.rels')).replace('</Relationships>',
        '<Relationship Id="rIdTheme" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>'));
      m.set('[Content_Types].xml', new TextDecoder().decode(m.get('[Content_Types].xml')).replace('</Types>',
        '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>'));
    });
    const c = (await rows(themed, { styles: true }))[0].styles[1].font.color;
    ok(/^[0-9A-F]{8}$/i.test(c), `theme colour came back as ${c}`);
  });
  await t('{ richText: true } gives runs; cells keep plain text', async () => {
    const b = await write((w) => w.addSheet('S', [[{ value: null, richText: [{ text: 'Net ', font: { bold: true } }, { text: 'revenue', font: { color: 'FFC00000' } }] }]]));
    const r = (await rows(b, { richText: true }))[0];
    eq(r.cells[0], 'Net revenue');
    eq(r.richText[0].map((x) => x.text), ['Net ', 'revenue']);
    ok(r.richText[0][0].font?.bold === true, 'bold run lost');
  });
  await t('{ errors: true } tells an error value from text reading #N/A', async () => {
    const b = await write((w) => w.addSheet('S', [[{ value: null, formula: '1/0' }, '#N/A']]));
    const r = (await rows(b, { errors: true }))[0];
    eq([r.cells, r.errors[0], r.errors[1]], [['#DIV/0!', '#N/A'], true, undefined]);
  });
  await t('{ formatted: true } examples and format kinds (en-US)', async () => {
    const cell = (value, numFmt) => ({ value, style: { numFmt } });
    const b = await write((w) => w.addSheet('S', [[
      cell(1234.5, '#,##0.00'), cell(0.256, '0.0%'), cell(-42, '#,##0;(#,##0)'),
      cell(new Date(Date.UTC(2026, 9, 8, 14, 5)), 'dd-mmm-yyyy h:mm AM/PM'),
      cell(1.5, '[h]:mm'), cell(0.5, '# ?/?'), cell(12345, '0.00E+00'), cell(1234.5, '"$"#,##0.00'),
      cell('abc', '"pre "@'), cell(150, '[>100]"big";"small"'),
    ]]));
    eq((await rows(b, { formatted: true }))[0].formatted,
      ['1,234.50', '25.6%', '(42)', '08-Oct-2026 2:05 PM', '36:00', ' 1/2', '1.23E+04', '$1,234.50', 'pre abc', 'big']);
  });
  await t('readWorkbook lists sheets with visibility, defined names, properties', async () => {
    const b = await write((w) => { w.addSheet('A', [[1]], { state: 'hidden' }); w.addSheet('B', [[1]]); },
      { properties: { title: 'T' }, definedNames: [{ name: 'Rate', ref: '0.18' }] });
    const i = await new SheetReader().readWorkbook(reader(b));
    eq([i.sheets, i.definedNames[0].name, i.properties.title], [[{ name: 'A', state: 'hidden' }, { name: 'B', state: 'visible' }], 'Rate', 'T']);
  });
  await t('Hyperlinks, images (with bytes) and notes read back in the writer format', async () => {
    const b = await write((w) => w.addSheet('S', [
      [{ value: 'Docs', hyperlink: 'https://example.com' }, { value: 'Q3', comment: { text: 'Restated', author: 'Ana' } }],
    ], { images: [{ data: fx.png, at: 'D1', height: 40 }] }));
    const res = await new SheetReader().parse(reader(b));
    for await (const _ of res);
    const m = await res.getMetadata();
    eq(m.hyperlinks.map(({ ref, hyperlink }) => ({ ref, hyperlink })), [{ ref: 'A1', hyperlink: 'https://example.com' }]);
    const imgs = await res.getImages();
    ok(imgs.length === 1 && imgs[0].data.length === fx.png.length && imgs[0].data.every((v, i) => v === fx.png[i]), 'image bytes differ');
    const back = await bytes(new SheetWriter().addSheet('S', [[1]], { images: imgs }).write());
    ok((await parts(back)).has('xl/media/image1.png'), 'images did not write back');
    eq(await res.getComments(), [{ ref: 'B1', text: 'Restated', author: 'Ana' }]);
  });
  await t('maxUncompressedBytes caps parts held in memory; Infinity disables', async () => {
    const b = await write((w) => w.addSheet('S', [['x'.repeat(5000)]]), { sharedStrings: true });
    await throwsLike(() => rows(b, { maxUncompressedBytes: 1000 }), null, 'small limit');
    eq((await rows(b, { maxUncompressedBytes: Infinity }))[0].cells[0].length, 5000);
  });
  await t('Crafted ZIP: entry inflating past its stated size is rejected', async () => {
    const b = (await write((w) => w.addSheet('S', [['y'.repeat(20000)]]))).slice();
    const dv = new DataView(b.buffer);
    for (let i = 0; i < b.length - 46; i++) {
      if (dv.getUint32(i, true) === 0x02014b50) {
        const nameLen = dv.getUint16(i + 28, true);
        const name = new TextDecoder().decode(b.subarray(i + 46, i + 46 + nameLen));
        if (name === 'xl/worksheets/sheet1.xml') {
          dv.setUint32(i + 24, 10, true);
          const local = dv.getUint32(i + 42, true);
          dv.setUint32(local + 22, 10, true);
        }
      }
    }
    await throwsLike(() => rows(b), null, 'inflating past size');
  });
  await t('Crafted ZIP: directory pointing outside the file is rejected', async () => {
    const b = (await write((w) => w.addSheet('S', [[1]]))).slice();
    const dv = new DataView(b.buffer);
    const eocd = b.length - 22;
    eq(dv.getUint32(eocd, true), 0x06054b50, 'EOCD');
    dv.setUint32(eocd + 16, b.length + 1000, true);
    await throwsLike(() => rows(b), null, 'outside file');
  });
  await t('Crafted ZIP: overlapping entries are rejected', async () => {
    const b = (await write((w) => w.addSheet('S', [[1]]))).slice();
    const dv = new DataView(b.buffer);
    const offsets = [];
    for (let i = 0; i < b.length - 46; i++) if (dv.getUint32(i, true) === 0x02014b50) offsets.push(i);
    const first = dv.getUint32(offsets[0] + 42, true);
    for (const i of offsets.slice(1)) dv.setUint32(i + 42, first, true);
    await throwsLike(() => rows(b), null, 'overlapping');
  });
  await t('One cell in column XFD pads its row; the limit bounds that padding', async () => {
    const b = await rezip(await write((w) => w.addSheet('S', Array.from({ length: 200 }, () => [1]))), async (m) => {
      const s = new TextDecoder().decode(m.get('xl/worksheets/sheet1.xml'));
      m.set('xl/worksheets/sheet1.xml', s.replace(/<c r="A(\d+)"/g, '<c r="XFD$1"'));
    });
    eq((await rows(b))[0].cells.length, 16384, 'row length');
    await throwsLike(() => rows(b, { maxUncompressedBytes: 100000 }), null, 'padding limit');
  });
  await t('Values returned as written: javascript: links and text starting with =', async () => {
    const b = await write((w) => w.addSheet('S', [[{ value: 'x', hyperlink: 'javascript:alert(1)' }, '=1+1']]));
    const res = await new SheetReader().parse(reader(b));
    const r = []; for await (const x of res) r.push(x);
    eq([(await res.getMetadata()).hyperlinks[0].hyperlink, r[0].cells[1]], ['javascript:alert(1)', '=1+1']);
  });
  await t('sheetToJson keys by header, repeated header gets _2; streamToCsv', async () => {
    const b = await write((w) => w.addSheet('S', [['Name', 'Name', 'Age'], ['Ana', 'B', 30], ['x,y', 'C', null]]));
    eq(await sheetToJson(await new SheetReader().parse(reader(b))), [{ Name: 'Ana', Name_2: 'B', Age: 30 }, { Name: 'x,y', Name_2: 'C', Age: null }]);
    const csv = (await streamToCsv(await new SheetReader().parse(reader(b)))).trim().split(/\r?\n/);
    eq(csv, ['Name,Name,Age', 'Ana,B,30', '"x,y",C,']);
  });

  // ---------- Writing ----------
  await t('write() returns a ReadableStream of a ZIP (.xlsx) file', async () => {
    const s = new SheetWriter().addSheet('Report', [['Header 1', 'Header 2', 'Header 3'], [1, 2, 3], ['Data', 'More Data', 'Even More Data']]).write();
    ok(s instanceof ReadableStream, 'not a ReadableStream');
    const b = await bytes(s);
    eq([b[0], b[1]], [0x50, 0x4b], 'zip signature');
    eq((await rows(b))[2].cells, ['Data', 'More Data', 'Even More Data']);
  });
  await t('AsyncIterable rows are pulled only as fast as the output is read', async () => {
    let pulled = 0;
    async function* gen() { for (let i = 0; i < 1e6; i++) { pulled++; yield [i, i, i, i, i]; } }
    const s = new SheetWriter().addSheet('S', gen()).write();
    const r = s.getReader(); await r.read();
    await new Promise((res) => setTimeout(res, 300));
    const atRest = pulled;
    await r.cancel();
    ok(atRest < 100000, `pulled ${atRest} rows with nobody reading`);
  });
  await t('Properties, defined names, hidden/veryHidden sheets, view settings', async () => {
    const b = await write((w) => {
      w.addSheet('Lookup', [[1]], { state: 'hidden' });
      w.addSheet('Secret', [[1]], { state: 'veryHidden' });
      w.addSheet('Data', [[1]], { view: { zoom: 90, showGridLines: false, rightToLeft: false } });
    }, {
      properties: { title: 'Q3 sales', creator: 'Finance', company: 'ACME' },
      definedNames: [{ name: 'TaxRate', ref: '0.18' }, { name: 'Sales', ref: 'Data!$B$2:$B$100' }, { name: 'Total', ref: 'Data!$B$101', sheet: 'Data' }],
    });
    const i = await new SheetReader().readWorkbook(reader(b));
    eq(i.sheets.map((s) => s.state), ['hidden', 'veryHidden', 'visible']);
    eq([i.properties.title, i.properties.creator, i.properties.company], ['Q3 sales', 'Finance', 'ACME']);
    eq(i.definedNames.map((d) => [d.name, d.ref, d.sheet]), [['TaxRate', '0.18', undefined], ['Sales', 'Data!$B$2:$B$100', undefined], ['Total', 'Data!$B$101', 'Data']]);
    const s3 = await text(b, 'xl/worksheets/sheet3.xml');
    ok(/zoomScale="90"/.test(s3) && /showGridLines="0"/.test(s3), 'view not written');
    const wb = await text(b, 'xl/workbook.xml');
    ok(/activeTab="2"/.test(wb), 'Excel would not open on the first visible sheet');
  });
  await t('Invalid or duplicate names, and no visible sheet, are rejected before writing', async () => {
    const reject = async (fn, what) => {
      let first = null;
      try { const w = new SheetWriter(); fn(w); const r = w.write().getReader(); first = await r.read(); }
      catch { return; }
      throw new Error(`${what} not rejected (first chunk ${first?.value?.length} bytes)`);
    };
    await reject((w) => w.addSheet('a/b', [[1]]), 'invalid sheet name');
    await reject((w) => { w.addSheet('A', [[1]]); w.addSheet('a', [[1]]); }, 'duplicate sheet name');
    await reject((w) => w.addSheet('A', [[1]], { state: 'hidden' }), 'no visible sheet');
    await throwsLike(async () => { const x = new SheetWriter({ definedNames: [{ name: 'A1', ref: '1' }] }); x.addSheet('S', [[1]]); await x.write().getReader().read(); }, null, 'invalid defined name');
  });

  // ---------- CSV ----------
  await t('parseCsv: quotes, line breaks, "", numbers, booleans, empty vs quoted empty, dates stay text', async () => {
    const csv = 'a,"b,c","line\nbreak","say ""hi""",12.5,TRUE,,"",2026-10-08\n';
    eq(await collect(parseCsv(csv)), [['a', 'b,c', 'line\nbreak', 'say "hi"', 12.5, true, null, '', '2026-10-08']]);
  });
  await t('parseCsv: ReadableStream input, { convert: false }, { delimiter: ";" }', async () => {
    const stream = new Blob(['x;1;FALSE\n']).stream();
    eq(await collect(parseCsv(stream, { delimiter: ';' })), [['x', 1, false]]);
    eq(await collect(parseCsv('1,TRUE\n', { convert: false })), [['1', 'TRUE']]);
  });
  await t('CSV to .xlsx in one line', async () => {
    const xlsx = new SheetWriter().addSheet('Data', parseCsv('a,1\nb,2\n')).write();
    eq((await rows(await bytes(xlsx))).map((r) => r.cells), [['a', 1], ['b', 2]]);
  });

  // ---------- .xls and .ods ----------
  // a merge's covered cells come back as trailing nulls from .xlsx only
  const strip = (r) => r.map((x) => { const c = [...x.cells]; while (c.length && c.at(-1) === null) c.pop(); return [x.rowNumber, c]; });
  await t('.xls, .ods and .xlsx of one Excel workbook read the same values', async () => {
    const [a, b, c] = await Promise.all([rows(fx.xlsx), rows(fx.xls), rows(fx.ods)]);
    eq(strip(b), strip(a), '.xls');
    eq(strip(c), strip(a), '.ods');
  });
  await t('.xls and .ods: readWorkbook gives the same sheets, names and title', async () => {
    const info = async (b) => { const i = await new SheetReader().readWorkbook(reader(b)); return [i.sheets, i.definedNames.map((d) => d.name).sort()]; };
    const [a, b, c] = await Promise.all([info(fx.xlsx), info(fx.xls), info(fx.ods)]);
    eq(b, a, '.xls'); eq(c, a, '.ods');
  });
  await t('formatted: true works on .xls and .ods; formulas: true on .ods', async () => {
    const [b, c] = await Promise.all([rows(fx.xls, { formatted: true }), rows(fx.ods, { formatted: true, formulas: true })]);
    ok(b.some((r) => r.formatted?.some((x) => x !== undefined)), '.xls formatted empty');
    ok(c.some((r) => r.formatted?.some((x) => x !== undefined)), '.ods formatted empty');
    ok(c.some((r) => r.formulas?.some((x) => x)), '.ods formulas empty');
  });
  await t('A password-protected file points to decryptWorkbook', async () => {
    await throwsLike(() => rows(fx.encrypted), /decryptWorkbook/, 'encrypted');
  });
  await t('OdsWriter: values, dates, formulas with results, merges, widths, freeze, hidden sheets, properties', async () => {
    const ods = await bytes(new OdsWriter({ properties: { title: 'Report' } })
      .addSheet('Data', [[1, new Date(Date.UTC(2026, 9, 8))], [2, { value: null, formula: 'SUM(A1:A2)' }], ['m', null]],
        { columnWidths: [20, 10], freezePanes: { row: 1 }, mergeCells: ['A3:B3'] })
      .addSheet('Hid', [[1]], { state: 'hidden' })
      .write());
    const res = await new SheetReader().parse(reader(ods));
    const r = []; for await (const x of res) r.push(x.cells);
    eq(r[0], [1, '2026-10-08T00:00:00.000Z']); eq(r[1], [2, 3]);
    const m = await res.getMetadata();
    eq([m.mergedCells, m.freezePanes?.row], [['A3:B3'], 1]);
    const i = await new SheetReader().readWorkbook(reader(ods));
    eq([i.sheets.map((s) => s.state), i.properties.title], [['visible', 'hidden'], 'Report']);
    const content = await text(ods, 'content.xml');
    ok(/of:=SUM\(\[\.A1:\.A2\]\)/.test(content), 'not OpenFormula');
    ok(/style:column-width/.test(content) || /style:column-width/.test(await text(ods, 'styles.xml')), 'no column widths');
  });

  // ---------- Editing ----------
  const base = await write((w) => {
    w.addSheet('Sheet1', Array.from({ length: 30 }, (_, i) => [
      i === 0 ? { value: 'Head', style: { font: { italic: true }, fill: { type: 'solid', fgColor: 'FF0000FF' } } } : `r${i + 1}`,
      i + 1, i === 1 ? { value: 5, style: { font: { color: 'FF00B050' } } } : null,
    ]));
    w.addSheet('Old', [['bye']]);
  });
  await t('Editing: README example runs; addresses count after the moves', async () => {
    const editor = new SheetEditor();
    editor.setCells('Sheet1', {
      B2: 42, C2: { formula: 'B2*2' }, D9: 'new cell', A3: null,
      A1: { style: { font: { bold: true }, fill: { type: 'solid', fgColor: 'FFFFFF00' } } },
      B3: { value: 7, style: { numFmt: '0.00' } },
    });
    editor.appendSheet('Sheet1', [['new', 'row']]);
    editor.insertRows('Sheet1', 5, 3);
    editor.deleteRows('Sheet1', 20, 2);
    editor.insertColumns('Sheet1', 'C');
    editor.addSheet('Summary', [['Total', { value: null, formula: 'SUM(Sheet1!B:B)' }]]);
    editor.deleteSheet('Old');
    const out = await bytes(editor.edit(reader(base)));
    const info = await new SheetReader().readWorkbook(reader(out));
    eq(info.sheets.map((s) => s.name), ['Sheet1', 'Summary']);
    const r = await rows(out, { sheetName: 'Sheet1', formulas: true, styles: true });
    const at = (n) => r.find((x) => x.rowNumber === n);
    eq(at(1).cells[0], 'Head');
    eq(at(2).cells[1], 42);
    ok(at(2).formulas?.some((f) => f && /B2\*2/.test(f)), `C2 formula: ${JSON.stringify(at(2).formulas)}`);
    eq(at(3).cells[0], null);
    eq(at(3).cells[1], 7);
    eq(at(3).styles[1]?.numFmt, '0.00');
    const st = at(1).styles[0];
    ok(st.font.bold === true && st.font.italic === true && st.fill.fgColor === 'FFFFFF00', `style merge: ${JSON.stringify(st)}`);
    ok(r.some((x) => x.cells.includes('new cell')), 'new cell missing');
    eq(r.at(-1).cells.filter((v) => v !== null), ['new', 'row']);
    eq(r.at(-1).rowNumber, 30 + 1 + 3 - 2, 'last row');
    ok([5, 6, 7].every((n) => !at(n) || at(n).cells.every((v) => v === null)), 'inserted rows 5-7 not empty');
  });
  await t('Edited cells keep their style; inserted rows do not copy neighbours\' formatting', async () => {
    const out = await bytes(new SheetEditor().setCells('Sheet1', { C2: 9 }).insertRows('Sheet1', 3, 1).edit(reader(base)));
    const r = await rows(out, { sheetName: 'Sheet1', styles: true });
    eq([r[1].cells[2], r[1].styles[2]?.font?.color], [9, 'FF00B050']);
    const row3 = r.find((x) => x.rowNumber === 3);
    ok(!row3 || row3.styles.every((s) => !s), 'inserted row has formats');
  });
  await t('Overwriting the first cell of a shared formula writes it out in the others', async () => {
    const src = await rezip(await write((w) => w.addSheet('S', [[1, 0], [2, 0], [3, 0]])), async (m) => {
      const s = new TextDecoder().decode(m.get('xl/worksheets/sheet1.xml'));
      let i = 0;
      m.set('xl/worksheets/sheet1.xml', s.replace(/<c r="B(\d)"([^>]*)>.*?<\/c>/g, (_, r, a) =>
        `<c r="B${r}"${a.replace(/ t="\w+"/, '')}>${i++ === 0 ? '<f t="shared" ref="B1:B3" si="0">A1*2</f>' : '<f t="shared" si="0"/>'}<v>0</v></c>`));
    });
    const out = await bytes(new SheetEditor().setCells('S', { B1: 100 }).edit(reader(src)));
    eq((await rows(out, { formulas: true })).map((r) => r.formulas?.[1]), [undefined, 'A2*2', 'A3*2']);
  });
  await t('.xlsm keeps its VBA project and content type through an edit', async () => {
    const vba = new Uint8Array(Array.from({ length: 3000 }, (_, i) => (i * 37) % 256));
    const src = await rezip(await write((w) => w.addSheet('Data', [['a', 1]])), async (m) => {
      const ct = new TextDecoder().decode(m.get('[Content_Types].xml'))
        .replace(/spreadsheetml\.sheet\.main\+xml/, 'NOPE').replace('application/vnd.openxmlformats-officedocument.NOPE', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml')
        .replace(/<Default /, '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/><Default ');
      m.set('[Content_Types].xml', ct);
      m.set('xl/_rels/workbook.xml.rels', new TextDecoder().decode(m.get('xl/_rels/workbook.xml.rels')).replace('</Relationships>',
        '<Relationship Id="rIdVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>'));
      m.set('xl/vbaProject.bin', vba);
    });
    ok((await text(src, '[Content_Types].xml')).includes('macroEnabled'), 'fixture not macro-enabled');
    const out = await bytes(new SheetEditor().setCells('Data', { B1: 2 }).insertRows('Data', 1, 1).edit(reader(src)));
    const p = await parts(out);
    ok(p.get('xl/vbaProject.bin')?.every((v, i) => v === vba[i]) && p.get('xl/vbaProject.bin').length === vba.length, 'VBA project changed');
    ok(new TextDecoder().decode(p.get('[Content_Types].xml')).includes('macroEnabled'), 'content type lost');
  });
  await t('deleteSheet: scoped names removed, other names become #REF!, formulas elsewhere untouched', async () => {
    const src = await write((w) => { w.addSheet('Keep', [[{ value: null, formula: 'Gone!A1+1' }]]); w.addSheet('Gone', [[1]]); },
      { definedNames: [{ name: 'Local', ref: 'Gone!$A$1', sheet: 'Gone' }, { name: 'Global', ref: 'Gone!$A$1' }] });
    const out = await bytes(new SheetEditor().deleteSheet('Gone').edit(reader(src)));
    const i = await new SheetReader().readWorkbook(reader(out));
    eq(i.definedNames.map((d) => d.name), ['Global']);
    ok(i.definedNames[0].ref.startsWith('#REF!'), i.definedNames[0].ref);
    eq((await rows(out, { formulas: true }))[0].formulas[0], 'Gone!A1+1');
  });
  await t('insertRows moves formulas, names, merges, CF, validations, links, filter, breaks, widths, tables, pictures, notes', async () => {
    const src = await write((w) => {
      w.addSheet('S', [
        ['h1', 'h2'], [1, 2], [3, { value: 'link', hyperlink: 'https://x.y', comment: { text: 'n' } }], [{ value: null, formula: 'SUM(A2:A3)' }, null],
      ], {
        mergeCells: ['A6:B6'], autoFilter: 'A1:B3', tables: [{ name: 'T_1', ref: 'A1:B3' }],
        conditionalFormats: [{ range: 'A2:A3', rule: { type: 'dataBar', color: 'FF06B6D4' } }],
        dataValidations: [{ sqref: 'B2:B3', type: 'whole', operator: 'between', formula1: '0', formula2: '9' }],
        images: [{ data: fx.png, at: 'D3' }], pageSetup: { printArea: 'A1:B4' },
      });
      w.addSheet('Other', [[{ value: null, formula: 'S!A3*2' }]]);
    }, { definedNames: [{ name: 'Rng', ref: 'S!$A$2:$A$3' }] });
    const out = await bytes(new SheetEditor().insertRows('S', 2, 2).edit(reader(src)));
    const s = await text(out, 'xl/worksheets/sheet1.xml');
    const need = { formula: /SUM\(A4:A5\)/, merge: /mergeCell ref="A8:B8"/, cf: /sqref="A4:A5"/, dv: /sqref="B4:B5"/, link: /hyperlink[^>]*ref="B5"/, filter: /autoFilter ref="A1:B5"/ };
    for (const [k, re] of Object.entries(need)) ok(re.test(s), `${k} not moved`);
    const p = await parts(out);
    const all = (re) => [...p.keys()].filter((n) => re.test(n)).map((n) => new TextDecoder().decode(p.get(n))).join('\n');
    ok(/ref="A1:B5"/.test(all(/tables\/table/)), 'table not grown');
    ok(/<xdr:row>4<\/xdr:row>/.test(all(/drawings\/drawing\d+\.xml$/)), 'picture not moved');
    ok(/ref="B5"/.test(all(/comments/)), 'note not moved');
    eq((await rows(out, { sheetName: 'Other', formulas: true }))[0].formulas[0], 'S!A5*2');
    const names = (await new SheetReader().readWorkbook(reader(out))).definedNames.map((d) => d.ref);
    ok(names.includes('S!$A$4:$A$5'), `defined name: ${names}`);
    ok(names.some((r) => /\$B\$6/.test(r)), `print area: ${names}`);
  });
  await t('Columns inserted inside a table are named Column1...; deleting its header row is refused', async () => {
    const src = await write((w) => w.addSheet('S', [['a', 'b', 'c'], [1, 2, 3]], { tables: [{ name: 'Tb', ref: 'A1:C2' }] }));
    const out = await bytes(new SheetEditor().insertColumns('S', 'B', 2).edit(reader(src)));
    const tbl = await text(out, 'xl/tables/table1.xml');
    ok(/name="Column1"/.test(tbl) && /name="Column2"/.test(tbl) && /ref="A1:E2"/.test(tbl), tbl.slice(0, 300));
    await throwsLike(() => bytes(new SheetEditor().deleteRows('S', 1, 1).edit(reader(src))), null, 'delete header row');
  });

  // ---------- Styles, formulas, conditional formats ----------
  await t('Styles example: cached SUM, default date format, internal link, rich text, data bar, images', async () => {
    const b = await write((w) => w.addSheet('Sales', [
      [{ value: 'Total Revenue', style: { font: { bold: true }, fill: { type: 'solid', fgColor: 'FF1E3A5F' } } }],
      [100], [250],
      [{ value: null, formula: '=SUM(A2:A3)' }],
      [new Date(Date.UTC(2026, 9, 8)), { value: new Date(Date.UTC(2026, 9, 8)), style: { numFmt: 'dd/mm/yyyy' } }],
      [{ value: 'Docs', hyperlink: 'https://example.com' }, { value: 'Back to top', hyperlink: '#Sales!A1' }],
      [{ value: null, richText: [{ text: 'Net ', font: { bold: true } }, { text: 'revenue', font: { color: 'FFC00000' } }] }],
    ], {
      freezePanes: { row: 1 }, autoFilter: 'A1:B1',
      conditionalFormats: [{ range: 'A2:A3', rule: { type: 'dataBar', color: 'FF06B6D4' } }],
      images: [{ data: fx.png, at: 'D1', height: 40 }, { data: fx.jpg, range: 'D4:H14', altText: 'Trend' }, { data: fx.gif, at: 'J1' }],
    }));
    const r = await rows(b, { formatted: true });
    eq(r[3].cells[0], 350, 'cached SUM');
    eq(r[4].formatted, ['2026-10-08', '08/10/2026']);
    const s = await text(b, 'xl/worksheets/sheet1.xml');
    ok(/location="Sales!A1"/.test(s), 'internal link');
    ok(/dataBar/.test(s), 'data bar');
    const d = await text(b, 'xl/drawings/drawing1.xml');
    // red-40x20.png at height 40 -> width 80 px = 762000 EMU, height 381000 EMU
    ok(/cx="762000" cy="381000"/.test(d), 'aspect ratio');
    ok(/descr="Trend"/.test(d) && /twoCellAnchor/.test(d), 'range image / alt text');
    eq((await parts(b)).has('xl/media/image3.gif'), true, 'gif');
  });
  await t('Colours: RGB hex C00000 and #C00000 are opaque ARGB', async () => {
    const b = await write((w) => w.addSheet('S', [[{ value: 1, style: { font: { color: 'C00000' } } }, { value: 1, style: { font: { color: '#C00000' } } }]]));
    eq((await rows(b, { styles: true }))[0].styles.map((s) => s.font.color), ['FFC00000', 'FFC00000']);
  });
  await t('A cell holds at most 32,767 characters', async () => {
    eq((await rows(await write((w) => w.addSheet('S', [['a'.repeat(32767)]]))))[0].cells[0].length, 32767);
    await throwsLike(async () => { const w = new SheetWriter(); w.addSheet('S', [['a'.repeat(32768)]]); await bytes(w.write()); }, null, '32,768 chars');
  });
  await t('A merge inside a table is refused (Excel tables cannot hold merged cells)', async () => {
    await throwsLike(() => write((w) => w.addSheet('S', [['a', 'b'], [1, 2]], { tables: [{ name: 'Tb', ref: 'A1:B2' }], mergeCells: ['A1:B1'] })), /merge/i, 'merge in table');
  });
  await t('Other sheet options example writes every option', async () => {
    const data = [['Name', 'Amount', 'Level'], ...Array.from({ length: 99 }, (_, i) => [i === 0 ? 'urgent' : `x${i}`, i * 20, i % 3])];
    const b = await write((w) => w.addSheet('Report', data, {
      conditionalFormats: [
        { range: 'B2:B100', rule: { type: 'cellIs', operator: 'greaterThan', formulae: [1000], style: { fill: { type: 'solid', fgColor: 'FFFFC7CE' } } } },
        { range: 'A2:A100', rule: { type: 'containsText', text: 'urgent', style: { font: { bold: true, color: 'FF9C0006' } } } },
        { range: 'C2:C100', rule: { type: 'iconSet', iconSet: '3TrafficLights1' } },
      ],
      dataValidations: [{ sqref: 'B2:B100', type: 'whole', operator: 'between', formula1: '0', formula2: '100', error: 'Use 0-100' }],
      tables: [{ name: 'Sales', ref: 'A1:C100' }],
      rows: { 1: { height: 24 }, 5: { outlineLevel: 1, hidden: true } },
      columns: [{ width: 30 }, { width: 12, outlineLevel: 1 }],
      protection: { password: 'secret', sort: true },
      pageSetup: { orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0, printArea: 'A1:C100', printTitleRows: '1', footer: '&CPage &P of &N' },
      tabColor: 'FF00B050', mergeCells: ['E1:G1'], columnWidths: [30, 12], autoFitColumns: true,
    }));
    const s = await text(b, 'xl/worksheets/sheet1.xml');
    for (const re of [/cellIs/, /containsText/, /3TrafficLights1/, /dataValidation /, /ht="24"/, /outlineLevel="1"/, /sheetProtection[^>]*password="/,
      /orientation="landscape"/, /paperSize="9"/, /fitToHeight="0"/, /&amp;CPage/, /tabColor rgb="FF00B050"/, /mergeCell ref="E1:G1"/, /width="30/])
      ok(re.test(s), `missing ${re}`);
    const wb = await text(b, 'xl/workbook.xml');
    ok(/_xlnm.Print_Area/.test(wb) && /_xlnm.Print_Titles/.test(wb), 'print area/titles');
    ok((await parts(b)).has('xl/tables/table1.xml'), 'table');
  });
  await t('Notes, including formatted notes', async () => {
    const b = await write((w) => w.addSheet('S', [[
      { value: 'Q3', comment: { text: 'Restated', author: 'Ana' } },
      { value: 'Q4', comment: { text: [{ text: 'Ana:', font: { bold: true } }, { text: ' restated' }] } },
    ]]));
    const res = await new SheetReader().parse(reader(b)); for await (const _ of res);
    eq((await res.getComments()).map((c) => c.text), ['Restated', 'Ana: restated']);
  });
  await t('Cached results: SUM/AVERAGE/COUNT/MIN/MAX/IF/CONCATENATE/&/comparisons/arithmetic, over formula cells; errors', async () => {
    const f = (formula) => ({ value: null, formula });
    const b = await write((w) => w.addSheet('S', [
      [2, 4, f('A1+B1')],
      [f('SUM(A1:C1)'), f('AVERAGE(A1:B1)'), f('COUNT(A1:C1)'), f('MIN(A1:C1)'), f('MAX(A1:C1)'), f('IF(C1>5,"big","small")'),
        f('CONCATENATE("a",A1)'), f('"x"&B1'), f('A1<B1'), f('(A1+B1)*3-1/2'), f('1/0')],
    ]));
    eq((await rows(b))[1].cells, [12, 3, 3, 2, 6, 'big', 'a2', 'x4', true, 17.5, '#DIV/0!']);
    async function* g() { yield [1]; yield [f('A1*2')]; }
    const s = await text(await bytes(new SheetWriter().addSheet('S', g()).write()), 'xl/worksheets/sheet1.xml');
    ok(!/<f>A1\*2<\/f><v>/.test(s), 'AsyncIterable row got a cached value');
  });
  await t('sharedStrings: true stores each string once; file smaller when many strings repeat', async () => {
    const data = Array.from({ length: 50000 }, (_, i) => [`value number ${i % 1000}`, `other ${(i * 7) % 1000}`]);
    const inline = await write((w) => w.addSheet('S', data));
    const shared = await write((w) => w.addSheet('S', data), { sharedStrings: true });
    const sst = await text(shared, 'xl/sharedStrings.xml');
    ok(/uniqueCount="2000"/.test(sst), 'not deduplicated');
    ok(shared.length < inline.length, `shared ${shared.length} vs inline ${inline.length}`);
    eq((await rows(shared))[49999].cells, data[49999]);
  });
  return results;
}
