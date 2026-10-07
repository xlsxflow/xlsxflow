// Builds reader fixtures with other JS xlsx writers (installed under scripts/competitors).
//   node scripts/fixtures/make-js.mjs && python scripts/fixtures/make-fixtures.py --oracle
import { createRequire } from 'module';
import { fileURLToPath } from 'url';
import path from 'path';

const here = path.dirname(fileURLToPath(import.meta.url));
const require = createRequire(path.join(here, '../competitors/package.json'));
const FIX = path.join(here, '../../test/fixtures');

const bulk = Array.from({ length: 5000 }, (_, i) => [i, `name ${i % 500}`, i * 0.5, i % 2 === 0, `unique-${i}`]);
const mixed = [
  ['text', 42, -3.5, 1e20, 1e-10, 0.1 + 0.2],
  [true, false, null, 'x & y < z > "q"', 'emoji \u{1F600} cjk 漢字'],
  ['line1\nline2', '  padded  ', '00123'],
];

// ExcelJS: shared strings, rich text, real Date cells, formulas with cached results
{
  const ExcelJS = require('exceljs');
  const wb = new ExcelJS.Workbook();
  const t = wb.addWorksheet('Types');
  mixed.forEach(r => t.addRow(r));
  t.getCell('A5').value = { richText: [{ font: { bold: true }, text: 'bold' }, { text: ' then plain' }] };
  t.getCell('B5').value = { formula: 'B1*2', result: 84 };
  t.getCell('C5').value = { formula: '"a"&"b"', result: 'ab' };
  t.getCell('D5').value = { error: '#DIV/0!' };
  const d = wb.addWorksheet('Dates');
  d.getCell('A1').value = new Date(Date.UTC(2024, 1, 29, 18, 30));
  d.getCell('A1').numFmt = 'yyyy-mm-dd hh:mm';
  d.getCell('B1').value = new Date(Date.UTC(1900, 0, 1));
  d.getCell('C1').value = 0.25;
  d.getCell('C1').numFmt = '0.00%';
  wb.addWorksheet('Bulk').addRows(bulk);
  t.mergeCells('A7:C7');
  await wb.xlsx.writeFile(path.join(FIX, 'exceljs.xlsx'));
}

// SheetJS: its own string/number/bool/date encodings, no styles part by default
{
  const XLSX = require('xlsx');
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(mixed), 'Types');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([[new Date(Date.UTC(2024, 1, 29))]], { cellDates: true }), 'Dates');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(bulk), 'Bulk');
  XLSX.writeFile(wb, path.join(FIX, 'sheetjs.xlsx'), { compression: true });
}

// xlsx-populate: starts from its own template workbook
{
  const XlsxPopulate = require('xlsx-populate');
  const wb = await XlsxPopulate.fromBlankAsync();
  const s = wb.sheet(0).name('Types');
  s.cell('A1').value(mixed.map(r => r.map(v => v ?? undefined)));
  s.cell('A6').value(45000).style('numberFormat', 'dd/mm/yyyy');
  wb.addSheet('Bulk').cell('A1').value(bulk);
  await wb.toFileAsync(path.join(FIX, 'xlsx-populate.xlsx'));
  // openpyxl can't parse xlsx-populate's template stylesheet, so expectations come from the source data
  const trim = r => { r = [...r]; while (r.length && r[r.length - 1] == null) r.pop(); return r; };
  const expected = { source: 'source data', sheets: [
    { name: 'Types', rows: [...mixed.map((r, i) => [i + 1, trim(r)]), [6, ['2023-03-15T00:00:00.000Z']]] },
    { name: 'Bulk', rows: bulk.map((r, i) => [i + 1, r]) },
  ] };
  (await import('fs')).writeFileSync(path.join(FIX, 'xlsx-populate.expected.json'), JSON.stringify(expected));
}

// XlsxFlow's own writer: styles, formulas, merges, AsyncIterable rows
{
  const { SheetWriter } = await import('../../src/index.ts').catch(() => import('../../dist/index.mjs'));
  const fs = await import('fs');
  const w = new SheetWriter();
  w.addSheet('Types', [...mixed, [{ value: null, formula: 'B1*2' }, { value: 'styled', style: { font: { bold: true } } }]], { mergeCells: ['A6:B6'] });
  w.addSheet('Dates', [[{ value: 45351.75, style: { numFmt: 'yyyy-mm-dd hh:mm' } }]]);
  w.addSheet('Bulk', (async function* () { yield* bulk; })());
  fs.writeFileSync(path.join(FIX, 'xlsxflow.xlsx'), new Uint8Array(await new Response(w.write()).arrayBuffer()));
}
console.log('JS fixtures written');
