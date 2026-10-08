// Source workbook for test/fixtures/libreoffice-made.ods: npx tsx scripts/fixtures/make-libreoffice-source.ts <out.xlsx>,
// then make-libreoffice-ods.py turns it into the .ods
import { writeFileSync } from 'node:fs';
import { SheetWriter } from '../../src/index';

const long = Array.from({ length: 1500 }, (_, i) => `abc${i} ` + (i % 100 === 0 ? '\u3042' : '')).join('');
const many = Array.from({ length: 3000 }, (_, i) => [`String number ${i}`, i % 7 === 0 ? `Unicode \u0416${i}` : `plain ${i}`, i * 1.5]);

const w = new SheetWriter({
  properties: { title: 'Fixture title', creator: 'Fixture author', subject: 'Fixture subject' },
  definedNames: [{ name: 'Numbers', ref: 'Data!$C$10:$C$3009' }, { name: 'Answer', ref: 'Data!$B$1' }],
});
w.addSheet('Data', [
  ['Text', 42, 3.14159, { value: new Date(Date.UTC(2026, 9, 8)), style: { numFmt: 'yyyy-mm-dd' } }, true,
    { value: null, formula: '1/0' }, { value: null, formula: 'B1*2' }, { value: null, formula: 'A1&"!"' },
    'h\u00e9llo \u2014 \u65e5\u672c', 1234567890123, -0.5, { value: 0.256, style: { numFmt: '0.0%' } },
    { value: new Date(Date.UTC(2026, 9, 8, 14, 4, 30)), style: { numFmt: 'yyyy-mm-dd hh:mm:ss' } },
    { value: null, formula: 'B1>40' }, { value: 12.5, style: { numFmt: '#,##0.00 "USD"' } }, 100, 'hidden column'],
  ['hidden row'],
  ['Merged'],
  [],
  [long],
  [], [], [], [],
  ...many,
], {
  mergeCells: ['A3:C3'],
  rows: { 2: { hidden: true } },
  columns: Array.from({ length: 17 }, (_, i) => i === 16 ? { hidden: true } : {}),
  freezePanes: { row: 1, col: 1 },
});
w.addSheet('Second', [['second sheet'], [null, 7]]);
w.addSheet('Secret', [['hidden sheet']], { state: 'hidden' });

const out = process.argv[2];
new Response(w.write()).arrayBuffer().then(b => { writeFileSync(out, new Uint8Array(b)); console.log('wrote', out); });
