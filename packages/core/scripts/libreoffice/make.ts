// Builds edited workbooks and the values a spreadsheet app must compute for them. Run by check.py,
// which first writes the openpyxl-made sources (openpyxl-rows.xlsx, openpyxl-cols.xlsx) into the
// output folder given as the first argument.
import { readFileSync, writeFileSync } from 'fs';
import { join } from 'path';
import { SheetEditor, SheetWriter, createBlobReader, parseCsv } from '../../src/index';

const PNG = Uint8Array.from(atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='), c => c.charCodeAt(0));
const bytes = async (s: ReadableStream<Uint8Array>) => new Uint8Array(await new Response(s).arrayBuffer());
const edit = async (src: Uint8Array, e: SheetEditor) => bytes(e.edit(createBlobReader(new Blob([src]))));
const f = (formula: string) => ({ value: null, formula });

async function main(out: string) {
  // File name -> { "Sheet!A1": value the app must show }
  const expected: Record<string, Record<string, unknown>> = {};
  const save = (name: string, data: Uint8Array, expect: Record<string, unknown>) => {
    writeFileSync(join(out, name), data);
    expected[name] = expect;
  };

  // Cell edits and restyles
  {
    const w = new SheetWriter();
    w.addSheet('Data', [[{ value: 'Title', style: { font: { bold: true } } }, 1, f('SUM(B1:B5)')], [null, 2], [null, 3], [null, 4], [null, 5]]);
    save('cells.xlsx', await edit(await bytes(w.write()), new SheetEditor().setCells('Data', {
      B2: 100, D1: { formula: 'C1*2' }, A1: { style: { font: { italic: true }, fill: { type: 'solid', fgColor: 'FFFFFF00' } } },
      B5: { value: 50, style: { numFmt: '0.00' } },
    })), { 'Data!C1': 158, 'Data!D1': 316, 'Data!A1': 'Title', 'Data!B5': 50 });
  }

  // Sheets added and deleted
  {
    const w = new SheetWriter();
    w.addSheet('One', [[7]]).addSheet('Two', [[1]]).addSheet('Three', [[f('One!A1*10')]]);
    save('sheets.xlsx', await edit(await bytes(w.write()), new SheetEditor().deleteSheet('Two')
      .addSheet('Added', [[f('One!A1+Three!A1'), { value: 'x', style: { font: { bold: true } } }]])),
    { 'Three!A1': 70, 'Added!A1': 77, 'Added!B1': 'x' });
  }

  // Rows inserted and deleted on a sheet with a table, note, merge, rules, validation and picture
  {
    const w = new SheetWriter();
    w.addSheet('Data', [
      ['Name', 'Qty', f('SUM(B2:B11)')],
      ...Array.from({ length: 10 }, (_, i) => [i === 4 ? { value: `item ${i + 1}`, comment: 'note' } : `item ${i + 1}`, i + 1]),
    ], {
      mergeCells: ['D5:E6'], tables: [{ name: 'T', ref: 'A1:B11' }],
      conditionalFormats: [{ range: 'B2:B11', rule: { type: 'cellIs', operator: 'greaterThan', formulae: [5], style: { font: { bold: true } } } }],
      dataValidations: [{ sqref: 'B2:B11', type: 'whole', operator: 'greaterThan', formula1: '0' }],
      images: [{ data: PNG, range: 'G8:H10' }], pageSetup: { printArea: 'A1:B11' },
    });
    w.addSheet('Other', [[f('Data!B6')], [f('SUM(Data!B2:B11)')], [f('Data!B9')]]);
    // Insert 2 rows at 3, then delete row 11 (the old row 9, value 8)
    save('rows.xlsx', await edit(await bytes(w.write()), new SheetEditor().insertRows('Data', 3, 2).deleteRows('Data', 11)),
      { 'Data!C1': 47, 'Data!B8': 5, 'Other!A1': 5, 'Other!A2': 47, 'Other!A3': '#REF!' });
  }

  // Columns inserted and deleted
  {
    const w = new SheetWriter();
    w.addSheet('Data', [[1, 2, 3, 4, 5, f('SUM(A1:E1)'), f('C1*2')]], { columns: [{ width: 8 }, { width: 9 }, { width: 10 }] });
    w.addSheet('Other', [[f('Data!D1')], [f('Data!C1')]]);
    // Insert a column at B, then delete column E (the old D, value 4)
    save('cols.xlsx', await edit(await bytes(w.write()), new SheetEditor().insertColumns('Data', 'B').deleteColumns('Data', 'E')),
      { 'Data!E1': 5, 'Data!F1': 11, 'Data!G1': 6, 'Other!A1': '#REF!', 'Other!A2': 3 });
  }

  // Writer features and CSV
  {
    const w = new SheetWriter();
    w.addSheet('Report', [
      ['Name', 'Amount'],
      [{ value: 'a', comment: { author: 'Ana', text: [{ text: 'Bold', font: { bold: true } }, { text: ' plain' }] } }, 10],
      ['b', 20],
      ['Total', f('SUM(B2:B3)')],
    ], {
      tables: [{ name: 'Sales', ref: 'A1:B3' }],
      conditionalFormats: [{ range: 'B2:B3', rule: { type: 'iconSet', iconSet: '3TrafficLights1' } }],
      protection: { password: 'secret' }, pageSetup: { orientation: 'landscape', printTitleRows: '1' },
      rows: { 1: { height: 24 } }, tabColor: 'FF00B050',
    });
    w.addSheet('CSV', parseCsv('x,y\n1,"a,b"\n2,true\n'));
    save('writer.xlsx', await bytes(w.write()), { 'Report!B4': 30, 'CSV!A3': 2, 'CSV!B2': 'a,b', 'CSV!B3': true });
  }

  // openpyxl-made workbooks with a chart, filters, page breaks and print titles
  const rowsSrc = new Uint8Array(readFileSync(join(out, 'openpyxl-rows.xlsx')));
  save('openpyxl-rows-insert.xlsx', await edit(rowsSrc, new SheetEditor().insertRows('Data', 5, 3)),
    { 'Data!D1': 550, 'Other!A1': 70, 'Other!A2': 550 });
  save('openpyxl-rows-delete.xlsx', await edit(rowsSrc, new SheetEditor().deleteRows('Data', 7, 2)),
    { 'Data!D1': 420, 'Other!A1': '#REF!', 'Other!A2': 420 });
  const colsSrc = new Uint8Array(readFileSync(join(out, 'openpyxl-cols.xlsx')));
  save('openpyxl-cols-insert.xlsx', await edit(colsSrc, new SheetEditor().insertColumns('Data', 'B', 2)),
    { 'Data!H1': 360, 'Data!H2': 15, 'Other!A1': 6, 'Other!A2': 414 });
  save('openpyxl-cols-delete.xlsx', await edit(colsSrc, new SheetEditor().deleteColumns('Data', 'B')),
    { 'Data!E1': 360, 'Data!E2': 15, 'Other!A1': 6, 'Other!A2': 414 });

  writeFileSync(join(out, 'expected.json'), JSON.stringify(expected, null, 1));
}

main(process.argv[2]).catch(err => {
  console.error(err);
  process.exit(1);
});
