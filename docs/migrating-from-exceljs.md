# Migrating from ExcelJS

This guide shows each common ExcelJS task next to its XlsxFlow equivalent. Every XlsxFlow example here runs as written against `@xlsxflow/core`; the ExcelJS examples were checked against ExcelJS 4.4.0.

```bash
npm uninstall exceljs
npm install @xlsxflow/core
```

## Why switch

- **Memory.** `workbook.xlsx.readFile` loads the whole workbook into memory. XlsxFlow streams rows in and out, so memory stays flat: 10M cells are written and read back in CI with Node's heap capped at 32 MB. ExcelJS has separate streaming classes, but they don't support everything the normal ones do.
- **Editing keeps the rest of the file.** ExcelJS rebuilds a file from its own model when it saves, so anything the model doesn't cover is lost. Loading a workbook with a chart or pivot tables and saving it again drops them. `SheetEditor` changes the cells you name and copies every other part of the file unchanged.
- **Maintenance.** ExcelJS's last release, 4.4.0, came out in October 2023.
- **Runtimes.** XlsxFlow uses Web APIs only, so the same code runs in Node, Bun, Deno, browsers and Cloudflare Workers.

## What works differently

| | ExcelJS | XlsxFlow |
|---|---|---|
| Model | The workbook is loaded into memory; any cell can be read or changed with `getCell` | Rows stream past once. Keep the rows you need, or change cells with `SheetEditor` |
| Row values | `row.values`, 1-based: `values[0]` is empty | `row.cells`, 0-based. `row.rowNumber` is the 1-based row number |
| Blank rows | Skipped by `eachRow` | Skipped too; `rowNumber` shows the gap |
| Dates | `Date` objects | ISO-8601 strings in UTC: `'2026-10-09T00:00:00.000Z'`. Use `new Date(value)` |
| Formula cells | `{ formula, result }` | `row.cells` holds the result. Pass `{ formulas: true }` to get `row.formulas` |
| Hyperlink cells | `{ text, hyperlink }` | `row.cells` holds the text. The links are in `(await rows.getMetadata()).hyperlinks` |
| Rich text cells | `{ richText: [...] }` | `row.cells` holds the plain text. Pass `{ richText: true }` to get `row.richText` |
| Colours | `{ argb: 'FFFF0000' }` | `'FFFF0000'` |
| Solid fill | `{ type: 'pattern', pattern: 'solid', fgColor: { argb } }` | `{ type: 'solid', fgColor }` |

## Reading

### A file on disk

ExcelJS:

```js
import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile('report.xlsx');
workbook.getWorksheet('Sheet1').eachRow((row, rowNumber) => {
  console.log(rowNumber, row.values.slice(1));
});
```

XlsxFlow:

```js
import { XlsxFlow } from '@xlsxflow/core';

for await (const row of await XlsxFlow.readFile('report.xlsx', { sheetName: 'Sheet1' })) {
  console.log(row.rowNumber, row.cells);
}
```

Omit `sheetName` to read the first sheet. The same loop replaces ExcelJS's streaming `WorkbookReader`.

### A buffer, a download or an upload

ExcelJS:

```js
await workbook.xlsx.load(buffer);
```

XlsxFlow takes any `Blob`, so a `File` from an upload or `fetch` works directly:

```js
import { SheetReader, createBlobReader } from '@xlsxflow/core';

const rows = await new SheetReader().parse(createBlobReader(new Blob([buffer])));
for await (const row of rows) console.log(row.cells);
```

### Rows as objects

ExcelJS has no built-in for this; XlsxFlow uses the header row as keys:

```js
import { XlsxFlow, sheetToJson } from '@xlsxflow/core';

const people = await sheetToJson(await XlsxFlow.readFile('report.xlsx'));
// [{ Name: 'Ana', Date: '2026-10-09T00:00:00.000Z', Total: 3 }, ...]
```

`sheetToJson` holds the whole result in memory; loop over the rows instead for large files.

### Sheet names, merged cells, formulas and styles

ExcelJS:

```js
const names = workbook.worksheets.map((ws) => ws.name);
const sheet = workbook.getWorksheet('Sheet1');
const formula = sheet.getCell('C2').formula;
const bold = sheet.getCell('A1').font?.bold;
```

XlsxFlow:

```js
import { SheetReader, createBlobReader } from '@xlsxflow/core';

const reader = new SheetReader();
const blob = new Blob([buffer]);
const { sheets } = await reader.readWorkbook(createBlobReader(blob)); // [{ name, state }, ...]

const rows = await reader.parse(createBlobReader(blob), { sheetName: sheets[0].name, formulas: true, styles: true });
for await (const row of rows) {
  if (row.rowNumber === 1) console.log(row.styles?.[0]?.font?.bold);
  if (row.rowNumber === 2) console.log(row.formulas?.[2]);
}
const { mergedCells, hyperlinks, freezePanes } = await rows.getMetadata();
```

## Writing

### Columns with headers and keys

ExcelJS:

```js
const workbook = new ExcelJS.Workbook();
const sheet = workbook.addWorksheet('Users');
sheet.columns = [
  { header: 'Name', key: 'name', width: 20 },
  { header: 'Email', key: 'email', width: 30 },
];
for (const user of users) sheet.addRow(user);
await workbook.xlsx.writeFile('users.xlsx');
```

XlsxFlow writes rows as arrays, so map the objects with the same column list:

```js
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';
import { SheetWriter } from '@xlsxflow/core';

const columns = [
  { header: 'Name', key: 'name', width: 20 },
  { header: 'Email', key: 'email', width: 30 },
];
const writer = new SheetWriter().addSheet('Users', [
  columns.map((c) => c.header),
  ...users.map((user) => columns.map((c) => user[c.key])),
], { columns: columns.map((c) => ({ width: c.width })) });

await writer.write().pipeTo(Writable.toWeb(createWriteStream('users.xlsx')));
```

### Saving to a buffer or a Blob

ExcelJS:

```js
const buffer = await workbook.xlsx.writeBuffer();
```

XlsxFlow: `write()` returns a `ReadableStream`, which a `Response` collects:

```js
import { SheetWriter } from '@xlsxflow/core';

const stream = new SheetWriter().addSheet('Data', [['a', 1]]).write();
const buffer = Buffer.from(await new Response(stream).arrayBuffer()); // or .blob() in a browser
```

In a server, return the stream itself instead of a buffer, so the file is sent as it is written. See [recipes](../examples/README.md).

### Millions of rows

ExcelJS needs its separate streaming writer, with a `commit()` per row:

```js
const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: 'big.xlsx' });
const sheet = workbook.addWorksheet('Data');
for (let i = 0; i < 1_000_000; i++) sheet.addRow([i, `row ${i}`]).commit();
sheet.commit();
await workbook.commit();
```

XlsxFlow's one writer takes an async generator and pulls rows only as fast as the file is written:

```js
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';
import { SheetWriter } from '@xlsxflow/core';

async function* rows() {
  for (let i = 0; i < 1_000_000; i++) yield [i, `row ${i}`];
}
await new SheetWriter().addSheet('Data', rows()).write().pipeTo(Writable.toWeb(createWriteStream('big.xlsx')));
```

### Styles

ExcelJS sets properties on each cell:

```js
const cell = sheet.getCell('A1');
cell.value = 'Total';
cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E3A5F' } };
cell.border = { bottom: { style: 'thin', color: { argb: 'FF000000' } } };
cell.alignment = { horizontal: 'center', wrapText: true };
sheet.getCell('B1').numFmt = '#,##0.00';
```

XlsxFlow puts the value and its style in the row:

```js
import { SheetWriter } from '@xlsxflow/core';

const writer = new SheetWriter().addSheet('Report', [[
  {
    value: 'Total',
    style: {
      font: { bold: true, color: 'FFFFFFFF' },
      fill: { type: 'solid', fgColor: 'FF1E3A5F' },
      border: { bottom: { style: 'thin', color: 'FF000000' } },
      alignment: { horizontal: 'center', wrapText: true },
    },
  },
  { value: 1234.5, style: { numFmt: '#,##0.00' } },
]]);
```

Formulas, hyperlinks, notes and rich text are cell properties too:

```js
import { SheetWriter } from '@xlsxflow/core';

const writer = new SheetWriter().addSheet('Data', [
  [1, 2, { value: null, formula: 'A1+B1' }],
  [{ value: 'Docs', hyperlink: 'https://example.com' }, { value: 'Q3', comment: { text: 'Restated', author: 'Ana' } }],
  [{ value: null, richText: [{ text: 'Net ', font: { bold: true } }, { text: 'revenue' }] }],
]);
```

### Sheet features

ExcelJS:

```js
sheet.mergeCells('A1:C1');
sheet.views = [{ state: 'frozen', ySplit: 1 }];
sheet.autoFilter = 'A1:C1';
sheet.getCell('B2').dataValidation = { type: 'list', formulae: ['"Yes,No"'] };
sheet.addConditionalFormatting({
  ref: 'C2:C100',
  rules: [{ type: 'cellIs', operator: 'greaterThan', formulae: [1000], style: { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } } } }],
});
const imageId = workbook.addImage({ buffer: png, extension: 'png' });
sheet.addImage(imageId, 'E1:G6');
```

XlsxFlow takes them as sheet options:

```js
import { SheetWriter } from '@xlsxflow/core';

const writer = new SheetWriter().addSheet('Data', [['Region', 'Approved', 'Total']], {
  mergeCells: ['A1:C1'],
  freezePanes: { row: 1 },
  autoFilter: 'A1:C1',
  dataValidations: [{ sqref: 'B2', type: 'list', formula1: '"Yes,No"' }],
  conditionalFormats: [{
    range: 'C2:C100',
    rule: { type: 'cellIs', operator: 'greaterThan', formulae: [1000], style: { fill: { type: 'solid', fgColor: 'FFFFC7CE' } } },
  }],
  images: [{ data: png, range: 'E1:G6' }],
});
```

Tables, outline levels, row heights, protection, page setup and tab colours are sheet options as well; see the [README](../README.md#styles-formulas--conditional-formats).

## Editing an existing file

ExcelJS loads the file, changes it in memory and writes everything out again:

```js
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.readFile('report.xlsx');
const sheet = workbook.getWorksheet('Sheet1');
sheet.getCell('B2').value = 42;
sheet.getCell('C2').value = { formula: 'B2*2' };
sheet.spliceRows(5, 0, [], [], []);
await workbook.xlsx.writeFile('report-edited.xlsx');
```

XlsxFlow lists the changes, then streams the file through them:

```js
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';
import { SheetEditor, createFileReader } from '@xlsxflow/core';

const editor = new SheetEditor();
editor.setCells('Sheet1', { B2: 42, C2: { formula: 'B2*2' } });
editor.insertRows('Sheet1', 5, 3);

const file = await createFileReader('report.xlsx');
try {
  await editor.edit(file).pipeTo(Writable.toWeb(createWriteStream('report-edited.xlsx')));
} finally {
  await file.close();
}
```

Edited cells keep their style. Inserting or deleting rows and columns moves formulas, merged cells, tables, charts and everything else that points at the moved cells, as Excel does. Charts, pivot tables and any other parts are copied unchanged.

## Not supported

- **Reading or changing any cell at any time.** There is no in-memory workbook. To look cells up, collect the rows you need into an array or a `Map`; to change cells, use `SheetEditor`.
- **Charts and pivot tables** are created by [`@xlsxflow/pro`](../README.md#free-and-pro); XlsxFlow keeps existing ones when it edits a file.
