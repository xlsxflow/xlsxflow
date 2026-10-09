# Migrating from SheetJS

This guide shows each common SheetJS (`xlsx`) task next to its XlsxFlow equivalent. Every XlsxFlow example here runs as written against `@xlsxflow/core`; the SheetJS examples were checked against SheetJS Community Edition 0.20.3.

```bash
npm uninstall xlsx
npm install @xlsxflow/core
```

## Why switch

- **The npm package is stuck.** `npm install xlsx` gets 0.18.5, from March 2022. npm audit reports two advisories against it, [CVE-2023-30533](https://github.com/advisories/GHSA-4r6h-8v6p-xvw6) (prototype pollution) and [CVE-2024-22363](https://github.com/advisories/GHSA-5pgg-2g8v-p4x9) (ReDoS). The fixed versions are published only on SheetJS's own CDN, so Dependabot and Renovate don't update them.
- **Memory.** SheetJS reads the whole workbook into memory before you see the first cell, and builds the whole file in memory before writing it. XlsxFlow streams rows in and out, so memory stays flat: 10M cells are written and read back in CI with Node's heap capped at 32 MB.
- **Styles.** The Community Edition neither reads nor writes cell styles, and saving a file it read drops the styles that were there. XlsxFlow reads and writes fonts, fills, borders, alignment and number formats, plus conditional formats, validations, tables and images.
- **Editing.** `SheetEditor` changes the cells you name and copies every other part of the file, charts and pivot tables included.

SheetJS reads and writes many more formats (`.xlsb`, `.numbers`, `.xls` writing and others). If you need those, keep SheetJS for them.

## What works differently

| | SheetJS | XlsxFlow |
|---|---|---|
| Model | The whole workbook as `wb.Sheets[name]['A1']` cell objects | Rows stream past once; keep the ones you need |
| Dates | Serial numbers such as `46304`; with `cellDates: true`, `Date` objects shifted by the local time zone | ISO-8601 strings in UTC: `'2026-10-09T00:00:00.000Z'` |
| Empty cells in `sheet_to_json` | The key is left out, unless you pass `defval` | The key is there with `null` |
| Blank rows with `header: 1` | Kept as `[]` | Skipped; `row.rowNumber` shows the gap |
| Formulas | `cell.f`, read with `cellFormula` | Pass `{ formulas: true }` to get `row.formulas` |
| Column widths | `ws['!cols'] = [{ wch: 20 }]` | `columnWidths: [20]` |
| Merged cells | `ws['!merges']` as `{ s, e }` objects | `mergeCells: ['A1:C1']` |

The time zone point matters in practice: in a UTC+5:30 time zone, SheetJS with `cellDates: true` read the date 9 October 2026 as `2026-10-08T18:30:00.000Z`. XlsxFlow returns `2026-10-09T00:00:00.000Z` wherever it runs.

## Reading

### Rows as objects

SheetJS:

```js
import * as XLSX from 'xlsx';
import * as fs from 'node:fs';
XLSX.set_fs(fs);

const workbook = XLSX.readFile('report.xlsx');
const people = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
```

XlsxFlow:

```js
import { XlsxFlow, sheetToJson } from '@xlsxflow/core';

const people = await sheetToJson(await XlsxFlow.readFile('report.xlsx'));
```

`sheetToJson` holds the result in memory, as `sheet_to_json` does. For large files, loop over the rows instead.

### Rows as arrays

SheetJS:

```js
const rows = XLSX.utils.sheet_to_json(workbook.Sheets.Sheet1, { header: 1 });
for (const cells of rows) console.log(cells);
```

XlsxFlow streams them one at a time:

```js
import { XlsxFlow } from '@xlsxflow/core';

for await (const row of await XlsxFlow.readFile('report.xlsx', { sheetName: 'Sheet1' })) {
  console.log(row.rowNumber, row.cells);
}
```

### A buffer, a download or an upload

SheetJS:

```js
const workbook = XLSX.read(await file.arrayBuffer());
```

XlsxFlow takes any `Blob`, including a `File` from `<input type="file">` or a `fetch` response:

```js
import { SheetReader, createBlobReader } from '@xlsxflow/core';

const reader = new SheetReader();
const { sheets } = await reader.readWorkbook(createBlobReader(file)); // like workbook.SheetNames, with visibility
for await (const row of await reader.parse(createBlobReader(file), { sheetName: sheets[0].name })) {
  console.log(row.cells);
}
```

For a Node `Buffer`, pass `createBlobReader(new Blob([buffer]))`.

### CSV out

SheetJS:

```js
const csv = XLSX.utils.sheet_to_csv(workbook.Sheets.Sheet1);
```

XlsxFlow:

```js
import { XlsxFlow, streamToCsv } from '@xlsxflow/core';

const csv = await streamToCsv(await XlsxFlow.readFile('report.xlsx'));
```

## Writing

### From arrays

SheetJS:

```js
const sheet = XLSX.utils.aoa_to_sheet([['Name', 'Total'], ['Ana', 3]]);
sheet['!cols'] = [{ wch: 20 }, { wch: 10 }];
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, sheet, 'Data');
XLSX.writeFile(workbook, 'out.xlsx');
```

XlsxFlow:

```js
import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';
import { SheetWriter } from '@xlsxflow/core';

const writer = new SheetWriter().addSheet('Data', [['Name', 'Total'], ['Ana', 3]], { columnWidths: [20, 10] });
await writer.write().pipeTo(Writable.toWeb(createWriteStream('out.xlsx')));
```

### From objects

SheetJS:

```js
const sheet = XLSX.utils.json_to_sheet(people);
```

XlsxFlow writes arrays, so put the keys in a header row:

```js
import { SheetWriter } from '@xlsxflow/core';

const keys = Object.keys(people[0]);
const writer = new SheetWriter().addSheet('People', [keys, ...people.map((p) => keys.map((k) => p[k]))]);
```

### To a buffer, or a download in the browser

SheetJS:

```js
const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
XLSX.writeFile(workbook, 'out.xlsx'); // in a browser, starts a download
```

XlsxFlow: `write()` returns a `ReadableStream`, which a `Response` collects:

```js
import { SheetWriter } from '@xlsxflow/core';

const stream = new SheetWriter().addSheet('Data', [['a', 1]]).write();
const buffer = Buffer.from(await new Response(stream).arrayBuffer());
```

For a browser download, see the [browser recipe](../examples/browser-download.mjs). In a server, return the stream itself, so the file is sent as it is written; see the [recipes](../examples/README.md).

### Formulas, merged cells, styles

SheetJS:

```js
sheet['C2'] = { t: 'n', f: 'A2+B2' };
sheet['!merges'] = [XLSX.utils.decode_range('A1:C1')];
// Cell styles need SheetJS Pro
```

XlsxFlow:

```js
import { SheetWriter } from '@xlsxflow/core';

const writer = new SheetWriter().addSheet('Data', [
  [{ value: 'Q3 report', style: { font: { bold: true, size: 14 }, alignment: { horizontal: 'center' } } }],
  [1, 2, { value: null, formula: 'A2+B2' }],
], { mergeCells: ['A1:C1'], freezePanes: { row: 1 } });
```

Simple formulas also get a cached result, so other programs that read the file without recalculating see a value. See the [README](../README.md#styles-formulas--conditional-formats) for fills, borders, number formats, conditional formats, validations, tables and images.

## Not supported

- **Other formats.** XlsxFlow reads `.xlsx`, `.xls` and `.ods`, and writes `.xlsx` and `.ods`. It doesn't read `.xlsb`, `.numbers` or HTML tables, and doesn't write `.xls`.
- **Random access to any cell.** There is no in-memory workbook. Collect the rows you need into an array or a `Map`; to change cells in a file, use `SheetEditor`.
