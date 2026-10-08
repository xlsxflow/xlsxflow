<div align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/xlsxflow/xlsxflow/main/assets/logo-wordmark-dark.svg" />
    <img src="https://raw.githubusercontent.com/xlsxflow/xlsxflow/main/assets/logo-wordmark.svg" alt="XlsxFlow" width="320" />
  </picture>
  <p><strong>Streaming .xlsx reader, writer and editor for JavaScript. Also reads .xls and .ods, and writes .ods.</strong></p>
  
  [![npm version](https://img.shields.io/npm/v/@xlsxflow/core.svg?style=flat-square)](https://www.npmjs.com/package/@xlsxflow/core)
  [![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square)](https://opensource.org/licenses/MIT)

  <p>
    <a href="#features">Features</a> •
    <a href="#installation">Installation</a> •
    <a href="#quick-start">Quick Start</a> •
    <a href="#styles-formulas--conditional-formats">Styles &amp; Formulas</a> •
    <a href="#compared-with-sheetjs-and-exceljs">Comparison</a> •
    <a href="#benchmarks">Benchmarks</a> •
    <a href="#free-and-pro">Free and Pro</a>
  </p>
</div>

---

**XlsxFlow** is a zero-dependency streaming reader, writer and editor for OpenXML (`.xlsx`) files. Built on native Web APIs (like `TransformStream` and `CompressionStream`), it handles millions of cells in flat memory.

Rows are read and written one at a time instead of loading the whole workbook, so it suits browsers, servers and edge runtimes alike.

## Features

- **No dependencies**: TypeScript on Web APIs (`ReadableStream`, `CompressionStream`, `Blob`).
- **Streaming**: rows are read and written one at a time, so memory stays flat as files grow (10M cells written with about 1 MB of extra heap; see [Benchmarks](#benchmarks)).
- **Runs anywhere with Web APIs**: tested on Node 20.12+, Bun, browsers and Cloudflare Workers (without `nodejs_compat`). Deno provides the same APIs but is not tested yet.
- **Read, write and edit**: stream rows out of a file, generate one on the fly, or change cells, rows, columns and sheets of an existing file while keeping everything else in it.
- **Styles and formulas**: fonts, fills, borders, alignment, number formats, conditional formats, validations, tables, notes, hyperlinks, autofilters, images, protection and page setup. Formulas and styles read back too.
- **Older and open formats**: the same reader opens Excel 97-2003 `.xls` files and OpenDocument `.ods` files, and `OdsWriter` writes `.ods`.

## Installation

```bash
# npm
npm install @xlsxflow/core

# pnpm
pnpm add @xlsxflow/core

# yarn
yarn add @xlsxflow/core
```

## Quick Start

### Parsing an Excel File (Streaming)

```typescript
import { SheetReader, createBlobReader } from '@xlsxflow/core';

// Browser / Edge: any Blob or File (e.g. from <input type="file">)
const blob = await fetch('https://example.com/data.xlsx').then(r => r.blob());

const reader = new SheetReader();
// Omit sheetName to read the first tab
const rows = await reader.parse(createBlobReader(blob), { sheetName: 'Sheet1' });

for await (const row of rows) {
  console.log(row.rowNumber, row.cells); // 1 ['ID', 'Name', 'Amount']
}

const meta = await rows.getMetadata(); // merged cells, hidden rows/cols, freeze panes
```

Dates come back as ISO-8601 strings. Opt in to more detail, each indexed like `row.cells`:

- `{ formulas: true }` gives `row.formulas`, with shared formulas expanded per cell.
- `{ styles: true }` gives `row.styles`, as `CellStyle` objects (the same shape the writer takes). Theme and palette colours are resolved to ARGB.
- `{ richText: true }` gives `row.richText`, the formatted runs of cells that have them. `row.cells` still holds the plain text.
- `{ formatted: true }` gives `row.formatted`, each cell's text as Excel (en-US) shows it: `1,234.50`, `25.6%`, `(42)`, `08-Oct-2026 2:05 PM`. It covers sections, conditions, dates and elapsed times, fractions, scientific notation, currency and text formats. Repeat fills (`*`) and colours are left out, and other locales are shown as en-US.

`await reader.readWorkbook(createBlobReader(blob))` lists the sheets with their visibility, the defined names and the document properties, without reading any sheet.

Hyperlinks are in `(await rows.getMetadata()).hyperlinks`, as `{ ref, hyperlink, tooltip? }` in the writer's format. `await rows.getImages()` returns the sheet's pictures in the writer's `images` format too (bytes included, read on that call), so they can be written back unchanged. Charts and shapes are skipped. `await rows.getComments()` returns the sheet's notes as `{ ref, text, author? }`.

Parts held in memory (workbook, shared strings, styles) are capped at 1 GiB uncompressed each, to stop zip bombs. The streamed worksheet is uncapped. Change both with `maxUncompressedBytes` (`Infinity` disables); `SheetEditor.edit(reader, { maxUncompressedBytes })` takes the same limit.

**Reading files from untrusted users** (uploads on a server): set `maxUncompressedBytes` to what you expect, such as `50_000_000`. Corrupt or crafted ZIPs (overlapping entries, entries that inflate past their stated size, directories pointing outside the file) are rejected either way. The limit also bounds the empty cells added before far-right cells, since one tiny cell in column XFD pads its row to 16,384 values. Values are returned as written: hyperlinks may be `javascript:` URLs and text may start with `=`, so check them before putting them in a web page or a CSV that a spreadsheet will open.

`sheetToJson(rows, headerRowIndex = 0)` turns the rows into objects keyed by the header row, and `streamToCsv(rows)` returns the sheet as CSV text. Both hold the whole result in memory.

In Node.js, read straight from disk with `XlsxFlow.readFile`, or pass `await createFileReader(path)` to any function that takes a reader:

```typescript
import { XlsxFlow } from '@xlsxflow/core';

for await (const row of await XlsxFlow.readFile('./data.xlsx')) {
  console.log(row.cells);
}
```

### Writing an Excel File

```typescript
import { SheetWriter } from '@xlsxflow/core';

const writer = new SheetWriter();

writer.addSheet('Report', [
  ['Header 1', 'Header 2', 'Header 3'],
  [1, 2, 3],
  ['Data', 'More Data', 'Even More Data'],
]);

// write() returns a ReadableStream<Uint8Array> of the .xlsx file
const blob = await new Response(writer.write()).blob();
```

Rows can also be an `AsyncIterable<Row>`, so millions of rows can be generated lazily; the writer only pulls rows as fast as the output is consumed.

Document properties and named ranges go to the constructor; visibility and view settings go to each sheet:

```typescript
const writer = new SheetWriter({
  properties: { title: 'Q3 sales', creator: 'Finance', company: 'ACME' }, // File > Info in Excel
  definedNames: [
    { name: 'TaxRate', ref: '0.18' },
    { name: 'Sales', ref: 'Data!$B$2:$B$100' },
    { name: 'Total', ref: 'Data!$B$101', sheet: 'Data' }, // scoped to one sheet
  ],
});
writer.addSheet('Lookup', lookupRows, { state: 'hidden' }); // or 'veryHidden'
writer.addSheet('Data', rows, { view: { zoom: 90, showGridLines: false, rightToLeft: false } });
```

Excel opens on the first visible sheet. Invalid or duplicate names, and a workbook with no visible sheet, are rejected before anything is written.

### Converting CSV

```typescript
import { SheetWriter, parseCsv } from '@xlsxflow/core';

// parseCsv streams rows from a string or a ReadableStream<Uint8Array>, e.g. file.stream()
const xlsx = new SheetWriter().addSheet('Data', parseCsv(csvStream)).write();
```

Quoted fields can hold delimiters, line breaks and `""`. Unquoted numbers and `TRUE`/`FALSE` become numbers and booleans, and empty fields become empty cells; `{ convert: false }` keeps everything as text, and `{ delimiter: ';' }` sets the delimiter. Dates stay text, since CSV files do not say which date order they use.

### .xls and .ods Files

`SheetReader` tells the format from the file's contents, so `parse` and `readWorkbook` work the same on `.xls`
(Excel 97-2003) and `.ods` (LibreOffice, Google Sheets, Excel's OpenDocument export) as on `.xlsx`:
values, dates, formula results, merged cells, hidden rows, columns and sheets, frozen panes, defined names
and document properties. `formatted: true` works on both; `formulas: true` works on `.ods`.

- `.xls` files are read whole into memory (capped by `maxUncompressedBytes`, 1 GiB by default). Formula text,
  styles, hyperlinks and notes are not read from them, and files older than Excel 97 (BIFF5 and earlier) are
  not supported.
- `.ods` files stream like `.xlsx`. Hyperlinks and notes are read too; `formatted` returns the text the file
  stores for each cell.
- A file saved with a password is rejected with an error that points to `decryptWorkbook` in
  [`@xlsxflow/pro`](#free-and-pro).

`OdsWriter` writes `.ods` with the same rows as `SheetWriter`: values, dates, formulas (converted to
OpenFormula), merged cells, column widths, frozen panes, hidden sheets and document properties. Cell styles,
hyperlinks, notes and images are not written to `.ods`.

```typescript
import { OdsWriter } from '@xlsxflow/core';

const ods = new OdsWriter({ properties: { title: 'Report' } })
  .addSheet('Data', rows, { columnWidths: [20, 10], freezePanes: { row: 1 } })
  .write(); // ReadableStream<Uint8Array>
```

### Editing an Existing File

```typescript
import { SheetEditor, createBlobReader } from '@xlsxflow/core';

const editor = new SheetEditor();
editor.setCells('Sheet1', {
  B2: 42, C2: { formula: 'B2*2' }, D9: 'new cell', A3: null,       // null clears
  A1: { style: { font: { bold: true }, fill: { type: 'solid', fgColor: 'FFFFFF00' } } }, // restyle, keep content
  B3: { value: 7, style: { numFmt: '0.00' } },
});
editor.appendSheet('Sheet1', [['new', 'row']]); // appended after the last existing row
editor.insertRows('Sheet1', 5, 3);  // 3 empty rows before row 5
editor.deleteRows('Sheet1', 20, 2); // rows 20-21
editor.insertColumns('Sheet1', 'C');  // or deleteColumns('Sheet1', 'C', 2)
editor.addSheet('Summary', [['Total', { value: null, formula: 'SUM(Sheet1!B:B)' }]]);
editor.deleteSheet('Old');
const edited = editor.edit(createBlobReader(existingBlob)); // ReadableStream<Uint8Array>
```

Edited cells keep their style. A style change is merged into the cell's current format: font properties, border sides and alignment settings you name change and the rest stay, while a fill or number format replaces the old one. The sheet streams through one row at a time, and every other part of the file is copied without being unpacked. Excel recalculates formulas when it opens the file. A shared formula whose first cell is overwritten is written out in full in the cells that used it.

Macro-enabled workbooks (`.xlsm`) keep their VBA project and content type through every edit.

`addSheet` takes an array of rows with values, formulas and styles; for hyperlinks, notes and sheet options, write the workbook with `SheetWriter`. `deleteSheet` removes names scoped to the sheet and turns other defined names that point at it into `#REF!`; formulas in other sheets that point at it are not rewritten. `insertRows`, `deleteRows`, `insertColumns` and `deleteColumns` move everything that points at the cells, as Excel does:
- formulas on every sheet and the workbook's defined names (print areas, named ranges);
- merged cells, conditional formats, validations, hyperlinks, the filter and its column filters, page breaks and column widths;
- tables, pictures, notes, sparklines, What-If data tables, chart series and pivot-table sources.

Ranges that span inserted rows or columns grow, and ranges over deleted ones shrink. References to deleted cells become `#REF!`. Columns inserted inside a table become table columns named Column1, Column2 and so on. Deleting a table's header row, all its data rows or all its columns is refused. Operations run in the order given, and `setCells` addresses count after the cells have moved. Inserted rows and columns are empty: they don't copy the formatting of their neighbours. Every sheet streams through the editor, because any of its formulas might point at the moved cells.

## Styles, Formulas & Conditional Formats

```typescript
import { SheetWriter } from '@xlsxflow/core';

const writer = new SheetWriter();
writer.addSheet('Sales', [
  [{ value: 'Total Revenue', style: { font: { bold: true }, fill: { type: 'solid', fgColor: 'FF1E3A5F' } } }],
  [100], [250],
  [{ value: null, formula: '=SUM(A2:A3)' }],          // cached result computed on write
  [new Date(), { value: new Date(), style: { numFmt: 'dd/mm/yyyy' } }], // UTC; default format yyyy-mm-dd[ hh:mm:ss]
  [{ value: 'Docs', hyperlink: 'https://example.com' }, { value: 'Back to top', hyperlink: '#Sales!A1' }],
  [{ value: null, richText: [{ text: 'Net ', font: { bold: true } }, { text: 'revenue', font: { color: 'FFC00000' } }] }],
], {
  freezePanes: { row: 1 },
  autoFilter: 'A1:B1',
  conditionalFormats: [{ range: 'A2:A3', rule: { type: 'dataBar', color: 'FF06B6D4' } }],
  images: [
    { data: logoPng, at: 'D1', height: 40 },               // PNG/JPEG/GIF bytes; width follows the aspect ratio
    { data: chartJpeg, range: 'D4:H14', altText: 'Trend' }, // stretched over the cells
  ],
});
```

Other sheet options:

```typescript
writer.addSheet('Report', rows, {
  conditionalFormats: [
    { range: 'B2:B100', rule: { type: 'cellIs', operator: 'greaterThan', formulae: [1000], style: { fill: { type: 'solid', fgColor: 'FFFFC7CE' } } } },
    { range: 'A2:A100', rule: { type: 'containsText', text: 'urgent', style: { font: { bold: true, color: 'FF9C0006' } } } },
    { range: 'C2:C100', rule: { type: 'iconSet', iconSet: '3TrafficLights1' } }, // also expression, top10, aboveAverage, duplicateValues...
  ],
  dataValidations: [{ sqref: 'B2:B100', type: 'whole', operator: 'between', formula1: '0', formula2: '100', error: 'Use 0-100' }],
  tables: [{ name: 'Sales', ref: 'A1:C100' }],  // header names come from row 1
  rows: { 1: { height: 24 }, 5: { outlineLevel: 1, hidden: true } },
  columns: [{ width: 30 }, { width: 12, outlineLevel: 1 }],
  protection: { password: 'secret', sort: true },  // Excel's legacy hash: deters edits, is not encryption
  pageSetup: { orientation: 'landscape', paperSize: 9, fitToWidth: 1, fitToHeight: 0, printArea: 'A1:C100', printTitleRows: '1', footer: '&CPage &P of &N' },
  tabColor: 'FF00B050',
  mergeCells: ['A1:C1'],
  columnWidths: [30, 12],  // in characters; `columns[i].width` wins where both are set
  autoFitColumns: true,    // widths from the longest value (array rows only)
});
// Notes: [{ value: 'Q3', comment: { text: 'Restated', author: 'Ana' } }]
// Formatted notes: comment: { text: [{ text: 'Ana:', font: { bold: true } }, { text: ' restated' }] }
```

Formulas are stored for Excel to calculate when it opens the file. For array rows, the writer also stores a cached result for simple formulas (`SUM`, `AVERAGE`, `COUNT`, `MIN`, `MAX`, `IF`, `CONCATENATE` and arithmetic), so other readers see a value.

Strings are written inline, which keeps memory flat. `new SheetWriter({ sharedStrings: true })` stores each distinct string once instead. Files are smaller when values repeat, but the distinct strings stay in memory until the file is finished.

## Compared with SheetJS and ExcelJS

Checked against each project's own documentation on 8 October 2026. "Pro" means a paid add-on.

| | XlsxFlow | SheetJS Community Edition | ExcelJS 4.4 |
|---|---|---|---|
| Streaming `.xlsx` read and write | Yes | No (streams CSV, HTML and JSON out) | Yes |
| Cell styles, read and write | Yes | No (SheetJS Pro) | Yes |
| Images | Yes | No | Yes |
| `.xls` | Read | Read and write | No |
| `.ods` | Read and write | Read and write | No |
| `.xlsb`, `.numbers` and other formats | No | Yes | No |
| Charts | Add (Pro) | No (SheetJS Pro) | No |
| Pivot tables | Add (Pro) | No (SheetJS Pro) | Partial, undocumented |
| Password-protected files | Open and save (Pro) | Old `.xls` obfuscation only (SheetJS Pro opens AES files) | No |
| Licence | MIT, Pro is paid | Apache 2.0 | MIT |

SheetJS reads and writes far more formats, and ExcelJS has a longer track record (its last release was in
October 2023). XlsxFlow focuses on `.xlsx`: streaming in flat memory, keeping everything in a file it edits,
and running on Web APIs alone.

## Benchmarks

Write benchmark: 10 numeric columns, at 100,000 rows (1M cells) and 1,000,000 rows (10M cells). Each library ran in its own process on Node v25.8.2 with a 4 GB heap limit, and "Heap" is the growth in heap usage. Times are from one run on a laptop with other apps open; runs on that machine varied by up to 2×, so treat differences under about 20% as a tie. Reproduce with `npx tsx scripts/benchmark-competitors.ts` (inside `packages/core`, after `pnpm build`). `BENCH_ROWS=1000000` runs 10M cells, `BENCH_LIBS=xlsxflow,exceljs` runs a subset, and a library still writing after `BENCH_TIMEOUT_MIN` minutes (default 10) is stopped. The 10M runs for SheetJS and excel4node used a 30-minute limit.

1M cells:

| Library | Write Time | File Size | Heap |
|---|---|---|---|
| **XlsxFlow** | **2,899 ms** | **2.9 MB** | **+2 MB** |
| ExcelJS 4.4 (streaming writer) | 3,397 ms | 3.0 MB | +9 MB |
| SheetJS 0.20.3, `compression: true` | 4,370 ms | 8.4 MB | +140 MB |
| SheetJS 0.20.3, default options | 5,379 ms | 31.4 MB | +140 MB |
| write-excel-file | 7,681 ms | 2.8 MB | +2 MB |
| xlsx-populate | 9,852 ms | 2.9 MB | +114 MB |
| excel4node | 14,973 ms | 3.1 MB | +205 MB |
| msexcel-builder | fails to run (`Invalid character in name: fileVersion`) | | |

10M cells:

| Library | Write Time | File Size | Heap |
|---|---|---|---|
| **XlsxFlow** | **19.5 s** | **29.9 MB** | **+1 MB** |
| ExcelJS 4.4 (streaming writer) | 22.0 s | 31.2 MB | +7 MB |
| write-excel-file | 68.1 s | 29.3 MB | +1 MB |
| xlsx-populate | 75.3 s | 30.0 MB | +1,118 MB |
| SheetJS 0.20.3 (with and without compression) | not finished after 30 min | | |
| excel4node | not finished after 30 min | | |

Rows are pulled from an async generator. The writer only generates rows as fast as the output stream is consumed, so memory stays flat as row count grows.

Read benchmark: a 100,000 × 10 file written by ExcelJS (shared strings, numbers, dates, booleans; 6.4 MB). Every library reads every cell. "Peak RSS" is the peak memory of the reading process. Median of three runs. Reproduce with `npx tsx scripts/benchmark-read-competitors.ts`.

| Library | Read Time | Peak RSS |
|---|---|---|
| **XlsxFlow** | **2,279 ms** | **96 MB** |
| ExcelJS 4.4 (streaming reader) | 2,669 ms | 263 MB |
| ExcelJS 4.4 | 3,680 ms | 667 MB |
| SheetJS 0.20.3 | 6,091 ms | 549 MB |

---

## Free and Pro

`@xlsxflow/core` is free and MIT licensed, including every feature on this page. [`@xlsxflow/pro`](https://www.npmjs.com/package/@xlsxflow/pro) is a paid add-on with a licence key:

| | Core (free) | Pro |
|---|---|---|
| Read, write and edit `.xlsx` / `.xlsm`, styles, formulas, images, tables, notes | Yes | Yes |
| Read `.xls`, read and write `.ods` | Yes | Yes |
| Fill Excel templates with data, repeating rows for lists | | Yes |
| Add column, bar, line, area and pie charts | | Yes |
| Add pivot tables | | Yes |
| Open and save password-protected `.xlsx` files | | Yes |

Pro is $5 per developer (local pricing at checkout), with a perpetual licence and a year of updates. See the [Pro README](https://www.npmjs.com/package/@xlsxflow/pro) for details.

## Changelog

See [CHANGELOG.md](https://github.com/xlsxflow/xlsxflow/blob/main/packages/core/CHANGELOG.md).

## Contributing and security

Bug reports and pull requests are welcome: see [CONTRIBUTING.md](https://github.com/xlsxflow/xlsxflow/blob/main/CONTRIBUTING.md). Report security issues privately as described in [SECURITY.md](https://github.com/xlsxflow/xlsxflow/blob/main/SECURITY.md).

## License

[MIT](https://github.com/xlsxflow/xlsxflow/blob/main/packages/core/LICENSE). Pro is a separate package under its own licence and does not change the terms of the core.
