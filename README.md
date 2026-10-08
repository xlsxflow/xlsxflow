<div align="center">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://raw.githubusercontent.com/xlsxflow/xlsxflow/main/assets/logo-wordmark-dark.svg" />
    <img src="https://raw.githubusercontent.com/xlsxflow/xlsxflow/main/assets/logo-wordmark.svg" alt="XlsxFlow" width="320" />
  </picture>
  <p><strong>The modern, streaming Excel engine for the web.</strong></p>
  
  [![npm version](https://img.shields.io/npm/v/@xlsxflow/core.svg?style=flat-square)](https://www.npmjs.com/package/@xlsxflow/core)
  [![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square)](https://opensource.org/licenses/MIT)
  [![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?style=flat-square)](http://makeapullrequest.com)

  <p>
    <a href="#features">Features</a> •
    <a href="#installation">Installation</a> •
    <a href="#quick-start">Quick Start</a> •
    <a href="#styles-formulas--conditional-formats">Styles &amp; Formulas</a> •
    <a href="#benchmarks">Benchmarks</a>
  </p>
</div>

---

**XlsxFlow** is a zero-dependency streaming reader, writer and editor for OpenXML (`.xlsx`) files. Built on native Web APIs (like `TransformStream` and `CompressionStream`), it handles millions of cells in flat memory.

Unlike DOM-based AST parsers (like ExcelJS or SheetJS), XlsxFlow processes files chunk-by-chunk on the fly, so it suits browsers, servers and edge runtimes that provide the same Web APIs.

## Features

- **No dependencies**: TypeScript on Web APIs (`ReadableStream`, `CompressionStream`, `Blob`).
- **Streaming**: rows are read and written one at a time, so memory stays flat as files grow (10M cells written with about 1 MB of extra heap; see [Benchmarks](#benchmarks)).
- **Runs anywhere with Web APIs**: tested on Node 20.12+, Bun, browsers and Cloudflare Workers (without `nodejs_compat`). Deno provides the same APIs but is not tested yet.
- **Read, write and edit**: stream rows out of a file, generate one on the fly, or change cells, rows, columns and sheets of an existing file while keeping everything else in it.
- **Styles and formulas**: fonts, fills, borders, alignment, number formats, conditional formats, validations, tables, notes, hyperlinks, autofilters, images, protection and page setup. Formulas and styles read back too.

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
const blob = await fetch('https://example.com/massive-data.xlsx').then(r => r.blob());

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

In Node.js, read straight from disk:

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

`addSheet` takes values, formulas and styles; for hyperlinks, notes and sheet options, write the workbook with `SheetWriter`. `deleteSheet` removes names scoped to the sheet and turns other defined names that point at it into `#REF!`; formulas in other sheets that point at it are not rewritten. `insertRows`, `deleteRows`, `insertColumns` and `deleteColumns` move everything that points at the cells, as Excel does:
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

More sheet options, matching what ExcelJS offers:

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
});
// Notes: [{ value: 'Q3', comment: { text: 'Restated', author: 'Ana' } }]
// Formatted notes: comment: { text: [{ text: 'Ana:', font: { bold: true } }, { text: ' restated' }] }
```

Strings are written inline, which keeps memory flat. `new SheetWriter({ sharedStrings: true })` stores each distinct string once instead. Files are smaller when values repeat, but the distinct strings stay in memory until the file is finished.

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

## Changelog

### v1.1.0

- **New:** workbook properties (title, author, company...), defined names, hidden and very hidden sheets, and sheet views (zoom, gridlines, headings, right-to-left) in `SheetWriter`; `SheetReader.readWorkbook` reads them back.
- **Fixed:** sheets with frozen panes were all marked as selected, so Excel opened them grouped.
- **New:** `{ formatted: true }` reports each cell's text as Excel shows it, from its number format.
- **Tested:** `.xlsm` files keep their macros through `SheetEditor`.
- **Tested:** runs in Cloudflare Workers without `nodejs_compat`; `scripts/workers` checks it.
- **Security:** fixed crashes, hangs and memory blowups from crafted files: a ZIP directory pointing outside the file crashed Node; unclosed tags and elements made the parser quadratic; repeated hidden columns, far-right cells, fraction formats with long denominators and large style tables were slow or memory-hungry; one deflated entry could be read many times under different names. `SheetEditor.edit` takes `maxUncompressedBytes`. Option values typed as enums are now escaped, appended rows encode control characters, and `sheetToJson` keeps a `__proto__` header as an ordinary key.
- **Corrected benchmarks:** the earlier ExcelJS write time (13.5 s) came from a cold first run; warm, it is 1.8 s. write-excel-file was listed as running out of memory, but the harness used its old API and never ran it. Failures are now reported as failures, not out-of-memory.
- **New:** `SheetEditor.insertRows`, `deleteRows`, `insertColumns` and `deleteColumns` move cells in existing files, with every reference to them: formulas on all sheets, defined names, merges, conditional formats, validations, hyperlinks, filters, page breaks, column widths, tables, pictures, notes, sparklines, data tables, chart series and pivot sources.
- **New:** `SheetEditor` restyles existing cells (`setCells` with `style`), adds sheets (`addSheet`) and deletes them (`deleteSheet`); `parseCsv` streams CSV rows, which `SheetWriter.addSheet` turns into xlsx; notes take formatted text runs.
- **New, closing ExcelJS gaps:** `SheetEditor.setCells` edits cells of existing files; cell notes on write and `getComments()` on read; conditional formats `cellIs`, `expression`, `top10`, `aboveAverage`, text rules, `duplicateValues`/`uniqueValues` and `iconSet`; Excel tables; sheet protection; page setup, margins, header/footer, print area and titles; row heights, hidden rows/columns and outline grouping; tab colour; validation operators and messages.
- **Changed:** `[Content_Types].xml` is now written last in the ZIP, since streamed sheets decide which parts exist. Readers use the ZIP's central directory, so entry order does not matter.
- **Fixed: reader dropped/corrupted cells at stream chunk boundaries** (the XML tokenizer discarded buffered characters between chunks). Large files from ExcelJS/SheetJS now read back exactly.
- **Fixed: styles pointed at the wrong font/fill/border** (off-by-one against the default entries), and styles used by `AsyncIterable` rows were missing from `styles.xml`.
- **Real backpressure** in the ZIP writer and worksheet stream; producer errors now error the output stream instead of hanging it.
- **Faster writes** via table-driven CRC-32 (1M numeric cells: ~1.6 s, previously ~32 s).
- Reader resolves sheets from `workbook.xml` (first tab by default, absolute targets, any attribute order).
- `SheetEditor`: handles empty `<sheetData/>`, copies untouched entries without recompressing, errors on unknown sheet names; now exported.
- Formula cached values keep their type; aggregates ignore text/blanks like Excel; formulas evaluate against their own sheet.
- Styles, formulas and conditional formatting are part of the MIT core (moved out of `src/pro`). Licensing code moved to a separate, unpublished `@xlsxflow/pro` package.
- npm package now ships compiled ESM + CJS builds with bundled type declarations instead of TypeScript source.
- Browser bundles no longer try to resolve Node `fs`.
- **Fixed: numbers shown as dates.** Number formats with quoted text, escapes or colours (`#,##0.00 "USD"`, `0 "days"`, `[Red]0.0`) were detected as date formats, so their values were returned as dates.
- **Fixed: unhandled promise rejection** (fatal in Node) when reading a corrupt file without calling `getMetadata()`. Corrupt ZIP entries now fail with an error that names the entry.
- **Hostile-input hardening:** the XML tokenizer is linear-time on giant tags, text nodes and CDATA sections (was quadratic). Cell references beyond column XFD and oversized hidden-column ranges are rejected or clamped instead of allocating without bound.
- **~2.5x faster reads:** the tokenizer emits one batch of tokens per chunk instead of one stream chunk per token.
- Reader correctness:
  - Workbook, sheets, shared strings and styles are located via package relationships, so Strict OOXML and non-standard part names work.
  - `_xHHHH_` escapes are decoded and XML line endings are normalized.
  - Empty `<v/>` reads as empty, and `-0` as `0`.
  - Out-of-order cells land in the right column.
  - Time-only values are no longer a day off, and datetimes keep millisecond precision.
- **New:** `Date` cell values, hyperlinks (URLs and in-workbook locations), `autoFilter`, and an opt-in shared string table on write. Reading can return formulas (`formulas: true`, shared formulas expanded) and styles (`styles: true`).
- **New:** rich text runs on write (`richText`) and read (`richText: true`); hyperlinks read back via `getMetadata()`; theme and indexed colours (with tints, and the workbook's own palette) resolved to ARGB when reading styles.
- **New:** embedded PNG/JPEG/GIF images (`images` sheet option), anchored to a cell at their own size or stretched over a range. Format and size come from the file header; data shared between sheets is stored once. `getImages()` reads them back.
- **Fixed:** `addSheet` accepted sheet names Excel refuses to open (over 31 characters, `\ / ? * : [ ]`, a leading or trailing apostrophe, or a duplicate name ignoring case). It now throws.
- Low-level building blocks are exported for add-ons: the ZIP reader and writer, `resolveWorkbookParts`, `readSharedStrings`, and `mapFormulaRefs`/`shiftFormula` (A1 reference rewriting that understands sheet qualifiers).
- The formula engine understands `$A$1` references and no longer logs to the console for formulas it cannot evaluate.
- **New:** 1 GiB default size cap on in-memory parts when reading (zip-bomb guard).
- **Fixed: files Excel would repair:** gradient fills were written as `<gradientStop>` (the element is `<stop>`), `vertical: 'middle'` was written verbatim (OOXML says `center`), and conditional formatting came after data validation (schema order is the other way round).
- Fixed: `minValue`/`maxValue` on data bars were ignored; every conditional format had priority 1; colours and ranges were not XML-escaped.
- Fixed: strings containing `_xHHHH_`, control characters or CR, or leading/trailing spaces, now survive a write/read round trip. `NaN`/`Infinity` are written as `#NUM!` instead of an invalid cell.
- Test corpus: fixtures from openpyxl, ExcelJS, SheetJS, xlsx-populate, XlsxFlow, plus hand-crafted edge cases, checked against an openpyxl oracle. Also a fuzz suite for corrupted ZIPs and hostile XML.

### v1.0.0 (Official Release)
> True O(1) Streaming Architecture for Writers & Editors

- **O(1) Memory Streaming Writer & Editor**: Removed the in-memory buffers. `SheetWriter` and `SheetEditor` now use `ZipStreamWriter` via Data Descriptors (Bit 3) to generate dynamic ZIP archives entirely on-the-fly, so memory stays flat.
- **Dynamic Date Deserialization**: Detection of `numFmtId` across workbooks to convert numeric epoch dates back into strict ISO-8601 strings during stream parsing.
- **Data Descriptors & Signature Scanning**: Fixed limitations with forward-only zip stream parsers by scanning for Data Descriptor headers `0x08074b50`, so workbooks parse without seeking.

---

### v0.3.0-beta
> Multi-Sheet Support, Auto Date Deserialization, & Massive XML Parsing Optimization

- **Multi-Sheet Writing**: You can now use `writer.addSheet()` multiple times to chain worksheets into a single exported `.xlsx` workbook.
- **Dynamic XML Structuring**: The zip packer dynamically adjusts `[Content_Types].xml`, `workbook.xml`, and relationships files.
- **Auto Date Deserialization**: `SheetReader` now pre-fetches `styles.xml` from the stream, parses `<cellXfs>` and `<numFmts>`, and heuristically identifies cells with date formats. Numeric Excel dates are returned as `ISO-8601` strings.
- **Quadratic XML stream bug fixed**: Fixed a buffer accumulation bug that made parsing O(n²) in `xml-stream.ts`. Reading 1 Million cells now parses fully in under ~4 seconds (down from ~6.4s) while consuming <60MB of peak heap overhead.
- **Portal App Update**: The interactive `/apps/portal` demo now dynamically exports workbooks containing 2 distinct sheets and verified Date cells.

---

### v0.2.0-beta
> Native Deflate Compression & Rich Text Support

- **SheetWriter is now async**: `write()` returns `Promise<Uint8Array>`
- **Native Deflate compression** via `CompressionStream('deflate-raw')`, no dependencies
- **ZIP binary upgraded**: compression method `0x08`, correct uncompressed size & CRC-32 in headers
- **File size reduction**: 1M cell file went from **30 MB → 2.9 MB** (90% smaller)
- **Read speed improved**: parse time dropped from **~9.4s → 6.4s** (less I/O from smaller file)
- **Rich Text / Inline String support**: `SheetReader` now parses `<t>` inside `<is>` and `<r>` elements
- **Type fixes**: `@types/node` added, `TextDecoderStream` cast resolved
- **ZIP backpressure deadlock** fixed with a background pump

#### v0.2.0-beta vs v0.1.0-beta comparison

| Metric | v0.1.0-beta | v0.2.0-beta | Delta |
|---|---|---|---|
| File size (1M cells) | 30 MB | 2.9 MB | **−90%** |
| Write time | ~3,600 ms | ~3,900 ms | ~+8% (compression overhead) |
| Read time | ~9,400 ms | ~6,400 ms | **−32%** (less disk I/O) |
| ZIP compression | Store (none) | Deflate (native) | |
| Write API | sync | async | |
| Rich text cells | no | yes | |

---

### v0.1.0-beta
> Initial Release

- Streaming SAX-style XLSX parser (`SheetReader`)
- Zero-dependency XLSX writer (`SheetWriter`) with Store compression
- ZIP stream parser built on native `TransformStream`
- Shared String Table (`xl/sharedStrings.xml`) support
- Pro tier architecture: `StyleEngine`, `FormulaEngine`, `ConditionalFormatter`
- Dead-Drop license system (Zero-DB, hardware-bound, offline)
- Next.js Portal (`apps/portal`) with interactive playground

---

## License

`@xlsxflow/core` is licensed under the [MIT License](https://github.com/xlsxflow/xlsxflow/blob/main/packages/core/LICENSE). A commercial `@xlsxflow/pro` add-on is planned; it will be a separate package under its own license and never changes the terms of the core.

