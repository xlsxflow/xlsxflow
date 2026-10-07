<div align="center">
  <img src="https://raw.githubusercontent.com/xlsxflow/xlsxflow/main/assets/logo.png" alt="XlsxFlow" width="200" />
  <h1>XlsxFlow</h1>
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

**XlsxFlow** is a zero-dependency, ultra-fast streaming parser and writer for OpenXML (`.xlsx`) files. Built on native Web APIs (like `TransformStream` and Web Crypto), it handles millions of cells with a completely flat memory profile.

Unlike DOM-based AST parsers (like ExcelJS or SheetJS), XlsxFlow processes files chunk-by-chunk on the fly, making it perfect for Edge environments (Cloudflare Workers, Vercel Edge, Next.js Server Actions) and client-side browser usage without crashing the heap.

## ✨ Features

- **Zero Dependencies**: Pure modern TypeScript, leveraging native browser/Node Web APIs.
- **True Streaming**: Parse gigabytes of Excel data using `ReadableStream` with almost zero memory overhead.
- **Edge Native**: Fully compatible with Node.js, Deno, Bun, Cloudflare Workers, and modern browsers.
- **Read & Write**: Stream massive `.xlsx` files and generate them on the fly.
- **Styles & Formulas**: Fonts, fills, borders, alignment, number formats, data bars / color scales, cached formula results, dates, hyperlinks, autofilters and images. Read formulas and styles back. All MIT, all free.

## 📦 Installation

```bash
# npm
npm install @xlsxflow/core

# pnpm
pnpm add @xlsxflow/core

# yarn
yarn add @xlsxflow/core
```

## 🚀 Quick Start

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

Hyperlinks are in `(await rows.getMetadata()).hyperlinks`, as `{ ref, hyperlink, tooltip? }` in the writer's format. `await rows.getImages()` returns the sheet's pictures in the writer's `images` format too (bytes included, read on that call), so they can be written back unchanged. Charts and shapes are skipped.

Parts held in memory (workbook, shared strings, styles) are capped at 1 GiB uncompressed each, to stop zip bombs. The streamed worksheet is uncapped. Change both with `maxUncompressedBytes` (`Infinity` disables).

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

### Appending to an Existing File

```typescript
import { SheetEditor, createBlobReader } from '@xlsxflow/core';

const editor = new SheetEditor();
editor.appendSheet('Sheet1', [['new', 'row']]); // appended after the last existing row
const edited = editor.edit(createBlobReader(existingBlob)); // ReadableStream<Uint8Array>
```

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

Strings are written inline, which keeps memory flat. `new SheetWriter({ sharedStrings: true })` stores each distinct string once instead. Files are smaller when values repeat, but the distinct strings stay in memory until the file is finished.

## 📊 Benchmarks

Write benchmark: 100,000 rows × 10 numeric columns (1M cells). Each library ran in its own process on Node v25.8.2, and "Heap" is the growth in heap usage. Times are from one warm run; repeated runs of the top three varied by under 10%. Reproduce with `npx tsx scripts/benchmark-competitors.ts` (inside `packages/core`, after `pnpm build`; `BENCH_LIBS=xlsxflow,exceljs` runs a subset).

| Library | Write Time | File Size | Heap |
|---|---|---|---|
| **XlsxFlow** | **1,685 ms** | **2.9 MB** | **+1 MB** |
| ExcelJS 4.4 (streaming writer) | 1,797 ms | 3.0 MB | +9 MB |
| SheetJS (`xlsx` 0.18.5 from npm) | 3,228 ms | 31.4 MB | +170 MB |
| write-excel-file | 4,313 ms | 2.8 MB | +2 MB |
| xlsx-populate | 7,500 ms | 2.9 MB | +114 MB |
| excel4node | 12,087 ms | 3.1 MB | +205 MB |
| msexcel-builder | fails to run (`Invalid character in name: fileVersion`) | | |

Rows are pulled from an async generator. The writer only generates rows as fast as the output stream is consumed, so memory stays flat as row count grows.

Read benchmark: a 100,000 × 10 file written by ExcelJS (shared strings, numbers, dates, booleans; 6.4 MB). Every library reads every cell. "Peak RSS" is the peak memory of the reading process. Median of three runs. Reproduce with `npx tsx scripts/benchmark-read-competitors.ts`.

| Library | Read Time | Peak RSS |
|---|---|---|
| **XlsxFlow** | **2,415 ms** | **92 MB** |
| ExcelJS (streaming reader) | 3,010 ms | 267 MB |
| ExcelJS | 4,227 ms | 660 MB |
| SheetJS (`xlsx`) | 6,035 ms | 544 MB |

---

## 📋 Changelog

### Unreleased

- **Corrected benchmarks:** the earlier ExcelJS write time (13.5 s) came from a cold first run; warm, it is 1.8 s. write-excel-file was listed as running out of memory, but the harness used its old API and never ran it. Failures are now reported as failures, not out-of-memory.
- **Renamed:** SheetForge is now XlsxFlow. Packages are `@xlsxflow/core` and `@xlsxflow/pro`, and `SheetForge.readFile` is `XlsxFlow.readFile`.
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

- **O(1) Memory Streaming Writer & Editor** — Ripped out in-memory buffers. `SheetWriter` and `SheetEditor` now use `ZipStreamWriter` via Data Descriptors (Bit 3) to generate dynamic ZIP archives entirely on-the-fly, reducing memory overhead to O(1) flat.
- **Dynamic Date Deserialization** — Robust detection of `numFmtId` across workbooks to reliably auto-convert numeric epoch dates back into strict ISO-8601 strings during stream parsing.
- **Data Descriptors & Signature Scanning** — Fixed limitations with forward-only zip stream parsers by scanning for Data Descriptor headers `0x08074b50`, achieving zero-seek streaming parsing of workbooks.
- **Production Ready** — Validated by extensive tests and rigorous benchmarking.

---

### v0.3.0-beta
> Multi-Sheet Support, Auto Date Deserialization, & Massive XML Parsing Optimization

- **Multi-Sheet Writing** — You can now use `writer.addSheet()` multiple times to chain worksheets into a single exported `.xlsx` workbook.
- **Dynamic XML Structuring** — The zip packer dynamically adjusts `[Content_Types].xml`, `workbook.xml`, and relationships files.
- **Auto Date Deserialization** — `SheetReader` now pre-fetches `styles.xml` from the stream, parses `<cellXfs>` and `<numFmts>`, and heuristically identifies cells with date formats. Numeric Excel epoch dates are automatically mapped directly to strict `ISO-8601` strings!
- **Exponential XML Stream Bug Fixed** — Found and eliminated an $O(N^2)$ buffer accumulation bug in `xml-stream.ts`. Reading 1 Million cells now parses fully in under ~4 seconds (down from ~6.4s) while consuming <60MB of peak heap overhead.
- **Portal App Update** — The interactive `/apps/portal` demo now dynamically exports workbooks containing 2 distinct sheets and verified Date cells.

---

### v0.2.0-beta
> Native Deflate Compression & Rich Text Support

- **SheetWriter is now async** — `write()` returns `Promise<Uint8Array>`
- **Native Deflate compression** via `CompressionStream('deflate-raw')` — no dependencies
- **ZIP binary upgraded** — compression method `0x08`, correct uncompressed size & CRC-32 in headers
- **File size reduction**: 1M cell file went from **30 MB → 2.9 MB** (90% smaller)
- **Read speed improved**: parse time dropped from **~9.4s → 6.4s** (less I/O from smaller file)
- **Rich Text / Inline String support** — `SheetReader` now parses `<t>` inside `<is>` and `<r>` elements
- **Type fixes**: `@types/node` added, `TextDecoderStream` cast resolved
- **ZIP backpressure deadlock** permanently fixed via concurrent background pump

#### v0.2.0-beta vs v0.1.0-beta comparison

| Metric | v0.1.0-beta | v0.2.0-beta | Delta |
|---|---|---|---|
| File size (1M cells) | 30 MB | 2.9 MB | **−90%** |
| Write time | ~3,600 ms | ~3,900 ms | ~+8% (compression overhead) |
| Read time | ~9,400 ms | ~6,400 ms | **−32%** (less disk I/O) |
| ZIP Compression | Store (none) | Deflate (native) | ✅ |
| Async write API | ❌ sync | ✅ async | ✅ |
| Rich text cells | ❌ | ✅ | ✅ |

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

## 📄 License

`@xlsxflow/core` is licensed under the [MIT License](packages/core/LICENSE). A commercial `@xlsxflow/pro` add-on is planned; it will be a separate package under its own license and never changes the terms of the core.

---
<div align="center">
  Built with 💻 and ☕ for modern web developers.
</div>
