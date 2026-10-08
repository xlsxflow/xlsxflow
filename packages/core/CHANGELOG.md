# Changelog

## 1.1.3

### Fixed

- `SheetEditor.setCells` refuses addresses past column XFD or row 1,048,576, instead of writing a file Excel reports as damaged.
- `SheetWriter` and `SheetEditor` refuse text longer than Excel's 32,767 characters per cell.
- Colours are checked: ARGB (`FFFF0000`) and RGB (`FF0000`, `#FF0000`) hex are accepted, anything else (`'red'`) throws instead of making Excel repair the file.
- Picture anchors past the last column or row throw; a range given the wrong way round (`F20:A2`) is turned round.
- `SheetReader`: a shared string index past the end of the table reads as an empty cell, not as the index.

### Added

- `checkCellText` is exported, for `@xlsxflow/pro`.

## 1.1.2

### Fixed

- `SheetWriter` throws an error for a row longer than 16,384 cells or past row 1,048,576, instead of writing a file Excel reports as damaged.
- `SheetEditor.setCells`: a date in a cell without a date format gets `yyyy-mm-dd` (or `yyyy-mm-dd hh:mm:ss` with a time), as in `SheetWriter`, instead of showing as a serial number.
- `OdsWriter` stores formula results, so readers that don't recalculate show them.
- `sheetToJson` keeps every column when headers repeat: the second `Name` becomes `Name_2`.
- `parseCsv` returns `null` for empty unquoted fields with `{ convert: false }` too.

### Added

- `StylePatcher` is exported, for `@xlsxflow/pro`.

## 1.1.1

- README: the XlsxFlow website address, and Pro at a flat $5.
- More npm keywords: `xls` and `ods`.

## 1.1.0

### Added

- `SheetWriter`: workbook properties (title, author, company and so on), defined names, hidden and very hidden sheets, and sheet views (zoom, gridlines, headings, right-to-left). `SheetReader.readWorkbook` reads them back.
- `{ formatted: true }` reports each cell's text as Excel shows it, from its number format.
- `SheetEditor.insertRows`, `deleteRows`, `insertColumns` and `deleteColumns`. Everything that points at the moved cells moves with them: formulas on all sheets, defined names, merges, conditional formats, validations, hyperlinks, filters, page breaks, column widths, tables, pictures, notes, sparklines, data tables, chart series and pivot sources.
- `SheetEditor.setCells` edits and restyles cells of existing files; `addSheet` and `deleteSheet` add and remove sheets.
- `parseCsv` streams CSV rows, which `SheetWriter.addSheet` turns into a sheet.
- `SheetReader` reads Excel 97-2003 `.xls` files and OpenDocument `.ods` files: values, dates, formula results, merges, hidden rows, columns and sheets, frozen panes, defined names and document properties. `.ods` also gives formulas, hyperlinks and notes.
- `OdsWriter` writes `.ods` files: values, dates, formulas, merges, column widths, frozen panes and hidden sheets.
- Password-protected `.xlsx` files are rejected with an error that points to `decryptWorkbook` in `@xlsxflow/pro`, instead of a ZIP error.
- Cell notes on write (plain or formatted text) and `getComments()` on read.
- Conditional formats `cellIs`, `expression`, `top10`, `aboveAverage`, text rules, `duplicateValues`/`uniqueValues` and `iconSet`.
- Excel tables, sheet protection, page setup (margins, header and footer, print area and titles), row heights, hidden rows and columns, outline grouping, tab colour, and validation operators and messages.
- `Date` cell values, hyperlinks (URLs and locations in the workbook), `autoFilter`, and an optional shared string table on write.
- Reading returns formulas (`formulas: true`, shared formulas expanded), styles (`styles: true`) and rich text (`richText: true`). Theme and indexed colours are resolved to ARGB.
- PNG, JPEG and GIF images (`images` sheet option), anchored to a cell or stretched over a range. `getImages()` reads them back.
- Low-level parts are exported for add-ons: the ZIP reader and writer, `resolveWorkbookParts`, `readSharedStrings`, and `mapFormulaRefs`/`shiftFormula`.
- Tested: `.xlsm` files keep their macros through `SheetEditor`; the library runs in Cloudflare Workers without `nodejs_compat` (`scripts/workers` checks it).

### Security

Fixed crashes, hangs and memory blowups caused by crafted files:

- A ZIP directory pointing outside the file crashed Node.
- Unclosed tags and elements made the parser quadratic. The XML tokenizer is now linear on giant tags, text nodes and CDATA.
- Repeated hidden columns, far-right cells, fraction formats with long denominators and large style tables were slow or used unbounded memory.
- One deflated entry could be read many times under different names.

Also:

- `SheetEditor.edit` takes `maxUncompressedBytes`.
- Option values typed as enums are escaped, and appended rows encode control characters.
- `sheetToJson` keeps a `__proto__` header as an ordinary key.
- In-memory parts are capped at 1 GiB uncompressed by default.

### Fixed

- Sheets with frozen panes were all marked as selected, so Excel opened them grouped.
- The reader dropped or corrupted cells at stream chunk boundaries.
- Styles pointed at the wrong font, fill or border, and styles used only by `AsyncIterable` rows were missing from `styles.xml`.
- Number formats with quoted text, escapes or colours (`#,##0.00 "USD"`, `[Red]0.0`) were taken for date formats.
- Reading a corrupt file without calling `getMetadata()` caused an unhandled promise rejection, which is fatal in Node.
- `addSheet` accepted sheet names Excel refuses to open. It now throws.
- Files Excel would repair:
  - gradient stops were written as `<gradientStop>`;
  - `vertical: 'middle'` was written as is instead of `center`;
  - conditional formats came after validations.
- Data bar `minValue`/`maxValue` were ignored, every conditional format had priority 1, and colours and ranges were not XML-escaped.
- Strings with `_xHHHH_`, control characters, CR, or leading and trailing spaces now survive a round trip. `NaN` and `Infinity` are written as `#NUM!`.
- Reader:
  - Strict OOXML and non-standard part names work;
  - empty `<v/>` reads as empty and `-0` as `0`;
  - out-of-order cells land in the right column;
  - time-only values are no longer a day off, and datetimes keep milliseconds.
  - frozen panes saved by Excel as `frozenSplit` (frozen after a split) are reported.
- `sheetToJson` ignored its `headerRowIndex` argument.

### Changed

- The package ships compiled ESM and CommonJS builds with type declarations.
- Browser bundles no longer try to resolve Node's `fs`.
- `[Content_Types].xml` is written last in the ZIP.
- Real backpressure in the ZIP writer and worksheet stream. Producer errors now error the output stream instead of hanging it.
- About 2.5× faster reads, and a table-driven CRC-32 for faster writes.
- Styles, formulas and conditional formats are part of the MIT core. Paid add-ons are in the separate `@xlsxflow/pro` package.
- Benchmarks corrected: the earlier ExcelJS write time came from a cold first run, and write-excel-file was run with an old API.

## 1.0.0

- `SheetWriter` and `SheetEditor` stream the ZIP with data descriptors instead of buffering the file in memory.
- Dates are returned as ISO-8601 strings, detected from each cell's number format.

## 0.3.0-beta

- Workbooks with several sheets.
- Fixed a buffer bug that made parsing quadratic.

## 0.2.0-beta

- Deflate compression with `CompressionStream`: a 1M-cell file went from 30 MB to 2.9 MB.
- Inline and rich text strings are read.

## 0.1.0-beta

- First release: streaming reader, writer, and shared strings.
