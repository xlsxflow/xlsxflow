// Reader for OpenDocument spreadsheets (.ods, as saved by LibreOffice, Google Sheets and Excel).
// content.xml streams through the same tokenizer as .xlsx sheets, so memory stays flat.

import { createXmlBatchParser, type XmlToken } from './xml-stream';
import { ParseResult, excelToIsoDate, type CellValue, type RowData, type SheetComment, type SheetMetadata } from './worksheet-parser';
import { colLetter, unescapeXml, xmlElements } from './utils';
import type { DefinedName, WorkbookProperties } from './types';

export const ODS_MIMETYPE = 'application/vnd.oasis.opendocument.spreadsheet';

const MAX_ROWS = 1 << 20;
const MAX_COLUMNS = 1 << 14;

type TokenSource = () => Promise<ReadableStream<Uint8Array>>;

async function* tokens(open: TokenSource): AsyncGenerator<XmlToken> {
  const reader = (await open()).pipeThrough(createXmlBatchParser()).getReader();
  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) return;
      yield* value;
    }
  } finally {
    reader.releaseLock();
  }
}

const repeat = (value: string | undefined, max: number) => Math.min(Math.max(parseInt(value ?? '1', 10) || 1, 1), max);

// "2026-10-08" or "2026-10-08T14:04:30" (local time in the file) as the ISO string the .xlsx reader returns
function odsDate(value: string): string | null {
  const m = /^(-?\d{4,})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d):(\d\d)(\.\d+)?)?/.exec(value);
  if (!m) return null;
  const ms = Date.UTC(+m[1], +m[2] - 1, +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0), Math.round(+(m[7] ?? 0) * 1000));
  return isNaN(ms) ? null : new Date(ms).toISOString();
}

// "PT14H04M30S" as a fraction of a day, shown like an .xlsx time-only cell
function odsTime(value: string): string | null {
  const m = /^(-)?P(?:(\d+)D)?T(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?$/.exec(value);
  if (!m) return null;
  const seconds = (+(m[2] ?? 0) * 24 + +(m[3] ?? 0)) * 3600 + +(m[4] ?? 0) * 60 + +(m[5] ?? 0);
  return excelToIsoDate((m[1] ? -seconds : seconds) / 86400);
}

// OpenFormula to Excel syntax: "of:=SUM([.A1:.B2];[$Sheet2.C3])" -> "SUM(A1:B2,Sheet2!C3)"
export function odsFormula(formula: string): string {
  let f = formula.replace(/^(?:of:)?=/, '');
  f = f.replace(/\[([^\]]*)\]/g, (_, ref: string) => ref.split(':').map(part => {
    const dot = part.lastIndexOf('.');
    const sheet = part.slice(0, dot).replace(/^\$/, '');
    const cell = part.slice(dot + 1);
    return sheet ? `${sheet.startsWith("'") ? sheet : /^[A-Za-z_][\w.]*$/.test(sheet) ? sheet : `'${sheet}'`}!${cell}` : cell;
  }).join(':').replace(/^([^!]+!)([^:]+):\1/, '$1$2:'));
  // Separators outside string literals
  return f.replace(/"[^"]*"|;/g, s => s === ';' ? ',' : s);
}

// "$Sheet1.$A$1:.$B$2" -> "Sheet1!$A$1:$B$2"
function odsRange(address: string): string {
  return address.split(' ').map(range => odsFormula(`[${range}]`)).join(',');
}

export interface OdsSheet { name: string; state: 'visible' | 'hidden' }

export async function readOdsWorkbook(content: TokenSource, readText: (name: string) => Promise<string>) {
  const sheets: OdsSheet[] = [];
  const definedNames: DefinedName[] = [];
  const hiddenStyles = new Set<string>();
  let style: string | undefined;
  for await (const t of tokens(content)) {
    if (t.type !== 'startElement') continue;
    const a = t.attributes;
    if (t.name === 'style' && a['family'] === 'table') style = a['name'];
    else if (t.name === 'table-properties' && style && a['display'] === 'false') hiddenStyles.add(style);
    else if (t.name === 'table' && a['name'] !== undefined) {
      sheets.push({ name: a['name'], state: hiddenStyles.has(a['style-name'] ?? '') ? 'hidden' : 'visible' });
    } else if (t.name === 'named-range' && a['name'] && a['cell-range-address']) {
      definedNames.push({ name: a['name'], ref: odsRange(a['cell-range-address']) });
    } else if (t.name === 'named-expression' && a['name'] && a['expression']) {
      definedNames.push({ name: a['name'], ref: odsFormula(a['expression']) });
    }
  }

  const meta = await readText('meta.xml');
  const text = (tag: string) => { for (const el of xmlElements(meta, tag)) return unescapeXml(el.body); return undefined; };
  const properties: WorkbookProperties = {};
  for (const [key, tag] of [['title', 'dc:title'], ['subject', 'dc:subject'], ['creator', 'meta:initial-creator'],
    ['keywords', 'meta:keyword'], ['description', 'dc:description']] as const) {
    const v = text(tag);
    if (v) properties[key] = v;
  }
  if (!properties.creator && text('dc:creator')) properties.creator = text('dc:creator');
  const created = text('meta:creation-date');
  if (created && !isNaN(Date.parse(created))) properties.created = new Date(created);
  return { sheets, definedNames, properties };
}

// Frozen panes are view settings, kept in settings.xml per sheet
async function freezePanes(readText: (name: string) => Promise<string>, sheet: string): Promise<SheetMetadata['freezePanes']> {
  const settings = await readText('settings.xml');
  // Per-sheet entries sit inside the view's own (unnamed) entry and hold no entries themselves
  for (const m of settings.matchAll(/<config:config-item-map-entry\s+config:name="([^"]*)"\s*>([\s\S]*?)<\/config:config-item-map-entry>/g)) {
    if (unescapeXml(m[1]) !== sheet) continue;
    const item = (name: string) => {
      for (const el of xmlElements(m[2], 'config:config-item')) if (el.attrs.includes(`config:name="${name}"`)) return parseInt(el.body, 10) || 0;
      return 0;
    };
    const cols = item('HorizontalSplitMode') === 2 ? item('HorizontalSplitPosition') : 0;
    const rows = item('VerticalSplitMode') === 2 ? item('VerticalSplitPosition') : 0;
    if (!rows && !cols) return undefined;
    const pane = rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight';
    return { row: rows, col: cols, topLeftCell: `${colLetter(cols ? item('PositionRight') : 0)}${(rows ? item('PositionBottom') : 0) + 1}`, activePane: pane };
  }
  return undefined;
}

export interface OdsParseOptions { sheetName?: string; formulas?: boolean; formatted?: boolean; maxPaddingCells?: number }

export function parseOds(content: TokenSource, readText: (name: string) => Promise<string>, options: OdsParseOptions): ParseResult {
  let resolveMeta!: (m: SheetMetadata) => void;
  let rejectMeta!: (e: unknown) => void;
  const metadata = new Promise<SheetMetadata>((res, rej) => { resolveMeta = res; rejectMeta = rej; });
  metadata.catch(() => {}); // surfaced by the row iterator; awaiting getMetadata() still sees it
  const comments: SheetComment[] = [];
  let finished: Promise<void> = Promise.resolve();

  async function* generate(): AsyncGenerator<RowData> {
    const meta: SheetMetadata = { mergedCells: [], hiddenRows: [], hiddenCols: [], hyperlinks: [] };
    let done!: () => void;
    finished = new Promise(r => { done = r; });
    try {
      let found: string | undefined;
      let depth = 0;           // table elements open inside the target sheet (a table in a cell's shape, say)
      let row = 0, col = 0, column = 0;
      let rowRepeat = 1, rowHidden = false;
      let cells: CellValue[] = [], formulas: (string | undefined)[] | undefined, formatted: (string | undefined)[] | undefined;
      let cell: Record<string, string> | undefined;
      let cellRepeat = 1;
      let text: string[] = [], para = -1, link: string | undefined;
      let note: { text: string[]; author?: string; inAuthor: boolean } | undefined;
      let padding = 0, rowPadding = 0;
      const checkPadding = (total: number) => {
        if (total > (options.maxPaddingCells ?? Infinity)) {
          throw new Error(`Security Error: rows padded with more than ${options.maxPaddingCells} empty cells (maxUncompressedBytes).`);
        }
      };
      let emitted = 0;

      for await (const t of tokens(content)) {
        if (!found) {
          if (t.type === 'startElement' && t.name === 'table' && (options.sheetName === undefined || t.attributes['name'] === options.sheetName)) {
            found = t.attributes['name'] ?? '';
          }
          continue;
        }
        if (t.type === 'startElement') {
          const a = t.attributes;
          if (t.name === 'table') { depth++; continue; }
          if (depth > 0) continue;
          if (note) {
            if (t.name === 'creator') note.inAuthor = true;
            else if (t.name === 'p') note.text.push('');
            continue;
          }
          switch (t.name) {
            case 'table-column': {
              const n = repeat(a['number-columns-repeated'], MAX_COLUMNS - column);
              if (a['visibility'] === 'collapse') for (let c = column; c < column + n; c++) meta.hiddenCols.push(c + 1);
              column += n;
              break;
            }
            case 'table-row':
              rowRepeat = repeat(a['number-rows-repeated'], MAX_ROWS - row);
              rowHidden = a['visibility'] === 'collapse';
              cells = []; formulas = undefined; formatted = undefined; col = 0; rowPadding = 0;
              break;
            case 'table-cell': case 'covered-table-cell':
              cell = a; cellRepeat = repeat(a['number-columns-repeated'], MAX_COLUMNS - col);
              text = []; para = -1; link = undefined;
              if (t.name === 'table-cell' && (a['number-columns-spanned'] || a['number-rows-spanned'])) {
                const cols = repeat(a['number-columns-spanned'], MAX_COLUMNS - col), rows = repeat(a['number-rows-spanned'], MAX_ROWS - row);
                if (cols > 1 || rows > 1) meta.mergedCells.push(`${colLetter(col)}${row + 1}:${colLetter(col + cols - 1)}${row + rows}`);
              }
              break;
            case 'annotation':
              note = { text: [], inAuthor: false };
              break;
            case 'p':
              if (cell) { text.push(''); para++; }
              break;
            case 's':
              if (cell && para >= 0) text[para] += ' '.repeat(repeat(a['c'], 1000));
              break;
            case 'tab':
              if (cell && para >= 0) text[para] += '\t';
              break;
            case 'line-break':
              if (cell && para >= 0) text[para] += '\n';
              break;
            case 'a':
              if (cell && a['href']) link = a['href'];
              break;
          }
        } else if (t.type === 'text') {
          if (depth > 0) continue;
          if (note) {
            if (note.inAuthor) note.author = (note.author ?? '') + t.value;
            else if (note.text.length) note.text[note.text.length - 1] += t.value;
          } else if (cell && para >= 0) text[para] += t.value;
        } else {
          if (t.name === 'table') {
            if (depth-- > 0) continue;
            break; // end of the sheet
          }
          if (depth > 0) continue;
          if (note) {
            if (t.name === 'creator') note.inAuthor = false;
            else if (t.name === 'annotation') {
              comments.push({ ref: `${colLetter(col)}${row + 1}`, text: note.text.join('\n'), ...(note.author ? { author: note.author } : {}) });
              note = undefined;
            }
            continue;
          }
          if ((t.name === 'table-cell' || t.name === 'covered-table-cell') && cell) {
            const a = cell;
            const shown = text.join('\n');
            let value: CellValue = null;
            switch (a['value-type']) {
              case 'float': case 'percentage': case 'currency': {
                const n = Number(a['value']);
                value = isNaN(n) ? shown || null : n === 0 ? 0 : n;
                break;
              }
              case 'date': value = odsDate(a['date-value'] ?? '') ?? shown; break;
              case 'time': value = odsTime(a['time-value'] ?? '') ?? shown; break;
              case 'boolean': value = a['boolean-value'] === 'true'; break;
              case 'string': value = a['string-value'] ?? shown; break;
              default: value = shown || null; // formula errors carry only their text
            }
            if (value !== null || a['formula']) {
              rowPadding += Math.max(0, col - cells.length);
              checkPadding(padding + rowPadding);
              while (cells.length < col) cells.push(null);
              for (let i = 0; i < cellRepeat; i++) {
                cells[col + i] = value;
                if (options.formulas && a['formula']) (formulas ??= [])[col + i] = odsFormula(a['formula']);
                if (options.formatted && value !== null) (formatted ??= [])[col + i] = shown;
              }
              if (link && value !== null) meta.hyperlinks.push({ ref: `${colLetter(col)}${row + 1}`, hyperlink: link });
            }
            col += cellRepeat;
            cell = undefined;
          } else if (t.name === 'table-row') {
            if (rowHidden) for (let r = row; r < row + rowRepeat; r++) meta.hiddenRows.push(r + 1);
            if (cells.length) {
              for (let i = 0; i < rowRepeat; i++) {
                // Each repeated copy carries its padding again, so a tiny file can't expand without bound
                padding += rowPadding;
                checkPadding(padding);
                if (++emitted > MAX_ROWS) throw new Error('Sheet has more than 1,048,576 rows');
                const data: RowData = { rowNumber: row + i + 1, cells: i ? cells.slice() : cells };
                if (formulas) data.formulas = formulas;
                if (formatted) data.formatted = formatted;
                yield data;
              }
            }
            row += rowRepeat;
          }
        }
      }
      if (!found) throw new Error(options.sheetName === undefined ? 'No worksheets found in the workbook.' : `Sheet with name "${options.sheetName}" not found in workbook.`);
      meta.freezePanes = await freezePanes(readText, found);
      resolveMeta(meta);
    } catch (e) {
      rejectMeta(e);
      throw e;
    } finally {
      done();
    }
  }

  const rows = generate();
  return new ParseResult(rows, metadata, async () => [], async () => { await finished; return comments; });
}
