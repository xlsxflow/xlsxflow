// Reader for Excel 97-2003 workbooks (.xls, BIFF8 records in a compound file) [MS-XLS].
// Returns the same rows and metadata as the .xlsx reader. The file is held in memory; .xls sheets
// are limited to 65,536 rows, so that stays modest.

import { CfbReader } from './cfb';
import { ParseResult, excelToIsoDate, type CellValue, type RowData, type SheetMetadata } from './worksheet-parser';
import { BUILTIN_NUM_FMTS, DISPLAY_NUM_FMTS, isBuiltinDateFormat, formatValue } from './number-format';
import { colLetter, isDateFormatCode } from './utils';
import type { DefinedName, WorkbookProperties } from './types';

const corrupt = (why: string) => new Error(`Corrupt .xls file: ${why}`);

const ERRORS: Record<number, string> = {
  0x00: '#NULL!', 0x07: '#DIV/0!', 0x0f: '#VALUE!', 0x17: '#REF!', 0x1d: '#NAME?', 0x24: '#NUM!', 0x2a: '#N/A', 0x2b: '#GETTING_DATA',
};

// Characters stored one byte each are UTF-16 code units with a zero high byte, i.e. Latin-1
function latin1(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i += 8192) s += String.fromCharCode(...bytes.subarray(i, i + 8192));
  return s;
}
const utf16 = new TextDecoder('utf-16le');
const chars = (bytes: Uint8Array, high: boolean) => high ? utf16.decode(bytes) : latin1(bytes);

const u16 = (b: Uint8Array, o: number) => {
  if (o + 2 > b.length) throw corrupt('record is cut short');
  return b[o] | (b[o + 1] << 8);
};
const u32 = (b: Uint8Array, o: number) => (u16(b, o) | (u16(b, o + 2) << 16)) >>> 0;
const f64 = (b: Uint8Array, o: number) => {
  if (o + 8 > b.length) throw corrupt('record is cut short');
  return new DataView(b.buffer, b.byteOffset + o, 8).getFloat64(0, true);
};

// XLUnicodeString (16-bit length) or ShortXLUnicodeString (8-bit length); returns the text and its size
function xlString(b: Uint8Array, o: number, lengthBytes: 1 | 2): [string, number] {
  const cch = lengthBytes === 1 ? b[o] : u16(b, o);
  const high = (b[o + lengthBytes] & 1) === 1;
  const start = o + lengthBytes + 1;
  const end = start + cch * (high ? 2 : 1);
  if (end > b.length) throw corrupt('string runs past its record');
  return [chars(b.subarray(start, end), high), end - o];
}

// RK numbers: a 30-bit integer or the top 30 bits of a double, optionally divided by 100
function rk(value: number): number {
  let n: number;
  if (value & 2) n = value >> 2;
  else {
    const view = new DataView(new ArrayBuffer(8));
    view.setUint32(4, value & 0xfffffffc, true);
    n = view.getFloat64(0, true);
  }
  return value & 1 ? n / 100 : n;
}

interface Rec { type: number; data: Uint8Array; next: number }

function recordAt(stream: Uint8Array, pos: number): Rec | undefined {
  if (pos + 4 > stream.length) return undefined;
  const len = u16(stream, pos + 2);
  if (pos + 4 + len > stream.length) throw corrupt(`record at ${pos} runs past the end of the stream`);
  return { type: u16(stream, pos), data: stream.subarray(pos + 4, pos + 4 + len), next: pos + 4 + len };
}

// The shared string table spans CONTINUE records. A string's characters may break across them, and
// each continuation then starts with a fresh flags byte saying whether the rest is 8- or 16-bit.
function parseSst(chunks: Uint8Array[]): string[] {
  let ci = 0, off = 8; // skip cstTotal and cstUnique
  const total = u32(chunks[0], 4);
  const strings: string[] = [];
  const advance = () => { while (ci < chunks.length && off >= chunks[ci].length) { ci++; off = 0; } };
  const read16 = () => { advance(); const v = u16(chunks[ci] ?? new Uint8Array(), off); off += 2; return v; };
  const read32 = () => { const lo = read16(); return (lo | (read16() << 16)) >>> 0; };
  const skip = (n: number) => {
    while (n > 0) {
      advance();
      if (ci >= chunks.length) throw corrupt('shared string table is cut short');
      const take = Math.min(n, chunks[ci].length - off);
      off += take; n -= take;
    }
  };
  for (let i = 0; i < total; i++) {
    advance();
    if (ci >= chunks.length) break; // some writers overstate the count
    const cch = read16();
    advance();
    const flags = chunks[ci]?.[off++];
    if (flags === undefined) throw corrupt('shared string table is cut short');
    let high = (flags & 1) === 1;
    const runs = flags & 8 ? read16() : 0;
    const ext = flags & 4 ? read32() : 0;
    let text = '';
    for (let left = cch; left > 0;) {
      if (off >= chunks[ci].length) {
        ci++; off = 0;
        if (ci >= chunks.length) throw corrupt('shared string table is cut short');
        high = (chunks[ci][off++] & 1) === 1;
      }
      const take = Math.min(left, Math.floor((chunks[ci].length - off) / (high ? 2 : 1)));
      if (take === 0) throw corrupt('a 16-bit character is split across records');
      text += chars(chunks[ci].subarray(off, off + take * (high ? 2 : 1)), high);
      off += take * (high ? 2 : 1);
      left -= take;
    }
    skip(runs * 4 + ext);
    strings.push(text);
  }
  return strings;
}

export interface XlsSheet { name: string; state: 'visible' | 'hidden' | 'veryHidden'; offset: number; isWorksheet: boolean }

export class XlsWorkbook {
  readonly sheets: XlsSheet[] = [];
  readonly definedNames: DefinedName[] = [];
  readonly properties: WorkbookProperties = {};
  private sst: string[] = [];
  private xfFormats: number[] = [];
  private formats = new Map<number, string>(BUILTIN_NUM_FMTS);
  private dateFormats = new Set<number>();
  private date1904 = false;
  private readonly stream: Uint8Array;

  constructor(cfb: CfbReader) {
    const stream = cfb.read('Workbook') ?? cfb.read('Book');
    if (!stream) throw new Error('Not an Excel workbook: the compound file has no Workbook stream');
    this.stream = stream;
    this.parseGlobals();
    readProperties(cfb, this.properties);
  }

  private parseGlobals() {
    const s = this.stream;
    const bof = recordAt(s, 0);
    if (!bof || (bof.type & 0xff) !== 0x09) throw corrupt('no BOF record');
    if (bof.type !== 0x0809 || u16(bof.data, 0) !== 0x0600) {
      throw new Error('Only Excel 97-2003 (BIFF8) .xls files are supported; this one is from Excel 95 or earlier');
    }
    const externSheets: number[] = []; // XTI index -> first sheet index
    const names: { name: string; builtin: boolean; scope: number; rgce: Uint8Array }[] = [];
    for (let r = recordAt(s, bof.next); r; r = recordAt(s, r.next)) {
      const d = r.data;
      switch (r.type) {
        case 0x000a: // EOF
          this.resolveNames(names, externSheets);
          return;
        case 0x002f: // FILEPASS
          throw new Error('This .xls file is password-protected. Encrypted .xls files are not supported; save it as .xlsx in Excel');
        case 0x0022: // DATEMODE
          this.date1904 = u16(d, 0) === 1;
          break;
        case 0x041e: { // FORMAT
          const id = u16(d, 0);
          const [code] = xlString(d, 2, 2);
          this.formats.set(id, code);
          if (isDateFormatCode(code)) this.dateFormats.add(id);
          break;
        }
        case 0x00e0: // XF
          this.xfFormats.push(u16(d, 2));
          break;
        case 0x0085: { // BOUNDSHEET
          const [name] = xlString(d, 6, 1);
          const state = d[4] & 3;
          this.sheets.push({ name, state: state === 1 ? 'hidden' : state === 2 ? 'veryHidden' : 'visible', offset: u32(d, 0), isWorksheet: d[5] === 0 });
          break;
        }
        case 0x00fc: { // SST, with its CONTINUE records
          const chunks = [d];
          let next = recordAt(s, r.next);
          while (next?.type === 0x003c) { chunks.push(next.data); r = next; next = recordAt(s, r.next); }
          this.sst = parseSst(chunks);
          break;
        }
        case 0x0017: // EXTERNSHEET
          for (let i = 0, n = u16(d, 0); i < n; i++) externSheets.push(u16(d, 2 + i * 6 + 2) === 0xfffe ? -1 : u16(d, 2 + i * 6 + 2));
          break;
        case 0x0018: { // NAME
          const flags = u16(d, 0);
          const cch = d[3], cce = u16(d, 4), scope = u16(d, 8);
          const high = (d[14] & 1) === 1;
          const nameEnd = 15 + cch * (high ? 2 : 1);
          const raw = d.subarray(15, nameEnd);
          const builtin = (flags & 0x20) !== 0;
          names.push({ name: builtin ? (BUILTIN_NAMES[raw[0]] ?? `_xlnm.${raw[0]}`) : chars(raw, high), builtin, scope, rgce: d.subarray(nameEnd, nameEnd + cce) });
          break;
        }
      }
    }
    throw corrupt('workbook globals have no EOF record');
  }

  // Names whose formula is a single reference or range on a sheet; other formulas are left out
  private resolveNames(names: { name: string; scope: number; rgce: Uint8Array }[], externSheets: number[]) {
    for (const { name, scope, rgce } of names) {
      const ref = this.reference(rgce, externSheets);
      if (ref === undefined) continue;
      const out: DefinedName = { name, ref };
      if (scope > 0 && this.sheets[scope - 1]) out.sheet = this.sheets[scope - 1].name;
      this.definedNames.push(out);
    }
  }

  private reference(rgce: Uint8Array, externSheets: number[]): string | undefined {
    const parts: string[] = [];
    for (let i = 0; i < rgce.length;) {
      const ptg = rgce[i] & 0x1f | 0x20; // reference class does not matter here
      const sheetName = (xti: number) => {
        const sheet = this.sheets[externSheets[xti]];
        if (!sheet) return undefined;
        return /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheet.name) ? sheet.name : `'${sheet.name.replace(/'/g, "''")}'`;
      };
      const cell = (row: number, col: number) => `$${colLetter(col & 0x3fff)}$${row + 1}`;
      if (ptg === 0x3a && i + 7 <= rgce.length) { // ptgRef3d
        const sheet = sheetName(u16(rgce, i + 1));
        if (!sheet) return undefined;
        parts.push(`${sheet}!${cell(u16(rgce, i + 3), u16(rgce, i + 5))}`);
        i += 7;
      } else if (ptg === 0x3b && i + 11 <= rgce.length) { // ptgArea3d
        const sheet = sheetName(u16(rgce, i + 1));
        if (!sheet) return undefined;
        const [r1, r2, c1, c2] = [u16(rgce, i + 3), u16(rgce, i + 5), u16(rgce, i + 7) & 0xff, u16(rgce, i + 9) & 0xff];
        // Whole rows and columns read as $1:$2 and $A:$B, as in .xlsx
        const area = r1 === 0 && r2 === 0xffff ? `$${colLetter(c1)}:$${colLetter(c2)}`
          : c1 === 0 && c2 === 0xff ? `$${r1 + 1}:$${r2 + 1}` : `${cell(r1, c1)}:${cell(r2, c2)}`;
        parts.push(`${sheet}!${area}`);
        i += 11;
      } else if (rgce[i] === 0x10 && parts.length) { // ptgUnion, as in print titles "A:A,1:1"
        i++;
      } else if (ptg === 0x29) { // ptgMemFunc: wraps the references that follow
        i += 3;
      } else {
        return undefined;
      }
    }
    return parts.length ? parts.join(',') : undefined;
  }

  private formatCode(xf: number): string | undefined {
    const id = this.xfFormats[xf];
    return id === undefined ? undefined : DISPLAY_NUM_FMTS.get(id) ?? this.formats.get(id);
  }

  private isDate(xf: number): boolean {
    const id = this.xfFormats[xf];
    return id !== undefined && (isBuiltinDateFormat(id) || this.dateFormats.has(id));
  }

  parseSheet(name: string | undefined, formatted: boolean): ParseResult {
    const sheet = name === undefined ? this.sheets.find(s => s.isWorksheet) : this.sheets.find(s => s.name === name);
    if (!sheet) throw new Error(name === undefined ? 'No worksheets found in the workbook.' : `Sheet with name "${name}" not found in workbook.`);
    if (!sheet.isWorksheet) throw new Error(`Sheet "${sheet.name}" is a chart or macro sheet, not a worksheet`);

    const rows = new Map<number, { cells: CellValue[]; formatted?: (string | undefined)[] }>();
    const meta: SheetMetadata = { mergedCells: [], hiddenRows: [], hiddenCols: [], hyperlinks: [] };
    const set = (row: number, col: number, xf: number, value: CellValue, raw?: number) => {
      let r = rows.get(row);
      if (!r) rows.set(row, r = { cells: [] });
      while (r.cells.length < col) r.cells.push(null);
      if (typeof value === 'number' && this.isDate(xf)) value = excelToIsoDate(value, this.date1904);
      else if (value === -0) value = 0;
      r.cells[col] = value;
      if (formatted && value !== null) {
        (r.formatted ??= [])[col] = typeof value === 'string' && ERROR_TEXTS.has(value) && raw === undefined
          ? value : formatValue(raw ?? value, this.formatCode(xf) ?? 'General', this.date1904);
      }
    };
    const blank = (row: number, col: number) => {
      let r = rows.get(row);
      if (!r) rows.set(row, r = { cells: [] });
      while (r.cells.length <= col) r.cells.push(null);
    };

    const s = this.stream;
    const bof = recordAt(s, sheet.offset);
    if (!bof || bof.type !== 0x0809) throw corrupt(`sheet "${sheet.name}" does not start with a BOF record`);
    let pendingString: { row: number; col: number; xf: number } | undefined;
    let frozen = false;
    let pane: SheetMetadata['freezePanes'];
    let depth = 0; // charts embedded in the sheet are nested BOF...EOF blocks
    for (let r = recordAt(s, bof.next); r; r = recordAt(s, r.next)) {
      const d = r.data;
      if (r.type === 0x0809) { depth++; continue; }
      if (r.type === 0x000a) { if (depth-- === 0) break; continue; }
      if (depth > 0) continue;
      switch (r.type) {
        case 0x0203: // NUMBER
          set(u16(d, 0), u16(d, 2), u16(d, 4), f64(d, 6), f64(d, 6));
          break;
        case 0x027e: { // RK
          const n = rk(u32(d, 6));
          set(u16(d, 0), u16(d, 2), u16(d, 4), n, n);
          break;
        }
        case 0x00bd: { // MULRK
          const row = u16(d, 0), first = u16(d, 2);
          for (let i = 0; 4 + i * 6 + 6 <= d.length - 2; i++) {
            const n = rk(u32(d, 4 + i * 6 + 2));
            set(row, first + i, u16(d, 4 + i * 6), n, n);
          }
          break;
        }
        case 0x00fd: // LABELSST
          set(u16(d, 0), u16(d, 2), u16(d, 4), this.sst[u32(d, 6)] ?? '');
          break;
        case 0x0204: case 0x00d6: // LABEL, RSTRING
          set(u16(d, 0), u16(d, 2), u16(d, 4), xlString(d, 6, 2)[0]);
          break;
        case 0x0205: { // BOOLERR
          const isError = d[7] === 1;
          set(u16(d, 0), u16(d, 2), u16(d, 4), isError ? ERRORS[d[6]] ?? '#VALUE!' : d[6] === 1);
          break;
        }
        case 0x0201: // BLANK
          blank(u16(d, 0), u16(d, 2));
          break;
        case 0x00be: // MULBLANK
          blank(u16(d, 0), u16(d, d.length - 2));
          break;
        case 0x0006: { // FORMULA: the cached result
          const row = u16(d, 0), col = u16(d, 2), xf = u16(d, 4);
          pendingString = undefined;
          if (u16(d, 12) !== 0xffff) set(row, col, xf, f64(d, 6), f64(d, 6));
          else if (d[6] === 0) pendingString = { row, col, xf }; // text follows in a STRING record
          else if (d[6] === 1) set(row, col, xf, d[8] === 1);
          else if (d[6] === 2) set(row, col, xf, ERRORS[d[8]] ?? '#VALUE!');
          else if (d[6] === 3) set(row, col, xf, '');
          break;
        }
        case 0x0207: // STRING
          if (pendingString) set(pendingString.row, pendingString.col, pendingString.xf, xlString(d, 0, 2)[0]);
          pendingString = undefined;
          break;
        case 0x00e5: // MERGECELLS
          for (let i = 0, n = u16(d, 0); i < n; i++) {
            const o = 2 + i * 8;
            meta.mergedCells.push(`${colLetter(u16(d, o + 4))}${u16(d, o) + 1}:${colLetter(u16(d, o + 6))}${u16(d, o + 2) + 1}`);
          }
          break;
        case 0x0208: // ROW
          if (u16(d, 12) & 0x20) meta.hiddenRows.push(u16(d, 0) + 1);
          break;
        case 0x007d: // COLINFO
          if (u16(d, 8) & 1) for (let c = u16(d, 0); c <= Math.min(u16(d, 2), 255); c++) meta.hiddenCols.push(c + 1);
          break;
        case 0x023e: // WINDOW2
          frozen = (u16(d, 0) & 0x08) !== 0;
          break;
        case 0x0041: { // PANE
          const col = u16(d, 0), row = u16(d, 2);
          pane = { row, col, topLeftCell: `${colLetter(u16(d, 6))}${u16(d, 4) + 1}`, activePane: PANES[d[8]] ?? 'bottomRight' };
          break;
        }
      }
    }
    if (frozen && pane) meta.freezePanes = pane;

    const sorted = [...rows.entries()].sort((a, b) => a[0] - b[0]);
    async function* generate(): AsyncGenerator<RowData> {
      for (const [row, r] of sorted) {
        const data: RowData = { rowNumber: row + 1, cells: r.cells };
        if (r.formatted) data.formatted = r.formatted;
        yield data;
      }
    }
    return new ParseResult(generate(), Promise.resolve(meta));
  }
}

const ERROR_TEXTS = new Set(Object.values(ERRORS));
const PANES = ['bottomRight', 'topRight', 'bottomLeft', 'topLeft'];
const BUILTIN_NAMES: Record<number, string> = {
  0x00: '_xlnm.Consolidate_Area', 0x01: '_xlnm.Auto_Open', 0x02: '_xlnm.Auto_Close', 0x03: '_xlnm.Extract', 0x04: '_xlnm.Database',
  0x05: '_xlnm.Criteria', 0x06: '_xlnm.Print_Area', 0x07: '_xlnm.Print_Titles', 0x08: '_xlnm.Recorder', 0x09: '_xlnm.Data_Form',
  0x0a: '_xlnm.Auto_Activate', 0x0b: '_xlnm.Auto_Deactivate', 0x0c: '_xlnm.Sheet_Title', 0x0d: '_xlnm._FilterDatabase',
};

// Title, author and so on, from the OLE property sets [MS-OLEPS]
function readProperties(cfb: CfbReader, out: WorkbookProperties) {
  const read = (stream: string, keys: Record<number, keyof WorkbookProperties>) => {
    const b = cfb.read(stream);
    if (!b || b.length < 48) return;
    try {
      const section = u32(b, 44);
      const count = u32(b, section + 4);
      let codepage = 1252;
      const values = new Map<number, number>();
      for (let i = 0; i < Math.min(count, 1000); i++) values.set(u32(b, section + 8 + i * 8), section + u32(b, section + 12 + i * 8));
      const cp = values.get(1);
      if (cp !== undefined && u16(b, cp) === 2) codepage = u16(b, cp + 4);
      for (const [id, key] of Object.entries(keys)) {
        const at = values.get(+id);
        if (at === undefined) continue;
        const type = u16(b, at);
        if (type === 0x1e) { // VT_LPSTR, in the property set's code page
          const len = u32(b, at + 4);
          const raw = b.subarray(at + 8, at + 8 + len);
          const text = codepage === 1200 ? utf16.decode(raw) : codepage === 65001 ? new TextDecoder().decode(raw) : latin1(raw);
          const value = text.replace(/\0+$/, '');
          if (value) (out as Record<string, unknown>)[key] = value;
        } else if (type === 0x40 && key === 'created') { // VT_FILETIME: 100 ns ticks since 1601
          const ticks = u32(b, at + 4) + u32(b, at + 8) * 2 ** 32;
          if (ticks) out.created = new Date(ticks / 10000 - 11644473600000);
        }
      }
    } catch {
      // Properties are optional; a damaged property set must not stop the workbook from opening
    }
  };
  read('\u0005SummaryInformation', { 2: 'title', 3: 'subject', 4: 'creator', 5: 'keywords', 6: 'description', 12: 'created' });
  read('\u0005DocumentSummaryInformation', { 2: 'category', 14: 'manager', 15: 'company' });
}
