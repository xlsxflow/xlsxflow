import type { ParseResult, CellValue } from './worksheet-parser';

export async function sheetToJson<T = Record<string, CellValue>>(
  parseResult: ParseResult,
  headerRowIndex: number = 0
): Promise<T[]> {
  const results: T[] = [];
  let headers: string[] = [];

  for await (const row of parseResult) {
    // Treat the first row we get as the header
    if (headers.length === 0) {
      headers = row.cells.map((c, i) => (c !== null && c !== undefined ? String(c) : `Column${i + 1}`));
      continue;
    }

    const obj: any = {};
    for (let i = 0; i < headers.length; i++) {
      obj[headers[i]] = row.cells[i] ?? null;
    }
    results.push(obj as T);
  }

  return results;
}

function escapeCsv(val: any): string {
  if (val === null || val === undefined) return '';
  const str = String(val);
  if (str.includes(',') || str.includes('"') || str.includes('\n') || str.includes('\r')) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

export async function streamToCsv(parseResult: ParseResult): Promise<string> {
  const rows: string[] = [];
  for await (const row of parseResult) {
    const csvRow = row.cells.map(escapeCsv).join(',');
    rows.push(csvRow);
  }
  return rows.join('\n');
}

function attr(tag: string, name: string): string | null {
  const m = new RegExp(`\\s${name}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`).exec(tag);
  if (!m) return null;
  return (m[1] ?? m[2])
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
}

// Resolves a relationship Target against the directory of the part that owns the .rels file.
function resolveTarget(baseDir: string, target: string): string {
  const parts = (target.startsWith('/') ? target.slice(1) : baseDir + target).split('/');
  const out: string[] = [];
  for (const p of parts) {
    if (p === '..') out.pop();
    else if (p !== '.' && p !== '') out.push(p);
  }
  return out.join('/');
}

export interface Relationship { id: string; type: string; path: string; external: boolean }

function parseRels(relsXml: string, baseDir: string): Relationship[] {
  const rels: Relationship[] = [];
  for (const [tag] of relsXml.matchAll(/<(?:\w+:)?Relationship\b[^>]*>/g)) {
    const id = attr(tag, 'Id');
    const target = attr(tag, 'Target');
    if (!id || !target) continue;
    // External targets (hyperlink URLs) are kept verbatim, not resolved as package paths
    const external = attr(tag, 'TargetMode') === 'External';
    rels.push({ id, type: attr(tag, 'Type') ?? '', path: external ? target : resolveTarget(baseDir, target), external });
  }
  return rels;
}

const dirOf = (path: string) => path.slice(0, path.lastIndexOf('/') + 1);
const relsPathOf = (path: string) => `${dirOf(path)}_rels/${path.slice(path.lastIndexOf('/') + 1)}.rels`;
// Relationship types end the same way in Transitional and Strict OOXML
const byType = (rels: Relationship[], suffix: string) => rels.find(r => !r.external && r.type.endsWith(suffix))?.path;

export interface WorkbookParts {
  workbookPath: string;
  workbookXml: string;
  sheets: Map<string, string>;   // sheet name -> ZIP entry path, in tab order
  sharedStrings?: string;
  styles?: string;
  theme?: string;
}

// Locates the workbook and its parts through the package relationships instead of assuming xl/ names.
export async function resolveWorkbookParts(readText: (path: string) => Promise<string>): Promise<WorkbookParts> {
  const rootRels = parseRels(await readText('_rels/.rels'), '');
  const workbookPath = byType(rootRels, '/officeDocument') ?? 'xl/workbook.xml';
  const workbookXml = await readText(workbookPath);
  const rels = parseRels(await readText(relsPathOf(workbookPath)), dirOf(workbookPath));

  const sheets = new Map<string, string>();
  for (const [tag] of workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)) {
    const name = attr(tag, 'name');
    const rId = attr(tag, 'r:id') ?? attr(tag, '\\w+:id');
    const rel = rId ? rels.find(r => r.id === rId && !r.external) : undefined;
    if (name && rel) sheets.set(name, rel.path);
  }
  return {
    workbookPath, workbookXml, sheets,
    sharedStrings: byType(rels, '/sharedStrings'), styles: byType(rels, '/styles'), theme: byType(rels, '/theme'),
  };
}

// The relationships of any part (its _rels/<name>.rels), targets resolved to ZIP paths unless external
export async function partRelationships(readText: (path: string) => Promise<string>, partPath: string): Promise<Relationship[]> {
  return parseRels(await readText(relsPathOf(partPath)), dirOf(partPath));
}

// Relationship id -> target of a worksheet's hyperlinks (from its .rels part)
export async function hyperlinkTargets(readText: (path: string) => Promise<string>, sheetPath: string): Promise<Map<string, string>> {
  const rels = await partRelationships(readText, sheetPath);
  return new Map(rels.filter(r => r.type.endsWith('/hyperlink')).map(r => [r.id, r.path]));
}

// Maps sheet name -> ZIP entry path for a workbook at xl/workbook.xml, in tab order.
export function resolveSheetPaths(workbookXml: string, relsXml: string): Map<string, string> {
  const rels = parseRels(relsXml, 'xl/');
  const paths = new Map<string, string>();
  for (const [tag] of workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)) {
    const name = attr(tag, 'name');
    const rId = attr(tag, 'r:id') ?? attr(tag, '\\w+:id');
    const rel = rId ? rels.find(r => r.id === rId && !r.external) : undefined;
    if (name && rel) paths.set(name, rel.path);
  }
  return paths;
}

// Decodes OOXML ST_Xstring escapes (_x000D_ = CR, _x005F_ = literal underscore).
export function decodeXString(s: string): string {
  if (s.indexOf('_x') === -1) return s;
  return s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16)));
}

// Excel's sheet size (A..XFD, 1..1048576)
export const MAX_COLUMNS = 16384;
export const MAX_ROWS = 1048576;

// 0-based column of letters like "BC" (or a reference like "BC12"); -1 when there are none
export function colIndex(ref: string): number {
  const match = ref.match(/[A-Za-z]+/);
  if (!match) return -1;
  const letters = match[0].toUpperCase();
  let idx = 0;
  for (let i = 0; i < letters.length; i++) idx = idx * 26 + (letters.charCodeAt(i) - 64);
  return idx - 1;
}

// 0-based column index -> letters (0 -> A, 26 -> AA)
export function colLetter(index: number): string {
  let col = '';
  for (let c = index + 1; c > 0; c = Math.floor((c - 1) / 26)) col = String.fromCharCode(((c - 1) % 26) + 65) + col;
  return col;
}

// A1 cell refs, whole-column (A:C) and whole-row (1:3) ranges. The lookarounds keep function names
// (LOG10(), ATAN2()) and unquoted sheet names (Q1!A1) from being taken for references.
const REF_RE = /(?<![A-Za-z0-9_.$])(?:(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(!])|(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(!])|(\$?)(\d+):(\$?)(\d+)(?![0-9!]))/g;

// How a reference sits in the formula: alone, or as the first / last cell of a range
export type RefRole = 'single' | 'start' | 'end';
// Returns the new 0-based column / 1-based row, or null when the reference no longer exists (#REF!).
// `sheet` is the sheet qualifier written before the reference, if any (unquoted).
export type RefMapper = (value: number, absolute: boolean, role: RefRole, sheet: string | undefined) => number | null;

const unquoteSheet = (quoted: string) => quoted.slice(1, -1).replace(/''/g, "'");

// Rewrites every A1 reference in a formula through mapCol / mapRow. String literals are left alone.
// A reference moved off the sheet, or a range turned inside out, becomes #REF!.
export function mapFormulaRefs(formula: string, mapCol: RefMapper, mapRow: RefMapper): string {
  const parts = formula.split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/);
  for (let i = 0; i < parts.length; i += 2) {
    const part = parts[i];
    let prevEnd = -1;
    let prevSheet: string | undefined;
    let prevRow = 0;
    parts[i] = part.replace(REF_RE, (m: string, ...g: any[]) => {
      const offset: number = g[12];
      const before = part.slice(0, offset);
      const isEnd = before.endsWith(':') && prevEnd === offset - 1;
      let sheet: string | undefined;
      if (isEnd) sheet = prevSheet;
      else if (before.endsWith('!')) {
        sheet = /([A-Za-z0-9_.\u00C0-\uFFFF]+)!$/.exec(before)?.[1]
          ?? (offset === 1 && i > 0 && parts[i - 1].startsWith("'") ? unquoteSheet(parts[i - 1]) : undefined);
      }
      const role: RefRole = isEnd ? 'end' : part[offset + m.length] === ':' ? 'start' : 'single';
      const col = (abs: string, letters: string, r: RefRole) => {
        const c = mapCol(colIndex(letters), !!abs, r, sheet);
        return c !== null && c >= 0 && c < MAX_COLUMNS ? abs + colLetter(c) : null;
      };
      const row = (abs: string, digits: string, r: RefRole) => {
        const n = mapRow(parseInt(digits, 10), !!abs, r, sheet);
        return n !== null && n >= 1 && n <= MAX_ROWS ? n : null;
      };
      prevEnd = offset + m.length;
      prevSheet = sheet;
      let out: string | null;
      if (g[1] !== undefined) {
        const c = col(g[0], g[1], role), r = row(g[2], g[3], role);
        // The end of a range must not land above its start (rows removed from under it)
        const inverted = r !== null && role === 'end' && r < prevRow;
        prevRow = r ?? 0;
        out = c === null || r === null || inverted ? null : c + g[2] + r;
      } else if (g[5] !== undefined) {
        const a = col(g[4], g[5], 'start'), b = col(g[6], g[7], 'end');
        out = a === null || b === null ? null : `${a}:${b}`;
      } else {
        const a = row(g[8], g[9], 'start'), b = row(g[10], g[11], 'end');
        out = a === null || b === null || b < a ? null : `${g[8]}${a}:${g[10]}${b}`;
      }
      return out ?? '#REF!';
    // A range with a dead end is dead as a whole
    }).replace(/\$?[A-Za-z]{1,3}\$?\d+:#REF!|#REF!:\$?[A-Za-z]{1,3}\$?\d+/g, '#REF!');
  }
  return parts.join('');
}

// Moves the relative references of a formula by (rows, cols), as Excel does when it fills a shared
// formula from its anchor cell into the other cells of the range. Absolute parts ($A$1) stay put.
export function shiftFormula(formula: string, rows: number, cols: number): string {
  if (!rows && !cols) return formula;
  return mapFormulaRefs(formula, (c, abs) => (abs ? c : c + cols), (r, abs) => (abs ? r : r + rows));
}

// Date -> Excel 1900-system serial (UTC). Serials 1..60 sit one day early because of Excel's
// phantom 1900-02-29, mirroring excelToIsoDate on the read side.
// Dates before 1900-01-01 have no serial in Excel; they come out as small/negative numbers.
export function dateToSerial(d: Date): number {
  const serial = (d.getTime() - Date.UTC(1899, 11, 30)) / 86400000;
  return serial >= 2 && serial < 61 ? serial - 1 : serial;
}

// Inverse of decodeXString: literal "_xHHHH_" text is protected, and characters XML cannot carry
// (control characters) or would normalise away (CR) are written as escapes.
export function encodeXString(s: string): string {
  return s
    .replace(/_(?=x[0-9A-Fa-f]{4}_)/g, '_x005F_')
    .replace(/[\x00-\x08\x0B\x0C\r\x0E-\x1F]/g, c => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`);
}

// True when a number format renders dates/times. Ignores quoted text, escapes, colours and
// locale tags, so "0 \"days\"", "[Red]0.0" and "#,##0 \"USD\"" are not mistaken for dates.
export function isDateFormatCode(code: string): boolean {
  const stripped = code
    .replace(/"[^"]*"/g, '')
    .replace(/\\./g, '')
    .replace(/[_*]./g, '')
    .replace(/\[(?!(?:h+|m+|s+)\])[^\]]*\]/gi, '');   // keep elapsed-time [h] [mm] [ss]
  return /[ymdhs]/i.test(stripped) || /\[(?:h+|m+|s+)\]/i.test(stripped);
}

// After cells change, cached formula results are stale: mark the workbook for a full recalculation
// on open, and drop the calculation chain, which lists formula cells by address (Excel reports a
// broken file when a listed cell no longer holds a formula). Returns the parts to rewrite and the
// parts to leave out of the package.
export async function recalcOnOpen(
  readText: (path: string) => Promise<string>,
  parts: WorkbookParts,
  workbookXml = parts.workbookXml,
): Promise<{ replace: Map<string, string>; drop: Set<string> }> {
  let wb = workbookXml;
  if (/<(?:\w+:)?calcPr\b/.test(wb)) {
    wb = wb.replace(/<((?:\w+:)?)calcPr\b([^>]*?)(\/?)>/, (_m, p: string, attrs: string, slash: string) =>
      `<${p}calcPr${attrs.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${slash}>`);
  } else {
    // calcPr comes before these in CT_Workbook
    const p = /<((?:\w+:)?)workbook\b/.exec(wb)?.[1] ?? '';
    const at = wb.search(/<(?:\w+:)?(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/(?:\w+:)?workbook>/);
    wb = wb.slice(0, at) + `<${p}calcPr fullCalcOnLoad="1"/>` + wb.slice(at);
  }
  const replace = new Map([[parts.workbookPath, wb]]);
  const drop = new Set<string>();
  const calcChain = (await partRelationships(readText, parts.workbookPath)).find(r => r.type.endsWith('/calcChain'))?.path;
  if (calcChain) {
    drop.add(calcChain);
    const relsPath = relsPathOf(parts.workbookPath);
    replace.set(relsPath, (await readText(relsPath)).replace(/<(?:\w+:)?Relationship\b[^>]*?Type="[^"]*\/calcChain"[^>]*\/>/, ''));
    const escaped = calcChain.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    replace.set('[Content_Types].xml', (await readText('[Content_Types].xml'))
      .replace(new RegExp(`<(?:\\w+:)?Override\\b[^>]*?PartName="/${escaped}"[^>]*/>`), ''));
  }
  return { replace, drop };
}
