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

interface Relationship { id: string; type: string; path: string; external: boolean }

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
    workbookXml, sheets,
    sharedStrings: byType(rels, '/sharedStrings'), styles: byType(rels, '/styles'), theme: byType(rels, '/theme'),
  };
}

// Relationship id -> target of a worksheet's hyperlinks (from its .rels part)
export async function hyperlinkTargets(readText: (path: string) => Promise<string>, sheetPath: string): Promise<Map<string, string>> {
  const rels = parseRels(await readText(relsPathOf(sheetPath)), dirOf(sheetPath));
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

// Moves the relative references of a formula by (rows, cols), as Excel does when it fills a shared
// formula from its anchor cell into the other cells of the range. String literals and quoted sheet
// names are left alone; a reference pushed off the sheet becomes #REF!.
export function shiftFormula(formula: string, rows: number, cols: number): string {
  if (!rows && !cols) return formula;
  const col = (abs: string, letters: string) => {
    if (abs) return abs + letters;
    const c = colIndex(letters) + cols;
    return c >= 0 && c < MAX_COLUMNS ? colLetter(c) : null;
  };
  const row = (abs: string, digits: string) => {
    if (abs) return abs + digits;
    const r = parseInt(digits, 10) + rows;
    return r >= 1 && r <= MAX_ROWS ? String(r) : null;
  };
  const shift = (_m: string, ...g: string[]) => {
    const out = g[1] !== undefined ? [col(g[0], g[1]), row(g[2], g[3])]
      : g[5] !== undefined ? [col(g[4], g[5]), ':', col(g[6], g[7])]
      : [row(g[8], g[9]), ':', row(g[10], g[11])];
    return out.includes(null) ? '#REF!' : out.join('');
  };
  return formula
    .split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/)
    .map((part, i) => (i % 2 ? part : part.replace(REF_RE, shift)))
    .join('');
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
