import { RandomAccessReader } from './random-access';
import { ZipRandomAccessParser } from './zip-random-access';
import { Row, SheetOptions, CellValue, CellStyle } from './types';
import { ZipStreamWriter } from './zip-stream-writer';
import { StylePatcher } from './style-patcher';
import {
  resolveWorkbookParts, partRelationships, recalcOnOpen, attr, dirOf, relsPathOf, validateSheetName,
  colIndex, colLetter, dateToSerial, encodeXString, shiftFormula, type WorkbookParts,
} from './utils';

// A new cell value; null clears it. The object form sets a formula (leading "=" optional) or a value,
// and can change the cell's style: only the style properties given change. With neither `value` nor
// `formula`, the cell keeps its content.
export type CellEdit = CellValue | { value?: CellValue; formula?: string; style?: CellStyle };
type EditedRows = Map<number, Map<number, CellEdit>>;

const KEEP = Symbol('keep');
function contentOf(edit: CellEdit): CellValue | { formula: string } | typeof KEEP {
  if (edit === null || typeof edit !== 'object' || edit instanceof Date) return edit ?? null;
  if (edit.formula !== undefined) return { formula: edit.formula };
  return 'value' in edit ? edit.value ?? null : KEEP;
}
const styleOf = (edit: CellEdit) =>
  edit !== null && typeof edit === 'object' && !(edit instanceof Date) ? edit.style : undefined;

const escapeXml = (val: string) =>
  val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
const unescapeXml = (s: string) =>
  s.replace(/&(?:#x([0-9a-fA-F]+)|#(\d+)|(amp|lt|gt|quot|apos));/g, (m, hex, dec, name) =>
    hex ? String.fromCodePoint(parseInt(hex, 16)) : dec ? String.fromCodePoint(parseInt(dec, 10))
      : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" } as Record<string, string>)[name] ?? m);

// The cell XML for an edit; `attrs` keeps the existing cell's style (s=...) and other attributes
function editedCell(p: string, ref: string, attrs: string, edit: CellValue | { formula: string }): string {
  if (edit === null || edit === undefined) return `<${p}c r="${ref}"${attrs}/>`;
  if (typeof edit === 'number') {
    return isFinite(edit) ? `<${p}c r="${ref}"${attrs}><${p}v>${edit}</${p}v></${p}c>` : `<${p}c r="${ref}"${attrs} t="e"><${p}v>#NUM!</${p}v></${p}c>`;
  }
  if (typeof edit === 'boolean') return `<${p}c r="${ref}"${attrs} t="b"><${p}v>${edit ? 1 : 0}</${p}v></${p}c>`;
  if (edit instanceof Date) {
    if (isNaN(edit.getTime())) throw new Error(`Invalid Date for cell ${ref}`);
    return `<${p}c r="${ref}"${attrs}><${p}v>${dateToSerial(edit)}</${p}v></${p}c>`;
  }
  if (typeof edit === 'object') return `<${p}c r="${ref}"${attrs}><${p}f>${escapeXml(edit.formula.replace(/^=/, ''))}</${p}f></${p}c>`;
  const text = escapeXml(encodeXString(edit));
  const t = /^\s|\s$/.test(edit) ? `<${p}t xml:space="preserve">${text}</${p}t>` : `<${p}t>${text}</${p}t>`;
  return `<${p}c r="${ref}"${attrs} t="inlineStr"><${p}is>${t}</${p}is></${p}c>`;
}

// The cells of a new sheet, as edits of an empty one
function rowsToEdits(name: string, rows: Row[]): EditedRows {
  const edits: EditedRows = new Map();
  rows.forEach((row, ri) => {
    const cells = new Map<number, CellEdit>();
    row.forEach((cell, ci) => {
      if (cell === null || cell === undefined) return;
      if (typeof cell !== 'object' || cell instanceof Date) {
        cells.set(ci, cell);
        return;
      }
      if (cell.hyperlink !== undefined || cell.comment !== undefined) {
        throw new Error(`Sheet "${name}": SheetEditor.addSheet does not write hyperlinks or notes; use SheetWriter.`);
      }
      // Rich text is written as plain text here
      const value = cell.richText && cell.value == null ? cell.richText.map(r => r.text).join('') : cell.value;
      cells.set(ci, { value, formula: cell.formula, style: cell.style });
    });
    if (cells.size) edits.set(ri + 1, cells);
  });
  return edits;
}

const WORKSHEET_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const WORKSHEET_CONTENT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml';

export class SheetEditor {
  private modifications = new Map<string, { rows: Row[], options: SheetOptions }>();
  private cellEdits = new Map<string, EditedRows>();
  private additions = new Map<string, Row[]>();
  private deletions = new Set<string>();

  // Appends `rows` after the last existing row of sheet `sheetName`.
  appendSheet(sheetName: string, rows: Row[], options: SheetOptions = {}) {
    this.modifications.set(sheetName, { rows, options });
  }

  // Sets cells of an existing sheet by address, e.g. { B2: 42, C2: { formula: 'B2*2' } }. A cell
  // keeps its style, so a date written over a date-formatted cell shows as a date. Formulas are
  // recalculated when Excel opens the file.
  setCells(sheetName: string, cells: Record<string, CellEdit>): this {
    let rows = this.cellEdits.get(sheetName);
    if (!rows) this.cellEdits.set(sheetName, rows = new Map());
    for (const [ref, edit] of Object.entries(cells)) {
      const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref);
      if (!m || parseInt(m[2], 10) < 1) throw new Error(`Invalid cell reference "${ref}".`);
      const r = parseInt(m[2], 10);
      if (!rows.has(r)) rows.set(r, new Map());
      rows.get(r)!.set(colIndex(m[1]), edit);
    }
    return this;
  }

  // Adds a sheet after the existing ones. Cells take values, formulas and styles; for hyperlinks,
  // notes and sheet options, write the workbook with SheetWriter.
  addSheet(sheetName: string, rows: Row[]): this {
    validateSheetName(sheetName, this.additions.keys());
    this.additions.set(sheetName, rows);
    return this;
  }

  // Removes a sheet. Names scoped to it go; other defined names that point at it become #REF!.
  // Formulas in other sheets that point at it are not rewritten.
  deleteSheet(sheetName: string): this {
    this.deletions.add(sheetName);
    return this;
  }

  edit(reader: RandomAccessReader): ReadableStream<Uint8Array> {
    const zipIn = new ZipRandomAccessParser(reader);
    const zipOut = new ZipStreamWriter();

    (async () => {
      try {
        await zipIn.parseCentralDirectory();

        // Resolve sheet paths up front via random access, independent of entry order.
        const readText = async (name: string) =>
          zipIn.has(name) ? this.readStreamToString(await zipIn.extractStream(name)) : '';
        const parts = await resolveWorkbookParts(readText);
        // Parts rewritten whole, read back by later steps
        const overlay = new Map<string, string>();
        const read = async (path: string) => overlay.get(path) ?? readText(path);
        const written = new Set<string>();
        const pathOf = (sheetName: string) => {
          if (this.deletions.has(sheetName)) throw new Error(`Sheet "${sheetName}" is being deleted.`);
          const path = parts.sheets.get(sheetName);
          if (!path || !zipIn.has(path)) throw new Error(`Sheet "${sheetName}" not found in workbook.`);
          return path;
        };

        const appendByPath = new Map<string, Row[]>();
        for (const [sheetName, mod] of this.modifications) appendByPath.set(pathOf(sheetName), mod.rows);
        const editsByPath = new Map<string, EditedRows>();
        for (const [sheetName, rows] of this.cellEdits) editsByPath.set(pathOf(sheetName), rows);

        for (const sheetName of this.deletions) await removeSheet(sheetName, parts, read, overlay, written);
        const added = new Map<string, EditedRows>();
        for (const [sheetName, rows] of this.additions) {
          const free = (path: string) => !zipIn.has(path) && !added.has(path);
          added.set(await insertSheet(sheetName, parts, read, overlay, free), rowsToEdits(sheetName, rows));
        }

        // Edited cells invalidate cached formula results
        if (editsByPath.size || added.size || this.deletions.size) {
          const { replace, drop } = await recalcOnOpen(read, parts, await read(parts.workbookPath));
          for (const [path, xml] of replace) overlay.set(path, xml);
          for (const path of drop) written.add(path);
        }

        // Styles gain the formats of restyled cells as the sheets stream, so they are written last
        let patcher: StylePatcher | undefined;
        const styled = [...editsByPath.values(), ...added.values()]
          .some(rows => [...rows.values()].some(cells => [...cells.values()].some(e => styleOf(e))));
        if (styled) {
          if (!parts.styles) throw new Error('Workbook has no styles part.');
          patcher = new StylePatcher(await read(parts.styles));
          written.add(parts.styles);
        }

        for (const [path, xml] of overlay) {
          if (written.has(path)) continue;
          await zipOut.addFile(path, new Response(xml).body!);
          written.add(path);
        }
        const decode = (bytes: ReadableStream<Uint8Array>) => bytes.pipeThrough(new TextDecoderStream() as any as TransformStream<Uint8Array, string>);
        const encode = (text: ReadableStream<string>) => text.pipeThrough(new TextEncoderStream() as any as TransformStream<string, Uint8Array>);
        if (added.size) {
          // A new sheet uses the workbook's namespace (Transitional or Strict)
          const root = /<((?:\w+:)?)workbook\b[^>]*>/.exec(await read(parts.workbookPath));
          const ns = (root && attr(root[0], root[1] ? `xmlns:${root[1].slice(0, -1)}` : 'xmlns')) ?? 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
          const empty = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${ns}"><sheetData/></worksheet>`;
          for (const [path, edits] of added) {
            await zipOut.addFile(path, encode(decode(new Response(empty).body!).pipeThrough(this.createEditTransform(edits, patcher))));
          }
        }

        for (const filename of zipIn.getFiles()) {
          if (written.has(filename)) continue;
          const rows = appendByPath.get(filename);
          const edits = editsByPath.get(filename);
          if (rows || edits) {
            let text = decode(await zipIn.extractStream(filename));
            if (edits) text = text.pipeThrough(this.createEditTransform(edits, patcher));
            if (rows) text = text.pipeThrough(this.createInjectTransform(rows));
            await zipOut.addFile(filename, encode(text));
          } else {
            // Untouched entries are copied without decompressing
            const record = zipIn.getRecord(filename);
            await zipOut.addCompressedFile(filename, await zipIn.extractRawStream(filename),
              record.uncompressedSize, record.compressedSize, record.crc, record.compressionMethod);
          }
        }
        if (patcher) await zipOut.addFile(parts.styles!, new Response(patcher.toXml()).body!);
        await zipOut.close();
      } catch (err) {
        zipOut.error(err);
      }
    })();

    return zipOut.stream;
  }

  // Streams a worksheet, rewriting edited cells as their rows pass by and adding rows and cells
  // that did not exist. Only one row at a time is held in memory.
  private createEditTransform(edits: EditedRows, patcher?: StylePatcher): TransformStream<string, string> {
    const pending = [...edits.keys()].sort((a, b) => a - b);
    let next = 0;
    let buffer = '';
    let p = '';              // namespace prefix of the sheet's elements
    let state: 'head' | 'rows' | 'tail' = 'head';
    // Shared formulas whose anchor cell was overwritten: their other cells get the formula written out
    const sharedAnchors = new Map<string, { text: string; row: number; col: number }>();
    const brokenShared = new Set<string>();

    // An edited cell; `existing` is the cell's current element, if any
    const cellXml = (ref: string, edit: CellEdit, existing?: RegExpMatchArray): string => {
      let attrs = existing ? existing[1].replace(/\sr="[^"]*"/, '') : '';
      const style = styleOf(edit);
      if (style) {
        const base = parseInt(/\ss="(\d+)"/.exec(attrs)?.[1] ?? '0', 10);
        attrs = `${attrs.replace(/\ss="[^"]*"/, '')} s="${patcher!.patch(base, style)}"`;
      }
      const content = contentOf(edit);
      if (content === KEEP) {
        return existing?.[2] !== undefined ? `<${p}c r="${ref}"${attrs}>${existing[2]}</${p}c>` : `<${p}c r="${ref}"${attrs}/>`;
      }
      // Keeps the style; drops the type and the metadata of rich values and dynamic arrays
      return editedCell(p, ref, attrs.replace(/\s(?:t|vm|cm)="[^"]*"/g, ''), content);
    };

    const newRow = (r: number) => {
      const cells = [...edits.get(r)!.entries()].sort((a, b) => a[0] - b[0])
        .map(([c, e]) => cellXml(colLetter(c) + r, e)).join('');
      return `<${p}row r="${r}">${cells}</${p}row>`;
    };
    const rowsBefore = (limit: number) => {
      let xml = '';
      for (; next < pending.length && pending[next] < limit; next++) xml += newRow(pending[next]);
      return xml;
    };

    const rewriteRow = (open: string, inner: string, r: number): string => {
      // A copy: the edits stay intact for another edit() call
      const rowEdits = edits.has(r) ? new Map(edits.get(r)) : undefined;
      if (rowEdits) next++;
      let changed = !!rowEdits;
      const cells: [number, string][] = [];
      let col = 0;
      const re = new RegExp(`<${p}c\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${p}c>)`, 'g');
      for (const m of inner.matchAll(re)) {
        const ref = /\sr="([A-Za-z]+)\d+"/.exec(m[1]);
        col = ref ? colIndex(ref[1]) : col;
        const edit = rowEdits?.get(col);
        const f = m[2] && new RegExp(`<${p}f\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${p}f>)`).exec(m[2]);
        const si = f && /\st="shared"/.test(f[1]) ? /\ssi="(\d+)"/.exec(f[1])?.[1] : undefined;
        let xml = m[0];
        const anchor = si !== undefined && f && f[2] !== undefined;
        if (anchor) sharedAnchors.set(si!, { text: unescapeXml(f![2]), row: r, col });
        if (rowEdits?.has(col)) {
          // the group's formula text lived in this cell
          if (anchor && contentOf(edit as CellEdit) !== KEEP) brokenShared.add(si!);
          xml = cellXml(colLetter(col) + r, edit as CellEdit, m);
          rowEdits.delete(col);
        } else if (si !== undefined && f && !anchor) {
          if (brokenShared.has(si)) {
            const a = sharedAnchors.get(si);
            if (a) {
              xml = xml.replace(f[0], `<${p}f>${escapeXml(shiftFormula(a.text, r - a.row, col - a.col))}</${p}f>`);
              changed = true;
            }
          }
        }
        cells.push([col, xml]);
        col++;
      }
      for (const [c, e] of rowEdits ?? []) cells.push([c, cellXml(colLetter(c) + r, e)]);
      if (!changed) return `${open}${inner}</${p}row>`;
      cells.sort((a, b) => a[0] - b[0]);
      return `${open.replace(/\sspans="[^"]*"/, '').replace(/\/>$/, '>')}${cells.map(c => c[1]).join('')}</${p}row>`;
    };

    return new TransformStream({
      transform(chunk, controller) {
        buffer += chunk;
        let out = '';
        while (true) {
          if (state === 'head') {
            const m = /<((?:\w+:)?)sheetData\b[^>]*?(\/?)>/.exec(buffer);
            if (!m) break;
            p = m[1];
            if (m[2]) {
              // Empty sheet: <sheetData/> becomes a pair holding the new rows
              out += buffer.slice(0, m.index) + `<${p}sheetData>${rowsBefore(Infinity)}</${p}sheetData>`;
              buffer = buffer.slice(m.index + m[0].length);
              state = 'tail';
            } else {
              out += buffer.slice(0, m.index + m[0].length);
              buffer = buffer.slice(m.index + m[0].length);
              state = 'rows';
            }
          } else if (state === 'rows') {
            const m = new RegExp(`^\\s*(?:<${p}row\\b([^>]*?)(?:/>|>([\\s\\S]*?)</${p}row>)|</${p}sheetData>)`).exec(buffer);
            if (!m) break;
            if (m[0].trimStart().startsWith(`</${p}sheetData`)) {
              out += rowsBefore(Infinity) + m[0];
              state = 'tail';
            } else {
              const r = parseInt(/\sr="(\d+)"/.exec(m[1])?.[1] ?? '0', 10);
              out += rowsBefore(r);
              const open = m[0].slice(0, m[0].indexOf('>') + 1);
              // Rows are only parsed when they have edits or take part in a shared formula
              out += edits.has(r) || (m[2] ?? '').includes('t="shared"') ? rewriteRow(open.trimStart(), m[2] ?? '', r) : m[0];
            }
            buffer = buffer.slice(m[0].length);
          } else {
            out += buffer;
            buffer = '';
            break;
          }
        }
        if (out) controller.enqueue(out);
      },
      flush(controller) {
        if (state !== 'tail') {
          controller.error(new Error('Worksheet has no complete <sheetData> element.'));
          return;
        }
        if (buffer) controller.enqueue(buffer);
      },
    });
  }

  private createInjectTransform(rows: Row[]): TransformStream<string, string> {
    let buffer = '';
    let maxRow = 0;
    let done = false;
    const rowsXml = (): string => rows.map((row, ri) => {
      const rowNum = maxRow + ri + 1;
      const cellsXml = row.map((cell, ci) => {
        const colRef = colLetter(ci) + rowNum;
        const val = typeof cell === 'object' && cell !== null && 'value' in cell ? cell.value : cell;
        if (val === null || val === undefined) return `<c r="${colRef}"/>`;
        if (typeof val === 'boolean') return `<c r="${colRef}" t="b"><v>${val ? 1 : 0}</v></c>`;
        if (typeof val === 'number') return `<c r="${colRef}"><v>${val}</v></c>`;
        if (typeof val === 'string') {
          return `<c r="${colRef}" t="inlineStr"><is><t>${escapeXml(val)}</t></is></c>`;
        }
        return `<c r="${colRef}"/>`;
      }).join('');
      return `<row r="${rowNum}">${cellsXml}</row>`;
    }).join('');

    return new TransformStream({
      transform(chunk, controller) {
        if (done) { controller.enqueue(chunk); return; }
        buffer += chunk;

        for (const m of buffer.matchAll(/<(?:\w+:)?row\b[^>]*?\sr="(\d+)"/g)) {
          const r = parseInt(m[1], 10);
          if (r > maxRow) maxRow = r;
        }

        const close = /<\/(?:\w+:)?sheetData>|<((?:\w+:)?sheetData)\b[^>]*\/>/.exec(buffer);
        if (close) {
          const before = buffer.slice(0, close.index);
          const after = buffer.slice(close.index + close[0].length);
          // Self-closing <sheetData/> (empty sheet) becomes an open/close pair
          const tagName = close[1];
          const inject = tagName
            ? `${close[0].slice(0, -2)}>${rowsXml()}</${tagName}>`
            : `${rowsXml()}${close[0]}`;
          controller.enqueue(before + inject + after);
          buffer = '';
          done = true;
        } else {
          // Hold back from the last '<' so a split tag is never emitted half-scanned
          const cut = buffer.lastIndexOf('<');
          if (cut > 0) {
            controller.enqueue(buffer.slice(0, cut));
            buffer = buffer.slice(cut);
          }
        }
      },
      flush(controller) {
        if (!done) {
          controller.error(new Error('Worksheet has no <sheetData> element.'));
          return;
        }
        if (buffer.length > 0) controller.enqueue(buffer);
      }
    });
  }

  private async readStreamToString(stream: ReadableStream<Uint8Array>): Promise<string> {
    const reader = stream.pipeThrough(new TextDecoderStream() as any as TransformStream<Uint8Array, string>).getReader();
    let result = '';
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      result += value;
    }
    return result;
  }
}

const sheetTags = (wb: string) => [...wb.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)].map(m => m[0]);
const relTag = (rels: string, id: string) =>
  new RegExp(`<(?:\\w+:)?Relationship\\b[^>]*?\\sId="${escapeRe(id)}"[^>]*/>`).exec(rels)?.[0];
const overrideTag = (types: string, path: string) =>
  new RegExp(`<(?:\\w+:)?Override\\b[^>]*?PartName="/${escapeRe(path)}"[^>]*/>`).exec(types)?.[0];

async function removeSheet(name: string, parts: WorkbookParts, read: (path: string) => Promise<string>,
  overlay: Map<string, string>, dropped: Set<string>) {
  let wb = await read(parts.workbookPath);
  const tags = sheetTags(wb);
  const idx = tags.findIndex(t => attr(t, 'name') === name);
  if (idx < 0) throw new Error(`Sheet "${name}" not found in workbook.`);
  if (!tags.some((t, i) => i !== idx && !/\sstate="(?:hidden|veryHidden)"/.test(t))) {
    throw new Error(`Cannot delete sheet "${name}": a workbook needs a visible sheet.`);
  }
  const rId = attr(tags[idx], 'r:id') ?? attr(tags[idx], '\\w+:id');
  wb = wb.replace(tags[idx], '');

  // Names scoped to the sheet go and later sheets' scopes shift down; references to it become #REF!
  const quoted = new RegExp(`'${escapeRe(name.replace(/'/g, "''"))}'!`, 'gi');
  const plain = /^[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*$/.test(name)
    ? new RegExp(`(^|[^\\w.'!])${escapeRe(name)}!`, 'gi') : undefined;
  wb = wb.replace(/<((?:\w+:)?)definedName\b([^>]*)>([\s\S]*?)<\/\1definedName>/g, (whole, _p, attrs: string, text: string) => {
    const local = attr(attrs, 'localSheetId');
    if (local !== null) {
      const n = parseInt(local, 10);
      return n === idx ? '' : n > idx ? whole.replace(/\slocalSheetId="\d+"/, ` localSheetId="${n - 1}"`) : whole;
    }
    const formula = unescapeXml(text);
    let mapped = formula.replace(quoted, '#REF!');
    if (plain) mapped = mapped.replace(plain, '$1#REF!');
    return mapped === formula ? whole : whole.replace(`>${text}<`, `>${escapeXml(mapped)}<`);
  });
  wb = wb.replace(/<(?:\w+:)?workbookView\b[^>]*>/g, view => view.replace(/\s(activeTab|firstSheet)="(\d+)"/g, (_m, a: string, v: string) => {
    const n = parseInt(v, 10);
    return ` ${a}="${n > idx ? n - 1 : Math.min(n, tags.length - 2)}"`;
  }));
  overlay.set(parts.workbookPath, wb);

  const path = (await partRelationships(read, parts.workbookPath)).find(r => r.id === rId)?.path;
  const relsPath = relsPathOf(parts.workbookPath);
  const rels = await read(relsPath);
  overlay.set(relsPath, rels.replace(relTag(rels, rId ?? '') ?? '', ''));
  if (path) {
    // The sheet's own drawings, notes and tables stay in the package, unreferenced
    const types = await read('[Content_Types].xml');
    overlay.set('[Content_Types].xml', types.replace(overrideTag(types, path) ?? '', ''));
    dropped.add(path);
    dropped.add(relsPathOf(path));
  }
}

// Registers a new, empty worksheet part after the existing sheets; returns its path
async function insertSheet(name: string, parts: WorkbookParts, read: (path: string) => Promise<string>,
  overlay: Map<string, string>, free: (path: string) => boolean): Promise<string> {
  let wb = await read(parts.workbookPath);
  const tags = sheetTags(wb);
  validateSheetName(name, tags.map(t => attr(t, 'name') ?? ''));
  const dir = dirOf(parts.workbookPath);
  let n = 1;
  while (!free(`${dir}worksheets/sheet${n}.xml`)) n++;
  const path = `${dir}worksheets/sheet${n}.xml`;

  const relsPath = relsPathOf(parts.workbookPath);
  let rels = await read(relsPath);
  let k = 1;
  while (relTag(rels, `rId${k}`)) k++;
  // Strict OOXML files use other type URIs: copy them from an existing sheet
  const sheetRel = (await partRelationships(read, parts.workbookPath)).find(r => r.type.endsWith('/worksheet'));
  rels = rels.replace(/<\/((?:\w+:)?)Relationships>/, (close, p: string) =>
    `<${p}Relationship Id="rId${k}" Type="${sheetRel?.type ?? WORKSHEET_TYPE}" Target="worksheets/sheet${n}.xml"/>${close}`);
  overlay.set(relsPath, rels);

  const sheetId = Math.max(0, ...tags.map(t => parseInt(attr(t, 'sheetId') ?? '0', 10))) + 1;
  const rPrefix = /<(?:\w+:)?sheet\b[^>]*?\s(\w+):id="/.exec(wb)?.[1] ?? 'r';
  wb = wb.replace(/<\/((?:\w+:)?)sheets>/, (close, p: string) =>
    `<${p}sheet name="${escapeXml(name)}" sheetId="${sheetId}" ${rPrefix}:id="rId${k}"/>${close}`);
  overlay.set(parts.workbookPath, wb);

  const types = await read('[Content_Types].xml');
  const contentType = sheetRel && /ContentType="([^"]*)"/.exec(overrideTag(types, sheetRel.path) ?? '')?.[1];
  overlay.set('[Content_Types].xml', types.replace(/<\/((?:\w+:)?)Types>/, (close, p: string) =>
    `<${p}Override PartName="/${path}" ContentType="${contentType ?? WORKSHEET_CONTENT}"/>${close}`));
  return path;
}
