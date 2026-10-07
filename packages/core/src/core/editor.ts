import { RandomAccessReader } from './random-access';
import { ZipRandomAccessParser } from './zip-random-access';
import { Row, SheetOptions, CellValue } from './types';
import { ZipStreamWriter } from './zip-stream-writer';
import { resolveWorkbookParts, recalcOnOpen, colIndex, colLetter, dateToSerial, encodeXString, shiftFormula } from './utils';

// A new cell value, or a formula (leading "=" optional). null clears the value.
export type CellEdit = CellValue | { formula: string };

const escapeXml = (val: string) =>
  val.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const unescapeXml = (s: string) =>
  s.replace(/&(?:#x([0-9a-fA-F]+)|#(\d+)|(amp|lt|gt|quot|apos));/g, (m, hex, dec, name) =>
    hex ? String.fromCodePoint(parseInt(hex, 16)) : dec ? String.fromCodePoint(parseInt(dec, 10))
      : ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" } as Record<string, string>)[name] ?? m);

// The cell XML for an edit; `attrs` keeps the existing cell's style (s=...) and other attributes
function editedCell(p: string, ref: string, attrs: string, edit: CellEdit): string {
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

export class SheetEditor {
  private modifications = new Map<string, { rows: Row[], options: SheetOptions }>();
  private cellEdits = new Map<string, Map<number, Map<number, CellEdit>>>();

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
        const pathOf = (sheetName: string) => {
          const path = parts.sheets.get(sheetName);
          if (!path || !zipIn.has(path)) throw new Error(`Sheet "${sheetName}" not found in workbook.`);
          return path;
        };

        const appendByPath = new Map<string, Row[]>();
        for (const [sheetName, mod] of this.modifications) appendByPath.set(pathOf(sheetName), mod.rows);
        const editsByPath = new Map<string, Map<number, Map<number, CellEdit>>>();
        for (const [sheetName, rows] of this.cellEdits) editsByPath.set(pathOf(sheetName), rows);

        // Edited cells invalidate cached formula results
        const written = new Set<string>();
        if (editsByPath.size) {
          const { replace, drop } = await recalcOnOpen(readText, parts);
          for (const [path, xml] of replace) {
            await zipOut.addFile(path, new Response(xml).body!);
            written.add(path);
          }
          for (const path of drop) written.add(path);
        }

        for (const filename of zipIn.getFiles()) {
          if (written.has(filename)) continue;
          const rows = appendByPath.get(filename);
          const edits = editsByPath.get(filename);
          if (rows || edits) {
            let text = (await zipIn.extractStream(filename))
              .pipeThrough(new TextDecoderStream() as any as TransformStream<Uint8Array, string>);
            if (edits) text = text.pipeThrough(this.createEditTransform(edits));
            if (rows) text = text.pipeThrough(this.createInjectTransform(rows));
            await zipOut.addFile(filename, text.pipeThrough(new TextEncoderStream() as any as TransformStream<string, Uint8Array>));
          } else {
            // Untouched entries are copied without decompressing
            const record = zipIn.getRecord(filename);
            await zipOut.addCompressedFile(filename, await zipIn.extractRawStream(filename),
              record.uncompressedSize, record.compressedSize, record.crc, record.compressionMethod);
          }
        }
        await zipOut.close();
      } catch (err) {
        zipOut.error(err);
      }
    })();

    return zipOut.stream;
  }

  // Streams a worksheet, rewriting edited cells as their rows pass by and adding rows and cells
  // that did not exist. Only one row at a time is held in memory.
  private createEditTransform(edits: Map<number, Map<number, CellEdit>>): TransformStream<string, string> {
    const pending = [...edits.keys()].sort((a, b) => a - b);
    let next = 0;
    let buffer = '';
    let p = '';              // namespace prefix of the sheet's elements
    let state: 'head' | 'rows' | 'tail' = 'head';
    // Shared formulas whose anchor cell was overwritten: their other cells get the formula written out
    const sharedAnchors = new Map<string, { text: string; row: number; col: number }>();
    const brokenShared = new Set<string>();

    const newRow = (r: number) => {
      const cells = [...edits.get(r)!.entries()].sort((a, b) => a[0] - b[0])
        .map(([c, e]) => editedCell(p, colLetter(c) + r, '', e)).join('');
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
          if (anchor) brokenShared.add(si!); // the group's formula text lived in this cell
          // Keeps the style; drops the type and the metadata of rich values and dynamic arrays
          const attrs = m[1].replace(/\s(?:r|t|vm|cm)="[^"]*"/g, '');
          xml = editedCell(p, colLetter(col) + r, attrs, edit as CellEdit);
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
      for (const [c, e] of rowEdits ?? []) cells.push([c, editedCell(p, colLetter(c) + r, '', e)]);
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
