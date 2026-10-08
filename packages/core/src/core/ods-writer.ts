// Writer for OpenDocument spreadsheets (.ods). Rows stream into content.xml one at a time.
// Writes values, formulas, dates, merged cells, column widths, frozen panes, hidden sheets and
// document properties; styles and the other .xlsx sheet options are left out.

import { ZipStreamWriter, crc32 } from './zip-stream-writer';
import { ODS_MIMETYPE } from './ods';
import { colIndex, validateSheetName } from './utils';
import type { CellValue, Row, StyledCell, WorkbookProperties } from './types';

export interface OdsSheetOptions {
  columnWidths?: number[];      // in characters, as in SheetOptions
  mergeCells?: string[];        // e.g. ["A1:C1"]
  freezePanes?: { row?: number; col?: number };
  state?: 'visible' | 'hidden';
}

export interface OdsWriterOptions {
  properties?: WorkbookProperties;
}

function escapeXml(val: unknown): string {
  // XML 1.0 cannot hold most control characters at all, so they are dropped
  return String(val).replace(/[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

const isStyledCell = (v: CellValue | StyledCell): v is StyledCell =>
  v !== null && typeof v === 'object' && !(v instanceof Date);

// Excel syntax to OpenFormula: "SUM(A1:B2,Sheet2!C3)" -> "of:=SUM([.A1:.B2];[$Sheet2.C3])"
const REF = /((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?(?![\w(!])/y;
export function toOpenFormula(formula: string): string {
  const f = formula.replace(/^=/, '');
  let out = '';
  for (let i = 0; i < f.length;) {
    const c = f[i];
    if (c === '"') { // string literal, copied as is
      let j = i + 1;
      while (j < f.length && !(f[j] === '"' && f[j + 1] !== '"')) j += f[j] === '"' ? 2 : 1;
      out += f.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    if (c === ',') { out += ';'; i++; continue; }
    // A reference starts where an identifier could not be continuing
    if (!/[\w.$]/.test(f[i - 1] ?? '')) {
      REF.lastIndex = i;
      const m = REF.exec(f);
      if (m) {
        const sheet = m[1] ? `$${m[1].slice(0, -1)}` : '';
        out += `[${sheet}.${m[2]}${m[3] ? `:${sheet}.${m[3]}` : ''}]`;
        i = REF.lastIndex;
        continue;
      }
    }
    out += c;
    i++;
  }
  return `of:=${out}`;
}

// Paragraphs per line, and runs of spaces as <text:s/>, since ODF collapses whitespace
function textXml(s: string): string {
  return s.split('\n').map(line => `<text:p>${escapeXml(line)
    .replace(/\t/g, '<text:tab/>')
    .replace(/^ | {2,}/g, m => m === ' ' ? '<text:s/>' : ` <text:s text:c="${m.length - 1}"/>`)}</text:p>`).join('');
}

const dateValue = (d: Date) => d.toISOString().slice(0, 19); // UTC, like SheetWriter
const hasTime = (d: Date) => d.getTime() % 86400000 !== 0;

function cellXml(cell: CellValue | StyledCell | undefined, span: string, covered: boolean): string {
  if (covered) return '<table:covered-table-cell/>';
  const value = isStyledCell(cell!) ? cell.value : cell;
  const formula = isStyledCell(cell!) && cell.formula ? ` table:formula="${escapeXml(toOpenFormula(cell.formula))}"` : '';
  if (value === null || value === undefined || (typeof value === 'number' && !isFinite(value))) {
    return formula || span ? `<table:table-cell${formula}${span}/>` : '<table:table-cell/>';
  }
  if (typeof value === 'number') {
    return `<table:table-cell${formula} office:value-type="float" office:value="${value}"${span}><text:p>${value}</text:p></table:table-cell>`;
  }
  if (typeof value === 'boolean') {
    return `<table:table-cell table:style-name="ceBool"${formula} office:value-type="boolean" office:boolean-value="${value}"${span}><text:p>${value ? 'TRUE' : 'FALSE'}</text:p></table:table-cell>`;
  }
  if (value instanceof Date) {
    const v = dateValue(value);
    const shown = hasTime(value) ? v.replace('T', ' ') : v.slice(0, 10);
    return `<table:table-cell table:style-name="${hasTime(value) ? 'ceDateTime' : 'ceDate'}"${formula} office:value-type="date" office:date-value="${v}"${span}><text:p>${shown}</text:p></table:table-cell>`;
  }
  return `<table:table-cell${formula} office:value-type="string"${span}>${textXml(String(value))}</table:table-cell>`;
}

const NS = 'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" '
  + 'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
  + 'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0" '
  + 'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2" xmlns:dc="http://purl.org/dc/elements/1.1/" '
  + 'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" office:version="1.3"';

interface Sheet { name: string; rows: Row[] | AsyncIterable<Row>; options: OdsSheetOptions }

export class OdsWriter {
  private sheets: Sheet[] = [];

  constructor(private readonly writerOptions: OdsWriterOptions = {}) {}

  addSheet(name: string, rows: Row[] | AsyncIterable<Row>, options: OdsSheetOptions = {}): this {
    validateSheetName(name, this.sheets.map(s => s.name));
    for (const range of options.mergeCells ?? []) parseRange(range);
    this.sheets.push({ name, rows, options });
    return this;
  }

  write(): ReadableStream<Uint8Array> {
    if (this.sheets.length === 0) this.addSheet('Sheet1', []);
    if (this.sheets.every(s => s.options.state === 'hidden')) throw new Error('A workbook needs at least one visible sheet');
    const zip = new ZipStreamWriter();
    const encoder = new TextEncoder();
    const text = (s: string) => new ReadableStream<Uint8Array>({ start(c) { c.enqueue(encoder.encode(s)); c.close(); } });

    (async () => {
      try {
        // The mimetype entry comes first and uncompressed, so tools can sniff the type
        const mime = encoder.encode(ODS_MIMETYPE);
        await zip.addCompressedFile('mimetype', new ReadableStream({ start(c) { c.enqueue(mime); c.close(); } }), mime.length, mime.length, crc32(mime), 0);
        await zip.addFile('content.xml', this.contentStream());
        await zip.addFile('styles.xml', text(`<?xml version="1.0" encoding="UTF-8"?><office:document-styles ${NS}><office:styles/></office:document-styles>`));
        await zip.addFile('meta.xml', text(this.metaXml()));
        await zip.addFile('settings.xml', text(this.settingsXml()));
        await zip.addFile('META-INF/manifest.xml', text('<?xml version="1.0" encoding="UTF-8"?>'
          + '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
          + `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${ODS_MIMETYPE}"/>`
          + ['content.xml', 'styles.xml', 'meta.xml', 'settings.xml'].map(p => `<manifest:file-entry manifest:full-path="${p}" manifest:media-type="text/xml"/>`).join('')
          + '</manifest:manifest>'));
        await zip.close();
      } catch (err) {
        zip.error(err);
      }
    })();
    return zip.stream;
  }

  private contentStream(): ReadableStream<Uint8Array> {
    const encoder = new TextEncoder();
    const sheets = this.sheets;
    async function* parts(): AsyncGenerator<string> {
      const columnStyles = new Map<number, string>();
      sheets.forEach(s => s.options.columnWidths?.forEach(w => { if (w > 0 && !columnStyles.has(w)) columnStyles.set(w, `co${columnStyles.size + 1}`); }));
      yield `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NS}><office:automatic-styles>`
        + '<number:date-style style:name="NDate"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>'
        + '<number:date-style style:name="NDateTime"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/><number:text> </number:text><number:hours number:style="long"/><number:text>:</number:text><number:minutes number:style="long"/><number:text>:</number:text><number:seconds number:style="long"/></number:date-style>'
        + '<number:boolean-style style:name="NBool"><number:boolean/></number:boolean-style>'
        + '<style:style style:name="ceBool" style:family="table-cell" style:data-style-name="NBool"/>'
        + '<style:style style:name="ceDate" style:family="table-cell" style:data-style-name="NDate"/>'
        + '<style:style style:name="ceDateTime" style:family="table-cell" style:data-style-name="NDateTime"/>'
        + '<style:style style:name="taVisible" style:family="table"><style:table-properties table:display="true"/></style:style>'
        + '<style:style style:name="taHidden" style:family="table"><style:table-properties table:display="false"/></style:style>'
        // Excel's width unit is about 7 pixels at 96 dpi, plus 5 pixels of padding
        + [...columnStyles].map(([w, name]) => `<style:style style:name="${name}" style:family="table-column"><style:table-column-properties style:column-width="${((w * 7 + 5) / 96).toFixed(4)}in"/></style:style>`).join('')
        + '</office:automatic-styles><office:body><office:spreadsheet>';

      for (const sheet of sheets) {
        yield `<table:table table:name="${escapeXml(sheet.name)}" table:style-name="${sheet.options.state === 'hidden' ? 'taHidden' : 'taVisible'}">`
          + (sheet.options.columnWidths ?? []).map(w => w > 0 ? `<table:table-column table:style-name="${columnStyles.get(w)}"/>` : '<table:table-column/>').join('');

        // Merges: spans go on the top-left cell, and the cells they cover are written as covered
        const merges = (sheet.options.mergeCells ?? []).map(parseRange);
        const lastMergeRow = Math.max(-1, ...merges.map(m => m.r2));
        const mergeAt = (r: number, c: number) => merges.find(m => m.r1 === r && m.c1 === c);
        const covered = (r: number, c: number) => merges.some(m => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2 && (r !== m.r1 || c !== m.c1));

        let r = 0, batch = '';
        const writeRow = (row: Row) => {
          const width = Math.max(row.length, ...merges.filter(m => r >= m.r1 && r <= m.r2).map(m => m.c2 + 1));
          let xml = '<table:table-row>';
          for (let c = 0; c < width; c++) {
            const m = mergeAt(r, c);
            const span = m ? ` table:number-columns-spanned="${m.c2 - m.c1 + 1}" table:number-rows-spanned="${m.r2 - m.r1 + 1}"` : '';
            xml += cellXml(row[c], span, covered(r, c));
          }
          r++;
          return xml + (width ? '' : '<table:table-cell/>') + '</table:table-row>';
        };
        for await (const row of sheet.rows) {
          batch += writeRow(row);
          if (batch.length > 1 << 16) { yield batch; batch = ''; }
        }
        while (r <= lastMergeRow) batch += writeRow([]); // merges below the last row still need their cells
        yield batch + '</table:table>';
      }
      yield '</office:spreadsheet></office:body></office:document-content>';
    }
    const it = parts();
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        const { done, value } = await it.next();
        if (done) controller.close(); else controller.enqueue(encoder.encode(value));
      },
      async cancel() { await it.return(undefined); },
    });
  }

  private metaXml(): string {
    const p = this.writerOptions.properties ?? {};
    const el = (tag: string, v?: string) => v ? `<${tag}>${escapeXml(v)}</${tag}>` : '';
    return `<?xml version="1.0" encoding="UTF-8"?><office:document-meta ${NS}><office:meta>`
      + '<meta:generator>XlsxFlow</meta:generator>'
      + el('dc:title', p.title) + el('dc:subject', p.subject) + el('meta:initial-creator', p.creator) + el('dc:creator', p.creator)
      + el('meta:keyword', p.keywords) + el('dc:description', p.description)
      + el('meta:creation-date', (p.created ?? new Date()).toISOString().slice(0, 19))
      + '</office:meta></office:document-meta>';
  }

  // Frozen panes are view settings: split mode 2 freezes, the positions say where
  private settingsXml(): string {
    const item = (name: string, type: string, value: string | number) => `<config:config-item config:name="${name}" config:type="${type}">${value}</config:config-item>`;
    const tables = this.sheets.filter(s => s.options.freezePanes?.row || s.options.freezePanes?.col).map(s => {
      const rows = s.options.freezePanes!.row ?? 0, cols = s.options.freezePanes!.col ?? 0;
      return `<config:config-item-map-entry config:name="${escapeXml(s.name)}">`
        + item('HorizontalSplitMode', 'short', cols ? 2 : 0) + item('VerticalSplitMode', 'short', rows ? 2 : 0)
        + item('HorizontalSplitPosition', 'int', cols) + item('VerticalSplitPosition', 'int', rows)
        + item('ActiveSplitRange', 'short', 2) + item('PositionLeft', 'int', 0) + item('PositionRight', 'int', cols)
        + item('PositionTop', 'int', 0) + item('PositionBottom', 'int', rows)
        + '</config:config-item-map-entry>';
    }).join('');
    const first = this.sheets.find(s => s.options.state !== 'hidden')!;
    return `<?xml version="1.0" encoding="UTF-8"?><office:document-settings ${NS}><office:settings>`
      + '<config:config-item-set config:name="ooo:view-settings"><config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>'
      + item('ViewId', 'string', 'view1') + item('ActiveTable', 'string', escapeXml(first.name))
      + `<config:config-item-map-named config:name="Tables">${tables}</config:config-item-map-named>`
      + '</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>';
  }
}

function parseRange(range: string) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(range);
  if (!m) throw new Error(`Invalid merge range "${range}"`);
  const [c1, r1, c2, r2] = [colIndex(m[1]), +m[2] - 1, colIndex(m[3]), +m[4] - 1];
  return { r1: Math.min(r1, r2), r2: Math.max(r1, r2), c1: Math.min(c1, c2), c2: Math.max(c1, c2) };
}
