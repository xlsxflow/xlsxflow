import { Row, StyledCell, CellValue, SheetOptions, SheetImage, CellComment, TableOptions, PageSetup, RichTextRun, WorkbookProperties, DefinedName } from './types';
import { StyleEngine, fontXml } from './style-engine';
import { ConditionalFormatter } from './conditional-formatter';
import { FormulaEngine } from './formula-engine';
import { ZipStreamWriter } from './zip-stream-writer';
import { encodeXString, colLetter, colIndex, dateToSerial, validateSheetName } from './utils';
import { imageInfo, drawingXml, type ImageInfo } from './image';

function escapeXml(val: unknown): string {
  return String(val)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function isStyledCell(v: CellValue | StyledCell): v is StyledCell {
  return typeof v === 'object' && v !== null && 'value' in v;
}

const DAY_MS = 86400000;

// "A1:C10" -> "$A$1:$C$10"
const absoluteRef = (ref: string) => ref.replace(/([A-Za-z]+)(\d+)/g, '$$$1$$$2');

function stringXml(s: string): string {
  const text = escapeXml(encodeXString(s));
  // Leading/trailing whitespace is dropped by Excel unless marked as preserved
  return /^\s|\s$/.test(s) ? `<t xml:space="preserve">${text}</t>` : `<t>${text}</t>`;
}

// Rich text runs; plain text is one unformatted run
const runsXml = (text: string | RichTextRun[]) => typeof text === 'string' ? `<r>${stringXml(text)}</r>`
  : text.map(run => `<r>${run.font ? `<rPr>${fontXml(run.font, 'rFont')}</rPr>` : ''}${stringXml(run.text)}</r>`).join('');

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

// What a sheet produced while streaming, for the parts written after it
interface SheetParts {
  links: string[];                                         // external hyperlink targets, rId1..n
  comments: { ref: string; row: number; col: number; comment: CellComment }[];
}

// Excel's legacy 16-bit sheet password hash (ECMA-376 Part 4, 14.7.1)
export function legacyPasswordHash(password: string): string {
  let hash = 0;
  for (let i = password.length - 1; i >= 0; i--) {
    hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
    hash ^= password.charCodeAt(i);
  }
  hash = ((hash >> 14) & 1) | ((hash << 1) & 0x7fff);
  return (hash ^ password.length ^ 0xce4b).toString(16).toUpperCase();
}

const quoteSheet = (name: string) => `'${name.replace(/'/g, "''")}'`;

export interface WriterOptions {
  // Store strings once in xl/sharedStrings.xml instead of inline. Smaller files when values repeat,
  // but every distinct string is kept in memory until the workbook is finished.
  sharedStrings?: boolean;
  properties?: WorkbookProperties;
  definedNames?: DefinedName[];
}

const W3CDTF = (d: Date) => d.toISOString().replace(/\.\d{3}Z$/, 'Z');
const propXml = (tag: string, v?: string) => v === undefined ? '' : `<${tag}>${escapeXml(v)}</${tag}>`;

function corePropsXml(p: WorkbookProperties): string {
  const created = p.created ? `<dcterms:created xsi:type="dcterms:W3CDTF">${W3CDTF(p.created)}</dcterms:created>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
    propXml('dc:title', p.title) + propXml('dc:subject', p.subject) + propXml('dc:creator', p.creator) + propXml('cp:keywords', p.keywords) +
    propXml('dc:description', p.description) + propXml('cp:category', p.category) + created + '</cp:coreProperties>';
}

function appPropsXml(p: WorkbookProperties): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>XlsxFlow</Application>${propXml('Manager', p.manager)}${propXml('Company', p.company)}</Properties>`;
}

// A name Excel accepts: letters, digits, _ . and \, not a cell reference (A1, R1C1) and not reserved
function validateDefinedName(name: string) {
  if (!/^[A-Za-z_\\][A-Za-z0-9_.\\]*$/.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^([Rr]\d*)?([Cc]\d*)?$/.test(name) || /^_xlnm\./i.test(name)) {
    throw new Error(`Invalid defined name "${name}"`);
  }
}

// Builds a complete .xlsx Uint8Array from structured row data.
// Uses a pure JS ZIP writer (no dependencies) to pack OOXML parts.
export class SheetWriter {
  private styleEngine = new StyleEngine();
  private conditionalFormatter = new ConditionalFormatter(style => this.styleEngine.registerDxf(style));
  private formulaEngine = new FormulaEngine();
  private sheets: { name: string; rows: Row[] | AsyncIterable<Row>; options: SheetOptions }[] = [];
  private sharedStrings = new Map<string, number>();
  private sharedStringRefs = 0;
  // Embedded pictures by their data object, so a logo used on every sheet is stored once
  private media = new Map<Uint8Array | ArrayBuffer, { path: string; bytes: Uint8Array; info: ImageInfo }>();

  constructor(private writerOptions: WriterOptions = {}) {}

  addSheet(name: string, rows: Row[] | AsyncIterable<Row>, options: SheetOptions = {}): this {
    validateSheetName(name, this.sheets.map(s => s.name));
    const zoom = options.view?.zoom;
    if (zoom !== undefined && !(zoom >= 10 && zoom <= 400)) throw new Error(`Zoom must be between 10 and 400, got ${zoom}`);
    this.sheets.push({ name, rows, options });
    return this;
  }

  // Accept rows directly or via options, returning a standard ReadableStream
  write(rows?: Row[], options: SheetOptions = {}): ReadableStream<Uint8Array> {
    if (rows) {
      this.addSheet(options.name ?? 'Sheet1', rows, options);
    }
    if (this.sheets.length === 0) {
      this.addSheet('Sheet1', []);
    }
    if (this.activeSheet() < 0) throw new Error('A workbook needs at least one visible sheet');
    const scopes = new Set<string>();
    for (const n of this.writerOptions.definedNames ?? []) {
      validateDefinedName(n.name);
      if (n.sheet !== undefined && !this.sheets.some(s => s.name === n.sheet)) throw new Error(`Sheet "${n.sheet}" of defined name "${n.name}" not found`);
      const key = `${n.sheet ?? ''}!${n.name.toLowerCase()}`;
      if (scopes.has(key)) throw new Error(`Duplicate defined name "${n.name}"`);
      scopes.add(key);
    }
    const props = this.writerOptions.properties;

    const zip = new ZipStreamWriter();
    for (const img of this.sheets.flatMap(s => s.options.images ?? [])) {
      if (this.media.has(img.data)) continue;
      const bytes = img.data instanceof Uint8Array ? img.data : new Uint8Array(img.data);
      const info = imageInfo(bytes);
      this.media.set(img.data, { path: `media/image${this.media.size + 1}.${info.ext}`, bytes, info });
    }
    let drawings = 0;
    const commentSheets: number[] = [];

    // Push the build process to background so we can return the stream immediately
    (async () => {
      try {
        const tables = this.planTables();
        await this.streamString(zip, '_rels/.rels', this.buildRootRels());
        await this.streamString(zip, 'xl/workbook.xml', this.buildWorkbookXml());
        await this.streamString(zip, 'xl/_rels/workbook.xml.rels', this.buildWorkbookRels());
        if (props) {
          await this.streamString(zip, 'docProps/core.xml', corePropsXml(props));
          await this.streamString(zip, 'docProps/app.xml', appPropsXml(props));
        }

        for (let i = 0; i < this.sheets.length; i++) {
          const sheet = this.sheets[i];
          // Formulas resolve against the sheet they live in
          this.formulaEngine.clear();
          if (Array.isArray(sheet.rows)) {
            this.formulaEngine.loadData(sheet.rows.map(row => row.map(cell => {
              const v = !isStyledCell(cell) ? cell : cell.richText && cell.value == null ? cell.richText.map(r => r.text).join('') : cell.value;
              return v instanceof Date ? dateToSerial(v) : v;
            })));
            if (sheet.options.autoFitColumns) this.applyAutoFit(sheet.rows, sheet.options);
          }
          const parts: SheetParts = { links: [], comments: [] };
          const sheetTables = tables.filter(t => t.sheet === i);
          await zip.addFile(`xl/worksheets/sheet${i + 1}.xml`, this.buildWorksheetXmlStream(sheet.rows, sheet.options, parts, sheetTables.length));
          const images = sheet.options.images ?? [];
          const drawing = images.length ? ++drawings : 0;
          const rels = this.buildSheetRels(parts.links, drawing, parts.comments.length ? i + 1 : 0, sheetTables.map(t => t.id));
          if (rels) await this.streamString(zip, `xl/worksheets/_rels/sheet${i + 1}.xml.rels`, rels);
          if (drawing) await this.writeDrawing(zip, drawing, images);
          if (parts.comments.length) {
            commentSheets.push(i + 1);
            await this.streamString(zip, `xl/comments${i + 1}.xml`, commentsXml(parts.comments));
            await this.streamString(zip, `xl/drawings/vmlDrawing${i + 1}.vml`, vmlXml(parts.comments, i + 1));
          }
          for (const t of sheetTables) await this.streamString(zip, `xl/tables/table${t.id}.xml`, tableXml(t.table, t.id, t.columns));
        }
        // Written last: whether a streamed sheet has comments is only known after streaming it
        await this.streamString(zip, '[Content_Types].xml', this.buildContentTypes(drawings, commentSheets, tables.length));

        // styles.xml goes last: streamed (AsyncIterable) rows register styles while being written.
        // Entry order is irrelevant to readers, which use the central directory.
        await this.streamString(zip, 'xl/sharedStrings.xml', this.buildSharedStringsXml());
        await this.streamString(zip, 'xl/styles.xml', this.styleEngine.toXml());

        for (const { path, bytes } of this.media.values()) {
          await zip.addFile(`xl/${path}`, new ReadableStream({ start(c) { c.enqueue(bytes); c.close(); } }));
        }
        await zip.close();
      } catch (err) {
        zip.error(err);
      }
    })();

    return zip.stream;
  }

  // Tables get workbook-wide ids; their header names come from the options or the header row
  private planTables(): { sheet: number; id: number; table: TableOptions; columns: string[] }[] {
    const plan: { sheet: number; id: number; table: TableOptions; columns: string[] }[] = [];
    const names = new Set<string>();
    this.sheets.forEach((sheet, i) => {
      for (const table of sheet.options.tables ?? []) {
        if (!/^[A-Za-z_\\][A-Za-z0-9_.]*$/.test(table.name) || names.has(table.name.toLowerCase())) {
          throw new Error(`Invalid or duplicate table name "${table.name}".`);
        }
        names.add(table.name.toLowerCase());
        const m = /^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/.exec(table.ref.replace(/\$/g, ''));
        if (!m) throw new Error(`Invalid table range "${table.ref}".`);
        const first = colIndex(m[1]);
        const width = colIndex(m[3]) - first + 1;
        let columns = table.columns;
        if (!columns) {
          if (!Array.isArray(sheet.rows)) throw new Error(`Table "${table.name}": pass columns when rows are an AsyncIterable.`);
          const header = sheet.rows[parseInt(m[2], 10) - 1] ?? [];
          columns = Array.from({ length: width }, (_, k) => {
            const cell = header[first + k];
            const v = isStyledCell(cell) ? cell.value : cell;
            return v == null ? '' : String(v);
          });
        }
        if (columns.length !== width) throw new Error(`Table "${table.name}" has ${width} columns but ${columns.length} names.`);
        if (columns.some(c => !c) || new Set(columns.map(c => c.toLowerCase())).size !== columns.length) {
          throw new Error(`Table "${table.name}" needs unique, non-empty header names.`);
        }
        plan.push({ sheet: i, id: plan.length + 1, table, columns });
      }
    });
    return plan;
  }

  private async writeDrawing(zip: ZipStreamWriter, n: number, images: SheetImage[]) {
    const targets: string[] = [];
    const rIds = images.map(img => {
      const target = `../${this.media.get(img.data)!.path}`;
      let k = targets.indexOf(target);
      if (k < 0) k = targets.push(target) - 1;
      return `rId${k + 1}`;
    });
    await this.streamString(zip, `xl/drawings/drawing${n}.xml`, drawingXml(images, images.map(img => this.media.get(img.data)!.info), rIds));
    const rels = targets.map((t, k) =>
      `<Relationship Id="rId${k + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${t}"/>`).join('');
    await this.streamString(zip, `xl/drawings/_rels/drawing${n}.xml.rels`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`);
  }

  private applyAutoFit(rows: Row[], options: SheetOptions) {
    const colWidths: number[] = [];
    for (const row of rows) {
      row.forEach((cell, ci) => {
        const val = isStyledCell(cell) ? cell.value : cell;
        const str = val instanceof Date ? 'yyyy-mm-dd hh:mm:ss' : val !== null && val !== undefined ? String(val) : '';
        // Add a little padding (e.g. 1.2 multiplier) + base width
        const estWidth = Math.min(255, Math.max(10, str.length * 1.2));
        if (!colWidths[ci] || estWidth > colWidths[ci]) colWidths[ci] = estWidth;
      });
    }
    options.columnWidths = options.columnWidths || [];
    for (let ci = 0; ci < colWidths.length; ci++) {
      if (options.columnWidths[ci] === undefined && colWidths[ci] !== undefined) {
        options.columnWidths[ci] = Math.round(colWidths[ci] * 100) / 100;
      }
    }
  }

  private async streamString(zip: ZipStreamWriter, filename: string, content: string) {
    const encoder = new TextEncoder();
    const data = encoder.encode(content);
    const stream = new ReadableStream({
      start(c) { c.enqueue(data); c.close(); }
    });
    await zip.addFile(filename, stream);
  }

  // Pull-based so rows are only generated as fast as the ZIP consumer drains them.
  private buildWorksheetXmlStream(rows: Row[] | AsyncIterable<Row>, options: SheetOptions, parts: SheetParts, tables: number): ReadableStream<Uint8Array> {
    const encoder = new TextEncoder();
    const chunks = this.worksheetXmlChunks(rows, options, parts, tables);
    return new ReadableStream({
      async pull(controller) {
        const { done, value } = await chunks.next();
        if (done) controller.close();
        else controller.enqueue(encoder.encode(value));
      },
      async cancel() {
        await chunks.return(undefined);
      }
    });
  }

  // Hyperlink targets and comments are collected in `parts` for the parts written after the sheet
  private async *worksheetXmlChunks(rows: Row[] | AsyncIterable<Row>, options: SheetOptions, parts: SheetParts, tables: number): AsyncGenerator<string> {
    const links = parts.links;
    const hyperlinks: string[] = [];
    const colCount = Math.max(options.columnWidths?.length ?? 0, options.columns?.length ?? 0);
    let colWidths = '';
    for (let i = 0; i < colCount; i++) {
      const c = options.columns?.[i] ?? {};
      const width = c.width ?? options.columnWidths?.[i];
      if (width === undefined && !c.hidden && !c.outlineLevel) continue;
      colWidths += `<col min="${i + 1}" max="${i + 1}"${width !== undefined ? ` width="${escapeXml(width)}" customWidth="1"` : ''}` +
        `${c.hidden ? ' hidden="1"' : ''}${c.outlineLevel ? ` outlineLevel="${escapeXml(c.outlineLevel)}"` : ''}/>`;
    }
    const rowOptions = Object.entries(options.rows ?? {}).map(([r, o]) => [parseInt(r, 10), o] as const).sort((a, b) => a[0] - b[0]);
    const rowAttrs = new Map(rowOptions.map(([r, o]) =>
      [r, `${o.height !== undefined ? ` ht="${escapeXml(o.height)}" customHeight="1"` : ''}${o.hidden ? ' hidden="1"' : ''}${o.outlineLevel ? ` outlineLevel="${escapeXml(o.outlineLevel)}"` : ''}`]));
    let nextRowOption = 0;
    // Rows that only exist for their options (height, hidden, outline) and hold no cells
    const optionRowsBefore = (limit: number) => {
      let xml = '';
      for (; nextRowOption < rowOptions.length && rowOptions[nextRowOption][0] < limit; nextRowOption++) {
        xml += `    <row r="${rowOptions[nextRowOption][0]}"${rowAttrs.get(rowOptions[nextRowOption][0])}/>\n`;
      }
      return xml;
    };
    const outlineRow = Math.max(0, ...rowOptions.map(([, o]) => o.outlineLevel ?? 0));
    const outlineCol = Math.max(0, ...(options.columns ?? []).map(c => c?.outlineLevel ?? 0));
    const page = options.pageSetup;
    const fitToPage = page && (page.fitToWidth !== undefined || page.fitToHeight !== undefined);
    const sheetPr = options.tabColor || fitToPage
      ? `<sheetPr>${options.tabColor ? `<tabColor rgb="${escapeXml(options.tabColor)}"/>` : ''}${fitToPage ? '<pageSetUpPr fitToPage="1"/>' : ''}</sheetPr>`
      : '';
    const sheetFormatPr = outlineRow || outlineCol
      ? `<sheetFormatPr defaultRowHeight="15"${outlineRow ? ` outlineLevelRow="${outlineRow}"` : ''}${outlineCol ? ` outlineLevelCol="${outlineCol}"` : ''}/>`
      : '';

    let pane = '';
    if (options.freezePanes) {
      const { row = 0, col = 0 } = options.freezePanes;
      let activePane = 'bottomRight';
      if (row > 0 && col === 0) activePane = 'bottomLeft';
      if (row === 0 && col > 0) activePane = 'topRight';
      const topLeftCell = colLetter(col) + (row + 1);
      pane = `<pane ySplit="${row}" xSplit="${col}" topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/>`;
    }
    const view = options.view ?? {};
    const viewAttrs = (view.showGridLines === false ? ' showGridLines="0"' : '') + (view.showHeadings === false ? ' showRowColHeaders="0"' : '') +
      (view.rightToLeft ? ' rightToLeft="1"' : '') + (view.zoom !== undefined ? ` zoomScale="${Math.round(view.zoom)}"` : '');
    const sheetViews = pane || viewAttrs ? `<sheetViews><sheetView${viewAttrs} workbookViewId="0">${pane}</sheetView></sheetViews>` : '';

    let header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
    header += `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n`;
    header += `  ${sheetPr}${sheetViews}${sheetFormatPr}\n`;
    header += `  ${colWidths ? `<cols>${colWidths}</cols>` : ''}\n`;
    header += `  <sheetData>\n`;
    yield header;

    let ri = 0;
    
    // Handle both Array and AsyncIterable
    const iterator = Symbol.asyncIterator in rows 
      ? (rows as AsyncIterable<Row>)[Symbol.asyncIterator]()
      : (async function*() { for (const r of rows as Row[]) yield r; })();

    let chunkStr = '';
    while (true) {
      const { done, value: row } = await iterator.next();
      if (done) break;

      const rowNum = ri + 1;
      chunkStr += optionRowsBefore(rowNum);
      if (rowOptions[nextRowOption]?.[0] === rowNum) nextRowOption++;
      const cellsXml = row.map((cell, ci) => {
        const colRef = colLetter(ci) + rowNum;
        const styledCell: StyledCell = isStyledCell(cell) ? cell : { value: cell };
        const val = styledCell.value;
        let cellStyle = styledCell.style;
        if (val instanceof Date && !cellStyle?.numFmt) {
          // A date needs a date format, or Excel shows the bare serial number
          cellStyle = { ...cellStyle, numFmt: val.getTime() % DAY_MS === 0 ? 'yyyy-mm-dd' : 'yyyy-mm-dd hh:mm:ss' };
        }
        const style = cellStyle ? this.styleEngine.registerStyle(cellStyle) : 0;
        const sAttr = style > 0 ? ` s="${style}"` : '';

        if (styledCell.comment !== undefined) {
          const comment = typeof styledCell.comment === 'string' ? { text: styledCell.comment } : styledCell.comment;
          parts.comments.push({ ref: colRef, row: rowNum - 1, col: ci, comment });
        }

        if (styledCell.hyperlink) {
          if (styledCell.hyperlink.startsWith('#')) {
            hyperlinks.push(`<hyperlink ref="${colRef}" location="${escapeXml(styledCell.hyperlink.slice(1))}"/>`);
          } else {
            links.push(styledCell.hyperlink);
            hyperlinks.push(`<hyperlink ref="${colRef}" r:id="rId${links.length}"/>`);
          }
        }

        if (styledCell.formula) {
          const result = this.formulaEngine.evaluate(styledCell.formula);
          let tAttr = '';
          let cachedVal = '';
          if (typeof result === 'number' && isFinite(result)) cachedVal = `<v>${result}</v>`;
          else if (typeof result === 'boolean') { tAttr = ' t="b"'; cachedVal = `<v>${result ? 1 : 0}</v>`; }
          else if (typeof result === 'string') {
            tAttr = result.startsWith('#') ? ' t="e"' : ' t="str"';
            cachedVal = `<v>${escapeXml(encodeXString(result))}</v>`;
          }
          return `<c r="${colRef}"${tAttr}${sAttr}><f>${escapeXml(styledCell.formula.replace(/^=/, ''))}</f>${cachedVal}</c>`;
        }

        if (styledCell.richText) {
          // Rich text is always inline, even with sharedStrings on
          return `<c r="${colRef}" t="inlineStr"${sAttr}><is>${runsXml(styledCell.richText)}</is></c>`;
        }
        if (val === null || val === undefined) return `<c r="${colRef}"${sAttr}/>`;
        if (typeof val === 'boolean') return `<c r="${colRef}" t="b"${sAttr}><v>${val ? 1 : 0}</v></c>`;
        if (typeof val === 'number') {
          // NaN/Infinity have no representation in a cell
          return isFinite(val) ? `<c r="${colRef}"${sAttr}><v>${val}</v></c>` : `<c r="${colRef}" t="e"${sAttr}><v>#NUM!</v></c>`;
        }
        if (val instanceof Date) {
          if (isNaN(val.getTime())) throw new Error(`Invalid Date in cell ${colRef}`);
          return `<c r="${colRef}"${sAttr}><v>${dateToSerial(val)}</v></c>`;
        }
        if (typeof val === 'string') {
          if (this.writerOptions.sharedStrings) {
            let index = this.sharedStrings.get(val);
            if (index === undefined) this.sharedStrings.set(val, index = this.sharedStrings.size);
            this.sharedStringRefs++;
            return `<c r="${colRef}" t="s"${sAttr}><v>${index}</v></c>`;
          }
          return `<c r="${colRef}" t="inlineStr"${sAttr}><is>${stringXml(val)}</is></c>`;
        }
        return `<c r="${colRef}"${sAttr}/>`;
      }).join('');
      
      chunkStr += `    <row r="${rowNum}"${rowAttrs.get(rowNum) ?? ''}>${cellsXml}</row>\n`;
      ri++;

      // Flush every ~64KB of string data to keep O(1) memory while avoiding chunk overhead
      if (chunkStr.length > 65536) {
        yield chunkStr;
        chunkStr = '';
      }
    }
    
    chunkStr += optionRowsBefore(Infinity);
    if (chunkStr.length > 0) {
      yield chunkStr;
    }

    let footer = `  </sheetData>\n`;
    if (options.protection) {
      const p = options.protection === true ? {} : options.protection;
      // In the file, an action set to 1 is locked; allowing it means writing 0
      const allowed = (['formatCells', 'formatColumns', 'formatRows', 'insertRows', 'insertColumns', 'deleteRows', 'deleteColumns', 'sort', 'autoFilter'] as const)
        .filter(k => p[k]).map(k => ` ${k}="0"`).join('');
      footer += `  <sheetProtection${p.password ? ` password="${legacyPasswordHash(p.password)}"` : ''} sheet="1" objects="1" scenarios="1"${allowed}/>\n`;
    }

    if (options.autoFilter) footer += `  <autoFilter ref="${escapeXml(options.autoFilter)}"/>\n`;

    if (options.mergeCells && options.mergeCells.length > 0) {
      const merges = options.mergeCells.map(ref => `<mergeCell ref="${escapeXml(ref)}"/>`).join('');
      footer += `  <mergeCells count="${options.mergeCells.length}">${merges}</mergeCells>\n`;
    }

    if (options.conditionalFormats) {
      footer += this.conditionalFormatter.toXml(options.conditionalFormats) + '\n';
    }

    if (options.dataValidations && options.dataValidations.length > 0) {
      const dvs = options.dataValidations.map(dv => {
        let attr = `sqref="${escapeXml(dv.sqref)}"`;
        if (dv.type) attr += ` type="${escapeXml(dv.type)}"`;
        if (dv.allowBlank !== undefined) attr += ` allowBlank="${dv.allowBlank ? 1 : 0}"`;
        if (dv.showInputMessage !== undefined) attr += ` showInputMessage="${dv.showInputMessage ? 1 : 0}"`;
        if (dv.errorStyle) attr += ` errorStyle="${escapeXml(dv.errorStyle)}"`;
        if (dv.operator) attr += ` operator="${escapeXml(dv.operator)}"`;
        if (dv.showErrorMessage !== undefined || dv.error) attr += ` showErrorMessage="${dv.showErrorMessage ?? true ? 1 : 0}"`;
        if (dv.prompt && dv.showInputMessage === undefined) attr += ' showInputMessage="1"';
        for (const k of ['errorTitle', 'error', 'promptTitle', 'prompt'] as const) {
          if (dv[k]) attr += ` ${k}="${escapeXml(dv[k]!)}"`;
        }
        let inner = '';
        if (dv.formula1) inner += `<formula1>${escapeXml(dv.formula1)}</formula1>`;
        if (dv.formula2) inner += `<formula2>${escapeXml(dv.formula2)}</formula2>`;
        return `<dataValidation ${attr}>${inner}</dataValidation>`;
      }).join('');
      footer += `  <dataValidations count="${options.dataValidations.length}">${dvs}</dataValidations>\n`;
    }

    if (hyperlinks.length) footer += `  <hyperlinks>${hyperlinks.join('')}</hyperlinks>\n`;
    footer += pageSetupXml(page);
    if (options.images?.length) footer += `  <drawing r:id="rIdDrawing"/>\n`;
    if (parts.comments.length) footer += `  <legacyDrawing r:id="rIdVml"/>\n`;
    if (tables) footer += `  <tableParts count="${tables}">${Array.from({ length: tables }, (_, k) => `<tablePart r:id="rIdTable${k + 1}"/>`).join('')}</tableParts>\n`;

    footer += `</worksheet>`;
    yield footer;
  }

  // Empty unless the sharedStrings option is on (strings are written inline by default)
  private buildSharedStringsXml(): string {
    let items = '';
    for (const s of this.sharedStrings.keys()) items += `<si>${stringXml(s)}</si>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${this.sharedStringRefs}" uniqueCount="${this.sharedStrings.size}">${items}</sst>`;
  }

  // '' when the sheet has no relationships
  private buildSheetRels(links: string[], drawing: number, comments: number, tableIds: number[]): string {
    let rels = links.map((target, i) =>
      `<Relationship Id="rId${i + 1}" Type="${REL}/hyperlink" Target="${escapeXml(target)}" TargetMode="External"/>`
    ).join('');
    if (drawing) rels += `<Relationship Id="rIdDrawing" Type="${REL}/drawing" Target="../drawings/drawing${drawing}.xml"/>`;
    if (comments) {
      rels += `<Relationship Id="rIdComments" Type="${REL}/comments" Target="../comments${comments}.xml"/>` +
        `<Relationship Id="rIdVml" Type="${REL}/vmlDrawing" Target="../drawings/vmlDrawing${comments}.vml"/>`;
    }
    tableIds.forEach((id, k) => { rels += `<Relationship Id="rIdTable${k + 1}" Type="${REL}/table" Target="../tables/table${id}.xml"/>`; });
    if (!rels) return '';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
  }

  private buildWorkbookXml(): string {
    const sheetsXml = this.sheets.map((s, i) =>
      `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}"${s.options.state && s.options.state !== 'visible' ? ` state="${s.options.state}"` : ''} r:id="rId${i + 1}"/>`
    ).join('');
    const userNames = (this.writerOptions.definedNames ?? []).map(n => {
      const local = n.sheet === undefined ? '' : ` localSheetId="${this.sheets.findIndex(s => s.name === n.sheet)}"`;
      const ref = n.ref.startsWith('=') ? n.ref.slice(1) : n.ref;
      return `<definedName name="${escapeXml(n.name)}"${n.comment ? ` comment="${escapeXml(n.comment)}"` : ''}${local}${n.hidden ? ' hidden="1"' : ''}>${escapeXml(ref)}</definedName>`;
    }).join('');
    const active = this.activeSheet();
    // Excel tracks each sheet's filter range in a hidden defined name
    const filterNames = this.sheets.map((s, i) => s.options.autoFilter
      ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${escapeXml(`'${s.name.replace(/'/g, "''")}'!${absoluteRef(s.options.autoFilter)}`)}</definedName>`
      : ''
    ).join('') + this.sheets.map((s, i) => {
      const page = s.options.pageSetup;
      let names = '';
      if (page?.printArea) names += `<definedName name="_xlnm.Print_Area" localSheetId="${i}">${escapeXml(`${quoteSheet(s.name)}!${absoluteRef(page.printArea)}`)}</definedName>`;
      if (page?.printTitleRows) {
        const rows = page.printTitleRows.replace(/\$/g, '').split(':').map(r => `$${r}`).join(':');
        names += `<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${escapeXml(`${quoteSheet(s.name)}!${rows.includes(':') ? rows : `${rows}:${rows}`}`)}</definedName>`;
      }
      return names;
    }).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${active > 0 ? `<bookViews><workbookView firstSheet="${active}" activeTab="${active}"/></bookViews>` : ''}
  <sheets>
    ${sheetsXml}
  </sheets>
  ${filterNames || userNames ? `<definedNames>${filterNames}${userNames}</definedNames>` : ''}
</workbook>`;
  }

  // The first visible sheet, which Excel opens on
  private activeSheet(): number {
    return this.sheets.findIndex(s => !s.options.state || s.options.state === 'visible');
  }

  private buildWorkbookRels(): string {
    const sheetRels = this.sheets.map((_, i) => 
      `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`
    ).join('\n  ');
    
    const ssId = this.sheets.length + 1;
    const stylesId = this.sheets.length + 2;

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${sheetRels}
  <Relationship Id="rId${ssId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId${stylesId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  private buildRootRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>${this.writerOptions.properties ? `
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` : ''}
</Relationships>`;
  }

  private buildContentTypes(drawings: number, commentSheets: number[], tables: number): string {
    const sheetOverrides = this.sheets.map((_, i) => 
      `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    ).join('\n  ');
    const exts = new Set([...this.media.values()].map(m => m.info.ext));
    const imageTypes = [...exts].map(e => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('');
    const drawingOverrides = Array.from({ length: drawings }, (_, i) =>
      `<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join('') +
      commentSheets.map(n => `<Override PartName="/xl/comments${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`).join('') +
      Array.from({ length: tables }, (_, i) =>
        `<Override PartName="/xl/tables/table${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`).join('') +
      (commentSheets.length ? '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>' : '');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${imageTypes}${drawingOverrides}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetOverrides}
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${this.writerOptions.properties ? `
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>` : ''}
</Types>`;
  }


}

function pageSetupXml(page?: PageSetup): string {
  if (!page) return '';
  let xml = '';
  if (page.gridLines || page.horizontalCentered) {
    xml += `  <printOptions${page.horizontalCentered ? ' horizontalCentered="1"' : ''}${page.gridLines ? ' gridLines="1"' : ''}/>\n`;
  }
  // Excel's "Normal" margins, in inches; all six are required
  const m = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3, ...page.margins };
  const e = escapeXml;
  xml += `  <pageMargins left="${e(m.left)}" right="${e(m.right)}" top="${e(m.top)}" bottom="${e(m.bottom)}" header="${e(m.header)}" footer="${e(m.footer)}"/>\n`;
  const attrs = [
    page.paperSize !== undefined && `paperSize="${e(page.paperSize)}"`,
    page.scale !== undefined && `scale="${e(page.scale)}"`,
    page.fitToWidth !== undefined && `fitToWidth="${e(page.fitToWidth)}"`,
    page.fitToHeight !== undefined && `fitToHeight="${e(page.fitToHeight)}"`,
    page.orientation && `orientation="${e(page.orientation)}"`,
  ].filter(Boolean).join(' ');
  if (attrs) xml += `  <pageSetup ${attrs}/>\n`;
  if (page.header || page.footer) {
    xml += `  <headerFooter>${page.header ? `<oddHeader>${escapeXml(page.header)}</oddHeader>` : ''}${page.footer ? `<oddFooter>${escapeXml(page.footer)}</oddFooter>` : ''}</headerFooter>\n`;
  }
  return xml;
}

function commentsXml(comments: SheetParts['comments']): string {
  const authors = [...new Set(comments.map(c => c.comment.author ?? ''))];
  const list = comments.map(c => `<comment ref="${c.ref}" authorId="${authors.indexOf(c.comment.author ?? '')}"><text>${runsXml(c.comment.text)}</text></comment>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>${authors.map(a => `<author>${escapeXml(a)}</author>`).join('')}</authors><commentList>${list}</commentList></comments>`;
}

// The legacy VML shapes Excel needs to show the note boxes (hidden until the cell is hovered)
function vmlXml(comments: SheetParts['comments'], sheet: number): string {
  const shapes = comments.map((c, i) =>
    `<v:shape id="_x0000_s${sheet * 1024 + i + 1}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto">` +
    `<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>` +
    `<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>` +
    `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>` +
    `<x:Anchor>${c.col + 1}, 15, ${c.row}, 10, ${c.col + 3}, 15, ${c.row + 4}, 4</x:Anchor><x:AutoFill>False</x:AutoFill>` +
    `<x:Row>${c.row}</x:Row><x:Column>${c.col}</x:Column></x:ClientData></v:shape>`).join('');
  return `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">` +
    `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${sheet}"/></o:shapelayout>` +
    `<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>` +
    `${shapes}</xml>`;
}

function tableXml(table: TableOptions, id: number, columns: string[]): string {
  const ref = table.ref.replace(/\$/g, '');
  const filter = table.autoFilter === false ? '' : `<autoFilter ref="${ref}"/>`;
  const cols = columns.map((c, k) => `<tableColumn id="${k + 1}" name="${escapeXml(c)}"/>`).join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="${id}" name="${escapeXml(table.name)}" displayName="${escapeXml(table.name)}" ref="${ref}" totalsRowShown="0">` +
    `${filter}<tableColumns count="${columns.length}">${cols}</tableColumns>` +
    `<tableStyleInfo name="${escapeXml(table.style ?? 'TableStyleMedium2')}" showFirstColumn="0" showLastColumn="0" showRowStripes="${table.showRowStripes === false ? 0 : 1}" showColumnStripes="0"/></table>`;
}
