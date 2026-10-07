import { Row, StyledCell, CellValue, SheetOptions, SheetImage } from './types';
import { StyleEngine, fontXml } from './style-engine';
import { ConditionalFormatter } from './conditional-formatter';
import { FormulaEngine } from './formula-engine';
import { ZipStreamWriter } from './zip-stream-writer';
import { encodeXString, colLetter, dateToSerial } from './utils';
import { imageInfo, drawingXml, type ImageInfo } from './image';

function escapeXml(val: string): string {
  return val
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

export interface WriterOptions {
  // Store strings once in xl/sharedStrings.xml instead of inline. Smaller files when values repeat,
  // but every distinct string is kept in memory until the workbook is finished.
  sharedStrings?: boolean;
}

// Builds a complete .xlsx Uint8Array from structured row data.
// Uses a pure JS ZIP writer (no dependencies) to pack OOXML parts.
export class SheetWriter {
  private styleEngine = new StyleEngine();
  private conditionalFormatter = new ConditionalFormatter();
  private formulaEngine = new FormulaEngine();
  private sheets: { name: string; rows: Row[] | AsyncIterable<Row>; options: SheetOptions }[] = [];
  private sharedStrings = new Map<string, number>();
  private sharedStringRefs = 0;
  // Embedded pictures by their data object, so a logo used on every sheet is stored once
  private media = new Map<Uint8Array | ArrayBuffer, { path: string; bytes: Uint8Array; info: ImageInfo }>();

  constructor(private writerOptions: WriterOptions = {}) {}

  addSheet(name: string, rows: Row[] | AsyncIterable<Row>, options: SheetOptions = {}): this {
    // Excel refuses to open a workbook that breaks these rules
    if (!name || name.length > 31 || /[\\/?*:[\]]/.test(name) || name.startsWith("'") || name.endsWith("'")) {
      throw new Error(`Invalid sheet name "${name}": 1-31 characters, none of \\ / ? * : [ ], and no leading or trailing apostrophe.`);
    }
    if (this.sheets.some(s => s.name.toLowerCase() === name.toLowerCase())) throw new Error(`Duplicate sheet name "${name}".`);
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

    const zip = new ZipStreamWriter();
    for (const img of this.sheets.flatMap(s => s.options.images ?? [])) {
      if (this.media.has(img.data)) continue;
      const bytes = img.data instanceof Uint8Array ? img.data : new Uint8Array(img.data);
      const info = imageInfo(bytes);
      this.media.set(img.data, { path: `media/image${this.media.size + 1}.${info.ext}`, bytes, info });
    }
    let drawings = 0;

    // Push the build process to background so we can return the stream immediately
    (async () => {
      try {
        await this.streamString(zip, '[Content_Types].xml', this.buildContentTypes());
        await this.streamString(zip, '_rels/.rels', this.buildRootRels());
        await this.streamString(zip, 'xl/workbook.xml', this.buildWorkbookXml());
        await this.streamString(zip, 'xl/_rels/workbook.xml.rels', this.buildWorkbookRels());

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
          const links: string[] = [];
          await zip.addFile(`xl/worksheets/sheet${i + 1}.xml`, this.buildWorksheetXmlStream(sheet.rows, sheet.options, links));
          const images = sheet.options.images ?? [];
          const drawing = images.length ? ++drawings : 0;
          if (links.length || drawing) await this.streamString(zip, `xl/worksheets/_rels/sheet${i + 1}.xml.rels`, this.buildSheetRels(links, drawing));
          if (drawing) await this.writeDrawing(zip, drawing, images);
        }

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
  private buildWorksheetXmlStream(rows: Row[] | AsyncIterable<Row>, options: SheetOptions, links: string[]): ReadableStream<Uint8Array> {
    const encoder = new TextEncoder();
    const chunks = this.worksheetXmlChunks(rows, options, links);
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

  // External hyperlink targets are pushed onto `links` (their index + 1 is the relationship id)
  private async *worksheetXmlChunks(rows: Row[] | AsyncIterable<Row>, options: SheetOptions, links: string[]): AsyncGenerator<string> {
    const hyperlinks: string[] = [];
    const colWidths = options.columnWidths
      ? options.columnWidths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('')
      : '';

    let sheetViews = '';
    if (options.freezePanes) {
      const { row = 0, col = 0 } = options.freezePanes;
      let activePane = 'bottomRight';
      if (row > 0 && col === 0) activePane = 'bottomLeft';
      if (row === 0 && col > 0) activePane = 'topRight';
      const topLeftCell = colLetter(col) + (row + 1);
      sheetViews = `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ySplit="${row}" xSplit="${col}" topLeftCell="${topLeftCell}" activePane="${activePane}" state="frozen"/></sheetView></sheetViews>`;
    }

    let header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
    header += `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n`;
    header += `  ${sheetViews}\n`;
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
          const runs = styledCell.richText.map(run =>
            `<r>${run.font ? `<rPr>${fontXml(run.font, 'rFont')}</rPr>` : ''}${stringXml(run.text)}</r>`).join('');
          return `<c r="${colRef}" t="inlineStr"${sAttr}><is>${runs}</is></c>`;
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
      
      chunkStr += `    <row r="${rowNum}">${cellsXml}</row>\n`;
      ri++;

      // Flush every ~64KB of string data to keep O(1) memory while avoiding chunk overhead
      if (chunkStr.length > 65536) {
        yield chunkStr;
        chunkStr = '';
      }
    }
    
    if (chunkStr.length > 0) {
      yield chunkStr;
    }

    let footer = `  </sheetData>\n`;

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
        if (dv.type) attr += ` type="${dv.type}"`;
        if (dv.allowBlank !== undefined) attr += ` allowBlank="${dv.allowBlank ? 1 : 0}"`;
        if (dv.showInputMessage !== undefined) attr += ` showInputMessage="${dv.showInputMessage ? 1 : 0}"`;
        if (dv.showErrorMessage !== undefined) attr += ` showErrorMessage="${dv.showErrorMessage ? 1 : 0}"`;
        let inner = '';
        if (dv.formula1) inner += `<formula1>${escapeXml(dv.formula1)}</formula1>`;
        if (dv.formula2) inner += `<formula2>${escapeXml(dv.formula2)}</formula2>`;
        return `<dataValidation ${attr}>${inner}</dataValidation>`;
      }).join('');
      footer += `  <dataValidations count="${options.dataValidations.length}">${dvs}</dataValidations>\n`;
    }

    if (hyperlinks.length) footer += `  <hyperlinks>${hyperlinks.join('')}</hyperlinks>\n`;
    if (options.images?.length) footer += `  <drawing r:id="rIdDrawing"/>\n`;

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

  private buildSheetRels(links: string[], drawing: number): string {
    let rels = links.map((target, i) =>
      `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXml(target)}" TargetMode="External"/>`
    ).join('');
    if (drawing) rels += `<Relationship Id="rIdDrawing" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${drawing}.xml"/>`;
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>`;
  }

  private buildWorkbookXml(): string {
    const sheetsXml = this.sheets.map((s, i) => 
      `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`
    ).join('');
    // Excel tracks each sheet's filter range in a hidden defined name
    const filterNames = this.sheets.map((s, i) => s.options.autoFilter
      ? `<definedName name="_xlnm._FilterDatabase" localSheetId="${i}" hidden="1">${escapeXml(`'${s.name.replace(/'/g, "''")}'!${absoluteRef(s.options.autoFilter)}`)}</definedName>`
      : ''
    ).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${sheetsXml}
  </sheets>
  ${filterNames ? `<definedNames>${filterNames}</definedNames>` : ''}
</workbook>`;
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
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  }

  private buildContentTypes(): string {
    const sheetOverrides = this.sheets.map((_, i) => 
      `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
    ).join('\n  ');
    const exts = new Set([...this.media.values()].map(m => m.info.ext));
    const imageTypes = [...exts].map(e => `<Default Extension="${e}" ContentType="image/${e}"/>`).join('');
    const drawingOverrides = this.sheets.filter(s => s.options.images?.length).map((_, i) =>
      `<Override PartName="/xl/drawings/drawing${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  ${imageTypes}${drawingOverrides}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetOverrides}
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  }


}
