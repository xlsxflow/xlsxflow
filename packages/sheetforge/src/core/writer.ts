import { Row, StyledCell, CellValue, SheetOptions } from '../pro/types';
import { StyleEngine } from '../pro/style-engine';
import { ConditionalFormatter } from '../pro/conditional-formatter';
import { FormulaEngine } from '../pro/formula-engine';
import { ZipStreamWriter } from './zip-stream-writer';

// Utility: convert 0-based col index to column letter (A, B, ..., Z, AA...)
function colLetter(colIndex: number): string {
  let col = '';
  let c = colIndex + 1;
  while (c > 0) {
    col = String.fromCharCode(((c - 1) % 26) + 65) + col;
    c = Math.floor((c - 1) / 26);
  }
  return col;
}

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

// Builds a complete .xlsx Uint8Array from structured row data.
// Uses a pure JS ZIP writer (no dependencies) to pack OOXML parts.
export class SheetWriter {
  private styleEngine = new StyleEngine();
  private conditionalFormatter = new ConditionalFormatter();
  private formulaEngine = new FormulaEngine();
  private sharedStrings: string[] = [];
  private sharedStringMap = new Map<string, number>();
  private sheets: { name: string; rows: Row[] | AsyncIterable<Row>; options: SheetOptions }[] = [];

  addSheet(name: string, rows: Row[] | AsyncIterable<Row>, options: SheetOptions = {}): this {
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
    
    // Push the build process to background so we can return the stream immediately
    (async () => {
      try {
        await this.streamString(zip, '[Content_Types].xml', this.buildContentTypes());
        await this.streamString(zip, '_rels/.rels', this.buildRootRels());
        await this.streamString(zip, 'xl/workbook.xml', this.buildWorkbookXml());
        await this.streamString(zip, 'xl/_rels/workbook.xml.rels', this.buildWorkbookRels());

        // Build sharedStrings and styles (for async iterable we can't pre-calculate them without 2 passes, 
        // so we must use inlineStr for strings if the user provides an AsyncIterable).
        const sheetsMetadata: { options: SheetOptions }[] = [];
        for (let i = 0; i < this.sheets.length; i++) {
          sheetsMetadata.push({ options: this.sheets[i].options });
        }

        // Pre-pass for arrays to register styles and load formula data
        for (let i = 0; i < this.sheets.length; i++) {
          const sheet = this.sheets[i];
          if (Array.isArray(sheet.rows)) {
            const rawData = sheet.rows.map((row: Row) =>
              row.map((cell: any) => isStyledCell(cell) ? cell.value : cell)
            );
            this.formulaEngine.loadData(rawData);
            
            // Pre-register styles
            sheet.rows.forEach(row => {
              row.forEach(cell => {
                if (isStyledCell(cell) && cell.style) {
                  this.styleEngine.registerStyle(cell.style);
                }
              });
            });
          }
        }

        // Write styles and sharedStrings FIRST so streaming readers can parse them early
        await this.streamString(zip, 'xl/sharedStrings.xml', this.buildSharedStringsXml());
        await this.streamString(zip, 'xl/styles.xml', this.styleEngine.toXml());

        // Then write worksheets
        for (let i = 0; i < this.sheets.length; i++) {
          const sheet = this.sheets[i];
          const stream = this.buildWorksheetXmlStream(sheet.rows, sheet.options);
          await zip.addFile(`xl/worksheets/sheet${i + 1}.xml`, stream);
        }

        await zip.close();
      } catch (err) {
        console.error('Writer Error:', err);
      }
    })();

    return zip.stream;
  }

  private async streamString(zip: any, filename: string, content: string) {
    const encoder = new TextEncoder();
    const data = encoder.encode(content);
    const stream = new ReadableStream({
      start(c) { c.enqueue(data); c.close(); }
    });
    await zip.addFile(filename, stream);
  }

  // ... (keep buildWorksheetXml, addSharedString, buildSharedStringsXml, buildWorkbookXml, buildWorkbookRels, buildRootRels, buildContentTypes unchanged)

  private buildWorksheetXmlStream(rows: Row[] | AsyncIterable<Row>, options: SheetOptions): ReadableStream<Uint8Array> {
    const encoder = new TextEncoder();
    const self = this;

    return new ReadableStream({
      async start(controller) {
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
        controller.enqueue(encoder.encode(header));

        let ri = 0;
        
        // Handle both Array and AsyncIterable
        const iterator = Symbol.asyncIterator in rows 
          ? (rows as AsyncIterable<Row>)[Symbol.asyncIterator]()
          : (async function*() { for (const r of rows as Row[]) yield r; })();

        while (true) {
          const { done, value: row } = await iterator.next();
          if (done) break;

          const rowNum = ri + 1;
          const cellsXml = row.map((cell, ci) => {
            const colRef = colLetter(ci) + rowNum;
            const styledCell = isStyledCell(cell) ? cell : { value: cell };
            const style = styledCell.style ? self.styleEngine.registerStyle(styledCell.style) : 0;
            const sAttr = style > 0 ? ` s="${style}"` : '';

            if (styledCell.formula) {
              const result = self.formulaEngine.evaluate(styledCell.formula);
              const cachedVal = typeof result === 'number' ? `<v>${result}</v>` : (result ? `<v>${escapeXml(String(result))}</v>` : '');
              return `<c r="${colRef}"${sAttr}><f>${escapeXml(styledCell.formula.replace(/^=/, ''))}</f>${cachedVal}</c>`;
            }

            const val = styledCell.value;
            if (val === null || val === undefined) return `<c r="${colRef}"${sAttr}/>`;
            if (typeof val === 'boolean') return `<c r="${colRef}" t="b"${sAttr}><v>${val ? 1 : 0}</v></c>`;
            if (typeof val === 'number') return `<c r="${colRef}"${sAttr}><v>${val}</v></c>`;
            if (typeof val === 'string') {
              // Always use inlineStr for streaming (avoids memory overhead of sharedStrings)
              return `<c r="${colRef}" t="inlineStr"${sAttr}><is><t>${escapeXml(val)}</t></is></c>`;
            }
            return `<c r="${colRef}"${sAttr}/>`;
          }).join('');
          
          controller.enqueue(encoder.encode(`    <row r="${rowNum}">${cellsXml}</row>\n`));
          ri++;
        }

        let footer = `  </sheetData>\n`;

        if (options.mergeCells && options.mergeCells.length > 0) {
          const merges = options.mergeCells.map(ref => `<mergeCell ref="${escapeXml(ref)}"/>`).join('');
          footer += `  <mergeCells count="${options.mergeCells.length}">${merges}</mergeCells>\n`;
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

        if (options.conditionalFormats) {
          footer += self.conditionalFormatter.toXml(options.conditionalFormats) + '\n';
        }

        footer += `</worksheet>`;
        controller.enqueue(encoder.encode(footer));
        controller.close();
      }
    });
  }

  private addSharedString(str: string): number {
    if (this.sharedStringMap.has(str)) return this.sharedStringMap.get(str)!;
    const idx = this.sharedStrings.length;
    this.sharedStrings.push(str);
    this.sharedStringMap.set(str, idx);
    return idx;
  }

  private buildSharedStringsXml(): string {
    const count = this.sharedStrings.length;
    const items = this.sharedStrings.map(s => `<si><t>${escapeXml(s)}</t></si>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${count}" uniqueCount="${count}">
${items}
</sst>`;
  }

  private buildWorkbookXml(): string {
    const sheetsXml = this.sheets.map((s, i) => 
      `<sheet name="${escapeXml(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`
    ).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    ${sheetsXml}
  </sheets>
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

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  ${sheetOverrides}
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  }


}
