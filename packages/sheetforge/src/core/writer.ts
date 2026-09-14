import { Row, StyledCell, CellValue, SheetOptions } from '../pro/types';
import { StyleEngine } from '../pro/style-engine';
import { ConditionalFormatter } from '../pro/conditional-formatter';
import { FormulaEngine } from '../pro/formula-engine';

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

  async write(rows: Row[], options: SheetOptions = {}): Promise<Uint8Array> {
    // Pre-load data into formula engine for evaluation
    const rawData = rows.map(row =>
      row.map(cell => isStyledCell(cell) ? cell.value as any : cell as any)
    );
    this.formulaEngine.loadData(rawData);

    const sheetXml = this.buildWorksheetXml(rows, options);
    const sharedStringsXml = this.buildSharedStringsXml();
    const stylesXml = this.styleEngine.toXml();
    const conditionalXml = options.conditionalFormats
      ? this.conditionalFormatter.toXml(options.conditionalFormats)
      : '';

    // Assemble OOXML parts into a ZIP
    const parts: Record<string, string> = {
      '[Content_Types].xml': this.buildContentTypes(),
      '_rels/.rels': this.buildRootRels(),
      'xl/workbook.xml': this.buildWorkbookXml(options.name ?? 'Sheet1'),
      'xl/_rels/workbook.xml.rels': this.buildWorkbookRels(),
      'xl/worksheets/sheet1.xml': sheetXml + conditionalXml,
      'xl/sharedStrings.xml': sharedStringsXml,
      'xl/styles.xml': stylesXml,
    };

    return this.packZip(parts);
  }

  // ... (keep buildWorksheetXml, addSharedString, buildSharedStringsXml, buildWorkbookXml, buildWorkbookRels, buildRootRels, buildContentTypes unchanged)

  private buildWorksheetXml(rows: Row[], options: SheetOptions): string {
    const colWidths = options.columnWidths
      ? options.columnWidths.map((w, i) =>
          `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`
        ).join('')
      : '';

    const rowsXml = rows.map((row, ri) => {
      const rowNum = ri + 1;
      const cellsXml = row.map((cell, ci) => {
        const colRef = colLetter(ci) + rowNum;
        const styledCell = isStyledCell(cell) ? cell : { value: cell };
        const style = styledCell.style ? this.styleEngine.registerStyle(styledCell.style) : 0;
        const sAttr = style > 0 ? ` s="${style}"` : '';

        if (styledCell.formula) {
          const result = this.formulaEngine.evaluate(styledCell.formula);
          const cachedVal = typeof result === 'number' ? `<v>${result}</v>` : (result ? `<v>${escapeXml(String(result))}</v>` : '');
          return `<c r="${colRef}"${sAttr}><f>${escapeXml(styledCell.formula.replace(/^=/, ''))}</f>${cachedVal}</c>`;
        }

        const val = styledCell.value;
        if (val === null || val === undefined) return `<c r="${colRef}"${sAttr}/>`;
        if (typeof val === 'boolean') return `<c r="${colRef}" t="b"${sAttr}><v>${val ? 1 : 0}</v></c>`;
        if (typeof val === 'number') return `<c r="${colRef}"${sAttr}><v>${val}</v></c>`;
        if (typeof val === 'string') {
          const idx = this.addSharedString(val);
          return `<c r="${colRef}" t="s"${sAttr}><v>${idx}</v></c>`;
        }
        return `<c r="${colRef}"${sAttr}/>`;
      }).join('');
      return `<row r="${rowNum}">${cellsXml}</row>`;
    }).join('');

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  ${colWidths ? `<cols>${colWidths}</cols>` : ''}
  <sheetData>
    ${rowsXml}
  </sheetData>
</worksheet>`;
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

  private buildWorkbookXml(sheetName: string): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="${escapeXml(sheetName)}" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;
  }

  private buildWorkbookRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
  }

  private buildRootRels(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
  }

  private buildContentTypes(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
  }

  private async compressData(data: Uint8Array): Promise<Uint8Array> {
    const stream = new ReadableStream({
      start(controller) {
        controller.enqueue(data);
        controller.close();
      }
    });
    const compressedStream = stream.pipeThrough(new CompressionStream('deflate-raw'));
    const reader = compressedStream.getReader();
    const chunks: Uint8Array[] = [];
    let totalLen = 0;
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      totalLen += value.length;
    }
    const result = new Uint8Array(totalLen);
    let offset = 0;
    for (const c of chunks) {
      result.set(c, offset);
      offset += c.length;
    }
    return result;
  }

  // Minimal pure-JS ZIP packer (Deflate compression — no dependencies)
  private async packZip(parts: Record<string, string>): Promise<Uint8Array> {
    const encoder = new TextEncoder();
    const entries: { filename: Uint8Array; data: Uint8Array; offset: number; uncompressedSize: number; crc: number }[] = [];
    const chunks: Uint8Array[] = [];
    let offset = 0;

    for (const [path, content] of Object.entries(parts)) {
      const filename = encoder.encode(path);
      const uncompressedData = encoder.encode(content);
      const uncompressedSize = uncompressedData.length;
      const crc = this.crc32(uncompressedData);
      
      const data = await this.compressData(uncompressedData);

      // Local file header
      const header = new Uint8Array(30 + filename.length);
      const view = new DataView(header.buffer);
      view.setUint32(0, 0x04034b50, true);  // signature
      view.setUint16(4, 20, true);           // version needed
      view.setUint16(6, 0, true);            // flags
      view.setUint16(8, 8, true);            // compression (Deflate)
      view.setUint16(10, 0, true);           // mod time
      view.setUint16(12, 0, true);           // mod date
      view.setUint32(14, crc, true);         // CRC-32 (uncompressed)
      view.setUint32(18, data.length, true); // compressed size
      view.setUint32(22, uncompressedSize, true); // uncompressed size
      view.setUint16(26, filename.length, true); // filename length
      view.setUint16(28, 0, true);           // extra field length
      header.set(filename, 30);

      entries.push({ filename, data, offset, uncompressedSize, crc });
      chunks.push(header, data);
      offset += header.length + data.length;
    }

    // Central directory
    const cdChunks: Uint8Array[] = [];
    let cdSize = 0;

    for (const entry of entries) {
      const cd = new Uint8Array(46 + entry.filename.length);
      const view = new DataView(cd.buffer);
      view.setUint32(0, 0x02014b50, true);   // central dir signature
      view.setUint16(4, 20, true);            // version made by
      view.setUint16(6, 20, true);            // version needed
      view.setUint16(8, 0, true);             // flags
      view.setUint16(10, 8, true);            // compression (Deflate)
      view.setUint16(12, 0, true);            // mod time
      view.setUint16(14, 0, true);            // mod date
      view.setUint32(16, entry.crc, true);    // CRC-32
      view.setUint32(20, entry.data.length, true); // compressed size
      view.setUint32(24, entry.uncompressedSize, true); // uncompressed size
      view.setUint16(28, entry.filename.length, true); // filename length
      view.setUint16(30, 0, true);            // extra field length
      view.setUint16(32, 0, true);            // comment length
      view.setUint16(34, 0, true);            // disk start
      view.setUint16(36, 0, true);            // internal attributes
      view.setUint32(38, 0, true);            // external attributes
      view.setUint32(42, entry.offset, true); // local header offset
      cd.set(entry.filename, 46);
      cdChunks.push(cd);
      cdSize += cd.length;
    }

    // End of central directory record
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, entries.length, true);
    eocdView.setUint16(10, entries.length, true);
    eocdView.setUint32(12, cdSize, true);
    eocdView.setUint32(16, offset, true);
    eocdView.setUint16(20, 0, true);

    const allChunks = [...chunks, ...cdChunks, eocd];
    const total = allChunks.reduce((s, c) => s + c.length, 0);
    const result = new Uint8Array(total);
    let pos = 0;
    for (const chunk of allChunks) {
      result.set(chunk, pos);
      pos += chunk.length;
    }
    return result;
  }

  private crc32(data: Uint8Array): number {
    let crc = 0xffffffff;
    for (const byte of data) {
      crc ^= byte;
      for (let i = 0; i < 8; i++) {
        crc = (crc & 1) ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
      }
    }
    return (crc ^ 0xffffffff) >>> 0;
  }
}
