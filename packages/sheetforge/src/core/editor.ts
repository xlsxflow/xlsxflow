import { createZipStreamParser, ZipEntry } from './zip-stream';
import { Row, SheetOptions } from '../pro/types';
import { ZipStreamWriter } from './zip-stream-writer';

export class SheetEditor {
  private modifications = new Map<string, { rows: Row[], options: SheetOptions }>();

  appendSheet(sheetName: string, rows: Row[], options: SheetOptions = {}) {
    this.modifications.set(sheetName, { rows, options });
  }

  edit(inputStream: ReadableStream<Uint8Array>): ReadableStream<Uint8Array> {
    const zipStream = inputStream.pipeThrough(createZipStreamParser());
    const zipOut = new ZipStreamWriter();

    (async () => {
      try {
        const zipReader = zipStream.getReader();
        let workbookXml = '';
        let relsXml = '';
        const sheetNameToPath = new Map<string, string>();
        const pathToSheetName = new Map<string, string>();

        while (true) {
          const { done, value } = await zipReader.read();
          if (done) break;

          const entry = value as ZipEntry;
          let s = entry.stream;
          if (entry.compressionMethod === 8) {
            s = s.pipeThrough(new DecompressionStream('deflate-raw') as any);
          }

          if (entry.filename === 'xl/workbook.xml') {
            workbookXml = await this.readStreamToString(s);
            this.updatePaths(workbookXml, relsXml, sheetNameToPath, pathToSheetName);
            await this.streamString(zipOut, entry.filename, workbookXml);
          } else if (entry.filename === 'xl/_rels/workbook.xml.rels') {
            relsXml = await this.readStreamToString(s);
            this.updatePaths(workbookXml, relsXml, sheetNameToPath, pathToSheetName);
            await this.streamString(zipOut, entry.filename, relsXml);
          } else {
            const sheetName = pathToSheetName.get(entry.filename);
            const mod = sheetName ? this.modifications.get(sheetName) : null;
            if (mod) {
              const textDecoder = new TextDecoderStream();
              const textEncoder = new TextEncoderStream();
              const injectStream = this.createInjectTransform(mod.rows);
              
              const modifiedStream = s
                .pipeThrough(textDecoder as any)
                .pipeThrough(injectStream)
                .pipeThrough(textEncoder as any) as unknown as ReadableStream<Uint8Array>;
              
              await zipOut.addFile(entry.filename, modifiedStream);
            } else {
              await zipOut.addFile(entry.filename, s);
            }
          }
        }
        await zipOut.close();
      } catch (err) {
        console.error('Editor Error:', err);
      }
    })();

    return zipOut.stream;
  }

  private updatePaths(workbookXml: string, relsXml: string, sheetNameToPath: Map<string, string>, pathToSheetName: Map<string, string>) {
    if (!workbookXml || !relsXml) return;
    const sheetRegex = /<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"/g;
    let match;
    while ((match = sheetRegex.exec(workbookXml)) !== null) {
      const name = match[1];
      const rId = match[2];
      const relRegex = new RegExp(`<Relationship [^>]*Id="${rId}" [^>]*Target="([^"]+)"`);
      const relMatch = relsXml.match(relRegex);
      if (relMatch) {
        let target = relMatch[1];
        if (target.startsWith('/xl/')) target = target.slice(4);
        const path = `xl/${target}`;
        sheetNameToPath.set(name, path);
        pathToSheetName.set(path, name);
      }
    }
  }

  private createInjectTransform(rows: Row[]): TransformStream<string, string> {
    let buffer = '';
    let maxRow = 0;
    const self = this;
    
    return new TransformStream({
      transform(chunk, controller) {
        buffer += chunk;
        
        const rowRegex = /<row [^>]*r="(\d+)"/g;
        let match;
        while ((match = rowRegex.exec(buffer)) !== null) {
          const r = parseInt(match[1], 10);
          if (r > maxRow) maxRow = r;
        }

        const idx = buffer.indexOf('</sheetData>');
        if (idx !== -1) {
          controller.enqueue(buffer.slice(0, idx));
          
          let appendedXml = '';
          rows.forEach((row, ri) => {
            const rowNum = maxRow + ri + 1;
            const cellsXml = row.map((cell, ci) => {
              const colRef = self.colLetter(ci) + rowNum;
              const val = typeof cell === 'object' && cell !== null && 'value' in cell ? cell.value : cell;
              if (val === null || val === undefined) return `<c r="${colRef}"/>`;
              if (typeof val === 'boolean') return `<c r="${colRef}" t="b"><v>${val ? 1 : 0}</v></c>`;
              if (typeof val === 'number') return `<c r="${colRef}"><v>${val}</v></c>`;
              if (typeof val === 'string') {
                return `<c r="${colRef}" t="inlineStr"><is><t>${self.escapeXml(val)}</t></is></c>`;
              }
              return `<c r="${colRef}"/>`;
            }).join('');
            appendedXml += `<row r="${rowNum}">${cellsXml}</row>`;
          });

          controller.enqueue(appendedXml);
          controller.enqueue(buffer.slice(idx));
          buffer = '';
        } else {
          if (buffer.length > 100) {
            controller.enqueue(buffer.slice(0, buffer.length - 100));
            buffer = buffer.slice(-100);
          }
        }
      },
      flush(controller) {
        if (buffer.length > 0) {
          controller.enqueue(buffer);
        }
      }
    });
  }

  private async streamString(zip: ZipStreamWriter, filename: string, content: string) {
    const encoder = new TextEncoder();
    const data = encoder.encode(content);
    const stream = new ReadableStream({
      start(c) { c.enqueue(data); c.close(); }
    });
    await zip.addFile(filename, stream);
  }

  private colLetter(colIndex: number): string {
    let col = '';
    let c = colIndex + 1;
    while (c > 0) {
      col = String.fromCharCode(((c - 1) % 26) + 65) + col;
      c = Math.floor((c - 1) / 26);
    }
    return col;
  }

  private escapeXml(val: string): string {
    return val
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  private async readStreamToString(stream: ReadableStream<Uint8Array>): Promise<string> {
    const textDecoder = new TextDecoderStream();
    const reader = stream.pipeThrough(textDecoder as any).getReader();
    let result = '';
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      result += value;
    }
    return result;
  }
}
