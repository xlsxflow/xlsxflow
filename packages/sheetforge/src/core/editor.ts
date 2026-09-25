import { RandomAccessReader } from './random-access';
import { ZipRandomAccessParser } from './zip-random-access';
import { Row, SheetOptions } from './types';
import { ZipStreamWriter } from './zip-stream-writer';
import { resolveWorkbookParts } from './utils';

export class SheetEditor {
  private modifications = new Map<string, { rows: Row[], options: SheetOptions }>();

  // Appends `rows` after the last existing row of sheet `sheetName`.
  appendSheet(sheetName: string, rows: Row[], options: SheetOptions = {}) {
    this.modifications.set(sheetName, { rows, options });
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
        const { sheets: sheetPaths } = await resolveWorkbookParts(readText);

        const modsByPath = new Map<string, Row[]>();
        for (const [sheetName, mod] of this.modifications) {
          const path = sheetPaths.get(sheetName);
          if (!path || !zipIn.has(path)) throw new Error(`Sheet "${sheetName}" not found in workbook.`);
          modsByPath.set(path, mod.rows);
        }

        for (const filename of zipIn.getFiles()) {
          const rows = modsByPath.get(filename);
          if (rows) {
            const modifiedStream = (await zipIn.extractStream(filename))
              .pipeThrough(new TextDecoderStream() as any as TransformStream<Uint8Array, string>)
              .pipeThrough(this.createInjectTransform(rows))
              .pipeThrough(new TextEncoderStream() as any as TransformStream<string, Uint8Array>);
            await zipOut.addFile(filename, modifiedStream);
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

  private createInjectTransform(rows: Row[]): TransformStream<string, string> {
    let buffer = '';
    let maxRow = 0;
    let done = false;
    const rowsXml = (): string => rows.map((row, ri) => {
      const rowNum = maxRow + ri + 1;
      const cellsXml = row.map((cell, ci) => {
        const colRef = this.colLetter(ci) + rowNum;
        const val = typeof cell === 'object' && cell !== null && 'value' in cell ? cell.value : cell;
        if (val === null || val === undefined) return `<c r="${colRef}"/>`;
        if (typeof val === 'boolean') return `<c r="${colRef}" t="b"><v>${val ? 1 : 0}</v></c>`;
        if (typeof val === 'number') return `<c r="${colRef}"><v>${val}</v></c>`;
        if (typeof val === 'string') {
          return `<c r="${colRef}" t="inlineStr"><is><t>${this.escapeXml(val)}</t></is></c>`;
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
