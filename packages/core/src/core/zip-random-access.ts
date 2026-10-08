import { RandomAccessReader } from './random-access';
import { crc32Update } from './zip-stream-writer';
import { isCfb } from './cfb';

export interface ZipRecord {
  filename: string;
  compressionMethod: number;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  localHeaderOffset: number;
}

export class ZipRandomAccessParser {
  private records = new Map<string, ZipRecord>();
  // Part names are case-insensitive in OPC: a rels target "Sheet1.xml" finds "sheet1.xml"
  private lowerCase = new Map<string, ZipRecord>();

  constructor(private reader: RandomAccessReader) {}

  async parseCentralDirectory(): Promise<void> {
    const size = this.reader.size;
    if (size < 22) throw new Error("File too small to be a valid ZIP.");

    // The EOCD record can be anywhere in the last 65535 + 22 bytes due to an optional file comment.
    const searchSize = Math.min(size, 65535 + 22);
    const searchStart = size - searchSize;
    const buffer = await this.reader.read(searchStart, searchSize);
    
    const dataView = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);

    // Search backwards for the EOCD signature (0x06054b50). The real record's comment reaches the end
    // of the file, so the signature bytes inside a comment are skipped; failing that (junk appended
    // after the archive), the last signature is used.
    let eocdOffset = -1;
    for (let i = buffer.length - 22; i >= 0; i--) {
      if (dataView.getUint32(i, true) !== 0x06054b50) continue;
      if (eocdOffset === -1) eocdOffset = i;
      if (i + 22 + dataView.getUint16(i + 20, true) === buffer.length) {
        eocdOffset = i;
        break;
      }
    }

    if (eocdOffset === -1) {
      if (isCfb(buffer.subarray(0, 8)) || searchStart > 0 && isCfb(await this.reader.read(0, 8))) {
        throw new Error('This file is a compound file, not a ZIP: an .xls, or a workbook saved with a password. Decrypt it first with decryptWorkbook from @xlsxflow/pro.');
      }
      throw new Error("End of Central Directory (EOCD) not found. This may not be a valid ZIP file.");
    }

    const totalRecords = dataView.getUint16(eocdOffset + 10, true);
    const cdOffset = dataView.getUint32(eocdOffset + 16, true);
    const cdSize = dataView.getUint32(eocdOffset + 12, true);
    if (totalRecords === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff) {
      throw new Error("ZIP64 archives are not supported.");
    }

    if (cdOffset + cdSize > searchStart + eocdOffset || cdSize < 46 * totalRecords) {
      throw new Error("Corrupt ZIP: the central directory is outside the file or too small for its records.");
    }

    // Read the Central Directory
    const cdBuffer = await this.reader.read(cdOffset, cdSize);
    const cdView = new DataView(cdBuffer.buffer, cdBuffer.byteOffset, cdBuffer.byteLength);

    let offset = 0;
    const textDecoder = new TextDecoder();

    for (let i = 0; i < totalRecords; i++) {
      if (offset + 46 > cdBuffer.length || cdView.getUint32(offset, true) !== 0x02014b50) {
        throw new Error(`Central Directory record signature mismatch at index ${i}`);
      }

      const compressionMethod = cdView.getUint16(offset + 10, true);
      const crc = cdView.getUint32(offset + 16, true);
      const compressedSize = cdView.getUint32(offset + 20, true);
      const uncompressedSize = cdView.getUint32(offset + 24, true);
      const filenameLength = cdView.getUint16(offset + 28, true);
      const extraFieldLength = cdView.getUint16(offset + 30, true);
      const fileCommentLength = cdView.getUint16(offset + 32, true);
      const localHeaderOffset = cdView.getUint32(offset + 42, true);

      if (offset + 46 + filenameLength > cdBuffer.length) {
        throw new Error(`Corrupt ZIP: central directory record ${i} runs past the directory`);
      }
      const filenameBuffer = cdBuffer.subarray(offset + 46, offset + 46 + filenameLength);
      const filename = textDecoder.decode(filenameBuffer);

      if (filename.includes('../') || filename.includes('..\\')) {
        throw new Error(`Security Exception: Path traversal detected in ZIP filename: ${filename}`);
      }

      const record = { filename, compressionMethod, crc, compressedSize, uncompressedSize, localHeaderOffset };
      this.records.set(filename, record);
      if (!this.lowerCase.has(filename.toLowerCase())) this.lowerCase.set(filename.toLowerCase(), record);

      offset += 46 + filenameLength + extraFieldLength + fileCommentLength;
    }

    // Entries must not share bytes: aliased names would let one small deflated entry be inflated many times
    const byOffset = [...this.records.values()].sort((a, b) => a.localHeaderOffset - b.localHeaderOffset);
    for (let i = 1; i < byOffset.length; i++) {
      const prev = byOffset[i - 1];
      if (prev.localHeaderOffset + 30 + prev.compressedSize > byOffset[i].localHeaderOffset) {
        throw new Error(`Corrupt ZIP: entries ${prev.filename} and ${byOffset[i].filename} overlap`);
      }
    }
  }

  has(filename: string): boolean {
    return this.records.has(filename) || this.lowerCase.has(filename.toLowerCase());
  }

  getFiles(): string[] {
    return Array.from(this.records.keys());
  }

  getRecord(filename: string): ZipRecord {
    const record = this.records.get(filename) ?? this.lowerCase.get(filename.toLowerCase());
    if (!record) throw new Error(`File ${filename} not found in ZIP.`);
    return record;
  }

  // Raw (still compressed) bytes of an entry, for pass-through copying.
  async extractRawStream(filename: string): Promise<ReadableStream<Uint8Array>> {
    const record = this.getRecord(filename);
    if (record.compressedSize === 0xffffffff || record.localHeaderOffset === 0xffffffff) {
      throw new Error(`ZIP64 entry not supported: ${filename}`);
    }

    // The local header has a variable length filename and extra field.
    if (record.localHeaderOffset + 30 > this.reader.size) {
      throw new Error(`Corrupt ZIP: entry ${filename} points past the end of the file`);
    }
    const headerBuffer = await this.reader.read(record.localHeaderOffset, 30);
    const headerView = new DataView(headerBuffer.buffer, headerBuffer.byteOffset, headerBuffer.byteLength);
    if (headerView.getUint32(0, true) !== 0x04034b50) {
      throw new Error(`Local File Header signature mismatch for ${filename}`);
    }
    const dataOffset = record.localHeaderOffset + 30 + headerView.getUint16(26, true) + headerView.getUint16(28, true);
    if (dataOffset + record.compressedSize > this.reader.size) {
      throw new Error(`Corrupt ZIP: entry ${filename} extends past the end of the file`);
    }
    return this.reader.stream(dataOffset, record.compressedSize);
  }

  async extractStream(filename: string): Promise<ReadableStream<Uint8Array>> {
    const record = this.getRecord(filename);
    if (record.compressionMethod !== 0 && record.compressionMethod !== 8) {
      throw new Error(`Unsupported compression method ${record.compressionMethod} for ${filename}`);
    }
    const stream = await this.extractRawStream(filename);
    const data = record.compressionMethod === 0 ? stream
      : stream.pipeThrough(new DecompressionStream('deflate-raw') as any as TransformStream<Uint8Array, Uint8Array>);

    // Node reports bad deflate data as a bare TypeError; name the entry instead
    const inflated = data.getReader();
    let total = 0;
    let crc = 0xffffffff;
    return new ReadableStream<Uint8Array>({
      async pull(controller) {
        try {
          const { done, value } = await inflated.read();
          if (done) {
            // Changed bytes must fail, not read as different data
            if (total !== record.uncompressedSize || ((crc ^ 0xffffffff) >>> 0) !== record.crc) {
              return controller.error(new Error(`Corrupt ZIP entry ${filename}: its data does not match the size and CRC-32 in the directory`));
            }
            return controller.close();
          }
          crc = crc32Update(crc, value);
          // The directory states each entry's size; inflating past it means a crafted entry (a zip bomb)
          total += value.length;
          if (total > record.uncompressedSize) {
            await inflated.cancel();
            return controller.error(new Error(`Corrupt ZIP entry ${filename}: inflates past its stated size of ${record.uncompressedSize} bytes`));
          }
          controller.enqueue(value);
        } catch (err: any) {
          const detail = err?.cause?.message || err?.message || String(err);
          controller.error(new Error(`Corrupt ZIP entry ${filename}: ${detail}`, { cause: err }));
        }
      },
      cancel(reason) {
        return inflated.cancel(reason);
      }
    });
  }
}
