const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (c >>> 1) ^ 0xedb88320 : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

const MAX_U32 = 0xffffffff;

// No ZIP64 support: fail loudly instead of writing a corrupt archive.
function assertNoZip64(value: number, what: string) {
  if (value > MAX_U32) throw new Error(`ZIP64 not supported: ${what} exceeds 4 GiB.`);
}

interface CdEntry {
  filename: Uint8Array;
  offset: number;
  uncompressedSize: number;
  compressedSize: number;
  crc: number;
  flags: number;
  method: number;
}

// A true Single-Pass Streaming ZIP Writer
export class ZipStreamWriter {
  private cdEntries: CdEntry[] = [];
  private offset = 0;
  private streamController!: ReadableStreamDefaultController<Uint8Array>;
  private resumePull: (() => void) | null = null;
  private cancelled = false;
  public stream: ReadableStream<Uint8Array>;
  private textEncoder = new TextEncoder();

  constructor(highWaterMarkBytes = 1 << 20) {
    this.stream = new ReadableStream<Uint8Array>({
      start: (controller) => {
        this.streamController = controller;
      },
      pull: () => {
        this.resumePull?.();
        this.resumePull = null;
      },
      // Consumer gone: unblock any pending write so the producer can see the error.
      cancel: () => {
        this.cancelled = true;
        this.resumePull?.();
        this.resumePull = null;
      }
    }, new ByteLengthQueuingStrategy({ highWaterMark: highWaterMarkBytes }));
  }

  // Adds a file to the zip. `inputStream` MUST be raw uncompressed data.
  async addFile(filenameStr: string, inputStream: ReadableStream<Uint8Array>): Promise<void> {
    const filename = this.textEncoder.encode(filenameStr);
    const startOffset = this.offset;
    await this.pushChunk(this.localHeader(filename, 0x0008, 8, 0, 0, 0));

    // Stream data, tracking sizes and CRC32
    let uncompressedSize = 0;
    let crc = 0xffffffff;

    const crcStream = new TransformStream<Uint8Array, Uint8Array>({
      transform: (chunk, controller) => {
        uncompressedSize += chunk.length;
        for (let i = 0; i < chunk.length; i++) {
          crc = CRC_TABLE[(crc ^ chunk[i]) & 0xff] ^ (crc >>> 8);
        }
        controller.enqueue(chunk);
      }
    });

    let compressedSize = 0;
    const reader = inputStream
      .pipeThrough(crcStream)
      .pipeThrough(new CompressionStream('deflate-raw') as any as TransformStream<Uint8Array, Uint8Array>)
      .getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      compressedSize += value.length;
      await this.pushChunk(value);
    }

    crc = (crc ^ 0xffffffff) >>> 0;
    assertNoZip64(uncompressedSize, filenameStr);
    assertNoZip64(compressedSize, filenameStr);

    // Data Descriptor
    const desc = new Uint8Array(16);
    const descView = new DataView(desc.buffer);
    descView.setUint32(0, 0x08074b50, true);
    descView.setUint32(4, crc, true);
    descView.setUint32(8, compressedSize, true);
    descView.setUint32(12, uncompressedSize, true);
    await this.pushChunk(desc);

    this.cdEntries.push({ filename, offset: startOffset, uncompressedSize, compressedSize, crc, flags: 0x0008, method: 8 });
  }

  // Adds a file that is ALREADY compressed (pass-through for the Editor)
  async addCompressedFile(filenameStr: string, compressedStream: ReadableStream<Uint8Array>, uncompressedSize: number, compressedSize: number, crc: number, method = 8): Promise<void> {
    const filename = this.textEncoder.encode(filenameStr);
    const startOffset = this.offset;
    await this.pushChunk(this.localHeader(filename, 0, method, crc, compressedSize, uncompressedSize));

    const reader = compressedStream.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      await this.pushChunk(value);
    }

    this.cdEntries.push({ filename, offset: startOffset, uncompressedSize, compressedSize, crc, flags: 0, method });
  }

  async close(): Promise<void> {
    if (this.cdEntries.length > 0xffff) throw new Error('ZIP64 not supported: more than 65535 entries.');
    const cdStartOffset = this.offset;

    for (const entry of this.cdEntries) {
      const cd = new Uint8Array(46 + entry.filename.length);
      const view = new DataView(cd.buffer);
      view.setUint32(0, 0x02014b50, true);
      view.setUint16(4, 20, true);            // version made by
      view.setUint16(6, 20, true);            // version needed
      view.setUint16(8, entry.flags, true);
      view.setUint16(10, entry.method, true);
      view.setUint32(16, entry.crc, true);
      view.setUint32(20, entry.compressedSize, true);
      view.setUint32(24, entry.uncompressedSize, true);
      view.setUint16(28, entry.filename.length, true);
      view.setUint32(42, entry.offset, true);
      cd.set(entry.filename, 46);
      await this.pushChunk(cd);
    }

    const cdSize = this.offset - cdStartOffset;
    assertNoZip64(this.offset, 'archive');

    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(8, this.cdEntries.length, true);
    eocdView.setUint16(10, this.cdEntries.length, true);
    eocdView.setUint32(12, cdSize, true);
    eocdView.setUint32(16, cdStartOffset, true);
    await this.pushChunk(eocd);

    this.streamController.close();
  }

  // Propagate a producer failure to whoever is reading `stream`.
  error(err: unknown) {
    try { this.streamController.error(err); } catch { /* already closed/errored */ }
  }

  private localHeader(filename: Uint8Array, flags: number, method: number, crc: number, compressedSize: number, uncompressedSize: number): Uint8Array {
    assertNoZip64(this.offset, 'archive');
    const header = new Uint8Array(30 + filename.length);
    const view = new DataView(header.buffer);
    view.setUint32(0, 0x04034b50, true);
    view.setUint16(4, 20, true);
    view.setUint16(6, flags, true);
    view.setUint16(8, method, true);
    view.setUint32(14, crc, true);
    view.setUint32(18, compressedSize, true);
    view.setUint32(22, uncompressedSize, true);
    view.setUint16(26, filename.length, true);
    header.set(filename, 30);
    return header;
  }

  // Enqueue and wait while the consumer's queue is full, so memory stays bounded.
  private async pushChunk(chunk: Uint8Array) {
    this.streamController.enqueue(chunk);
    this.offset += chunk.length;
    while ((this.streamController.desiredSize ?? 1) <= 0) {
      if (this.cancelled) throw new Error('ZIP stream cancelled by consumer.');
      await new Promise<void>(resolve => { this.resumePull = resolve; });
    }
  }
}
