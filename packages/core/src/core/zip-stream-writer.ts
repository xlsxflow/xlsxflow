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
  return (crc32Update(0xffffffff, bytes) ^ 0xffffffff) >>> 0;
}

// Older runtimes lack CompressionStream or its 'deflate-raw' format; say what is needed instead of a bare ReferenceError
export function deflateRaw(kind: 'compress' | 'decompress'): TransformStream<Uint8Array, Uint8Array> {
  try {
    return (kind === 'compress' ? new CompressionStream('deflate-raw') : new DecompressionStream('deflate-raw')) as any;
  } catch {
    throw new Error(`XlsxFlow needs ${kind === 'compress' ? 'CompressionStream' : 'DecompressionStream'} with 'deflate-raw': Chrome 103+, Firefox 113+, Safari 16.4+, Node 20.12+, Deno, Bun or Cloudflare Workers`);
  }
}

// Running CRC-32 over chunks: start at 0xffffffff, finish with (crc ^ 0xffffffff) >>> 0
export function crc32Update(crc: number, bytes: Uint8Array): number {
  for (let i = 0; i < bytes.length; i++) crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  return crc;
}

const MAX_U32 = 0xffffffff;

// General-purpose flag bit 11: the entry name is UTF-8 (set only when it isn't plain ASCII)
const utf8Flag = (name: Uint8Array) => name.some(b => b > 0x7f) ? 0x0800 : 0;

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
  private waiting: (() => void)[] = [];
  private cancelled = false;
  public stream: ReadableStream<Uint8Array>;
  private textEncoder = new TextEncoder();

  constructor(highWaterMarkBytes = 1 << 20) {
    this.stream = new ReadableStream<Uint8Array>({
      start: (controller) => {
        this.streamController = controller;
      },
      pull: () => this.wake(),
      // Consumer gone: unblock any pending write so the producer can see the error.
      cancel: () => {
        this.cancelled = true;
        this.wake();
      }
    }, new ByteLengthQueuingStrategy({ highWaterMark: highWaterMarkBytes }));
  }

  // Adds a file to the zip. `inputStream` MUST be raw uncompressed data.
  async addFile(filenameStr: string, inputStream: ReadableStream<Uint8Array>): Promise<void> {
    const filename = this.textEncoder.encode(filenameStr);
    const flags = 0x0008 | utf8Flag(filename);
    const startOffset = this.offset;
    await this.pushChunk(this.localHeader(filename, flags, 8, 0, 0, 0));

    // Stream data, tracking sizes and CRC32
    let uncompressedSize = 0;
    let crc = 0xffffffff;

    const compressor = deflateRaw('compress');
    const writer = compressor.writable.getWriter();
    const reader = compressor.readable.getReader();
    const input = inputStream.getReader();
    // Node's and Bun's CompressionStream accept thousands of writes without backpressure, so only a few
    // chunks are fed ahead: a write resolves once the chunk is compressed, which waits while nobody reads
    // the output. Rows are then only pulled as fast as the ZIP is consumed.
    // A few writes in flight keep the compressor busy while the next chunk is built
    const inFlight: Promise<void>[] = [];
    const feed = (async () => {
      try {
        while (true) {
          const { done, value } = await input.read();
          if (done) break;
          await this.roomInQueue();
          uncompressedSize += value.length;
          crc = crc32Update(crc, value);
          const written = writer.write(value);
          written.catch(() => {}); // after a failure, writes nobody awaits any more must not go unhandled
          inFlight.push(written);
          if (inFlight.length >= 4) await inFlight.shift();
        }
        await Promise.all(inFlight);
        await writer.close();
      } catch (err) {
        // Stop the source too (a row generator's finally runs, a database cursor closes)
        await input.cancel(err).catch(() => {});
        await writer.abort(err).catch(() => {});
        throw err;
      }
    })();

    let compressedSize = 0;
    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        compressedSize += value.length;
        await this.pushChunk(value);
      }
      await feed;
    } catch (err) {
      await reader.cancel(err).catch(() => {});
      await input.cancel(err).catch(() => {});
      await feed.catch(() => {});
      throw err;
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

    this.cdEntries.push({ filename, offset: startOffset, uncompressedSize, compressedSize, crc, flags, method: 8 });
  }

  // Adds a file that is ALREADY compressed (pass-through for the Editor)
  async addCompressedFile(filenameStr: string, compressedStream: ReadableStream<Uint8Array>, uncompressedSize: number, compressedSize: number, crc: number, method = 8): Promise<void> {
    const filename = this.textEncoder.encode(filenameStr);
    const flags = utf8Flag(filename);
    const startOffset = this.offset;
    await this.pushChunk(this.localHeader(filename, flags, method, crc, compressedSize, uncompressedSize));

    const reader = compressedStream.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      await this.pushChunk(value);
    }

    this.cdEntries.push({ filename, offset: startOffset, uncompressedSize, compressedSize, crc, flags, method });
  }

  async close(): Promise<void> {
    // 0xFFFF in the end record means "see the ZIP64 record", which readers then look for
    if (this.cdEntries.length >= 0xffff) throw new Error('ZIP64 not supported: 65535 entries or more.');
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
    if (this.cancelled) throw new Error('ZIP stream cancelled by consumer.');
    this.streamController.enqueue(chunk);
    this.offset += chunk.length;
    await this.roomInQueue();
  }

  private async roomInQueue() {
    while ((this.streamController.desiredSize ?? 1) <= 0) {
      if (this.cancelled) throw new Error('ZIP stream cancelled by consumer.');
      await new Promise<void>(resolve => this.waiting.push(resolve));
    }
    if (this.cancelled) throw new Error('ZIP stream cancelled by consumer.');
  }

  private wake() {
    const waiting = this.waiting;
    this.waiting = [];
    for (const resolve of waiting) resolve();
  }
}
