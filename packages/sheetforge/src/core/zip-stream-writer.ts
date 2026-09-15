// A true Single-Pass Streaming ZIP Writer
export class ZipStreamWriter {
  private cdEntries: { filename: Uint8Array; offset: number; uncompressedSize: number; compressedSize: number; crc: number }[] = [];
  private offset = 0;
  private streamController!: ReadableStreamDefaultController<Uint8Array>;
  public stream: ReadableStream<Uint8Array>;
  private textEncoder = new TextEncoder();

  constructor() {
    this.stream = new ReadableStream({
      start: (controller) => {
        this.streamController = controller;
      }
    });
  }

  // Adds a file to the zip. `inputStream` MUST be raw uncompressed data.
  async addFile(filenameStr: string, inputStream: ReadableStream<Uint8Array>): Promise<void> {
    const filename = this.textEncoder.encode(filenameStr);
    
    // 1. Write Local File Header with Data Descriptor flag (Bit 3)
    const header = new Uint8Array(30 + filename.length);
    const view = new DataView(header.buffer);
    view.setUint32(0, 0x04034b50, true);  // signature
    view.setUint16(4, 20, true);          // version needed
    view.setUint16(6, 0x0008, true);      // flags: Bit 3 = Data Descriptor
    view.setUint16(8, 8, true);           // compression (Deflate)
    view.setUint16(10, 0, true);          // mod time
    view.setUint16(12, 0, true);          // mod date
    view.setUint32(14, 0, true);          // CRC (0 for now)
    view.setUint32(18, 0, true);          // Compressed Size (0 for now)
    view.setUint32(22, 0, true);          // Uncompressed Size (0 for now)
    view.setUint16(26, filename.length, true);
    view.setUint16(28, 0, true);          // extra field length
    header.set(filename, 30);

    const startOffset = this.offset;
    this.pushChunk(header);

    // 2. Stream Data, track sizes and CRC32
    let uncompressedSize = 0;
    let crc = 0xffffffff;

    const crcStream = new TransformStream<Uint8Array, Uint8Array>({
      transform: (chunk, controller) => {
        uncompressedSize += chunk.length;
        for (const byte of chunk) {
          crc ^= byte;
          for (let i = 0; i < 8; i++) {
            crc = (crc & 1) ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
          }
        }
        controller.enqueue(chunk);
      }
    });

    let compressedSize = 0;
    const sizeTrackerStream = new TransformStream<Uint8Array, Uint8Array>({
      transform: (chunk, controller) => {
        compressedSize += chunk.length;
        controller.enqueue(chunk);
      }
    });

    const compressed = inputStream
      .pipeThrough(crcStream)
      .pipeThrough(new CompressionStream('deflate-raw') as any)
      .pipeThrough(sizeTrackerStream);

    const reader = compressed.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      this.pushChunk(value as Uint8Array);
    }

    crc = (crc ^ 0xffffffff) >>> 0;

    // 3. Write Data Descriptor
    const desc = new Uint8Array(16);
    const descView = new DataView(desc.buffer);
    descView.setUint32(0, 0x08074b50, true); // signature
    descView.setUint32(4, crc, true);
    descView.setUint32(8, compressedSize, true);
    descView.setUint32(12, uncompressedSize, true);
    this.pushChunk(desc);

    // 4. Save metadata for Central Directory
    this.cdEntries.push({ filename, offset: startOffset, uncompressedSize, compressedSize, crc });
  }

  // Adds a file that is ALREADY compressed (useful for pass-through Editor)
  async addCompressedFile(filenameStr: string, compressedStream: ReadableStream<Uint8Array>, uncompressedSize: number, compressedSize: number, crc: number): Promise<void> {
    const filename = this.textEncoder.encode(filenameStr);
    
    // Write standard Local File Header (no Bit 3 because we know sizes)
    const header = new Uint8Array(30 + filename.length);
    const view = new DataView(header.buffer);
    view.setUint32(0, 0x04034b50, true);  
    view.setUint16(4, 20, true);          
    view.setUint16(6, 0, true);           // No Data Descriptor needed
    view.setUint16(8, 8, true);           
    view.setUint16(10, 0, true);          
    view.setUint16(12, 0, true);          
    view.setUint32(14, crc, true);        
    view.setUint32(18, compressedSize, true); 
    view.setUint32(22, uncompressedSize, true); 
    view.setUint16(26, filename.length, true);
    view.setUint16(28, 0, true);          
    header.set(filename, 30);

    const startOffset = this.offset;
    this.pushChunk(header);

    const reader = compressedStream.getReader();
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      this.pushChunk(value as Uint8Array);
    }

    this.cdEntries.push({ filename, offset: startOffset, uncompressedSize, compressedSize, crc });
  }

  async close(): Promise<void> {
    const cdStartOffset = this.offset;
    
    // Central Directory
    for (const entry of this.cdEntries) {
      const cd = new Uint8Array(46 + entry.filename.length);
      const view = new DataView(cd.buffer);
      view.setUint32(0, 0x02014b50, true);   
      view.setUint16(4, 20, true);            
      view.setUint16(6, 20, true);            
      view.setUint16(8, 0, true);             
      view.setUint16(10, 8, true);            
      view.setUint16(12, 0, true);            
      view.setUint16(14, 0, true);            
      view.setUint32(16, entry.crc, true);    
      view.setUint32(20, entry.compressedSize, true); 
      view.setUint32(24, entry.uncompressedSize, true); 
      view.setUint16(28, entry.filename.length, true); 
      view.setUint16(30, 0, true);            
      view.setUint16(32, 0, true);            
      view.setUint16(34, 0, true);            
      view.setUint16(36, 0, true);            
      view.setUint32(38, 0, true);            
      view.setUint32(42, entry.offset, true); 
      cd.set(entry.filename, 46);
      this.pushChunk(cd);
    }

    const cdEndOffset = this.offset;
    const cdSize = cdEndOffset - cdStartOffset;

    // End of central directory record
    const eocd = new Uint8Array(22);
    const eocdView = new DataView(eocd.buffer);
    eocdView.setUint32(0, 0x06054b50, true);
    eocdView.setUint16(4, 0, true);
    eocdView.setUint16(6, 0, true);
    eocdView.setUint16(8, this.cdEntries.length, true);
    eocdView.setUint16(10, this.cdEntries.length, true);
    eocdView.setUint32(12, cdSize, true);
    eocdView.setUint32(16, cdStartOffset, true);
    eocdView.setUint16(20, 0, true);
    this.pushChunk(eocd);

    this.streamController.close();
  }

  private pushChunk(chunk: Uint8Array) {
    this.streamController.enqueue(chunk);
    this.offset += chunk.length;
  }
}
