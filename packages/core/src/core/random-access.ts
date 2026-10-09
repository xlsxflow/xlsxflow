export interface RandomAccessReader {
  size: number;
  read(offset: number, length: number): Promise<Uint8Array>;
  stream(offset: number, length: number): ReadableStream<Uint8Array>;
  close(): Promise<void>;
}

// Offsets come from the file itself, so a corrupt one must fail here, not allocate or crash
function checkRange(offset: number, length: number, size: number) {
  if (!(offset >= 0 && length >= 0 && offset + length <= size)) {
    throw new Error(`Corrupt ZIP: read of ${length} bytes at ${offset} is outside the ${size}-byte file`);
  }
}

// Built-in adapter for Web Blob/File
export function createBlobReader(blob: Blob): RandomAccessReader {
  return {
    size: blob.size,
    async read(offset: number, length: number): Promise<Uint8Array> {
      checkRange(offset, length, blob.size);
      const slice = blob.slice(offset, offset + length);
      return new Uint8Array(await slice.arrayBuffer());
    },
    stream(offset: number, length: number): ReadableStream<Uint8Array> {
      // Bun's Blob.slice(start, end).stream() runs on to the end of the original blob, so stop at length
      let left = length;
      return blob.slice(offset, offset + length).stream().pipeThrough(new TransformStream<Uint8Array, Uint8Array>({
        start(controller) { if (left <= 0) controller.terminate(); },
        transform(chunk, controller) {
          controller.enqueue(chunk.length > left ? chunk.subarray(0, left) : chunk);
          left -= chunk.length;
          if (left <= 0) controller.terminate();
        }
      }));
    },
    async close() {}
  };
}

// Built-in adapter for Node.js fs.promises
// Dynamic, bundler-ignored imports so browser bundles don't try to resolve Node built-ins
export async function createFileReader(filePath: string): Promise<RandomAccessReader> {
  const fs = await import(/* webpackIgnore: true */ /* turbopackIgnore: true */ 'fs/promises');
  const { createReadStream } = await import(/* webpackIgnore: true */ /* turbopackIgnore: true */ 'fs');
  const { Readable } = await import(/* webpackIgnore: true */ /* turbopackIgnore: true */ 'stream');
  const stat = await fs.stat(filePath);
  const fd = await fs.open(filePath, 'r');

  return {
    size: stat.size,
    async read(offset: number, length: number): Promise<Uint8Array> {
      checkRange(offset, length, stat.size);
      const buffer = new Uint8Array(length);
      let filled = 0;
      while (filled < length) {
        const { bytesRead } = await fd.read(buffer, filled, length - filled, offset + filled);
        if (bytesRead === 0) throw new Error(`Unexpected end of file reading ${filePath} at ${offset + filled}`);
        filled += bytesRead;
      }
      return buffer;
    },
    stream(offset: number, length: number): ReadableStream<Uint8Array> {
      if (length === 0) return new ReadableStream({ start(c) { c.close(); } });
      // Opens its own handle, so streams outlive close(). toWeb keeps backpressure.
      return Readable.toWeb(createReadStream(filePath, { start: offset, end: offset + length - 1 })) as ReadableStream<Uint8Array>;
    },
    async close() {
      await fd.close();
    }
  };
}
