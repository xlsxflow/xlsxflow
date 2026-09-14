// Implements a forward-only ZIP parser using TransformStream
// It searches for the Local File Header signature (0x04034b50)
export interface ZipEntry {
  filename: string;
  compressionMethod: number;
  stream: ReadableStream<Uint8Array>;
}

export function createZipStreamParser(): TransformStream<Uint8Array, ZipEntry> {
  let buffer = new Uint8Array(0);
  let currentFile: { filename: string, compressedSize: number, controller: ReadableStreamDefaultController<Uint8Array> } | null = null;
  let currentBytesRead = 0;

  return new TransformStream({
    start(controller) {
      // Initialization if needed
    },
    transform(chunk, controller) {
      // Append new chunk to buffer
      const newBuffer = new Uint8Array(buffer.length + chunk.length);
      newBuffer.set(buffer);
      newBuffer.set(chunk, buffer.length);
      buffer = newBuffer;

      // Simple state machine to parse ZIP Local File Headers
      while (buffer.length >= 30) {
        if (currentFile) {
          // We are currently reading file data
          const remaining = currentFile.compressedSize - currentBytesRead;
          if (buffer.length >= remaining) {
            // Finish this file
            currentFile.controller.enqueue(buffer.slice(0, remaining));
            currentFile.controller.close();
            buffer = buffer.slice(remaining);
            currentFile = null;
            currentBytesRead = 0;
            continue;
          } else {
            // Push all available data to current file
            currentFile.controller.enqueue(buffer);
            currentBytesRead += buffer.length;
            buffer = new Uint8Array(0);
            break;
          }
        }

        const dataView = new DataView(buffer.buffer, buffer.byteOffset, buffer.byteLength);
        const signature = dataView.getUint32(0, true);

        if (signature === 0x04034b50) {
          // Local File Header found
          const compressionMethod = dataView.getUint16(8, true);
          const compressedSize = dataView.getUint32(18, true);
          const filenameLength = dataView.getUint16(26, true);
          const extraFieldLength = dataView.getUint16(28, true);
          const headerLength = 30 + filenameLength + extraFieldLength;

          if (buffer.length < headerLength) {
            // Wait for more data to read the full header
            break;
          }

          const filenameBuffer = buffer.slice(30, 30 + filenameLength);
          const filename = new TextDecoder().decode(filenameBuffer);

          buffer = buffer.slice(headerLength);

          let fileController: ReadableStreamDefaultController<Uint8Array>;
          const stream = new ReadableStream<Uint8Array>({
            start(c) {
              fileController = c;
            }
          });

          // Enqueue the new ZipEntry to the consumer
          controller.enqueue({ filename, compressionMethod, stream });

          // Handle 0-byte files instantly
          if (compressedSize === 0) {
            fileController!.close();
            currentFile = null;
          } else {
            currentFile = { filename, compressedSize, controller: fileController! };
          }
        } else if (signature === 0x02014b50) {
          // Central Directory found, we can stop parsing forward
          break;
        } else {
          // If we lose sync or hit something else, advance buffer by 1 to resync
          // (In a robust implementation, this would search for the next signature)
          buffer = buffer.slice(1);
        }
      }
    },
    flush(controller) {
      if (currentFile) {
        currentFile.controller.close();
      }
    }
  });
}
