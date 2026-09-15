import { describe, it, expect } from 'vitest';
import { SheetReader } from '../src/core/index';
import { createByteLimitStream } from '../src/core/zip-stream';
import { createXmlStreamParser } from '../src/core/xml-stream';
import { parseWorksheet } from '../src/core/worksheet-parser';

describe('SheetForge Fuzzer and Security', () => {
  it('should throw when maxUncompressedBytes is exceeded', async () => {
    // Generate a massive string of spaces
    const massiveData = new Uint8Array(10 * 1024 * 1024); // 10MB
    massiveData.fill(32);

    const stream = new ReadableStream({
      start(controller) {
        controller.enqueue(massiveData);
        controller.close();
      }
    });

    const limitedStream = stream.pipeThrough(createByteLimitStream(5 * 1024 * 1024)); // 5MB limit
    const reader = limitedStream.getReader();

    let errorThrown = false;
    try {
      while (true) {
        const { done } = await reader.read();
        if (done) break;
      }
    } catch (e: any) {
      errorThrown = true;
      expect(e.message).toContain('Security Error: Stream exceeded maximum uncompressed size');
    }
    
    expect(errorThrown).toBe(true);
  });

  it('xml parser should not crash on incomplete tags', async () => {
    const garbageXml = new TextEncoder().encode('<worksheet><sheetData><row><c t="inlineStr"><is><t>Hello');
    const stream = new ReadableStream({
      start(controller) {
        controller.enqueue(garbageXml);
        controller.close();
      }
    });

    const xmlStream = stream.pipeThrough(createXmlStreamParser());
    const parseResult = parseWorksheet(xmlStream, new Map(), new Map());
    
    // It should yield whatever it can without infinite looping
    let count = 0;
    try {
      for await (const row of parseResult) {
        count++;
      }
      // Wait for metadata
      await parseResult.getMetadata();
    } catch (e) {
      // It might throw due to unexpected end, which is fine
    }
    // As long as it finishes execution, fuzzer passes.
    expect(count).toBe(0);
  });
});
