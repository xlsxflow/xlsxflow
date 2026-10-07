import { describe, it, expect } from 'vitest';
import { SheetReader } from '../src/core/index';
import { createXmlStreamParser } from '../src/core/xml-stream';
import { parseWorksheet } from '../src/core/worksheet-parser';

describe('XlsxFlow Fuzzer and Security', () => {
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

  it('xml parser output does not depend on chunk boundaries', async () => {
    const xml = '<?xml version="1.0"?><worksheet><sheetData>' +
      Array.from({ length: 200 }, (_, r) =>
        `<row r="${r + 1}" spans="1:3"><c r="A${r + 1}"><v>${r}</v></c><c r="B${r + 1}" t="inlineStr"><is><t>a &amp; b ${r}</t></is></c><!-- c --><c r="C${r + 1}" t="str"><v><![CDATA[x<y]]></v></c></row>`
      ).join('') + '</sheetData></worksheet>';
    const bytes = new TextEncoder().encode(xml);

    const tokens = async (chunkSizes: () => number) => {
      const out: any[] = [];
      const stream = new ReadableStream<Uint8Array>({
        start(c) {
          for (let i = 0; i < bytes.length;) { const n = chunkSizes(); c.enqueue(bytes.slice(i, i + n)); i += n; }
          c.close();
        }
      });
      const reader = stream.pipeThrough(createXmlStreamParser()).getReader();
      for (let t = await reader.read(); !t.done; t = await reader.read()) {
        // Adjacent text tokens may legitimately split; merge them before comparing
        const last = out[out.length - 1];
        if (t.value.type === 'text' && last?.type === 'text') last.value += t.value.value;
        else out.push({ ...t.value });
      }
      return out;
    };

    const whole = await tokens(() => bytes.length);
    let seed = 7;
    const rand = () => (seed = (seed * 1103515245 + 12345) % 2147483648) % 97 + 1;
    expect(await tokens(() => 1)).toEqual(whole);
    expect(await tokens(rand)).toEqual(whole);
  });
});
