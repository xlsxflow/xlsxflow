import { describe, it, expect } from 'vitest';
import { sheetToJson, streamToCsv } from '../src/core/utils';
import { SheetWriter } from '../src/core/writer';
import { XlsxFlow } from '../src/index';
import { createBlobReader, RandomAccessReader } from '../src/core/random-access';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';

async function streamToUint8Array(stream: ReadableStream<Uint8Array>): Promise<Uint8Array> {
  const reader = stream.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    chunks.push(value);
    total += value.length;
  }
  const result = new Uint8Array(total);
  let offset = 0;
  for (const c of chunks) {
    result.set(c, offset);
    offset += c.length;
  }
  return result;
}

describe('DX Utilities', () => {
  it('should map stream to JSON objects', async () => {
    const writer = new SheetWriter();
    writer.addSheet('Data', [
      ['ID', 'Name', 'Active'],
      [1, 'Alice', true],
      [2, 'Bob', false]
    ]);
    const stream = writer.write();
    const bytes = await streamToUint8Array(stream);
    
    // Convert to mock file reader
    const blob = new Blob([bytes.buffer as ArrayBuffer]);
    const fileReader = createBlobReader(blob);
    
    // Parse
    const { SheetReader } = await import('../src/core/index');
    const reader = new SheetReader();
    const parseResult = await reader.parse(fileReader);
    
    const json = await sheetToJson<{ ID: number; Name: string; Active: boolean }>(parseResult);
    
    expect(json.length).toBe(2);
    expect(json[0]).toEqual({ ID: 1, Name: 'Alice', Active: true });
    expect(json[1]).toEqual({ ID: 2, Name: 'Bob', Active: false });
  });

  it('should export stream to CSV', async () => {
    const writer = new SheetWriter();
    writer.addSheet('Data', [
      ['ID', 'Name', 'Desc'],
      [1, 'Alice', 'Hello, World'],
      [2, 'Bob', 'Quote: "Yes"']
    ]);
    const stream = writer.write();
    const bytes = await streamToUint8Array(stream);
    
    const blob = new Blob([bytes.buffer as ArrayBuffer]);
    const fileReader = createBlobReader(blob);
    
    const { SheetReader } = await import('../src/core/index');
    const reader = new SheetReader();
    const parseResult = await reader.parse(fileReader);
    
    const csv = await streamToCsv(parseResult);
    
    expect(csv).toContain('ID,Name,Desc');
    expect(csv).toContain('1,Alice,"Hello, World"');
    expect(csv).toContain('2,Bob,"Quote: ""Yes"""');
  });

  it('should auto-fit columns', async () => {
    const writer = new SheetWriter();
    writer.addSheet('Data', [
      ['Short', 'A very long string indeed']
    ], { autoFitColumns: true });
    
    const stream = writer.write();
    const bytes = await streamToUint8Array(stream);
    
    const blob = new Blob([bytes.buffer as ArrayBuffer]);
    const fileReader = createBlobReader(blob);
    const zip = new ZipRandomAccessParser(fileReader);
    await zip.parseCentralDirectory();
    
    const sheetStream = await zip.extractStream('xl/worksheets/sheet1.xml');
    const textReader = sheetStream.pipeThrough(new TextDecoderStream() as any).getReader();
    let sheetXml = '';
    while (true) {
      const res = await textReader.read();
      if (res.done) break;
      sheetXml += res.value;
    }
    
    // Ensure cols were written
    expect(sheetXml).toContain('<cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="30" customWidth="1"/></cols>');
  });
});
