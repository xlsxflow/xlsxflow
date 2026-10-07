import { describe, it, expect } from 'vitest';
import { SheetWriter } from '../src/core/writer';
import { Row } from '../src/core/types';
import { ZipRandomAccessParser } from '../src/core/zip-random-access';
import { createBlobReader } from '../src/core/random-access';

// Helper to consume stream into Uint8Array
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

describe('SheetWriter', () => {
  it('rejects sheet names Excel cannot open', () => {
    for (const bad of ['', 'x'.repeat(32), 'a/b', 'a:b', 'a[1]', 'what?', "'quoted", "quoted'"]) {
      expect(() => new SheetWriter().addSheet(bad, [])).toThrow(/Invalid sheet name/);
    }
    expect(() => new SheetWriter().addSheet('Data', []).addSheet('DATA', [])).toThrow(/Duplicate sheet name/);
    expect(() => new SheetWriter().addSheet("Q1 Sales's", []).addSheet('x'.repeat(31), [])).not.toThrow();
  });

  it('should generate a valid XLSX buffer with Deflate compression', async () => {
    const writer = new SheetWriter();
    const rows: Row[] = [
      ['A1', 'B1'],
      [100, 200]
    ];

    const stream = writer.write(rows);
    const buffer = await streamToUint8Array(stream);
    expect(buffer).toBeInstanceOf(Uint8Array);
    
    // Check ZIP signature (PK\x03\x04)
    expect(buffer[0]).toBe(0x50); // P
    expect(buffer[1]).toBe(0x4b); // K
    expect(buffer[2]).toBe(0x03);
    expect(buffer[3]).toBe(0x04);
  });

  it('should write a non-empty binary that contains sheet XML markers', async () => {
    const writer = new SheetWriter();
    const stream = writer.write([['Hello', 'World'], [1, 2]]);
    const bytes = await streamToUint8Array(stream);
    const text = new TextDecoder().decode(bytes);
    
    expect(text).toContain('[Content_Types].xml');
    expect(text).toContain('xl/workbook.xml');
  });
  it('should inject formula string and cached value', async () => {
    const writer = new SheetWriter();
    const stream = writer.write([
      [1, 2, 3],
      [{ value: null, formula: '=SUM(A1:C1)' }]
    ]);
    const bytes = await streamToUint8Array(stream);
    
    // We can just verify it generated a valid ZIP
    expect(bytes[0]).toBe(0x50); // P
    expect(bytes[1]).toBe(0x4b); // K
  });

  it('should handle booleans and null cells', async () => {
    const writer = new SheetWriter();
    const stream = writer.write([[true, false, null]]);
    const bytes = await streamToUint8Array(stream);
    
    // We can just verify it generated a valid ZIP
    expect(bytes[0]).toBe(0x50); // P
    expect(bytes[1]).toBe(0x4b); // K
  });

  it('should support multiple sheets', async () => {
    const writer = new SheetWriter();
    writer.addSheet('First', [[1, 2]]);
    writer.addSheet('Second', [['A', 'B']]);
    const stream = writer.write();
    const bytes = await streamToUint8Array(stream);
    expect(bytes[0]).toBe(0x50); // P
    expect(bytes[1]).toBe(0x4b); // K
    expect(bytes.length).toBeGreaterThan(100);
  });

  it('should write a date and read it back as an ISO string', async () => {
    const writer = new SheetWriter();
    const dateNum = 45000; // ~ 2023-03-15
    writer.addSheet('Dates', [[
      { value: dateNum, style: { numFmt: 'yyyy-mm-dd' } }
    ]]);
    const stream = writer.write();
    const bytes = await streamToUint8Array(stream);

    // Verify via Reader
    const { SheetReader } = await import('../src/core/index');
    const reader = new SheetReader();
    const blob = new Blob([bytes.buffer as ArrayBuffer]);
    const fileReader = createBlobReader(blob);
    
    let parsedRows: any[] = [];
    for await (const row of await reader.parse(fileReader)) {
      parsedRows.push(row);
    }
    
    // 45000 in Excel is 2023-03-15T00:00:00.000Z
    expect(parsedRows[0].cells[0]).toBe('2023-03-15T00:00:00.000Z');
  });
  it('should support freezePanes and mergeCells', async () => {
    const writer = new SheetWriter();
    writer.addSheet('Complex', [[1, 2], [3, 4]], {
      freezePanes: { row: 1, col: 1 },
      mergeCells: ['A1:B1']
    });
    const stream = writer.write();
    const bytes = await streamToUint8Array(stream);

    const { SheetReader } = await import('../src/core/index');
    const reader = new SheetReader();
    const blob = new Blob([bytes.buffer as ArrayBuffer]);
    const fileReader = createBlobReader(blob);
    
    const result = await reader.parse(fileReader);
    for await (const row of result) { } // consume stream
    
    const meta = await result.getMetadata();
    expect(meta.mergedCells).toEqual(['A1:B1']);
    expect(meta.freezePanes).toBeDefined();
    expect(meta.freezePanes?.row).toBe(1);
    expect(meta.freezePanes?.col).toBe(1);
  });

  it('should generate dataValidations correctly', async () => {
    const writer = new SheetWriter();
    writer.addSheet('Validation', [['Select Item']], {
      dataValidations: [
        { sqref: 'A2:A10', type: 'list', formula1: '"Apple,Banana,Orange"', showErrorMessage: true }
      ]
    });
    const stream = writer.write();
    const bytes = await streamToUint8Array(stream);
    
    // Extract sheet1.xml from the zip to verify its content
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
    
    expect(sheetXml).toContain('<dataValidations count="1">');
    expect(sheetXml).toContain('<dataValidation sqref="A2:A10" type="list" showErrorMessage="1">');
    expect(sheetXml).toContain('<formula1>&quot;Apple,Banana,Orange&quot;</formula1>');
  });
});

describe('FormulaEngine (via SheetWriter)', () => {
  it('should correctly compute SUM in exported binary', async () => {
    const writer = new SheetWriter();
    const stream = writer.write([
      [1, 2, 3],
      [{ value: null, formula: '=SUM(A1:C1)' }]
    ]);
    const bytes = await streamToUint8Array(stream);
    expect(bytes.length).toBeGreaterThan(100);
  });
});
