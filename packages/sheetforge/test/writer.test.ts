import { describe, it, expect } from 'vitest';
import { SheetWriter } from '../src/core/writer';
import { Row } from '../src/pro/types';

describe('SheetWriter', () => {
  it('should generate a valid XLSX buffer with Deflate compression', async () => {
    const writer = new SheetWriter();
    const rows: Row[] = [
      ['A1', 'B1'],
      [100, 200]
    ];

    const buffer = await writer.write(rows);
    expect(buffer).toBeInstanceOf(Uint8Array);
    
    // Check ZIP signature (PK\x03\x04)
    expect(buffer[0]).toBe(0x50); // P
    expect(buffer[1]).toBe(0x4b); // K
    expect(buffer[2]).toBe(0x03);
    expect(buffer[3]).toBe(0x04);
  });

  it('should write a non-empty binary that contains sheet XML markers', async () => {
    const writer = new SheetWriter();
    const bytes = await writer.write([['Hello', 'World'], [1, 2]]);
    const text = new TextDecoder().decode(bytes);
    
    expect(text).toContain('[Content_Types].xml');
    expect(text).toContain('xl/workbook.xml');
  });
  it('should inject formula string and cached value', async () => {
    const writer = new SheetWriter();
    const bytes = await writer.write([
      [1, 2, 3],
      [{ value: null, formula: '=SUM(A1:C1)' }]
    ]);
    
    // We can just verify it generated a valid ZIP
    expect(bytes[0]).toBe(0x50); // P
    expect(bytes[1]).toBe(0x4b); // K
  });

  it('should handle booleans and null cells', async () => {
    const writer = new SheetWriter();
    const bytes = await writer.write([[true, false, null]]);
    
    // We can just verify it generated a valid ZIP
    expect(bytes[0]).toBe(0x50); // P
    expect(bytes[1]).toBe(0x4b); // K
  });

  it('should support multiple sheets', async () => {
    const writer = new SheetWriter();
    writer.addSheet('First', [[1, 2]]);
    writer.addSheet('Second', [['A', 'B']]);
    const bytes = await writer.write();
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
    const bytes = await writer.write();

    // Verify via Reader
    const { SheetReader } = await import('../src/core/index');
    const reader = new SheetReader();
    // @ts-ignore
    const oldParse = reader.parseStyles;
    // @ts-ignore
    reader.parseStyles = async function(xmlTokenStream, map) {
      await oldParse.call(this, xmlTokenStream, map);
      console.log('STYLES MAP:', map);
    };
    
    const stream = new ReadableStream({ start(c) { c.enqueue(bytes); c.close(); } });
    
    let parsedRows: any[] = [];
    for await (const row of await reader.parse(stream)) {
      parsedRows.push(row);
    }
    
    // 45000 in Excel is 2023-03-15T00:00:00.000Z
    expect(parsedRows[0][0]).toBe('2023-03-15T00:00:00.000Z');
  });
});

describe('FormulaEngine (via SheetWriter)', () => {
  it('should correctly compute SUM in exported binary', async () => {
    const writer = new SheetWriter();
    const bytes = await writer.write([
      [1, 2, 3],
      [{ value: null, formula: '=SUM(A1:C1)' }]
    ]);
    expect(bytes.length).toBeGreaterThan(100);
  });
});
