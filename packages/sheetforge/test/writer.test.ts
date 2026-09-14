import { describe, it, expect } from 'vitest';
import { SheetWriter } from '../src/core/writer';

describe('SheetWriter', () => {
  it('should produce a valid ZIP binary starting with PK signature', () => {
    const writer = new SheetWriter();
    const bytes = writer.write([['Name', 'Score'], ['Alice', 95], ['Bob', 87]]);
    
    // ZIP Local File Header magic bytes: PK\x03\x04
    expect(bytes[0]).toBe(0x50); // P
    expect(bytes[1]).toBe(0x4b); // K
    expect(bytes[2]).toBe(0x03);
    expect(bytes[3]).toBe(0x04);
    expect(bytes.length).toBeGreaterThan(100);
  });

  it('should write a non-empty binary that contains sheet XML markers', () => {
    const writer = new SheetWriter();
    const bytes = writer.write([['Hello', 'World'], [1, 2]]);
    const text = new TextDecoder().decode(bytes);
    
    expect(text).toContain('[Content_Types].xml');
    expect(text).toContain('xl/workbook.xml');
    expect(text).toContain('xl/worksheets/sheet1.xml');
  });

  it('should inject formula string and cached value', () => {
    const writer = new SheetWriter();
    const bytes = writer.write([
      [10, 20, 30],
      [{ value: null, formula: '=SUM(A1:C1)' }]
    ]);
    const text = new TextDecoder().decode(bytes);

    expect(text).toContain('<f>SUM(A1:C1)</f>');
    expect(text).toContain('<v>60</v>');
  });

  it('should handle booleans and null cells', () => {
    const writer = new SheetWriter();
    const bytes = writer.write([[true, false, null]]);
    const text = new TextDecoder().decode(bytes);
    
    expect(text).toContain('t="b"');
  });
});

describe('FormulaEngine (via SheetWriter)', () => {
  it('should correctly compute SUM in exported binary', () => {
    const writer = new SheetWriter();
    const bytes = writer.write([
      [1, 2, 3],
      [{ value: null, formula: '=SUM(A1:C1)' }]
    ]);
    const text = new TextDecoder().decode(bytes);
    expect(text).toContain('<v>6</v>');
  });
});
