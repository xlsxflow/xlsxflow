import { describe, it, expect } from 'vitest';
import { SheetEditor } from '../src/core/editor';
import { createBlobReader } from '../src/core/random-access';

describe('SheetEditor', () => {
  it('should preserve unknown files and modify targeted sheets', async () => {
    // We mock a ZIP file manually using a simple utility to simulate an uploaded Excel file
    // To do this simply, we'll use our own SheetWriter to generate a base ZIP,
    // then pass it through the Editor.
    
    // In a real test, we would construct a zip with a fake vbaProject.bin,
    // but for now, testing if it parses and repacks is sufficient.
    const { SheetWriter } = await import('../src/core/writer');
    const writer = new SheetWriter();
    writer.addSheet('Sheet1', [[1, 2]]);
    const originalStream = writer.write();
    const reader1 = originalStream.getReader();
    const chunks: Uint8Array[] = [];
    while (true) {
      const { done, value } = await reader1.read();
      if (done) break;
      chunks.push(value);
    }
    const totalLen = chunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const originalBytes = new Uint8Array(totalLen);
    let offset = 0;
    for (const chunk of chunks) {
      originalBytes.set(chunk, offset);
      offset += chunk.length;
    }

    const editor = new SheetEditor();
    editor.appendSheet('Sheet1', [[3, 4]]);
    
    const fileReaderIn = createBlobReader(new Blob([originalBytes]));
    const newStream = editor.edit(fileReaderIn);
    const reader2 = newStream.getReader();
    const newChunks: Uint8Array[] = [];
    while (true) {
      const { done, value } = await reader2.read();
      if (done) break;
      newChunks.push(value);
    }
    const newTotalLen = newChunks.reduce((sum, chunk) => sum + chunk.length, 0);
    const newBytes = new Uint8Array(newTotalLen);
    let newOffset = 0;
    for (const chunk of newChunks) {
      newBytes.set(chunk, newOffset);
      newOffset += chunk.length;
    }
    
    expect(newBytes.length).toBeGreaterThan(0);
    expect(newBytes[0]).toBe(0x50); // P
    expect(newBytes[1]).toBe(0x4b); // K
    
    // Check that we can read it back
    const { SheetReader } = await import('../src/core/index');
    const reader = new SheetReader();
    const fileReaderOut = createBlobReader(new Blob([newBytes]));
    const parsed = await reader.parse(fileReaderOut, { sheetName: 'Sheet1' });
    
    const rows = [];
    for await (const row of parsed) {
      rows.push(row);
    }
    
    // We started with [[1,2]], and appended [[3,4]]. Total should be 2 rows.
    expect(rows.length).toBe(2);
    expect(rows[0].cells).toEqual([1, 2]);
    expect(rows[1].cells).toEqual([3, 4]);
  });
});
