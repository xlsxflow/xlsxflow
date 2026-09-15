import { describe, it, expect } from 'vitest';
import { createXmlStreamParser } from '../src/core/xml-stream';
import { parseWorksheet } from '../src/core/worksheet-parser';

describe('XML SAX-like Stream Parser', () => {
  it('should parse simple XML into tokens', async () => {
    const xml = `<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`;
    
    const stream = new ReadableStream({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(xml));
        controller.close();
      }
    });

    const tokenStream = stream.pipeThrough(createXmlStreamParser());
    const reader = tokenStream.getReader();
    
    const tokens = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      tokens.push(value);
    }

    expect(tokens).toContainEqual({ type: 'startElement', name: 'row', attributes: { r: '1' } });
    expect(tokens).toContainEqual({ type: 'startElement', name: 'c', attributes: { r: 'A1', t: 's' } });
    expect(tokens).toContainEqual({ type: 'text', value: '0' });
    expect(tokens).toContainEqual({ type: 'endElement', name: 'c' });
  });

  it('should handle CDATA and HTML comments gracefully', async () => {
    const xml = `<worksheet><!-- ignore me --><c><![CDATA[<foo> & bar]]></c></worksheet>`;
    
    const stream = new ReadableStream({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(xml));
        controller.close();
      }
    });

    const tokenStream = stream.pipeThrough(createXmlStreamParser());
    const reader = tokenStream.getReader();
    
    const tokens = [];
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      tokens.push(value);
    }

    expect(tokens).toContainEqual({ type: 'startElement', name: 'worksheet', attributes: {} });
    expect(tokens).toContainEqual({ type: 'startElement', name: 'c', attributes: {} });
    expect(tokens).toContainEqual({ type: 'text', value: '<foo> & bar' });
    expect(tokens).toContainEqual({ type: 'endElement', name: 'c' });
    expect(tokens).toContainEqual({ type: 'endElement', name: 'worksheet' });
  });
});

describe('Worksheet Row Assembler', () => {
  it('should assemble rows correctly using shared strings', async () => {
    const xml = `
      <worksheet>
        <sheetData>
          <row>
            <c t="s"><v>0</v></c>
            <c><v>123.45</v></c>
            <c t="b"><v>1</v></c>
          </row>
        </sheetData>
      </worksheet>
    `;

    const stream = new ReadableStream({
      start(controller) {
        controller.enqueue(new TextEncoder().encode(xml));
        controller.close();
      }
    });

    const tokenStream = stream.pipeThrough(createXmlStreamParser());
    
    const sharedStrings = new Map<number, string>();
    sharedStrings.set(0, 'Hello World');
    const styles = new Map<number, number>();

    const rowGenerator = parseWorksheet(tokenStream, sharedStrings, styles);
    
    const rows = [];
    for await (const row of rowGenerator) {
      rows.push(row);
    }

    expect(rows.length).toBe(1);
    expect(rows[0]).toEqual(['Hello World', 123.45, true]);
  });
});
