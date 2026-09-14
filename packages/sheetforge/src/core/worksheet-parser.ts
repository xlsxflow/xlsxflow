import { XmlToken } from './xml-stream';

export type CellValue = string | number | boolean | null;

export async function* parseWorksheet(
  xmlTokenStream: ReadableStream<XmlToken>,
  sharedStrings: Map<number, string>
): AsyncGenerator<CellValue[]> {
  const reader = xmlTokenStream.getReader();
  
  let currentRow: CellValue[] = [];
  let currentCellType: string | null = null;
  let currentCellValue: string = '';
  let inRow = false;
  let inCell = false;
  let inValue = false;

  try {
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;

      const token = value as XmlToken;

      if (token.type === 'startElement') {
        if (token.name === 'row') {
          inRow = true;
          currentRow = [];
        } else if (token.name === 'c') {
          inCell = true;
          currentCellType = token.attributes['t'] || null; // 's' means shared string
          currentCellValue = '';
        } else if (token.name === 'v') {
          inValue = true;
        }
      } else if (token.type === 'text') {
        if (inValue) {
          currentCellValue += token.value;
        }
      } else if (token.type === 'endElement') {
        if (token.name === 'row') {
          yield currentRow;
          inRow = false;
        } else if (token.name === 'c') {
          // Resolve cell value
          let resolvedValue: CellValue = currentCellValue;
          
          if (currentCellType === 's' && currentCellValue) {
            const index = parseInt(currentCellValue, 10);
            resolvedValue = sharedStrings.get(index) ?? currentCellValue;
          } else if (currentCellType === 'b') {
            resolvedValue = currentCellValue === '1';
          } else if (currentCellType === 'n' || (!currentCellType && currentCellValue)) {
             const num = Number(currentCellValue);
             resolvedValue = isNaN(num) ? currentCellValue : num;
          }
          
          currentRow.push(resolvedValue);
          inCell = false;
          currentCellType = null;
        } else if (token.name === 'v') {
          inValue = false;
        }
      }
    }
  } finally {
    reader.releaseLock();
  }
}
