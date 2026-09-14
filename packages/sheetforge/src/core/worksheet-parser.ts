import { XmlToken } from './xml-stream';

export type CellValue = string | number | boolean | null;

export function excelToIsoDate(serial: number): string {
  let days = Math.floor(serial);
  const time = serial - days;
  // Excel bug: assumes 1900 is a leap year. Serial 60 is Feb 29 1900.
  if (days <= 60) days++; 
  // Milliseconds since 1899-12-30
  const epoch = Date.UTC(1899, 11, 30);
  const ms = epoch + days * 86400000 + time * 86400000;
  return new Date(ms).toISOString();
}

export async function* parseWorksheet(
  xmlTokenStream: ReadableStream<XmlToken>,
  sharedStrings: Map<number, string>,
  styles: Map<number, number>
): AsyncGenerator<CellValue[]> {
  const reader = xmlTokenStream.getReader();
  
  let currentRow: CellValue[] = [];
  let currentCellType: string | null = null;
  let currentStyleId: number | null = null;
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
          currentCellType = token.attributes['t'] || null; // 's' means shared string, 'inlineStr' means inline
          currentStyleId = token.attributes['s'] ? parseInt(token.attributes['s'], 10) : null;
          currentCellValue = '';
        } else if (token.name === 'v' || token.name === 't') {
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
             if (isNaN(num)) {
               resolvedValue = currentCellValue;
             } else {
               if (currentStyleId !== null && styles.get(currentStyleId) === 14) {
                 resolvedValue = excelToIsoDate(num);
               } else {
                 resolvedValue = num;
               }
             }
          } else if (currentCellType === 'inlineStr' || currentCellType === 'str') {
             resolvedValue = currentCellValue;
          }
          
          currentRow.push(resolvedValue);
          inCell = false;
          currentCellType = null;
          currentStyleId = null;
        } else if (token.name === 'v' || token.name === 't') {
          inValue = false;
        }
      }
    }
  } finally {
    reader.releaseLock();
  }
}
