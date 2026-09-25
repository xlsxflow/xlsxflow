export { SheetReader, type ParseOptions, sheetToJson, streamToCsv, createBlobReader, createFileReader, type RandomAccessReader } from './core/index';
export { SheetEditor } from './core/editor';
export { SheetWriter, type WriterOptions } from './core/writer';
export { createXmlStreamParser } from './core/xml-stream';
export { parseWorksheet, type RowData, type SheetMetadata } from './core/worksheet-parser';
export type { CellValue, StyledCell, Row, CellStyle, CellFont, CellFill, GradientFill, CellBorder, CellAlignment, BorderSide, RichTextRun, SheetOptions, ConditionalFormat, DataValidation } from './core/types';
export { StyleEngine } from './core/style-engine';
export { FormulaEngine } from './core/formula-engine';
export { ConditionalFormatter } from './core/conditional-formatter';

export const SheetForge = {
  async readFile(filePath: string, options?: import('./core/index').ParseOptions) {
    const { createFileReader } = await import('./core/random-access');
    const { SheetReader } = await import('./core/index');
    const reader = new SheetReader();
    const fileReader = await createFileReader(filePath);
    // All random reads happen inside parse(); the row stream opens its own handle.
    try {
      return await reader.parse(fileReader, options);
    } finally {
      await fileReader.close();
    }
  }
};
