export { SheetReader, type ParseOptions, type WorkbookInfo, sheetToJson, streamToCsv, createBlobReader, createFileReader, type RandomAccessReader } from './core/index';
export { SheetEditor, type CellEdit } from './core/editor';
export { parseCsv, type CsvOptions } from './core/csv';
export { SheetWriter, legacyPasswordHash, type WriterOptions } from './core/writer';
export { createXmlStreamParser } from './core/xml-stream';
export { parseWorksheet, type RowData, type SheetMetadata, type SheetComment } from './core/worksheet-parser';
export type { CellValue, StyledCell, Row, CellStyle, CellFont, CellFill, GradientFill, CellBorder, CellAlignment, BorderSide, RichTextRun, SheetOptions, SheetImage, CellComment, HighlightStyle, ComparisonOperator, RowOptions, ColumnOptions, SheetProtection, PageSetup, TableOptions, SheetView, WorkbookProperties, DefinedName, ConditionalFormatRule, ConditionalFormat, DataValidation } from './core/types';
export { StyleEngine } from './core/style-engine';
export { FormulaEngine } from './core/formula-engine';
export { ConditionalFormatter } from './core/conditional-formatter';

// Low-level building blocks (used by @xlsxflow/pro)
export { readSharedStrings } from './core/index';
export { ZipRandomAccessParser } from './core/zip-random-access';
export { ZipStreamWriter } from './core/zip-stream-writer';
export {
  resolveWorkbookParts, partRelationships, recalcOnOpen, mapFormulaRefs, shiftFormula, dateToSerial,
  colIndex, colLetter, encodeXString, decodeXString,
  type WorkbookParts, type Relationship, type RefMapper, type RefRole,
} from './core/utils';
export { anchorXml, drawingPartXml, DRAWING_NS, type Placement } from './core/image';

export const XlsxFlow = {
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
