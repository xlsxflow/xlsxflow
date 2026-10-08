export { SheetReader, type ParseOptions, type WorkbookInfo, sheetToJson, streamToCsv, createBlobReader, createFileReader, type RandomAccessReader } from './core/index';
export { SheetEditor, type CellEdit } from './core/editor';
export { parseCsv, type CsvOptions } from './core/csv';
export { SheetWriter, legacyPasswordHash, type WriterOptions } from './core/writer';
export { OdsWriter, type OdsSheetOptions, type OdsWriterOptions } from './core/ods-writer';
export { createXmlStreamParser } from './core/xml-stream';
export { parseWorksheet, type RowData, type SheetMetadata, type SheetComment } from './core/worksheet-parser';
export type { CellValue, StyledCell, Row, CellStyle, CellFont, CellFill, GradientFill, CellBorder, CellAlignment, BorderSide, RichTextRun, SheetOptions, SheetImage, CellComment, HighlightStyle, ComparisonOperator, RowOptions, ColumnOptions, SheetProtection, PageSetup, TableOptions, SheetView, WorkbookProperties, DefinedName, ConditionalFormatRule, ConditionalFormat, DataValidation } from './core/types';
export { StyleEngine } from './core/style-engine';
export { FormulaEngine } from './core/formula-engine';
export { ConditionalFormatter } from './core/conditional-formatter';

// Low-level building blocks (used by @xlsxflow/pro)
export { readSharedStrings } from './core/index';
export { ZipRandomAccessParser } from './core/zip-random-access';
export { ZipStreamWriter, crc32 } from './core/zip-stream-writer';
export { CfbReader, isCfb, type CfbEntry } from './core/cfb';
export { StylePatcher } from './core/style-patcher';
export {
  resolveWorkbookParts, partRelationships, recalcOnOpen, mapFormulaRefs, shiftFormula, dateToSerial,
  colIndex, colLetter, encodeXString, decodeXString, checkCellText,
  type WorkbookParts, type Relationship, type RefMapper, type RefRole,
} from './core/utils';
export { anchorXml, drawingPartXml, DRAWING_NS, type Placement } from './core/image';

export const XlsxFlow = {
  async readFile(filePath: string, options?: import('./core/index').ParseOptions) {
    const { createFileReader } = await import('./core/random-access');
    const { SheetReader } = await import('./core/index');
    const reader = new SheetReader();
    const fileReader = await createFileReader(filePath);
    // Comments, images and .ods content are read after parse() returns, when the handle is closed:
    // those reads open a short-lived handle of their own
    let closed = false;
    const lateRead = async (offset: number, length: number) => {
      const late = await createFileReader(filePath);
      try { return await late.read(offset, length); } finally { await late.close(); }
    };
    try {
      return await reader.parse({
        size: fileReader.size,
        read: (offset, length) => closed ? lateRead(offset, length) : fileReader.read(offset, length),
        stream: (offset, length) => fileReader.stream(offset, length),
        close: async () => {},
      }, options);
    } finally {
      closed = true;
      await fileReader.close();
    }
  }
};
