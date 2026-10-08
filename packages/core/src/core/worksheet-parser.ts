import { XmlToken } from './xml-stream';
import { decodeXString, colIndex, shiftFormula, MAX_COLUMNS, unescapeXml, xmlElements } from './utils';
import { applyFontElement, ColorResolver } from './style-reader';
import { formatValue } from './number-format';
import type { CellStyle, RichTextRun, SheetImage } from './types';

export type CellValue = string | number | boolean | null;

export function excelToIsoDate(serial: number, is1904: boolean = false): string {
  let days = Math.floor(serial);
  // Excel stores times with ms precision; rounding avoids 18:29:59.999 for 18:30
  const timeMs = Math.round((serial - days) * 86400000);

  let epoch: number;
  if (is1904) {
    // 1904 date system starts on Jan 1, 1904
    epoch = Date.UTC(1904, 0, 1);
  } else {
    // 1900 date system has the leap year bug (phantom Feb 29 1900 = serial 60), so serials
    // 1..60 sit one day later than the 1899-12-30 epoch implies. Serial 0 (pure time) does not.
    if (days >= 1 && days <= 60) days++;
    epoch = Date.UTC(1899, 11, 30);
  }

  return new Date(epoch + days * 86400000 + timeMs).toISOString();
}

export interface SheetComment {
  ref: string;
  text: string;
  author?: string;
}

// Notes from a comments part. Threaded comments (Excel 365) also write their text here.
export function parseComments(xml: string): SheetComment[] {
  const authors = [...xmlElements(xml, 'author')].map(el => unescapeXml(el.body));
  const out: SheetComment[] = [];
  for (const el of xmlElements(xml, 'comment')) {
    const ref = /\sref="([^"]*)"/.exec(el.attrs)?.[1];
    if (!ref) continue;
    let text = '';
    // phonetic hints (rPh) hold their own <t>, which is not part of the note
    for (const t of xmlElements(el.body, 't|rPh')) if (t.name === 't') text += t.body;
    const comment: SheetComment = { ref, text: decodeXString(unescapeXml(text)) };
    const author = authors[parseInt(/\sauthorId="(\d+)"/.exec(el.attrs)?.[1] ?? '-1', 10)];
    if (author) comment.author = author;
    out.push(comment);
  }
  return out;
}


export interface SheetMetadata {
  mergedCells: string[];
  hiddenRows: number[];
  hiddenCols: number[];
  freezePanes?: { row: number; col: number; topLeftCell: string; activePane: string };
  // hyperlink is a URL, or "#Sheet!A1" for a place in the workbook (the writer's format)
  hyperlinks: { ref: string; hyperlink: string; tooltip?: string }[];
}

// Collects the formatted runs (<r>) of one rich string (<si> or <is>)
export class RichTextCollector {
  runs: RichTextRun[] = [];
  private run: RichTextRun | undefined;
  private inRPr = false;
  private inText = false;

  constructor(private color: ColorResolver) {}

  reset() {
    this.runs = [];
    this.run = undefined;
    this.inRPr = this.inText = false;
  }

  token(t: XmlToken) {
    if (t.type === 'startElement') {
      if (t.name === 'r') this.run = { text: '' };
      else if (!this.run) return;
      else if (t.name === 'rPr') {
        this.inRPr = true;
        this.run.font = {};
      } else if (this.inRPr) {
        const font = this.run.font!;
        applyFontElement(font, t.name, t.attributes, a => { const c = this.color(a); if (c) font.color = c; });
      } else if (t.name === 't') this.inText = true;
    } else if (t.type === 'endElement') {
      if (t.name === 'rPr') this.inRPr = false;
      else if (t.name === 't') this.inText = false;
      else if (t.name === 'r' && this.run) {
        this.run.text = decodeXString(this.run.text);
        if (this.run.font && !Object.keys(this.run.font).length) delete this.run.font;
        this.runs.push(this.run);
        this.run = undefined;
      }
    } else if (this.inText && this.run) {
      this.run.text += t.value;
    }
  }
}

export interface RowData {
  rowNumber: number;
  cells: CellValue[];
  // Only with the `formulas` / `styles` parse options, and only on rows that have any.
  // Indexed like `cells`; formulas have no leading "=".
  formulas?: (string | undefined)[];
  styles?: (CellStyle | undefined)[]; // shared between cells with the same style: do not mutate
  richText?: (RichTextRun[] | undefined)[]; // only cells whose text has formatted runs
  formatted?: (string | undefined)[]; // with the `formatted` option: each cell's text as Excel shows it
}

export interface WorksheetOptions {
  formulas?: boolean;
  cellStyles?: (CellStyle | undefined)[]; // by style index; set to report styles
  numFmts?: (string | undefined)[];         // number format code by style index; set to report formatted text
  richText?: ColorResolver;                // set to report rich text runs
  sharedRichText?: Map<number, RichTextRun[]>;
  hyperlinkTargets?: Map<string, string>;  // relationship id -> URL
  images?: () => Promise<SheetImage[]>;
  comments?: () => Promise<SheetComment[]>;
  maxPaddingCells?: number;                // empty cells added before far-right cells; the reader ties it to maxUncompressedBytes
}

export class ParseResult implements AsyncIterable<RowData> {
  constructor(
    private generator: AsyncGenerator<RowData>,
    private metadataPromise: Promise<SheetMetadata>,
    private images: () => Promise<SheetImage[]> = async () => [],
    private comments: () => Promise<SheetComment[]> = async () => []
  ) {}

  [Symbol.asyncIterator]() {
    return this.generator;
  }

  async getMetadata(): Promise<SheetMetadata> {
    return this.metadataPromise;
  }

  // The sheet's pictures, in the writer's `images` format. Image files are read on this call.
  getImages(): Promise<SheetImage[]> {
    return this.images();
  }

  // The sheet's notes (comments), in sheet order
  getComments(): Promise<SheetComment[]> {
    return this.comments();
  }
}

export function parseWorksheet(
  xmlTokenStream: ReadableStream<XmlToken | XmlToken[]>,
  sharedStrings: Map<number, string>,
  styles: Map<number, number>,
  is1904: boolean = false,
  options: WorksheetOptions = {}
): ParseResult {
  let resolveMeta!: (m: SheetMetadata) => void;
  let rejectMeta!: (e: any) => void;
  const metadataPromise = new Promise<SheetMetadata>((res, rej) => {
    resolveMeta = res;
    rejectMeta = rej;
  });
  // The row iterator already surfaces parse errors; callers who never ask for metadata
  // must not get an unhandled rejection (fatal by default in Node).
  metadataPromise.catch(() => {});
  const mergedCells: string[] = [];
  const hiddenRows: number[] = [];
  const hiddenColMarks = new Uint8Array(MAX_COLUMNS + 1); // by column; repeated <col> ranges cost nothing extra
  let padding = 0;
  const hyperlinks: SheetMetadata['hyperlinks'] = [];
  let freezePanes: SheetMetadata['freezePanes'] = undefined;

  async function* generateRows(): AsyncGenerator<RowData> {
    const reader = xmlTokenStream.getReader();
    
    let currentRow: CellValue[] = [];
    let currentCellType: string | null = null;
    let currentStyleId: number | null = null;
    let currentCellValue: string | null = null;
    let inRow = false;
    let inCell = false;
    let inValue = false;
    let inFormula = false;
    let inPhonetic = false;
    let currentFormula = '';
    let currentFormulaSi: string | null = null;
    let currentFormulaRef: string | null = null;
    let hasFormula = false;
    let rowFormulas: (string | undefined)[] | undefined;
    let rowStyles: (CellStyle | undefined)[] | undefined;
    let rowRichText: (RichTextRun[] | undefined)[] | undefined;
    let rowFormatted: (string | undefined)[] | undefined;
    let rawNumber: number | undefined; // a date cell's serial, for formatted text
    const runs = options.richText ? new RichTextCollector(options.richText) : undefined;
    
    let skipDepth = 0;

    // Shared formula anchors by si: the text and the cell it was written for
    const sharedFormulas = new Map<string, { formula: string; row: number; col: number }>();

    let currentRowNumber = 0;
    let currentColIndex = 0;
    let currentCellCol = 0;

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        for (const token of (Array.isArray(value) ? value : [value]) as XmlToken[]) {
          if (runs && inCell) runs.token(token);
          if (token.type === 'startElement') {
            if (token.name === 'AlternateContent') {
              skipDepth++;
            }
            if (skipDepth > 0) {
              if (token.name === 'Fallback') {
                // we could choose to parse fallbacks, but often they are duplicate
              }
              continue;
            }

            if (token.name === 'row') {
              inRow = true;
              currentRow = [];
              rowFormulas = rowStyles = rowRichText = rowFormatted = undefined;
              currentColIndex = 0;
              const r = token.attributes['r'];
              if (r) {
                currentRowNumber = parseInt(r, 10);
              } else {
                currentRowNumber++;
              }
              if (token.attributes['hidden'] === '1' || token.attributes['hidden'] === 'true') {
                hiddenRows.push(currentRowNumber);
              }
            } else if (token.name === 'col') {
              if (token.attributes['hidden'] === '1' || token.attributes['hidden'] === 'true') {
                // Clamp to the sheet's real width: a hostile max="999999999" must not allocate a billion entries
                const min = Math.max(1, parseInt(token.attributes['min'] || '1', 10) || 1);
                const max = Math.min(MAX_COLUMNS, parseInt(token.attributes['max'] || '0', 10) || 0);
                if (min <= max) hiddenColMarks.fill(1, min, max + 1);
              }
            } else if (token.name === 'c') {
              inCell = true;
              currentCellType = token.attributes['t'] || null;
              currentStyleId = token.attributes['s'] ? parseInt(token.attributes['s'], 10) : null;
              currentCellValue = null;
              hasFormula = false;
              runs?.reset();

              const r = token.attributes['r'];
              const refCol = r ? colIndex(r) : -1;
              currentCellCol = refCol >= 0 ? refCol : currentColIndex;
              if (currentCellCol >= MAX_COLUMNS) {
                throw new Error(`Invalid cell reference "${r}": beyond the last column (XFD)`);
              }
              // A tiny cell at XFD pads its row to 16,384 cells, so padding counts against the byte limit
              padding += Math.max(0, currentCellCol - currentRow.length);
              if (padding > (options.maxPaddingCells ?? Infinity)) {
                throw new Error(`Security Error: rows padded with more than ${options.maxPaddingCells} empty cells (maxUncompressedBytes).`);
              }
              while (currentRow.length < currentCellCol) currentRow.push(null);
            } else if (token.name === 'rPh') {
              inPhonetic = true; // phonetic hints are not part of the cell text
            } else if ((token.name === 'v' || token.name === 't') && !inPhonetic) {
              inValue = true;
              if (currentCellValue === null) currentCellValue = '';
            } else if (token.name === 'f') {
              inFormula = true;
              hasFormula = true;
              currentFormula = '';
              currentFormulaSi = token.attributes['si'] || null;
              currentFormulaRef = token.attributes['ref'] || null;
            } else if (token.name === 'hyperlink') {
              const { ref, id, location, tooltip } = token.attributes;
              const target = id ? options.hyperlinkTargets?.get(id) : undefined;
              const hyperlink = target ? (location ? `${target}#${location}` : target) : location ? `#${location}` : '';
              if (ref && hyperlink) hyperlinks.push(tooltip ? { ref, hyperlink, tooltip } : { ref, hyperlink });
            } else if (token.name === 'mergeCell') {
              if (token.attributes['ref']) mergedCells.push(token.attributes['ref']);
            } else if (token.name === 'pane') {
              // frozenSplit: frozen after being split, which Excel also writes
              if (token.attributes['state'] === 'frozen' || token.attributes['state'] === 'frozenSplit') {
                freezePanes = {
                  row: parseInt(token.attributes['ySplit'] || '0', 10),
                  col: parseInt(token.attributes['xSplit'] || '0', 10),
                  topLeftCell: token.attributes['topLeftCell'] || '',
                  activePane: token.attributes['activePane'] || 'bottomRight'
                };
              }
            }
          } else if (token.type === 'text') {
            if (skipDepth > 0) continue;
            if (inValue) {
              currentCellValue = (currentCellValue || '') + token.value;
            } else if (inFormula) {
              currentFormula += token.value;
            }
          } else if (token.type === 'endElement') {
            if (token.name === 'AlternateContent') {
              skipDepth--;
              continue;
            }
            if (skipDepth > 0) continue;

            if (token.name === 'row') {
              const row: RowData = { rowNumber: currentRowNumber, cells: currentRow };
              if (rowFormulas) row.formulas = rowFormulas;
              if (rowStyles) row.styles = rowStyles;
              if (rowRichText) row.richText = rowRichText;
              if (rowFormatted) row.formatted = rowFormatted;
              yield row;
              inRow = false;
            } else if (token.name === 'c') {
              let resolvedValue: CellValue = null;
            
              // If the cell had a shared formula reference but no <v>, we still don't have the evaluated value
              // since we are streaming. A full formula engine would calculate it. 
              // For now, if currentCellValue is null but we have a formula, we could return the formula string 
              // or null. We'll stick to yielding null for cached value.
            
              if (currentCellValue !== null) {
                if (currentCellType === 'inlineStr' || currentCellType === 'str') {
                  // Strings may legitimately be empty
                  resolvedValue = decodeXString(currentCellValue);
                } else if (currentCellValue === '') {
                  // <v></v> on a non-string cell carries no value
                  resolvedValue = null;
                } else if (currentCellType === 's') {
                  const index = parseInt(currentCellValue, 10);
                  // An index past the table points nowhere: the cell is empty
                  resolvedValue = sharedStrings.get(index) ?? null;
                } else if (currentCellType === 'b') {
                  resolvedValue = currentCellValue === '1' || currentCellValue === 'true';
                } else if (currentCellType === 'e') {
                  resolvedValue = currentCellValue;
                } else {
                  const num = Number(currentCellValue);
                  if (isNaN(num)) {
                    resolvedValue = currentCellValue; // e.g. t="d" ISO dates
                  } else if (currentStyleId !== null && styles.get(currentStyleId) === 14) {
                    rawNumber = num;
                    resolvedValue = excelToIsoDate(num, is1904);
                  } else {
                    resolvedValue = num === 0 ? 0 : num; // normalise -0
                  }
                }
              }
            
              if (options.formulas && hasFormula) {
                let formula: string | undefined = currentFormula ? decodeXString(currentFormula) : undefined;
                if (!formula && currentFormulaSi !== null) {
                  // A follower cell of a shared formula: the anchor's text, moved to this cell
                  const anchor = sharedFormulas.get(currentFormulaSi);
                  if (anchor) formula = shiftFormula(anchor.formula, currentRowNumber - anchor.row, currentCellCol - anchor.col);
                }
                if (formula) (rowFormulas ??= [])[currentCellCol] = formula;
              }
              if (runs) {
                const cellRuns = currentCellType === 's' && currentCellValue
                  ? options.sharedRichText?.get(parseInt(currentCellValue, 10))
                  : runs.runs.length ? runs.runs : undefined;
                if (cellRuns) (rowRichText ??= [])[currentCellCol] = cellRuns;
              }
              if (options.numFmts && resolvedValue !== null) {
                const code = currentStyleId !== null ? options.numFmts[currentStyleId] : undefined;
                (rowFormatted ??= [])[currentCellCol] = currentCellType === 'e' ? String(resolvedValue)
                  : formatValue(rawNumber ?? resolvedValue, code ?? 'General', is1904);
              }
              rawNumber = undefined;
              if (options.cellStyles && currentStyleId !== null) {
                const style = options.cellStyles[currentStyleId];
                if (style) (rowStyles ??= [])[currentCellCol] = style;
              }

              // Indexed write: also correct when a cell's r points left of the previous cell
              currentRow[currentCellCol] = resolvedValue;
              currentColIndex = currentCellCol + 1;
              inCell = false;
              currentCellType = null;
              currentStyleId = null;
              currentCellValue = null;
            } else if (token.name === 'rPh') {
              inPhonetic = false;
            } else if (token.name === 'v' || token.name === 't') {
              inValue = false;
            } else if (token.name === 'f') {
              inFormula = false;
              if (currentFormulaSi !== null && currentFormulaRef !== null && currentFormula) {
                sharedFormulas.set(currentFormulaSi, { formula: decodeXString(currentFormula), row: currentRowNumber, col: currentCellCol });
              }
            }
          }
        }
      }
      const hiddenCols: number[] = [];
      hiddenColMarks.forEach((hidden, c) => { if (hidden) hiddenCols.push(c); });
      resolveMeta({ mergedCells, hiddenRows, hiddenCols, freezePanes, hyperlinks });
    } catch (e) {
      rejectMeta(e);
      throw e;
    } finally {
      reader.releaseLock();
    }
  }

  return new ParseResult(generateRows(), metadataPromise, options.images, options.comments);
}
