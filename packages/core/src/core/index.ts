import { RandomAccessReader } from './random-access';
import { ZipRandomAccessParser } from './zip-random-access';
import { createXmlBatchParser, XmlToken } from './xml-stream';
import { parseWorksheet, parseComments, ParseResult, RichTextCollector } from './worksheet-parser';
import { resolveWorkbookParts, partRelationships, hyperlinkTargets, decodeXString, isDateFormatCode, attr, unescapeXml } from './utils';
import { parseThemeColors, colorResolver, applyFontElement, ColorResolver } from './style-reader';
import { readSheetImages } from './image';
import type { CellStyle, CellFont, CellFill, GradientFill, CellBorder, BorderSide, CellAlignment, RichTextRun, DefinedName, WorkbookProperties } from './types';
export { SheetWriter } from './writer';
export { createFileReader, createBlobReader, type RandomAccessReader } from './random-access';
export { sheetToJson, streamToCsv, resolveSheetPaths } from './utils';
export { SheetEditor } from './editor';

export interface ParseOptions {
  sheetName?: string;
  // Zip bomb protection: max uncompressed bytes per part. Parts held in memory (workbook, rels,
  // shared strings, styles) default to 1 GiB; the streamed worksheet is only limited when set.
  // Pass Infinity to disable.
  maxUncompressedBytes?: number;
  formulas?: boolean; // report cell formulas in RowData.formulas
  styles?: boolean;   // report cell styles in RowData.styles
  richText?: boolean; // report formatted text runs in RowData.richText
  formatted?: boolean; // report each cell's text as Excel shows it (en-US) in RowData.formatted
}

// Number formats Excel does not write into styles.xml (ECMA-376 Part 1, 18.8.30)
const BUILTIN_NUM_FMTS: [number, string][] = [
  [1, '0'], [2, '0.00'], [3, '#,##0'], [4, '#,##0.00'], [9, '0%'], [10, '0.00%'], [11, '0.00E+00'],
  [12, '# ?/?'], [13, '# ??/??'], [14, 'mm-dd-yy'], [15, 'd-mmm-yy'], [16, 'd-mmm'], [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'], [19, 'h:mm:ss AM/PM'], [20, 'h:mm'], [21, 'h:mm:ss'], [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'], [38, '#,##0 ;[Red](#,##0)'], [39, '#,##0.00;(#,##0.00)'], [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'], [46, '[h]:mm:ss'], [47, 'mmss.0'], [48, '##0.0E+0'], [49, '@'],
];

// Formats Excel stores by id only and shows by locale, as en-US Excel shows them
const DISPLAY_NUM_FMTS = new Map<number, string>([
  [5, '"$"#,##0_);("$"#,##0)'], [6, '"$"#,##0_);[Red]("$"#,##0)'], [7, '"$"#,##0.00_);("$"#,##0.00)'], [8, '"$"#,##0.00_);[Red]("$"#,##0.00)'],
  [14, 'm/d/yyyy'], [22, 'm/d/yyyy h:mm'],
  [41, '_(* #,##0_);_(* \\(#,##0\\);_(* "-"_);_(@_)'], [42, '_("$"* #,##0_);_("$"* \\(#,##0\\);_("$"* "-"_);_(@_)'],
  [43, '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'], [44, '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'],
]);

const DEFAULT_MAX_PART_BYTES = 1 << 30;

export interface WorkbookInfo {
  sheets: { name: string; state: 'visible' | 'hidden' | 'veryHidden' }[]; // in tab order
  definedNames: DefinedName[]; // includes Excel's own, such as _xlnm.Print_Area
  properties: WorkbookProperties;
}

// Text of the first <tag> element in xml, unescaped
const elementText = (xml: string, tag: string) => {
  const m = new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`).exec(xml);
  return m ? unescapeXml(m[1]) : undefined;
};

function createByteLimitStream(maxBytes: number): TransformStream<Uint8Array, Uint8Array> {
  let bytesRead = 0;
  return new TransformStream({
    transform(chunk, controller) {
      bytesRead += chunk.length;
      if (bytesRead > maxBytes) {
        controller.error(new Error(`Security Error: Stream exceeded maximum uncompressed size of ${maxBytes} bytes.`));
      } else {
        controller.enqueue(chunk);
      }
    }
  });
}

// Shared string table (index -> text) of a workbook's sharedStrings part
export async function readSharedStrings(stream: ReadableStream<Uint8Array>): Promise<Map<number, string>> {
  const map = new Map<number, string>();
  await SheetReader.parseSharedStrings(stream.pipeThrough(createXmlBatchParser()), map);
  return map;
}

export class SheetReader {
  // Sheet names and visibility, defined names and document properties, without reading any sheet
  async readWorkbook(reader: RandomAccessReader, options?: Pick<ParseOptions, 'maxUncompressedBytes'>): Promise<WorkbookInfo> {
    const zip = new ZipRandomAccessParser(reader);
    await zip.parseCentralDirectory();
    const max = options?.maxUncompressedBytes ?? DEFAULT_MAX_PART_BYTES;
    const readText = async (name: string) => zip.has(name)
      ? this.readStreamToString(max !== Infinity ? (await zip.extractStream(name)).pipeThrough(createByteLimitStream(max)) : await zip.extractStream(name))
      : '';
    const parts = await resolveWorkbookParts(readText);

    const sheets: WorkbookInfo['sheets'] = [];
    for (const [tag] of parts.workbookXml.matchAll(/<(?:\w+:)?sheet\b[^>]*>/g)) {
      const name = attr(tag, 'name');
      const state = attr(tag, 'state');
      if (name !== null) sheets.push({ name: unescapeXml(name), state: state === 'hidden' || state === 'veryHidden' ? state : 'visible' });
    }
    const definedNames: DefinedName[] = [];
    for (const [, , open, text] of parts.workbookXml.matchAll(/<((?:\w+:)?)definedName\b([^>]*)>([\s\S]*?)<\/\1definedName>/g)) {
      const local = attr(open, 'localSheetId');
      const comment = attr(open, 'comment');
      const name: DefinedName = { name: unescapeXml(attr(open, 'name') ?? ''), ref: unescapeXml(text) };
      if (local !== null && sheets[+local]) name.sheet = sheets[+local].name;
      if (comment !== null) name.comment = unescapeXml(comment);
      if (/^(?:1|true)$/.test(attr(open, 'hidden') ?? '')) name.hidden = true;
      definedNames.push(name);
    }

    const rootRels = await partRelationships(readText, '');
    const core = await readText(rootRels.find(r => r.type.endsWith('/core-properties'))?.path ?? 'docProps/core.xml');
    const app = await readText(rootRels.find(r => r.type.endsWith('/extended-properties'))?.path ?? 'docProps/app.xml');
    const properties: WorkbookProperties = {};
    for (const [key, tag] of [['title', 'dc:title'], ['subject', 'dc:subject'], ['creator', 'dc:creator'], ['keywords', 'cp:keywords'],
      ['description', 'dc:description'], ['category', 'cp:category']] as const) {
      const v = elementText(core, tag);
      if (v !== undefined) properties[key] = v;
    }
    const created = elementText(core, 'dcterms:created');
    if (created && !isNaN(Date.parse(created))) properties.created = new Date(created);
    for (const [key, tag] of [['company', 'Company'], ['manager', 'Manager']] as const) {
      const v = elementText(app, tag);
      if (v !== undefined) properties[key] = v;
    }
    return { sheets, definedNames, properties };
  }

  async parse(reader: RandomAccessReader, options?: ParseOptions): Promise<ParseResult> {
    const zip = new ZipRandomAccessParser(reader);
    await zip.parseCentralDirectory();

    const sharedStrings = new Map<number, string>();
    const styles = new Map<number, number>();
    const limit = (st: ReadableStream<Uint8Array>, max = options?.maxUncompressedBytes ?? DEFAULT_MAX_PART_BYTES) =>
      max !== Infinity ? st.pipeThrough(createByteLimitStream(max)) : st;
    const readText = async (name: string) =>
      zip.has(name) ? this.readStreamToString(limit(await zip.extractStream(name))) : '';

    // 1. Locate workbook, sheets, shared strings and styles via package relationships
    const parts = await resolveWorkbookParts(readText);
    let worksheetZipPath: string | undefined;
    if (options?.sheetName) {
      worksheetZipPath = parts.sheets.get(options.sheetName);
      if (!worksheetZipPath) throw new Error(`Sheet with name "${options.sheetName}" not found in workbook.`);
    } else {
      worksheetZipPath = parts.sheets.values().next().value
        ?? zip.getFiles().find(f => /^xl\/worksheets\/[^/]+\.xml$/.test(f));
      if (!worksheetZipPath) throw new Error("No worksheets found in ZIP.");
    }
    const is1904 = /<(?:\w+:)?workbookPr\b[^>]*\bdate1904="(?:1|true)"/.test(parts.workbookXml);

    // 2. Parse Styles (if exists). Colours can refer to the theme, only read when they are reported.
    const theme = (options?.styles || options?.richText) && parts.theme ? parseThemeColors(await readText(parts.theme)) : [];
    const stylesPath = parts.styles ?? 'xl/styles.xml';
    let cellStyles: (CellStyle | undefined)[] = [];
    let color = colorResolver(theme);
    let formats: (string | undefined)[] = [];
    if (zip.has(stylesPath)) {
      const xmlStream = limit(await zip.extractStream(stylesPath)).pipeThrough(createXmlBatchParser());
      ({ cellStyles, color, formats } = await this.parseStyles(xmlStream, styles, theme));
    }

    // 3. Parse Shared Strings (if exists)
    const sstPath = parts.sharedStrings ?? 'xl/sharedStrings.xml';
    const sharedRichText = options?.richText ? new Map<number, RichTextRun[]>() : undefined;
    if (zip.has(sstPath)) {
      const xmlStream = limit(await zip.extractStream(sstPath)).pipeThrough(createXmlBatchParser());
      await SheetReader.parseSharedStrings(xmlStream, sharedStrings, sharedRichText && { map: sharedRichText, color });
    }

    // 4. Parse Worksheet
    if (!zip.has(worksheetZipPath)) {
      throw new Error(`Worksheet ${worksheetZipPath} not found in ZIP.`);
    }
    // Rows are yielded as they stream, so an unlimited worksheet costs time, not memory
    const xmlStream = limit(await zip.extractStream(worksheetZipPath), options?.maxUncompressedBytes ?? Infinity)
      .pipeThrough(createXmlBatchParser());
    return parseWorksheet(xmlStream, sharedStrings, styles, is1904, {
      formulas: options?.formulas,
      cellStyles: options?.styles ? cellStyles : undefined,
      numFmts: options?.formatted ? formats : undefined,
      richText: options?.richText ? color : undefined,
      sharedRichText,
      hyperlinkTargets: await hyperlinkTargets(readText, worksheetZipPath),
      comments: async () => {
        const rel = (await partRelationships(readText, worksheetZipPath)).find(r => !r.external && r.type.endsWith('/comments'));
        return rel ? parseComments(await readText(rel.path)) : [];
      },
      images: () => readSheetImages(readText, async path => zip.has(path)
        ? new Uint8Array(await new Response(limit(await zip.extractStream(path))).arrayBuffer()) : undefined, worksheetZipPath),
    });
  }

  static async parseSharedStrings(
    xmlTokenStream: ReadableStream<XmlToken | XmlToken[]>,
    map: Map<number, string>,
    rich?: { map: Map<number, RichTextRun[]>; color: ColorResolver }
  ) {
    const reader = xmlTokenStream.getReader();
    const runs = rich && new RichTextCollector(rich.color);
    let index = 0;
    let inText = false;
    let inPhonetic = false;
    let currentString = '';

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        for (const token of (Array.isArray(value) ? value : [value]) as XmlToken[]) {
          runs?.token(token);
          if (token.type === 'startElement' && token.name === 'si') {
            currentString = '';
            runs?.reset();
          } else if (token.type === 'startElement' && token.name === 'rPh') {
            inPhonetic = true; // phonetic hints are not part of the cell text
          } else if (token.type === 'endElement' && token.name === 'rPh') {
            inPhonetic = false;
          } else if (token.type === 'startElement' && token.name === 't' && !inPhonetic) {
            inText = true;
          } else if (token.type === 'text' && inText) {
            currentString += token.value;
          } else if (token.type === 'endElement' && token.name === 't') {
            inText = false;
          } else if (token.type === 'endElement' && token.name === 'si') {
            if (runs?.runs.length) rich!.map.set(index, runs.runs);
            map.set(index++, decodeXString(currentString));
          }
        }
      }
    } finally {
      reader.releaseLock();
    }
  }

  // Fills `dateStyles` (style index -> 14 when it formats dates), returns each style as a CellStyle and the
  // colour resolver for this workbook (theme + its indexed palette).
  private async parseStyles(xmlTokenStream: ReadableStream<XmlToken | XmlToken[]>, dateStyles: Map<number, number>, theme: string[]) {
    const reader = xmlTokenStream.getReader();
    const numFmts = new Map<number, string>(BUILTIN_NUM_FMTS);
    const customDateFmts = new Set<number>();
    const fonts: (CellFont | undefined)[] = [];
    const fills: (CellFill | GradientFill | undefined)[] = [];
    const borders: (CellBorder | undefined)[] = [];
    const cellStyles: (CellStyle | undefined)[] = [];
    const formats: (string | undefined)[] = []; // number format code by style index, for formatted text

    // Only direct children of these lists count: dxfs and cellStyleXfs reuse the same element names
    let section = '';
    let font: CellFont | undefined;
    let fill: CellFill | GradientFill | undefined;
    let pattern = '';
    let border: CellBorder | undefined;
    let side: BorderSide | undefined;
    let xf: CellStyle | undefined;
    // A custom palette (<colors>) comes after the styles that use it, so colours are resolved at the end
    const palette: string[] = [];
    const colors: [Record<string, string>, (argb: string) => void][] = [];
    const later = (a: Record<string, string>, set: (argb: string) => void) => colors.push([a, set]);

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        for (const token of (Array.isArray(value) ? value : [value]) as XmlToken[]) {
          if (token.type === 'startElement') {
            const a = token.attributes;
            if (font) { const f = font; if (applyFontElement(f, token.name, a, c => later(c, argb => { f.color = argb; }))) continue; }
            switch (token.name) {
              case 'numFmts': case 'fonts': case 'fills': case 'borders': case 'cellXfs':
                section = token.name;
                break;
              case 'numFmt':
                if (section === 'numFmts') {
                  const id = parseInt(a['numFmtId'] || '0', 10);
                  numFmts.set(id, a['formatCode'] || '');
                  if (isDateFormatCode(a['formatCode'] || '')) customDateFmts.add(id);
                }
                break;
              case 'font':
                if (section === 'fonts') font = {};
                break;
              case 'rgbColor':
                if (section === 'indexedColors') palette.push(a['rgb'] ?? '');
                break;
              case 'indexedColors':
                section = token.name;
                break;
              case 'color':
                if (side) { const target = side; later(a, c => { target.color = c; }); }
                else if (font) { const target = font; later(a, c => { target.color = c; }); }
                else if (fill?.type === 'gradient' && fill.stops.length) { const stop = fill.stops[fill.stops.length - 1]; later(a, c => { stop.color = c; }); }
                break;
              case 'fill':
                if (section === 'fills') fills.push(fill = undefined);
                break;
              case 'patternFill':
                pattern = a['patternType'] || '';
                break;
              case 'fgColor':
                if (section === 'fills' && pattern === 'solid') {
                  const solid: CellFill = fill = { type: 'solid', fgColor: '' };
                  later(a, c => { solid.fgColor = c; });
                }
                break;
              case 'gradientFill':
                if (section === 'fills') fill = { type: 'gradient', degree: Number(a['degree'] || 0), stops: [] };
                break;
              case 'stop':
                if (fill?.type === 'gradient') fill.stops.push({ position: Number(a['position'] || 0), color: '' });
                break;
              case 'border':
                if (section === 'borders') border = {};
                break;
              case 'left': case 'right': case 'top': case 'bottom':
                if (border && a['style'] && a['style'] !== 'none') {
                  side = border[token.name as 'left'] = { style: a['style'] as BorderSide['style'] };
                }
                break;
              case 'xf':
                if (section === 'cellXfs') {
                  const index = cellStyles.length;
                  const numFmtId = parseInt(a['numFmtId'] || '0', 10);
                  // Builtin date/time ids (27-36 and 50-58 are dates in CJK locales)
                  const isDate = (numFmtId >= 14 && numFmtId <= 22) || (numFmtId >= 27 && numFmtId <= 36)
                    || (numFmtId >= 45 && numFmtId <= 47) || (numFmtId >= 50 && numFmtId <= 58) || customDateFmts.has(numFmtId);
                  dateStyles.set(index, isDate ? 14 : 0);
                  formats.push(DISPLAY_NUM_FMTS.get(numFmtId) ?? numFmts.get(numFmtId));

                  xf = {};
                  const fontId = parseInt(a['fontId'] || '0', 10);
                  const fillId = parseInt(a['fillId'] || '0', 10);
                  const borderId = parseInt(a['borderId'] || '0', 10);
                  // Font 0 is the workbook default; fills 0 and 1 are the mandatory none/gray125
                  if (fontId > 0 && fonts[fontId]) xf.font = fonts[fontId];
                  if (fillId > 1 && fills[fillId]) xf.fill = fills[fillId];
                  if (borders[borderId]) xf.border = borders[borderId];
                  if (numFmtId > 0 && numFmts.has(numFmtId)) xf.numFmt = numFmts.get(numFmtId);
                  cellStyles.push(xf);
                }
                break;
              case 'alignment':
                if (xf) {
                  const alignment: CellAlignment = {};
                  if (a['horizontal']) alignment.horizontal = a['horizontal'] as CellAlignment['horizontal'];
                  if (a['vertical']) alignment.vertical = a['vertical'] as CellAlignment['vertical'];
                  if (a['wrapText'] === '1' || a['wrapText'] === 'true') alignment.wrapText = true;
                  if (Object.keys(alignment).length) xf.alignment = alignment;
                }
                break;
            }
          } else if (token.type === 'endElement') {
            switch (token.name) {
              case 'numFmts': case 'fonts': case 'fills': case 'borders': case 'cellXfs': case 'indexedColors':
                section = '';
                break;
              case 'font':
                if (font) fonts.push(font);
                font = undefined;
                break;
              case 'fill':
                if (section === 'fills') fills[fills.length - 1] = fill;
                fill = undefined;
                pattern = '';
                break;
              case 'left': case 'right': case 'top': case 'bottom':
                side = undefined;
                break;
              case 'border':
                if (border) borders.push(Object.keys(border).length ? border : undefined);
                border = undefined;
                break;
              case 'xf':
                xf = undefined;
                break;
            }
          }
        }
      }
    } finally {
      reader.releaseLock();
    }
    const color = colorResolver(theme, palette.length ? palette : undefined);
    for (const [a, set] of colors) {
      const argb = color(a);
      if (argb) set(argb);
    }
    // Colours that did not resolve (auto, unknown theme slot) are dropped; so is a solid fill without one
    for (const f of fills) if (f?.type === 'solid' && !f.fgColor) fills[fills.indexOf(f)] = undefined;
    for (const f of fills) if (f?.type === 'gradient') f.stops = f.stops.filter(st => st.color);
    for (const st of cellStyles) if (st?.fill && !fills.includes(st.fill)) delete st.fill;
    // Default-looking styles report as no style
    return { cellStyles: cellStyles.map(st => (st && Object.keys(st).length ? st : undefined)), color, formats };
  }

  private async readStreamToString(stream: ReadableStream<Uint8Array>): Promise<string> {
    const textDecoder = new TextDecoderStream();
    const reader = stream.pipeThrough(textDecoder as any).getReader();
    let result = '';
    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      result += value;
    }
    return result;
  }
}

