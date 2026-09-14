import { createZipStreamParser, ZipEntry } from './zip-stream';
import { createXmlStreamParser, XmlToken } from './xml-stream';
import { parseWorksheet, CellValue } from './worksheet-parser';

export interface ParseOptions {
  sheetName?: string;
}

export class SheetReader {
  async parse(stream: ReadableStream<Uint8Array>, options?: ParseOptions) {
    const zipStream = stream.pipeThrough(createZipStreamParser());
    const zipReader = zipStream.getReader();

    let sharedStrings = new Map<number, string>();
    let styles = new Map<number, number>();
    let workbookXml = '';
    let workbookRelsXml = '';
    let targetSheetPath: string | null = null;
    let worksheetStream: ReadableStream<Uint8Array> | null = null;

    // We must pump the ZIP stream concurrently in the background to prevent backpressure deadlock.
    // This promise resolves as soon as we find the target worksheet stream.
    const findWorksheetPromise = new Promise<ReadableStream<Uint8Array>>((resolve, reject) => {
      (async () => {
        try {
          while (true) {
            const { done, value } = await zipReader.read();
            if (done) {
              if (!worksheetStream) reject(new Error("Target worksheet not found in stream."));
              break;
            }
            
            const entry = value as ZipEntry;
            
            if (entry.filename === 'xl/workbook.xml') {
              workbookXml = await this.readStreamToString(entry);
              this.tryResolveSheetPath(workbookXml, workbookRelsXml, options, (path) => { targetSheetPath = path; });
            } else if (entry.filename === 'xl/_rels/workbook.xml.rels') {
              workbookRelsXml = await this.readStreamToString(entry);
              this.tryResolveSheetPath(workbookXml, workbookRelsXml, options, (path) => { targetSheetPath = path; });
            } else if (entry.filename === 'xl/sharedStrings.xml') {
              let s = entry.stream;
              if (entry.compressionMethod === 8) s = s.pipeThrough(new DecompressionStream('deflate-raw') as any);
              const xmlStream = s.pipeThrough(createXmlStreamParser());
              await this.parseSharedStrings(xmlStream, sharedStrings);
            } else if (entry.filename === 'xl/styles.xml') {
              let s = entry.stream;
              if (entry.compressionMethod === 8) s = s.pipeThrough(new DecompressionStream('deflate-raw') as any);
              const xmlStream = s.pipeThrough(createXmlStreamParser());
              await this.parseStyles(xmlStream, styles);
            } else if (entry.filename.startsWith('xl/worksheets/')) {
              if (options?.sheetName && !targetSheetPath) {
                reject(new Error(`Cannot stream worksheet: workbook.xml must precede worksheets in the ZIP to resolve by name.`));
                await entry.stream.pipeTo(new WritableStream());
                continue;
              }

              const isTarget = targetSheetPath ? `xl/${targetSheetPath}` === entry.filename : true;
              
              if (isTarget && !worksheetStream) {
                worksheetStream = entry.stream;
                if (entry.compressionMethod === 8) {
                  worksheetStream = worksheetStream.pipeThrough(new DecompressionStream('deflate-raw') as any);
                }
                // Resolve the promise so the caller gets the stream immediately!
                resolve(worksheetStream);
                // DO NOT BREAK! We must keep reading the rest of the ZIP stream and dumping it
                // into WritableStreams so that `zipReader` doesn't block due to backpressure.
              } else {
                await entry.stream.pipeTo(new WritableStream());
              }
            } else {
              // Discard irrelevant files
              await entry.stream.pipeTo(new WritableStream());
            }
          }
        } catch (e) {
          reject(e);
        } finally {
          zipReader.releaseLock();
        }
      })();
    });

    const streamFound = await findWorksheetPromise;
    const xmlStream = streamFound.pipeThrough(createXmlStreamParser());
    return parseWorksheet(xmlStream, sharedStrings, styles);
  }

  private tryResolveSheetPath(workbookXml: string, relsXml: string, options: ParseOptions | undefined, setPath: (p: string) => void) {
    if (!options?.sheetName || !workbookXml || !relsXml) return;
    
    // 1. Find sheet r:id by name
    const sheetRegex = new RegExp(`<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"`, 'g');
    let match;
    let rId: string | null = null;
    while ((match = sheetRegex.exec(workbookXml)) !== null) {
      if (match[1] === options.sheetName) {
        rId = match[2];
        break;
      }
    }

    if (!rId) throw new Error(`Sheet with name "${options.sheetName}" not found in workbook.`);

    // 2. Find Target path by r:id
    const relRegex = new RegExp(`<Relationship[^>]*Id="${rId}"[^>]*Target="([^"]+)"`, 'g');
    const relMatch = relRegex.exec(relsXml);
    if (!relMatch) throw new Error(`Relationship for sheet "${options.sheetName}" not found.`);

    setPath(relMatch[1]);
  }

  private async parseSharedStrings(xmlTokenStream: ReadableStream<XmlToken>, map: Map<number, string>) {
    const reader = xmlTokenStream.getReader();
    let index = 0;
    let inText = false;
    let currentString = '';

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        const token = value as XmlToken;
        if (token.type === 'startElement' && token.name === 't') {
          inText = true;
          currentString = '';
        } else if (token.type === 'text' && inText) {
          currentString += token.value;
        } else if (token.type === 'endElement' && token.name === 't') {
          inText = false;
          map.set(index++, currentString);
        }
      }
    } finally {
      reader.releaseLock();
    }
  }

  private async parseStyles(xmlTokenStream: ReadableStream<XmlToken>, map: Map<number, number>) {
    const reader = xmlTokenStream.getReader();
    let index = 0;
    let inCellXfs = false;
    let inNumFmts = false;
    const customDateFmts = new Set<number>();

    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;

        const token = value as XmlToken;
        if (token.type === 'startElement') {
          if (token.name === 'numFmts') inNumFmts = true;
          else if (token.name === 'cellXfs') inCellXfs = true;
          else if (token.name === 'numFmt' && inNumFmts) {
            const id = parseInt(token.attributes['numFmtId'] || '0', 10);
            const formatCode = token.attributes['formatCode'] || '';
            if (/[ymdhms]/i.test(formatCode)) {
              customDateFmts.add(id);
            }
          } else if (token.name === 'xf' && inCellXfs) {
            const numFmtId = parseInt(token.attributes['numFmtId'] || '0', 10);
            const isDate = (numFmtId >= 14 && numFmtId <= 22) || (numFmtId >= 45 && numFmtId <= 47) || customDateFmts.has(numFmtId);
            map.set(index++, isDate ? 14 : 0);
          }
        } else if (token.type === 'endElement') {
          if (token.name === 'numFmts') inNumFmts = false;
          else if (token.name === 'cellXfs') inCellXfs = false;
        }
      }
    } finally {
      reader.releaseLock();
    }
  }

  private async readStreamToString(entry: ZipEntry): Promise<string> {
    let stream = entry.stream;
    if (entry.compressionMethod === 8) stream = stream.pipeThrough(new DecompressionStream('deflate-raw') as any);
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
