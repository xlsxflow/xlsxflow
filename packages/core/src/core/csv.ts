import { CellValue } from './types';

export interface CsvOptions {
  delimiter?: string;  // default ","
  convert?: boolean;   // default true: unquoted numbers and TRUE/FALSE become numbers and booleans. Empty unquoted fields are null either way
}

// Numbers as Excel would read them; "007" and "1e5x" stay text
const NUMBER = /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/;

function convertField(s: string): CellValue {
  if (s === '') return null;
  if (NUMBER.test(s)) return Number(s);
  const upper = s.toUpperCase();
  return upper === 'TRUE' ? true : upper === 'FALSE' ? false : s;
}

// Streams the rows of a CSV file (RFC 4180: quoted fields may hold delimiters, line breaks and ""),
// one row at a time. Feed it straight to SheetWriter.addSheet to convert CSV to xlsx.
export async function* parseCsv(
  input: ReadableStream<Uint8Array> | string,
  options: CsvOptions = {},
): AsyncGenerator<CellValue[]> {
  const delimiter = options.delimiter ?? ',';
  if (delimiter.length !== 1 || delimiter === '"' || delimiter === '\n' || delimiter === '\r') {
    throw new Error(`Invalid CSV delimiter ${JSON.stringify(delimiter)}.`);
  }
  const convert = options.convert !== false;

  let row: CellValue[] = [];
  let field = '';
  let quoted = false;     // the current field started with a quote
  let inQuotes = false;
  let skipLF = false;     // the previous character was a \r ending a row
  let first = true;
  const rows: CellValue[][] = [];

  const endField = () => {
    // An empty unquoted field is null in both modes; a quoted "" stays an empty string
    row.push(quoted ? field : !convert ? (field === '' ? null : field) : convertField(field));
    field = '';
    quoted = false;
  };
  const endRow = () => {
    // A blank line is an empty row
    if (row.length || field !== '' || quoted) endField();
    rows.push(row);
    row = [];
  };

  const feed = (text: string) => {
    if (first && text) {
      if (text.charCodeAt(0) === 0xfeff) text = text.slice(1); // byte order mark
      first = false;
    }
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (skipLF) {
        skipLF = false;
        if (c === '\n') continue;
      }
      if (inQuotes) {
        if (c === '"') inQuotes = false;
        else field += c;
      } else if (c === '"') {
        // A quote right after a quoted field's closing quote is an escaped ("") quote
        if (quoted) field += '"';
        else if (field === '') quoted = true;
        else { field += c; continue; } // a stray quote inside an unquoted field is kept
        inQuotes = true;
      } else if (c === delimiter) endField();
      else if (c === '\n' || c === '\r') {
        endRow();
        skipLF = c === '\r';
      } else field += c;
    }
  };

  if (typeof input === 'string') {
    feed(input);
    yield* rows.splice(0);
  } else {
    const reader = input.pipeThrough(new TextDecoderStream() as any as TransformStream<Uint8Array, string>).getReader();
    try {
      while (true) {
        const { done, value } = await reader.read();
        if (done) break;
        feed(value);
        yield* rows.splice(0);
      }
    } finally {
      // Stops reading when the caller ends the loop early
      await reader.cancel().catch(() => {});
    }
  }
  if (inQuotes) throw new Error('CSV ends inside a quoted field.');
  if (row.length || field !== '' || quoted) endRow();
  yield* rows.splice(0);
}
