// Read a spreadsheet uploaded with a form (<input type="file" name="file">) on a server that uses Web
// Requests: Next.js route handlers, Cloudflare Workers, Bun, Deno, Hono and Remix.
// Rows are read one at a time, so a large upload doesn't have to fit in memory as a workbook.
import { SheetReader, createBlobReader } from '@xlsxflow/core';

const MAX_ROWS = 10_000;

export async function POST(request) {
  const file = (await request.formData()).get('file');
  if (!(file instanceof Blob)) return Response.json({ error: 'No file uploaded' }, { status: 400 });

  try {
    // Uploads are untrusted: cap how much the file may unpack to (the default is 1 GiB per part).
    // .xlsx, .xls and .ods are all read the same way.
    const rows = await new SheetReader().parse(createBlobReader(file), { maxUncompressedBytes: 50_000_000 });
    let header;
    const records = [];
    for await (const row of rows) {
      if (!header) { header = row.cells.map(String); continue; }
      if (records.length === MAX_ROWS) return Response.json({ error: `More than ${MAX_ROWS} rows` }, { status: 413 });
      records.push(Object.fromEntries(header.map((key, i) => [key, row.cells[i] ?? null])));
    }
    return Response.json({ columns: header ?? [], records });
  } catch (e) {
    // Not a spreadsheet, damaged, password-protected, or over the size limit
    return Response.json({ error: e.message }, { status: 422 });
  }
}
