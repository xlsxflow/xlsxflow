# Recipes

Copy-paste examples for common jobs. Each one runs in CI against the current build, so they stay correct as the library changes.

| Recipe | For |
|---|---|
| [web-response.mjs](web-response.mjs) | Sending a generated `.xlsx` as a download from Next.js route handlers, Cloudflare Workers, Bun, Deno, Hono or Remix, streamed while it is written |
| [node-http.mjs](node-http.mjs) | The same from Express, Fastify or plain `node:http` |
| [upload.mjs](upload.mjs) | Reading an uploaded `.xlsx`, `.xls` or `.ods` file into JSON on the server, with limits for untrusted files |
| [database-export.mjs](database-export.mjs) | Exporting a database table without loading it into memory (`node:sqlite`; the same pattern works with any cursor) |
| [browser-download.mjs](browser-download.mjs) | An "Export to Excel" button in the browser, with no server |

Moving from another library? See [Migrating from ExcelJS](../docs/migrating-from-exceljs.md) and [Migrating from SheetJS](../docs/migrating-from-sheetjs.md).
