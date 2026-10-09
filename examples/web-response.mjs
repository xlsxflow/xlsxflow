// Send a generated .xlsx as a download from any handler that returns a Web Response:
// Next.js route handlers, Cloudflare Workers, Bun.serve, Deno.serve, Hono and Remix.
// The file streams out while it is written, so memory stays flat however many rows there are.
import { SheetWriter } from '@xlsxflow/core';

// Stand-in for your data source: a database cursor, paginated API calls, anything async
async function* orders() {
  for (let id = 1; id <= 10_000; id++) yield { id, customer: `Customer ${id % 50}`, total: id * 1.5, placed: new Date(Date.UTC(2026, 0, 1 + (id % 365))) };
}

async function* rows() {
  yield ['Order', 'Customer', 'Total', 'Placed'];
  for await (const o of orders()) yield [o.id, o.customer, { value: o.total, style: { numFmt: '#,##0.00' } }, o.placed];
}

export function GET() {
  const xlsx = new SheetWriter().addSheet('Orders', rows(), { freezePanes: { row: 1 }, columnWidths: [10, 16, 12, 12] }).write();
  return new Response(xlsx, {
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="orders.xlsx"',
    },
  });
}

// Next.js: save as app/orders/route.js (it uses the GET export).
// Cloudflare Workers, Bun and Deno: export this default, or pass GET to Bun.serve({ fetch }) / Deno.serve().
export default { fetch: GET };
