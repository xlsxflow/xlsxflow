// Runs the built library inside Cloudflare's Workers runtime (workerd), without nodejs_compat.
// From packages/core, after `pnpm build`:
//   npx wrangler@4 dev --config scripts/workers/wrangler.toml
// then open http://localhost:8787. It answers {"ok":true} with the checks that passed, or 500 with the error.
import { SheetWriter, SheetReader, SheetEditor, createBlobReader } from '../../dist/index.mjs';

const bytes = async (s) => new Uint8Array(await new Response(s).arrayBuffer());
const reader = (b) => createBlobReader(new Blob([b]));
async function rows(b, options) {
  const out = [];
  for await (const r of await new SheetReader().parse(reader(b), options)) out.push(r);
  return out;
}
function check(name, ok, detail) {
  if (!ok) throw new Error(`${name} failed: ${JSON.stringify(detail)}`);
  return name;
}

export default {
  async fetch() {
    const passed = [];
    try {
      async function* many() { for (let i = 0; i < 100000; i++) yield [i, `row ${i}`, i * 0.5]; }
      const big = await bytes(new SheetWriter().addSheet('Big', many()).write());
      let count = 0, last;
      for await (const r of await new SheetReader().parse(reader(big))) { count++; last = r.cells; }
      passed.push(check('stream 100k rows', count === 100000 && last[0] === 99999, { count, last }));

      const w = new SheetWriter({ properties: { title: 'Edge' }, definedNames: [{ name: 'Rate', ref: '0.18' }] });
      w.addSheet('Hidden', [['x']], { state: 'hidden' });
      w.addSheet('Data', [
        [{ value: 1234.5, style: { numFmt: '#,##0.00', font: { bold: true } } }, new Date(Date.UTC(2026, 9, 8)), { value: null, formula: 'A1*Rate' }],
      ], { view: { zoom: 120 } });
      const book = await bytes(w.write());
      const info = await new SheetReader().readWorkbook(reader(book));
      passed.push(check('workbook info', info.properties.title === 'Edge' && info.definedNames[0]?.name === 'Rate' && info.sheets[0].state === 'hidden', info));
      const data = await rows(book, { sheetName: 'Data', formatted: true, styles: true });
      passed.push(check('formatted text', data[0].formatted?.[0] === '1,234.50' && data[0].formatted?.[1] === '2026-10-08', data[0]));

      const edited = await bytes(new SheetEditor().setCells('Data', { B2: { value: 'added' } }).insertRows('Data', 1, 1).edit(reader(book)));
      const after = await rows(edited, { sheetName: 'Data' });
      passed.push(check('edit', after.some(r => r.cells.includes('added')), after.map(r => r.cells)));

      return Response.json({ ok: true, passed });
    } catch (e) {
      return Response.json({ ok: false, passed, error: String(e?.stack ?? e) }, { status: 500 });
    }
  },
};
