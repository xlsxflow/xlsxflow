// Checks the README's claims against the built library (pnpm build first). claims.mjs needs only Web APIs,
// so the same checks run on every runtime:
//   node scripts/claims/run.mjs
//   bun scripts/claims/run.mjs
//   deno run --allow-read scripts/claims/run.mjs
//   node scripts/claims/run.mjs --browsers    (Chromium, Firefox and WebKit via Playwright)
//   node scripts/claims/run.mjs --worker      (Cloudflare Workers via wrangler dev, without nodejs_compat)
import { readFileSync } from 'node:fs';
import { run, FIXTURES } from './claims.mjs';

function report(where, results) {
  for (const r of results) if (!r.ok) console.log(`FAIL ${r.name}\n     ${r.error}`);
  const passed = results.filter((r) => r.ok).length;
  console.log(`${where}: ${passed}/${results.length} README claims passed`);
  return passed === results.length && passed > 0;
}

const mode = globalThis.Deno ? Deno.args[0] : process.argv[2];
let ok = true;
if (mode === '--browsers') ok = await (await import('./browsers.mjs')).default(report);
else if (mode === '--worker') ok = await (await import('./worker-run.mjs')).default(report);
else {
  const lib = await import('../../dist/index.mjs');
  const fx = Object.fromEntries(Object.entries(FIXTURES).map(([k, f]) => [k, new Uint8Array(readFileSync(new URL(`../../test/fixtures/${f}`, import.meta.url)))]));
  const where = globalThis.Bun ? `Bun ${Bun.version}` : globalThis.Deno ? `Deno ${Deno.version.deno}` : `Node ${process.version}`;
  ok = report(where, await run(lib, fx));
}
if (!ok) process.exit(1);
