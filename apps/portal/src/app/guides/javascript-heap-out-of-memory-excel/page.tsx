import type { Metadata } from "next";
import Link from "next/link";
import PageShell, { REPO, link } from "../../page-shell";

export const metadata: Metadata = {
  title: "Fixing \"JavaScript heap out of memory\" when reading or writing Excel files · XlsxFlow",
  description: "Why ExcelJS and SheetJS run out of memory on large .xlsx files in Node, and how to fix it: streaming readers and writers, dense mode, and not loading every row at once.",
  alternates: { canonical: "/guides/javascript-heap-out-of-memory-excel" },
  openGraph: { title: "Fixing \"JavaScript heap out of memory\" with Excel files", url: "/guides/javascript-heap-out-of-memory-excel" },
};

const pre = "bg-head border border-grid p-4 text-[13px] leading-relaxed font-mono overflow-x-auto";
const code = "font-mono text-sm";

const ERROR = `FATAL ERROR: Reached heap limit Allocation failed - JavaScript heap out of memory`;

const EXCELJS_WRITE = `import ExcelJS from 'exceljs';

const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: 'big.xlsx' });
const sheet = workbook.addWorksheet('Data');
for await (const user of db.users.cursor()) sheet.addRow([user.id, user.name]).commit();
sheet.commit();
await workbook.commit();`;

const EXCELJS_READ = `const reader = new ExcelJS.stream.xlsx.WorkbookReader('big.xlsx');
for await (const worksheet of reader) {
  for await (const row of worksheet) console.log(row.values);
}`;

const SHEETJS = `const wb = XLSX.readFile('big.xlsx', { dense: true });`;

const XLSXFLOW_WRITE = `import { createWriteStream } from 'node:fs';
import { Writable } from 'node:stream';
import { SheetWriter } from '@xlsxflow/core';

async function* rows() {
  yield ['ID', 'Name'];
  for await (const user of db.users.cursor()) yield [user.id, user.name];
}
await new SheetWriter().addSheet('Users', rows()).write()
  .pipeTo(Writable.toWeb(createWriteStream('users.xlsx')));`;

const XLSXFLOW_READ = `import { XlsxFlow } from '@xlsxflow/core';

for await (const row of await XlsxFlow.readFile('big.xlsx')) {
  console.log(row.rowNumber, row.cells);
}`;

export default function HeapGuide() {
  return (
    <PageShell>
      <article className="max-w-3xl mx-auto px-4 sm:px-6 py-12 sm:py-16 flex flex-col gap-5 leading-relaxed">
        <h1 className="display text-[clamp(36px,5.5vw,56px)]">&ldquo;JavaScript heap out of memory&rdquo; with Excel files</h1>
        <p className="text-lg text-muted">
          You export a report to .xlsx, or read an uploaded spreadsheet, and Node stops with:
        </p>
        <pre className={pre}><code>{ERROR}</code></pre>
        <p>
          The usual cause is that the whole workbook is held in memory at once. A spreadsheet library keeps an object
          for every cell, which takes many times the file&apos;s size: in our{" "}
          <a href={`${REPO}#benchmarks`} className={link}>read benchmark</a>, a 6.4 MB file took 549 MB of memory with
          SheetJS and 667 MB with ExcelJS&apos;s normal reader. A file ten times bigger needs more than Node&apos;s default heap.
        </p>

        <h2 className="text-2xl font-semibold mt-6">1. Check where the rows come from</h2>
        <p>
          A streaming writer can&apos;t help if your code first loads every row from the database into an array.
          Fetch the rows in pages or through a cursor, and hand them to the writer as you go. The examples below use
          <code className={code}> db.users.cursor()</code> as a stand-in for that.
        </p>

        <h2 className="text-2xl font-semibold mt-6">2. If you use ExcelJS: switch to its streaming classes</h2>
        <p>
          <code className={code}>workbook.xlsx.writeFile</code>, <code className={code}>writeBuffer</code> and{" "}
          <code className={code}>readFile</code> build the whole workbook in memory. ExcelJS has separate streaming
          classes. For writing, commit each row as you add it:
        </p>
        <pre className={pre}><code>{EXCELJS_WRITE}</code></pre>
        <p>For reading:</p>
        <pre className={pre}><code>{EXCELJS_READ}</code></pre>
        <p className="text-muted">
          The streaming classes don&apos;t support everything the normal ones do, so check the features you rely on.
          ExcelJS&apos;s last release was in October 2023.
        </p>

        <h2 className="text-2xl font-semibold mt-6">3. If you use SheetJS: dense mode helps, but only so far</h2>
        <p>
          SheetJS reads the whole workbook before you see a cell, and builds the whole file before writing it. Dense
          mode stores cells in arrays rather than one object per cell address, which uses less memory:
        </p>
        <pre className={pre}><code>{SHEETJS}</code></pre>
        <p className="text-muted">
          It doesn&apos;t make reading or writing .xlsx stream, so very large files still don&apos;t fit. Note also that
          the <code className={code}>xlsx</code> package on npm is stuck at 0.18.5 from 2022, with two published
          security advisories; newer versions are installed from the SheetJS CDN.
        </p>

        <h2 className="text-2xl font-semibold mt-6">4. Raising the heap limit</h2>
        <p>
          <code className={code}>node --max-old-space-size=4096 app.js</code> buys time, but the memory needed still
          grows with the file, and serverless functions and containers have a hard limit you can&apos;t raise this way.
        </p>

        <h2 className="text-2xl font-semibold mt-6">5. Use a library that streams by default</h2>
        <p>
          I wrote <a href={REPO} className={link}>XlsxFlow</a> for this case, so take this section with that in mind.
          It reads and writes .xlsx one row at a time with a single API, so there&apos;s no separate streaming mode to
          switch to. Its CI writes and reads back 10 million cells with Node&apos;s heap capped at 32 MB. Writing:
        </p>
        <pre className={pre}><code>{XLSXFLOW_WRITE}</code></pre>
        <p>Reading:</p>
        <pre className={pre}><code>{XLSXFLOW_READ}</code></pre>
        <p>
          It also runs in browsers, Bun, Deno and Cloudflare Workers. What it doesn&apos;t do: jump to any cell at any
          time, since rows stream past once. If you&apos;re moving from another library, there are guides for{" "}
          <a href="https://github.com/xlsxflow/xlsxflow/blob/main/docs/migrating-from-exceljs.md" className={link}>ExcelJS</a> and{" "}
          <a href="https://github.com/xlsxflow/xlsxflow/blob/main/docs/migrating-from-sheetjs.md" className={link}>SheetJS</a>, and{" "}
          <a href="https://github.com/xlsxflow/xlsxflow/blob/main/examples/README.md" className={link}>recipes</a> for
          Express, Next.js, Workers and database exports.
        </p>
        <p className="text-muted">
          Need a one-off conversion rather than code? The <Link href="/csv-to-xlsx" className={link}>CSV to Excel</Link> and{" "}
          <Link href="/xlsx-to-csv" className={link}>Excel to CSV</Link> converters stream in your browser.
        </p>
      </article>
    </PageShell>
  );
}
