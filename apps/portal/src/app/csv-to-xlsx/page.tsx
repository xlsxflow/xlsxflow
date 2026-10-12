import type { Metadata } from "next";
import Converter from "../converter";
import PageShell, { REPO, link } from "../page-shell";

export const metadata: Metadata = {
  title: "CSV to Excel (.xlsx) converter, in your browser · XlsxFlow",
  description: "Convert a CSV file to an Excel .xlsx workbook in your browser. Nothing is uploaded, large files work, and leading zeros are kept.",
  alternates: { canonical: "/csv-to-xlsx" },
  openGraph: { title: "CSV to Excel (.xlsx) converter · XlsxFlow", url: "/csv-to-xlsx" },
};

const CODE = `import { SheetWriter, parseCsv } from '@xlsxflow/core';

// Rows stream from the CSV into the .xlsx one at a time
const xlsx = new SheetWriter().addSheet('Data', parseCsv(file.stream())).write();
const blob = await new Response(xlsx).blob();`;

export default function CsvToXlsx() {
  return (
    <PageShell>
      <div className="max-w-3xl mx-auto px-4 sm:px-6 py-12 sm:py-16 flex flex-col gap-8">
        <div className="flex flex-col gap-4">
          <h1 className="display text-[clamp(40px,6vw,64px)]">CSV to Excel converter</h1>
          <p className="text-lg text-muted max-w-[60ch]">
            Turn a .csv file into an .xlsx workbook. It runs in your browser, so the file never leaves your computer,
            and it streams, so files with hundreds of thousands of rows work too.
          </p>
        </div>

        <Converter mode="csv-to-xlsx" />

        <section className="flex flex-col gap-4 leading-relaxed">
          <h2 className="text-xl font-semibold">What you get</h2>
          <ul className="list-disc pl-5 flex flex-col gap-2 text-muted">
            <li>Numbers become real numbers you can add up. Values with leading zeros, such as <code className="font-mono text-sm">007</code> or phone numbers, stay as text, so the zeros are kept.</li>
            <li>The header row is bold and frozen, and columns are sized to fit the first 100 rows.</li>
            <li>Commas, semicolons, tabs and pipes are detected from the first line. Quoted fields can contain line breaks and quotes.</li>
            <li>Dates stay as text, because a CSV file doesn&apos;t say whether 03/04 is March or April.</li>
            <li>Excel holds 1,048,576 rows per sheet; anything beyond that is left out, and you&apos;ll be told.</li>
          </ul>
        </section>

        <section className="flex flex-col gap-4 leading-relaxed">
          <h2 className="text-xl font-semibold">Do the same in your own code</h2>
          <p className="text-muted">
            This page is built on <a href={REPO} className={`${link} text-ink`}>XlsxFlow</a>, an open-source JavaScript
            library. The converter above comes down to this, and the same code runs in Node, Bun, Deno and Cloudflare Workers:
          </p>
          <pre className="bg-head border border-grid p-4 text-[13px] leading-relaxed font-mono overflow-x-auto"><code>{CODE}</code></pre>
          <p className="text-muted">
            <code className="font-mono text-sm text-ink">npm install @xlsxflow/core</code>. MIT licence, no dependencies.
          </p>
        </section>
      </div>
    </PageShell>
  );
}
