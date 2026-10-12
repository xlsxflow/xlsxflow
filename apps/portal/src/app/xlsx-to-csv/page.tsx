import type { Metadata } from "next";
import Converter from "../converter";
import PageShell, { REPO, link } from "../page-shell";

export const metadata: Metadata = {
  title: "Excel to CSV converter (.xlsx, .xls, .ods), in your browser · XlsxFlow",
  description: "Convert an Excel .xlsx, .xls or .ods sheet to CSV in your browser. Nothing is uploaded; pick the sheet and the separator.",
  alternates: { canonical: "/xlsx-to-csv" },
  openGraph: { title: "Excel to CSV converter · XlsxFlow", url: "/xlsx-to-csv" },
};

const CODE = `import { SheetReader, createBlobReader } from '@xlsxflow/core';

const rows = await new SheetReader().parse(createBlobReader(file), { formatted: true });
for await (const row of rows) {
  console.log(row.formatted); // each cell's text as Excel shows it
}`;

export default function XlsxToCsv() {
  return (
    <PageShell>
      <div className="max-w-3xl mx-auto px-4 sm:px-6 py-12 sm:py-16 flex flex-col gap-8">
        <div className="flex flex-col gap-4">
          <h1 className="display text-[clamp(40px,6vw,64px)]">Excel to CSV converter</h1>
          <p className="text-lg text-muted max-w-[60ch]">
            Save a sheet from an .xlsx, .xls or .ods file as CSV. It runs in your browser, so the file never leaves your
            computer, and rows are read one at a time, so large workbooks work too.
          </p>
        </div>

        <Converter mode="xlsx-to-csv" />

        <section className="flex flex-col gap-4 leading-relaxed">
          <h2 className="text-xl font-semibold">What you get</h2>
          <ul className="list-disc pl-5 flex flex-col gap-2 text-muted">
            <li>Each cell as Excel shows it: dates, currencies and percentages keep their formatting, and formulas give their results.</li>
            <li>A UTF-8 file that Excel opens with accents and other scripts intact.</li>
            <li>Comma, semicolon or tab separators. Many European versions of Excel expect semicolons.</li>
            <li>Empty rows stay where they were, so row numbers line up with the original sheet.</li>
            <li>One sheet per file; choose which one if the workbook has several.</li>
          </ul>
        </section>

        <section className="flex flex-col gap-4 leading-relaxed">
          <h2 className="text-xl font-semibold">Do the same in your own code</h2>
          <p className="text-muted">
            This page is built on <a href={REPO} className={`${link} text-ink`}>XlsxFlow</a>, an open-source JavaScript
            library that reads spreadsheets as a stream in browsers, Node, Bun, Deno and Cloudflare Workers:
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
