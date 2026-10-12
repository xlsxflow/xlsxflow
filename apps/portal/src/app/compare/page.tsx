import type { Metadata } from "next";
import PageShell, { REPO, link } from "../page-shell";
import CompareTool from "./compare-tool";

export const metadata: Metadata = {
  title: "Compare XlsxFlow, ExcelJS and SheetJS on your own Excel file · XlsxFlow",
  description: "Read your own .xlsx file with XlsxFlow, ExcelJS and SheetJS in your browser and compare time, memory and what each one found. Nothing is uploaded.",
  alternates: { canonical: "/compare" },
  openGraph: { title: "Compare Excel libraries on your own file · XlsxFlow", url: "/compare" },
};

export default function Compare() {
  return (
    <PageShell>
      <div className="max-w-4xl mx-auto px-4 sm:px-6 py-12 sm:py-16 flex flex-col gap-8">
        <div className="flex flex-col gap-4">
          <h1 className="display text-[clamp(40px,6vw,64px)]">Compare on your own file</h1>
          <p className="text-lg text-muted max-w-[62ch]">
            Benchmarks on someone else&apos;s data only go so far. Open one of your spreadsheets and it&apos;s read by
            XlsxFlow, ExcelJS and SheetJS, one after another, right here in your browser. Nothing is uploaded.
          </p>
        </div>

        <CompareTool />

        <section className="flex flex-col gap-4 leading-relaxed">
          <h2 className="text-xl font-semibold">How it&apos;s measured</h2>
          <ul className="list-disc pl-5 flex flex-col gap-2 text-muted">
            <li>Each library reads the first sheet in its own new Web Worker, so one doesn&apos;t slow down or add memory to the next. Loading the library&apos;s code isn&apos;t timed.</li>
            <li>Each is used the lightest way it offers in a browser: XlsxFlow streams the rows, ExcelJS loads the workbook (its streaming reader needs Node), and SheetJS reads only the first sheet in dense mode.</li>
            <li><strong className="font-medium text-ink">Memory when done</strong> is what the worker uses once reading finishes, including the library&apos;s code. ExcelJS and SheetJS hand you the whole workbook, so it&apos;s still in memory; XlsxFlow hands you rows one at a time and keeps none. Browsers report this only in Chrome and Edge.</li>
            <li><strong className="font-medium text-ink">Rows</strong> and <strong className="font-medium text-ink">Cells</strong> count those with a value. A count that differs from XlsxFlow&apos;s is shown in red; open the file to see which is right, and <a href={`${REPO}/issues`} className={`${link} text-ink`}>tell us</a> if it&apos;s XlsxFlow that&apos;s wrong.</li>
            <li>It&apos;s one run on your machine, so small differences are noise. The <a href={`${REPO}#benchmarks`} className={`${link} text-ink`}>benchmarks in the README</a> use medians of several runs in Node, with the scripts to repeat them.</li>
          </ul>
        </section>

        <section className="flex flex-col gap-3 leading-relaxed">
          <h2 className="text-xl font-semibold">What the numbers don&apos;t show</h2>
          <p className="text-muted">
            SheetJS reads many more formats, and ExcelJS has years of use behind it. XlsxFlow can&apos;t jump to any cell
            at any time, because rows stream past once. If that fits how you use spreadsheets, there are guides for
            moving from <a href="https://github.com/xlsxflow/xlsxflow/blob/main/docs/migrating-from-exceljs.md" className={`${link} text-ink`}>ExcelJS</a> and{" "}
            <a href="https://github.com/xlsxflow/xlsxflow/blob/main/docs/migrating-from-sheetjs.md" className={`${link} text-ink`}>SheetJS</a>.
          </p>
          <p className="text-sm text-muted">SheetJS and ExcelJS belong to their respective owners. XlsxFlow isn&apos;t affiliated with or endorsed by either project.</p>
        </section>
      </div>
    </PageShell>
  );
}
