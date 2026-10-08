import Link from "next/link";
import SheetDemo from "./sheet-demo";
import LicenceForm from "./licence-form";
import CopyCommand from "./copy-command";

const CHECKOUT_URL = process.env.NEXT_PUBLIC_CHECKOUT_URL;
const REPO = "https://github.com/xlsxflow/xlsxflow";
const CONTACT = "palikaomkar.22.cse@anits.edu.in";

const READ = `import { SheetReader, createBlobReader } from '@xlsxflow/core';

const rows = await new SheetReader().parse(createBlobReader(file));
for await (const row of rows) {
  console.log(row.rowNumber, row.cells); // 1 ['Order', 'Date', …]
}`;

const WRITE = `import { SheetWriter } from '@xlsxflow/core';

const writer = new SheetWriter();
writer.addSheet('Report', rowsFromDatabase()); // an array or async iterable
const xlsx = writer.write(); // ReadableStream<Uint8Array>
return new Response(xlsx);`;

const FEATURES: [string, boolean][] = [
  ["Read .xlsx, .xlsm, .xls and .ods", true],
  ["Write and edit .xlsx, write .ods", true],
  ["Styles, formulas, images, tables, notes and conditional formats", true],
  ["Fill Excel templates with data, with rows that repeat for lists", false],
  ["Column, bar, line, area and pie charts", false],
  ["Pivot tables", false],
  ["Open and save password-protected .xlsx files", false],
];

const Check = ({ label }: { label: string }) => (
  <svg viewBox="0 0 16 16" className="w-4 h-4 inline-block text-accent" role="img" aria-label={label}>
    <path d="M3 8.5l3 3 7-7" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const link = "underline decoration-grid underline-offset-4 hover:decoration-accent";

export default function Home() {
  return (
    <div className="flex-1 flex flex-col">
      <header className="border-b border-grid">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 h-14 flex items-center justify-between gap-4">
          <Link href="/" className="flex items-center gap-2 font-semibold text-[17px]">
            <img src="/icon.svg" alt="" className="w-7 h-7" />
            XlsxFlow
          </Link>
          <nav className="flex items-center gap-4 sm:gap-6 text-sm text-muted">
            <a href={`${REPO}#readme`} className="hover:text-ink">Docs</a>
            <a href="#pro" className="hover:text-ink">Pro</a>
            <a href="#licence" className="hidden sm:inline hover:text-ink">Licence key</a>
            <a href={REPO} className="hover:text-ink">GitHub</a>
          </nav>
        </div>
      </header>

      <main className="flex-1">
        {/* Phones: one column. Tablets: headline beside the intro, sheet below. Desktops: sheet on the right. */}
        <section className="max-w-6xl mx-auto px-4 sm:px-6 pt-12 sm:pt-16 pb-16 grid gap-x-8 lg:gap-x-12 gap-y-6 md:grid-cols-2 lg:grid-cols-[minmax(0,5fr)_minmax(0,7fr)]">
          <h1 className="display text-[clamp(48px,6.6vw,84px)] lg:col-start-1 lg:row-start-1 lg:self-end">Read, write and edit xlsx as a stream.</h1>
          <div className="flex flex-col gap-6 md:col-start-2 md:row-start-1 lg:col-start-1 lg:row-start-2">
            <p className="text-lg text-muted max-w-[38ch]">
              A JavaScript library with no dependencies. Rows go through one at a time, so memory stays flat whether the
              file has a hundred rows or a million. Runs in Node, Bun, browsers and Cloudflare Workers.
            </p>
            <CopyCommand command="npm install @xlsxflow/core" />
            <p className="text-sm text-muted">
              Also reads .xls and .ods, and writes .ods. Free under the MIT licence.
            </p>
          </div>
          <div className="min-w-0 mt-4 md:col-span-2 lg:mt-0 lg:col-span-1 lg:col-start-2 lg:row-start-1 lg:row-span-2 lg:self-center">
            <SheetDemo />
            <p className="mt-2 text-sm text-muted">
              This sample was written by SheetWriter in your browser a moment ago, then read back by SheetReader. Open
              your own file to try it; it stays on your computer.
            </p>
          </div>
        </section>

        <section className="border-t border-grid">
          <div className="max-w-6xl mx-auto px-4 sm:px-6 py-16 grid gap-8 lg:grid-cols-2">
            {[["Read a file", READ], ["Stream one out of a server", WRITE]].map(([title, code]) => (
              <div key={title} className="min-w-0">
                <h2 className="font-semibold mb-3">{title}</h2>
                <pre className="bg-head border border-grid p-4 text-[13px] leading-relaxed font-mono overflow-x-auto"><code>{code}</code></pre>
              </div>
            ))}
            <p className="lg:col-span-2 text-muted">
              Styles, formulas, merges, frozen panes, images and more are covered in the <a href={`${REPO}#readme`} className={`${link} text-ink`}>documentation</a>.
            </p>
          </div>
        </section>

        <section id="pro" className="border-t border-grid scroll-mt-4">
          <div className="max-w-6xl mx-auto px-4 sm:px-6 py-16 grid gap-10 md:grid-cols-[minmax(0,7fr)_minmax(0,5fr)]">
            <div className="min-w-0">
              <h2 className="display text-[clamp(40px,5vw,56px)] mb-6">Free, and Pro for $5.</h2>
              <div className="overflow-x-auto">
                <table className="w-full text-[15px] border-t border-grid">
                  <thead>
                    <tr className="text-left text-sm text-muted">
                      <th className="py-2 pr-4 font-medium">What it does</th>
                      <th className="py-2 px-3 font-medium text-center w-16 sm:w-24">Core</th>
                      <th className="py-2 pl-3 font-medium text-center w-16 sm:w-24">Pro</th>
                    </tr>
                  </thead>
                  <tbody>
                    {FEATURES.map(([what, core]) => (
                      <tr key={what} className="border-t border-grid">
                        <td className="py-2.5 pr-4">{what}</td>
                        <td className="py-2.5 px-3 text-center">{core ? <Check label="Included" /> : <span className="sr-only">Not included</span>}</td>
                        <td className="py-2.5 pl-3 text-center"><Check label="Included" /></td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="flex flex-col gap-5 md:pt-3">
              <p>
                <strong className="font-semibold">@xlsxflow/pro</strong> is an add-on to the core package for the work
                that usually needs Excel itself: templates, charts, pivot tables and passwords.
              </p>
              <ul className="flex flex-col gap-2 text-muted list-disc pl-5">
                <li>$5 per developer, with local pricing at checkout.</li>
                <li>Perpetual licence, with every version released within a year of your order.</li>
                <li>Build servers and CI don&apos;t need their own key.</li>
                <li>Full refund within 14 days if it doesn&apos;t work for you, processed by Polar.</li>
              </ul>
              <div className="flex flex-wrap items-center gap-x-5 gap-y-3">
                {CHECKOUT_URL && (
                  <a href={CHECKOUT_URL} className="px-5 py-2.5 rounded-sm bg-accent text-accent-ink font-medium hover:opacity-90">Buy Pro for $5</a>
                )}
                <a href="/pro-licence.txt" className={link}>Licence agreement</a>
              </div>
            </div>
          </div>
        </section>

        <section id="licence" className="border-t border-grid scroll-mt-4">
          <div className="max-w-6xl mx-auto px-4 sm:px-6 py-16 grid gap-8 md:grid-cols-[minmax(0,5fr)_minmax(0,7fr)]">
            <div className="flex flex-col gap-4 max-w-prose">
              <h2 className="text-2xl font-semibold">Get your licence key</h2>
              <p className="text-muted">
                Bought Pro? Enter the order ID from your Polar receipt and the email you used. The key is made on the spot
                and works offline; nothing is stored here.
              </p>
              <p className="text-muted">
                Lost the receipt or stuck? Email <a href={`mailto:${CONTACT}`} className={`${link} text-ink`}>{CONTACT}</a>.
              </p>
            </div>
            <div className="max-w-xl w-full">
              <LicenceForm />
            </div>
          </div>
        </section>
      </main>

      <footer className="border-t border-grid text-sm text-muted">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 py-6 flex flex-wrap items-center justify-between gap-4">
          <span>XlsxFlow</span>
          <nav className="flex flex-wrap gap-x-5 gap-y-2">
            <a href={REPO} className="hover:text-ink">GitHub</a>
            <a href="https://www.npmjs.com/package/@xlsxflow/core" className="hover:text-ink">npm</a>
            <a href="/pro-licence.txt" className="hover:text-ink">Pro licence</a>
            <a href="/privacy" className="hover:text-ink">Privacy</a>
            <a href={`mailto:${CONTACT}`} className="hover:text-ink">Contact</a>
          </nav>
        </div>
      </footer>
    </div>
  );
}
