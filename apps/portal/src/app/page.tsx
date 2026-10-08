import Link from "next/link";
import SheetDemo from "./sheet-demo";
import LicenceForm from "./licence-form";
import CopyCommand from "./copy-command";

const CHECKOUT_URL = process.env.NEXT_PUBLIC_CHECKOUT_URL;
const REPO = "https://github.com/xlsxflow/xlsxflow";
const CONTACT = "palikaomkar@gmail.com";

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

// Kept in step with "Compared with SheetJS and ExcelJS" in the root README
const COMPARE: [string, string, string, string][] = [
  ["Streaming .xlsx read and write", "Yes", "No (streams CSV, HTML and JSON out)", "Yes"],
  ["Cell styles, read and write", "Yes", "No (SheetJS Pro)", "Yes"],
  ["Images", "Yes", "No", "Yes"],
  [".xls", "Read", "Read and write", "No"],
  [".ods", "Read and write", "Read and write", "No"],
  [".xlsb, .numbers and other formats", "No", "Yes", "No"],
  ["Charts", "Add (Pro)", "No (SheetJS Pro)", "No"],
  ["Pivot tables", "Add (Pro)", "No (SheetJS Pro)", "Partial, undocumented"],
  ["Password-protected files", "Open and save (Pro)", "Old .xls obfuscation only (SheetJS Pro opens AES files)", "No"],
  ["Licence", "MIT, Pro is paid", "Apache 2.0", "MIT"],
];

// From the benchmarks in the root README
const BENCH: [string, string, string, string, string][] = [
  ["XlsxFlow", "2.9 s", "+2 MB", "2.3 s", "96 MB"],
  ["ExcelJS 4.4, streaming", "3.4 s", "+9 MB", "2.7 s", "263 MB"],
  ["SheetJS 0.20.3, compressed", "4.4 s", "+140 MB", "6.1 s", "549 MB"],
];

const Check = ({ label }: { label: string }) => (
  <svg viewBox="0 0 16 16" className="w-4 h-4 inline-block text-accent" role="img" aria-label={label}>
    <path d="M3 8.5l3 3 7-7" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

const link = "underline decoration-grid underline-offset-4 hover:decoration-accent";

// Read once at build time, so visitors' browsers never call GitHub. Hidden when unknown or zero.
async function starCount(): Promise<number> {
  try {
    const res = await fetch("https://api.github.com/repos/xlsxflow/xlsxflow", { headers: { Accept: "application/vnd.github+json" } });
    return res.ok ? ((await res.json()) as { stargazers_count?: number }).stargazers_count ?? 0 : 0;
  } catch {
    return 0;
  }
}

const GitHubMark = () => (
  <svg viewBox="0 0 16 16" className="w-4 h-4" fill="currentColor" aria-hidden>
    <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z" />
  </svg>
);

export default async function Home() {
  const stars = await starCount();
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
            <a href="#compare" className="hidden sm:inline hover:text-ink">Compare</a>
            <a href="#pro" className="hover:text-ink">Pro</a>
            <a href="#licence" className="hidden sm:inline hover:text-ink">Licence key</a>
            <a href={REPO} className="flex items-center gap-1.5 hover:text-ink" aria-label={stars ? `GitHub, ${stars} stars` : "GitHub"}>
              <GitHubMark />
              <span className="hidden sm:inline">GitHub</span>
              {stars > 0 && <span className="font-mono text-xs border border-grid rounded-sm px-1.5 py-0.5">★ {stars.toLocaleString("en-US")}</span>}
            </a>
          </nav>
        </div>
      </header>

      <main className="flex-1">
        {/* Phones: one column. Tablets: headline beside the intro, sheet below. Desktops: sheet on the right. */}
        <section className="max-w-6xl mx-auto px-4 sm:px-6 pt-12 sm:pt-16 pb-16 grid gap-x-8 lg:gap-x-12 gap-y-6 md:grid-cols-2 lg:grid-cols-[minmax(0,5fr)_minmax(0,7fr)]">
          <h1 className="display text-[clamp(48px,6.6vw,84px)] lg:col-start-1 lg:row-start-1 lg:self-end">Read, write and edit xlsx as a stream.</h1>
          <div className="flex flex-col gap-6 md:col-start-2 md:row-start-1 lg:col-start-1 lg:row-start-2">
            <p className="text-lg text-muted max-w-[60ch]">
              A JavaScript library with no dependencies. Rows go through one at a time, so memory stays flat whether the
              file has a hundred rows or a million. Runs in Node, Bun, browsers and Cloudflare Workers.
            </p>
            {/* Side by side only while the intro spans the page */}
            <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:gap-5 md:flex-col md:items-start md:gap-3">
              <CopyCommand command="npm install @xlsxflow/core" />
              <p className="text-sm text-muted">
                Also reads .xls and .ods, and writes .ods. Free under the MIT licence.
              </p>
            </div>
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

        <section id="compare" className="border-t border-grid scroll-mt-4">
          <div className="max-w-6xl mx-auto px-4 sm:px-6 py-16 flex flex-col gap-6">
            <h2 className="display text-[clamp(40px,5vw,56px)]">Compared with SheetJS and ExcelJS</h2>
            <p className="text-muted max-w-[68ch]">
              SheetJS reads and writes far more formats, and ExcelJS has a longer track record. XlsxFlow focuses on .xlsx:
              streaming in flat memory, keeping everything in a file it edits, and running on Web APIs alone.
            </p>
            <div className="overflow-x-auto -mx-4 px-4 sm:mx-0 sm:px-0">
              <table className="w-full min-w-160 text-[15px] border-t border-grid">
                <thead>
                  <tr className="text-left text-sm text-muted">
                    <th className="py-2 pr-4 font-medium w-[28%]"><span className="sr-only">Feature</span></th>
                    <th className="py-2 pr-4 font-medium text-ink">XlsxFlow</th>
                    <th className="py-2 pr-4 font-medium">SheetJS Community Edition</th>
                    <th className="py-2 font-medium">ExcelJS 4.4</th>
                  </tr>
                </thead>
                <tbody>
                  {COMPARE.map(([what, ours, sheetjs, exceljs]) => (
                    <tr key={what} className="border-t border-grid align-top">
                      <th scope="row" className="py-2.5 pr-4 text-left font-normal text-muted">{what}</th>
                      <td className="py-2.5 pr-4">{ours}</td>
                      <td className="py-2.5 pr-4">{sheetjs}</td>
                      <td className="py-2.5">{exceljs}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <h3 className="text-lg font-semibold mt-6">Speed and memory</h3>
            <div className="overflow-x-auto -mx-4 px-4 sm:mx-0 sm:px-0">
              <table className="w-full min-w-140 max-w-3xl text-[15px] border-t border-grid">
                <thead>
                  <tr className="text-left text-sm text-muted">
                    <th className="py-2 pr-4 font-medium"><span className="sr-only">Library</span></th>
                    <th className="py-2 px-3 font-medium text-right">Write 1M cells</th>
                    <th className="py-2 px-3 font-medium text-right">Heap growth</th>
                    <th className="py-2 px-3 font-medium text-right">Read 1M cells</th>
                    <th className="py-2 pl-3 font-medium text-right">Peak memory</th>
                  </tr>
                </thead>
                <tbody>
                  {BENCH.map(([name, ...cells]) => (
                    <tr key={name} className="border-t border-grid">
                      <th scope="row" className={`py-2.5 pr-4 text-left ${name === "XlsxFlow" ? "font-semibold" : "font-normal"}`}>{name}</th>
                      {cells.map((c, i) => <td key={i} className={`py-2.5 text-right ${i === 3 ? "pl-3" : "px-3"}`}>{c}</td>)}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="text-sm text-muted max-w-[80ch] flex flex-col gap-2">
              <p>
                Features were checked against each project&apos;s own documentation on 8 October 2026:{" "}
                <a href="https://docs.sheetjs.com/docs/miscellany/formats" className={link}>SheetJS formats</a>,{" "}
                <a href="https://sheetjs.com/pro" className={link}>SheetJS Pro</a> and the{" "}
                <a href="https://github.com/exceljs/exceljs#readme" className={link}>ExcelJS README</a>. ExcelJS&apos;s last
                release was in October 2023.
              </p>
              <p>
                Timings are from one laptop on Node 25 and varied by up to 2× between runs, so treat differences under
                about 20% as a tie. The <a href={`${REPO}#benchmarks`} className={link}>method, more libraries and the
                scripts</a> are in the README. Spotted something out of date? <a href={`mailto:${CONTACT}`} className={link}>Tell us</a> and
                we&apos;ll correct it.
              </p>
              <p>SheetJS and ExcelJS belong to their respective owners. XlsxFlow isn&apos;t affiliated with or endorsed by either project.</p>
            </div>
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
                <li>$5 per developer, one-time payment.</li>
                <li>Perpetual licence, with every version released within a year of your order.</li>
                <li>Build servers and CI don&apos;t need their own key.</li>
                <li>Full refund within 14 days if it doesn&apos;t work for you, processed by Polar.</li>
              </ul>
              <div>
                <h3 className="font-semibold mb-2">How it works</h3>
                <ol className="flex flex-col gap-2 list-decimal pl-5 text-muted">
                  <li>Install it: <code className="font-mono text-sm text-ink">npm install @xlsxflow/core @xlsxflow/pro</code></li>
                  <li>Buy Pro. Polar emails you a receipt with your order ID.</li>
                  <li>Enter the order ID and your email in <a href="#licence" className={`${link} text-ink`}>Get your licence key</a> below.</li>
                  <li>Call <code className="font-mono text-sm text-ink">await setLicenseKey(key)</code> once at startup. Without a key, Pro functions throw an error.</li>
                </ol>
              </div>
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
