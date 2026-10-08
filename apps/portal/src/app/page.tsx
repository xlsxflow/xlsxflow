"use client";

import { useState, useCallback, useRef } from "react";
import { validateOrderAndGenerateLicense } from "./actions";
import { SheetReader, SheetWriter, createBlobReader } from "@xlsxflow/core";
import type { Row } from "@xlsxflow/core";

type CellValue = string | number | boolean | null;

const CHECKOUT_URL = process.env.NEXT_PUBLIC_CHECKOUT_URL;
const CONTACT = "palikaomkar.22.cse@anits.edu.in";

// 0 -> A, 25 -> Z, 26 -> AA
const columnName = (i: number): string => (i >= 26 ? columnName(Math.floor(i / 26) - 1) : "") + String.fromCharCode(65 + (i % 26));

export default function Home() {
  const [orderId, setOrderId] = useState("");
  const [email, setEmail] = useState("");
  const [license, setLicense] = useState<string | null>(null);
  const [licenseError, setLicenseError] = useState<string | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);

  const [rows, setRows] = useState<CellValue[][]>([]);
  const [fileName, setFileName] = useState<string | null>(null);
  const [isParsing, setIsParsing] = useState(false);
  const [parseError, setParseError] = useState<string | null>(null);
  const [isDragging, setIsDragging] = useState(false);

  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleGenerateLicense = async () => {
    setLicenseError(null);
    setIsGenerating(true);
    try {
      const result = await validateOrderAndGenerateLicense(orderId, email);
      if (result.success) {
        setLicense(result.license!);
      } else {
        setLicenseError(result.error ?? "Unknown error");
      }
    } catch (e) {
      setLicenseError(e instanceof Error ? e.message : String(e));
    } finally {
      setIsGenerating(false);
    }
  };

  const downloadLicense = () => {
    if (!license) return;
    const blob = new Blob([license], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'xlsxflow-licence.txt';
    a.click();
    URL.revokeObjectURL(url);
  };

  const parseFile = useCallback(async (file: File) => {
    setIsParsing(true);
    setFileName(file.name);
    setRows([]);
    setParseError(null);

    try {
      const reader = new SheetReader();
      const result = await reader.parse(createBlobReader(file));

      const parsed: CellValue[][] = [];
      for await (const row of result) {
        parsed.push(row.cells);
        // Stream render: update every 500 rows for live feel
        if (parsed.length % 500 === 0) {
          setRows([...parsed]);
        }
      }

      setRows(parsed);
    } catch (e) {
      setParseError(e instanceof Error ? e.message : String(e));
    } finally {
      setIsParsing(false);
    }
  }, []);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (!file) return;
    if (/\.xls[xm]$/i.test(file.name)) parseFile(file);
    else { setFileName(file.name); setRows([]); setParseError("Only .xlsx and .xlsm files can be read here."); }
  }, [parseFile]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) parseFile(file);
  };

  const exportDemo = async () => {
    const writer = new SheetWriter();
    const data: Row[] = [
      [{ value: 'Product', style: { font: { bold: true, color: 'FFFFFFFF' }, fill: { type: 'solid' as const, fgColor: 'FF1e3a5f' } } },
       { value: 'Revenue', style: { font: { bold: true, color: 'FFFFFFFF' }, fill: { type: 'solid' as const, fgColor: 'FF1e3a5f' } } },
       { value: 'Units', style: { font: { bold: true, color: 'FFFFFFFF' }, fill: { type: 'solid' as const, fgColor: 'FF1e3a5f' } } }],
      ['Desk lamp', 15000, 3000],
      ['Office chair', 42000, 8400],
      ['Standing desk', 98000, 4900],
      [{ value: 'Total', style: { font: { bold: true } } },
       { value: null, formula: '=SUM(B2:B4)' },
       { value: null, formula: '=SUM(C2:C4)' }],
    ];
    
    writer.addSheet('Sales Data', data, {
      columnWidths: [30, 20, 15],
      conditionalFormats: [{
        range: 'B2:B4',
        rule: { type: 'dataBar', color: 'FF06b6d4' }
      }]
    });

    const todayExcel = 25569 + (Date.now() / 86400000);
    writer.addSheet('Metadata', [
      ['Property', 'Value'],
      ['Generated On', { value: todayExcel, style: { numFmt: 'yyyy-mm-dd hh:mm:ss' } }],
      ['Generator', 'XlsxFlow']
    ], { columnWidths: [20, 30] });

    const bytes = await new Response(writer.write()).arrayBuffer();
    const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'xlsxflow-demo.xlsx';
    a.click();
    URL.revokeObjectURL(url);
  };

  const columnCount = rows.reduce((n, row) => Math.max(n, row.length), 0);

  return (
    <div className="min-h-screen bg-slate-950 text-white font-sans">
      {/* Header */}
      <header className="border-b border-slate-800 bg-slate-950/80 backdrop-blur-md sticky top-0 z-50">
        <div className="max-w-7xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <img src="/icon.svg" alt="" className="w-9 h-9" />
            <span className="text-xl font-bold tracking-tight">XlsxFlow</span>
          </div>
          <nav className="flex items-center gap-4 md:gap-6 text-sm font-medium text-slate-400">
            <a href="#pro" className="hover:text-white transition-colors">Pro</a>
            <a href="#playground" className="hidden sm:inline hover:text-white transition-colors">Try it</a>
            <a href="https://github.com/xlsxflow/xlsxflow" className="hover:text-white transition-colors">GitHub</a>
            <button onClick={exportDemo}
              className="hidden sm:flex items-center gap-2 bg-cyan-500/10 hover:bg-cyan-500/20 text-cyan-400 border border-cyan-500/30 px-4 py-1.5 rounded-lg transition-all text-sm font-medium">
              Download a sample .xlsx
            </button>
          </nav>
        </div>
      </header>

      {/* Hero */}
      <section className="relative max-w-7xl mx-auto px-6 pt-20 pb-12 text-center overflow-hidden">
        <div className="absolute inset-0 bg-[radial-gradient(ellipse_at_center,_var(--tw-gradient-stops))] from-cyan-900/20 via-slate-950 to-slate-950 pointer-events-none"/>
        <div className="relative">
          <h1 className="text-5xl md:text-6xl font-black tracking-tighter mb-4 bg-gradient-to-br from-white via-slate-200 to-slate-400 bg-clip-text text-transparent">
            Read, write and edit xlsx<br/>as a stream.
          </h1>
          <p className="text-slate-400 text-lg max-w-2xl mx-auto">
            No dependencies. Runs in Node, Bun, browsers and Cloudflare Workers. Memory stays flat
            however many rows you stream.
          </p>
          <div className="mt-8 flex flex-wrap justify-center gap-3 text-sm">
            <code className="bg-slate-900 border border-slate-800 rounded-lg px-4 py-2 text-slate-300">npm install @xlsxflow/core</code>
            <a href="https://github.com/xlsxflow/xlsxflow#readme" className="rounded-lg px-4 py-2 border border-slate-700 text-slate-300 hover:text-white hover:border-slate-500 transition-colors">Documentation</a>
          </div>
        </div>
      </section>

      <main className="max-w-7xl mx-auto px-6 pb-20 grid grid-cols-1 lg:grid-cols-2 gap-8">

        {/* Pro Section */}
        <section id="pro" className="flex flex-col gap-5">
          <div className="flex items-center gap-3">
            <h2 className="text-xl font-bold">Free and Pro</h2>
          </div>

          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl flex flex-col gap-4 text-sm text-slate-300">
            <p>
              <span className="font-semibold text-white">@xlsxflow/core</span> is free and MIT licensed: read, write and edit
              .xlsx and .xlsm files with styles, formulas, images, tables and notes.
            </p>
            <p>
              <span className="font-semibold text-white">@xlsxflow/pro</span> adds two things: filling Excel templates with data,
              repeating rows for lists, and adding column, bar, line, area and pie charts.
            </p>
            <ul className="list-disc pl-5 text-slate-400 flex flex-col gap-1">
              <li>$5 per developer, with local pricing at checkout.</li>
              <li>Perpetual licence, with every version released within a year of your order.</li>
              <li>Build servers and CI don&apos;t need their own key.</li>
              <li>Full refund within 14 days if it doesn&apos;t work for you.</li>
            </ul>
            <div className="flex flex-wrap items-center gap-4">
              {CHECKOUT_URL && (
                <a href={CHECKOUT_URL} className="bg-gradient-to-r from-cyan-500 to-blue-600 hover:from-cyan-400 hover:to-blue-500 text-white font-semibold py-2 px-4 rounded-xl transition-colors">
                  Buy Pro
                </a>
              )}
              <a href="/pro-licence.txt" className="text-cyan-400 hover:text-cyan-300">Read the licence agreement</a>
            </div>
          </div>

          <h3 id="licence" className="text-lg font-bold mt-2">Get your licence key</h3>

          <div className="bg-slate-900 border border-slate-800 rounded-2xl p-6 shadow-xl flex flex-col gap-5">
            <div className="flex flex-col gap-2">
              <label htmlFor="order-id-input" className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Order ID</label>
              <input type="text" id="order-id-input" value={orderId} onChange={e => setOrderId(e.target.value)}
                className="bg-slate-950 border border-slate-700 rounded-xl px-4 py-3 text-white text-sm focus:outline-none focus:ring-2 focus:ring-cyan-500 placeholder:text-slate-600 transition-all"
                placeholder="The order ID on your Polar receipt" />
            </div>
            <div className="flex flex-col gap-2">
              <label htmlFor="order-email-input" className="text-xs font-semibold text-slate-400 uppercase tracking-wide">Email</label>
              <input type="email" id="order-email-input" autoComplete="email" value={email} onChange={e => setEmail(e.target.value)}
                className="bg-slate-950 border border-slate-700 rounded-xl px-4 py-3 text-white text-sm focus:outline-none focus:ring-2 focus:ring-cyan-500 placeholder:text-slate-600 transition-all"
                placeholder="The email you bought with" />
            </div>

            {licenseError && (
              <div className="text-red-400 text-xs bg-red-500/10 border border-red-500/30 rounded-lg px-4 py-3">
                {licenseError}
              </div>
            )}

            <button id="generate-license-btn" onClick={handleGenerateLicense} disabled={isGenerating || !orderId || !email}
              className="bg-gradient-to-r from-cyan-500 to-blue-600 hover:from-cyan-400 hover:to-blue-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-semibold py-3 px-4 rounded-xl shadow-lg shadow-cyan-500/20 transition-all transform hover:-translate-y-0.5 active:translate-y-0 flex items-center justify-center gap-2">
              {isGenerating ? <><span className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin"/>Checking your order…</> : 'Get licence key'}
            </button>

            {license && (
              <div className="flex flex-col gap-3 animate-in fade-in slide-in-from-bottom-2 duration-300">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-medium text-emerald-400">Your licence key</span>
                  <button onClick={downloadLicense} className="text-xs bg-slate-800 hover:bg-slate-700 px-3 py-1.5 rounded-lg text-slate-300 transition-colors font-medium">Download</button>
                </div>
                <pre className="bg-slate-950 border border-slate-800 rounded-xl p-4 text-xs text-emerald-400 font-mono whitespace-pre-wrap break-all max-h-48 scrollbar-thin">
                  {license}
                </pre>
                <p className="text-xs text-slate-500">Don&apos;t publish it, for example in a public repository. Activate it with <code className="text-slate-400">await setLicenseKey(key)</code> from <code className="text-slate-400">@xlsxflow/pro</code>, for example with the key in an <code className="text-slate-400">XLSXFLOW_LICENSE</code> environment variable.</p>
              </div>
            )}
          </div>
        </section>

        {/* Playground Section */}
        <section id="playground" className="flex flex-col gap-5">
          <div className="flex items-center gap-3">
            <h2 className="text-xl font-bold">Try the reader</h2>
          </div>

          {/* Drop Zone */}
          <div
            id="xlsx-drop-zone"
            onDragOver={e => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current?.click()}
            className={`relative rounded-2xl border-2 border-dashed transition-all cursor-pointer flex flex-col items-center justify-center p-10 gap-3 text-center
              ${isDragging
                ? 'border-cyan-400 bg-cyan-500/10 scale-[1.02]'
                : 'border-slate-700 bg-slate-900 hover:border-slate-500 hover:bg-slate-900/80'}`}
          >
            <input ref={fileInputRef} type="file" accept=".xlsx,.xlsm" className="hidden" onChange={handleFileChange} />
            <div className={`w-14 h-14 rounded-full flex items-center justify-center transition-all ${isDragging ? 'bg-cyan-500/20' : 'bg-slate-800'}`}>
              <svg className={`w-7 h-7 transition-colors ${isDragging ? 'text-cyan-400' : 'text-slate-400'}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.5" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
              </svg>
            </div>
            <div>
              <p className="font-semibold text-slate-200">Drop an .xlsx file here</p>
              <p className="text-slate-500 text-sm mt-1">Or click to browse. The file is read in your browser and never uploaded.</p>
            </div>
            {isParsing && (
              <div className="flex items-center gap-2 text-cyan-400 text-sm font-medium">
                <span className="w-4 h-4 border-2 border-cyan-400/30 border-t-cyan-400 rounded-full animate-spin"/>
                Streaming {fileName}…
              </div>
            )}
          </div>

          {parseError && (
            <div className="text-red-400 text-xs bg-red-500/10 border border-red-500/30 rounded-lg px-4 py-3">
              Could not read {fileName}: {parseError}
            </div>
          )}

          {/* Data Grid */}
          {rows.length > 0 && (
            <div className="flex flex-col gap-2 animate-in fade-in slide-in-from-bottom-2 duration-300">
              <div className="flex items-center justify-between">
                <span className="text-xs font-medium text-slate-400">{fileName} · <span className="text-emerald-400">{rows.length.toLocaleString()} rows</span></span>
                <span className="text-xs text-slate-600">{columnCount} columns</span>
              </div>
              <div className="overflow-auto rounded-xl border border-slate-800 max-h-80 shadow-inner">
                <table className="text-xs w-full border-collapse">
                  <thead className="sticky top-0 bg-slate-900 z-10">
                    <tr>
                      {Array.from({ length: columnCount }, (_, i) => (
                        <th key={i} className="px-3 py-2 text-left text-slate-400 font-semibold border-b border-slate-800 whitespace-nowrap">
                          {columnName(i)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.slice(0, 200).map((row, ri) => (
                      <tr key={ri} className="hover:bg-slate-800/50 transition-colors">
                        {row.map((cell, ci) => (
                          <td key={ci} className="px-3 py-1.5 border-b border-slate-800/50 text-slate-300 whitespace-nowrap max-w-[200px] overflow-hidden text-ellipsis">
                            {cell === null || cell === undefined ? <span className="text-slate-700">—</span> : String(cell)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
                {rows.length > 200 && (
                  <div className="text-center text-xs text-slate-600 py-3 border-t border-slate-800">
                    Showing 200 of {rows.length.toLocaleString()} rows
                  </div>
                )}
              </div>
            </div>
          )}
        </section>
      </main>

      {/* Footer */}
      <footer className="border-t border-slate-800 py-8 text-slate-500 text-xs">
        <div className="max-w-7xl mx-auto px-6 flex flex-wrap justify-center gap-x-6 gap-y-2">
          <a href="https://github.com/xlsxflow/xlsxflow" className="hover:text-slate-300">GitHub</a>
          <a href="https://www.npmjs.com/package/@xlsxflow/core" className="hover:text-slate-300">npm</a>
          <a href="/pro-licence.txt" className="hover:text-slate-300">Pro licence agreement</a>
          <a href="/privacy" className="hover:text-slate-300">Privacy</a>
          <a href={`mailto:${CONTACT}`} className="hover:text-slate-300">Contact</a>
        </div>
      </footer>
    </div>
  );
}
