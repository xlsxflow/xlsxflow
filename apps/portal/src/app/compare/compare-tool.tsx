"use client";

import { useRef, useState } from "react";
import { SheetWriter } from "@xlsxflow/core";
import type { Lib, Result } from "./read.worker";

type Run = Result | { skipped: string } | "running" | "waiting";

const LIBS: [Lib, string][] = [["xlsxflow", "XlsxFlow"], ["exceljs", "ExcelJS 4.4.0"], ["sheetjs", "SheetJS 0.20.3"]];
const BIG = 50 * 1024 * 1024; // above this, the in-memory libraries can crash the tab
const ACCEPT = /\.(xlsx|xlsm|xls|ods)$/i;

const mb = (bytes: number) => `${(bytes / 1024 / 1024).toFixed(bytes < 10 * 1024 * 1024 ? 1 : 0)} MB`;

// 100,000 rows x 10 columns, written in the browser
async function makeSample(): Promise<File> {
  const rows = async function* () {
    yield ["ID", "Customer", "Region", "Product", "Units", "Price", "Total", "Date", "Paid", "Note"];
    for (let i = 1; i <= 100_000; i++) {
      yield [i, `Customer ${i % 997}`, ["North", "South", "East", "West"][i % 4], `Product ${i % 50}`, 1 + (i % 40),
        (i % 500) / 4, (1 + (i % 40)) * ((i % 500) / 4), new Date(Date.UTC(2026, 0, 1 + (i % 365))), i % 3 === 0, `Order ${i}`];
    }
  };
  const blob = await new Response(new SheetWriter().addSheet("Orders", rows()).write()).blob();
  return new File([blob], "sample-100k-rows.xlsx");
}

// Memory used by the worker that just finished. Needs a cross-origin isolated page and Chrome or Edge.
async function workerMemory(): Promise<number | null> {
  const measure = (performance as unknown as { measureUserAgentSpecificMemory?: () => Promise<{ breakdown: { bytes: number; attribution: { scope: string }[] }[] }> }).measureUserAgentSpecificMemory;
  if (!globalThis.crossOriginIsolated || !measure) return null;
  try {
    // Chrome may wait for its next garbage collection before answering; don't hold the table up for long
    const result = await Promise.race([measure.call(performance), new Promise<null>(r => setTimeout(r, 30_000, null))]);
    return result && result.breakdown.filter(b => b.attribution.some(a => a.scope === "DedicatedWorkerGlobalScope")).reduce((sum, b) => sum + b.bytes, 0);
  } catch {
    return null; // some embedded browsers have the function but refuse to run it
  }
}

function readWith(lib: Lib, file: File): Promise<{ result: Result; memory: number | null }> {
  const worker = new Worker(new URL("./read.worker.ts", import.meta.url), { type: "module" });
  return new Promise<Result>(resolve => {
    worker.onmessage = e => resolve(e.data);
    worker.onerror = e => resolve({ error: e.message || "The library stopped with an error." });
    worker.postMessage({ lib, file });
  })
    .then(async result => ({ result, memory: "error" in result ? null : await workerMemory() }))
    .finally(() => worker.terminate());
}

export default function CompareTool() {
  const [file, setFile] = useState<File | null>(null);
  const [runs, setRuns] = useState<Record<Lib, Run>>({ xlsxflow: "waiting", exceljs: "waiting", sheetjs: "waiting" });
  const [memory, setMemory] = useState<Partial<Record<Lib, number | null>>>({});
  const [busy, setBusy] = useState(false);
  const [all, setAll] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const input = useRef<HTMLInputElement>(null);

  const start = async (f: File, runAll: boolean) => {
    setFile(f); setError(null); setBusy(true); setMemory({});
    const next: Record<Lib, Run> = { xlsxflow: "waiting", exceljs: "waiting", sheetjs: "waiting" };
    setRuns({ ...next });
    for (const [lib] of LIBS) {
      if (lib === "exceljs" && !/\.xls[xm]$/i.test(f.name)) next[lib] = { skipped: "Doesn't read .xls or .ods files." };
      else if (lib !== "xlsxflow" && f.size > BIG && !runAll) next[lib] = { skipped: "Not run: reading a file this size whole can crash the tab. Tick the box above to try." };
      else {
        setRuns({ ...next, [lib]: "running" });
        const { result, memory: m } = await readWith(lib, f);
        next[lib] = result;
        setMemory(prev => ({ ...prev, [lib]: m }));
      }
      setRuns({ ...next });
    }
    setBusy(false);
  };

  const open = (f: File | undefined) => {
    if (!f || busy) return;
    if (!ACCEPT.test(f.name)) { setError("Choose an .xlsx, .xlsm, .xls or .ods file."); return; }
    start(f, all);
  };

  const ours = runs.xlsxflow;
  const base = typeof ours === "object" && "rows" in ours ? ours : null;
  const measured = Object.values(memory).some(m => m != null);

  return (
    <div className="flex flex-col gap-4">
      <div
        className={`relative border bg-paper ${dragging ? "border-accent" : "border-grid"}`}
        onDragOver={e => { e.preventDefault(); setDragging(true); }}
        onDragLeave={e => { if (!e.currentTarget.contains(e.relatedTarget as Node)) setDragging(false); }}
        onDrop={e => { e.preventDefault(); setDragging(false); open(e.dataTransfer.files[0]); }}
      >
        <div className="flex flex-wrap items-center gap-3 px-4 py-3 border-b border-grid">
          <button disabled={busy} onClick={() => input.current?.click()} className="px-4 py-2 rounded-sm bg-accent text-accent-ink font-medium hover:opacity-90 disabled:opacity-50">
            Choose a spreadsheet
          </button>
          <button disabled={busy} onClick={() => makeSample().then(f => start(f, all))} className="px-4 py-2 rounded-sm border border-grid hover:border-ink disabled:opacity-50">
            Use a 100,000-row sample
          </button>
          <input ref={input} type="file" accept=".xlsx,.xlsm,.xls,.ods" className="sr-only" tabIndex={-1} onChange={e => { open(e.target.files?.[0]); e.target.value = ""; }} />
          <span className="text-sm text-muted min-w-0 truncate">{file ? `${file.name}, ${mb(file.size)}` : "or drop a file here"}</span>
        </div>
        {file && file.size > BIG && (
          <label className="flex items-center gap-2 px-4 py-3 border-b border-grid text-sm">
            <input type="checkbox" checked={all} disabled={busy} onChange={e => setAll(e.target.checked)} className="accent-accent" />
            Also run ExcelJS and SheetJS on this {mb(file.size)} file (the tab may crash)
          </label>
        )}

        <div className="overflow-x-auto">
          <table className="w-full min-w-150 text-[15px]">
            <thead>
              <tr className="text-left text-sm text-muted border-b border-grid">
                <th className="py-2 px-4 font-medium">Library</th>
                <th className="py-2 px-3 font-medium text-right">Time</th>
                <th className="py-2 px-3 font-medium text-right">Memory when done</th>
                <th className="py-2 px-3 font-medium text-right">Rows</th>
                <th className="py-2 px-4 font-medium text-right">Cells</th>
              </tr>
            </thead>
            <tbody aria-live="polite">
              {LIBS.map(([lib, name]) => {
                const run = runs[lib];
                const m = memory[lib];
                return (
                  <tr key={lib} className="border-b border-grid last:border-b-0 align-top">
                    <th scope="row" className={`py-2.5 px-4 text-left ${lib === "xlsxflow" ? "font-semibold" : "font-normal"}`}>{name}</th>
                    {run === "waiting" || run === "running" ? (
                      <td colSpan={4} className="py-2.5 px-3 text-muted">{run === "running" ? "Reading…" : file ? "Waiting" : ""}</td>
                    ) : "skipped" in run || "error" in run ? (
                      <td colSpan={4} className={`py-2.5 px-3 ${"error" in run ? "text-error" : "text-muted"}`}>{"error" in run ? `Failed: ${run.error}` : run.skipped}</td>
                    ) : (
                      <>
                        <td className="py-2.5 px-3 text-right">{Math.max(1, Math.round(run.ms)).toLocaleString()} ms</td>
                        <td className="py-2.5 px-3 text-right">{m == null ? "–" : mb(m)}</td>
                        <td className={`py-2.5 px-3 text-right ${base && run.rows !== base.rows ? "text-error" : ""}`}>{run.rows.toLocaleString()}</td>
                        <td className={`py-2.5 px-4 text-right ${base && run.cells !== base.cells ? "text-error" : ""}`}>{run.cells.toLocaleString()}</td>
                      </>
                    )}
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {dragging && (
          <div className="absolute inset-0 grid place-items-center bg-paper/90 pointer-events-none">
            <p className="font-medium">Drop to compare. Nothing is uploaded.</p>
          </div>
        )}
      </div>
      {error && <p className="text-error text-sm">{error}</p>}
      {file && !busy && !measured && (
        <p className="text-sm text-muted">Memory is measured only in Chrome and Edge, which can report what a worker uses.</p>
      )}
    </div>
  );
}
