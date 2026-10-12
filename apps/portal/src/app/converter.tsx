"use client";

import { useRef, useState } from "react";
import { SheetReader, SheetWriter, createBlobReader, parseCsv } from "@xlsxflow/core";
import type { Row, CellValue } from "@xlsxflow/core";

type Mode = "csv-to-xlsx" | "xlsx-to-csv";
interface Done { url: string; name: string; rows: number; bytes: number; ms: number; note?: string }

const MAX_ROWS = 1048576; // Excel's last row
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DELIMITERS: [string, string][] = [[",", "Comma"], [";", "Semicolon"], ["\t", "Tab"], ["|", "Pipe"]];
const SPEC = {
  "csv-to-xlsx": { accept: ".csv,.tsv,.txt", test: /\.(csv|tsv|txt)$/i, kind: "CSV" },
  "xlsx-to-csv": { accept: ".xlsx,.xlsm,.xls,.ods", test: /\.(xls[xm]?|ods)$/i, kind: "spreadsheet" },
};

const size = (n: number) => n < 1024 * 1024 ? `${Math.max(1, Math.round(n / 1024))} KB` : `${(n / 1024 / 1024).toFixed(1)} MB`;
const baseName = (f: File) => f.name.replace(/\.[^.]+$/, "");

// The delimiter used most on the first line, outside quotes
async function guessDelimiter(f: File): Promise<string> {
  const head = (await f.slice(0, 65536).text()).replace(/"[^"]*"/g, "");
  const line = head.split(/\r?\n/, 1)[0];
  let best = ",", most = 0;
  for (const [d] of DELIMITERS) {
    const n = line.split(d).length - 1;
    if (n > most) { best = d; most = n; }
  }
  return best;
}

async function csvToXlsx(f: File, delimiter: string, header: boolean, convert: boolean, progress: (n: number) => void): Promise<Omit<Done, "ms">> {
  const source = parseCsv(f.stream(), { delimiter, convert });
  // The first 100 rows set the column widths; the rest stream straight through
  // (next() rather than for-await with break, which would close the generator)
  const first: CellValue[][] = [];
  for (let r = await source.next(); !r.done; r = await source.next()) { first.push(r.value); if (first.length === 100) break; }
  const widths: number[] = [];
  for (const r of first) r.forEach((v, i) => { widths[i] = Math.min(60, Math.max(widths[i] ?? 8, String(v ?? "").length + 2)); });

  let rows = 0, cut = false;
  async function* all(): AsyncGenerator<CellValue[]> {
    yield* first;
    yield* source;
  }
  async function* out(): AsyncGenerator<Row> {
    for await (const r of all()) {
      if (rows === MAX_ROWS) { cut = true; break; }
      rows++;
      if (rows % 10000 === 0) progress(rows);
      yield rows === 1 && header ? r.map(value => ({ value, style: { font: { bold: true } } })) : r;
    }
  }
  const writer = new SheetWriter().addSheet(baseName(f).slice(0, 31) || "Sheet1", out(), { columnWidths: widths, freezePanes: header ? { row: 1 } : undefined });
  const blob = await new Response(writer.write()).blob();
  return {
    url: URL.createObjectURL(new Blob([blob], { type: XLSX_TYPE })),
    name: `${baseName(f)}.xlsx`, rows, bytes: blob.size,
    note: cut ? `Excel holds at most ${MAX_ROWS.toLocaleString()} rows a sheet, so the rows after that were left out.` : undefined,
  };
}

async function xlsxToCsv(f: File, sheetName: string | undefined, nameSheet: boolean, delimiter: string, progress: (n: number) => void): Promise<Omit<Done, "ms">> {
  const escape = (v: string) => v.includes(delimiter) || /["\r\n]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
  const text = (v: CellValue, shown: string | undefined) =>
    shown ?? (v == null ? "" : v instanceof Date ? v.toISOString().replace("T00:00:00.000Z", "") : typeof v === "boolean" ? String(v).toUpperCase() : String(v));
  // A byte order mark, so Excel opens the file as UTF-8
  const parts: string[] = ["﻿"];
  let chunk: string[] = [], rows = 0, last = 0;
  for await (const row of await new SheetReader().parse(createBlobReader(f), { sheetName, formatted: true })) {
    // Missing rows become blank lines, as in Excel's own CSV export
    for (let gap = last ? row.rowNumber - last - 1 : 0; gap > 0; gap--) chunk.push("");
    last = row.rowNumber;
    chunk.push(row.cells.map((v, i) => escape(text(v, row.formatted?.[i]))).join(delimiter));
    if (++rows % 10000 === 0) { parts.push(chunk.join("\r\n") + "\r\n"); chunk = []; progress(rows); }
  }
  parts.push(chunk.join("\r\n"));
  const blob = new Blob(parts, { type: "text/csv;charset=utf-8" });
  return { url: URL.createObjectURL(blob), name: `${baseName(f)}${nameSheet ? `-${sheetName}` : ""}.csv`, rows, bytes: blob.size };
}

interface Options { delimiter: string; header: boolean; convert: boolean; sheet?: string; sheets: string[] }

export default function Converter({ mode }: { mode: Mode }) {
  const spec = SPEC[mode];
  const [file, setFile] = useState<File | null>(null);
  const [opts, setOpts] = useState<Options>({ delimiter: ",", header: true, convert: true, sheets: [] });
  const [busy, setBusy] = useState<number | null>(null);
  const [done, setDone] = useState<Done | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const input = useRef<HTMLInputElement>(null);
  const run = useRef(0);
  const lastUrl = useRef<string | undefined>(undefined);
  const { delimiter, header, convert, sheet, sheets } = opts;

  const go = async (f: File, o: Options) => {
    const id = ++run.current;
    setFile(f); setOpts(o); setBusy(0); setDone(null); setError(null);
    if (lastUrl.current) URL.revokeObjectURL(lastUrl.current);
    lastUrl.current = undefined;
    const start = performance.now();
    const progress = (n: number) => { if (id === run.current) setBusy(n); };
    try {
      const r = await (mode === "csv-to-xlsx"
        ? csvToXlsx(f, o.delimiter, o.header, o.convert, progress)
        : xlsxToCsv(f, o.sheet, o.sheets.length > 1, o.delimiter, progress));
      if (id === run.current) { lastUrl.current = r.url; setDone({ ...r, ms: performance.now() - start }); }
      else URL.revokeObjectURL(r.url);
    } catch (e) {
      if (id === run.current) setError(e instanceof Error ? e.message : String(e));
    } finally {
      if (id === run.current) setBusy(null);
    }
  };
  // Options apply straight away to the file already chosen
  const change = (o: Partial<Options>) => file ? go(file, { ...opts, ...o }) : setOpts({ ...opts, ...o });

  const open = async (f: File | undefined) => {
    if (!f) return;
    const id = ++run.current;
    setDone(null);
    if (!spec.test.test(f.name)) { setFile(null); setError(`Choose a ${mode === "csv-to-xlsx" ? ".csv, .tsv or .txt" : ".xlsx, .xlsm, .xls or .ods"} file.`); return; }
    setError(null);
    try {
      if (mode === "csv-to-xlsx") {
        const d = f.name.toLowerCase().endsWith(".tsv") ? "\t" : await guessDelimiter(f);
        if (id === run.current) go(f, { ...opts, delimiter: d });
      } else {
        const info = await new SheetReader().readWorkbook(createBlobReader(f));
        const visible = info.sheets.filter(s => s.state === "visible").map(s => s.name);
        if (id === run.current) go(f, { ...opts, sheets: visible, sheet: visible[0] });
      }
    } catch (e) {
      setFile(null);
      setError(e instanceof Error ? e.message : String(e));
    }
  };

  const field = "border border-grid bg-paper rounded-sm px-2 py-1";

  return (
    <div
      className={`relative border bg-paper ${dragging ? "border-accent" : "border-grid"}`}
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={e => { if (!e.currentTarget.contains(e.relatedTarget as Node)) setDragging(false); }}
      onDrop={e => { e.preventDefault(); setDragging(false); open(e.dataTransfer.files[0]); }}
    >
      <div className="flex flex-wrap items-center gap-3 px-4 py-3 border-b border-grid">
        <button onClick={() => input.current?.click()} className="px-4 py-2 rounded-sm bg-accent text-accent-ink font-medium hover:opacity-90">
          Choose a {spec.kind} file
        </button>
        <input ref={input} type="file" accept={spec.accept} className="sr-only" tabIndex={-1} onChange={e => { open(e.target.files?.[0]); e.target.value = ""; }} />
        <span className="text-sm text-muted min-w-0 truncate">{file ? `${file.name}, ${size(file.size)}` : "or drop it here"}</span>
      </div>

      <div className="flex flex-wrap items-center gap-x-6 gap-y-3 px-4 py-3 border-b border-grid text-sm">
        {mode === "xlsx-to-csv" && sheets.length > 1 && (
          <label className="flex items-center gap-2">Sheet
            <select value={sheet} onChange={e => change({ sheet: e.target.value })} className={field}>
              {sheets.map(s => <option key={s}>{s}</option>)}
            </select>
          </label>
        )}
        <label className="flex items-center gap-2">{mode === "csv-to-xlsx" ? "Separated by" : "Separate with"}
          <select value={delimiter} onChange={e => change({ delimiter: e.target.value })} className={field}>
            {DELIMITERS.map(([d, name]) => <option key={name} value={d}>{name}</option>)}
          </select>
        </label>
        {mode === "csv-to-xlsx" && (
          <>
            <label className="flex items-center gap-2"><input type="checkbox" checked={header} onChange={e => change({ header: e.target.checked })} className="accent-accent" />First row is a header</label>
            <label className="flex items-center gap-2"><input type="checkbox" checked={convert} onChange={e => change({ convert: e.target.checked })} className="accent-accent" />Turn numbers into numbers</label>
          </>
        )}
      </div>

      <div className="px-4 py-5 min-h-24 flex flex-col gap-3" aria-live="polite">
        {error ? (
          <p className="text-error text-sm">{error}</p>
        ) : busy !== null ? (
          <p className="text-sm text-muted">Converting… {busy ? `${busy.toLocaleString()} rows so far` : ""}</p>
        ) : done ? (
          <>
            <div className="flex flex-wrap items-center gap-x-4 gap-y-2">
              <a href={done.url} download={done.name} className="px-4 py-2 rounded-sm border border-ink font-medium hover:bg-head">Download {done.name}</a>
              <span className="text-sm text-muted">{done.rows.toLocaleString()} rows, {size(done.bytes)}, in {Math.max(1, Math.round(done.ms)).toLocaleString()} ms</span>
            </div>
            {done.note && <p className="text-sm text-muted">{done.note}</p>}
          </>
        ) : (
          <p className="text-sm text-muted">The file is converted in your browser. It is never uploaded.</p>
        )}
      </div>

      {dragging && (
        <div className="absolute inset-0 grid place-items-center bg-paper/90 pointer-events-none">
          <p className="font-medium">Drop to convert it here. Nothing is uploaded.</p>
        </div>
      )}
    </div>
  );
}
