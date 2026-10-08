"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { SheetReader, SheetWriter, createBlobReader } from "@xlsxflow/core";
import type { Row, CellStyle } from "@xlsxflow/core";

type Cell = string | number | boolean | null;
interface ShownRow { n: number; cells: Cell[]; text: (string | undefined)[]; formulas: (string | undefined)[] }

const SHOWN = 200; // rows kept for the grid; the rest are only counted
const MIN_COLS = 8;
const MIN_ROWS = 14;
const ACCEPT = /\.(xls[xm]?|ods)$/i;

// 0 -> A, 25 -> Z, 26 -> AA
const columnName = (i: number): string => (i >= 26 ? columnName(Math.floor(i / 26) - 1) : "") + String.fromCharCode(65 + (i % 26));

// 5,000 orders, made here in the browser with SheetWriter
async function makeSample(): Promise<File> {
  const head: CellStyle = { font: { bold: true } };
  const date: CellStyle = { numFmt: "dd-mmm-yyyy" };
  const money: CellStyle = { numFmt: "$#,##0.00" };
  const regions = ["North", "South", "East", "West"];
  const products: [string, number][] = [["Desk lamp", 24.5], ["Office chair", 189], ["Standing desk", 420], ["Monitor arm", 79.9], ["Notebook", 3.25]];
  let seed = 7;
  const next = () => (seed = (seed * 48271) % 2147483647) / 2147483647;
  const rows: Row[] = [["Order", "Date", "Region", "Product", "Units", "Unit price", "Total"].map(value => ({ value, style: head }))];
  for (let r = 2; r <= 5001; r++) {
    const [product, price] = products[Math.floor(next() * products.length)];
    const units = 1 + Math.floor(next() * 40);
    rows.push([
      10000 + r - 1,
      { value: new Date(Date.UTC(2026, 0, 1 + Math.floor((r - 2) / 18))), style: date },
      regions[Math.floor(next() * regions.length)],
      product,
      units,
      { value: price, style: money },
      { value: units * price, formula: `=E${r}*F${r}`, style: money },
    ]);
  }
  const writer = new SheetWriter({ properties: { title: "Orders", creator: "XlsxFlow" } });
  writer.addSheet("Orders", rows, { columnWidths: [10, 13, 10, 16, 8, 11, 12], freezePanes: { row: 1, col: 0 } });
  const blob = await new Response(writer.write()).blob();
  return new File([blob], "orders-sample.xlsx", { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
}

export default function SheetDemo() {
  const [file, setFile] = useState<File | null>(null);
  const [isSample, setIsSample] = useState(true);
  const [sheets, setSheets] = useState<string[]>([]);
  const [sheet, setSheet] = useState<string | null>(null);
  const [rows, setRows] = useState<ShownRow[]>([]);
  const [count, setCount] = useState(0);
  const [ms, setMs] = useState<number | null>(null);
  const [reading, setReading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [dragging, setDragging] = useState(false);
  const [sel, setSel] = useState<[number, number]>([0, 0]);
  const run = useRef(0);
  const input = useRef<HTMLInputElement>(null);

  const read = useCallback(async (f: File, sheetName?: string) => {
    const id = ++run.current;
    setFile(f);
    setError(null);
    setRows([]);
    setCount(0);
    setMs(null);
    setSel([0, 0]);
    setReading(true);
    try {
      const reader = createBlobReader(f);
      const start = performance.now();
      if (!sheetName) {
        const info = await new SheetReader().readWorkbook(reader);
        const visible = info.sheets.filter(s => s.state === "visible").map(s => s.name);
        if (id !== run.current) return;
        setSheets(visible);
        sheetName = visible[0];
      }
      setSheet(sheetName ?? null);
      const result = await new SheetReader().parse(reader, { sheetName, formulas: true, formatted: true });
      const shown: ShownRow[] = [];
      let n = 0;
      for await (const row of result) {
        if (id !== run.current) return;
        n++;
        if (shown.length < SHOWN) shown.push({ n: row.rowNumber, cells: row.cells, text: row.formatted ?? [], formulas: row.formulas ?? [] });
        if (n % 1000 === 0) { setCount(n); if (shown.length <= SHOWN) setRows([...shown]); }
      }
      setRows(shown);
      setCount(n);
      setMs(performance.now() - start);
    } catch (e) {
      if (id === run.current) setError(e instanceof Error ? e.message : String(e));
    } finally {
      if (id === run.current) setReading(false);
    }
  }, []);

  useEffect(() => {
    makeSample().then(f => read(f)).catch(e => setError(String(e)));
  }, [read]);

  const open = (f: File | undefined) => {
    if (!f) return;
    if (!ACCEPT.test(f.name)) {
      run.current++;
      setFile(f); setRows([]); setCount(0); setMs(null); setSheets([]); setReading(false);
      setError("This reader opens .xlsx, .xlsm, .xls and .ods files.");
      return;
    }
    setIsSample(false);
    read(f);
  };

  const download = () => {
    if (!file) return;
    const url = URL.createObjectURL(file);
    const a = document.createElement("a");
    a.href = url;
    a.download = file.name;
    a.click();
    URL.revokeObjectURL(url);
  };

  const cols = Math.max(MIN_COLS, rows.reduce((m, r) => Math.max(m, r.cells.length), 0));
  const lastRow = rows.length ? rows[rows.length - 1].n : 0;
  const filler = Array.from({ length: Math.max(0, MIN_ROWS - rows.length) }, (_, i) => lastRow + i + 1);
  const [sr, sc] = sel;
  const selRow = rows[sr];
  const selFormula = selRow?.formulas[sc];
  const selValue = selRow?.cells[sc];
  const bar = selFormula ? `=${selFormula}` : selValue == null ? "" : typeof selValue === "boolean" ? String(selValue).toUpperCase() : String(selValue);

  const onKey = (e: React.KeyboardEvent) => {
    const move: Record<string, [number, number]> = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    const d = move[e.key];
    if (!d || !rows.length) return;
    e.preventDefault();
    const r = Math.min(rows.length - 1, Math.max(0, sr + d[0]));
    const c = Math.min(cols - 1, Math.max(0, sc + d[1]));
    setSel([r, c]);
    document.getElementById(`cell-${r}-${c}`)?.scrollIntoView({ block: "nearest", inline: "nearest" });
  };

  return (
    <div
      className={`relative border bg-paper text-ink ${dragging ? "border-accent" : "border-grid"}`}
      onDragOver={e => { e.preventDefault(); setDragging(true); }}
      onDragLeave={e => { if (!e.currentTarget.contains(e.relatedTarget as Node)) setDragging(false); }}
      onDrop={e => { e.preventDefault(); setDragging(false); open(e.dataTransfer.files[0]); }}
    >
      {/* Title bar */}
      <div className="flex flex-wrap items-center gap-x-4 gap-y-2 px-3 py-2 border-b border-grid">
        <span className="font-medium text-sm truncate min-w-0 grow basis-full sm:basis-0">{file?.name ?? "Making a sample file"}</span>
        <div className="flex items-center gap-2 text-sm">
          {isSample && file && (
            <button onClick={download} className="px-3 py-1 border border-grid hover:border-ink rounded-sm">Download it</button>
          )}
          <button onClick={() => input.current?.click()} className="px-3 py-1 rounded-sm bg-accent text-accent-ink font-medium hover:opacity-90">Open your file</button>
          <input ref={input} type="file" accept=".xlsx,.xlsm,.xls,.ods" className="sr-only" tabIndex={-1} onChange={e => { open(e.target.files?.[0]); e.target.value = ""; }} />
        </div>
      </div>

      {/* Formula bar */}
      <div className="flex items-stretch border-b border-grid text-[13px] font-mono">
        <span className="w-[72px] shrink-0 px-2 py-1.5 border-r border-grid text-muted">{rows.length ? `${columnName(sc)}${selRow?.n ?? ""}` : ""}</span>
        <span className="px-2 py-1.5 border-r border-grid text-muted italic font-sans" aria-hidden>fx</span>
        <output className="px-2 py-1.5 truncate min-w-0 flex-1" aria-label="Selected cell">{bar}</output>
      </div>

      {/* Grid */}
      <div
        className="overflow-auto h-[360px] sm:h-[420px] outline-none focus-visible:outline-2 focus-visible:outline-accent focus-visible:-outline-offset-2"
        tabIndex={0}
        role="region"
        aria-label={`Cells of ${sheet ?? "the sheet"}. Use the arrow keys to move.`}
        onKeyDown={onKey}
      >
        {error ? (
          <div className="p-6 text-sm max-w-prose">
            <p className="text-error font-medium">Couldn&apos;t read {file?.name}.</p>
            <p className="mt-1 text-muted">{error}</p>
          </div>
        ) : (
          <table className="sheet">
            <thead>
              <tr>
                <th aria-label="Row" />
                {Array.from({ length: cols }, (_, c) => <th key={c} className={c === sc && rows.length ? "on" : ""}>{columnName(c)}</th>)}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, r) => (
                <tr key={row.n}>
                  <th className={r === sr ? "on" : ""}>{row.n}</th>
                  {Array.from({ length: cols }, (_, c) => {
                    const v = row.cells[c];
                    const align = typeof v === "number" ? "text-right" : typeof v === "boolean" ? "text-center" : "";
                    return (
                      <td key={c} id={`cell-${r}-${c}`} className={align} aria-selected={r === sr && c === sc} onClick={() => setSel([r, c])}>
                        {row.text[c] ?? (v == null ? "" : typeof v === "boolean" ? String(v).toUpperCase() : String(v))}
                      </td>
                    );
                  })}
                </tr>
              ))}
              {filler.map(n => (
                <tr key={`f${n}`}>
                  <th>{n}</th>
                  {Array.from({ length: cols }, (_, c) => <td key={c} />)}
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Sheet tabs and status */}
      <div className="flex flex-wrap items-center justify-between gap-x-4 gap-y-1 border-t border-grid bg-head text-[13px]">
        <div className="flex overflow-x-auto" role="tablist" aria-label="Sheets">
          {sheets.map(name => (
            <button key={name} role="tab" aria-selected={name === sheet}
              onClick={() => name !== sheet && file && read(file, name)}
              className={`px-3 py-1.5 border-r border-grid whitespace-nowrap ${name === sheet ? "bg-paper text-accent font-medium" : "text-muted hover:text-ink"}`}>
              {name}
            </button>
          ))}
        </div>
        <p className="px-3 py-1.5 text-muted" aria-live="polite">
          {reading
            ? `Reading… ${count.toLocaleString()} rows so far`
            : ms !== null
              ? `${count.toLocaleString()} rows read in ${Math.max(1, Math.round(ms)).toLocaleString()} ms${count > SHOWN ? `. The first ${SHOWN} are shown.` : "."}`
              : ""}
        </p>
      </div>

      {dragging && (
        <div className="absolute inset-0 grid place-items-center bg-paper/90 pointer-events-none">
          <p className="font-medium">Drop to read it here. Nothing is uploaded.</p>
        </div>
      )}
    </div>
  );
}
