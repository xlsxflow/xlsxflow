// Reads the first sheet of a file with one library and counts what it found. Each library runs in a
// fresh worker, so one run's memory never counts against the next.
import { SheetReader, createBlobReader } from "@xlsxflow/core";

export type Lib = "xlsxflow" | "exceljs" | "sheetjs";
export type Result = { ms: number; rows: number; cells: number } | { error: string };

const filled = (v: unknown) => v !== null && v !== undefined && v !== "";
// Kept on the worker's global until it ends, so a memory measurement sees what the library holds after reading
const keep = (workbook: unknown) => { (self as unknown as { held: unknown }).held = workbook; };

async function read(lib: Lib, file: File): Promise<Result> {
  let rows = 0, cells = 0;
  const count = (n: number) => { if (n) { rows++; cells += n; } };
  // Libraries load before the clock starts
  if (lib === "xlsxflow") {
    const start = performance.now();
    for await (const row of await new SheetReader().parse(createBlobReader(file))) count(row.cells.filter(filled).length);
    return { ms: performance.now() - start, rows, cells };
  }
  if (lib === "exceljs") {
    const ExcelJS = (await import("exceljs")).default;
    const start = performance.now();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(await file.arrayBuffer());
    workbook.worksheets[0]?.eachRow(row => { let n = 0; row.eachCell(c => { if (filled(c.value)) n++; }); count(n); });
    keep(workbook);
    return { ms: performance.now() - start, rows, cells };
  }
  const XLSX = await import("xlsx");
  const start = performance.now();
  // Dense mode and only the first sheet: SheetJS's lightest way to read
  const workbook = XLSX.read(await file.arrayBuffer(), { dense: true, sheets: 0 });
  const data = (workbook.Sheets[workbook.SheetNames[0]]?.["!data"] ?? []) as ({ t: string; v?: unknown } | undefined)[][];
  for (const row of data) count(row ? row.filter(c => c && c.t !== "z" && filled(c.v)).length : 0);
  keep(workbook);
  return { ms: performance.now() - start, rows, cells };
}

self.onmessage = async (e: MessageEvent<{ lib: Lib; file: File }>) => {
  let result: Result;
  try {
    result = await read(e.data.lib, e.data.file);
  } catch (err) {
    result = { error: err instanceof Error ? err.message : String(err) };
  }
  self.postMessage(result);
};
