// Browser: an "Export to Excel" button. Builds the file in the page and starts a download,
// with no server involved. Works with any bundler (Vite, webpack, Next.js client components).
import { SheetWriter } from '@xlsxflow/core';

export async function downloadXlsx(filename, header, records) {
  const rows = [
    header.map((h) => ({ value: h, style: { font: { bold: true }, fill: { type: 'solid', fgColor: 'FFE2E8F0' } } })),
    ...records,
  ];
  const blob = await new Response(new SheetWriter().addSheet('Sheet1', rows, { freezePanes: { row: 1 }, autoFitColumns: true }).write()).blob();
  const url = URL.createObjectURL(blob);
  const link = Object.assign(document.createElement('a'), { href: url, download: filename });
  link.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// downloadXlsx('people.xlsx', ['Name', 'Age'], [['Ana', 34], ['Bo', 27]]);
