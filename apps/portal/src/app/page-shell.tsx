import Link from "next/link";

export const REPO = "https://github.com/xlsxflow/xlsxflow";
export const link = "underline decoration-grid underline-offset-4 hover:decoration-accent";

// Header and footer for the tool and guide pages
export default function PageShell({ children }: { children: React.ReactNode }) {
  return (
    <div className="flex-1 flex flex-col">
      <header className="border-b border-grid">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 h-14 flex items-center justify-between gap-4">
          <Link href="/" className="flex items-center gap-2 font-semibold text-[17px]">
            <img src="/icon.svg" alt="XlsxFlow logo" aria-hidden className="w-7 h-7" />
            XlsxFlow
          </Link>
          <nav className="flex items-center gap-4 sm:gap-6 text-sm text-muted">
            <Link href="/csv-to-xlsx" className="hover:text-ink">CSV to Excel</Link>
            <Link href="/xlsx-to-csv" className="hover:text-ink">Excel to CSV</Link>
            <a href={REPO} className="hidden sm:inline hover:text-ink">GitHub</a>
          </nav>
        </div>
      </header>
      <main className="flex-1">{children}</main>
      <footer className="border-t border-grid text-sm text-muted">
        <div className="max-w-6xl mx-auto px-4 sm:px-6 py-6 flex flex-wrap items-center justify-between gap-4">
          <Link href="/" className="hover:text-ink">XlsxFlow</Link>
          <nav className="flex flex-wrap gap-x-5 gap-y-2">
            <a href={REPO} className="hover:text-ink">GitHub</a>
            <a href="https://www.npmjs.com/package/@xlsxflow/core" className="hover:text-ink">npm</a>
            <Link href="/guides/javascript-heap-out-of-memory-excel" className="hover:text-ink">Out of memory with Excel files</Link>
            <a href="/privacy" className="hover:text-ink">Privacy</a>
          </nav>
        </div>
      </footer>
    </div>
  );
}
