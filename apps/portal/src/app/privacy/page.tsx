import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Privacy · XlsxFlow",
  description: "What the XlsxFlow website does with your data.",
};

export default function Privacy() {
  return (
    <main className="min-h-screen bg-slate-950 text-slate-300 font-sans">
      <div className="max-w-2xl mx-auto px-6 py-16 flex flex-col gap-5 text-sm leading-relaxed">
        <a href="/" className="text-cyan-400 hover:text-cyan-300">XlsxFlow</a>
        <h1 className="text-3xl font-bold text-white">Privacy</h1>
        <p className="text-slate-500">Last updated 8 October 2026</p>

        <h2 className="text-lg font-semibold text-white mt-4">Licence keys</h2>
        <p>
          When you ask for a licence key, the order ID and email you enter are sent to Polar, which sold you XlsxFlow Pro,
          to check that the order exists, is paid and was placed with that email. The key is generated from the order ID
          and shown to you. We don&apos;t store your email, your order ID or the key on this website.
        </p>

        <h2 className="text-lg font-semibold text-white mt-4">Payments</h2>
        <p>
          Polar Software, Inc. handles checkout, payment, taxes and invoices as merchant of record, under its own
          privacy policy. We receive your order details, including your name and email, from Polar.
        </p>

        <h2 className="text-lg font-semibold text-white mt-4">The file reader</h2>
        <p>Files you open in &quot;Try the reader&quot; are read in your browser. They are never uploaded.</p>

        <h2 className="text-lg font-semibold text-white mt-4">Logs and cookies</h2>
        <p>
          This website sets no cookies and uses no analytics. The hosting provider keeps standard request logs, such as
          IP addresses and times, for security and troubleshooting.
        </p>

        <h2 className="text-lg font-semibold text-white mt-4">The libraries</h2>
        <p>@xlsxflow/core and @xlsxflow/pro run in your own code. They don&apos;t contact us or send any data.</p>

        <h2 className="text-lg font-semibold text-white mt-4">Contact</h2>
        <p>
          Questions, or requests to see or delete data about you:{" "}
          <a href="mailto:palikaomkar.22.cse@anits.edu.in" className="text-cyan-400 hover:text-cyan-300">palikaomkar.22.cse@anits.edu.in</a>
        </p>
      </div>
    </main>
  );
}
