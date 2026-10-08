import type { Metadata } from "next";
import Link from "next/link";

export const metadata: Metadata = {
  title: "Privacy · XlsxFlow",
  description: "What the XlsxFlow website does with your data.",
  alternates: { canonical: "/privacy" },
  openGraph: { title: "Privacy · XlsxFlow", url: "/privacy" },
};

export default function Privacy() {
  return (
    <main className="flex-1">
      <div className="max-w-prose mx-auto px-4 sm:px-6 py-12 sm:py-16 flex flex-col gap-4 leading-relaxed">
        <Link href="/" className="flex items-center gap-2 font-semibold"><img src="/icon.svg" alt="" className="w-6 h-6" />XlsxFlow</Link>
        <h1 className="display text-[48px] mt-6">Privacy</h1>
        <p className="text-muted text-sm">Last updated 8 October 2026</p>

        <h2 className="text-lg font-semibold mt-6">Licence keys</h2>
        <p>
          When you ask for a licence key, the order ID and email you enter are sent to Polar, which sold you XlsxFlow Pro,
          to check that the order exists, is paid and was placed with that email. The key is generated from the order ID
          and shown to you. We don&apos;t store your email, your order ID or the key on this website.
        </p>

        <h2 className="text-lg font-semibold mt-6">Payments</h2>
        <p>
          Polar Software, Inc. handles checkout, payment, taxes and invoices as merchant of record, under its own
          privacy policy. We receive your order details, including your name and email, from Polar.
        </p>

        <h2 className="text-lg font-semibold mt-6">The file reader</h2>
        <p>Files you open in the spreadsheet on the home page are read in your browser. They are never uploaded.</p>

        <h2 className="text-lg font-semibold mt-6">Logs and cookies</h2>
        <p>
          This website sets no cookies and uses no analytics. The hosting provider keeps standard request logs, such as
          IP addresses and times, for security and troubleshooting.
        </p>

        <h2 className="text-lg font-semibold mt-6">The libraries</h2>
        <p>@xlsxflow/core and @xlsxflow/pro run in your own code. They don&apos;t contact us or send any data.</p>

        <h2 className="text-lg font-semibold mt-6">Contact</h2>
        <p>
          Questions, or requests to see or delete data about you:{" "}
          <a href="mailto:palikaomkar@gmail.com" className="underline decoration-grid underline-offset-4 hover:decoration-accent">palikaomkar@gmail.com</a>
        </p>
      </div>
    </main>
  );
}
