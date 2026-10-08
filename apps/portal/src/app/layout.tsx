import type { Metadata } from "next";
import { Archivo, IBM_Plex_Mono } from "next/font/google";
import "./globals.css";

const archivo = Archivo({
  variable: "--font-archivo",
  subsets: ["latin"],
  axes: ["wdth"],
});

const plexMono = IBM_Plex_Mono({
  variable: "--font-plex-mono",
  subsets: ["latin"],
  weight: ["400", "500"],
});

export const metadata: Metadata = {
  metadataBase: new URL("https://www.xlsxflow.workers.dev"),
  title: "XlsxFlow: streaming Excel (.xlsx) reader and writer for JavaScript",
  description: "Read, write and edit .xlsx files as a stream in Node, Bun, browsers and Cloudflare Workers, with flat memory at millions of cells. Zero dependencies, MIT licence. Also reads .xls and .ods.",
  alternates: { canonical: "/" },
  openGraph: {
    type: "website",
    siteName: "XlsxFlow",
    title: "XlsxFlow: streaming Excel (.xlsx) reader and writer for JavaScript",
    description: "Read, write and edit .xlsx files as a stream, with flat memory at millions of cells. Zero dependencies, MIT licence.",
    url: "/",
  },
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html lang="en" className={`${archivo.variable} ${plexMono.variable} h-full antialiased`}>
      <body className="min-h-full flex flex-col">{children}</body>
    </html>
  );
}
