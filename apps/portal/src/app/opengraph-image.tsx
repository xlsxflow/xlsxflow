import { ImageResponse } from "next/og";
import { readFile } from "node:fs/promises";

// The picture shown when the site is shared on LinkedIn, X, Slack and so on. Made at build time.
export const alt = "XlsxFlow: read, write and edit xlsx as a stream";
export const size = { width: 1200, height: 630 };
export const contentType = "image/png";
export const dynamic = "force-static";

const INK = "#1c1e21", MUTED = "#5b6168", GRID = "#dcdfe3", HEAD = "#f3f4f6", ACCENT = "#4338ca";
const COLS = ["A", "B", "C", "D", "E"];
const WIDTHS = [104, 150, 96, 76, 124];
const ROWS = [
  ["Order", "Date", "Region", "Units", "Total"],
  ["10001", "01-Jan-2026", "North", "24", "$588.00"],
  ["10002", "01-Jan-2026", "South", "32", "$6,048.00"],
  ["10003", "01-Jan-2026", "West", "32", "$2,556.80"],
  ["", "", "", "", ""],
  ["1010000", "31-Dec-2026", "East", "17", "$3,213.00"],
];
const ROW_NUMBERS = ["1", "2", "3", "4", "…", "1000001"];

const cell = { display: "flex", alignItems: "center", height: 44, padding: "0 12px", whiteSpace: "nowrap", borderRight: `1px solid ${GRID}`, borderBottom: `1px solid ${GRID}` };

export default async function Image() {
  const logo = `data:image/svg+xml;base64,${(await readFile(new URL("./icon.svg", import.meta.url))).toString("base64")}`;
  return new ImageResponse(
    (
      <div style={{ width: "100%", height: "100%", display: "flex", flexDirection: "column", background: "#ffffff", color: INK, padding: "56px 64px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 14, fontSize: 32 }}>
          <img src={logo} width={44} height={44} alt="" />
          XlsxFlow
        </div>
        <div style={{ display: "flex", fontSize: 60, letterSpacing: -2, marginTop: 28, lineHeight: 1.05 }}>
          Read, write and edit xlsx as a stream.
        </div>
        <div style={{ display: "flex", marginTop: 34, gap: 48, alignItems: "flex-start" }}>
          <div style={{ display: "flex", flexDirection: "column", border: `1px solid ${GRID}`, fontSize: 19 }}>
            <div style={{ display: "flex", background: HEAD, color: MUTED }}>
              <div style={{ ...cell, width: 96, justifyContent: "center" }} />
              {COLS.map((c, i) => <div key={c} style={{ ...cell, width: WIDTHS[i], justifyContent: "center" }}>{c}</div>)}
            </div>
            {ROWS.map((row, r) => (
              <div key={r} style={{ display: "flex" }}>
                <div style={{ ...cell, width: 96, justifyContent: "center", background: HEAD, color: MUTED, fontSize: 17 }}>{ROW_NUMBERS[r]}</div>
                {row.map((v, i) => (
                  <div key={i} style={{ ...cell, width: WIDTHS[i], justifyContent: r > 0 && (i === 0 || i >= 3) ? "flex-end" : "flex-start", color: v === "…" ? MUTED : INK }}>{v}</div>
                ))}
              </div>
            ))}
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 16, fontSize: 25, color: MUTED, marginTop: 4 }}>
            <div style={{ display: "flex", color: ACCENT, fontSize: 30 }}>1,000,000 rows,</div>
            <div style={{ display: "flex", color: ACCENT, fontSize: 30, marginTop: -12 }}>flat memory.</div>
            <div style={{ display: "flex" }}>Node, Bun, browsers, Workers</div>
            <div style={{ display: "flex" }}>Zero dependencies · MIT</div>
            <div style={{ display: "flex", marginTop: 10, padding: "12px 18px", border: `1px solid ${GRID}`, fontFamily: "monospace", fontSize: 22, color: INK }}>npm install @xlsxflow/core</div>
          </div>
        </div>
      </div>
    ),
    size,
  );
}
