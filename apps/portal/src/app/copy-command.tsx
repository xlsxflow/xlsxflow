"use client";

import { useState } from "react";

export default function CopyCommand({ command }: { command: string }) {
  const [copied, setCopied] = useState(false);
  return (
    <div className="self-start flex items-stretch border border-grid font-mono text-sm max-w-full">
      <code className="px-3 py-2 overflow-x-auto whitespace-nowrap"><span className="text-muted select-none">$ </span>{command}</code>
      <button
        onClick={async () => { await navigator.clipboard.writeText(command); setCopied(true); setTimeout(() => setCopied(false), 1500); }}
        className="px-3 border-l border-grid font-sans text-muted hover:text-ink shrink-0"
      >
        {copied ? "Copied" : "Copy"}
      </button>
    </div>
  );
}
