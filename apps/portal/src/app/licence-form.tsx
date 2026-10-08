"use client";

import { useEffect, useState, useSyncExternalStore } from "react";

type Result = { success: boolean; license?: string; error?: string; pending?: boolean };

async function request(body: object): Promise<Result> {
  const res = await fetch("/api/licence", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`The licence service answered ${res.status}. Try again in a few minutes.`);
  return await res.json() as Result;
}

export default function LicenceForm() {
  const [orderId, setOrderId] = useState("");
  const [email, setEmail] = useState("");
  const [key, setKey] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [copied, setCopied] = useState(false);
  // Read on the client only: the page is prerendered without a query string
  const checkoutId = useSyncExternalStore(() => () => {}, () => new URLSearchParams(location.search).get("checkout_id"), () => null);
  const confirming = !!checkoutId && !key && !error;

  // Polar sends buyers back with ?checkout_id=…; the order shows up a few seconds after payment
  useEffect(() => {
    if (!checkoutId) return;
    let stopped = false;
    (async () => {
      try {
        for (let attempt = 0; attempt < 10 && !stopped; attempt++) {
          const result = await request({ checkoutId });
          if (stopped) return;
          if (result.success) {
            setKey(result.license!);
            history.replaceState(null, "", `${location.pathname}#licence`);
            return;
          }
          if (!result.pending) { setError(result.error ?? "The licence service gave no reason."); return; }
          await new Promise(r => setTimeout(r, 3000));
        }
        if (!stopped) setError("Your payment is taking longer than usual to confirm. Reload this page in a minute, or enter the order ID from your receipt below.");
      } catch (e) {
        if (!stopped) setError(e instanceof Error ? e.message : String(e));
      }
    })();
    return () => { stopped = true; };
  }, [checkoutId]);

  const submit = async (e: React.SubmitEvent<HTMLFormElement>) => {
    e.preventDefault();
    setError(null);
    setBusy(true);
    try {
      const result = await request({ orderId, email });
      if (result.success) setKey(result.license!);
      else setError(result.error ?? "The licence service gave no reason. Try again in a few minutes.");
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  const download = () => {
    if (!key) return;
    const url = URL.createObjectURL(new Blob([key], { type: "text/plain" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = "xlsxflow-licence.txt";
    a.click();
    URL.revokeObjectURL(url);
  };

  const copy = async () => {
    if (!key) return;
    await navigator.clipboard.writeText(key);
    setCopied(true);
  };

  const field = "w-full bg-paper border border-grid rounded-sm px-3 py-2 text-[15px] placeholder:text-muted/70 hover:border-muted focus:border-accent focus:outline-none focus:ring-1 focus:ring-accent";

  return (
    <div className="border border-grid p-5 sm:p-6">
      {confirming && <p role="status" className="mb-4 text-sm">Thanks for buying. Confirming your payment and making your key…</p>}
      <form onSubmit={submit} className="flex flex-col gap-4">
        <label className="flex flex-col gap-1.5">
          <span className="text-sm font-medium">Order ID</span>
          <input value={orderId} onChange={e => setOrderId(e.target.value)} required spellCheck={false}
            className={`${field} font-mono text-sm`} placeholder="e.g. 3f2a9c1e-…" />
        </label>
        <label className="flex flex-col gap-1.5">
          <span className="text-sm font-medium">Email you bought with</span>
          <input type="email" autoComplete="email" value={email} onChange={e => setEmail(e.target.value)} required className={field} />
        </label>
        {error && <p role="alert" className="text-sm text-error">{error}</p>}
        <button type="submit" disabled={busy || !orderId || !email}
          className="self-start px-4 py-2 rounded-sm bg-accent text-accent-ink font-medium hover:opacity-90 disabled:opacity-40 disabled:cursor-not-allowed">
          {busy ? "Checking your order…" : "Get licence key"}
        </button>
      </form>

      {key && (
        <div className="mt-6 pt-5 border-t border-grid flex flex-col gap-3">
          <div className="flex flex-wrap items-center justify-between gap-2">
            <p className="text-sm font-medium text-ok">Your licence key</p>
            <div className="flex gap-2 text-sm">
              <button onClick={copy} className="px-3 py-1 border border-grid hover:border-ink rounded-sm">{copied ? "Copied" : "Copy"}</button>
              <button onClick={download} className="px-3 py-1 border border-grid hover:border-ink rounded-sm">Download</button>
            </div>
          </div>
          <pre className="bg-head border border-grid p-3 text-xs font-mono whitespace-pre-wrap break-all max-h-40 overflow-auto">{key}</pre>
          <p className="text-sm text-muted">
            Keep it out of public repositories. Put it in an <code className="font-mono text-ink">XLSXFLOW_LICENSE</code> environment
            variable and call <code className="font-mono text-ink">await setLicenseKey(process.env.XLSXFLOW_LICENSE)</code> once at startup.
          </p>
        </div>
      )}
    </div>
  );
}
