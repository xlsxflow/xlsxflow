// Cloudflare Worker: serves the static site from `out/` and issues licence keys at POST /api/licence.
// Secrets come from `wrangler secret put` in production and `.dev.vars` locally.

interface Env {
  ASSETS: { fetch(request: Request): Promise<Response> };
  POLAR_ACCESS_TOKEN?: string;
  POLAR_PRODUCT_ID?: string;
  POLAR_API_URL?: string;
  ALLOW_SANDBOX?: string;
  LICENSE_PRIVATE_KEY?: string;
}

const CONTACT = "palikaomkar@gmail.com";
const ORDER_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// A reason the buyer can act on; anything else is logged and shown as a generic error
class Refusal extends Error {}

const base64url = (bytes: Uint8Array) => btoa(String.fromCharCode(...bytes)).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");

// Keys are `<payload>.<signature>` (base64url), signed with Ed25519. @xlsxflow/pro verifies them
// offline against the public key in its license.ts.
async function createLicenseKey(orderId: string, privateKey: string): Promise<string> {
  const der = Uint8Array.from(atob(privateKey), c => c.charCodeAt(0));
  const key = await crypto.subtle.importKey("pkcs8", der, { name: "Ed25519" }, false, ["sign"]);
  const payload = base64url(new TextEncoder().encode(JSON.stringify({ orderId, tier: "pro", issuedAt: new Date().toISOString() })));
  const signature = new Uint8Array(await crypto.subtle.sign("Ed25519", key, new TextEncoder().encode(payload)));
  return `${payload}.${base64url(signature)}`;
}

type Order = { id?: unknown; checkout_id?: unknown; customer?: { email?: unknown }; product_id?: unknown; status?: unknown };
type Result = { success: true; license: string } | { success: false; error: string; pending?: true };

async function issueLicence(body: unknown, env: Env): Promise<Result> {
  try {
    const { orderId: rawId, email, checkoutId } = (body ?? {}) as { orderId?: unknown; email?: unknown; checkoutId?: unknown };
    // Polar IDs are UUIDs; checking the shape keeps the URLs below to exactly one order or checkout
    if (checkoutId === undefined && (typeof rawId !== "string" || !ORDER_ID.test(rawId.trim()))) {
      throw new Refusal("That doesn't look like a Polar order ID. Copy it from your receipt.");
    }
    if (checkoutId !== undefined && (typeof checkoutId !== "string" || !ORDER_ID.test(checkoutId))) {
      throw new Refusal("That checkout link is incomplete. Enter your order ID and email instead.");
    }
    if (checkoutId === undefined && (typeof email !== "string" || !email.includes("@"))) throw new Refusal("Enter the email you bought with.");

    const { POLAR_ACCESS_TOKEN: token, POLAR_PRODUCT_ID: productId, LICENSE_PRIVATE_KEY: privateKey } = env;
    if (!token || !productId || !privateKey) throw new Error("POLAR_ACCESS_TOKEN, POLAR_PRODUCT_ID or LICENSE_PRIVATE_KEY is not set.");
    const apiUrl = env.POLAR_API_URL || "https://api.polar.sh/v1/orders";
    // Sandbox orders are free, so only a test deployment may accept them
    if (apiUrl.includes("sandbox") && env.ALLOW_SANDBOX !== "1") {
      throw new Error("POLAR_API_URL points at the Polar sandbox. Set ALLOW_SANDBOX=1 only for a test deployment.");
    }

    const headers = { Authorization: `Bearer ${token}` };
    let order: Order;
    if (typeof checkoutId === "string") {
      // The checkout ID comes from Polar's redirect after payment and isn't on receipts, so it is
      // enough on its own. The order appears a few seconds after the redirect.
      const res = await fetch(`${apiUrl}/?checkout_id=${checkoutId.toLowerCase()}&limit=1`, { headers });
      if (!res.ok) {
        if (res.status === 422) throw new Refusal("That checkout link is incomplete. Enter your order ID and email instead.");
        throw new Error(`Polar API error: ${res.status} ${res.statusText}`);
      }
      // Checked again here: if Polar ever ignored the filter, any buyer's order would come back
      const found = (await res.json() as { items?: Order[] }).items?.find(o => String(o.checkout_id).toLowerCase() === checkoutId.toLowerCase());
      if (!found) return { success: false, pending: true, error: "Your payment is still being confirmed." };
      order = found;
    } else {
      const res = await fetch(`${apiUrl}/${(rawId as string).trim().toLowerCase()}`, { headers });
      if (!res.ok) {
        if (res.status === 404 || res.status === 422) throw new Refusal("Order not found or invalid.");
        throw new Error(`Polar API error: ${res.status} ${res.statusText}`);
      }
      order = await res.json() as Order;
      // The order ID is printed on receipts, so it alone must not be enough. Same message as an
      // unknown order, so the form can't be used to test which emails bought.
      const buyer = order.customer?.email;
      if (typeof buyer !== "string" || buyer.trim().toLowerCase() !== (email as string).trim().toLowerCase()) {
        throw new Refusal("Order not found or invalid.");
      }
    }
    if (typeof order.id !== "string" || !ORDER_ID.test(order.id)) throw new Error("Polar returned an order without an ID.");
    const orderId = order.id.toLowerCase();
    if (order.status === "pending") return { success: false, pending: true, error: "Your payment is still being confirmed." };
    // The product decides the tier, not the amount: discounts change the price
    if (order.product_id !== productId) throw new Refusal("This order is not for XlsxFlow Pro.");
    if (order.status !== "paid") {
      throw new Refusal(`This order is ${typeof order.status === "string" ? order.status.replace(/_/g, " ") : "not paid"}.`);
    }

    return { success: true, license: await createLicenseKey(orderId, privateKey) };
  } catch (error) {
    if (error instanceof Refusal) return { success: false, error: error.message };
    console.error("Licence request failed:", error);
    return { success: false, error: `Something went wrong on our side. Please try again later, or email ${CONTACT} with your order ID.` };
  }
}

const worker = {
  async fetch(request: Request, env: Env): Promise<Response> {
    const url = new URL(request.url);
    if (url.pathname !== "/api/licence") return env.ASSETS.fetch(request);
    if (request.method !== "POST") return new Response("Method not allowed", { status: 405, headers: { Allow: "POST" } });
    // JSON only: a cross-site form can't send it without a CORS preflight, which this endpoint never allows
    if (!request.headers.get("content-type")?.startsWith("application/json")) return new Response("Unsupported media type", { status: 415 });
    const body = await request.json().catch(() => null);
    return Response.json(await issueLicence(body, env), { headers: { "Cache-Control": "no-store", "X-Content-Type-Options": "nosniff" } });
  },
};

export default worker;
