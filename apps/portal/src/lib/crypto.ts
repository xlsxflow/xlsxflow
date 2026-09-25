import { createPrivateKey, sign } from "crypto";

// Licence keys are `<payload>.<signature>` (base64url), signed with Ed25519.
// @sheetforge/pro verifies them offline against the matching public key in its license.ts.
export function createLicenseKey(orderId: string): string {
  const privateKey = process.env.LICENSE_PRIVATE_KEY;
  if (!privateKey) throw new Error("LICENSE_PRIVATE_KEY is missing on the server.");
  const key = createPrivateKey({ key: Buffer.from(privateKey, "base64"), format: "der", type: "pkcs8" });

  const payload = Buffer.from(JSON.stringify({ orderId, tier: "pro", issuedAt: new Date().toISOString() })).toString("base64url");
  return `${payload}.${sign(null, Buffer.from(payload), key).toString("base64url")}`;
}
