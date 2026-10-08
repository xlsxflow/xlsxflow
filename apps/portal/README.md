# XlsxFlow website

A static Next.js site (reader demo, free vs Pro, licence key form) served by a Cloudflare Worker. The Worker
(`worker/index.ts`) also answers `POST /api/licence`: it checks with Polar that the order is paid, is for the
Pro product and was placed with the email given, and returns a key signed with Ed25519.

## Run locally

```bash
pnpm dev                                  # the site only; the key form needs the Worker
pnpm build && npx wrangler@4 dev          # site and Worker on http://localhost:8787
```

`npx wrangler@4 dev` reads `.dev.vars` (git-ignored):

| Variable | Value |
|---|---|
| `POLAR_ACCESS_TOKEN` | Polar organisation access token with read access to orders |
| `POLAR_PRODUCT_ID` | ID of the Pro product |
| `POLAR_API_URL` | `https://sandbox-api.polar.sh/v1/orders` for a sandbox token. Production is the default |
| `ALLOW_SANDBOX` | `1` to accept sandbox orders. Sandbox orders are free, so never set it on the live site |
| `LICENSE_PRIVATE_KEY` | Base64 PKCS#8 DER Ed25519 key matching `PUBLIC_KEY` in the `@xlsxflow/pro` source |

`NEXT_PUBLIC_CHECKOUT_URL` (a Polar checkout link for Pro) is read by `pnpm build` from `.env.local` or the
environment, and shows the "Buy Pro" button. Without it the button is hidden.

## Deploy to Cloudflare

Once, from this folder:

```bash
npx wrangler@4 login
npx wrangler@4 secret put POLAR_ACCESS_TOKEN
npx wrangler@4 secret put POLAR_PRODUCT_ID
npx wrangler@4 secret put LICENSE_PRIVATE_KEY
```

Use the production Polar token and product. Each `secret put` prompts for the value, so it never appears in
the shell history. Then, for every release:

```bash
pnpm build && npx wrangler@4 deploy
```

The site is served at `https://xlsxflow.<your-subdomain>.workers.dev`. To use your own domain, add it under
the Worker's Settings → Domains & Routes in the Cloudflare dashboard. Security headers are in `public/_headers`.
