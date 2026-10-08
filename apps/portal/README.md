# XlsxFlow portal

Next.js site with a reader demo and the Pro licence key page. A buyer enters their Polar order ID, the server
checks that the order is paid and is for the Pro product, and returns a key signed with Ed25519.

## Run

```bash
pnpm dev
```

`.env.local` needs:

| Variable | Value |
|---|---|
| `POLAR_ACCESS_TOKEN` | Polar organisation access token with read access to orders |
| `POLAR_PRODUCT_ID` | ID of the Pro product |
| `POLAR_API_URL` | Optional. `https://sandbox-api.polar.sh/v1/orders` for a sandbox token; production otherwise |
| `ALLOW_SANDBOX` | Optional. A production build refuses a sandbox `POLAR_API_URL`, since sandbox orders are free; set `1` only for a test deployment |
| `LICENSE_PRIVATE_KEY` | Base64 PKCS#8 DER Ed25519 key matching `PUBLIC_KEY` in the `@xlsxflow/pro` source |
| `NEXT_PUBLIC_CHECKOUT_URL` | Optional. A Polar checkout link for Pro, shown as the "Buy Pro" button. Set at build time; without it the button is hidden |

Set the same variables in the deployment, with a production token and product ID.
