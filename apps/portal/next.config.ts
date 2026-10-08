import type { NextConfig } from "next";
import path from "path";

// A static site in `out/`, served by the Worker in worker/index.ts. Security headers are in
// public/_headers, since a static export has no server to add them.
const nextConfig: NextConfig = {
  output: "export",
  turbopack: {
    root: path.resolve(__dirname, "../.."),
  },
};

export default nextConfig;
