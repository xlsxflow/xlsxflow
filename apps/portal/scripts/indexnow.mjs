// Tells Bing and the other IndexNow search engines that the site's pages changed. Run after a deploy:
//   node scripts/indexnow.mjs
// The key is public by design: search engines check it against /<key>.txt on the site.
import { readFileSync } from "node:fs";

const KEY = "4760382830a4567b4fd675f4a75e01b9";
const HOST = "www.xlsxflow.workers.dev";

const sitemap = readFileSync(new URL("../public/sitemap.xml", import.meta.url), "utf8");
const urlList = [...sitemap.matchAll(/<loc>([^<]+)<\/loc>/g)].map(m => m[1]);

const res = await fetch("https://api.indexnow.org/indexnow", {
  method: "POST",
  headers: { "Content-Type": "application/json; charset=utf-8" },
  body: JSON.stringify({ host: HOST, key: KEY, keyLocation: `https://${HOST}/${KEY}.txt`, urlList }),
});
// 200 and 202 both mean accepted; 403 means the key file isn't live yet
console.log(`IndexNow: ${res.status} ${res.statusText} for ${urlList.length} URLs`);
if (res.status !== 200 && res.status !== 202) process.exit(1);
