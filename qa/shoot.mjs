// Visual QA: render slides of a deck in every requested theme and screenshot contact sheets.
// Usage: node qa/shoot.mjs [--themes=clarity,midnight] [--deck=/public/samples/ai-rollout.json] [--per=6] [--only=1,2,3] [--mode=thumb] [--strict]
import { createServer } from "node:http";
import { readFile, mkdir } from "node:fs/promises";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
process.env.PLAYWRIGHT_DISABLE_FORCED_CHROMIUM_PROXIED_LOOPBACK = "1";
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "/opt/node22/lib/node_modules/playwright");
const root = fileURLToPath(new URL("..", import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).map((arg) => arg.replace(/^--/, "").split("=")));
const types = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".json": "application/json", ".jpg": "image/jpeg", ".png": "image/png", ".svg": "image/svg+xml" };
const server = createServer(async (req, res) => {
  const path = normalize(decodeURIComponent(new URL(req.url, "http://x").pathname)).replace(/^([/\\])+/, "");
  try { const body = await readFile(join(root, path)); res.writeHead(200, { "content-type": types[extname(path)] || "application/octet-stream" }).end(body); }
  catch { res.writeHead(404).end("not found"); }
});
await new Promise((resolve) => server.listen(0, "127.0.0.1", resolve));
const port = server.address().port;
// The session's HTTPS proxy re-signs traffic with its own CA; trust exactly that key (see /root/.ccr/README.md).
const spki = process.env.PROXY_CA_SPKI || "";
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: spki ? [`--ignore-certificate-errors-spki-list=${spki}`] : [] });
const page = await browser.newPage({ viewport: { width: Math.max(1960, Number(args.w || 960) * Number(args.cols || 2) + 40), height: 1200 } });
const themes = (args.themes || "clarity").split(",");
const deck = args.deck || "/public/samples/ai-rollout.json";
const per = Number(args.per || 6);
const outDir = join(root, "qa", "out");
await mkdir(outDir, { recursive: true });
const total = JSON.parse(await readFile(join(root, deck), "utf8")).slideData.length;
const problems = [];
for (const theme of themes) {
  const batches = args.only ? [args.only] : Array.from({ length: Math.ceil(total / per) }, (_, b) => Array.from({ length: per }, (_, k) => b * per + k).filter((i) => i < total).join(","));
  for (const [b, only] of batches.entries()) {
    const url = `http://127.0.0.1:${port}/qa/gallery.html?deck=${encodeURIComponent(deck)}&theme=${theme}&only=${only}&mode=${args.mode || "thumb"}&cols=${args.cols || 2}&w=${args.w || 960}`;
    await page.goto(url);
    const done = await page.waitForFunction(() => window.__done, null, { timeout: 60000 }).then((handle) => handle.jsonValue());
    if (done.error) { console.error(theme, done.error); problems.push(`${theme} batch ${b + 1}: ${done.error}`); continue; }
    for (const result of done) if (result.issues.length || result.fs < 1 || result.ts < 1) problems.push(`${theme} #${result.index + 1} ${result.type}: fs=${result.fs} ts=${result.ts} ${result.issues.map((i) => `${i.field}:${i.message}`).join(" | ")}`);
    const file = join(outDir, `${theme}-${String(b + 1).padStart(2, "0")}.png`);
    await page.locator(".sheet").screenshot({ path: file });
    console.log("saved", file);
  }
}
console.log(problems.length ? `fit notes:\n${problems.join("\n")}` : "no fit notes");
await browser.close();
server.close();
// --strict (CI): any fit note or failed batch fails the run.
if ("strict" in args && problems.length) process.exitCode = 1;
