// How closely does "見た目どおりに取り込む" look like the original PowerPoint deck? Each deck is drawn twice - the original by
// LibreOffice (tools/fidelity.py render), the imported deck by the studio engine (qa/gallery.html) - and every slide is scored:
//   ssim    structural similarity of the pictures        text    share of the original's characters still on the slide
//   blocks  coarse light/dark layout                      pictures  how many pictures the original shows vs the studio
// Side-by-side pictures (original | studio | difference) are written for every slide so the worst ones can be looked at.
// Usage: node qa/fidelity.mjs [--corpus=DIR] [--only=name,name] [--out=DIR] [--port=8796]
//                             [--baseline=FILE] [--update-baseline] [--strict]
//   no --corpus      the decks test/pptx_fixtures.py builds (a company template and a plain 4:3 deck) - these are what CI guards
//   --corpus=DIR     every .pptx in DIR (a folder of real decks stays on your machine; nothing in it is committed)
//   --baseline       a JSON of {deck: {ssimMean, textMean}}; with --strict a deck that falls below it fails the run
import { execFileSync, spawn } from "node:child_process";
import { createServer } from "node:http";
import { mkdir, readFile, readdir, rm, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { basename, extname, join, normalize, relative } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
process.env.PLAYWRIGHT_DISABLE_FORCED_CHROMIUM_PROXIED_LOOPBACK = "1";
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "/opt/node22/lib/node_modules/playwright");
const root = fileURLToPath(new URL("..", import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).map((arg) => arg.replace(/^--/, "").split("=")));
const python = process.env.PYTHON_BIN || "python3";
const outRoot = args.out ? join(process.cwd(), args.out) : join(root, "qa", "out", "fidelity");
const baselineFile = args.baseline ? join(process.cwd(), args.baseline) : join(root, "qa", "fidelity-baseline.json");

// ---- the decks to measure
let decks = [];
if (args.corpus) {
  const dir = join(process.cwd(), args.corpus);
  decks = (await readdir(dir)).filter((f) => /\.pptx$/i.test(f)).sort().map((f) => join(dir, f));
} else {
  const fixtures = join(outRoot, "fixtures");
  await mkdir(fixtures, { recursive: true });
  execFileSync(python, [join(root, "test", "pptx_fixtures.py"), fixtures], { stdio: "ignore" });
  decks = (await readdir(fixtures)).filter((f) => /\.pptx$/i.test(f)).sort().map((f) => join(fixtures, f));
}
if (args.only) {
  const wanted = String(args.only).split(",");
  decks = decks.filter((d) => wanted.includes(basename(d, ".pptx")) || wanted.includes(basename(d)));
}
if (!decks.length) { console.error("no decks to measure"); process.exit(2); }

// ---- the studio server (the import API) and a static server for the gallery page
const port = Number(args.port || 8796);
const studio = spawn(process.execPath, ["server.js"], { cwd: root, env: { ...process.env, PORT: String(port), HOST: "127.0.0.1" }, stdio: "ignore" });
const base = `http://127.0.0.1:${port}`;
let up = false;
for (let i = 0; i < 80 && !up; i++) {
  try { up = (await fetch(`${base}/healthz`)).ok; } catch { /* not yet */ }
  if (!up) await new Promise((r) => setTimeout(r, 250));
}
if (!up) { studio.kill(); console.error("the studio server did not start"); process.exit(2); }
const types = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".mjs": "text/javascript", ".css": "text/css", ".json": "application/json", ".png": "image/png", ".woff2": "font/woff2" };
const files = createServer(async (req, res) => {
  const path = normalize(decodeURIComponent(new URL(req.url, "http://x").pathname)).replace(/^([/\\])+/, "");
  try {
    // /__out/... is where this run keeps the imported decks (it may be outside the repository); everything else is the repository.
    const body = await readFile(path.startsWith("__out/") ? join(outRoot, path.slice(6)) : join(root, path));
    res.writeHead(200, { "content-type": types[extname(path)] || "application/octet-stream" }).end(body);
  } catch { res.writeHead(404).end("not found"); }
});
await new Promise((resolve) => files.listen(0, "127.0.0.1", resolve));
const filesBase = `http://127.0.0.1:${files.address().port}`;

const spki = process.env.PROXY_CA_SPKI || "";
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: spki ? [`--ignore-certificate-errors-spki-list=${spki}`] : [] });
const page = await browser.newPage({ viewport: { width: 1000, height: 700 } });

async function measure(deckFile) {
  const name = basename(deckFile, ".pptx");
  const dir = join(outRoot, name);
  await rm(dir, { recursive: true, force: true });
  const [origDir, studioDir, compareDir] = ["orig", "studio", "compare"].map((d) => join(dir, d));
  for (const d of [origDir, studioDir, compareDir]) await mkdir(d, { recursive: true });

  execFileSync(python, [join(root, "tools", "fidelity.py"), "render", deckFile, origDir], { stdio: ["ignore", "ignore", "inherit"] });

  const response = await fetch(`${base}/api/import?mode=exact&name=${encodeURIComponent(basename(deckFile))}`, { method: "POST", body: await readFile(deckFile) });
  const imported = await response.json();
  if (!response.ok) throw new Error(`import: ${imported.error}`);
  await writeFile(join(dir, "deck.json"), JSON.stringify(imported));
  const slides = imported.slideData;
  const visible = slides.map((s, i) => (s.hidden ? -1 : i)).filter((i) => i >= 0);

  const url = `${filesBase}/qa/gallery.html?deck=${encodeURIComponent(`/__out/${name}/deck.json`)}&theme=clarity&mode=thumb&cols=1&w=960&from=0&count=100000&repaint=0`;
  await page.goto(url);
  const done = await page.waitForFunction(() => window.__done, null, { timeout: 120000 }).then((h) => h.jsonValue());
  if (done.error) throw new Error(`gallery: ${done.error}`);
  await page.addStyleTag({ content: ".cell > b { display: none !important; }" });
  const cells = await page.$$(".cell");
  const studioSlides = [];
  for (const [n, i] of visible.entries()) {
    const cell = cells[i];
    // A clip of the page, not an element screenshot: that waits for the element to hold still and gives up when it never does.
    const box = await cell.boundingBox();
    const file = `slide-${String(n + 1).padStart(3, "0")}.png`;
    await page.screenshot({ path: join(studioDir, file), clip: { x: box.x + (await page.evaluate(() => scrollX)), y: box.y + (await page.evaluate(() => scrollY)), width: box.width, height: box.height }, fullPage: true });
    const info = await cell.evaluate((el) => ({
      text: el.querySelector(".hs-slide")?.innerText ?? "",
      pictures: el.querySelectorAll(".hs-slide .hs-obj[data-kind=image]").length || el.querySelectorAll(".hs-slide img").length,
    }));
    studioSlides.push({ file, ...info });
  }
  await writeFile(join(studioDir, "studio.json"), JSON.stringify({ slides: studioSlides }));

  const summary = JSON.parse(execFileSync(python, [join(root, "tools", "fidelity.py"), "compare", origDir, studioDir, compareDir], { encoding: "utf-8" }));
  const report = JSON.parse(await readFile(join(compareDir, "report.json"), "utf-8"));
  return { name, summary, rows: report.rows, stats: imported.stats, compareDir };
}

const results = [];
for (const deck of decks) {
  process.stdout.write(`${basename(deck)} ... `);
  try {
    const r = await measure(deck);
    results.push(r);
    console.log(`ssim ${r.summary.ssimMean} (min ${r.summary.ssimMin})  text ${r.summary.textMean}  ${r.summary.slides} slides`);
  } catch (error) {
    results.push({ name: basename(deck, ".pptx"), error: error.message });
    console.log(`FAILED: ${error.message}`);
  }
}
await browser.close();
files.close();
studio.kill();

// ---- the report
const lines = ["| deck | slides | ssim mean | ssim min | blocks | text mean | text min | pictures (orig/studio) | unsupported | skipped |", "|---|---|---|---|---|---|---|---|---|---|"];
for (const r of results) {
  if (r.error) { lines.push(`| ${r.name} | - | - | - | - | - | - | - | - | ${r.error} |`); continue; }
  const s = r.summary;
  lines.push(`| ${r.name} | ${s.slides}/${s.originalPages} | ${s.ssimMean} | ${s.ssimMin} | ${s.blocksMean} | ${s.textMean} | ${s.textMin} | ${s.pictures}/${s.studioPictures} | ${r.stats?.unsupported ?? "-"} | ${r.stats?.skipped ?? "-"} |`);
}
lines.push("", "Worst slides (lowest ssim):");
for (const r of results.filter((x) => !x.error)) {
  const worst = [...r.rows].sort((a, b) => a.ssim - b.ssim).slice(0, 3);
  for (const w of worst) lines.push(`- ${r.name} #${w.slide}: ssim ${w.ssim} blocks ${w.blocks} text ${w.text}${w.missing.length ? ` missing: ${w.missing.join(" / ")}` : ""}  → ${relative(root, join(r.compareDir, w.pair))}`);
}
await mkdir(outRoot, { recursive: true });
await writeFile(join(outRoot, "report.md"), lines.join("\n") + "\n");
await writeFile(join(outRoot, "report.json"), JSON.stringify(results.map(({ name, summary, error, stats }) => ({ name, summary, error, stats })), null, 1));
console.log("\n" + lines.join("\n"));

// A deck may not look worse than its baseline by more than this (another machine draws text a little differently).
const SSIM_SLACK = 0.05;
const TEXT_SLACK = 0.01;
// Whatever the baseline says, a slide this far off is broken: its layout is somewhere else, or words are gone.
const BLOCKS_FLOOR = 0.93;
const TEXT_FLOOR = 0.95;
const broken = results.filter((r) => !r.error).flatMap((r) => r.rows.filter((w) => w.blocks < BLOCKS_FLOOR || w.text < TEXT_FLOOR).map((w) => ({ deck: r.name, ...w })));
if (broken.length) {
  console.log(`\nSlides far from the original (blocks < ${BLOCKS_FLOOR} or text < ${TEXT_FLOOR}):`);
  for (const w of broken) console.log(`- ${w.deck} #${w.slide}: blocks ${w.blocks} text ${w.text}${w.missing.length ? ` missing: ${w.missing.join(" / ")}` : ""}`);
}

// ---- the baseline: a deck must not look worse than it did
const current = Object.fromEntries(results.filter((r) => !r.error).map((r) => [r.name, { ssimMean: r.summary.ssimMean, textMean: r.summary.textMean }]));
if ("update-baseline" in args) {
  await writeFile(baselineFile, JSON.stringify(current, null, 2) + "\n");
  console.log(`\nbaseline written: ${relative(process.cwd(), baselineFile)}`);
}
let failed = results.some((r) => r.error) || ("strict" in args && broken.length > 0);
if ("strict" in args && existsSync(baselineFile)) {
  const baseline = JSON.parse(await readFile(baselineFile, "utf-8"));
  for (const [name, now] of Object.entries(current)) {
    const was = baseline[name];
    if (!was) continue;
    if (now.ssimMean < was.ssimMean - SSIM_SLACK) { console.log(`FAIL ${name}: ssim ${now.ssimMean} < baseline ${was.ssimMean} - ${SSIM_SLACK}`); failed = true; }
    if (now.textMean < was.textMean - TEXT_SLACK) { console.log(`FAIL ${name}: text ${now.textMean} < baseline ${was.textMean} - ${TEXT_SLACK}`); failed = true; }
  }
}
process.exit(failed ? 1 : 0);
