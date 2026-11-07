// End-to-end smoke test of the studio in a real browser (needs the server running: `npm start`).
// Usage: node qa/studio-smoke.mjs [--base=http://127.0.0.1:8787]
// Opens the sample deck, edits, switches the theme (its theme colours follow), previews motion, presents and exports an HTML file.
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
process.env.PLAYWRIGHT_DISABLE_FORCED_CHROMIUM_PROXIED_LOOPBACK = "1";
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "/opt/node22/lib/node_modules/playwright");
const root = fileURLToPath(new URL("..", import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).map((arg) => arg.replace(/^--/, "").split("=")));
const base = args.base || "http://127.0.0.1:8787";
const outDir = join(root, "qa", "out");
await mkdir(outDir, { recursive: true });
const spki = process.env.PROXY_CA_SPKI || "";
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: spki ? [`--ignore-certificate-errors-spki-list=${spki}`] : [] });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const shot = async (name, target = page) => { const file = join(outDir, `studio-${name}.png`); await target.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };

await page.goto(base);
await page.evaluate(() => { localStorage.clear(); });
await page.goto(base);
await page.waitForTimeout(800);
await shot("create");

await step("open the sample deck", async () => {
  await page.click("#sampleDeckBtn");
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length >= 40, null, { timeout: 20000 });
  await page.waitForFunction(() => !/チェック中/.test(document.getElementById("issueSummary").textContent), null, { timeout: 60000 });
  await page.waitForTimeout(600);
  console.log("     issue chip:", await page.textContent("#issueSummary"));
  await shot("edit");
});

await step("select a KPI slide and open the form", async () => {
  await page.click(".film-item:nth-child(5)");
  await page.click("#formTab");
  await page.waitForTimeout(400);
  await shot("inspector");
});

await step("inline edit the title", async () => {
  await page.click(".slide-wrap .hs-title");
  await page.keyboard.type("インライン編集のテスト");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(500);
  const title = await page.textContent(".slide-wrap .hs-title");
  if (!title.includes("インライン編集のテスト")) throw new Error(`title not updated: ${title}`);
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  console.log("     after undo:", await page.textContent(".slide-wrap .hs-title"));
});

await step("add a YouTube video by URL", async () => {
  await page.click('button:has-text("URL・YouTube")');
  await page.fill("#mediaUrlInput", "https://www.youtube.com/watch?v=dQw4w9WgXcQ");
  await page.click("#mediaUrlApply");
  await page.waitForSelector(".slide-wrap .hs-placed", { timeout: 5000 });
  await page.waitForTimeout(600);
  await shot("media");
  await page.keyboard.press("Control+z");
});

await step("upload a video and play it while presenting", async () => {
  // A short clip recorded in the page itself (no video files needed on disk).
  const b64 = await page.evaluate(async () => {
    const canvas = document.createElement("canvas");
    canvas.width = 320; canvas.height = 180;
    const ctx = canvas.getContext("2d");
    const recorder = new MediaRecorder(canvas.captureStream(24), { mimeType: "video/webm" });
    const chunks = [];
    recorder.ondataavailable = (event) => chunks.push(event.data);
    recorder.start();
    let t = 0;
    const timer = setInterval(() => { ctx.fillStyle = `hsl(${t * 9}, 70%, 55%)`; ctx.fillRect(0, 0, 320, 180); t += 1; }, 40);
    await new Promise((resolve) => setTimeout(resolve, 1500));
    clearInterval(timer);
    const stopped = new Promise((resolve) => { recorder.onstop = resolve; });
    recorder.stop();
    await stopped;
    const bytes = new Uint8Array(await new Blob(chunks, { type: "video/webm" }).arrayBuffer());
    let text = "";
    for (const byte of bytes) text += String.fromCharCode(byte);
    return btoa(text);
  });
  const clip = join(outDir, "clip.webm");
  await writeFile(clip, Buffer.from(b64, "base64"));
  await page.setInputFiles('#inspector input[type=file][accept^="video"]', clip);
  await page.waitForSelector('.slide-wrap .hs-placed[data-kind="video"] video', { timeout: 8000 });
  await page.waitForTimeout(500);
  await shot("video");
  await page.click("#presentBtn");
  await page.waitForSelector(".hs-player .hs-placed video", { timeout: 8000 });
  await page.waitForTimeout(1200);
  const playing = await page.evaluate(() => { const video = document.querySelector(".hs-player .hs-placed video"); return Boolean(video && !video.paused && video.src.startsWith("blob:")); });
  if (!playing) throw new Error("the video does not play in the presentation");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  await page.keyboard.press("Control+z");
});

await step("grid and outline views", async () => {
  await page.click('label:has-text("一覧")');
  await page.waitForTimeout(800);
  await shot("grid");
  await page.click('label:has-text("構成")');
  await page.waitForTimeout(400);
  await shot("outline");
  await page.click('label:has-text("1枚")');
});

await step("design dialog: the deck's motion; デザイン → テーマ switches the theme and back (one undo step)", async () => {
  await page.click("#designBtn");
  await page.waitForTimeout(1500);
  await shot("design");
  if (!(await page.isVisible("#designDialog .template-note"))) throw new Error("the design note is missing");
  await page.click('#designDialog .dialog-foot [data-close]');
  await page.waitForTimeout(800);
  const before = await page.getAttribute(".slide-wrap .hs-slide", "data-theme");
  await page.click('.rb-tabs [role=tab]:has-text("デザイン")');
  await page.click('.rb-body .rb-btn:has-text("テーマ")');
  await page.waitForSelector(".rb-pop .rb-theme");
  if ((await page.locator(".rb-pop .rb-theme").count()) !== 8) throw new Error("the theme gallery should show the eight themes");
  await shot("theme-gallery");
  await page.click('.rb-pop .rb-theme[data-theme-id="midnight"]');
  await page.waitForFunction(() => document.querySelector(".slide-wrap .hs-slide")?.dataset.theme === "midnight", null, { timeout: 5000 });
  await shot("theme-midnight");
  await page.keyboard.press(process.platform === "darwin" ? "Meta+z" : "Control+z");
  await page.waitForFunction((t) => document.querySelector(".slide-wrap .hs-slide")?.dataset.theme === t, before, { timeout: 5000 });
  await page.click('.rb-tabs [role=tab]:has-text("ホーム")');
});

await step("motion preview", async () => {
  await page.locator('.film-item:has(.hs-slide[data-type="process"])').first().click();
  await page.click('button:has-text("▶ 動きを確認")');
  await page.waitForTimeout(2600);
  await shot("motion");
  await page.keyboard.press("Escape");
});

await step("present", async () => {
  await page.click(".film-item:nth-child(2)");
  await page.click("#presentBtn");
  await page.waitForSelector(".hs-player", { timeout: 10000 });
  await page.waitForTimeout(1600);
  await shot("present-1");
  for (let i = 0; i < 4; i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(700); }
  await shot("present-2");
  const popup = page.context().waitForEvent("page", { timeout: 5000 });
  await page.keyboard.press("p");
  const view = await popup;
  await view.waitForSelector(".pv-now .hs-slide", { timeout: 5000 });
  await view.setViewportSize({ width: 1180, height: 760 });
  await view.waitForTimeout(600);
  await shot("presenter-view", view);
  await view.close();
  await page.click(".hs-player-stage [data-detail]").catch(() => {});
  await page.waitForTimeout(500);
  await page.keyboard.press("Escape");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

let exported = null;
await step("export HTML", async () => {
  const download = page.waitForEvent("download", { timeout: 60000 });
  await page.click("#downloadBtn");
  await page.waitForTimeout(800);
  if (await page.isVisible("#exportCheckDialog[open]")) {
    await shot("check");
    await page.click("#exportCheckGoBtn");
  }
  const file = await download;
  exported = join(outDir, "export.html");
  await file.saveAs(exported);
  console.log("     exported", exported);
});

if (exported) {
  await step("play the exported file", async () => {
    const view = await context.newPage();
    view.on("pageerror", (error) => errors.push(`export pageerror: ${error.message}`));
    await view.goto(pathToFileURL(exported).href);
    await view.waitForSelector(".hs-player .hs-slide", { timeout: 20000 });
    await view.waitForTimeout(1800);
    await shot("export-1", view);
    for (let i = 0; i < 6; i += 1) { await view.keyboard.press("ArrowRight"); await view.waitForTimeout(500); }
    await view.waitForTimeout(900);
    await shot("export-2", view);
    await view.close();
  });
}

await writeFile(join(outDir, "studio-errors.txt"), errors.join("\n"));
console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
