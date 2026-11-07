// Everyday editing in a real browser: templates, layout change, add/duplicate/delete, notes, find & replace,
// command palette, JSON, history, and restoring the work after a reload.
// Usage: node qa/studio-editing.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
process.env.PLAYWRIGHT_DISABLE_FORCED_CHROMIUM_PROXIED_LOOPBACK = "1";
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "/opt/node22/lib/node_modules/playwright");
const root = fileURLToPath(new URL("..", import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).map((arg) => arg.replace(/^--/, "").split("=")));
const base = args.base || "http://127.0.0.1:8787";
const outDir = join(root, "qa", "out");
await mkdir(outDir, { recursive: true });
const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const shot = async (name) => { const file = join(outDir, `edit-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };
const films = () => page.locator(".film-item").count();
const stageType = () => page.getAttribute(".slide-wrap .hs-slide", "data-type");

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);

await step("a deck saved in a theme comes back in that theme", async () => {
  await page.evaluate(() => localStorage.setItem("hpptx-studio-current-v1", JSON.stringify({ selected: 1, deck: {
    title: "前の版の資料", theme: "kinari", accent: "#2d4b78", transition: "fade", motion: { backdrop: "waves" },
    slides: [{ type: "title", title: "前の版の資料", backdrop: "orbits" }, { type: "content", title: "本文", takeaway: "結論", points: ["A", "B"] }, { type: "closing" }],
  } })));
  await page.reload();
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length === 3, null, { timeout: 10000 });
  const themes = await page.$$eval(".slide-wrap .hs-slide, .film-item .hs-slide", (els) => [...new Set(els.map((el) => el.dataset.theme))]);
  if (themes.join() !== "kinari") throw new Error(`themes: ${themes}`);
  await shot("deck-in-its-theme");
  await page.evaluate(() => localStorage.clear());
  await page.goto(base);
});

await step("start from a template", async () => {
  if ((await page.locator("#createThemes .theme-chip").count()) !== 8) throw new Error("the eight themes should be offered");
  await page.click('#createThemes .theme-chip[data-theme-id="forest"]');
  if ((await page.getAttribute('#createThemes .theme-chip[data-theme-id="forest"]', "aria-checked")) !== "true") throw new Error("the chosen theme is not marked");
  await page.click("#blankDeckBtn");
  await page.waitForSelector("#templateDialog[open] .type-grid button");
  await page.waitForTimeout(600);
  await shot("templates");
  await page.click('#templateGrid button:has-text("企画提案")');
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length === 8);
  if ((await page.getAttribute(".slide-wrap .hs-slide", "data-theme")) !== "forest") throw new Error("the chosen theme was not used");
  const chip = await page.textContent("#issueSummary");
  if (!/構成の指摘/.test(chip)) throw new Error(`placeholders should be flagged: ${chip}`);
});

await step("change a layout (content is carried over)", async () => {
  await page.click(".film-item:nth-child(4)");
  await page.click("#formTab");
  await page.click('.inspector-head button:has-text("変更")');
  await page.waitForSelector("#typeDialog[open] #typeGrid button");
  await page.waitForTimeout(800);
  await shot("layouts");
  await page.click('#typeGrid button:has-text("ロードマップ")');
  await page.waitForTimeout(400);
  if ((await stageType()) !== "roadmap") throw new Error(`type is ${await stageType()}`);
  const text = await page.textContent(".slide-wrap .hs-slide");
  if (!text.includes("柱1")) throw new Error("the cards' text was lost");
});

await step("add, duplicate, delete and undo", async () => {
  const before = await films();
  await page.click('.inspector-foot button:has-text("後ろに追加")');
  await page.click('#typeGrid button:has-text("ひと言メッセージ")');
  await page.waitForTimeout(300);
  if ((await films()) !== before + 1 || (await stageType()) !== "statement") throw new Error("insert failed");
  await page.click('.inspector-foot button:has-text("複製")');
  await page.waitForTimeout(300);
  if ((await films()) !== before + 2) throw new Error("duplicate failed");
  await page.click('.inspector-foot button:has-text("削除")');
  await page.waitForTimeout(300);
  if ((await films()) !== before + 1) throw new Error("delete failed");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  if ((await films()) !== before + 2) throw new Error("undo failed");
});

await step("edit a card from the form and add a click-to-open detail", async () => {
  await page.click(".film-item:nth-child(4)");
  await page.waitForTimeout(200);
  const field = page.locator('#inspector [data-path="items[0].title"]');
  await field.fill("現場の声を集める");
  await page.waitForTimeout(400);
  const text = await page.textContent(".slide-wrap .hs-slide");
  if (!text.includes("現場の声を集める")) throw new Error("the slide did not follow the form");
  await page.click('#inspector button:has-text("詳細を追加")');
  await page.fill("#inspector .detail-card textarea", "店舗ヒアリング30件から抽出した課題です。");
  await page.waitForTimeout(400);
  const badge = await page.locator(".slide-wrap .hs-detail-badge").count();
  if (!badge) throw new Error("no detail badge on the slide");
  await page.click('#inspector .motion-grid button:has-text("クリックで")');
  await page.waitForTimeout(300);
  if ((await page.getAttribute(".slide-wrap .hs-slide", "data-build")) !== "click") throw new Error("build not set");
  await shot("details");
});

await step("speaker notes (quick) for every slide", async () => {
  await page.click(".more-menu summary");
  await page.click("#notesBtn");
  await page.click('label:has-text("簡易（すぐ作成）")');
  await page.click('label:has-text("すべて作り直す")');
  await page.click("#notesRun");
  await page.waitForTimeout(300);
  // The notes show under the slide (the status bar's ノート), as in PowerPoint.
  if ((await page.getAttribute("#notesToggle", "aria-pressed")) !== "true") await page.click("#notesToggle");
  await page.waitForSelector("#notesPane:not([hidden])");
  const notes = await page.innerText("#notesInput");
  if (!notes || notes.length < 20) throw new Error("no notes");
});

await step("find and replace", async () => {
  await page.click(".more-menu summary");
  await page.click("#replaceBtn");
  await page.fill("#findInput", "柱");
  await page.fill("#replaceInput", "施策");
  const count = await page.textContent("#findCount");
  if (!/か所/.test(count)) throw new Error(count);
  await page.click("#replaceAllBtn");
  await page.waitForTimeout(300);
});

await step("command palette jumps to a slide and offers no theme switch", async () => {
  await page.keyboard.press("Control+k");
  await page.fill("#commandInput", "テーマ");
  await page.waitForTimeout(200);
  if (await page.locator('#commandList :text("サンセット")').count()) throw new Error("theme commands should be gone");
  await page.keyboard.press("Escape");
  await page.keyboard.press("Control+k");
  await page.fill("#commandInput", "スケジュール");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  if ((await stageType()) !== "timeline") throw new Error(`landed on ${await stageType()}`);
  await shot("palette-jump");
});

await step("checks list the placeholders and move to them", async () => {
  await page.click("#issueSummary");
  await page.waitForSelector("#exportCheckDialog[open] .check-item");
  await shot("checks");
  await page.click("#exportCheckDialog .check-item >> nth=0 >> button:has-text(\"移動\")");
  await page.waitForTimeout(300);
});

await step("the work comes back after a reload", async () => {
  const count = await films();
  await page.waitForTimeout(600);
  await page.reload();
  await page.waitForFunction((n) => document.querySelectorAll(".film-item").length === n, count, { timeout: 10000 });
  if ((await page.getAttribute(".slide-wrap .hs-slide", "data-theme")) !== "forest") throw new Error("the theme was lost on reload");
});

await step("history keeps versions to go back to", async () => {
  await page.click("#historyBtn");
  await page.waitForSelector("#historyDialog[open] .history-item");
  await page.waitForTimeout(400);
  await shot("history");
  await page.keyboard.press("Escape");
});

await step("JSON round trip", async () => {
  await page.click("#importJsonBtn").catch(() => {});
  await page.click("#createTab");
  await page.click("#importJsonBtn");
  await page.waitForSelector("#jsonDialog[open]");
  const json = await page.inputValue("#slideDataInput");
  const data = JSON.parse(json);
  data.slideData[0].title = "JSONから開いた資料";
  await page.fill("#slideDataInput", JSON.stringify(data));
  await page.click("#applyJsonBtn");
  await page.waitForTimeout(500);
  const title = await page.textContent(".slide-wrap .hs-slide");
  if (!title.includes("JSONから開いた資料")) throw new Error("JSON not applied");
});

await writeFile(join(outDir, "edit-errors.txt"), errors.join("\n"));
console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
