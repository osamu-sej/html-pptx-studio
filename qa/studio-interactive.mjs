// Pages the audience works with, in a real browser: the "動く資料" template, evidence panels, a ranking that
// re-sorts by view, bars that move from before to after, a simulator's sliders, measures that fill a gap, the
// "いま" line, the overview (O), the automatic demo (D), and the exported file with and without #static.
// Usage: node qa/studio-interactive.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir } from "node:fs/promises";
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
const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|fonts/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
// The slide as a picture of its place on the page. An element screenshot waits for the element to hold still and gives up
// when the stage re-renders or keeps moving; clipping the page does not depend on that.
const stageShot = async (name) => { const box = await page.locator(".slide-wrap").boundingBox(); await page.screenshot({ path: join(outDir, `interactive-${name}.png`), clip: box }); };
const shot = async (name, target = page) => { const file = join(outDir, `interactive-${name}.png`); await target.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };
const expect = (value, wanted, what) => { if (JSON.stringify(value) !== JSON.stringify(wanted)) throw new Error(`${what}: ${JSON.stringify(value)}`); };
const P = "#presenter .hs-player";
const now = (sel) => page.$eval(`${P} .hs-player-slide:last-of-type ${sel}`, (el) => el.textContent.trim());
const counter = () => page.textContent(`${P} .hs-player-count`).then((text) => text.trim());
const film = (n) => page.click(`.film-item:nth-child(${n})`);

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);

await step("the template: a worked deck with one action per page", async () => {
  await page.click("#blankDeckBtn");
  await page.waitForSelector("#templateDialog[open] .type-grid button");
  await page.click('#templateGrid button:has-text("動く資料")');
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length === 8);
  const types = await page.$$eval(".film-item .hs-slide", (els) => els.map((el) => el.dataset.type));
  expect(types, ["title", "headerCards", "imageText", "imageText", "simulator", "gap", "gantt", "closing"], "slides");
  const chip = await page.textContent("#issueSummary");
  if (/構成の指摘/.test(chip)) throw new Error(`the worked example should be clean: ${chip}`);
});

await step("editing: simulator and gap forms, the formula recalculates, evidence fields", async () => {
  await page.click("#formTab");
  await film(5);
  await page.waitForSelector("#f-formula");
  expect(await page.textContent(".slide-wrap .hs-sim-out"), "3,600", "the result from the starting values");
  await page.fill("#f-formula", "a × b × c ÷ 50");
  await page.waitForTimeout(500);
  expect(await page.textContent(".slide-wrap .hs-sim-out"), "7,200", "the result after changing the formula");
  await stageShot("edit-simulator");
  await page.fill("#f-formula", "a × b × c ÷ 100");
  await film(6);
  await page.waitForSelector('#inspector [data-path="measures[0].title"], #inspector input[value="業務別テンプレート"]');
  expect(await page.textContent(".slide-wrap .hs-gap-num"), "+100", "a finished gap slide shows every measure on");
  await stageShot("edit-gap");
  await film(2);
  await page.waitForSelector("#inspector .detail-evidence[open] textarea");
  const rows = await page.locator("#inspector .detail-card").first().locator(".detail-evidence textarea").first().inputValue();
  if (!/資料作成,620時間/.test(rows)) throw new Error(`breakdown rows: ${rows}`);
  await page.locator("#inspector .detail-card").first().scrollIntoViewIfNeeded();
  await shot("inspector-evidence");
  await film(3);
  await page.waitForSelector('#inspector select option[value="rank"]:checked', { state: "attached" });
  await page.locator("#inspector textarea[data-path=image]").scrollIntoViewIfNeeded();
  await shot("inspector-rank");
});

await step("present: a ring shows what can be clicked; the key message opens the evidence panel", async () => {
  await film(2);
  await page.click("#presentBtn");
  await page.waitForSelector(`${P} .hs-slide[data-type="headerCards"]`);
  await page.waitForSelector(`${P} .hs-ring`, { timeout: 4000 });
  await page.waitForTimeout(1600);
  await page.click(`${P} .hs-takeaway-row`);
  await page.waitForSelector(`${P} .hs-evidence`);
  await page.waitForTimeout(700);
  await shot("evidence-rows");
  expect(await counter(), "2 / 8", "the click opened the panel instead of moving on");
  await page.click(`${P} .hs-evidence [data-tab="source"]`);
  if (!(await page.isVisible(`${P} .hs-evidence [data-pane="source"]`))) throw new Error("the source tab did not show");
  await shot("evidence-source");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  if (await page.$(`${P} .hs-evidence`)) throw new Error("Esc did not close the panel");
  expect(await counter(), "2 / 8", "Esc closed the panel only");
});

await step("present: switching the view re-sorts the ranking (and does not move on)", async () => {
  await page.keyboard.press("ArrowRight");
  await page.waitForSelector(`${P} .hs-rank`);
  await page.waitForTimeout(1800);
  const order = () => page.$$eval(`${P} .hs-player-slide:last-of-type .hs-rank-row`, (rows) => rows.map((row) => [row.dataset.label, Number(row.style.getPropertyValue("--r"))]).sort((a, b) => a[1] - b[1]).map(([label]) => label));
  const before = await order();
  await page.click(`${P} .hs-rank-views [data-view="1"]`);
  await page.waitForTimeout(1000);
  const after = await order();
  expect(await counter(), "3 / 8", "switching the view kept the slide");
  if (!(await page.$(`${P} .hs-rank-views [data-view="1"][aria-selected="true"]`))) throw new Error("the view button is not marked");
  await shot("rank-view");
  if (before.join() !== "資料作成,議事録,データ集計,その他" || after[0] !== "資料作成") throw new Error(`order: ${before} → ${after}`);
});

await step("present: bars move from before to after and leave the difference dashed", async () => {
  await page.keyboard.press("ArrowRight");
  await page.waitForSelector(`${P} .hs-shift`);
  const width = () => page.$eval(`${P} .hs-player-slide:last-of-type .hs-shift-row .hs-shift-bar`, (el) => el.getBoundingClientRect().width);
  await page.waitForTimeout(1000);
  const early = await width();
  await page.waitForTimeout(1300);
  await shot("shift-mid");
  await page.waitForTimeout(1600);
  const late = await width();
  await shot("shift-end");
  if (!(early > late * 1.6)) throw new Error(`the first bar should shrink from 90 to 35 (${early} → ${late})`);
});

await step("present: dragging a slider recalculates, then the arrow keys move on again", async () => {
  await page.keyboard.press("ArrowRight");
  await page.waitForSelector(`${P} .hs-sim`);
  await page.waitForTimeout(1800);
  expect(await now(".hs-sim-out"), "3,600", "starting result");
  const slider = page.locator(`${P} .hs-player-slide:last-of-type input[data-sim="2"]`);
  const box = await slider.boundingBox();
  await page.mouse.move(box.x + box.width * 0.44, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width * 0.99, box.y + box.height / 2, { steps: 8 });
  await page.mouse.up();
  await page.waitForTimeout(400);
  const result = await now(".hs-sim-out");
  const diff = await now(".hs-sim-diff");
  await shot("simulator-moved");
  if (result === "3,600") throw new Error("the result did not change");
  if (!/上回る|あと/.test(diff)) throw new Error(`difference: ${diff}`);
  expect(await counter(), "5 / 8", "dragging kept the slide");
  await page.keyboard.press("ArrowRight");
  await page.waitForSelector(`${P} .hs-gap`, { timeout: 4000 });
});

await step("present: each click switches the next measure on; a measure can be switched by hand", async () => {
  await page.waitForTimeout(1500);
  expect(await now(".hs-gap-num"), "1,400", "the gap before any measure");
  await shot("gap-start");
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(900);
  expect(await now(".hs-gap-num"), "800", "after the first measure");
  await page.click(`${P} .hs-player-slide:last-of-type [data-measure="2"]`);
  await page.waitForTimeout(900);
  expect(await now(".hs-gap-num"), "400", "after switching the third by hand");
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(900);
  expect(await now(".hs-gap-state"), "目標を上回る", "after the second measure");
  expect(await now(".hs-gap-num"), "+100", "surplus");
  await shot("gap-filled");
  expect(await counter(), "6 / 8", "still on the gap slide");
});

await step("present: the plan shows 'いま'; O opens the overview; D runs the demo to the overview", async () => {
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowRight");
  await page.waitForSelector(`${P} .hs-slide[data-type="gantt"]`, { timeout: 4000 });
  await page.waitForSelector(`${P} .hs-gnow`, { state: "attached" });
  await page.keyboard.press("o");
  await page.waitForSelector(`${P} .hs-player-grid:not([hidden])`);
  await page.keyboard.press("Escape");
  await page.waitForSelector(`${P} .hs-player-grid[hidden]`, { state: "attached" });
  await page.keyboard.press("d");
  await page.waitForSelector(`${P} .hs-demo-cursor`);
  await page.waitForTimeout(5200);
  await shot("demo");
  await page.waitForSelector(`${P} .hs-player-grid:not([hidden])`, { timeout: 40000 });
  if (await page.$(`${P} .hs-demo-cursor`)) throw new Error("the demo cursor stayed after the demo");
  await shot("demo-end");
  await page.keyboard.press("Escape");
  // A demo stops as soon as someone presses a key.
  await page.keyboard.press("Home");
  await page.waitForTimeout(600);
  await page.keyboard.press("d");
  await page.waitForSelector(`${P} .hs-demo-cursor`);
  await page.waitForTimeout(1500);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  if (await page.$(`${P} .hs-demo-cursor`)) throw new Error("Esc did not stop the demo");
  if (await page.isHidden("#presenter")) throw new Error("Esc ended the presentation instead of stopping the demo");
  await page.keyboard.press("Escape");
  await page.waitForSelector("#presenter.hidden", { state: "attached", timeout: 5000 });
});

await step("the exported file: works the same, and #static shows every page finished", async () => {
  // A step that failed above may have left the presentation open.
  for (let i = 0; i < 3 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
  const download = page.waitForEvent("download", { timeout: 60000 });
  await page.click("#downloadBtn");
  await page.waitForTimeout(800);
  if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn");
  const file = join(outDir, "interactive-export.html");
  await (await download).saveAs(file);
  const view = await context.newPage();
  view.on("pageerror", (error) => errors.push(`export pageerror: ${error.message}`));
  await view.goto(`${pathToFileURL(file).href}#5`);
  await view.waitForSelector(".hs-player .hs-sim", { timeout: 20000 });
  await view.waitForTimeout(1500);
  await view.evaluate(() => { const range = document.querySelector('.hs-player input[data-sim="0"]'); range.value = "1200"; range.dispatchEvent(new Event("input", { bubbles: true })); });
  const moved = await view.textContent(".hs-player .hs-sim-out");
  if (moved !== "7,200") throw new Error(`the exported simulator: ${moved}`);
  await view.goto(`${pathToFileURL(file).href}#6-static`);
  await view.reload();
  await view.waitForSelector(".hs-player .hs-gap", { timeout: 20000 });
  await view.waitForTimeout(400);
  const still = await view.evaluate(() => ({
    playing: document.querySelectorAll(".hs-player .hs-slide.hs-play").length,
    num: document.querySelector(".hs-player .hs-gap-num").textContent,
    hash: location.hash,
  }));
  await shot("export-static", view);
  await view.keyboard.press("ArrowRight");
  await view.waitForSelector('.hs-player .hs-slide[data-type="gantt"]', { timeout: 5000 });
  await view.waitForTimeout(300);
  const hidden = await view.$$eval(".hs-player .hs-hidden", (els) => els.length);
  await view.close();
  if (still.playing) throw new Error("#static should not play entrances");
  if (still.num !== "+100") throw new Error(`#static should show every measure on: ${still.num}`);
  if (hidden) throw new Error("#static should show every step of the plan");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
