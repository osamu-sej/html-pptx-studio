// The motion IDE in a real browser: HTML's moves put on a PowerPoint deck automatically as it comes in (one click
// takes them back, a setting turns it off), and the animation pane listing everything that moves on the slide by
// when it happens — the slide's own motion (transition, entrance, build, title, background, emphasis), the
// animations, and the objects' interactions — each changed or taken off right there.
// Usage: node qa/studio-motion-ide.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { execFileSync } from "node:child_process";
import { mkdir } from "node:fs/promises";
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
const decks = join(outDir, "import");
await mkdir(decks, { recursive: true });
execFileSync(process.env.PYTHON_BIN || "python3", [join(root, "test", "pptx_fixtures.py"), decks], { stdio: "ignore" });

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 } });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const shot = async (name) => { const file = join(outDir, `motion-ide-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
  for (let i = 0; i < 3 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
const POWERPOINT = new Set(["fade", "flyIn"]);
const HTML_FX = new Set(["maskRise", "chartGrow", "countUp", "draw", "typewriter", "decode", "blurIn"]);
async function importTemplate() {
  await page.setInputFiles("#importDeckFile", join(decks, "template.pptx"));
  await page.waitForSelector("#importModeDialog[open]");
  await page.click('#importModeDialog [data-mode="exact"]');
  await page.waitForSelector("#msgBar .callout", { timeout: 30000 });
  await page.waitForTimeout(500);
}

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.waitForSelector("#importDeckBtn");
await importTemplate();

await step("a PowerPoint deck gets HTML's moves as it comes in; PowerPoint's own animations stay, in their order", async () => {
  const bar = await page.textContent("#msgBar");
  assert(/HTMLの動きを自動で付けました/.test(bar), `message bar: ${bar}`);
  const d = await deck();
  const timeline = d.slides[1].timeline || [];
  assert(timeline.some((e) => HTML_FX.has(e.fx)), `HTML moves on slide 2: ${timeline.map((e) => e.fx)}`);
  assert(timeline.filter((e) => POWERPOINT.has(e.fx)).map((e) => e.fx).join() === "fade,flyIn", "PowerPoint's animations stay in order");
  assert(d.slides.flatMap((s) => s.elements || []).some((o) => o.hover), "cards answer the mouse");
  await shot("auto-on-import");
});

await step("「元に戻す」 takes the automatic moves back: the deck as it came from PowerPoint", async () => {
  await page.click('#msgBar button:has-text("元に戻す")');
  await page.waitForTimeout(400);
  const d = await deck();
  assert((d.slides[1].timeline || []).map((e) => e.fx).join() === "fade,flyIn", `only PowerPoint's: ${(d.slides[1].timeline || []).map((e) => e.fx)}`);
  assert(!d.slides.flatMap((s) => s.elements || []).some((o) => o.hover), "no hovers left");
  assert(await page.locator('#msgBar button:has-text("HTMLの動きをおまかせで付ける")').count(), "the bar offers おまかせ again");
  await page.click('#msgBar button:has-text("HTMLの動きをおまかせで付ける")');
  await page.waitForTimeout(400);
  assert((await deck()).slides[1].timeline.some((e) => HTML_FX.has(e.fx)), "added again by hand");
});

await step("the animation pane lists everything that moves on the slide, by when it happens", async () => {
  await page.click(".film-item:nth-child(2)");
  await page.waitForTimeout(300);
  await page.click('#msgBar button:has-text("アニメーションを編集")');
  await page.waitForSelector("#animPane:not([hidden]) .an-auto");
  const heads = await page.$$eval("#animPane .an-sec-head b", (els) => els.map((el) => el.textContent));
  assert(heads.join() === "スライドの自動の動き,アニメーション,インタラクション（HTML）,資料全体", `sections: ${heads}`);
  assert(await page.locator('#animPane select[data-auto="transition"]').count(), "the transition row");
  assert(/元の見た目を守る/.test(await page.textContent("#animPane .an-auto")), "a page brought over says why it has no layout motion");
  assert((await page.locator("#animPane .an-row").count()) >= 3, "animations listed");
  assert((await page.locator("#animPane .an-ix-row").count()) >= 2, "interactions listed");
  await shot("pane");
});

await step("changing and taking off right in the pane: a transition, a chip; ⌘Z brings it back", async () => {
  await page.selectOption('#animPane select[data-auto="transition"]', "wipe");
  await page.waitForTimeout(300);
  assert((await deck()).slides[1].transition === "wipe", "slide transition");
  const row = page.locator("#animPane .an-ix-row").first();
  const id = await row.getAttribute("data-id");
  const before = (await deck()).slides[1].elements.find((o) => o.id === id);
  await row.locator(".an-ix-chip button").first().click();
  await page.waitForTimeout(300);
  const after = (await deck()).slides[1].elements.find((o) => o.id === id);
  assert(before.hover && !after.hover, "the hover is taken off");
  await page.click("#stageBody", { position: { x: 5, y: 5 } });
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  assert((await deck()).slides[1].elements.find((o) => o.id === id).hover === before.hover, "⌘Z brings it back");
  await page.locator("#animPane .an-ix-row .an-ix-name").first().click();
  await page.waitForTimeout(300);
  assert((await page.evaluate(() => window.__hpptx.selection())).includes(id), "the row's name selects the object");
  assert(await page.locator('.rb-tabs [role=tab][aria-selected="true"]:has-text("インタラクション")').count(), "and opens the インタラクション tab");
});

await step("an AI-made slide: its automatic motion rows change the slide (build, title, entrance)", async () => {
  await page.evaluate(() => localStorage.clear());
  await page.goto(base);
  await page.click("#sampleDeckBtn");
  await page.waitForSelector(".slide-wrap .hs-slide");
  const index = await page.evaluate(() => window.__hpptx.deck().slides.findIndex((s) => ["cards", "content", "headerCards", "bulletCards"].includes(s.type)));
  await page.click(`.film-item:nth-child(${index + 1})`);
  await page.waitForTimeout(300);
  await page.click("#animTab");
  await page.waitForSelector("#animPane .an-auto select");
  const keys = await page.$$eval("#animPane .an-auto select", (els) => els.map((el) => el.dataset.auto));
  for (const key of ["transition", "entrance", "animation", "kinetic", "backdrop", "emphasis"]) assert(keys.includes(key), `row ${key}: ${keys}`);
  await page.selectOption('#animPane select[data-auto="animation"]', "click");
  await page.waitForTimeout(300);
  assert((await deck()).slides[index].animation === "click", "build: click");
  assert(/動き：クリックで/.test(await page.textContent("#slidePos")), "the status bar follows");
  await page.selectOption('#animPane select[data-auto="entrance"]', "none");
  await page.waitForTimeout(300);
  assert((await deck()).slides[index].entrance === "none", "entrance: none");
  await page.selectOption('#animPane select[data-auto="entrance"]', "auto");
  await page.waitForTimeout(300);
  assert(!("entrance" in (await deck()).slides[index]), "back to the deck's setting");
  await shot("ai-slide");
});

await step("インタラクション → 「取り込み時に自動」 off: the next import comes in as it is", async () => {
  await page.click('.rb-tabs [role=tab]:has-text("インタラクション")');
  const toggle = page.locator('.rb-body .rb-btn:has-text("取り込み時に自動")');
  assert((await toggle.getAttribute("aria-pressed")) === "true", "on by default");
  await toggle.click();
  await page.waitForTimeout(200);
  assert(await page.evaluate(() => localStorage.getItem("hpptx-auto-html")) === "0", "remembered off");
  await page.click("#createTab");
  await importTemplate();
  assert(!/自動で付けました/.test(await page.textContent("#msgBar")), "no automatic moves");
  assert((await deck()).slides[1].timeline.map((e) => e.fx).join() === "fade,flyIn", "PowerPoint's animations only");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
