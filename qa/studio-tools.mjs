// PowerPoint's helpers in a real browser. 検索 (⌥Q): a command found in hiragana and run with Enter (its tab shown,
// the button lit), one found by its description and chosen with the arrow keys, one inside a group folded by a narrow
// window, the recent commands, the ribbon's commands in ⌘K. クイック アクセス ツール バー: shown from the display
// options, 元に戻す and やり直し, a ribbon button added by a right-click and run without leaving the tab, taken off
// again, その他のコマンド (add, order, below the ribbon) kept after a reload. スポイト: a shape's fill taken from
// another shape on the slide. 音声読み上げ: the slide read block by block (the master's marks left out), paused and
// resumed, an object alone. 拡大 in a slide show: the magnifier and a click, a drag, − / ＋, Esc, another slide.
// Usage: node qa/studio-tools.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
await mkdir(outDir, { recursive: true });

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 } });
// The browser's own eyedropper cannot be driven here: the slide is sampled instead (as in Firefox and Safari).
// The voice is stood in for: what would be said is kept in window.__spoken, each block "said" in 120 ms.
await context.addInitScript(() => {
  window.EyeDropper = undefined;
  window.__spoken = [];
  const synth = {
    current: null,
    speak(u) { this.current = u; window.__spoken.push(u.text); setTimeout(() => { if (this.current === u) { this.current = null; u.onend?.({}); } }, 120); },
    cancel() { this.current = null; },
    pause() {}, resume() {}, addEventListener() {},
    getVoices: () => [{ name: "Kyoko", lang: "ja-JP" }],
  };
  Object.defineProperty(window, "speechSynthesis", { value: synth, configurable: true });
  window.SpeechSynthesisUtterance = class { constructor(text) { this.text = text; } };
});
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g|ERR_ABORTED/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name) => { const file = join(outDir, `tools-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const tab = (label) => page.click(`.rb-tabs [role=tab]:text-is("${label}")`);
const currentTab = () => page.textContent('.rb-tabs [role=tab][aria-selected="true"]');
const rulers = () => page.evaluate(() => Boolean(document.querySelector(".with-rulers")));
const items = () => page.$$eval(".rb-pop .rb-search-item", (els) => els.map((el) => ({ label: el.querySelector(".rb-search-label").textContent, place: el.querySelector("small").textContent })));
const menuItem = (label) => page.locator(`.rb-pop button:has-text("${label}")`).first().click();
const slideCount = () => page.evaluate(() => window.__hpptx.deck().slides.length);

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.click("#sampleDeckBtn");
await page.waitForSelector(".film-item");
await page.waitForSelector("#ribbon:not([hidden]) .rb-tabs");

await step("⌥Q focuses the search; 「るーらー」 finds ルーラー (表示 › 表示/非表示); Enter runs it there, the tab shown and the button lit", async () => {
  await page.locator(".film-item").nth(1).click();
  await page.keyboard.press("Alt+KeyQ");
  assert(await page.evaluate(() => document.activeElement?.classList.contains("rb-search-input")), "the search box has the focus");
  await page.keyboard.type("るーらー");
  await page.waitForSelector(".rb-pop .rb-search-item");
  const [first] = await items();
  assert(first.label === "ルーラー" && first.place === "表示 › 表示/非表示", `first: ${JSON.stringify(first)}`);
  await shot("results");
  const before = await rulers();
  await page.keyboard.press("Enter");
  // The button stays lit for 1.6 s: look for it first, then at what the command did.
  await page.waitForSelector(".rb-body .rb-btn.rb-found", { timeout: 1200 }).catch(() => {});
  assert(await page.locator(".rb-body .rb-btn.rb-found").count() === 1, "the button lit where it lives");
  await page.waitForTimeout(300);
  assert((await rulers()) === !before, "the rulers toggled");
  assert((await currentTab()) === "表示", `tab: ${await currentTab()}`);
});

await step("found by its description and chosen with ↓ ↑: アニメーション ウィンドウ opens", async () => {
  await tab("ホーム");
  await page.click(".rb-search-input");
  await page.keyboard.type("あにめーしょん うぃんどう");
  await page.waitForSelector(".rb-pop .rb-search-item");
  const list = await items();
  assert(list[0].label === "アニメーション ウィンドウ" && list[0].place.startsWith("アニメーション"), `results: ${JSON.stringify(list)}`);
  if (list.length > 1) { await page.keyboard.press("ArrowDown"); await page.keyboard.press("ArrowUp"); }
  assert(await page.locator(".rb-pop .rb-search-item.on").first().textContent().then((t) => t.includes("アニメーション ウィンドウ")), "the first result is the one chosen");
  await page.keyboard.press("Enter");
  await page.waitForSelector(".an-pane-head", { timeout: 5000 });
});

await step("a narrow window folds the ribbon's groups: a command inside a folded group runs too (選択ウィンドウ)", async () => {
  await page.setViewportSize({ width: 960, height: 900 });
  await tab("ホーム");
  await page.waitForTimeout(300);
  const folded = await page.$$eval(".rb-body > .rb-group.folded > .rb-label", (els) => els.map((el) => el.textContent));
  assert(folded.includes("編集"), `folded groups: ${folded}`);
  await page.click(".rb-search-input");
  await page.keyboard.type("選択ウィンドウ");
  await page.waitForSelector(".rb-pop .rb-search-item");
  const k = (await items()).findIndex((it) => it.place === "ホーム › 編集");
  assert(k >= 0, `results: ${JSON.stringify(await items())}`);
  await page.locator(".rb-pop .rb-search-item").nth(k).click();
  await page.waitForSelector(".fp-selection", { timeout: 5000 });
  await page.setViewportSize({ width: 1600, height: 1000 });
});

await step("the search box empty: the commands used last come first", async () => {
  await page.click(".rb-search-input");
  await page.waitForSelector(".rb-pop .rb-search-item");
  assert((await page.textContent(".rb-pop .rb-menu-head")).includes("最近使ったコマンド"), "the heading");
  const labels = (await items()).map((it) => it.label);
  assert(labels[0] === "選択ウィンドウ" && labels.includes("ルーラー") && labels.includes("アニメーション ウィンドウ"), `recent: ${labels}`);
  await page.keyboard.press("Escape");
});

await step("⌘K: the ribbon's commands are in the palette too (once something is typed)", async () => {
  await page.click(".slide-wrap", { position: { x: 5, y: 5 } }).catch(() => {});
  await page.keyboard.press("Control+KeyK");
  await page.waitForSelector("#commandDialog[open]");
  assert(!(await page.locator('#commandList .command-group:text-is("リボン")').count()), "not before typing");
  await page.fill("#commandInput", "ルーラー");
  await page.waitForSelector('#commandList .command-group:text-is("リボン")');
  const before = await rulers();
  await page.locator("#commandList .command-item", { hasText: "ルーラー" }).first().click();
  await page.waitForTimeout(300);
  assert((await rulers()) === !before, "the rulers toggled from the palette");
});

await step("クイック アクセス ツール バー: shown from the display options; 元に戻す and やり直し follow the edits", async () => {
  await page.click(".rb-options");
  await menuItem("クイック アクセス ツール バーを表示する");
  await page.waitForSelector(".rb-tabs .rb-qat");
  const keys = await page.$$eval(".rb-qat .rb-qat-btn", (els) => els.map((el) => el.dataset.qat));
  assert(keys.join(",") === "save,undo,redo,fromStart", `buttons: ${keys}`);
  const n = await slideCount();
  await tab("ホーム");
  await page.locator('.rb-body .rb-btn:has-text("新しいスライド")').first().click();
  await menuItem("白紙");
  await page.waitForFunction((count) => window.__hpptx.deck().slides.length === count + 1, n);
  await page.waitForSelector('.rb-qat-btn[data-qat="undo"]:not([disabled])');
  await page.click('.rb-qat-btn[data-qat="undo"]');
  await page.waitForFunction((count) => window.__hpptx.deck().slides.length === count, n);
  await page.waitForSelector('.rb-qat-btn[data-qat="redo"]:not([disabled])');
  await page.click('.rb-qat-btn[data-qat="redo"]');
  await page.waitForFunction((count) => window.__hpptx.deck().slides.length === count + 1, n);
  await shot("qat");
});

await step("a right-click adds a ribbon button to the toolbar; it runs from there without leaving the tab; and comes off", async () => {
  await tab("表示");
  await page.locator('.rb-body .rb-btn[title^="スライドの上と左に目盛り"]').first().click({ button: "right" });
  await menuItem("クイック アクセス ツール バーに追加");
  const added = page.locator('.rb-qat-btn[data-qat^="view|"]');
  await added.waitFor();
  await tab("ホーム");
  const before = await rulers();
  await added.click();
  await page.waitForTimeout(300);
  assert((await rulers()) === !before, "the rulers toggled from the toolbar");
  assert((await currentTab()) === "ホーム", `the tab stays: ${await currentTab()}`);
  await added.click({ button: "right" });
  await menuItem("クイック アクセス ツール バーから削除");
  await page.waitForFunction(() => !document.querySelector('.rb-qat-btn[data-qat^="view|"]'));
});

await step("その他のコマンド: a command added and moved up, the toolbar below the ribbon, kept after a reload", async () => {
  await page.click(".rb-qat-more");
  await menuItem("その他のコマンド");
  await page.waitForSelector(".qat-dialog[open]");
  await page.fill(".qat-dialog .qat-filter", "グリッド線");
  await page.waitForFunction(() => document.querySelectorAll(".qat-dialog .qat-all option").length > 0);
  await page.selectOption(".qat-dialog .qat-all", { index: 0 });
  const picked = await page.$eval(".qat-dialog .qat-all", (el) => el.value);
  await page.click(".qat-dialog .qat-add");
  await page.selectOption(".qat-dialog .qat-items", picked);
  await page.click(".qat-dialog .qat-up");
  await page.check(".qat-dialog .qat-below");
  await page.click(".qat-dialog .qat-ok");
  await page.waitForSelector(".rb-qat-row .rb-qat.below");
  const keys = await page.$$eval(".rb-qat .rb-qat-btn", (els) => els.map((el) => el.dataset.qat));
  assert(keys.length === 5 && keys[3] === picked, `buttons: ${keys}`);
  await shot("qat-below");
  await page.reload();
  if (await page.isVisible("#sampleDeckBtn")) await page.click("#sampleDeckBtn");
  await page.waitForSelector("#ribbon:not([hidden]) .rb-qat-row .rb-qat");
  const again = await page.$$eval(".rb-qat .rb-qat-btn", (els) => els.map((el) => el.dataset.qat));
  assert(JSON.stringify(again) === JSON.stringify(keys), `after a reload: ${again}`);
});

// ---------------------------------------------------------------- スポイト・音声読み上げ・拡大
const ribbonBtn = async (label) => {
  const sel = `.rb-btn:not(.rb-folded):has-text("${label}")`;
  const direct = page.locator(`.rb-body ${sel}`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct.click();
  for (const folded of await page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = page.locator(`.rb-pop .rb-fold ${sel}`).first();
    if (await inside.count()) return inside.click();
    await page.keyboard.press("Escape");
  }
  throw new Error(`no ribbon button "${label}"`);
};
/** A point of the slide on the stage (slide pixels, 1920 × 1080) on the screen. */
const at = async (x, y) => { const r = await page.locator(".slide-wrap .hs-slide").first().boundingBox(); return [r.x + (x * r.width) / 1920, r.y + (y * r.height) / 1080]; };
const objects = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide()?.elements ?? [])));
async function drawRect(x1, y1, x2, y2) {
  await tab("挿入");
  await ribbonBtn("図形");
  await page.click('.rb-pop .rb-gallery button[title="正方形/長方形"]');
  await page.mouse.move(...(await at(x1, y1)));
  await page.mouse.down();
  await page.mouse.move(...(await at(x2, y2)), { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(200);
  await tab("図形の書式");
  return (await objects()).at(-1);
}

await step("スポイト: a shape's fill taken from another shape on the slide (the selection kept); Esc gives up", async () => {
  await tab("ホーム");
  await ribbonBtn("新しいスライド");
  await menuItem("白紙");
  await page.waitForSelector(".slide-wrap .hs-slide");
  const a = await drawRect(200, 300, 700, 700);
  await page.locator('.rb-body .rb-btn[title="図形の塗りつぶし"]').first().click();
  await page.click('.rb-pop .rb-sw[aria-label="オレンジ"]');
  assert((await objects()).find((o) => o.id === a.id).fill === "#ffc000", "the first shape is orange");
  const b = await drawRect(1000, 300, 1500, 700);
  await page.locator('.rb-body .rb-btn[title="図形の塗りつぶし"]').first().click();
  await page.click(".rb-pop .rb-eyedrop");
  await page.waitForSelector("body.ed-eyedropping");
  await page.mouse.move(...(await at(450, 500)));
  await page.waitForSelector(".ed-eyedrop-swatch:not([hidden])");
  assert((await page.textContent(".ed-eyedrop-swatch")) === "#FFC000", `the colour under the pointer: ${await page.textContent(".ed-eyedrop-swatch")}`);
  await shot("eyedropper");
  await page.mouse.click(...(await at(450, 500)));
  await page.waitForTimeout(200);
  assert((await objects()).find((o) => o.id === b.id).fill === "#ffc000", `taken: ${(await objects()).find((o) => o.id === b.id).fill}`);
  const sel = await page.evaluate(() => [...(window.__hpptx.selection() || [])]);
  assert(sel.length === 1 && sel[0] === b.id, `the shape picked for stays selected: ${JSON.stringify(sel)}`);
  await page.locator('.rb-body .rb-btn[title="図形の塗りつぶし"]').first().click();
  await page.click(".rb-pop .rb-eyedrop");
  await page.waitForSelector("body.ed-eyedropping");
  await page.keyboard.press("Escape");
  await page.waitForFunction(() => !document.body.classList.contains("ed-eyedropping"));
});

await step("音声読み上げ: the slide read block by block, lit as it goes (decoration not read); paused, resumed; an object alone", async () => {
  await page.locator(".film-item").nth(1).click();
  await page.waitForTimeout(300);
  const title = (await page.textContent(".slide-wrap .hs-slide .hs-title")).trim();
  await page.evaluate(() => { window.__spoken.length = 0; });
  await tab("校閲");
  await ribbonBtn("音声");
  await page.waitForSelector(".ra-bar");
  await page.waitForSelector(".slide-wrap .ra-reading");
  await shot("read-aloud");
  await page.waitForFunction(() => window.__spoken.length >= 2);
  await page.click(".ra-bar .ra-play");
  assert((await page.textContent(".ra-bar .ra-play")) === "▶", "paused");
  const said = await page.evaluate(() => window.__spoken.length);
  await page.waitForTimeout(400);
  assert((await page.evaluate(() => window.__spoken.length)) === said, "nothing more is said while paused");
  await page.click(".ra-bar .ra-play");
  await page.waitForFunction(() => !document.querySelector(".ra-bar"), null, { timeout: 15000 });
  const spoken = await page.evaluate(() => window.__spoken.slice());
  assert(spoken[0].replace(/\s/g, "") === title.replace(/\s/g, ""), `the title first: ${spoken[0]} / ${title}`);
  assert(!(await page.locator(".ra-reading").count()), "the light goes when it ends");
  // An object selected: only its words.
  const blank = await page.evaluate(() => window.__hpptx.deck().slides.findIndex((s) => (s.elements || []).length >= 2));
  await page.locator(".film-item").nth(blank).click();
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  await page.mouse.click(...(await at(300, 900)));
  await page.keyboard.type("読み上げるのはこの文だけ");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  await page.evaluate(() => { window.__spoken.length = 0; });
  await tab("校閲");
  await ribbonBtn("音声");
  await page.waitForFunction(() => !document.querySelector(".ra-bar"), null, { timeout: 8000 });
  const only = await page.evaluate(() => window.__spoken.slice());
  assert(only.length === 1 && only[0] === "読み上げるのはこの文だけ", `the object alone: ${JSON.stringify(only)}`);
});

await step("拡大 in a slide show: the magnifier and a click zoom in there (the show stays), a drag looks around, − / ＋, Esc, another slide starts whole", async () => {
  await page.locator(".film-item").nth(1).click();
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector("#presenter .hs-player");
  await page.waitForTimeout(600);
  const index = () => page.evaluate(() => document.querySelector("#presenter .hs-player-count")?.textContent);
  const where = await index();
  const transform = () => page.evaluate(() => document.querySelector("#presenter .hs-player-stage").style.transform);
  await page.mouse.move(800, 700);
  await page.mouse.move(800, 960);
  await page.click('#presenter .hs-player-btn[title^="スライドを拡大"]');
  await page.waitForSelector("#presenter .hs-player.zoom-pick");
  await page.mouse.click(400, 300);
  await page.waitForSelector("#presenter .hs-player.zoomed");
  assert(/scale\(2\)/.test(await transform()), `zoomed twice: ${await transform()}`);
  assert((await index()) === where, "the click did not move the show on");
  const before = await transform();
  await page.mouse.move(800, 500);
  await page.mouse.down();
  await page.mouse.move(650, 420, { steps: 5 });
  await page.mouse.up();
  assert((await transform()) !== before && /scale\(2\)/.test(await transform()), `dragged: ${await transform()}`);
  assert((await index()) === where, "a drag does not move the show on");
  await shot("show-zoom");
  await page.keyboard.press("-");
  assert(/scale\(1\.5\)/.test(await transform()), `zoomed out a step: ${await transform()}`);
  await page.keyboard.press("Escape");
  assert(!(await page.locator("#presenter .hs-player.zoomed").count()) && (await transform()) === "", "Esc shows the whole slide");
  assert(await page.locator("#presenter .hs-player").count(), "and the show is still on");
  await page.keyboard.press("+");
  await page.waitForSelector("#presenter .hs-player.zoomed");
  await page.keyboard.press("ArrowRight");
  await page.waitForFunction(() => !document.querySelector("#presenter .hs-player.zoomed"));
  await page.keyboard.press("Escape");
  await page.keyboard.press("Escape");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
