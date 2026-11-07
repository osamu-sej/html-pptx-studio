// PowerPoint's graphics in a real browser: 図形の結合 (接合・型抜き/合成・切り出し・重なり抽出・単純型抜き, the live
// preview, the order of selection, ⌘Z, editing the points of a shape with a hole) and SmartArt (the gallery, the
// text pane, levels, layouts, colours, converting to shapes and from a text box).
// Usage: node qa/studio-graphics.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name) => { const file = join(outDir, `graphics-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const objects = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide()?.elements ?? [])));
const selection = () => page.evaluate(() => window.__hpptx.selection());
let box = null;
const measure = async () => { box = await page.locator(".slide-wrap .hs-slide").boundingBox(); };
const at = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
const drag = async ([x1, y1], [x2, y2], { steps = 8 } = {}) => { await measure(); await page.mouse.move(...at(x1, y1)); await page.mouse.down(); await page.mouse.move(...at(x2, y2), { steps }); await page.mouse.up(); await page.waitForTimeout(250); };
const click = async (x, y, opts = {}) => { await measure(); await page.mouse.click(...at(x, y), opts); await page.waitForTimeout(150); };
const tab = (label) => page.click(`.rb-tabs [role=tab]:text-is("${label}")`);
const ribbonBtn = async (label) => {
  const direct = page.locator(`.rb-body .rb-btn:has-text("${label}")`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct.click();
  for (const folded of await page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = page.locator(`.rb-pop .rb-fold .rb-btn:has-text("${label}")`).first();
    if (await inside.count()) return inside.click();
  }
  throw new Error(`no ribbon button "${label}"`);
};
const menuItem = (label) => page.locator(`.rb-pop .rb-menu button:has-text("${label}")`).first().click();
const undo = async () => { await page.keyboard.press("Control+z"); await page.waitForTimeout(300); };
const drawShape = async (title, from, to) => {
  await tab("挿入");
  await page.locator('.rb-btn.big:has-text("図形")').first().click();
  await page.waitForSelector(".rb-gallery.shapes");
  await page.click(`.rb-gallery button[title="${title}"]`);
  await drag(from, to);
};
/** Select the two shapes in this order (Shift+click adds). */
const pickInOrder = async (a, b) => {
  await click(1800, 1000);
  await click(...a);
  await page.keyboard.down("Shift");
  await click(...b);
  await page.keyboard.up("Shift");
};
const merge = async (label) => {
  await tab("図形の書式");
  await ribbonBtn("図形の結合");
  await menuItem(label);
  await page.waitForTimeout(400);
};

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.click("#sampleDeckBtn");
await page.waitForSelector(".film-item");
await page.waitForSelector("#ribbon:not([hidden]) .rb-tabs");
await tab("挿入");
await ribbonBtn("新しいスライド");
await menuItem("白紙");
await page.waitForTimeout(400);

let rect = null;
let oval = null;
await step("図形の結合: two shapes, the menu's live preview, 接合 keeps the first shape's id and format; ⌘Z", async () => {
  await drawShape("正方形/長方形", [300, 300], [800, 650]);
  await drawShape("楕円", [600, 450], [1100, 850]);
  [rect, oval] = await objects();
  assert(rect?.shape === "rect" && oval?.shape === "ellipse", "two shapes");
  await pickInOrder([400, 400], [1000, 750]);
  assert((await selection()).join() === `${rect.id},${oval.id}`, `selected in order: ${await selection()}`);
  await tab("図形の書式");
  await ribbonBtn("図形の結合");
  await page.waitForSelector('.rb-pop .rb-menu button:has-text("接合")');
  for (const label of ["接合", "型抜き/合成", "切り出し", "重なり抽出", "単純型抜き"]) assert(await page.locator(`.rb-pop .rb-menu button:has-text("${label}")`).count(), label);
  await page.hover('.rb-pop .rb-menu button:has-text("接合")');
  await page.waitForSelector(".ed-merge-preview path", { timeout: 5000 });
  await shot("merge-preview");
  await menuItem("接合");
  await page.waitForTimeout(400);
  const list = await objects();
  assert(list.length === 1 && list[0].id === rect.id && list[0].shape === "custom", `one shape: ${JSON.stringify(list.map((o) => [o.id, o.shape]))}`);
  assert(list[0].fill === (rect.fill ?? list[0].fill), "the first shape's fill");
  assert(Math.abs(list[0].x - 300) < 3 && Math.abs(list[0].x + list[0].w - 1100) < 3 && Math.abs(list[0].y + list[0].h - 850) < 3, `the box around both: ${list[0].x},${list[0].y},${list[0].w},${list[0].h}`);
  assert(!(await page.locator(".ed-merge-preview").count()), "the preview is gone");
  await shot("merge-union");
  await undo();
  assert((await objects()).length === 2, "⌘Z brings both back");
});

await step("型抜き/合成・重なり抽出・切り出し", async () => {
  await pickInOrder([400, 400], [1000, 750]);
  await merge("型抜き/合成");
  let list = await objects();
  assert(list.length === 1 && list[0].path.parts?.length >= 1, `combine has more than one ring: ${list[0]?.path?.parts?.length}`);
  assert(await page.evaluate((id) => document.querySelector(`.slide-wrap .hs-obj[data-el="${id}"] path`)?.getAttribute("fill-rule"), list[0].id) === "evenodd", "drawn even-odd");
  await shot("merge-combine");
  await undo();
  await pickInOrder([400, 400], [1000, 750]);
  await merge("重なり抽出");
  list = await objects();
  assert(list.length === 1 && list[0].x >= 598 && list[0].y >= 448 && list[0].x + list[0].w <= 802, `only the overlap: ${list[0].x},${list[0].y},${list[0].w},${list[0].h}`);
  await undo();
  await pickInOrder([400, 400], [1000, 750]);
  await merge("切り出し");
  list = await objects();
  assert(list.length === 3, `three pieces: ${list.length}`);
  assert(list.every((o) => o.shape === "custom") && list.some((o) => o.id === rect.id), "pieces; the first keeps the id");
  assert((await selection()).length === 3, "all pieces selected");
  await undo();
});

await step("単純型抜き: the order of selection decides what stays; the hole's points can be edited", async () => {
  await pickInOrder([1000, 750], [400, 400]);
  await merge("単純型抜き");
  let list = await objects();
  assert(list.length === 1 && list[0].id === oval.id, "the oval (selected first) stays");
  const abs = list[0].path.pts.map(([fx, fy]) => [list[0].x + fx * list[0].w, list[0].y + fy * list[0].h]);
  assert(abs.some(([x, y]) => Math.abs(x - 800) < 1.5 && y < 650) && abs.some(([x, y]) => Math.abs(y - 650) < 1.5 && x < 800), "the rectangle's corner is cut out of the oval");
  await undo();
  // A small square inside a big one: subtracting cuts a hole (a second ring).
  await click(1800, 1000);
  await page.keyboard.press("Control+a");
  await page.keyboard.press("Delete");
  await drawShape("正方形/長方形", [400, 200], [1200, 900]);
  await drawShape("正方形/長方形", [700, 450], [900, 650]);
  const [big] = await objects();
  await pickInOrder([500, 300], [800, 550]);
  await merge("単純型抜き");
  list = await objects();
  assert(list.length === 1 && list[0].id === big.id && list[0].path.parts?.length === 1, `a hole: ${JSON.stringify(list[0].path)}`);
  await click(500, 300, { button: "right" });
  await page.click('.ed-menu button:has-text("頂点の編集")');
  await page.waitForTimeout(200);
  const handles = await page.locator(".ed-vertex").count();
  assert(handles === 8, `the points of both rings: ${handles}`);
  // Drag one corner of the hole.
  const hole = page.locator('.ed-vertex[data-handle^="vtx:1:"]').first();
  const hb = await hole.boundingBox();
  await page.mouse.move(hb.x + hb.width / 2, hb.y + hb.height / 2);
  await page.mouse.down();
  await page.mouse.move(hb.x + hb.width / 2 - 30, hb.y + hb.height / 2 - 30, { steps: 4 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  list = await objects();
  assert(list[0].path.parts?.length === 1 && Math.abs(list[0].x - 400) < 2 && Math.abs(list[0].w - 800) < 2, "the hole moved, the outside stayed");
  await page.keyboard.press("Escape");
  await shot("merge-hole");
});

const smartart = async () => (await objects()).find((o) => o.kind === "smartart");

await step("挿入 → SmartArt: the gallery by kind with previews; 基本ステップ goes in with the text pane open", async () => {
  await click(1800, 1000);
  await page.keyboard.press("Control+a");
  await page.keyboard.press("Delete");
  await tab("挿入");
  await ribbonBtn("SmartArt");
  await page.waitForSelector(".sa-gallery .sa-gallery-item");
  const all = await page.locator(".sa-gallery-item").count();
  assert(all >= 14, `every layout: ${all}`);
  await page.click('.sa-gallery-kinds button:text-is("手順")');
  const steps = await page.$$eval(".sa-gallery-item span", (els) => els.map((el) => el.textContent));
  assert(steps.includes("基本ステップ") && !steps.includes("組織図"), `手順 only: ${steps}`);
  assert(await page.locator(".sa-gallery-item .sa-preview .hs-sa-node").count() > 5, "each with a preview");
  await shot("smartart-gallery");
  await page.click('.sa-gallery-item[data-layout="process"]');
  await page.waitForTimeout(400);
  const o = await smartart();
  assert(o?.smartart.layout === "process" && o.smartart.items.length >= 3, `inserted: ${JSON.stringify(o?.smartart)}`);
  assert(await page.locator('.rb-tabs [role=tab]:text-is("SmartArt のデザイン")').count(), "the contextual tab");
  assert(await page.isVisible(".sa-pane"), "the text pane");
  assert(await page.locator(`.slide-wrap .hs-obj[data-el="${o.id}"] .hs-sa-node`).count() >= 3, "drawn on the slide");
});

await step("テキスト ウィンドウ: type, Enter for a new item, Tab for a level; the slide follows", async () => {
  const first = page.locator('.sa-pane input[data-i="0"]');
  await first.click();
  await first.fill("");
  await page.keyboard.type("企画");
  await page.waitForTimeout(200);
  let o = await smartart();
  assert(o.smartart.items[0].text === "企画", `typed: ${o.smartart.items[0].text}`);
  assert((await page.textContent(`.slide-wrap .hs-obj[data-el="${o.id}"]`)).includes("企画"), "on the slide");
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.keyboard.type("市場調査");
  await page.keyboard.press("Tab");
  await page.waitForTimeout(300);
  o = await smartart();
  assert(o.smartart.items[1].text === "市場調査" && o.smartart.items[1].level === 1, `a sub-item: ${JSON.stringify(o.smartart.items.slice(0, 3))}`);
  assert(await page.evaluate(() => document.activeElement?.dataset.i === "1"), "the caret stays on the line");
  await shot("smartart-textpane");
});

await step("SmartArt のデザイン: add a shape, level up, layout gallery, colours, right to left", async () => {
  await tab("SmartArt のデザイン");
  let n = (await smartart()).smartart.items.length;
  await ribbonBtn("図形の");
  await menuItem("後に図形を追加");
  await page.waitForTimeout(300);
  let o = await smartart();
  assert(o.smartart.items.length === n + 1, "a shape added");
  // The new (empty) line has the caret: type into it.
  await page.keyboard.type("新しい工程");
  await page.waitForTimeout(200);
  o = await smartart();
  assert(o.smartart.items.some((it) => it.text === "新しい工程"), "typed into the new shape");
  await page.locator('.sa-pane input[data-i="1"]').click();
  await ribbonBtn("レベル上げ");
  await page.waitForTimeout(300);
  o = await smartart();
  assert(o.smartart.items[1].level === 0, "市場調査 moved up a level");
  await page.locator('.rb-body .rb-group:has(.rb-label:text-is("レイアウト")) .rb-drop').first().click().catch(async () => { await ribbonBtn("レイアウト"); });
  await page.waitForSelector(".sa-gallery .sa-gallery-item");
  await page.click('.sa-gallery-item[data-layout="hierarchy"]');
  await page.waitForTimeout(300);
  assert((await smartart()).smartart.layout === "hierarchy", "layout changed");
  await ribbonBtn("色の");
  await page.waitForSelector(".sa-color-grid .sa-color");
  await page.click('.sa-color[title="カラフル - アクセント 2 から 4"]');
  await page.waitForTimeout(300);
  assert((await smartart()).smartart.color === "colorful", "colour changed");
  await ribbonBtn("右から左");
  await page.waitForTimeout(200);
  assert((await smartart()).smartart.rtl === true, "right to left");
  await ribbonBtn("右から左");
  await page.waitForTimeout(200);
  await shot("smartart-design");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(200);
  assert((await smartart()).smartart.rtl === true, "⌘Z undoes one change");
  await page.keyboard.press("Control+y");
});

await step("double-click an item opens it in the text pane; an animation brings the items in one by one", async () => {
  await page.keyboard.press("Escape");
  const o = await smartart();
  const node = page.locator(`.slide-wrap .hs-obj[data-el="${o.id}"] .hs-sa-node[data-item="2"]`).first();
  const b = await node.boundingBox();
  await page.mouse.dblclick(b.x + b.width / 2, b.y + b.height / 2);
  await page.waitForTimeout(300);
  assert(await page.isVisible(".sa-pane"), "the pane opens");
  assert(await page.evaluate(() => document.activeElement?.dataset.i === "2"), `on that item: ${await page.evaluate(() => document.activeElement?.dataset.i)}`);
  await page.keyboard.press("Escape");
  await tab("アニメーション");
  await page.locator('.an-quick .an-fx[data-fx="in:fade"]').click();
  await page.waitForTimeout(300);
  await ribbonBtn("効果の");
  await menuItem("1つずつ");
  await page.waitForTimeout(300);
  const d = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide().timeline)));
  assert(d.some((e) => e.el === o.id && e.by === "item"), `one by one: ${JSON.stringify(d)}`);
});

await step("slide show: the items come in one after another", async () => {
  await tab("スライド ショー");
  await ribbonBtn("このスライド");
  await page.waitForSelector("#presenter .hs-player-slide .hs-sa-step", { state: "attached", timeout: 10000 });
  await page.waitForTimeout(500);
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(250);
  const early = await page.$$eval("#presenter .hs-player-slide .hs-sa-step:not(.hs-sa-fixed)", (els) => els.map((el) => Number(getComputedStyle(el).opacity) > 0.5 && getComputedStyle(el).visibility !== "hidden"));
  await page.waitForTimeout(4000);
  const late = await page.$$eval("#presenter .hs-player-slide .hs-sa-step:not(.hs-sa-fixed)", (els) => els.map((el) => Number(getComputedStyle(el).opacity) > 0.5 && getComputedStyle(el).visibility !== "hidden"));
  assert(late.every(Boolean), `all shown at the end: ${late}`);
  assert(early.filter(Boolean).length < late.length, `not all at once: ${early}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

await step("変換: 図形に変換 (grouped shapes) and テキストに変換; ホーム → SmartArt に変換 keeps the levels", async () => {
  await page.keyboard.press("Escape");
  let o = await smartart();
  const pickSmartArt = async (id) => { await page.locator(`.slide-wrap .hs-obj[data-el="${id}"] .hs-sa-node`).first().click(); await page.waitForTimeout(200); };
  await pickSmartArt(o.id);
  await tab("SmartArt のデザイン");
  await ribbonBtn("変換");
  await menuItem("図形に変換");
  await page.waitForTimeout(400);
  let list = await objects();
  assert(!list.some((x) => x.kind === "smartart") && list.length > 4 && list.every((x) => x.group && x.group === list[0].group), `grouped shapes: ${list.length}`);
  await undo();
  o = await smartart();
  assert(o, "⌘Z brings the SmartArt back");
  await pickSmartArt(o.id);
  await tab("SmartArt のデザイン");
  await ribbonBtn("変換");
  await menuItem("テキストに変換");
  await page.waitForTimeout(400);
  list = await objects();
  const text = list.find((x) => x.kind === "text");
  assert(text && /<ul>/.test(text.text) && text.text.includes("企画"), `a bulleted text box: ${text?.text}`);
  await tab("ホーム");
  const convert = page.locator('.rb-body .rb-btn[title^="SmartArt に変換"]');
  if (await convert.count() && await convert.first().isVisible()) await convert.first().click();
  else {
    for (const folded of await page.locator(".rb-body .rb-folded").all()) { await folded.click(); const inside = page.locator('.rb-pop .rb-btn[title^="SmartArt に変換"]'); if (await inside.count()) { await inside.first().click(); break; } }
  }
  await page.waitForSelector(".sa-gallery .sa-gallery-item");
  await page.click('.sa-gallery-item[data-layout="vlist"]');
  await page.waitForTimeout(400);
  o = await smartart();
  assert(o && o.id === text.id && o.smartart.layout === "vlist", "the text box became a SmartArt");
  assert(o.smartart.items.some((it) => it.level > 0) && o.smartart.items[0].text === "企画", `levels kept: ${JSON.stringify(o.smartart.items)}`);
  await shot("smartart-converted");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
