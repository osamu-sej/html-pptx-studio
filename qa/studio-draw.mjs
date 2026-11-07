// PowerPoint's 描画 tab and the rest of 挿入 in a real browser: pens and highlighters (strokes in a row join one ink
// object), the eraser, the lasso, ink to shapes, 描画で再生 and its animation; ズーム (slide, summary, section) and
// its tab, カメオ, 数式 (insert, double-click to edit), ワードアート, スライド番号 fields, 段組み, 文字種の変換 — and a
// zoom clicked in the show goes to its slide and back.
// Usage: node qa/studio-draw.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: [...browserArgs, "--use-fake-ui-for-media-stream", "--use-fake-device-for-media-stream"] });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, permissions: ["camera"] });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name) => { const file = join(outDir, `draw-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const objects = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide()?.elements ?? [])));
const slide = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide())));
const selection = () => page.evaluate(() => window.__hpptx.selection());
const index = () => page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide()));
let box = null;
const measure = async () => {
  // The stage is drawn again a few times just after a change: a slide replaced while it was being measured has no box,
  // so it is measured again.
  box = null;
  for (let t = 0; t < 20 && !box; t += 1) {
    if (t) await page.waitForTimeout(100);
    await page.locator("#stageBody .slide-wrap .hs-slide").first().waitFor({ state: "visible", timeout: 5000 });
    box = await page.locator("#stageBody .slide-wrap .hs-slide").first().boundingBox();
  }
  if (!box) throw new Error(`no slide on the stage (${await page.locator("#stageBody .slide-wrap .hs-slide").count()} found, view ${await page.evaluate(() => document.querySelector("#stageBody")?.className)})`);
};
const at = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
/** A pen stroke through these points (slide pixels). */
const stroke = async (points, { steps = 10 } = {}) => {
  await measure();
  await page.mouse.move(...at(...points[0]));
  await page.mouse.down();
  for (const p of points.slice(1)) await page.mouse.move(...at(...p), { steps });
  await page.mouse.up();
  await page.waitForTimeout(300);
};
const click = async (x, y, opts = {}) => { await measure(); await page.mouse.click(...at(x, y), opts); await page.waitForTimeout(150); };
const tab = (label) => page.click(`.rb-tabs [role=tab]:text-is("${label}")`);
/** A ribbon button by its words or its tooltip (opening a folded group when the window is narrow). */
const ribbonBtn = async (label, { title = false } = {}) => {
  const sel = title ? `.rb-btn:not(.rb-folded)[title^="${label}"]` : `.rb-btn:not(.rb-folded):has-text("${label}")`;
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
const menuItem = (label) => page.locator(`.rb-pop .rb-menu button:has-text("${label}")`).first().click();
const undo = async () => { await page.keyboard.press("Control+z"); await page.waitForTimeout(300); };

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

let inkId = null;
await step("描画: the tab, a pen, two strokes in a row make one ink object (drawn as two paths); ⌘Z", async () => {
  await tab("描画");
  for (const label of ["選択", "投げ縄", "消しゴム", "インクを", "描画で", "発表で"]) assert(await page.locator(`.rb-body .rb-btn:has-text("${label}")`).count(), `button ${label}`);
  assert((await page.locator(".rb-body .ink-pen").count()) === 6, "six pens");
  await page.click('.ink-pen[aria-label="ペン（青）"]');
  await page.waitForSelector(".ed-ink-layer");
  assert(await page.locator('.ink-pen[aria-label="ペン（青）"].on').count(), "the pen is on");
  await stroke([[300, 300], [450, 360], [600, 420]]);
  await stroke([[320, 470], [480, 500], [640, 520]]);
  const list = await objects();
  assert(list.length === 1 && list[0].kind === "ink" && list[0].strokes.length === 2, `one ink object with two strokes: ${JSON.stringify(list.map((o) => [o.kind, o.strokes?.length]))}`);
  assert(list[0].strokes[0].color === "#0070c0", "blue ink");
  inkId = list[0].id;
  assert((await page.locator(`.slide-wrap .hs-obj[data-el="${inkId}"] .hs-ink path`).count()) === 2, "drawn as two paths");
  assert(await page.isVisible(".ed-ink-layer"), "the pen stays in hand after a stroke");
  await shot("pen");
  await undo();
  assert((await objects())[0]?.strokes.length === 1, "⌘Z takes back the last stroke");
  await page.keyboard.press("Control+y");
  await page.waitForTimeout(300);
  assert((await objects())[0]?.strokes.length === 2, "⌘Y puts it back");
});

await step("蛍光ペン and the width menu: a wide, pale stroke apart from the rest", async () => {
  await page.click('.ink-pen[aria-label="蛍光ペン（黄）"]');
  await stroke([[300, 960], [700, 960], [1000, 965]]);
  const list = await objects();
  const hl = list.find((o) => o.id !== inkId);
  assert(hl?.kind === "ink" && hl.strokes[0].highlighter && hl.strokes[0].width === 28, `a highlighter stroke: ${JSON.stringify(hl?.strokes?.[0] && { ...hl.strokes[0], pts: undefined })}`);
  await ribbonBtn("太さ");
  await menuItem("極細");
  await page.click('.ink-pen[aria-label="ペン（黒）"]');
  await ribbonBtn("太さ");
  await menuItem("太（16px）");
  await page.waitForSelector(".ed-ink-layer");
  await stroke([[1300, 900], [1500, 920], [1700, 940]]);
  const thick = (await objects()).at(-1);
  assert(thick.strokes[0].width === 16 && thick.strokes[0].color === "#000000", `16px black: ${thick.strokes[0].width} ${thick.strokes[0].color}`);
  await undo();
  await undo();
});

await step("消しゴム: a stroke the eraser crosses goes; the other stays", async () => {
  await ribbonBtn("消しゴム");
  await page.waitForSelector(".ed-ink-layer.tool-eraser");
  await stroke([[450, 320], [450, 400]]);
  const ink = (await objects()).find((o) => o.id === inkId);
  assert(ink?.strokes.length === 1, `one stroke left: ${ink?.strokes.length}`);
  await undo();
  assert((await objects()).find((o) => o.id === inkId)?.strokes.length === 2, "⌘Z brings it back");
});

await step("投げ縄: the strokes inside the loop become their own ink object, selected", async () => {
  await ribbonBtn("投げ縄");
  await page.waitForSelector(".ed-ink-layer.tool-lasso");
  await stroke([[270, 440], [700, 440], [700, 570], [270, 570], [270, 445]], { steps: 6 });
  const list = await objects();
  const picked = await selection();
  assert(picked.length === 1, `one object selected: ${picked}`);
  const caught = list.find((o) => o.id === picked[0]);
  assert(caught?.kind === "ink" && caught.strokes.length === 1 && caught.y > 400, `the lower stroke: ${JSON.stringify(caught && [caught.y, caught.strokes.length])}`);
  assert(list.find((o) => o.id === inkId)?.strokes.length === 1, "the first object keeps the upper stroke");
  assert(!(await page.isVisible(".ed-ink-layer")), "the lasso hands over to selecting");
  await shot("lasso");
  await undo();
});

await step("インクを図形に変換: on, a rectangle and a line drawn by hand become shapes; off again", async () => {
  await click(1850, 1040);
  await tab("描画");
  await ribbonBtn("インクを");
  assert((await page.locator('.rb-body .rb-btn[aria-pressed="true"]:has-text("インクを")').count()) === 1, "pressed");
  await page.click('.ink-pen[aria-label="ペン（青）"]');
  const before = (await objects()).length;
  await stroke([[1000, 300], [1400, 302], [1402, 600], [1000, 598], [1001, 305]], { steps: 12 });
  await stroke([[1000, 760], [1250, 770], [1500, 780]]);
  const made = (await objects()).slice(before);
  assert(made.length === 2, `two new objects: ${JSON.stringify(made.map((o) => o.kind))}`);
  assert(made[0].kind === "shape" && made[0].shape === "rect" && made[0].fill === "none" && made[0].stroke === "#0070c0", `a rectangle: ${JSON.stringify({ kind: made[0].kind, shape: made[0].shape, fill: made[0].fill })}`);
  assert(Math.abs(made[0].x - 1000) < 8 && Math.abs(made[0].w - 402) < 12, `where it was drawn: ${made[0].x} ${made[0].w}`);
  assert(made[1].kind === "line", `a line: ${made[1].kind}`);
  await shot("ink-to-shape");
  await ribbonBtn("インクを");
  await page.keyboard.press("Escape");
  assert(!(await page.isVisible(".ed-ink-layer")), "Esc puts the pen down");
});

await step("描画で再生 (the stage plays the ink being written) and 発表で再生 (a 線を描く animation for each ink object)", async () => {
  await ribbonBtn("描画で");
  await page.waitForSelector(".motion-banner", { timeout: 3000 });
  await page.waitForSelector(".motion-banner", { state: "detached", timeout: 12000 });
  await ribbonBtn("発表で");
  const s = await slide();
  const inks = s.elements.filter((o) => o.kind === "ink").map((o) => o.id);
  const draws = (s.timeline || []).filter((e) => e.fx === "draw" && inks.includes(e.el));
  assert(draws.length === inks.length && draws[0].start === "click", `one draw per ink object: ${JSON.stringify(s.timeline)}`);
  await undo();
});

await step("数式: insert from the dialog (TeX → MathML), double-click to edit", async () => {
  await tab("挿入");
  await ribbonBtn("数式");
  await page.waitForSelector(".eq-dialog[open]");
  await page.fill(".eq-input", "\\frac{a+b}{2}");
  assert(await page.locator(".eq-preview math mfrac").count(), "the preview shows a fraction");
  await page.click(".eq-ok");
  await page.waitForTimeout(300);
  let eq = (await objects()).find((o) => o.kind === "equation");
  assert(eq?.tex === "\\frac{a+b}{2}", `equation object: ${eq?.tex}`);
  assert(await page.locator(`.slide-wrap .hs-obj[data-el="${eq.id}"] math mfrac`).count(), "drawn as MathML");
  await measure();
  const r = await page.locator(`.slide-wrap .hs-obj[data-el="${eq.id}"]`).boundingBox();
  await page.mouse.dblclick(r.x + r.width / 2, r.y + r.height / 2);
  await page.waitForSelector(".eq-dialog[open]");
  assert((await page.inputValue(".eq-input")) === "\\frac{a+b}{2}", "the dialog opens with the equation");
  await page.click('.eq-samples button:has-text("三平方の定理")');
  await page.click(".eq-ok");
  await page.waitForTimeout(300);
  eq = (await objects()).find((o) => o.id === eq.id);
  assert(eq.tex === "a^2 + b^2 = c^2", `updated: ${eq.tex}`);
  assert(await page.locator(`.slide-wrap .hs-obj[data-el="${eq.id}"] math msup`).count(), "drawn with powers");
  await shot("equation");
});

let wordId = null;
await step("ワードアート: a text style in the theme's colours, typed into at once; 文字種の変換 and 段組み", async () => {
  await ribbonBtn("ワードアート");
  await page.waitForSelector(".wa-gallery .wa-item");
  await page.locator(".wa-item").first().click();
  await page.waitForTimeout(400);
  assert((await page.evaluate(() => window.__hpptx.typing())) === "text", "typing in the new text box");
  await page.keyboard.type("acme results");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(250);
  const wa = (await objects()).at(-1);
  wordId = wa.id;
  const accent = await page.evaluate(() => window.SlideEngine.themeColors(window.__hpptx.deck().theme).accent1);
  assert(wa.kind === "text" && wa.color === accent && wa.bold && /acme results/.test(wa.text) && !/ここに文字/.test(wa.text), `word art: ${JSON.stringify({ color: wa.color, bold: wa.bold, text: wa.text })}`);
  await tab("ホーム");
  await ribbonBtn("文字種の変換", { title: true });
  await menuItem("各単語の先頭文字を大文字");
  let o = (await objects()).find((x) => x.id === wordId);
  assert(/Acme Results/.test(o.text), `title case: ${o.text}`);
  await ribbonBtn("文字種の変換", { title: true });
  await menuItem("全角にする");
  o = (await objects()).find((x) => x.id === wordId);
  assert(/Ａｃｍｅ　Ｒｅｓｕｌｔｓ/.test(o.text), `full width: ${o.text}`);
  await ribbonBtn("段組み", { title: true });
  await menuItem("2段");
  o = (await objects()).find((x) => x.id === wordId);
  assert(o.cols === 2, `two columns: ${o.cols}`);
  assert((await page.evaluate((id) => getComputedStyle(document.querySelector(`.slide-wrap .hs-obj[data-el="${id}"] .hs-obj-tx`)).columnCount, wordId)) === "2", "drawn in two columns");
});

await step("スライド番号: a field that shows the slide's number (and follows when slides move)", async () => {
  await click(1850, 1040);
  await tab("挿入");
  await ribbonBtn("スライド番号");
  await menuItem("スライド番号");
  await page.waitForTimeout(300);
  const field = (await objects()).at(-1);
  assert(/data-field="slideno"/.test(field.text), `a field: ${field.text}`);
  const shown = async () => (await page.locator(`.slide-wrap .hs-obj[data-el="${field.id}"] [data-field]`).textContent()).trim();
  const n = await index();
  assert((await shown()) === String(n + 1), `shows ${n + 1}: ${await shown()}`);
  // ⌘↓ on the thumbnails moves the slide (on the stage it would move the selected object).
  await page.locator(`.film-item[data-index="${n}"]`).click();
  await page.keyboard.press("Control+ArrowDown");
  await page.waitForTimeout(400);
  assert((await index()) === n + 1, `the slide moved down: ${await index()}`);
  assert((await shown()) === String(n + 2), `after moving down it shows ${n + 2}: ${await shown()}`);
  await undo();
});

let zoomId = null;
await step("ズーム: a slide zoom (picker of slides), the ズーム tab, ズームに戻る; カメオ", async () => {
  await ribbonBtn("ズーム");
  await menuItem("スライド ズーム");
  await page.waitForSelector(".zm-picker .zm-cell");
  await page.locator('.zm-cell[data-index="2"]').click();
  await page.waitForTimeout(400);
  const deck = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
  const zoom = (await objects()).find((o) => o.kind === "zoom");
  zoomId = zoom?.id;
  assert(zoom && zoom.target === deck.slides[2].sid, `a zoom to slide 3: ${JSON.stringify(zoom)}`);
  assert(await page.locator(`.slide-wrap .hs-obj[data-el="${zoomId}"] .hs-zoom .hs-slide`).count(), "drawn with the slide's picture");
  await page.waitForSelector('.rb-tabs [role=tab]:text-is("ズーム")');
  await tab("ズーム");
  await ribbonBtn("ズームに");
  assert((await objects()).find((o) => o.id === zoomId).back === false, "ズームに戻る off");
  await ribbonBtn("ズームに");
  assert((await objects()).find((o) => o.id === zoomId).back === undefined, "and on again");
  await tab("挿入");
  await ribbonBtn("カメオ");
  await page.waitForTimeout(300);
  const cam = (await objects()).find((o) => o.kind === "camera");
  assert(cam && cam.mask === "ellipse", `a round camera: ${JSON.stringify(cam)}`);
  await shot("zoom-camera");
});

await step("発表: clicking the zoom goes to its slide, and at the end of it the show comes back", async () => {
  const here = await index();
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player");
  await page.waitForTimeout(1200);
  assert(await page.locator(".hs-player .hs-camera-feed").count(), "the camera is in the show");
  const counter = async () => (await page.textContent("#presenter .hs-player-count")).trim().split(" ")[0];
  assert((await counter()) === String(here + 1), `the show starts here: ${await counter()}`);
  await page.locator(`#presenter .hs-obj[data-el="${zoomId}"]`).last().click();
  await page.waitForTimeout(1500);
  assert((await counter()) === "3", `the zoom goes to slide 3: ${await counter()}`);
  for (let i = 0; i < 15 && (await counter()) === "3"; i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(350); }
  assert((await counter()) === String(here + 1), `at the end of slide 3 the show comes back: ${await counter()}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(500);
});

await step("サマリー ズーム: a new slide with a zoom to each chosen slide; ⌘Z takes it away", async () => {
  if (await page.isVisible(".hs-player")) await page.keyboard.press("Escape");
  const count = await page.evaluate(() => window.__hpptx.deck().slides.length);
  await tab("挿入");
  await ribbonBtn("ズーム");
  await menuItem("サマリー ズーム");
  await page.waitForSelector(".zm-picker .zm-cell");
  for (const i of [1, 2, 3]) await page.locator(`.zm-cell[data-index="${i}"]`).click();
  assert((await page.textContent(".zm-ok")).includes("3枚"), "three chosen");
  await page.click(".zm-ok");
  await page.waitForTimeout(500);
  const s = await slide();
  assert((await page.evaluate(() => window.__hpptx.deck().slides.length)) === count + 1, "one more slide");
  assert(s.title === "サマリー" && s.elements.filter((o) => o.kind === "zoom").length === 3, `three zooms: ${JSON.stringify(s.elements.map((o) => o.kind))}`);
  await shot("summary-zoom");
  await undo();
  assert((await page.evaluate(() => window.__hpptx.deck().slides.length)) === count, "⌘Z removes it");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
