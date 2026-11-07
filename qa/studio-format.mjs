// 図の形式 and グラフ要素 and 校閲 in a real browser: 背景の削除 and 透明色を指定 (the new picture's pixels are checked),
// 図の圧縮 (a 3000-pixel picture comes back 1280 wide), グラフ要素を追加 (data labels, legend) and 行/列の切り替え,
// 比較 with a deck in the library (a changed slide taken back), インクの非表示, and the right-click menu's 図として保存,
// リンクを開く / リンクの削除 and 図とサイズのリセット.
// Usage: node qa/studio-format.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { execFileSync } from "node:child_process";
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
// Test pictures: a navy square on a white background (and a large one for compression).
const logo = join(outDir, "format-logo.png");
const big = join(outDir, "format-big.png");
execFileSync("python3", ["-c", `
from PIL import Image, ImageDraw
im = Image.new("RGB", (400, 300), (255, 255, 255)); d = ImageDraw.Draw(im); d.rectangle([100, 75, 300, 225], fill=(31, 56, 100)); im.save("${logo}")
im = Image.new("RGB", (3000, 2000), (220, 228, 242)); d = ImageDraw.Draw(im)
for i in range(0, 3000, 50): d.line([(i, 0), (3000 - i, 2000)], fill=(31, 56, 100), width=3)
im.save("${big}")
`]);

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 } });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name) => { const file = join(outDir, `format-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
const slide = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide())));
const tab = (label) => page.click(`.rb-tabs [role=tab]:text-is("${label}")`);
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
const menuItem = (label) => page.locator(`.rb-pop .rb-menu button:has-text("${label}")`).first().click();
const undo = async () => { await page.keyboard.press("Control+z"); await page.waitForTimeout(300); };
const withFiles = async (paths, open) => { const [chooser] = await Promise.all([page.waitForEvent("filechooser"), open()]); await chooser.setFiles(paths); };
/** The alpha at (fx, fy) of a picture object's image (fractions of the image), and its natural size. */
const pixel = (id, fx, fy) => page.evaluate(async ([id, fx, fy]) => {
  const img = document.querySelector(`#stageBody .hs-obj[data-el="${id}"] img`);
  await img.decode?.().catch(() => {});
  const c = document.createElement("canvas");
  c.width = img.naturalWidth;
  c.height = img.naturalHeight;
  const g = c.getContext("2d");
  g.drawImage(img, 0, 0);
  const d = g.getImageData(Math.floor(fx * c.width), Math.floor(fy * c.height), 1, 1).data;
  return { a: d[3], w: c.width, h: c.height };
}, [id, fx, fy]);
const insertPicture = async (file) => {
  const count = (await slide()).elements?.filter((o) => o.kind === "image").length || 0;
  await tab("挿入");
  await ribbonBtn("画像");
  await withFiles([file], () => menuItem("このデバイス"));
  await page.waitForFunction((n) => (window.__hpptx.slide().elements || []).filter((o) => o.kind === "image").length > n, count, { timeout: 8000 });
  await page.waitForTimeout(500);
  return (await slide()).elements.at(-1);
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

await step("背景の削除: the white around the square goes, the square stays (one ⌘Z)", async () => {
  const o = await insertPicture(logo);
  assert(o.kind === "image", "a picture");
  await page.waitForSelector('.rb-tabs [role=tab]:text-is("図の形式")');
  await tab("図の形式");
  await ribbonBtn("背景の");
  await page.waitForFunction((src) => window.__hpptx.slide().elements.at(-1).src !== src, o.src, { timeout: 8000 });
  await page.waitForTimeout(400);
  const n = (await slide()).elements.at(-1);
  assert(/^idb:/.test(n.src) && n.id === o.id, `a new picture: ${n.src}`);
  assert((await pixel(n.id, 0.02, 0.02)).a === 0, "the corner is transparent");
  assert((await pixel(n.id, 0.5, 0.5)).a === 255, "the square stays");
  await shot("background");
  await undo();
  assert((await slide()).elements.at(-1).src === o.src, "⌘Z brings the old picture back");
});

await step("透明色を指定: a click on the white makes the white transparent", async () => {
  const o = (await slide()).elements.at(-1);
  await tab("図の形式");
  await ribbonBtn("透明色を指定");
  await page.waitForSelector(".ed-pick-color");
  const r = await page.locator(`#stageBody .hs-obj[data-el="${o.id}"]`).boundingBox();
  await page.mouse.click(r.x + r.width * 0.05, r.y + r.height * 0.05);
  await page.waitForFunction((src) => window.__hpptx.slide().elements.at(-1).src !== src, o.src, { timeout: 8000 });
  await page.waitForTimeout(400);
  const n = (await slide()).elements.at(-1);
  assert((await pixel(n.id, 0.9, 0.9)).a === 0 && (await pixel(n.id, 0.5, 0.5)).a === 255, "white gone, navy kept");
});

await step("図の圧縮: a 3000-pixel picture becomes 1280 pixels wide", async () => {
  const o = await insertPicture(big);
  const before = await pixel(o.id, 0.5, 0.5);
  assert(before.w >= 2000, `a large picture first: ${before.w}`);
  await tab("図の形式");
  await ribbonBtn("図の圧縮");
  await page.waitForSelector(".compress-dialog[open]");
  await page.selectOption(".compress-dialog select", "1280");
  await page.click(".compress-dialog .cp-ok");
  await page.waitForFunction((src) => window.__hpptx.slide().elements.find((x) => x.kind === "image" && x.src !== src && x.fileName?.includes("big")), o.src, { timeout: 10000 });
  await page.waitForTimeout(500);
  const n = (await slide()).elements.find((x) => x.id === o.id);
  const after = await pixel(n.id, 0.5, 0.5);
  assert(after.w === 1280, `1280 wide: ${after.w}`);
});

await step("グラフ要素を追加 and 行/列の切り替え", async () => {
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="clustered-bar"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  let o = (await slide()).elements.at(-1);
  const node = () => page.locator(`#stageBody .hs-obj[data-el="${o.id}"]`);
  assert((await node().locator(".hs-val").count()) > 0, "values shown");
  await tab("グラフのデザイン");
  await ribbonBtn("グラフ要素");
  await page.locator('.rb-pop .rb-menu button:has-text("なし")').first().click();
  await page.waitForTimeout(300);
  assert((await node().locator(".hs-val").count()) === 0, "data labels off");
  await ribbonBtn("グラフ要素");
  await page.locator('.rb-pop .rb-menu button:text-is("下")').click().catch(async () => menuItem("下"));
  await page.waitForTimeout(300);
  assert(await node().locator(".hs-legend.at-bottom").count(), "the legend below");
  const before = (await slide()).elements.at(-1).chart;
  await ribbonBtn("行/列の切り替え");
  await page.waitForTimeout(300);
  o = (await slide()).elements.at(-1);
  assert(JSON.stringify(o.chart.labels) === JSON.stringify(before.series.map((s) => s.name)) && o.chart.series.length === before.labels.length, `transposed: ${JSON.stringify(o.chart.labels)}`);
  await shot("chart-elements");
});

await step("比較: with the deck saved in the library, a changed title is found and taken back", async () => {
  await page.click("#saveDeckBtn");
  await page.waitForTimeout(800);
  const original = (await deck()).slides[3].title;
  await page.locator('.film-item[data-index="3"]').click();
  await page.waitForTimeout(300);
  await page.evaluate(() => { window.__hpptx.slide().title = "書き換えたタイトル"; });
  await page.locator('.film-item[data-index="2"]').click();
  await page.locator('.film-item[data-index="3"]').click();
  await page.waitForTimeout(300);
  await tab("校閲");
  await ribbonBtn("比較");
  await page.waitForSelector(".compare-dialog[open] .st-source");
  await page.locator(`.compare-dialog .st-source:has-text("${(await deck()).title.slice(0, 6)}")`).first().click();
  await page.waitForSelector(".compare-dialog .cmp-item");
  const items = await page.$$eval(".compare-dialog .cmp-item", (els) => els.map((el) => [el.dataset.kind, el.querySelector(".cmp-what").textContent]));
  assert(items.some(([kind, what]) => kind === "changed" && what.includes("タイトル")), `found: ${JSON.stringify(items)}`);
  await shot("compare");
  await page.locator('.compare-dialog .cmp-item[data-kind="changed"] .cmp-take').first().click();
  await page.waitForTimeout(400);
  assert((await deck()).slides[3].title === original, `taken back: ${(await deck()).slides[3].title}`);
  await page.keyboard.press("Escape");
});

await step("インクの非表示: ink on the slide is hidden while editing, and shown again", async () => {
  await page.evaluate(() => { const s = window.__hpptx.slide(); s.elements = [...(s.elements || []), { id: "ink1", kind: "ink", x: 200, y: 300, w: 400, h: 120, strokes: [{ pts: [[0, 0], [1, 1]], color: "#1f3864", width: 6 }] }]; });
  await page.locator('.film-item[data-index="2"]').click();
  await page.locator('.film-item[data-index="3"]').click();
  await page.waitForTimeout(300);
  assert(await page.isVisible('#stageBody .hs-obj[data-el="ink1"]'), "ink shown");
  await tab("校閲");
  await ribbonBtn("インクの");
  assert(await page.locator("#stageBody.hide-ink").count(), "hidden class");
  assert(!(await page.isVisible('#stageBody .hs-obj[data-el="ink1"]')), "ink hidden");
  await ribbonBtn("インクの");
  assert(await page.isVisible('#stageBody .hs-obj[data-el="ink1"]'), "ink back");
});

/** The colour at (fx, fy) of a picture object's image. */
const rgb = (id, fx, fy) => page.evaluate(async ([id, fx, fy]) => {
  const img = document.querySelector(`#stageBody .hs-obj[data-el="${id}"] img`);
  await img.decode?.().catch(() => {});
  const c = document.createElement("canvas");
  c.width = img.naturalWidth;
  c.height = img.naturalHeight;
  const g = c.getContext("2d");
  g.drawImage(img, 0, 0);
  return Array.from(g.getImageData(Math.floor(fx * c.width), Math.floor(fy * c.height), 1, 1).data.slice(0, 3));
}, [id, fx, fy]);

await step("アート効果「線画」: the square's edge drawn dark on white (one ⌘Z back); 色の変更「アクセント 1（濃色）」 tones the picture", async () => {
  const o = await insertPicture(logo);
  await tab("図の形式");
  await ribbonBtn("アート");
  await menuItem("線画");
  await page.waitForFunction((src) => window.__hpptx.slide().elements.at(-1).src !== src, o.src, { timeout: 15000 });
  await page.waitForTimeout(400);
  const n = (await slide()).elements.at(-1);
  const inside = await rgb(n.id, 0.5, 0.5);
  const edge = await rgb(n.id, 0.25, 0.5);
  assert(inside[0] > 200 && edge[0] < 150, `line drawing: inside ${inside}, edge ${edge}`);
  await shot("art-lines");
  await undo();
  assert((await slide()).elements.at(-1).src === o.src, "⌘Z brings the picture back");
  await tab("図の形式");
  await page.locator('.rb-body .rb-btn[title^="色の変更"]').first().click();
  await menuItem("アクセント 1（濃色）");
  await page.waitForFunction((src) => window.__hpptx.slide().elements.at(-1).src !== src, o.src, { timeout: 15000 });
  await page.waitForTimeout(400);
  const toned = await rgb((await slide()).elements.at(-1).id, 0.5, 0.5);
  assert(toned[2] >= toned[1] && toned[1] >= toned[0] && toned[2] < 160, `a tone of the blue accent: ${toned}`);
});

/** A ribbon button found by the start of its tooltip (the ▾ halves of split buttons have no words), folded groups too. */
const byTitle = async (start) => {
  const sel = `.rb-btn:not(.rb-folded)[title^="${start}"]`;
  const direct = page.locator(`.rb-body ${sel}`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct.click();
  for (const folded of await page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = page.locator(`.rb-pop .rb-fold ${sel}`).first();
    if (await inside.count()) return inside.click();
    await page.keyboard.press("Escape");
  }
  throw new Error(`no ribbon button titled "${start}…"`);
};

await step("箇条書き ▾ ◆ and 段落番号 ▾ ①; 文字列の方向 270 度; リンクのスクリーンヒント", async () => {
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  const r = await page.locator(".slide-wrap .hs-slide").first().boundingBox();
  await page.mouse.click(r.x + r.width * 0.2, r.y + r.height * 0.75);
  await page.waitForTimeout(200);
  await page.keyboard.type("一行目");
  await page.keyboard.press("Enter");
  await page.keyboard.type("二行目");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  const id = (await slide()).elements.at(-1).id;
  const text = async () => (await slide()).elements.find((o) => o.id === id).text;
  await tab("ホーム");
  await byTitle("箇条書きの種類");
  await page.click('.rb-pop .rb-list-sw[data-list-style="diamond"]');
  await page.waitForTimeout(200);
  assert((await text()).startsWith('<ul data-style="diamond">'), `bullets: ${await text()}`);
  const marker = await page.evaluate((id) => getComputedStyle(document.querySelector(`#stageBody .hs-obj[data-el="${id}"] ul`)).listStyleType, id);
  assert(marker.includes("◆"), `drawn with ◆: ${marker}`);
  await byTitle("段落番号の種類");
  await page.click('.rb-pop .rb-list-sw[data-list-style="circled"]');
  await page.waitForTimeout(200);
  assert((await text()).startsWith('<ol data-style="circled">'), `numbers: ${await text()}`);
  await byTitle("文字列の方向");
  await menuItem("270 度");
  await page.waitForTimeout(200);
  assert((await slide()).elements.find((o) => o.id === id).textRot === 270, "turned 270°");
  assert(await page.locator(`#stageBody .hs-obj[data-el="${id}"] .hs-obj-text.is-rot270`).count(), "drawn turned");
  await shot("bullets-direction");
  await byTitle("文字列の方向");
  await menuItem("横書き");
  await page.waitForTimeout(200);
  assert(!(await slide()).elements.find((o) => o.id === id).textRot, "横書き again");
  // ヒント設定: the words shown when the pointer is on it (with a click that moves on).
  await tab("挿入");
  await ribbonBtn("リンク");
  await page.waitForSelector(".rb-pop .rb-form");
  await page.selectOption('.rb-pop .rb-form select[aria-label="クリックしたときの動作"]', "next");
  await page.fill(".rb-pop .rb-link-tip", "次のページへ進みます");
  await page.click('.rb-pop .rb-form button:has-text("設定する")');
  await page.waitForTimeout(200);
  const o = (await slide()).elements.find((x) => x.id === id);
  assert(o.action?.type === "next" && o.tip === "次のページへ進みます", `the link and its ScreenTip: ${JSON.stringify({ action: o.action, tip: o.tip })}`);
});

await step("トリミング ▾: 縦横比 16:9 crops the middle of a 4:3 picture; 枠に合わせる shows it whole again", async () => {
  const pic = await insertPicture(logo);
  await tab("図の形式");
  await byTitle("トリミングのオプション");
  await menuItem("16:9");
  await page.waitForTimeout(300);
  let o = (await slide()).elements.find((x) => x.id === pic.id);
  assert(Math.abs(o.w / o.h - 16 / 9) < 0.02 && o.crop && o.crop.t > 0.1 && !o.crop.l, `16:9: ${o.w}×${o.h} ${JSON.stringify(o.crop)}`);
  await byTitle("トリミングのオプション");
  await menuItem("枠に合わせる");
  await page.waitForTimeout(300);
  o = (await slide()).elements.find((x) => x.id === pic.id);
  assert(!o.crop && Math.abs(o.w / o.h - 4 / 3) < 0.02, `fit: ${o.w}×${o.h} ${JSON.stringify(o.crop)}`);
});

await step("デザイン → 背景の書式設定: a theme colour, a picture at 50%, すべてに適用, 背景のリセット", async () => {
  await page.locator(".film-item").nth(2).click();
  await tab("デザイン");
  await byTitle("このスライドの背景");
  await page.waitForSelector(".bg-dialog[open]");
  // アクセント 1、明るく 80% of the deck's theme.
  const tint = await page.evaluate(() => window.SlideEngine.paletteSheet(window.__hpptx.deck().theme).columns[4].variants[0].hex);
  await page.click(`.bg-dialog .bg-sw[data-color="${tint}"]`);
  await page.waitForTimeout(250);
  assert((await slide()).background?.color === tint, `the colour: ${JSON.stringify((await slide()).background)}`);
  const painted = await page.evaluate(() => getComputedStyle(document.querySelector("#stageBody .slide-wrap .hs-bg")).backgroundColor);
  const [tr, tg, tb] = [1, 3, 5].map((i) => parseInt(tint.slice(i, i + 2), 16));
  assert(painted.includes(`${tr}, ${tg}, ${tb}`), `drawn under the slide: ${painted}`);
  await withFiles([logo], () => page.click(".bg-dialog .bg-file"));
  await page.waitForFunction(() => window.__hpptx.slide().background?.image, null, { timeout: 8000 });
  await page.locator(".bg-dialog .bg-transparency").fill("50");
  await page.waitForTimeout(250);
  let bg = (await slide()).background;
  assert(bg.image?.startsWith("idb:") && bg.transparency === 0.5 && !bg.color, `the picture at 50%: ${JSON.stringify(bg)}`);
  assert(await page.evaluate(() => getComputedStyle(document.querySelector("#stageBody .slide-wrap .hs-bg-img")).opacity) === "0.5", "drawn at half");
  await shot("background");
  await page.click(".bg-dialog .bg-all");
  await page.waitForTimeout(250);
  assert((await deck()).slides.every((s) => s.background?.image === bg.image), "every slide");
  await page.click(".bg-dialog .bg-reset");
  await page.waitForTimeout(250);
  assert(!(await slide()).background, "this slide's background reset");
  await page.click('.bg-dialog .btn-primary');
  await undo();
  await undo();
  bg = (await slide()).background;
  assert(!(await deck()).slides.some((s) => s.background), `two undos: as before (${JSON.stringify(bg)})`);
});

await step("均等割り付け (⇧⌘J), オートコレクト ((c) → ©, 「・ 」で箇条書き, 「1. 」で段落番号)", async () => {
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  const r = await page.locator(".slide-wrap .hs-slide").first().boundingBox();
  await page.mouse.click(r.x + r.width * 0.25, r.y + r.height * 0.6);
  await page.waitForTimeout(200);
  await page.keyboard.type("著作権 (c) と --> 矢印");
  await page.keyboard.press("Enter");
  await page.keyboard.type("・ 一つ目");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const o = (await slide()).elements.at(-1);
  assert(/著作権 © と → 矢印/.test(o.text), `symbols: ${o.text}`);
  assert(/<ul><li>一つ目<\/li><\/ul>/.test(o.text), `a bullet from 「・ 」: ${o.text}`);
  await page.keyboard.press("Control+Shift+j");
  await page.waitForTimeout(200);
  assert((await slide()).elements.find((x) => x.id === o.id).align === "distributed", "⇧⌘J: 均等割り付け");
  const css = await page.evaluate((id) => { const s = getComputedStyle(document.querySelector(`#stageBody .hs-obj[data-el="${id}"] .hs-obj-tx`)); return `${s.textAlign}|${s.textAlignLast}`; }, o.id);
  assert(css === "justify|justify", `drawn justified to the last line: ${css}`);
  // A numbered paragraph from 「1. 」, in a new box; turned off, it stays as typed.
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  await page.mouse.click(r.x + r.width * 0.25, r.y + r.height * 0.8);
  await page.waitForTimeout(200);
  await page.keyboard.type("① 最初");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const n = (await slide()).elements.at(-1);
  assert(/^<ol data-style="circled"><li>最初<\/li><\/ol>$/.test(n.text), `numbers from 「① 」: ${n.text}`);
  await tab("校閲");
  await byTitle("入力中に (c)");
  await page.waitForSelector(".ac-dialog[open]");
  await page.uncheck('.ac-dialog input[data-ac="replace"]');
  await page.click(".ac-dialog .ac-ok");
  assert(JSON.parse(await page.evaluate(() => localStorage.getItem("hpptx-autocorrect"))).replace === false, "kept");
  await page.locator(`#stageBody .hs-obj[data-el="${n.id}"]`).dblclick();
  await page.keyboard.press("End");
  await page.keyboard.type(" (c)");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  assert(/最初 \(c\)/.test((await slide()).elements.find((x) => x.id === n.id).text), "off: typed as it is");
  await page.evaluate(() => localStorage.removeItem("hpptx-autocorrect"));
});

await step("表: 均等割り付け, セルの余白 (狭い・ユーザー設定), 文字列の方向 (縦書き)", async () => {
  await tab("挿入");
  await ribbonBtn("表");
  await page.waitForSelector(".tb-pick");
  await page.locator(".tb-pick button").nth(1 * 10 + 2).click();
  await page.waitForTimeout(400);
  await page.keyboard.type("見出し");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  const cell = async () => (await slide()).elements.find((x) => x.id === id).cells[0][0];
  assert(JSON.stringify(await page.evaluate(() => window.__hpptx.selection())) === JSON.stringify([id]), "the table is chosen (typing ended)");
  await tab("レイアウト");
  await byTitle("均等割り付け");
  await page.waitForTimeout(200);
  assert((await cell()).align === "distributed", `the cells distributed: ${JSON.stringify(await cell())}`);
  await byTitle("セルの余白");
  await menuItem("狭い");
  await page.waitForTimeout(200);
  assert(JSON.stringify((await cell()).pad) === JSON.stringify([7.37, 7.37, 7.37, 7.37]), `狭い (0.13 cm): ${JSON.stringify((await cell()).pad)}`);
  await byTitle("セルの余白");
  await menuItem("ユーザー設定の余白");
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", "0.2 0.5");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  const pad = (await cell()).pad;
  assert(Math.abs(pad[0] - 11.34) < 0.1 && Math.abs(pad[1] - 28.35) < 0.1 && Math.abs(pad[2] - pad[0]) < 0.01 && Math.abs(pad[3] - pad[1]) < 0.01, `上下 0.2・左右 0.5 cm: ${JSON.stringify(pad)}`);
  await byTitle("文字列の方向");
  await page.waitForTimeout(200);
  assert((await cell()).vertical === true, "縦書き");
  assert(await page.locator(`#stageBody .hs-obj[data-el="${id}"] .hs-cell-tx.is-vertical`).count(), "drawn vertical");
  await shot("table-cells");
});

await step("パスワードを使用して暗号化: the file asks the password, a wrong one is refused, the right one opens the show", async () => {
  await page.click(".rb-file");
  await page.locator('.rb-pop .rb-menu button:has-text("情報")').click();
  await page.waitForSelector(".fileinfo-dialog[open]");
  await page.click(".fileinfo-dialog .fi-lock");
  await page.waitForSelector(".lock-dialog[open]");
  await page.fill(".lock-dialog .lock-pw", "pass-2026");
  await page.fill(".lock-dialog .lock-again", "pass-2025");
  await page.click(".lock-dialog .lock-go");
  assert((await page.textContent(".lock-dialog .lock-msg")).includes("違います"), "the two must match");
  await page.fill(".lock-dialog .lock-again", "pass-2026");
  const [download] = await Promise.all([page.waitForEvent("download", { timeout: 120000 }), page.click(".lock-dialog .lock-go")]);
  const file = join(outDir, "format-locked.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  const words = [(await deck()).title, ...(await deck()).slides.map((x) => x.title)].filter((w) => typeof w === "string" && w.length >= 3);
  assert(html.includes('id="hs-locked"') && !html.includes("pass-2026"), "locked, the password not kept");
  assert(!words.some((w) => html.includes(w)), `nothing of the deck readable: ${words.filter((w) => html.includes(w)).join(" / ")}`);
  const viewer = await context.newPage();
  viewer.on("pageerror", (error) => errors.push(`locked pageerror: ${error.message}`));
  await viewer.goto(`file://${file}`);
  await viewer.fill("#pw", "wrong");
  await viewer.click("#go");
  await viewer.waitForFunction(() => document.getElementById("msg")?.textContent.includes("違います"), null, { timeout: 20000 });
  await viewer.fill("#pw", "pass-2026");
  await viewer.click("#go");
  await viewer.waitForSelector(".hs-player .hs-slide", { timeout: 30000 });
  assert(await viewer.locator(".hs-player .hs-slide").count(), "the show opens");
  await viewer.close();
});

await step("テーブル デザイン → 罫線の作成: ペンの太さ 3 pt・黒で外枠、内側の横罫線、枠なし", async () => {
  await tab("挿入");
  await ribbonBtn("表");
  await page.waitForSelector(".tb-pick");
  await page.locator(".tb-pick button").nth(2 * 10 + 2).click();
  await page.waitForTimeout(400);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  const cells = async () => (await slide()).elements.find((x) => x.id === id).cells;
  await tab("テーブル デザイン");
  await byTitle("罫線を引くペンの太さ");
  await menuItem("3 pt");
  await byTitle("選んだセル（なければ表全体）に罫線");
  await menuItem("外枠");
  await page.waitForTimeout(300);
  let c = await cells();
  assert(JSON.stringify(c[0][0].bt) === JSON.stringify({ c: "#000000", w: 6 }) && JSON.stringify(c[2][2].bb) === JSON.stringify({ c: "#000000", w: 6 }) && !c[1][1].bt, `外枠 at 3 pt: ${JSON.stringify(c[0][0])} ${JSON.stringify(c[1][1])}`);
  const drawn = await page.evaluate((id) => getComputedStyle(document.querySelector(`#stageBody .hs-obj[data-el="${id}"] td[data-r="0"][data-c="0"]`)).borderTopWidth, id);
  assert(parseFloat(drawn) >= 5, `drawn 6px: ${drawn}`);
  await byTitle("選んだセル（なければ表全体）に罫線");
  await menuItem("横罫線（内側）");
  await page.waitForTimeout(300);
  c = await cells();
  assert(c[0][1].bb && c[1][1].bt && c[1][1].bb && !c[1][1].bl, `inside horizontal lines: ${JSON.stringify(c[1][1])}`);
  await shot("table-borders");
  await byTitle("選んだセル（なければ表全体）に罫線");
  await menuItem("枠なし");
  await page.waitForTimeout(300);
  c = await cells();
  assert(c.flat().every((cell) => cell.merged || (cell.bt === "none" && cell.bb === "none" && cell.bl === "none" && cell.br === "none")), "枠なし");
});

await step("箇条書きと段落番号…: 開始 3, アクセント, 125%; 下線の種類 二重線; グラフの軸の書式 0〜200", async () => {
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  const r = await page.locator(".slide-wrap .hs-slide").first().boundingBox();
  await page.mouse.click(r.x + r.width * 0.3, r.y + r.height * 0.45);
  await page.waitForTimeout(200);
  await page.keyboard.type("三つ目");
  await page.keyboard.press("Enter");
  await page.keyboard.type("四つ目");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  const o = async () => (await slide()).elements.find((x) => x.id === id);
  await tab("ホーム");
  await byTitle("段落番号の種類");
  await page.click('.rb-pop .rb-list-sw[data-list-style="decimal"]');
  await page.waitForTimeout(200);
  await byTitle("段落番号の種類");
  await page.click(".rb-pop .rb-list-more");
  await page.waitForSelector(".list-dialog[open]");
  await page.fill(".list-dialog .ld-start", "3");
  await page.selectOption(".list-dialog .ld-color", "accent");
  await page.selectOption(".list-dialog .ld-size", "125");
  await page.click(".list-dialog .ld-ok");
  await page.waitForTimeout(300);
  const text = (await o()).text;
  assert(/^<ol[^>]*\bstart="3"/.test(text) && /data-mark="accent"/.test(text) && /data-msize="125"/.test(text), `the list: ${text}`);
  const marker = await page.evaluate((id) => { const li = document.querySelector(`#stageBody .hs-obj[data-el="${id}"] li`); const m = getComputedStyle(li, "::marker"); return `${m.color}|${m.fontSize}|${li.parentElement.start}`; }, id);
  const accentRgb = await page.evaluate(() => { const a = window.SlideEngine.themeColors(window.__hpptx.deck().theme).accent1; return [1, 3, 5].map((i) => parseInt(a.slice(i, i + 2), 16)).join(", "); });
  assert(marker.includes(accentRgb) && marker.endsWith("|3"), `drawn: marks in the accent, from 3 (${marker})`);
  await byTitle("下線の種類");
  await menuItem("二重線");
  await page.waitForTimeout(200);
  const u = await o();
  assert(u.underline && u.uline === "double", `underlined double: ${JSON.stringify({ underline: u.underline, uline: u.uline })}`);
  assert((await page.evaluate((id) => getComputedStyle(document.querySelector(`#stageBody .hs-obj[data-el="${id}"] .hs-obj-tx`)).textDecorationStyle, id)) === "double", "drawn double");
  // グラフ: 軸の書式 (最小値・最大値).
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="line"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const chartId = (await slide()).elements.at(-1).id;
  await tab("グラフのデザイン");
  await ribbonBtn("グラフ要素");
  await menuItem("軸の書式");
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", "0");
  await page.keyboard.press("Enter");
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", "200");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  const opts = (await slide()).elements.find((x) => x.id === chartId).chart.opts || {};
  assert(opts.axisMin === 0 && opts.axisMax === 200, `the axis: ${JSON.stringify(opts)}`);
  const ticks = await page.$$eval(`#stageBody .hs-obj[data-el="${chartId}"] .hs-tick`, (els) => els.map((el) => el.textContent));
  assert(ticks[0] === "0" && ticks.at(-1) === "200", `ticks 0–200: ${ticks}`);
});

await step("グラフの色の変更（カラフル）; 表の行の高さ・列の幅（cm）; ファイル → オプション", async () => {
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="clustered-bar"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const chartId = (await slide()).elements.at(-1).id;
  await tab("グラフのデザイン");
  await byTitle("グラフの色（テーマの色");
  await page.click('.rb-pop .tb-color-row[data-colors="colorful2"]');
  await page.waitForTimeout(300);
  assert((await slide()).elements.find((x) => x.id === chartId).chart.colors === "colorful2", "the chart's colours");
  const fill = await page.evaluate((id) => { const bar = document.querySelector(`#stageBody .hs-obj[data-el="${id}"] .hs-bar`); return bar ? getComputedStyle(bar).fill : ""; }, chartId);
  const accent2 = await page.evaluate(() => { const c = window.SlideEngine.themeColors(window.__hpptx.deck().theme).accent2; return [1, 3, 5].map((i) => parseInt(c.slice(i, i + 2), 16)).join(", "); });
  assert(fill.includes(accent2), `drawn from アクセント 2: ${fill} (expected ${accent2})`);
  // 表: the second row 2 cm high, the table taller by the difference.
  await tab("挿入");
  await ribbonBtn("表");
  await page.waitForSelector(".tb-pick");
  await page.locator(".tb-pick button").nth(2 * 10 + 1).click();
  await page.waitForTimeout(400);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const t0 = (await slide()).elements.at(-1);
  await tab("レイアウト");
  await byTitle("選んだ行（なければ全部）の高さ（cm）");
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", "2");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(300);
  const t1 = (await slide()).elements.find((x) => x.id === t0.id);
  const rowsCm = t1.rows.map((f) => Math.round((f * t1.h) / (144 / 2.54) * 10) / 10);
  assert(rowsCm.every((v) => Math.abs(v - 2) < 0.05), `every row 2 cm (none chosen: all): ${rowsCm}`);
  // ファイル → オプション: spelling on, the toolbar below the ribbon, then both back.
  await page.click(".rb-file");
  await page.locator('.rb-pop .rb-menu button:has-text("オプション…")').click();
  await page.waitForSelector(".options-dialog[open]");
  const spellBefore = await page.isChecked('.options-dialog input[data-opt="spell"]');
  await page.click('.options-dialog input[data-opt="spell"]');
  await page.selectOption('.options-dialog select[data-opt="qat"]', "below");
  await page.waitForTimeout(300);
  assert(await page.locator(".rb-qat-row").count(), "the Quick Access Toolbar below the ribbon");
  assert((await page.evaluate(() => localStorage.getItem("hpptx-spellcheck"))) === (spellBefore ? "0" : "1"), "spelling switched and kept");
  await page.selectOption('.options-dialog select[data-opt="qat"]', "off");
  await page.click('.options-dialog input[data-opt="spell"]');
  await page.click(".options-dialog .btn-primary");
  assert(!(await page.locator(".rb-qat-row").count()), "the toolbar off again");
});

await step("クイック レイアウト（グラフ）; 図形の塗りつぶし → 図・並べて表示・色に戻す", async () => {
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="clustered-bar"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const chartId = (await slide()).elements.at(-1).id;
  const chartOf = async () => (await slide()).elements.find((x) => x.id === chartId).chart;
  await tab("グラフのデザイン");
  await byTitle("グラフ要素の組み合わせ");
  assert(await page.locator(".rb-pop .tb-layout-row").count() === 7, "seven layouts, each with a picture");
  await page.click('.rb-pop .tb-layout-row[data-layout="layout5"]');
  await page.waitForTimeout(300);
  let c = await chartOf();
  assert(c.opts?.axisX && c.opts?.axisY && c.opts.labels === false, `レイアウト 5: axis titles, no data labels: ${JSON.stringify(c.opts)}`);
  await byTitle("グラフ要素の組み合わせ");
  await page.click('.rb-pop .tb-layout-row[data-layout="layout7"]');
  await page.waitForTimeout(300);
  c = await chartOf();
  assert(c.opts?.legend === "none" && c.opts.grid === false && c.opts.labels === false && !c.opts.axisX, `レイアウト 7: the chart alone: ${JSON.stringify(c.opts)}`);
  await undo();
  assert((await chartOf()).opts?.axisX, "⌘Z goes back to レイアウト 5 in one step");
  // グラフ スタイル: the whole look at once (colours, elements, bar gap), the axis titles kept.
  await byTitle("グラフのスタイル一覧");
  assert(await page.locator(".rb-pop .tb-layout-row[data-style]").count() === 8, "eight styles, each with a picture");
  assert(await page.locator(".rb-pop .tb-style-pic svg rect").count() > 20, "the pictures are drawn");
  await shot("chart-styles");
  await page.click('.rb-pop .tb-layout-row[data-style="style3"]');
  await page.waitForTimeout(300);
  c = await chartOf();
  assert(c.colors === "mono1" && c.opts?.gap === 20 && c.opts.axisX && c.opts.labels !== false, `スタイル 3: ${JSON.stringify({ colors: c.colors, opts: c.opts })}`);
  assert(await page.locator('#stageBody .hs-obj-chart[data-colors="mono1"]').count() === 1, "drawn in the monochrome set");
  await byTitle("グラフのスタイル一覧");
  assert(await page.locator('.rb-pop button.on .tb-layout-row[data-style="style3"]').count() === 1, "the style it matches is marked");
  await page.click('.rb-pop .tb-layout-row[data-style="style4"]');
  await page.waitForTimeout(300);
  c = await chartOf();
  assert(c.colors === "gray" && c.opts.gap === 200 && c.opts.labels === false && c.opts.legend === "bottom", `スタイル 4: ${JSON.stringify(c.opts)}`);
  await undo();
  assert((await chartOf()).colors === "mono1", "⌘Z goes back to スタイル 3 in one step");
  await undo();
  assert(!(await chartOf()).colors, "and one more to the layout it was");
  // 図で塗りつぶし: a rectangle filled with the logo, then tiled, then a colour again.
  await page.keyboard.press("Escape");
  await tab("挿入");
  await ribbonBtn("図形");
  await page.click('.rb-pop .rb-gallery button[title="正方形/長方形"]');
  const box = await page.locator("#stageBody .hs-slide").first().boundingBox();
  const pt = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
  await page.mouse.move(...pt(1300, 300));
  await page.mouse.down();
  await page.mouse.move(...pt(1700, 600), { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  const rect = (await slide()).elements.at(-1);
  assert(rect.kind === "shape", `the rectangle: ${rect.kind}`);
  await tab("図形の書式");
  await byTitle("図形の塗りつぶし");
  await withFiles(logo, () => page.click('.rb-pop .rb-extra[data-fill="picture"]'));
  await page.waitForFunction((id) => window.__hpptx.slide().elements.find((o) => o.id === id)?.fillImg, rect.id, { timeout: 8000 });
  let o = (await slide()).elements.find((x) => x.id === rect.id);
  assert(o.fillImg.startsWith("idb:"), `the picture kept in this browser: ${o.fillImg.slice(0, 20)}`);
  const drawn = await page.evaluate((id) => { const el = document.querySelector(`#stageBody .hs-obj[data-el="${id}"]`); const p = el?.querySelector("pattern"); return p ? { fill: el.querySelector("path").getAttribute("fill"), id: p.id, w: p.getAttribute("width"), href: p.querySelector("image").getAttribute("href") } : null; }, rect.id);
  assert(drawn && drawn.fill === `url(#${drawn.id})` && drawn.href.startsWith("blob:"), `drawn with the picture: ${JSON.stringify(drawn)}`);
  await shot("picture-fill");
  await byTitle("図形の塗りつぶし");
  await page.click('.rb-pop .rb-extra[data-fill="tile"]');
  await page.waitForTimeout(300);
  o = (await slide()).elements.find((x) => x.id === rect.id);
  assert(o.fillTile === true, "並べて表示");
  assert(await page.evaluate((id) => document.querySelector(`#stageBody .hs-obj[data-el="${id}"] pattern`)?.getAttribute("width"), rect.id) === "160", "tiles");
  await byTitle("図形の塗りつぶし");
  await page.locator(".rb-pop .rb-sw").first().click();
  await page.waitForTimeout(300);
  o = (await slide()).elements.find((x) => x.id === rect.id);
  assert(!o.fillImg && !o.fillTile, `a colour replaces the picture: ${JSON.stringify({ fillImg: o.fillImg, fillTile: o.fillTile })}`);
  await undo();
  assert((await slide()).elements.find((x) => x.id === rect.id).fillImg, "⌘Z brings the picture back");
});

await step("フォント（⇧⌘F）・段落（インデント・ぶら下げ）の設定を1回の⌘Zで; グリッドとガイド; スライドに合わせて配置; ⌘Dが移動を繰り返す", async () => {
  await page.keyboard.press("Escape");
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  const box = await page.locator("#stageBody .hs-slide").first().boundingBox();
  const pt = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
  await page.mouse.click(...pt(300, 860));
  await page.waitForTimeout(200);
  await page.keyboard.type("Seven Eleven");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  const obj = async () => (await slide()).elements.find((x) => x.id === id);
  // フォント: 太字・32pt・アクセント 1・二重下線・小型英大文字・広く, in one step.
  const accent = await page.evaluate(() => window.SlideEngine.themeColors(window.__hpptx.deck().theme).accent1);
  await page.keyboard.press("Control+Shift+F");
  await page.waitForSelector(".font-dialog[open]");
  await page.selectOption('.font-dialog select[name="スタイル"]', "bold");
  await page.fill('.font-dialog input[name="サイズ"]', "32");
  await page.click(`.font-dialog .fp-sw[data-color="${accent}"]`);
  await page.selectOption('.font-dialog select[name="下線のスタイル"]', "double");
  await page.check('.font-dialog input[name="smallCaps"]');
  await page.selectOption('.font-dialog select[name="文字間隔"]', "0.05");
  await shot("font-dialog");
  await page.click(".font-dialog .fmt-ok");
  await page.waitForTimeout(300);
  let o = await obj();
  assert(o.bold && o.fs === 64 && o.color === accent && o.underline && o.uline === "double" && o.caps === "small" && o.ls === 0.05, `the font settings: ${JSON.stringify({ bold: o.bold, fs: o.fs, color: o.color, underline: o.underline, uline: o.uline, caps: o.caps, ls: o.ls })}`);
  assert(await page.evaluate((i) => getComputedStyle(document.querySelector(`#stageBody .hs-obj[data-el="${i}"] .hs-obj-tx`)).fontVariantCaps, id) === "small-caps", "drawn in small capitals");
  await undo();
  o = await obj();
  assert(!o.bold && !o.caps && !o.ls && o.fs !== 64, `one ⌘Z takes all of the dialog back: ${JSON.stringify({ bold: o.bold, caps: o.caps, ls: o.ls, fs: o.fs })}`);
  await page.keyboard.press("Control+y");
  await page.waitForTimeout(300);
  // 段落: 中央揃え, 1 cm before the text with a 0.5 cm hanging first line, 12 pt after, 1.5 lines.
  await tab("ホーム");
  await byTitle("段落の設定");
  await page.waitForSelector(".para-dialog[open]");
  await page.selectOption('.para-dialog select[name="配置"]', "center");
  await page.fill('.para-dialog input[name="テキストの前"]', "1");
  await page.selectOption('.para-dialog select[name="最初の行"]', "hanging");
  await page.fill('.para-dialog input[name="幅"]', "0.5");
  await page.fill('.para-dialog input[name="段落後"]', "12");
  await page.selectOption('.para-dialog select[name="行間"]', "1.5");
  await page.click(".para-dialog .fmt-ok");
  await page.waitForTimeout(300);
  o = await obj();
  assert(o.align === "center" && o.lh === 1.5 && Math.abs(o.psp - 12 / 32) < 0.01, `alignment, line and paragraph spacing: ${JSON.stringify({ align: o.align, lh: o.lh, psp: o.psp })}`);
  assert(/margin-left: ?85\.0\dpx/.test(o.text) && /text-indent: ?-28\.3\dpx/.test(o.text), `1.5 cm in, the first line 0.5 cm out: ${o.text}`);
  const ind = await page.evaluate((i) => { const p = document.querySelector(`#stageBody .hs-obj[data-el="${i}"] .hs-obj-tx p`); const cs = getComputedStyle(p); return [cs.marginLeft, cs.textIndent]; }, id);
  assert(parseFloat(ind[0]) > 80 && parseFloat(ind[1]) < -25, `drawn indented: ${ind}`);
  // グリッドとガイド: 0.5 cm, snap on.
  await tab("表示");
  await byTitle("グリッドとガイドの設定");
  await page.waitForSelector(".grid-dialog[open]");
  await page.selectOption('.grid-dialog select[name="間隔"]', "0.5");
  await page.check('.grid-dialog input[name="snapGrid"]');
  await page.check('.grid-dialog input[name="grid"]');
  await page.click(".grid-dialog .fmt-ok");
  await page.waitForTimeout(300);
  const ed = await page.evaluate(() => JSON.parse(localStorage.getItem("hpptx-editor-gridStep") || "null"));
  assert(ed === 0.5, `the spacing kept: ${ed}`);
  assert(await page.locator("#stageBody .ed-grid").count(), "the grid is shown");
  await byTitle("グリッドとガイドの設定");
  await page.waitForSelector(".grid-dialog[open]");
  await page.uncheck('.grid-dialog input[name="snapGrid"]');
  await page.uncheck('.grid-dialog input[name="grid"]');
  await page.selectOption('.grid-dialog select[name="間隔"]', "0.25");
  await page.click(".grid-dialog .fmt-ok");
  await page.waitForTimeout(300);
  // ⌘D, move the copy, ⌘D again: the third lands as far again.
  await page.keyboard.press("Control+d");
  await page.waitForTimeout(300);
  const copy1 = (await slide()).elements.at(-1);
  for (let i = 0; i < 5; i += 1) await page.keyboard.press("Shift+ArrowRight");
  await page.waitForTimeout(300);
  const moved = (await slide()).elements.find((x) => x.id === copy1.id);
  await page.keyboard.press("Control+d");
  await page.waitForTimeout(300);
  const copy2 = (await slide()).elements.at(-1);
  const src = await obj();
  const [dx, dy] = [moved.x - src.x, moved.y - src.y];
  assert(Math.abs(copy2.x - moved.x - dx) < 0.6 && Math.abs(copy2.y - moved.y - dy) < 0.6, `the move repeated: source ${src.x},${src.y} copy ${moved.x},${moved.y} next ${copy2.x},${copy2.y}`);
  // 配置 → スライドに合わせて配置: two boxes both to the slide's right edge.
  await page.keyboard.press("Escape");
  await page.keyboard.press("Control+a");
  await tab("ホーム");
  await ribbonBtn("配置");
  await menuItem("スライドに合わせて配置");
  await ribbonBtn("配置");
  await menuItem("右揃え");
  await page.waitForTimeout(300);
  const all = (await slide()).elements.filter((x) => [id, copy1.id, copy2.id].includes(x.id));
  assert(all.every((x) => Math.abs(x.x + x.w - 1920) < 1), `all at the slide's right edge: ${all.map((x) => x.x + x.w)}`);
  await ribbonBtn("配置");
  await menuItem("選択したオブジェクトを揃える");
  assert(await page.evaluate(() => JSON.parse(localStorage.getItem("hpptx-editor-alignTo") || "null")) === "auto", "back to aligning with each other");
});

await step("グラフ フィルター（系列・項目を隠す）とデータ テーブル; 図のレイアウト（キャプション付きで並べる）", async () => {
  await page.keyboard.press("Escape");
  await page.locator(".film-item").nth(4).click();
  await page.waitForTimeout(300);
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="clustered-bar"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const chartId = (await slide()).elements.at(-1).id;
  const chartOf = async () => (await slide()).elements.find((x) => x.id === chartId).chart;
  const c0 = await chartOf();
  await tab("グラフのデザイン");
  await byTitle("グラフ フィルター");
  await page.locator(`.rb-pop .rb-menu button:has-text("${c0.labels[0]}")`).first().click();
  await page.waitForTimeout(300);
  let c = await chartOf();
  assert(JSON.stringify(c.opts?.hideLabels) === "[0]" && c.labels.length === c0.labels.length, `the first category hidden, its data kept: ${JSON.stringify(c.opts)}`);
  const axis = await page.evaluate((id) => document.querySelector(`#stageBody .hs-obj[data-el="${id}"]`).textContent, chartId);
  assert(!axis.includes(c0.labels[0]), `not drawn: ${axis.slice(0, 120)}`);
  await byTitle("データ ラベル・凡例・目盛線");
  await page.locator('.rb-pop .rb-menu button:has-text("表示する")').nth(2).click();
  await page.waitForTimeout(300);
  c = await chartOf();
  assert(c.opts?.table === true, `データ テーブル on: ${JSON.stringify(c.opts)}`);
  const cols = await page.evaluate((id) => document.querySelectorAll(`#stageBody .hs-obj[data-el="${id}"] .hs-chart-table thead th`).length, chartId);
  assert(cols === c0.labels.length, `a column a shown category (and the names): ${cols}`);
  await shot("chart-filter-table");
  await byTitle("グラフ フィルター");
  await page.locator('.rb-pop .rb-menu button:has-text("すべて表示")').click();
  await page.waitForTimeout(300);
  assert(!(await chartOf()).opts?.hideLabels, "all shown again");
  // 図のレイアウト: two pictures in a row with captions.
  await page.keyboard.press("Escape");
  const p1 = await insertPicture(logo);
  const p2 = await insertPicture(logo);
  await page.click("#formatTab").catch(() => {});
  await page.locator(`#formatPane .fp-sel-list li[data-id="${p1.id}"]`).click();
  await page.locator(`#formatPane .fp-sel-list li[data-id="${p2.id}"]`).click({ modifiers: ["Shift"] });
  await page.waitForTimeout(200);
  await tab("図の形式");
  await byTitle("図のレイアウト");
  await menuItem("キャプション付きの図（横に並べる）");
  await page.waitForTimeout(400);
  const list = (await slide()).elements;
  const [a, b] = [p1.id, p2.id].map((id) => list.find((o) => o.id === id));
  const caps = list.filter((o) => o.kind === "text" && o.group && o.group === a.group);
  assert(a.group && a.group === b.group && caps.length === 2, `pictures and two captions grouped: ${JSON.stringify(caps.map((o) => o.text))}`);
  assert(Math.abs(a.y - b.y) < 0.5 && Math.abs(a.w - b.w) < 0.5 && a.fit === "cover", `side by side, the same size: ${JSON.stringify([a, b].map((o) => [o.x, o.y, o.w, o.h, o.fit]))}`);
  await shot("picture-layout");
  await undo();
  assert(!(await slide()).elements.find((o) => o.id === p1.id).group, "⌘Z puts them back");
});

await step("右クリック: 図として保存（画像は PNG・図形は SVG）; リンクを開く・リンクの削除; 図とサイズのリセット", async () => {
  await page.keyboard.press("Escape");
  await page.locator(".film-item").nth(5).click();
  await page.waitForTimeout(300);
  const pic = await insertPicture(logo);
  // The picture trimmed (its left quarter), brightened, turned 90°, outlined; and a shape with a link.
  await page.evaluate((id) => {
    const s = window.__hpptx.slide();
    Object.assign(s.elements.find((o) => o.id === id), { x: 300, y: 300, w: 300, h: 300, rot: 90, crop: { l: 0.25, t: 0, r: 0, b: 0 }, bright: 0.2, stroke: "#808080", strokeW: 4 });
    s.elements.push({ id: "qaLink", kind: "shape", shape: "roundRect", x: 1000, y: 320, w: 420, h: 200, fill: "#DEEBF7", stroke: "none", text: "<p>リンク先</p>", fs: 40, color: "#1F3864", action: { type: "url", href: "https://example.com/acme" } });
  }, pic.id);
  const at = await page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide()));
  await page.locator(".film-item").nth(at + 1).click();
  await page.locator(".film-item").nth(at).click();
  await page.waitForSelector('#stageBody .hs-obj[data-el="qaLink"]');
  const rightClick = async (id, item) => {
    const r = await page.locator(`#stageBody .slide-wrap .hs-obj[data-el="${id}"]`).first().boundingBox();
    await page.mouse.click(r.x + r.width / 2, r.y + r.height / 2, { button: "right" });
    await page.locator(`.ed-menu button:has-text("${item}")`).first().click();
  };
  // The file names given to the browser (a headless browser may call a download made after a wait "download").
  await page.evaluate(() => { window.__names = []; const click = HTMLAnchorElement.prototype.click; HTMLAnchorElement.prototype.click = function () { if (this.download) window.__names.push(this.download); return click.call(this); }; });
  // 図として保存 (a picture): PNG as it shows on the slide.
  await rightClick(pic.id, "図として保存");
  await page.waitForSelector(".ps-dialog[open]");
  const [png] = await Promise.all([page.waitForEvent("download", { timeout: 20000 }), page.click(".ps-dialog .ps-go")]);
  const pngFile = join(outDir, "format-saved-picture.png");
  await png.saveAs(pngFile);
  const names = await page.evaluate(() => window.__names);
  assert(/^[^/\\]+\.png$/.test(names.at(-1)), `named after the picture: ${names}`);
  const facts = JSON.parse(execFileSync("python3", ["-c", `
from PIL import Image; import json, sys
im = Image.open(sys.argv[1]).convert("RGBA"); w, h = im.size
print(json.dumps({"w": w, "h": h, "top": im.getpixel((w // 2, int(h * .12))), "bottom": im.getpixel((w // 2, int(h * .88))), "left": im.getpixel((int(w * .12), h // 2)), "edge": im.getpixel((1, h // 2))}))`, pngFile]).toString());
  const navy = (p) => p[2] > p[0] + 40 && p[0] < 90 && p[3] > 200;
  const white = (p) => p.slice(0, 3).every((v) => v > 230);
  assert(Math.abs(facts.w - facts.h) <= 2 && facts.w >= 300 && facts.w <= 320, `the frame (with its outline): ${facts.w}×${facts.h}`);
  assert(navy(facts.top) && white(facts.bottom) && white(facts.left), `trimmed and turned (the navy square at the top): ${JSON.stringify(facts)}`);
  assert(facts.top[2] > 105, `brightened as on the slide: ${facts.top}`);
  assert(facts.edge[0] < 160 && facts.edge[3] > 200, `outlined in grey: ${facts.edge}`);
  // 図として保存 (a shape): SVG with its words.
  await rightClick("qaLink", "図として保存");
  await page.waitForSelector(".ps-dialog[open]");
  await page.selectOption(".ps-dialog .ps-format", "svg");
  const [svg] = await Promise.all([page.waitForEvent("download", { timeout: 20000 }), page.click(".ps-dialog .ps-go")]);
  const svgFile = join(outDir, "format-saved-shape.svg");
  await svg.saveAs(svgFile);
  assert(/\.svg$/.test((await page.evaluate(() => window.__names)).at(-1)), "an .svg file");
  const text = await readFile(svgFile, "utf8");
  assert(/^<\?xml/.test(text) && text.includes("<svg") && text.includes("リンク先") && /width="420"/.test(text), `an SVG of the shape: ${text.slice(0, 160)}`);
  const svgPng = await page.evaluate(async (t) => {
    const img = new Image();
    await new Promise((ok, no) => { img.onload = ok; img.onerror = no; img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(t)}`; });
    const c = document.createElement("canvas");
    c.width = img.naturalWidth; c.height = img.naturalHeight;
    const g = c.getContext("2d");
    g.drawImage(img, 0, 0);
    const d = g.getImageData(10, c.height / 2, 1, 1).data;
    return { w: c.width, h: c.height, fill: [...d] };
  }, text);
  assert(svgPng.w === 420 && svgPng.h === 200 && svgPng.fill[2] > 200 && svgPng.fill[0] < 240 && svgPng.fill[3] > 200, `the SVG draws the shape's light blue: ${JSON.stringify(svgPng)}`);
  // リンクを開く opens it in a new tab; リンクの削除 takes it off (one ⌘Z brings it back).
  await context.route("https://example.com/**", (route) => route.fulfill({ contentType: "text/html", body: "<title>acme</title>" }));
  const [popup] = await Promise.all([context.waitForEvent("page", { timeout: 10000 }), rightClick("qaLink", "リンクを開く")]);
  await popup.waitForLoadState().catch(() => {});
  assert(popup.url() === "https://example.com/acme", `the link opened in a new tab: ${popup.url()}`);
  await popup.close();
  await rightClick("qaLink", "リンクの削除");
  await page.waitForTimeout(300);
  assert(!(await slide()).elements.find((o) => o.id === "qaLink").action, "the link is gone");
  const r = await page.locator('#stageBody .slide-wrap .hs-obj[data-el="qaLink"]').boundingBox();
  await page.mouse.click(r.x + r.width / 2, r.y + r.height / 2, { button: "right" });
  assert(!(await page.locator('.ed-menu button:has-text("リンクを開く")').count()), "no リンクを開く without a link");
  await page.keyboard.press("Escape");
  await undo();
  assert((await slide()).elements.find((o) => o.id === "qaLink").action?.href === "https://example.com/acme", "⌘Z brings the link back");
  // 図とサイズのリセット: the picture's own 400 × 300, not turned or trimmed, its middle where it was (one ⌘Z).
  await rightClick(pic.id, "図とサイズのリセット");
  await page.waitForTimeout(300);
  let o = (await slide()).elements.find((x) => x.id === pic.id);
  assert(o.w === 400 && o.h === 300 && !o.rot && !o.crop && !o.bright && o.stroke == null && o.x === 250 && o.y === 300, `reset: ${JSON.stringify(o)}`);
  await undo();
  o = (await slide()).elements.find((x) => x.id === pic.id);
  assert(o.rot === 90 && o.crop && o.w === 300, "⌘Z puts it back");
  // 図の形式 → リセット ▾ → 図のリセット keeps the size.
  await page.click("#formatTab").catch(() => {});
  await page.locator(`#formatPane .fp-sel-list li[data-id="${pic.id}"]`).click();
  await tab("図の形式");
  await byTitle("図のリセット・図とサイズのリセット");
  await menuItem("図のリセット");
  await page.waitForTimeout(300);
  o = (await slide()).elements.find((x) => x.id === pic.id);
  assert(!o.crop && !o.bright && o.stroke == null && o.w === 300 && o.rot === 90, `the look reset, the size kept: ${JSON.stringify(o)}`);
  await shot("picture-save-reset");
});

// A new blank slide for the D22 steps, and helpers to point at it.
const stageAt = async (x, y) => { const b = await page.locator("#stageBody .slide-wrap .hs-slide").first().boundingBox(); return [b.x + (x * b.width) / 1920, b.y + (y * b.height) / 1080]; };
const objRightClick = async (id, item) => {
  const r = await page.locator(`#stageBody .slide-wrap .hs-obj[data-el="${id}"]`).first().boundingBox();
  await page.mouse.click(r.x + r.width / 2, r.y + r.height / 2, { button: "right" });
  await page.locator(`.ed-menu button:has-text("${item}")`).first().click();
};
const emptyRightClick = async (item) => {
  await page.mouse.click(...(await stageAt(1700, 1000)), { button: "right" });
  await page.locator(`.ed-menu button:has-text("${item}")`).first().click();
};
const redraw = async () => {
  const at = await page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide()));
  await page.locator(".film-item").nth(at - 1).click();
  await page.locator(".film-item").nth(at).click();
  await page.waitForTimeout(300);
};
const pickInPane = async (ids) => {
  await page.click("#formatTab").catch(() => {});
  for (const [i, id] of ids.entries()) await page.locator(`#formatPane .fp-sel-list li[data-id="${id}"]`).click(i ? { modifiers: ["Shift"] } : {});
  await page.waitForTimeout(200);
};
await step("描画モードのロック（右クリック）; 既定の図形・既定の線に設定; オプションで元に戻す", async () => {
  await page.keyboard.press("Escape");
  await tab("挿入");
  await ribbonBtn("新しいスライド");
  await menuItem("白紙");
  await page.waitForTimeout(400);
  await tab("挿入");
  await ribbonBtn("図形");
  assert((await page.textContent(".rb-pop .rb-gallery-hint")).includes("描画モードのロック"), "the gallery says how");
  await page.locator('.rb-pop .rb-gallery button[data-shape="ellipse"]').last().click({ button: "right" });
  for (const x of [400, 800, 1200]) { await page.mouse.click(...(await stageAt(x, 420))); await page.waitForTimeout(200); }
  let list = (await slide()).elements || [];
  assert(list.filter((o) => o.shape === "ellipse").length === 3, `three ellipses with one pick: ${list.map((o) => o.shape)}`);
  assert(await page.locator(".ed-drawing").count(), "still drawing");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  assert(!(await page.locator(".ed-drawing").count()), "Esc ends the lock");
  // The first one in orange with dark blue bold 30 px words (standard colours) becomes the deck's default shape.
  const [e1] = list.filter((o) => o.shape === "ellipse");
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((o) => o.id === id), { fill: "#ffc000", color: "#002060", bold: true, fs: 30 }), e1.id);
  await redraw();
  await objRightClick(e1.id, "既定の図形に設定");
  await page.waitForTimeout(200);
  let d = (await deck()).objectDefaults;
  assert(d?.shape?.fill === "#ffc000" && d.shape.color === "#002060" && d.shape.bold === true && d.shape.fs === 30 && !("text" in d.shape) && !("x" in d.shape), `the default shape: ${JSON.stringify(d)}`);
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[data-shape="rect"]').last().click();
  await page.mouse.click(...(await stageAt(400, 800)));
  await page.waitForTimeout(300);
  list = (await slide()).elements;
  const rect = list.at(-1);
  assert(rect.shape === "rect" && rect.fill === "#ffc000" && rect.color === "#002060" && rect.bold === true && rect.fs === 30, `a new shape takes the default: ${JSON.stringify(rect)}`);
  // A line: grey 6 px becomes the default line; a new arrow takes it (with its own arrowhead).
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[title="直線"]').first().click();
  await page.mouse.move(...(await stageAt(900, 760)));
  await page.mouse.down();
  await page.mouse.move(...(await stageAt(1300, 760)), { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  const line = (await slide()).elements.at(-1);
  assert(line.kind === "line", `a line: ${line.kind}`);
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((o) => o.id === id), { stroke: "#7f7f7f", strokeW: 6 }), line.id);
  await redraw();
  await page.mouse.click(...(await stageAt(1100, 760)), { button: "right" });
  await page.locator('.ed-menu button:has-text("既定の線に設定")').click();
  await page.waitForTimeout(200);
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[title="矢印"]').first().click();
  await page.mouse.move(...(await stageAt(900, 900)));
  await page.mouse.down();
  await page.mouse.move(...(await stageAt(1300, 900)), { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  const arrow = (await slide()).elements.at(-1);
  assert(arrow.kind === "line" && arrow.stroke === "#7f7f7f" && arrow.strokeW === 6 && arrow.tail === "triangle", `a new arrow takes the default line: ${JSON.stringify(arrow)}`);
  await shot("default-shapes");
  // ファイル → オプション: what is set, and back to the studio's own.
  await page.click(".rb-file");
  await page.locator('.rb-pop .rb-menu button:has-text("オプション…")').click();
  await page.waitForSelector(".options-dialog[open]");
  assert((await page.textContent(".options-dialog .opt-defaults")).includes("図形・線"), `the summary: ${await page.textContent(".options-dialog .opt-defaults")}`);
  await page.click(".options-dialog .opt-defaults-reset");
  await page.waitForTimeout(200);
  assert(!(await deck()).objectDefaults, "the defaults are gone");
  await page.click(".options-dialog .btn-primary");
  await undo();
  assert((await deck()).objectDefaults?.shape?.fill === "#ffc000", "⌘Z brings them back");
});

await step("図として貼り付け（2つの図形を1枚の画像に）; 2つを図として保存（SVG）; 背景の保存", async () => {
  const list = (await slide()).elements;
  const [e1, e2] = list.filter((o) => o.shape === "ellipse");
  await pickInPane([e1.id, e2.id]);
  await objRightClick(e1.id, "コピー");
  await emptyRightClick("図として貼り付け");
  await page.waitForFunction((n) => window.__hpptx.slide().elements.length > n, list.length, { timeout: 10000 });
  const pic = (await slide()).elements.at(-1);
  assert(pic.kind === "image" && /^idb:/.test(pic.src), `a picture: ${JSON.stringify(pic).slice(0, 160)}`);
  // Both ellipses (240 wide, 400 apart) in one picture.
  assert(Math.abs(pic.w - 640) <= 1 && Math.abs(pic.h - 240) <= 1, `the size of what was copied: ${pic.w}×${pic.h}`);
  const px = await page.evaluate(async (id) => {
    const img = document.querySelector(`#stageBody .hs-obj[data-el="${id}"] img`);
    await img.decode?.().catch(() => {});
    const c = document.createElement("canvas");
    c.width = img.naturalWidth; c.height = img.naturalHeight;
    const g = c.getContext("2d");
    g.drawImage(img, 0, 0);
    const at = (fx, fy) => [...g.getImageData(Math.floor(fx * c.width), Math.floor(fy * c.height), 1, 1).data];
    return { w: c.width, left: at(0.19, 0.5), right: at(0.81, 0.5), middle: at(0.5, 0.5) };
  }, pic.id);
  assert(Math.abs(px.w - pic.w * 2) <= 4, `drawn at twice the size: ${px.w}`);
  assert(px.left[0] > 190 && px.left[2] < 200 && px.left[3] > 200, `the brown one on the left: ${px.left}`);
  assert(px.right[2] > px.right[0] && px.right[3] > 200, `the blue one on the right: ${px.right}`);
  assert(px.middle[3] < 30, `see-through between them: ${px.middle}`);
  await shot("paste-as-picture");
  await undo();
  assert(!(await slide()).elements.some((o) => o.id === pic.id), "⌘Z takes the picture back");
  // 図として保存 with two chosen: one SVG holding both.
  await page.evaluate(() => { window.__names = []; const click = HTMLAnchorElement.prototype.click; HTMLAnchorElement.prototype.click = function () { if (this.download) window.__names.push(this.download); return click.call(this); }; });
  await pickInPane([e1.id, e2.id]);
  await objRightClick(e2.id, "図として保存");
  await page.waitForSelector(".ps-dialog[open]");
  assert((await page.textContent(".ps-dialog")).includes("2個"), "two objects in one picture");
  await page.selectOption(".ps-dialog .ps-format", "svg");
  const [svg] = await Promise.all([page.waitForEvent("download", { timeout: 20000 }), page.click(".ps-dialog .ps-go")]);
  const file = join(outDir, "format-saved-two.svg");
  await svg.saveAs(file);
  const text = await readFile(file, "utf8");
  const width = Number(text.match(/<svg[^>]* width="(\d+)"/)?.[1]);
  assert(Math.abs(width - 640) <= 1 && (text.match(/data-el=/g) || []).length === 2, `both in one SVG: ${width}, ${(text.match(/data-el=/g) || []).length}`);
  assert((await page.evaluate(() => window.__names)).at(-1) === "図.svg", "named 図.svg");
  // 背景の保存: the slide's background picture as a file.
  await page.keyboard.press("Escape");
  await page.evaluate(() => { window.__hpptx.slide().background = { image: "asset:storeOperations" }; });
  await redraw();
  await page.mouse.click(...(await stageAt(1700, 1000)));
  const [bg] = await Promise.all([page.waitForEvent("download", { timeout: 20000 }), emptyRightClick("背景の保存")]);
  const bgFile = join(outDir, "format-saved-background.jpg");
  await bg.saveAs(bgFile);
  assert(/^背景\.(jpg|png|webp)$/.test((await page.evaluate(() => window.__names)).at(-1)), `named 背景: ${await page.evaluate(() => window.__names)}`);
  const size = execFileSync("python3", ["-c", "from PIL import Image; import sys; print(*Image.open(sys.argv[1]).size)", bgFile]).toString().trim().split(" ").map(Number);
  assert(size[0] > 600, `the background picture itself: ${size}`);
  await page.evaluate(() => { delete window.__hpptx.slide().background; });
  await redraw();
});

await step("マウスの通過: 動作設定で「次のスライド」とサウンド; 発表中にマウスを乗せると進む", async () => {
  const rect = (await slide()).elements.find((o) => o.shape === "rect");
  await pickInPane([rect.id]);
  await objRightClick(rect.id, "リンク・動作の設定");
  await page.waitForSelector('#formatPane select[aria-label="マウスの通過時の動作"]');
  await page.selectOption('#formatPane select[aria-label="マウスの通過時の動作"]', "next");
  await page.selectOption('#formatPane select[aria-label="マウスの通過時の動作：サウンド"]', "chime");
  await page.selectOption('#formatPane select[aria-label="クリックしたときの動作：サウンド"]', "click");
  await page.click('#formatPane .rb-link-form button:has-text("設定する")');
  await page.waitForTimeout(300);
  const o = (await slide()).elements.find((x) => x.id === rect.id);
  assert(JSON.stringify(o.overAction) === JSON.stringify({ type: "next", sound: "chime" }) && JSON.stringify(o.action) === JSON.stringify({ type: "sound", sound: "click" }), `set: ${JSON.stringify([o.action, o.overAction])}`);
  await page.mouse.move(5, 5);
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(`.hs-player-stage .hs-obj[data-el="${rect.id}"][data-over="next"]`);
  await page.waitForTimeout(1200);
  const before = (await page.textContent(".hs-player-count")).trim();
  const b = await page.locator(`.hs-player-stage .hs-obj[data-el="${rect.id}"]`).boundingBox();
  await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2, { steps: 4 });
  await page.waitForFunction((was) => document.querySelector(".hs-player-count")?.textContent.trim() !== was, before, { timeout: 5000 });
  const after = (await page.textContent(".hs-player-count")).trim();
  assert(Number(after.split("/")[0]) === Number(before.split("/")[0]) + 1, `moved on: ${before} → ${after}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

const freshSlide = async () => {
  await page.keyboard.press("Escape");
  await tab("挿入");
  await ribbonBtn("新しいスライド");
  await menuItem("白紙");
  await page.waitForTimeout(400);
};
/** A paste of what Excel (text + a table) or a page puts on the clipboard, sent to the document or to a field. */
const pasteClip = (selector, { text = "", html = "" }) => page.evaluate(({ selector, text, html }) => {
  const dt = new DataTransfer();
  dt.setData("text/plain", text);
  if (html) dt.setData("text/html", html);
  (selector ? document.querySelector(selector) : document).dispatchEvent(new ClipboardEvent("paste", { clipboardData: dt, bubbles: true, cancelable: true }));
}, { selector, text, html });
await step("表の貼り付け: Excelのセル（タブ区切り）がスライドで表に、表のセルに貼ると続くセルに入る; Webの表は結合も保つ", async () => {
  await freshSlide();
  await page.mouse.click(...(await stageAt(1700, 1000)));
  await pasteClip(null, { text: "品目\t売上\t前年比\nお茶\t1,200\t98%\n水\t800\t105%\n" });
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).some((o) => o.kind === "table"), null, { timeout: 5000 });
  let t = (await slide()).elements.find((o) => o.kind === "table");
  assert(t.cells.length === 3 && t.cells[0].length === 3, `3 × 3: ${t.cells.length} × ${t.cells[0].length}`);
  const words = (c) => (c.text || "").replace(/<[^>]+>/g, "");
  assert(words(t.cells[0][0]) === "品目" && words(t.cells[1][1]) === "1,200" && t.cells[1][1].align === "right", `the cells: ${JSON.stringify(t.cells[1])}`);
  assert(await page.locator(".rb-tabs [role=tab]:text-is(\"テーブル デザイン\")").count(), "the table's tab opens");
  await shot("paste-table");
  // Into the cell under 水 / 前年比's row: 2 × 2 cells from the second row, second column on; the table grows by a column.
  await page.mouse.dblclick(...(await (async () => { const r = await page.locator("#stageBody .slide-wrap .hs-obj[data-kind='table'] td[data-r='2'][data-c='1']").first().boundingBox(); return [r.x + r.width / 2, r.y + r.height / 2]; })()));
  await page.waitForSelector(".ed-typing-tx");
  await pasteClip(".ed-typing-tx", { text: "甲\t乙\n丙\t丁" });
  await page.waitForTimeout(400);
  t = (await slide()).elements.find((o) => o.kind === "table");
  assert(t.cells[2].length === 3 && words(t.cells[2][1]) === "甲" && words(t.cells[2][2]) === "乙", `filled from the cell on: ${t.cells[2].map(words)}`);
  assert(t.cells.length === 4 && words(t.cells[3][1]) === "丙" && words(t.cells[3][2]) === "丁", `a row added: ${t.cells.map((r) => r.map(words))}`);
  assert(words(t.cells[1][1]) === "1,200" && words(t.cells[2][0]) === "水", "the others stay");
  await page.keyboard.press("Escape");
  await undo();
  assert((await slide()).elements.find((o) => o.kind === "table").cells.length === 3, "⌘Z takes the fill back in one step");
  // A web page's table: a span and a line break.
  await page.mouse.click(...(await stageAt(1700, 1000)));
  await pasteClip(null, { text: "計画\n2", html: '<table><tr><th colspan="2">計画</th></tr><tr><td>A<br>B</td><td>2</td></tr></table>' });
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).filter((o) => o.kind === "table").length === 2, null, { timeout: 5000 });
  const web = (await slide()).elements.filter((o) => o.kind === "table").at(-1);
  assert(web.cells[0][0].cs === 2 && web.cells[0][1].merged && words(web.cells[1][0]) === "AB", `the page's table: ${JSON.stringify(web.cells)}`);
  // Plain words still become a text box.
  await page.mouse.click(...(await stageAt(1700, 1000)));
  await pasteClip(null, { text: "ただの文章です" });
  await page.waitForTimeout(300);
  assert((await slide()).elements.at(-1).kind === "text", "words stay words");
});

await step("再グループ化: グループ解除したあと、ひとつ選んで右クリック → 再グループ化で元のグループに戻る", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[data-shape="ellipse"]').last().click({ button: "right" });
  for (const x of [500, 900]) { await page.mouse.click(...(await stageAt(x, 500))); await page.waitForTimeout(200); }
  await page.keyboard.press("Escape");
  const [a, b] = (await slide()).elements.filter((o) => o.shape === "ellipse");
  await pickInPane([a.id, b.id]);
  await page.keyboard.press("Control+g");
  await page.waitForTimeout(250);
  let list = (await slide()).elements;
  const group = list.find((o) => o.id === a.id).group;
  assert(group && list.find((o) => o.id === b.id).group === group, "grouped");
  await page.keyboard.press("Control+Shift+g");
  await page.waitForTimeout(250);
  assert(!(await slide()).elements.some((o) => o.group), "taken apart");
  await pickInPane([a.id]);
  await objRightClick(a.id, "再グループ化");
  await page.waitForTimeout(250);
  list = (await slide()).elements;
  const again = list.find((o) => o.id === a.id).group;
  assert(again && list.find((o) => o.id === b.id).group === again, "put together again");
  await undo();
  assert(!(await slide()).elements.some((o) => o.group), "⌘Z takes it back");
});

await step("第2軸: 複合グラフの折れ線が右側の別の目盛りで描かれる（グラフ要素 → 第2軸）", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="combo"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const id = (await slide()).elements.at(-1).id;
  await page.evaluate((id) => {
    const o = window.__hpptx.slide().elements.find((x) => x.id === id);
    o.chart.labels = ["4月", "5月", "6月", "7月"];
    o.chart.series = [{ name: "売上", values: [12000, 15000, 18000, 21000] }, { name: "前年比", values: [96, 102, 108, 111] }];
  }, id);
  await redraw();
  const ys = () => page.$$eval(`#stageBody .hs-obj[data-el="${id}"] circle.hs-mark`, (els) => els.map((c) => Number(c.getAttribute("cy"))));
  const flat = await ys();
  assert(flat.length === 4 && Math.max(...flat) - Math.min(...flat) < 3, `the line sits on the bars' scale: ${flat}`);
  await pickInPane([id]);
  await tab("グラフのデザイン");
  await byTitle("データ ラベル・凡例・目盛線");
  await page.locator('.rb-pop .rb-menu button:has-text("折れ線を第2軸にする")').click();
  await page.waitForTimeout(400);
  assert((await slide()).elements.find((o) => o.id === id).chart.opts?.axis2 === true, "opts.axis2");
  const climbs = await ys();
  assert(Math.max(...climbs) - Math.min(...climbs) > 40, `a scale of its own: ${climbs}`);
  const ticks = await page.$$eval(`#stageBody .hs-obj[data-el="${id}"] .hs-tick2`, (els) => els.map((t) => t.textContent));
  assert(ticks.length === 5 && Number(ticks.at(-1)) >= 111, `the right scale: ${ticks}`);
  await shot("second-axis");
  await byTitle("データ ラベル・凡例・目盛線");
  await page.locator('.rb-pop .rb-menu button:has-text("折れ線を第2軸にする")').click();
  await page.waitForTimeout(300);
  assert(!(await slide()).elements.find((o) => o.id === id).chart.opts?.axis2, "off again");
});

await step("スライドのサイズ → スライド番号の開始番号: 表紙を0にすると次のページが1になる（元に戻せる）", async () => {
  await page.keyboard.press("Escape");
  await page.locator(".film-item").nth(1).click();
  await page.waitForTimeout(400);
  const pageNo = () => page.evaluate(() => String(Number((document.querySelector("#stageBody .slide-wrap .hs-foot .hs-page")?.textContent || "").split("/")[0].trim())));
  const before = await pageNo();
  assert(before === "2", `the second slide is page 2: ${before}`);
  await tab("デザイン");
  await byTitle("スライドのサイズ");
  await page.waitForSelector(".size-dialog[open]");
  await page.fill('.size-dialog input[name="firstNumber"]', "0");
  await page.click(".size-dialog .fmt-ok");
  await page.waitForTimeout(500);
  assert((await deck()).firstNumber === 0, `firstNumber: ${(await deck()).firstNumber}`);
  assert((await pageNo()) === "1", `now page 1: ${await pageNo()}`);
  await undo();
  assert((await pageNo()) === "2" && !(await deck()).firstNumber && (await deck()).firstNumber !== 0, "⌘Z puts the numbering back");
});

await step("線の書式: 複合線（二重線）・線端（丸）・結合点（面取り）が図形の枠線に効く（二重線は本当に中が透ける）", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[data-shape="rect"]').last().click();
  await page.mouse.click(...(await stageAt(700, 500)));
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((o) => o.id === id), { stroke: "#1f3864", strokeW: 12 }), id);
  await redraw();
  await pickInPane([id]);
  await tab("図形の書式");
  await byTitle("図形の枠線");
  await page.locator('.rb-pop [data-lines="cmpd"] button[data-cmpd="dbl"]').click();
  await byTitle("図形の枠線");
  await page.locator('.rb-pop [data-lines="cap"] button[data-cap="round"]').click();
  await byTitle("図形の枠線");
  await page.locator('.rb-pop [data-lines="join"] button[data-join="bevel"]').click();
  await page.waitForTimeout(300);
  const o = (await slide()).elements.find((x) => x.id === id);
  assert(o.cmpd === "dbl" && o.cap === "round" && o.join === "bevel", `kept: ${JSON.stringify([o.cmpd, o.cap, o.join])}`);
  const drawn = await page.evaluate((id) => { const el = document.querySelector(`#stageBody .slide-wrap .hs-obj[data-el="${id}"]`); const p = [...el.querySelectorAll("svg path")].find((x) => x.getAttribute("mask")); return { mask: Boolean(el.querySelector("mask")), cap: p?.getAttribute("stroke-linecap"), join: p?.getAttribute("stroke-linejoin") }; }, id);
  assert(drawn.mask && drawn.cap === "round" && drawn.join === "bevel", `drawn: ${JSON.stringify(drawn)}`);
  // The double line really has a see-through middle: across the left edge, the colour shows twice with the slide between.
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((x) => x.id === id), { strokeW: 36 }), id);
  await redraw();
  await page.mouse.click(...(await stageAt(1700, 1000)));
  const r = await page.locator(`#stageBody .slide-wrap .hs-obj[data-el="${id}"]`).first().boundingBox();
  const strip = join(outDir, "format-compound-strip.png");
  await page.screenshot({ path: strip, clip: { x: Math.max(0, r.x - 24), y: r.y + r.height / 2 - 4, width: 48, height: 8 } });
  const bands = JSON.parse(execFileSync("python3", ["-c", `
from PIL import Image; import json, sys
im = Image.open(sys.argv[1]).convert("RGB"); w, h = im.size
row = [im.getpixel((x, h // 2)) for x in range(w)]
navy = lambda p: p[2] > p[0] + 25 and p[0] < 120
runs = 0; prev = False
for p in row:
    cur = navy(p)
    if cur and not prev: runs += 1
    prev = cur
print(json.dumps({"runs": runs, "navy": sum(navy(p) for p in row)}))`, strip]).toString());
  assert(bands.runs === 2, `two bands of colour with the slide between: ${JSON.stringify(bands)}`);
  await shot("compound-line");
  await undo();
  assert(!(await slide()).elements.find((x) => x.id === id).join, "⌘Z takes the last change back");
});

await step("塗りつぶし → グラデーション: アクセント 1 → 明るく・左方向で左が薄く右が濃い（テーマの色）、単色を選ぶと外れる", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[data-shape="rect"]').last().click();
  await page.mouse.click(...(await stageAt(900, 500)));
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((o) => o.id === id), { w: 800, h: 300, x: 500, y: 400 }), id);
  await redraw();
  await pickInPane([id]);
  await tab("図形の書式");
  await byTitle("図形の塗りつぶし");
  await page.locator('.rb-pop [data-fill="gradient"]').click();
  await page.waitForSelector(".gradient-dialog[open]");
  await page.selectOption('.gradient-dialog select[aria-label="色の組み合わせ"]', "accent1");
  await page.selectOption('.gradient-dialog select[aria-label="方向"]', "left");
  await page.click(".gradient-dialog .fmt-ok");
  await page.waitForTimeout(400);
  const o = (await slide()).elements.find((x) => x.id === id);
  const [ga, gb] = await page.evaluate(() => { const a = window.SlideEngine.themeColors(window.__hpptx.deck().theme).accent1; return [a, window.SlideEngine.lumColor(a, 40000, 60000)]; });
  assert(o.gradient?.angle === 180 && o.gradient.stops[0].color === ga && o.gradient.stops[1].color === gb && o.fill === ga, `the gradient: ${JSON.stringify([o.gradient, o.fill])}`);
  await page.mouse.click(...(await stageAt(1700, 1000)));
  const r = await page.locator(`#stageBody .slide-wrap .hs-obj[data-el="${id}"]`).first().boundingBox();
  const strip = join(outDir, "format-gradient-strip.png");
  await page.screenshot({ path: strip, clip: { x: r.x + 4, y: r.y + r.height / 2, width: r.width - 8, height: 4 } });
  const lum = JSON.parse(execFileSync("python3", ["-c", `
from PIL import Image; import json, sys
im = Image.open(sys.argv[1]).convert("RGB"); w, h = im.size
print(json.dumps({"left": sum(im.getpixel((6, 1))) / 3, "right": sum(im.getpixel((w - 7, 1))) / 3}))`, strip]).toString());
  assert(lum.left > lum.right + 15, `light on the left, deeper on the right: ${JSON.stringify(lum)}`);
  await shot("gradient");
  await pickInPane([id]);
  await byTitle("図形の塗りつぶし");
  await page.locator('.rb-pop .rb-sw').nth(2).click();
  await page.waitForTimeout(300);
  assert(!(await slide()).elements.find((x) => x.id === id).gradient, "a plain colour takes the gradient off");
  await undo();
  assert((await slide()).elements.find((x) => x.id === id).gradient, "⌘Z brings it back");
});

await step("図のスタイル（クイック スタイル）: 二重線の枠（黒）、標準で戻る; 文字カウント", async () => {
  await freshSlide();
  const pic = await insertPicture(logo);
  await pickInPane([pic.id]);
  await tab("図の形式");
  await byTitle("図のスタイル：枠線・角丸");
  await page.locator('.rb-pop [data-picture-style="二重線の枠（黒）"]').click();
  await page.waitForTimeout(300);
  let o = (await slide()).elements.find((x) => x.id === pic.id);
  assert(!o.mask && o.stroke === "#000000" && o.cmpd === "dbl" && o.strokeW === 18, `the style: ${JSON.stringify([o.mask, o.stroke, o.cmpd, o.strokeW])}`);
  assert(await page.evaluate((id) => Boolean(document.querySelector(`#stageBody .slide-wrap .hs-obj[data-el="${id}"] mask`)), pic.id), "the double line is drawn");
  await shot("picture-style");
  await byTitle("図のスタイル：枠線・角丸");
  await page.locator('.rb-pop [data-picture-style="標準（枠なし）"]').click();
  await page.waitForTimeout(300);
  o = (await slide()).elements.find((x) => x.id === pic.id);
  assert(!o.mask && !o.stroke && !o.cmpd, `back to plain: ${JSON.stringify([o.mask, o.stroke, o.cmpd])}`);
  assert(o.w === pic.w && o.src === pic.src, "the size and the picture stay");
  // 文字カウント.
  await page.keyboard.press("Escape");
  await tab("校閲");
  await ribbonBtn("文字");
  await page.waitForSelector(".wc-dialog[open]");
  const count = (name) => page.textContent(`.wc-dialog [data-count="${name}"]`).then((t) => Number(t.replace(/,/g, "")));
  const total = (await deck()).slides.length;
  assert((await count("スライド")) === total, `slides: ${await count("スライド")} of ${total}`);
  const withoutNotes = await count("段落");
  await page.check('.wc-dialog label:has-text("スピーカー ノートを含める") input');
  await page.waitForTimeout(200);
  assert((await count("段落")) >= withoutNotes, "the notes add paragraphs");
  assert((await count("文字数（スペースを含める）")) >= (await count("文字数（スペースを含めない）")) && (await count("文字数（スペースを含めない）")) > 100, "characters counted");
  await shot("word-count");
  await page.keyboard.press("Escape");
});

await step("グラフの数値の書式（表示単位 万・小数1桁・後ろに円）と誤差範囲（固定値）", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="line"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const id = (await slide()).elements.at(-1).id;
  await page.evaluate((id) => {
    const o = window.__hpptx.slide().elements.find((x) => x.id === id);
    o.chart.labels = ["4月", "5月", "6月"];
    o.chart.series = [{ name: "売上", values: [12000, 15000, 24000] }];
  }, id);
  await redraw();
  await pickInPane([id]);
  await tab("グラフのデザイン");
  const texts = (sel) => page.$$eval(`#stageBody .slide-wrap .hs-obj[data-el="${id}"] ${sel}`, (els) => els.map((t) => t.textContent));
  await byTitle("データ ラベル・凡例・目盛線");
  await page.locator('.rb-pop .rb-menu button:has-text("数値の書式（桁数・表示単位・記号）")').click();
  await page.waitForSelector(".nf-dialog[open]");
  await page.selectOption('.nf-dialog select[name="表示単位"]', "10000");
  await page.selectOption('.nf-dialog select[name="小数点以下の桁数"]', "1");
  await page.fill('.nf-dialog input[name="数値の後ろの記号"]', "円");
  await page.click(".nf-dialog .nf-ok");
  await page.waitForTimeout(400);
  const nf = (await slide()).elements.find((o) => o.id === id).chart.opts?.numFmt;
  assert(nf?.scale === 10000 && nf.decimals === 1 && nf.suffix === "円", `kept: ${JSON.stringify(nf)}`);
  const labels = await texts(".hs-val");
  assert(labels.includes("2.4万円"), `the data label: ${labels}`);
  const ticks = (await texts(".hs-tick")).filter((t) => /\d/.test(t));
  assert(ticks.length >= 4 && ticks.every((t) => !t.includes("円") && /万$|^0/.test(t)), `the axis: ${ticks}`);
  await shot("number-format");
  // 誤差範囲: ±2000.
  await byTitle("データ ラベル・凡例・目盛線");
  await page.locator('.rb-pop .rb-menu button:has-text("誤差範囲を付ける")').click();
  await page.fill("#askInput", "1");
  await page.keyboard.press("Enter");
  await page.waitForFunction(() => document.querySelector("#askTitle")?.textContent.includes("固定値"));
  await page.fill("#askInput", "2000");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(400);
  assert(JSON.stringify((await slide()).elements.find((o) => o.id === id).chart.opts?.errorBars) === JSON.stringify({ type: "fixed", amount: 2000 }), "errorBars kept");
  const count = await page.$$eval(`#stageBody .slide-wrap .hs-obj[data-el="${id}"] line.hs-err`, (els) => els.length);
  assert(count === 9, `a bar and two caps on each of three marks: ${count}`);
  await shot("error-bars");
  await byTitle("データ ラベル・凡例・目盛線");
  await page.locator('.rb-pop .rb-menu button:has-text("誤差範囲をなくす")').click();
  await byTitle("データ ラベル・凡例・目盛線");
  await page.locator('.rb-pop .rb-menu button:has-text("数値の書式をもとに戻す")').click();
  await page.waitForTimeout(300);
  const back = (await slide()).elements.find((o) => o.id === id).chart;
  assert(!back.opts?.errorBars && !back.opts?.numFmt, `both off again: ${JSON.stringify(back.opts)}`);
  assert((await texts(".hs-val")).includes("24,000"), "numbers as before");
});

await step("グラフの書式設定（折れ線の滑らかさ・マーカー・近似曲線・目盛間隔・項目の逆順、円の角度・穴・切り出し、棒の間隔と重なり）", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="line"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const id = (await slide()).elements.at(-1).id;
  const setData = async (type, series) => {
    await page.evaluate(({ id, type, series }) => {
      const o = window.__hpptx.slide().elements.find((x) => x.id === id);
      o.chart.type = type;
      o.chart.labels = ["4月", "5月", "6月", "7月"];
      o.chart.series = series;
    }, { id, type, series });
    await redraw();
    await pickInPane([id]);
  };
  const chartOf = async () => (await slide()).elements.find((o) => o.id === id).chart;
  const q = (sel) => `#stageBody .slide-wrap .hs-obj[data-el="${id}"] ${sel}`;
  const openFormat = async () => {
    await tab("グラフのデザイン");
    await byTitle("データ ラベル・凡例・目盛線");
    await page.locator('.rb-pop .rb-menu button:has-text("グラフの書式設定")').click();
    await page.waitForSelector(".cf-dialog[open]");
  };
  const cf = ".cf-dialog[open]";
  // A line: smooth, square marks, an exponential trendline with its equation, a scale every 5,000, the months turned round.
  await setData("line", [{ name: "売上", values: [12000, 15000, 24000, 41000] }]);
  await openFormat();
  assert((await page.$$(`${cf} fieldset.cf-section`)).length === 4, "a line shows データ ラベル・折れ線・軸・近似曲線 (no bars, no pie)");
  assert(!(await page.$(`${cf} input[name="gap"]`)) && !(await page.$(`${cf} input[name="angle"]`)), "the other kinds' fields are not offered");
  await page.check(`${cf} input[name="smooth"]`);
  await page.selectOption(`${cf} select[name="marker"]`, "square");
  await page.fill(`${cf} input[name="step"]`, "5000");
  await page.selectOption(`${cf} select[name="trend"]`, "exp");
  await page.check(`${cf} input[name="eq"]`);
  await page.check(`${cf} input[name="r2"]`);
  await page.check(`${cf} input[name="reverse"]`);
  await page.click(".cf-dialog .cf-ok");
  await page.waitForTimeout(500);
  let opts = (await chartOf()).opts;
  assert(opts?.smooth === true && opts.marker === "square" && opts.axisStep === 5000 && opts.trend === "exp" && opts.trendEq === true && opts.trendR2 === true && opts.reverse === true, `kept: ${JSON.stringify(opts)}`);
  assert(/ C/.test((await page.getAttribute(q("path.hs-draw"), "d")) || ""), "the line is a curve");
  assert((await page.$$(q("polygon.hs-marker"))).length === 4, "a square on each point");
  assert((await page.$$(q(".hs-trend"))).length === 1, "the trendline");
  const eq = await page.textContent(q(".hs-trend-eq"));
  assert(/^y = .*e\^.*x/.test(eq) && eq.includes("R² = "), `the equation: ${eq}`);
  const ticks = (await page.$$eval(q(".hs-tick"), (els) => els.map((t) => Number(t.textContent.replace(/,/g, ""))))).filter(Number.isFinite);
  assert(ticks.length >= 3 && ticks.every((v, i) => i === 0 || Math.abs(v - ticks[i - 1] - 5000) < 1e-6), `a scale every 5,000: ${ticks}`);
  const months = await page.$$eval(q("svg text:not(.hs-tick):not(.hs-val):not(.hs-trend-eq)"), (els) => els.map((t) => t.textContent));
  assert(months.join() === "7月,6月,5月,4月", `the months turned round: ${months}`);
  await shot("chart-format-line");
  // The menu names the kind now; one undo takes the whole dialog back.
  await tab("グラフのデザイン");
  await byTitle("データ ラベル・凡例・目盛線");
  assert(await page.locator('.rb-pop .rb-menu button:has-text("近似曲線（指数）…")').count() === 1, "the menu names the trendline's kind");
  await page.keyboard.press("Escape");
  await undo();
  opts = (await chartOf()).opts;
  assert(!opts?.smooth && !opts?.trend && !opts?.reverse && !opts?.axisStep, `one undo undoes the dialog: ${JSON.stringify(opts)}`);

  // A donut: the first slice turned 90°, a thicker ring, the slices pulled out.
  await setData("donut", [{ name: "構成比", values: [50, 30, 12, 8] }]);
  const ring = async () => Number(await page.getAttribute(q(".hs-arc"), "stroke-width"));
  const before = await ring();
  await openFormat();
  assert((await page.$$(`${cf} fieldset.cf-section`)).length === 2 && await page.$(`${cf} input[name="hole"]`), "a donut shows ドーナツ・軸 (the hole too)");
  await page.fill(`${cf} input[name="angle"]`, "90");
  await page.fill(`${cf} input[name="explode"]`, "10");
  await page.fill(`${cf} input[name="hole"]`, "40");
  await page.click(".cf-dialog .cf-ok");
  await page.waitForTimeout(500);
  opts = (await chartOf()).opts;
  assert(opts?.angle === 90 && opts.explode === 10 && opts.hole === 40, `kept: ${JSON.stringify(opts)}`);
  const turns = await page.$$eval(q(".hs-arc"), (els) => els.map((a) => a.getAttribute("transform")));
  assert(turns.length === 4 && turns.every((t) => t.startsWith("translate(")) && turns[0].includes("rotate(0 "), `pulled out and turned: ${turns[0]}`);
  assert((await ring()) > before, "a smaller hole: a thicker ring");
  await shot("chart-format-donut");

  // Clustered bars: nearly no gap, the series overlapping by half.
  await setData("clustered-bar", [{ name: "今期", values: [3, 5, 4, 6] }, { name: "前期", values: [2, 4, 3, 5] }]);
  const xs = async () => page.$$eval(q(".hs-bar"), (els) => els.map((b) => Number(b.getAttribute("d").match(/^M([-\d.]+),/)[1])));
  const widthBefore = (await xs())[1] - (await xs())[0];
  await openFormat();
  assert(await page.$(`${cf} input[name="gap"]`) && await page.$(`${cf} input[name="overlap"]`), "bars offer the gap and the overlap");
  await page.fill(`${cf} input[name="gap"]`, "20");
  await page.fill(`${cf} input[name="overlap"]`, "50");
  await page.click(".cf-dialog .cf-ok");
  await page.waitForTimeout(500);
  opts = (await chartOf()).opts;
  assert(opts?.gap === 20 && opts.overlap === 50, `kept: ${JSON.stringify(opts)}`);
  const after = await xs();
  assert(after[1] - after[0] > 0 && after[1] - after[0] < widthBefore, `the series closer together: ${widthBefore} → ${after[1] - after[0]}`);
  await shot("chart-format-bars");
  // A chart with nothing to format keeps the menu entry off.
  await setData("waterfall", [{ name: "利益", values: [100, 25, -10, 115] }]);
  await tab("グラフのデザイン");
  await byTitle("データ ラベル・凡例・目盛線");
  assert(await page.locator('.rb-pop .rb-menu button:has-text("グラフの書式設定")').isDisabled(), "disabled for a waterfall");
  await page.keyboard.press("Escape");
});

await step("グラフのデータ ラベル（すべての点・折れ線のラベル位置・円のスライスのラベル）", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="line"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const id = (await slide()).elements.at(-1).id;
  const setData = async (type, series) => {
    await page.evaluate(({ id, type, series }) => {
      const o = window.__hpptx.slide().elements.find((x) => x.id === id);
      o.chart.type = type;
      o.chart.labels = ["4月", "5月", "6月", "7月"];
      o.chart.series = series;
    }, { id, type, series });
    await redraw();
    await pickInPane([id]);
  };
  const chartOf = async () => (await slide()).elements.find((o) => o.id === id).chart;
  const q = (sel) => `#stageBody .slide-wrap .hs-obj[data-el="${id}"] ${sel}`;
  const cf = ".cf-dialog[open]";
  const openFormat = async () => {
    await tab("グラフのデザイン");
    await byTitle("データ ラベル・凡例・目盛線");
    await page.locator('.rb-pop .rb-menu button:has-text("グラフの書式設定")').click();
    await page.waitForSelector(cf);
  };
  // A line: only the last point and the highest are written; "すべての点" writes them all, here below the points.
  await setData("line", [{ name: "売上", values: [12, 30, 22, 41] }]);
  const written = async () => (await page.$$eval(q("text.hs-val"), (els) => els.map((t) => t.textContent)));
  assert((await written()).length === 1, `the usual: ${await written()}`);
  await openFormat();
  await page.check(`${cf} input[name="labelAll"]`);
  await page.selectOption(`${cf} select[name="labelPos"]`, "below");
  await page.click(".cf-dialog .cf-ok");
  await page.waitForTimeout(500);
  let opts = (await chartOf()).opts;
  assert(opts?.labelAll === true && opts.labelPos === "below", `kept: ${JSON.stringify(opts)}`);
  assert((await written()).join() === "12,30,22,41", `every point: ${await written()}`);
  const ys = await page.$$eval(q("text.hs-val"), (els) => els.map((t) => Number(t.getAttribute("y"))));
  const dots = await page.$$eval(q("circle[fill]"), (els) => els.map((c) => Number(c.getAttribute("cy"))).filter(Boolean));
  assert(ys.every((y, i) => y > dots[i]), `below the points: ${ys} vs ${dots}`);
  await shot("chart-labels-line");

  // A donut: names and shares on the slices, on leader lines; the list at the right gives way.
  await setData("donut", [{ name: "構成比", values: [50, 30, 12, 8] }]);
  assert((await page.$$(q("rect"))).length === 4, "the list at the right first");
  await openFormat();
  await page.check(`${cf} input[name="sl-category"]`);
  await page.check(`${cf} input[name="sl-percent"]`);
  await page.click(".cf-dialog .cf-ok");
  await page.waitForTimeout(500);
  opts = (await chartOf()).opts;
  assert(JSON.stringify(opts?.sliceLabels) === JSON.stringify(["category", "percent"]), `kept: ${JSON.stringify(opts)}`);
  const slices = await page.$$eval(q(".hs-slice-label"), (els) => els.map((t) => [...t.querySelectorAll("tspan")].map((s) => s.textContent).join("/")));
  assert(slices.join() === "4月/50%,5月/30%,6月/12%,7月/8%", `a label each: ${slices}`);
  assert((await page.$$(q(".hs-leader"))).length === 4 && (await page.$$(q("rect"))).length === 0, "leader lines, and no list");
  await shot("chart-labels-donut");
  // Values only: the names are not on the slices, so the list comes back.
  await openFormat();
  await page.uncheck(`${cf} input[name="sl-category"]`);
  await page.uncheck(`${cf} input[name="sl-percent"]`);
  await page.check(`${cf} input[name="sl-value"]`);
  await page.click(".cf-dialog .cf-ok");
  await page.waitForTimeout(500);
  assert((await page.$$(q("rect"))).length === 4 && (await page.$$(q(".hs-slice-label"))).length === 4, "the list stays; the values on the slices");
  await undo();
  assert((await page.$$(q(".hs-slice-label"))).length === 4 && (await page.$$(q("rect"))).length === 0, "one undo goes back to the names");
});

await step("検索・置換（大文字小文字・全角半角・単語・このスライドだけ・見つかった場所の一覧・1回のUndo）", async () => {
  await freshSlide();
  await page.mouse.click(...(await stageAt(1700, 1000)));
  await pasteClip(null, { text: "DXとＤＸとdxとADXの話" });
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).some((o) => o.kind === "text"), null, { timeout: 5000 });
  await page.keyboard.press("Escape");
  await redraw();
  const here = await page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide()));
  const words = async () => ((await slide()).elements.find((o) => o.kind === "text")?.text || "").replace(/<[^>]+>/g, "");
  const open = async () => {
    await page.keyboard.press("Escape");
    await page.keyboard.press("Control+f");
    await page.waitForSelector("#replaceDialog[open]");
  };
  const found = async () => page.textContent("#findCount");
  await open();
  // The options are what they were last time in this browser; start from the usual ones.
  for (const id of ["#findCase", "#findWord", "#findWidth"]) if (await page.isChecked(id)) await page.uncheck(id);
  await page.selectOption("#findScope", "slide");
  await page.fill("#findInput", "dx");
  assert(/4か所見つかりました/.test(await found()), `case and width alike: ${await found()}`);
  assert((await page.$$("#findHits .find-hit")).length >= 1, "the list of where");
  const hit = await page.textContent("#findHits .find-hit");
  assert(hit.includes(`${here + 1}枚目`) && hit.includes("DX") && hit.includes("4か所"), `the slide and the words round it: ${hit}`);
  await page.check("#findCase");
  assert(/1か所見つかりました/.test(await found()), `大文字と小文字を区別する: ${await found()}`);
  await page.uncheck("#findCase");
  await page.check("#findWidth");
  assert(/3か所見つかりました/.test(await found()), `全角と半角を区別する (ＤＸ is not DX): ${await found()}`);
  await page.uncheck("#findWidth");
  await page.check("#findWord");
  assert(/3か所見つかりました/.test(await found()), `完全に一致する単語だけ (ADX is not one): ${await found()}`);
  // The options stay for next time.
  await page.keyboard.press("Escape");
  await open();
  assert(await page.isChecked("#findWord"), "the option stays");
  await page.uncheck("#findWord");
  // Only this slide: another slide with the same word is left alone.
  await page.selectOption("#findScope", "slide");
  assert(/4か所見つかりました/.test(await found()), `このスライドだけ: ${await found()}`);
  await page.fill("#replaceInput", "DX推進");
  await page.click("#replaceAllBtn");
  await page.waitForTimeout(400);
  assert((await words()) === "DX推進とDX推進とDX推進とADX推進の話", `all four replaced as plain text: ${await words()}`);
  await undo();
  assert((await words()) === "DXとＤＸとdxとADXの話", `one undo: ${await words()}`);
  // Clicking a hit goes to the slide.
  await page.keyboard.press("Control+Home");
  await page.waitForTimeout(300);
  await open();
  await page.selectOption("#findScope", "deck");
  await page.fill("#findInput", "ＡＤＸ");
  await page.click("#findHits .find-hit");
  await page.waitForTimeout(400);
  assert(!(await page.isVisible("#replaceDialog[open]")), "the dialog closes");
  assert(await page.evaluate((n) => window.__hpptx.slide() === window.__hpptx.deck().slides[n], here), "and the slide with the words is on the stage");
});

await step("繰り返し（F4・やり直す操作がないときの ⌘Y）: 直前の書式を次に選んだ部品へ、1回の⌘Zで戻る", async () => {
  await freshSlide();
  const place = async (x, y) => {
    await tab("挿入");
    await ribbonBtn("図形");
    await page.locator('.rb-pop .rb-gallery button[data-shape="rect"]').last().click();
    await page.mouse.click(...(await stageAt(x, y)));
    await page.waitForTimeout(300);
    return (await slide()).elements.at(-1).id;
  };
  const a = await place(500, 500);
  const b = await place(1500, 500);
  assert(a !== b, "two rectangles");
  const look = async (id, key) => (await slide()).elements.find((o) => o.id === id)[key];
  const before = await look(b, "fill");
  // A's fill; then B takes it with F4.
  await pickInPane([a]);
  await tab("図形の書式");
  await byTitle("図形の塗りつぶし");
  await page.locator(".rb-pop .rb-sw").nth(3).click();
  await page.waitForTimeout(300);
  const picked = await look(a, "fill");
  assert(picked && picked !== before, `A's new fill: ${picked} (B ${before})`);
  await pickInPane([b]);
  await page.keyboard.press("F4");
  await page.waitForTimeout(300);
  assert((await look(b, "fill")) === picked, `F4 gave B the same fill: ${await look(b, "fill")}`);
  // A's outline colour; B again, this time with ⌘Y (nothing to redo).
  await pickInPane([a]);
  await tab("図形の書式");
  await byTitle("図形の枠線");
  await page.locator(".rb-pop .rb-sw").nth(2).click();
  await page.waitForTimeout(300);
  const line = await look(a, "stroke");
  await pickInPane([b]);
  await page.keyboard.press("Control+y");
  await page.waitForTimeout(300);
  assert((await look(b, "stroke")) === line, `⌘Y gave B the same outline: ${await look(b, "stroke")} / ${line}`);
  assert((await look(b, "fill")) === picked, "and kept the fill");
  await undo();
  assert((await look(b, "stroke")) !== line && (await look(b, "fill")) === picked, "one ⌘Z takes back only the repeat");
  // With something to redo, ⌘Y redoes.
  await page.keyboard.press("Control+y");
  await page.waitForTimeout(300);
  assert((await look(b, "stroke")) === line, "⌘Y redoes what was undone");
});

await step("PowerPoint のファンクション キー: ⇧F9 でグリッド線、F7 で表記ゆれチェック", async () => {
  await freshSlide();
  await page.keyboard.press("Escape");
  const grid = () => page.$$("#stageBody .ed-grid").then((l) => l.length);
  const before = await grid();
  await page.keyboard.press("Shift+F9");
  await page.waitForTimeout(300);
  assert((await grid()) !== before, `⇧F9 toggles the grid: ${before} → ${await grid()}`);
  await page.keyboard.press("Shift+F9");
  await page.waitForTimeout(300);
  assert((await grid()) === before, "and back");
  await page.keyboard.press("F7");
  await page.waitForSelector(".proof-dialog[open]", { timeout: 4000 });
  await page.keyboard.press("Escape");
});

await step("図形の効果（影・反射・光彩・ぼかし）: メニューとオプション、影はブランドの指摘、1回の⌘Zで戻る", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[data-shape="rect"]').last().click();
  await page.mouse.click(...(await stageAt(900, 500)));
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((o) => o.id === id), { w: 800, h: 300, x: 500, y: 400 }), id);
  await redraw();
  await pickInPane([id]);
  const obj = () => page.evaluate((id) => { const o = window.__hpptx.slide().elements.find((x) => x.id === id); return { shadow: o.shadow, reflect: o.reflect, glow: o.glow, soft: o.soft }; }, id);
  const q = (sel) => `#stageBody .slide-wrap .hs-obj[data-el="${id}"] ${sel}`;
  await tab("図形の書式");
  // 影 → 右下
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("影：右下")').click();
  await page.waitForTimeout(400);
  let fx = await obj();
  assert(fx.shadow && fx.shadow.dx === 8 && fx.shadow.dy === 8, `shadow kept: ${JSON.stringify(fx.shadow)}`);
  assert(await page.getAttribute(q(""), "data-shadow") === "1", "marked on the object");
  assert(/drop-shadow/.test((await page.getAttribute(q(".hs-obj-rot"), "style")) || ""), "drawn");
  // 反射 中 + 光彩 8pt + ぼかし 5pt
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("反射：中")').click();
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("光彩：8 pt")').click();
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("ぼかし：5 pt")').click();
  await page.waitForTimeout(400);
  fx = await obj();
  assert(fx.reflect?.size === 0.5 && fx.glow?.r === 16 && fx.soft === 10, `three more: ${JSON.stringify(fx)}`);
  const style = (await page.getAttribute(q(".hs-obj-rot"), "style")) || "";
  assert(/box-reflect/.test(style) && (style.match(/drop-shadow/g) || []).length === 3, `reflection and glow drawn: ${style}`);
  assert(/mask-image/.test((await page.getAttribute(q("svg.hs-obj-geom"), "style")) || ""), "the soft edge masks the shape");
  await shot("effects");
  // 効果のオプション: 影の距離と角度、光彩の色
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("効果のオプション")').click();
  await page.waitForSelector(".fx-dialog[open]");
  await page.fill('.fx-dialog [name="shadowDistance"]', "5");
  await page.fill('.fx-dialog [name="shadowAngle"]', "90");
  await page.selectOption('.fx-dialog [name="glowColor"]', "#ffc000");
  await page.uncheck('.fx-dialog [name="reflectOn"]');
  await page.click(".fx-dialog .fmt-ok");
  await page.waitForTimeout(400);
  fx = await obj();
  assert(fx.shadow.dx === 0 && fx.shadow.dy === 10 && fx.glow.color === "#ffc000" && !fx.reflect, `the dialog's numbers: ${JSON.stringify(fx)}`);
  // 1回の⌘Zでダイアログの変更だけが戻る
  await undo();
  fx = await obj();
  assert(fx.reflect && fx.shadow.dx === 8, `one undo: ${JSON.stringify(fx)}`);
  // すべての効果をなくす
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("すべての効果をなくす")').click();
  await page.waitForTimeout(300);
  fx = await obj();
  assert(!fx.shadow && !fx.reflect && !fx.glow && !fx.soft, `all gone: ${JSON.stringify(fx)}`);
  assert(await page.getAttribute(q(""), "data-shadow") === null, "and the mark with them");
});

await step("パターンの塗りつぶし: 図形の塗りつぶし → パターン…（一覧から選び前景・背景はテーマの色）、描画、単色に戻す、1回の⌘Zで戻る", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[data-shape="rect"]').last().click();
  await page.mouse.click(...(await stageAt(900, 500)));
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((o) => o.id === id), { w: 800, h: 300, x: 500, y: 400 }), id);
  await redraw();
  await pickInPane([id]);
  const obj = () => page.evaluate((id) => { const o = window.__hpptx.slide().elements.find((x) => x.id === id); return { fill: o.fill, pattern: o.pattern || null }; }, id);
  await tab("図形の書式");
  await byTitle("図形の塗りつぶし");
  await page.click('.rb-pop .rb-extra[data-fill="pattern"]');
  await page.waitForSelector(".pattern-dialog[open]");
  assert(await page.locator(".pattern-dialog .pat-tile").count() >= 34, "the kinds are drawn small to choose from");
  await page.click('.pattern-dialog .pat-tile[data-pattern="dkDnDiag"]');
  const [pfg, pbg] = await page.evaluate(() => { const col = window.SlideEngine.paletteSheet(window.__hpptx.deck().theme).columns[4]; return [col.hex, col.variants[0].hex]; });
  await page.selectOption('.pattern-dialog select[name="前景"]', pfg);
  await page.selectOption('.pattern-dialog select[name="背景"]', pbg);
  await page.click(".pattern-dialog .fmt-ok");
  await page.waitForTimeout(400);
  let o = await obj();
  assert(o.pattern?.kind === "dkDnDiag" && o.pattern.fg === pfg && o.pattern.bg === pbg && o.fill === pbg, `a pattern: ${JSON.stringify(o)}`);
  const drawn = await page.evaluate((id) => { const el = document.querySelector(`#stageBody .hs-obj[data-el="${id}"]`); const p = el?.querySelector("pattern"); return p ? { fill: el.querySelector("svg > path").getAttribute("fill"), id: p.id, bg: p.querySelector("rect").getAttribute("fill"), fg: p.querySelector("path").getAttribute("fill") } : null; }, id);
  assert(drawn && drawn.fill === `url(#${drawn.id})` && drawn.fg === pfg && drawn.bg === pbg, `drawn as a pattern: ${JSON.stringify(drawn)}`);
  await shot("pattern");
  // The dialog shows what the shape has; a plain colour takes the pattern off, and ⌘Z brings it back in one step.
  await byTitle("図形の塗りつぶし");
  await page.click('.rb-pop .rb-extra[data-fill="pattern"]');
  await page.waitForSelector(".pattern-dialog[open]");
  assert(await page.locator('.pattern-dialog .pat-tile.on[data-pattern="dkDnDiag"]').count() === 1, "the pattern it has is marked");
  await page.keyboard.press("Escape");
  await byTitle("図形の塗りつぶし");
  await page.locator(".rb-pop .rb-sw").nth(3).click();
  await page.waitForTimeout(300);
  assert((await obj()).pattern === null, "a plain colour replaces the pattern");
  await undo();
  assert((await obj()).pattern?.kind === "dkDnDiag", "⌘Z brings the pattern back");
  await byTitle("図形の塗りつぶし");
  await page.click('.rb-pop .rb-extra[data-fill="nopattern"]');
  await page.waitForTimeout(300);
  o = await obj();
  assert(o.pattern === null && o.fill === pbg, `パターンを外す keeps the background as the colour: ${JSON.stringify(o)}`);
});

await step("3-D 回転: 図形の効果のギャラリー（右を奥に・等角図法）とオプション（X・Y・遠近）、描画のCSS、1回の⌘Zで戻る", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("図形");
  await page.locator('.rb-pop .rb-gallery button[data-shape="rect"]').last().click();
  await page.mouse.click(...(await stageAt(900, 500)));
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((o) => o.id === id), { w: 800, h: 300, x: 500, y: 400 }), id);
  await redraw();
  await pickInPane([id]);
  const rot3d = () => page.evaluate((id) => window.__hpptx.slide().elements.find((x) => x.id === id).rot3d || null, id);
  const css = () => page.getAttribute(`#stageBody .slide-wrap .hs-obj[data-el="${id}"] .hs-obj-rot`, "style");
  await tab("図形の書式");
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("3-D：右を奥に")').click();
  await page.waitForTimeout(400);
  let r = await rot3d();
  assert(r && r.x === 0 && r.y === 28, `右を奥に: ${JSON.stringify(r)}`);
  assert(/perspective\([\d.]+px\) rotateX\(0deg\) rotateY\(28deg\)/.test((await css()) || ""), `drawn turned: ${await css()}`);
  await shot("rot3d");
  await byTitle("図形の効果");
  assert(await page.locator('.rb-pop .rb-menu button.on:has-text("3-D：右を奥に")').count() === 1, "the preset it is shows as chosen");
  await page.locator('.rb-pop .rb-menu button:has-text("3-D：等角図法：左上")').click();
  await page.waitForTimeout(300);
  r = await rot3d();
  assert(r && r.x === 35 && r.y === -45 && r.p === 0 && !/perspective/.test((await css()) || ""), `等角図法 is flat: ${JSON.stringify(r)} ${await css()}`);
  // オプション: X 10, Y -10, 遠近 60.
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("効果のオプション")').click();
  await page.waitForSelector(".fx-dialog[open]");
  await page.fill('.fx-dialog [name="rotX"]', "10");
  await page.fill('.fx-dialog [name="rotY"]', "-10");
  await page.fill('.fx-dialog [name="rotP"]', "60");
  await page.click(".fx-dialog .fmt-ok");
  await page.waitForTimeout(300);
  r = await rot3d();
  assert(r && r.x === 10 && r.y === -10 && r.p === 60, `the dialog: ${JSON.stringify(r)}`);
  await undo();
  r = await rot3d();
  assert(r && r.x === 35 && r.p === 0, `one undo goes back to 等角図法: ${JSON.stringify(r)}`);
  await byTitle("図形の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("3-D 回転なし")').click();
  await page.waitForTimeout(300);
  assert((await rot3d()) === null && !/rotateX/.test((await css()) || ""), "no turn at all");
});

await step("パレート図: 順不同の値が大きい順の棒になり、累積比率の線が右軸（0〜100%）に出る・書式設定の棒の間隔・1回の⌘Zで戻る", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="pareto"]');
  await page.waitForTimeout(400);
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  const id = (await slide()).elements.at(-1).id;
  const chartOf = async () => (await slide()).elements.find((x) => x.id === id).chart;
  let c = await chartOf();
  assert(c.type === "pareto" && c.labels.length === 6 && c.series.length === 1, `a pareto with its sample: ${JSON.stringify({ type: c.type, n: c.labels.length })}`);
  const texts = () => page.$$eval(`#stageBody .hs-obj[data-el="${id}"] .hs-chart svg text`, (list) => list.map((t) => t.textContent));
  let t = await texts();
  const order = ["弁当", "パン", "おにぎり", "惣菜", "サラダ", "デザート"];
  assert(JSON.stringify(t.filter((x) => order.includes(x))) === JSON.stringify(order), `the bars come largest first: ${t.join(",")}`);
  const right = await page.$$eval(`#stageBody .hs-obj[data-el="${id}"] .hs-tick2`, (list) => list.map((x) => x.textContent));
  assert(right.join(",") === "0%,25%,50%,75%,100%", `the right axis is in percent: ${right}`);
  assert(t.includes("100%"), "the line ends at 100%");
  await shot("pareto");
  // The values stay as typed (the order is only drawn); the gap is the bars' own.
  assert(JSON.stringify(c.labels) === JSON.stringify(["おにぎり", "弁当", "パン", "サラダ", "デザート", "惣菜"]), "the data keeps the order it was typed in");
  await tab("グラフのデザイン");
  await byTitle("データ ラベル・凡例・目盛線");
  await page.locator('.rb-pop .rb-menu button:has-text("グラフの書式設定")').click();
  await page.waitForSelector(".cf-dialog[open]");
  await page.fill('.cf-dialog input[name="gap"]', "40");
  await page.click(".cf-dialog .cf-ok");
  await page.waitForTimeout(300);
  c = await chartOf();
  assert(c.opts?.gap === 40 && !c.opts.axis2, `the gap is kept: ${JSON.stringify(c.opts)}`);
  await undo();
  assert(!(await chartOf()).opts?.gap, "⌘Z takes the gap back in one step");
});

await step("表の並べ替え・フィルター: 昇順・降順（数式は行といっしょに動く）、条件に合う行だけを表示、解除、1回の⌘Zで戻る", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("表");
  await page.waitForSelector(".tb-pick");
  await page.click('.tb-pick-more[data-sheet="insert"]');
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).some((o) => o.kind === "table" && o.sheet), null, { timeout: 5000 });
  const table = async () => (await slide()).elements.filter((o) => o.kind === "table").at(-1);
  const words = (o, r, c) => (o.cells[r][c].text || "").replace(/<[^>]+>/g, "");
  let o = await table();
  assert(words(o, 1, 0) === "店舗" && words(o, 2, 0) === "宅配", "店舗 (1,200) above 宅配 (480) to begin with");
  // B2 is being typed in: 昇順 sorts the data rows by column B (the header and the total row stay).
  await byTitle("選んだセルの列の昇順");
  await page.waitForFunction(() => /宅配/.test((window.__hpptx.slide().elements.find((x) => x.kind === "table").cells[1][0].text || "")), null, { timeout: 5000 });
  o = await table();
  assert(words(o, 1, 0) === "宅配" && words(o, 2, 0) === "店舗" && words(o, 3, 0) === "合計" && words(o, 0, 0) === "項目", "宅配 first, the header and the total stay");
  assert(o.cells[1][3].f === "=SUM(B2:C2)" && words(o, 1, 3) === "1,040", `the formula moved with its row: ${o.cells[1][3].f} ${words(o, 1, 3)}`);
  assert(words(o, 3, 1) === "1,680", "the total still adds both rows");
  await shot("sheet-sorted");
  await page.keyboard.press("Escape");
  await undo();
  assert(words(await table(), 1, 0) === "店舗", "⌘Z puts the rows back in one step");
  await page.locator("#stageBody .hs-obj td[data-r='1'][data-c='1']").first().dblclick();
  await page.waitForSelector(".tb-fx:not([disabled])");
  await byTitle("選んだセルの列の降順");
  await page.waitForTimeout(400);
  assert(words(await table(), 1, 0) === "店舗", "already in 降順: nothing moves");
  // フィルター: column B larger than 1000 → 宅配 is hidden, not deleted.
  await byTitle("ある列の条件に合う行だけを表示する");
  await page.waitForSelector(".filter-dialog[open]");
  await page.selectOption('.filter-dialog [name="col"]', "1");
  await page.selectOption('.filter-dialog [name="op"]', "gt");
  await page.fill('.filter-dialog [name="value"]', "1000");
  await page.click(".filter-dialog .fmt-ok");
  await page.waitForFunction(() => (window.__hpptx.slide().elements.find((x) => x.kind === "table").hide || []).length === 1, null, { timeout: 5000 });
  o = await table();
  assert(JSON.stringify(o.hide) === "[2]" && o.filter.op === "gt" && o.cells.length === 4, `row 3 (宅配) hidden, all four rows kept: ${JSON.stringify({ hide: o.hide, filter: o.filter })}`);
  const trs = await page.$$eval(`#stageBody .hs-obj[data-kind="table"] tr`, (list) => list.map((tr) => tr.getAttribute("data-r")));
  assert(trs.join(",") === "0,1,3", `the hidden row is not drawn: ${trs}`);
  await shot("sheet-filtered");
  assert(await page.locator('.rb-btn[title^="ある列の条件に合う行だけを表示する"][aria-pressed="true"]').count() === 1, "the button shows a filter is on");
  // The sum still counts the hidden row (as Excel's SUM does), and the rows left are not squeezed.
  assert(words(o, 3, 1) === "1,680", "SUM counts hidden rows too");
  // 解除.
  await byTitle("フィルターを解除");
  await page.waitForFunction(() => !(window.__hpptx.slide().elements.find((x) => x.kind === "table").filter), null, { timeout: 5000 });
  o = await table();
  assert(!o.hide && o.cells.length === 4, "every row is back");
  assert((await page.$$eval(`#stageBody .hs-obj[data-kind="table"] tr`, (list) => list.length)) === 4, "and drawn");
});

await step("段落のタブ設定: Tab キーでタブを入れる・段落ダイアログでタブ位置（左・右）を決める・描画で位置にそろう・1回の⌘Zで戻る", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  await page.mouse.click(...(await stageAt(300, 400)));
  await page.waitForTimeout(250);
  // At the start of a paragraph Tab indents; after words it types a tab character.
  await page.keyboard.type("品名");
  await page.keyboard.press("Tab");
  await page.keyboard.type("数量");
  await page.keyboard.press("Tab");
  await page.keyboard.type("金額");
  await page.waitForTimeout(200);
  assert(await page.evaluate(() => document.querySelector(".ed-typing-tx")?.textContent === "品名\t数量\t金額"), `tabs typed: ${JSON.stringify(await page.evaluate(() => document.querySelector(".ed-typing-tx")?.textContent))}`);
  assert(await page.locator(".ed-typing-tx.ed-has-tab").count() === 1, "the box shows its tabs while it is edited");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  const obj = () => page.evaluate((id) => window.__hpptx.slide().elements.find((x) => x.id === id), id);
  let o = await obj();
  assert(/品名\t数量\t金額/.test(o.text), `the tab characters are kept: ${JSON.stringify(o.text)}`);
  const widths = () => page.$$eval(`#stageBody .hs-obj[data-el="${id}"] .hs-tabbox`, (list) => list.map((b) => b.getAttribute("style") || "(natural)"));
  assert((await widths()).join("|") === "width: 144px; text-align: left|width: 144px; text-align: left|(natural)", `an inch to a tab to begin with: ${await widths()}`);
  // 段落 → タブ位置: 左 5 cm, 右 15 cm.
  await tab("ホーム");
  await byTitle("段落の設定");
  await page.waitForSelector(".para-dialog[open]");
  await page.fill('.para-dialog input[name="tabs"]', "左 5, 右 15");
  await page.click(".para-dialog .fmt-ok");
  await page.waitForTimeout(400);
  o = await obj();
  assert(/data-tabs="l\d+(\.\d+)?,r\d+(\.\d+)?"/.test(o.text), `the stops are on the paragraph: ${o.text}`);
  const stops = await page.evaluate((id) => window.SlideEngine.parseTabs(/data-tabs="([^"]*)"/.exec(window.__hpptx.slide().elements.find((x) => x.id === id).text)[1]), id);
  assert(Math.abs(stops[0].x - 5 * 144 / 2.54) < 1 && Math.abs(stops[1].x - 15 * 144 / 2.54) < 1, `5 cm and 15 cm: ${JSON.stringify(stops)}`);
  // A box wide enough for the stops (words past the box's right edge wrap, as they do in PowerPoint).
  await page.evaluate((id) => Object.assign(window.__hpptx.slide().elements.find((x) => x.id === id), { w: 1500, x: 200 }), id);
  await redraw();
  await pickInPane([id]);
  const drawn = await widths();
  assert(drawn.length === 3 && /text-align: right/.test(drawn[2]) && /text-align: left/.test(drawn[0]), `boxes at the stops: ${drawn}`);
  assert(Math.abs(parseFloat(drawn[0].match(/width: ([\d.]+)px/)[1]) - stops[0].x) < 0.5, "the first word reaches the left stop");
  await shot("tabs");
  // The dialog shows what is set; ⌘Z takes the stops back in one step.
  await byTitle("段落の設定");
  await page.waitForSelector(".para-dialog[open]");
  assert(/^左 5(\.0+)?, 右 15/.test(await page.inputValue('.para-dialog input[name="tabs"]')), `shown again: ${await page.inputValue('.para-dialog input[name="tabs"]')}`);
  await page.keyboard.press("Escape");
  await undo();
  o = await obj();
  assert(!/data-tabs/.test(o.text) && /品名\t数量\t金額/.test(o.text), "⌘Z: the stops go, the words stay");
});

await step("文字の効果（文字の影・光彩）: ホームの「文字の影」・図形の書式の文字の効果・オプション、影はブランドの指摘、1回の⌘Zで戻る", async () => {
  await freshSlide();
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  await page.mouse.click(...(await stageAt(500, 400)));
  await page.waitForTimeout(250);
  await page.keyboard.type("Shadow and Glow");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const id = (await slide()).elements.at(-1).id;
  const obj = () => page.evaluate((id) => { const o = window.__hpptx.slide().elements.find((x) => x.id === id); return { tshadow: o.tshadow, tglow: o.tglow, shadow: o.shadow, glow: o.glow }; }, id);
  const q = (sel) => `#stageBody .slide-wrap .hs-obj[data-el="${id}"] ${sel}`;
  const textStyle = async () => (await page.getAttribute(q(".hs-obj-tx"), "style")) || "";
  // ホーム → 文字の影: on, then off.
  await tab("ホーム");
  await byTitle("文字の影");
  await page.waitForTimeout(400);
  let fx = await obj();
  assert(fx.tshadow && fx.tshadow.dx === 3 && fx.tshadow.dy === 3 && !fx.shadow, `the letters' shadow, not the box's: ${JSON.stringify(fx)}`);
  assert(await page.getAttribute(q(""), "data-tshadow") === "1", "marked on the object");
  assert(/text-shadow:[^;]*3px 3px 4px/.test(await textStyle()), `drawn on the words: ${await textStyle()}`);
  assert(!/drop-shadow/.test((await page.getAttribute(q(".hs-obj-rot"), "style")) || ""), "the box has no shadow");
  assert(await page.locator('.rb-btn[title^="文字の影"][aria-pressed="true"]').count() === 1, "the button shows it is on");
  await byTitle("文字の影");
  await page.waitForTimeout(300);
  assert(!(await obj()).tshadow && await page.getAttribute(q(""), "data-tshadow") === null, "off again, and the mark with it");
  // 図形の書式 → 文字の効果: 光彩 5 pt and 影 下.
  await tab("図形の書式");
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("光彩：5 pt")').click();
  await page.waitForTimeout(300);
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("影：下")').click();
  await page.waitForTimeout(400);
  fx = await obj();
  assert(fx.tglow?.r === 10 && fx.tshadow?.dx === 0 && fx.tshadow?.dy === 3, `a glow and a shadow: ${JSON.stringify(fx)}`);
  assert((await textStyle()).match(/rgba\(0, ?112, ?192/g)?.length === 2, `the halo twice: ${await textStyle()}`);
  await shot("text-effects");
  // 文字の効果のオプション: 影の距離と角度、光彩の色。
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("文字の効果のオプション")').click();
  await page.waitForSelector(".tfx-dialog[open]");
  await page.fill('.tfx-dialog [name="shadowDistance"]', "5");
  await page.fill('.tfx-dialog [name="shadowAngle"]', "90");
  await page.selectOption('.tfx-dialog [name="glowColor"]', "#ffc000");
  await page.click(".tfx-dialog .fmt-ok");
  await page.waitForTimeout(400);
  fx = await obj();
  assert(fx.tshadow.dx === 0 && fx.tshadow.dy === 10 && fx.tglow.color === "#ffc000", `the dialog's numbers: ${JSON.stringify(fx)}`);
  await undo();
  fx = await obj();
  assert(fx.tshadow.dy === 3 && fx.tglow.color === "#0070c0", `one undo: ${JSON.stringify(fx)}`);
  // 文字の輪郭: 色と太さ（青・1.5 pt）。
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("輪郭：青")').click();
  await page.waitForTimeout(300);
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("輪郭の太さ：1.5 pt")').click();
  await page.waitForTimeout(300);
  let outline = await page.evaluate((id) => window.__hpptx.slide().elements.find((x) => x.id === id).toutline, id);
  assert(outline?.w === 3 && outline.color === "#0070c0", `an outline of 1.5 pt in blue: ${JSON.stringify(outline)}`);
  assert(/-webkit-text-stroke:[^;]*3px/.test(await textStyle()), `drawn round the letters: ${await textStyle()}`);
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("文字の効果のオプション")').click();
  await page.waitForSelector(".tfx-dialog[open]");
  await page.fill('.tfx-dialog [name="outlineWidth"]', "2");
  await page.selectOption('.tfx-dialog [name="outlineColor"]', "#808080");
  await page.click(".tfx-dialog .fmt-ok");
  await page.waitForTimeout(300);
  outline = await page.evaluate((id) => window.__hpptx.slide().elements.find((x) => x.id === id).toutline, id);
  assert(outline?.w === 4 && outline.color === "#808080", `the dialog: ${JSON.stringify(outline)}`);
  await undo();
  outline = await page.evaluate((id) => window.__hpptx.slide().elements.find((x) => x.id === id).toutline, id);
  assert(outline?.w === 3 && outline.color === "#0070c0", `one undo: ${JSON.stringify(outline)}`);
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("輪郭なし")').click();
  await page.waitForTimeout(300);
  assert(!(await page.evaluate((id) => window.__hpptx.slide().elements.find((x) => x.id === id).toutline, id)), "the outline is gone");
  // 光彩だけならブランドの指摘は出ない。すべての文字の効果をなくす。
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("影なし")').click();
  await page.waitForTimeout(300);
  assert(!(await obj()).tshadow && await page.getAttribute(q(""), "data-tshadow") === null, "a glow alone puts no shadow on the letters");
  await byTitle("文字の効果");
  await page.locator('.rb-pop .rb-menu button:has-text("すべての文字の効果をなくす")').click();
  await page.waitForTimeout(300);
  fx = await obj();
  assert(!fx.tshadow && !fx.tglow && !/text-shadow/.test(await textStyle()), `all gone: ${JSON.stringify(fx)}`);
});

await step("スライド マスター表示: 全スライド共通の部品を置く・背景グラフィックを表示しない・検索・Undo", async () => {
  await freshSlide();
  const here = await page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide()));
  const count = async () => page.evaluate(() => (window.__hpptx.deck().masterObjects || []).length);
  const own = async () => page.evaluate(() => (window.__hpptx.slide().elements || []).length);
  const before = await own();
  await tab("表示");
  await byTitle("スライド マスター表示");
  await page.waitForSelector(".callout-master");
  assert(await page.evaluate(() => window.__hpptx.master()), "the studio is in the master view");
  assert(/すべてのスライド/.test(await page.textContent(".callout-master")), "the message says what it is");
  // A text box pasted here goes to the master, not to the slide.
  await page.mouse.click(...(await stageAt(1700, 1000)));
  await pasteClip(null, { text: "マスターの文字Z" });
  await page.waitForFunction(() => (window.__hpptx.deck().masterObjects || []).length === 1, null, { timeout: 5000 });
  assert((await own()) === before, "the slide's own objects are not touched");
  assert(await page.evaluate(() => window.__hpptx.deck().masterObjects[0].text.includes("マスターの文字Z")), "the words are in the master");
  await shot("slide-master");
  // Close; the words are on this slide and on the thumbnails of the others.
  await page.click(".btn-master-close");
  await page.waitForTimeout(500);
  assert(!(await page.$(".callout-master")), "closed");
  assert((await page.$$eval("#stageBody .hs-master-layer", (els) => els.map((e) => e.textContent).join(""))).includes("マスターの文字Z"), "drawn under this slide");
  const thumbs = await page.$$eval("#filmstrip .hs-master-layer", (els) => els.length);
  assert(thumbs >= 2, `the thumbnails of the other slides show it too: ${thumbs}`);
  // 背景グラフィックを表示しない
  await tab("デザイン");
  await byTitle("スライド マスターに置いた図形");
  await page.waitForTimeout(400);
  assert(await page.evaluate(() => window.__hpptx.slide().hideMaster === true) && !(await page.$("#stageBody .hs-master-layer")), "hidden on this slide");
  await byTitle("スライド マスターに置いた図形");
  await page.waitForTimeout(400);
  assert(!(await page.evaluate(() => window.__hpptx.slide().hideMaster)) && (await page.$("#stageBody .hs-master-layer")), "and back");
  // 検索: the master's words are found, and a click takes you to the master.
  await page.keyboard.press("Control+f");
  await page.waitForSelector("#replaceDialog[open]");
  await page.selectOption("#findScope", "deck");
  await page.fill("#findInput", "マスターの文字Z");
  assert(/1か所/.test(await page.textContent("#findCount")), `found: ${await page.textContent("#findCount")}`);
  assert((await page.textContent("#findHits .find-hit")).includes("マスター"), "listed as the master");
  await page.click("#findHits .find-hit");
  await page.waitForSelector(".callout-master");
  // Undo inside the master view takes the words away; opening another slide leaves the view.
  // (the two 背景グラフィック toggles above are undo steps of their own, then the paste)
  for (let i = 0; i < 4 && (await count()) > 0; i += 1) await undo();
  assert((await count()) === 0, `undone: ${await count()}`);
  assert(await page.evaluate(() => window.__hpptx.master()), "still in the master view after undo");
  for (let i = 0; i < 4 && (await count()) < 1; i += 1) { await page.keyboard.press("Control+y"); await page.waitForTimeout(300); }
  assert((await count()) === 1, "redone");
  await page.click("#filmstrip .film-item >> nth=0");
  await page.waitForTimeout(400);
  assert(!(await page.$(".callout-master")) && !(await page.evaluate(() => window.__hpptx.master())), "choosing a slide leaves the master view");
  assert(await page.evaluate((i) => window.__hpptx.deck().slides.length > i, here), "the deck is whole");
});

await step("類義語（⇧F7・校閲）: 選んだ言葉の言い換えが出て、選ぶと入れ替わる; 文字を選ばないときはコピー; 辞書にない言葉", async () => {
  await freshSlide();
  await page.mouse.click(...(await stageAt(1700, 1000)));
  await pasteClip(null, { text: "最大の課題は人手不足です" });
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).some((o) => o.kind === "text"), null, { timeout: 5000 });
  const words = async () => ((await slide()).elements.find((o) => o.kind === "text")?.text || "").replace(/<[^>]+>/g, "");
  // Type into the text box and select 課題.
  await page.keyboard.press("Enter");
  await page.waitForSelector('#stageBody [contenteditable="true"]');
  await page.evaluate(() => {
    const box = document.querySelector('#stageBody [contenteditable="true"]');
    const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      const at = n.data.indexOf("課題");
      if (at >= 0) { box.focus(); window.getSelection().setBaseAndExtent(n, at, n, at + 2); return; }
    }
  });
  await page.waitForTimeout(200);
  await page.keyboard.press("Shift+F7");
  await page.waitForSelector(".syn-dialog[open]");
  assert((await page.inputValue('.syn-dialog [name="word"]')) === "課題", "the selected word is looked up");
  const chips = await page.$$eval(".syn-dialog .syn-word", (els) => els.map((e) => e.textContent));
  assert(chips.includes("問題") && chips.includes("論点"), `its synonyms: ${chips}`);
  assert(await page.isDisabled(".syn-dialog .syn-ai"), "the AI button waits for Codex");
  await shot("synonyms");
  await page.click('.syn-dialog .syn-word:text-is("問題")');
  await page.waitForTimeout(300);
  assert(!(await page.$(".syn-dialog")), "the dialog closes");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  assert((await words()) === "最大の問題は人手不足です", `the pick took its place: ${await words()}`);
  // From the ribbon with nothing selected: type a word, pick = copy.
  await page.keyboard.press("Escape");
  await tab("校閲");
  await ribbonBtn("類義語");
  await page.waitForSelector(".syn-dialog[open]");
  await page.fill('.syn-dialog [name="word"]', "強化する");
  assert((await page.$$eval(".syn-dialog .syn-word", (els) => els.map((e) => e.textContent))).includes("拡充"), "an ending is taken off");
  await page.fill('.syn-dialog [name="word"]', "ＫＰＩ");
  assert(/辞書に見つかりません/.test(await page.textContent(".syn-dialog .syn-status")), "an unknown word says so");
  await page.keyboard.press("Escape");
});

await step("Excel スプレッドシート: 数式つきの表（挿入・数式バー・オートSUM・桁区切り・Excel/CSVの保存と読み込み）", async () => {
  await freshSlide();
  const table = async () => (await slide()).elements.filter((o) => o.kind === "table").at(-1);
  const words = (o, r, c) => (o.cells[r][c].text || "").replace(/<[^>]+>/g, "");
  await tab("挿入");
  await ribbonBtn("表");
  await page.waitForSelector(".tb-pick");
  await page.click('.tb-pick-more[data-sheet="insert"]');
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).some((o) => o.kind === "table" && o.sheet), null, { timeout: 5000 });
  let o = await table();
  assert(words(o, 1, 3) === "2,550" && words(o, 3, 3) === "3,590", `the sums are worked out: ${words(o, 1, 3)} ${words(o, 3, 3)}`);
  assert(o.cells[1][3].f === "=SUM(B2:C2)", "the formula is kept");
  // Typing starts in B2: the formula bar names the cell and shows its words.
  await page.waitForSelector(".tb-fx:not([disabled])");
  assert((await page.textContent(".tb-fname")) === "B2" && (await page.inputValue(".tb-fx")) === "1,200", "the bar shows B2");
  await shot("sheet");
  // Changing B2 in the bar re-works every sum; Enter goes down to B3.
  await page.fill(".tb-fx", "2,000");
  await page.press(".tb-fx", "Enter");
  await page.waitForFunction(() => { const t = window.__hpptx.slide().elements.find((x) => x.kind === "table"); return /3,350/.test(t.cells[1][3].text || ""); }, null, { timeout: 5000 });
  o = await table();
  assert(words(o, 3, 3) === "4,390", `the total follows: ${words(o, 3, 3)}`);
  assert((await page.textContent(".tb-fname")) === "B3", "Enter moved down to B3");
  // A formula typed in the cell itself; while it is edited the cell shows the formula.
  await page.keyboard.type("=B2*2");
  assert((await page.inputValue(".tb-fx")) === "=B2*2", "typing in the cell shows in the bar");
  await page.keyboard.press("Enter");
  await page.waitForFunction(() => window.__hpptx.slide().elements.find((x) => x.kind === "table").cells[2][1].f === "=B2*2", null, { timeout: 5000 });
  o = await table();
  assert(words(o, 2, 1) === "4000", `=B2*2 is 4000: ${words(o, 2, 1)}`);
  assert((await page.textContent(".tb-fname")) === "B4" && (await page.inputValue(".tb-fx")) === "=SUM(B2:B3)", "B4 shows its formula while it is edited");
  // 0 で割る: the error shows in the cell.
  await page.fill(".tb-fx", "=1/0");
  await page.press(".tb-fx", "Enter");
  await page.waitForFunction(() => /#DIV\/0!/.test(window.__hpptx.slide().elements.find((x) => x.kind === "table").cells[3][1].text || ""), null, { timeout: 5000 });
  // 関数の挿入 puts a function in the cell being edited.
  await page.keyboard.press("Escape");
  await page.locator("#stageBody .hs-obj td[data-r='3'][data-c='2']").first().dblclick();
  await page.waitForSelector(".tb-fx:not([disabled])");
  await page.fill(".tb-fx", "=");
  await page.click('.tb-fbar .rb-drop');
  await menuItem("AVERAGE");
  await page.waitForTimeout(200);
  assert(/^=AVERAGE\($/.test(await page.inputValue(".tb-fx")), `the function is typed: ${await page.inputValue(".tb-fx")}`);
  await page.keyboard.type("C2:C3)");
  await page.keyboard.press("Enter");
  await page.waitForFunction(() => window.__hpptx.slide().elements.find((x) => x.kind === "table").cells[3][2].f === "=AVERAGE(C2:C3)", null, { timeout: 5000 });
  assert(words(await table(), 3, 2) === "955", `the average: ${words(await table(), 3, 2)}`);
  await page.keyboard.press("Escape");
  // Excel に保存: a workbook that keeps the formulas.
  const [download] = await Promise.all([page.waitForEvent("download"), ribbonBtn("Excel に保存")]);
  const xlsx = join(outDir, "format-sheet.xlsx");
  await download.saveAs(xlsx);
  assert(/\.xlsx$/.test(download.suggestedFilename()) || download.suggestedFilename() === "download", `an .xlsx: ${download.suggestedFilename()}`);
  const inside = execFileSync("python3", ["-c", `import zipfile,sys;z=zipfile.ZipFile("${xlsx}");print(z.read("xl/worksheets/sheet1.xml").decode())`]).toString();
  assert(inside.includes("<f>SUM(B2:C2)</f>") && inside.includes("<f>B2*2</f>") && inside.includes("<f>AVERAGE(C2:C3)</f>"), "the formulas are in the workbook");
  // CSV に保存: UTF-8 with the values.
  const [csvDownload] = await Promise.all([page.waitForEvent("download"), ribbonBtn("CSV に保存")]);
  const csvFile = join(outDir, "format-sheet.csv");
  await csvDownload.saveAs(csvFile);
  assert((await readFile(csvFile, "utf8")).includes('店舗,"2,000","1,350","3,350"'), "the CSV has the values as they are shown (as Excel saves them)");
  // Excel／CSV から: the workbook read back in place of this table keeps formulas and the values.
  await page.locator("#stageBody .hs-obj td[data-r='1'][data-c='1']").first().click();
  await withFiles([xlsx], () => ribbonBtn("Excel／CSV から"));
  await page.waitForFunction(() => window.__hpptx.slide().elements.find((x) => x.kind === "table").cells[1][3].f === "=SUM(B2:C2)", null, { timeout: 15000 });
  o = await table();
  assert(o.sheet && words(o, 1, 3) === "3,350" && o.cells[2][1].f === "=B2*2" && words(o, 3, 2) === "955", `read back: ${words(o, 1, 3)} ${words(o, 3, 2)}`);
  // A CSV makes a new table; オートSUM adds the column above, 桁区切り rewrites a typed number.
  const csv = join(outDir, "format-in.csv");
  await writeFile(csv, "項目,金額\n売上,1200\n費用,800\n差引,=B2-B3\n");
  await tab("挿入");
  await ribbonBtn("表");
  await page.waitForSelector(".tb-pick");
  await withFiles([csv], () => page.click('.tb-pick-more[data-sheet="import"]'));
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).filter((o) => o.kind === "table").length === 2, null, { timeout: 8000 });
  o = await table();
  assert(o.sheet && o.cells[3][1].f === "=B2-B3" && words(o, 3, 1) === "400", `the CSV's formula works: ${words(o, 3, 1)}`);
  assert(o.cells[1][1].align === "right", "numbers sit on the right");
  await page.locator("#stageBody .hs-obj").last().locator("td[data-r='3'][data-c='1']").dblclick();
  await tab("テーブル デザイン");
  await ribbonBtn("オートSUM");
  await page.waitForFunction(() => window.__hpptx.slide().elements.filter((x) => x.kind === "table").at(-1).cells[3][1].f === "=SUM(B2:B3)", null, { timeout: 5000 });
  assert(words(await table(), 3, 1) === "2000", "オートSUM: 1200 + 800");
  await page.keyboard.press("Escape");
  await page.locator("#stageBody .hs-obj").last().locator("td[data-r='1'][data-c='1']").dblclick();
  await byTitle("桁区切り");
  await page.waitForFunction(() => /1,200/.test(window.__hpptx.slide().elements.filter((x) => x.kind === "table").at(-1).cells[1][1].text || ""), null, { timeout: 5000 });
  await page.keyboard.press("Escape");
  // 数式を使う をやめる: the values stay, the formulas go.
  await byTitle("表のセルに数式");
  await page.waitForFunction(() => !window.__hpptx.slide().elements.filter((x) => x.kind === "table").at(-1).sheet, null, { timeout: 5000 });
  o = await table();
  assert(!o.cells.flat().some((c) => c.f) && words(o, 3, 1) === "2000", "values only");
  await shot("sheet-csv");
});

await step("配布資料マスター・ノート マスター: 用紙の向き・ヘッダーとフッター・全ページの文字を決め、印刷が従い、1回の⌘Zで戻る", async () => {
  await freshSlide();
  await tab("表示");
  await byTitle("配布資料マスター");
  await page.waitForSelector(".pm-dialog[open]");
  assert((await page.$$(".pm-dialog .pm-slotbox")).length === 6, "six slides on a handout page to begin with");
  const wide = () => page.$eval(".pm-dialog .pm-sheet", (el) => el.getBoundingClientRect().width);
  const w0 = await wide();
  await page.selectOption(".pm-dialog .pm-form > label:first-child select", "landscape");
  assert((await wide()) > w0 * 1.2, "the page turns sideways");
  await page.fill('.pm-dialog input[aria-label="ヘッダーの文字"]', "社内限りの配布資料");
  await page.check('.pm-dialog input[data-hf="footer"]');
  await page.fill('.pm-dialog input[aria-label="フッターの文字"]', "ACME");
  await page.click(".pm-add-text");
  await page.fill(".pm-obj-text", "全ページの注記");
  assert(await page.locator(".pm-dialog .pm-page .pr-objects").textContent().then((t) => t.includes("全ページの注記")), "the words are on the page");
  // The header is dragged down and to the right.
  const box = await page.locator('.pm-dialog [data-pm="header"]').boundingBox();
  await page.mouse.move(box.x + 6, box.y + 4);
  await page.mouse.down();
  await page.mouse.move(box.x + 86, box.y + 44, { steps: 4 });
  await page.mouse.up();
  await shot("print-master");
  await page.click(".pm-ok");
  await page.waitForFunction(() => window.__hpptx.deck().printMasters?.handout, null, { timeout: 5000 });
  const m = await page.evaluate(() => window.__hpptx.deck().printMasters.handout);
  assert(m.orientation === "landscape" && m.header.text === "社内限りの配布資料" && m.footer.on === true && m.footer.text === "ACME", `kept: ${JSON.stringify(m).slice(0, 200)}`);
  assert(m.header.x > 48 + 100 && m.header.y > 22 + 40, `the header moved: ${m.header.x},${m.header.y}`);
  assert(m.objects.length === 1 && /全ページの注記/.test(m.objects[0].text), "the words are kept");
  // 印刷 starts from it: the fields, a sideways page, the words on every page.
  await page.keyboard.press("Control+p");
  await page.waitForSelector(".print-dialog[open] .pr-sheet");
  await page.selectOption('.print-dialog select[name="layout"]', "h6");
  await page.waitForTimeout(500);
  assert((await page.inputValue('.print-dialog input[name="headerText"]')) === "社内限りの配布資料" && (await page.isChecked('.print-dialog input[name="footer"]')), "the header and footer fields start as the master has them");
  const sheet = await page.$eval(".pr-preview .pr-sheet", (el) => ({ w: el.getBoundingClientRect().width, h: el.getBoundingClientRect().height }));
  assert(sheet.w > sheet.h, `a landscape page: ${JSON.stringify(sheet)}`);
  assert((await page.textContent(".pr-preview .pr-ph-header")).includes("社内限りの配布資料") && (await page.textContent(".pr-preview .pr-ph-footer")).includes("ACME"), "the header and footer are on the page");
  assert((await page.textContent(".pr-preview .pr-objects")).includes("全ページの注記"), "and the words every page carries");
  await shot("print-master-page");
  // The notes pages are drawn on the ノート マスター (one page for each slide, with the slide and the notes on it).
  await page.selectOption('.print-dialog select[name="layout"]', "notes");
  await page.selectOption('.print-dialog select[name="range"]', "current");
  await page.waitForTimeout(500);
  assert((await page.textContent(".pr-count")).includes("1ページ"), `a notes page: ${await page.textContent(".pr-count")}`);
  assert((await page.$$(".pr-preview .pr-notes .pr-notes-slide")).length === 1 && (await page.$$(".pr-preview .pr-notes .pr-notes-text")).length === 1, "the slide and the notes are on it");
  await page.keyboard.press("Escape");
  // ノート マスター: the slide and the notes can each be moved; no change keeps nothing.
  await tab("表示");
  await byTitle("ノート マスター");
  await page.waitForSelector('.pm-dialog[open] [data-pm="slide"]');
  assert(await page.$('.pm-dialog [data-pm="notes"]'), "the notes box is there too");
  await page.click('.pm-dialog [data-pm="slide"]');
  assert(await page.$(".pm-dialog .pm-box.sel .pm-grip"), "a chosen box has its size handle");
  await shot("notes-master");
  await page.click(".pm-ok");
  await page.waitForTimeout(300);
  assert(!(await page.evaluate(() => window.__hpptx.deck().printMasters?.notes)), "nothing was changed, nothing is kept");
  // ⌘Z takes the handout master back.
  await undo();
  assert(!(await page.evaluate(() => window.__hpptx.deck().printMasters?.handout)), "one ⌘Z undoes the whole master");
});

await step("画面切り替え「変形」: 複製したスライドで動かした部品が、プレビューでも発表でも位置と大きさでつながって動く", async () => {
  await freshSlide();
  await page.mouse.click(...(await stageAt(1700, 1000)));
  await pasteClip(null, { text: "つながる文字" });
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).some((o) => o.kind === "text"), null, { timeout: 5000 });
  const first = await page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide()));
  const part = async () => (await slide()).elements.find((o) => o.kind === "text");
  const x0 = (await part()).x;
  // The slide is copied (the copy keeps the ids) and the text is moved on the copy.
  await page.locator("#filmstrip .film-item").nth(first).click();
  const count = await page.evaluate(() => window.__hpptx.deck().slides.length);
  await page.keyboard.press("Control+d");
  await page.waitForFunction((n) => window.__hpptx.deck().slides.length === n + 1, count, { timeout: 5000 });
  await page.waitForTimeout(400);
  const box = await page.locator('#stageBody .hs-obj[data-kind="text"]').first().boundingBox();
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  for (let i = 0; i < 16; i += 1) await page.keyboard.press("Shift+ArrowRight");
  await page.keyboard.press("Escape");
  const moved = await part();
  assert(moved.x > x0 + 100 && moved.id, `moved on the copy: ${x0} → ${moved.x}`);
  // 画面切り替え → 変形: the message says one part travels, and the preview carries it.
  await tab("画面切り替え");
  await page.click('.an-tr[data-tr="morph"]');
  await page.waitForFunction(() => /1 個がつながって動きます/.test(document.body.textContent), null, { timeout: 4000 });
  await page.waitForSelector(".transition-preview .hs-morph-layer", { timeout: 3000 });
  assert((await page.$$eval(".transition-preview .hs-morph-layer .hs-obj", (els) => els.length)) >= 1, "the old text travels in a layer over the new slide");
  await shot("morph-preview");
  await page.waitForSelector(".transition-preview .hs-morph-layer", { state: "detached", timeout: 4000 });
  assert((await slide()).transition === "morph", "the slide's way in is 変形");
  // Presenting: from the first slide, → plays it.
  await page.keyboard.press("Escape");
  await page.locator("#filmstrip .film-item").nth(first).click();
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player .hs-slide", { timeout: 8000 });
  await page.waitForTimeout(600);
  await page.keyboard.press("ArrowRight");
  await page.waitForSelector(".hs-player .hs-morph-layer", { timeout: 3000 });
  const travelling = await page.evaluate(() => { const n = document.querySelector(".hs-player .hs-morph-layer .hs-obj"); return Boolean(n && n.getAnimations().length); });
  assert(travelling, "the old text is being animated across");
  await shot("morph-present");
  await page.waitForSelector(".hs-player .hs-morph-layer", { state: "detached", timeout: 4000 });
  assert(await page.evaluate(() => /つながる文字/.test(document.querySelector(".hs-player .hs-slide")?.textContent || "")), "and the new slide shows the text where it ends");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

await step("挿入 → QR コード: アドレスは QR コードの図になりクリックで開く動作が付き、ただの文字はリンクなしの図になる", async () => {
  await freshSlide();
  const qr = async (words) => {
    const count = (await slide()).elements?.filter((o) => o.kind === "image").length || 0;
    await tab("挿入");
    await ribbonBtn("QR");
    await page.waitForSelector("#askDialog[open]");
    await page.fill("#askInput", words);
    await page.click("#askOk");
    await page.waitForFunction((n) => (window.__hpptx.slide().elements || []).filter((o) => o.kind === "image").length > n, count, { timeout: 8000 });
    return (await slide()).elements.filter((o) => o.kind === "image").at(-1);
  };
  const link = await qr("https://example.com/acme");
  assert(/^data:image\/svg\+xml/.test(link.src) && link.w === 360 && link.h === 360, `an SVG picture: ${link.src.slice(0, 40)}`);
  assert(link.action?.type === "url" && link.action.href === "https://example.com/acme", `a link: ${JSON.stringify(link.action)}`);
  assert(/https:\/\/example\.com\/acme/.test(link.alt), `its alternative text names it: ${link.alt}`);
  await page.waitForFunction((id) => { const img = document.querySelector(`#stageBody .hs-obj[data-el="${id}"] img`); return Boolean(img && img.complete && img.naturalWidth > 0); }, link.id, { timeout: 5000 });
  await shot("qr");
  const plain = await qr("会議室 3F");
  assert(!plain.action && /会議室 3F/.test(plain.alt), "words that are not an address are a picture only");
  // The QR code of an address read back: the picture encodes the same bytes for the same words.
  const again = await qr("https://example.com/acme");
  assert(again.src === link.src, "the same words make the same picture");
});

// 画像として保存 shares this tab: a browser that accepts sharing its own tab, as a person would by choosing it.
const capBrowser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: [...browserArgs, "--auto-accept-this-tab-capture", "--use-fake-ui-for-media-stream"] });
await step("画像として保存: the current slide as a PNG of the slide itself, and every slide in a ZIP", async () => {
  const cap = await (await capBrowser.newContext({ viewport: { width: 1600, height: 900 }, acceptDownloads: true })).newPage();
  cap.on("pageerror", (error) => errors.push(`capture pageerror: ${error.message}`));
  await cap.goto(base);
  await cap.evaluate(() => localStorage.clear());
  await cap.goto(base);
  await cap.click("#sampleDeckBtn");
  await cap.waitForSelector(".film-item");
  await cap.locator(".film-item").nth(1).click();
  const fileMenu = () => cap.click(".rb-file");
  await fileMenu();
  await cap.locator('.rb-pop button:has-text("画像として保存")').first().click();
  await cap.waitForSelector(".imgx-dialog[open]");
  await cap.check('.imgx-dialog input[value="one"]');
  const [png] = await Promise.all([cap.waitForEvent("download", { timeout: 60000 }), cap.click(".imgx-dialog .ix-go")]);
  const pngFile = join(outDir, "format-slide.png");
  await png.saveAs(pngFile);
  const facts = JSON.parse(execFileSync("python3", ["-c", `
from PIL import Image; import json, sys
im = Image.open(sys.argv[1]).convert("RGB"); w, h = im.size
colors = len(set(im.resize((64, 36)).getdata()))
px = im.load()
dark = lambda p: sum(p) < 360
# The foot of the slide: the deck's name at the bottom left and the page number at the bottom right (the whole
# slide is in the picture, not a part of it).
left = sum(dark(px[x, y]) for x in range(0, int(w * .3), 2) for y in range(int(h * .88), h, 2))
right = sum(dark(px[x, y]) for x in range(int(w * .8), w, 2) for y in range(int(h * .88), h, 2))
print(json.dumps({"w": w, "h": h, "colors": colors, "corner": im.getpixel((5, 5)), "left": left, "right": right}))`, pngFile]).toString());
  assert(Math.abs(facts.w / facts.h - 16 / 9) < 0.03 && facts.w >= 640, `a slide-shaped picture: ${facts.w}×${facts.h}`);
  assert(facts.colors > 12, `the slide drawn (not blank): ${facts.colors} colours`);
  assert(facts.corner.every((v) => v > 200), `the slide's own white corner, not the dark show around it: ${facts.corner}`);
  assert(facts.left > 5 && facts.right > 5, `the whole slide (the foot's words at both corners: ${facts.left}, ${facts.right})`);
  assert(facts.w >= 1500, `at the screen's resolution: ${facts.w}px wide`);
  await fileMenu();
  await cap.locator('.rb-pop button:has-text("画像として保存")').first().click();
  await cap.waitForSelector(".imgx-dialog[open]");
  const total = await cap.evaluate(() => Number(document.querySelector(".imgx-dialog").textContent.match(/（(\d+)枚/)?.[1] || 0));
  const [zip] = await Promise.all([cap.waitForEvent("download", { timeout: 240000 }), cap.click(".imgx-dialog .ix-go")]);
  const zipFile = join(outDir, "format-slides.zip");
  await zip.saveAs(zipFile);
  const names = execFileSync("python3", ["-c", "import zipfile,sys;z=zipfile.ZipFile(sys.argv[1]);assert z.testzip() is None;print('|'.join(z.namelist()))", zipFile]).toString().trim().split("|");
  assert(names.length === total && names[0] === "スライド01.png", `a picture a slide: ${names.length} of ${total} (${names[0]})`);
  await cap.close();
});
await step("アニメーション GIF の作成: 2 slides at 極小 play into a looping GIF of the slides (frames that changed only)", async () => {
  const cap = await (await capBrowser.newContext({ viewport: { width: 1600, height: 900 }, acceptDownloads: true })).newPage();
  cap.on("pageerror", (error) => errors.push(`gif pageerror: ${error.message}`));
  await cap.goto(base);
  await cap.evaluate(() => localStorage.clear());
  await cap.goto(base);
  await cap.click("#sampleDeckBtn");
  await cap.waitForSelector(".film-item");
  await cap.click(".rb-file");
  await cap.locator('.rb-pop button:has-text("アニメーション GIF の作成")').first().click();
  await cap.waitForSelector(".gif-dialog[open]");
  await cap.selectOption(".gif-dialog .gif-size", "xs");
  await cap.fill(".gif-dialog .gif-secs", "1.5");
  await cap.fill(".gif-dialog .gif-from", "2");
  await cap.fill(".gif-dialog .gif-to", "3");
  assert((await cap.textContent(".gif-dialog .gif-estimate")).includes("2枚"), "two slides");
  const [gif] = await Promise.all([cap.waitForEvent("download", { timeout: 120000 }), cap.click(".gif-dialog .gif-go")]);
  const gifFile = join(outDir, "format-show.gif");
  await gif.saveAs(gifFile);
  const facts = JSON.parse(execFileSync("python3", ["-c", `
from PIL import Image; import json, sys
im = Image.open(sys.argv[1]); w, h = im.size; n = im.n_frames; total = 0
for i in range(n): im.seek(i); total += im.info.get("duration", 0)
im.seek(n - 1); last = im.convert("RGB")
px = last.load(); green = sum(1 for x in range(0, w, 2) for y in range(int(h * .08), int(h * .2)) if px[x, y][1] > 100 and px[x, y][0] < 90 and px[x, y][2] < 110)
print(json.dumps({"format": im.format, "w": w, "h": h, "n": n, "loop": im.info.get("loop"), "ms": total, "green": green, "colors": len(set(last.resize((64, 36)).getdata()))}))`, gifFile]).toString());
  assert(facts.format === "GIF" && facts.w === 320 && facts.h === 180, `a 320×180 GIF: ${JSON.stringify(facts)}`);
  assert(facts.n >= 2 && facts.loop === 0, `animated and looping: ${facts.n} frames, loop ${facts.loop}`);
  assert(facts.ms >= 2000 && facts.ms <= 12000, `about the two slides' time: ${facts.ms} ms`);
  assert(facts.colors > 8, `the slide drawn: ${JSON.stringify(facts)}`);
  await cap.close();
});
await capBrowser.close();

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
