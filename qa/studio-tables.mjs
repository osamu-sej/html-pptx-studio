// Tables, charts, links in words, cropping and shapes drawn by hand, in a real browser: a table from the 挿入
// grid typed with Tab, cells merged (Shift+click), rows and columns added and removed, a style, a cell fill, a
// column line dragged; a chart from 挿入 → グラフ with its data sheet (and a paste from Excel); a link on words
// that opens while presenting; a picture cropped on the stage; 曲線・フリーフォーム drawn and a point moved —
// then the presentation and the exported file.
// Usage: node qa/studio-tables.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir, readFile } from "node:fs/promises";
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
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true, permissions: ["clipboard-read", "clipboard-write"] });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const shot = async (name) => { const file = join(outDir, `tables-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await page.waitForTimeout(150); if (await page.isVisible(".slide-wrap .hs-slide")) await measure(); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };
const slideNow = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide() ?? null)));
const objects = async () => (await slideNow())?.elements ?? [];
const find = async (kind) => (await objects()).find((o) => o.kind === kind);
const text = (html) => String(html || "").replace(/<[^>]+>/g, "");
const assert = (ok, message) => { if (!ok) throw new Error(message); };
let box = null;
const at = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
const measure = async () => { box = await page.locator(".slide-wrap .hs-slide").boundingBox(); };
const tab = (label) => page.click(`.rb-tabs [role=tab]:has-text("${label}")`);
const ribbon = async (label) => {
  const direct = page.locator(`.rb-body .rb-btn:has-text("${label}")`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct.click();
  for (const folded of await page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = page.locator(`.rb-pop .rb-fold .rb-btn:has-text("${label}")`).first();
    if (await inside.count()) return inside.click();
  }
  throw new Error(`no ribbon button "${label}"`);
};
const menuItem = (label) => page.locator(`.rb-pop button:has-text("${label}")`).first().click();
const cell = (r, c) => page.locator(`.slide-wrap .hs-obj-table td[data-r="${r}"][data-c="${c}"]`);
// Double-click / click a cell where it is on screen (the selection frame lies over it).
const cellPoint = async (r, c) => { const b = await cell(r, c).boundingBox(); return [b.x + b.width / 2, b.y + b.height / 2]; };
const typeIn = async (r, c) => { await page.mouse.dblclick(...(await cellPoint(r, c))); await page.waitForTimeout(250); };
const shortcut = (key) => page.keyboard.press(process.platform === "darwin" ? `Meta+${key}` : `Control+${key}`);

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);

await step("a 4 × 3 table from the 挿入 grid, typed with Tab (a row is added after the last cell)", async () => {
  await page.click("#sampleDeckBtn");
  await page.waitForSelector(".film-item");
  await ribbon("新しいスライド");
  await menuItem("白紙");
  await page.waitForTimeout(400);
  await measure();
  await tab("挿入");
  await ribbon("表");
  await page.waitForSelector(".tb-pick");
  await page.locator(".tb-pick button").nth(2 * 10 + 3).hover();
  assert((await page.textContent(".tb-pick-label")).includes("4 × 3"), "the grid says 4 × 3");
  await page.locator(".tb-pick button").nth(2 * 10 + 3).click();
  await page.waitForTimeout(400);
  for (const t of ["項目", "2025年", "2026年", "前年比", "売上", "1,200", "1,440", "120%", "客数", "980", "1,020", "104%"]) { await page.keyboard.type(t); await page.keyboard.press("Tab"); }
  await page.keyboard.type("客単価");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  const t = await find("table");
  assert(t.cells.length === 4 && t.cols.length === 4, `size ${t.cells.length} × ${t.cols.length}`);
  assert(text(t.cells[0][3].text) === "前年比" && text(t.cells[2][3].text) === "104%" && text(t.cells[3][0].text) === "客単価", `cells ${JSON.stringify(t.cells.map((row) => row.map((c) => text(c.text))))}`);
  const tabs = await page.$$eval(".rb-tabs [role=tab]", (els) => els.map((el) => el.textContent).join());
  assert(/テーブル デザイン,レイアウト/.test(tabs), `contextual tabs: ${tabs}`);
  await shot("typed");
});

await step("merge two cells with Shift+click, then split them", async () => {
  await typeIn(3, 1);
  await page.keyboard.down("Shift");
  await page.mouse.click(...(await cellPoint(3, 2)));
  await page.keyboard.up("Shift");
  await page.waitForTimeout(200);
  assert(await page.locator(".slide-wrap td.ed-cell-sel").count() === 2, "two cells picked");
  await tab("レイアウト");
  await ribbon("セルの結合");
  await page.waitForTimeout(300);
  let t = await find("table");
  assert(t.cells[3][1].cs === 2 && t.cells[3][2].merged, `merged: ${JSON.stringify(t.cells[3])}`);
  await ribbon("セルの分割");
  await page.waitForTimeout(300);
  t = await find("table");
  assert(!t.cells[3][1].cs && !t.cells[3][2].merged, `split: ${JSON.stringify(t.cells[3])}`);
});

await step("rows and columns: insert below and right, delete a row; undo", async () => {
  await ribbon("下に行を挿入");
  await page.waitForTimeout(300);
  let t = await find("table");
  assert(t.cells.length === 5, `rows: ${t.cells.length}`);
  await ribbon("右に列を挿入");
  await page.waitForTimeout(300);
  t = await find("table");
  assert(t.cols.length === 5, `cols: ${t.cols.length}`);
  await ribbon("削除");
  await menuItem("行の削除");
  await page.waitForTimeout(300);
  t = await find("table");
  assert(t.cells.length === 4, `after delete: ${t.cells.length}`);
  await page.keyboard.press("Escape");
  await shortcut("z");
  await page.waitForTimeout(300);
  assert((await find("table")).cells.length === 5, "undo");
  await shortcut("y").catch(() => {});
});

await step("テーブル デザイン: a PowerPoint table style, header off and on, a cell fill from the theme's colours", async () => {
  await typeIn(1, 0);
  await tab("テーブル デザイン");
  await page.click('.tb-style[title="淡色スタイル 1 - アクセント 1"]');
  await page.waitForTimeout(300);
  let t = await find("table");
  assert(t.style === "light1", `style ${t.style}`);
  await ribbon("塗りつぶし");
  await page.locator('.rb-pop .rb-sheet .rb-sw[aria-label="アクセント 1、明るく 80%"]').click();
  await page.waitForTimeout(300);
  t = await find("table");
  const tint = await page.evaluate(() => window.SlideEngine.paletteSheet(window.__hpptx.deck().theme).columns[4].variants[0].hex);
  assert(t.cells[1][0].fill === tint, `fill: ${JSON.stringify(t.cells[1][0])} (expected ${tint})`);
  await page.keyboard.press("Escape");
  await shot("styled");
});

await step("drag a column line on the stage", async () => {
  const t = await find("table");
  await page.click(".slide-wrap", { position: { x: 5, y: 5 } }).catch(() => {});
  await page.mouse.click(...at(t.x + 10, t.y + t.h + 40));
  await page.locator("#formatTab").click();
  const list = await objects();
  await page.locator("#formatPane .fp-sel-list li").nth(list.length - 1 - list.findIndex((o) => o.kind === "table")).click();
  await page.waitForTimeout(200);
  const handle = page.locator('.ed-layer .ed-tline.col[data-handle="tbl:col:0"]');
  const b = await handle.boundingBox();
  // Mid-way down the first row (not where a row line crosses).
  const y = b.y + (b.height * t.rows[0]) / 2;
  await page.mouse.move(b.x + b.width / 2, y);
  await page.mouse.down();
  await page.mouse.move(b.x + b.width / 2 + 60, y, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  const after = await find("table");
  assert(after.cols[0] > t.cols[0] + 0.02, `first column wider: ${t.cols[0]} → ${after.cols[0]}`);
});

await step("a chart from 挿入 → グラフ, its data sheet (typed and pasted from Excel)", async () => {
  await tab("挿入");
  await ribbon("グラフ");
  await page.click('.tb-chart-grid button[data-chart="stacked-bar"]');
  await page.waitForSelector(".tb-sheet-dialog[open]");
  await page.fill('.tb-sheet input[aria-label="項目1"]', "東京");
  await page.evaluate(() => {
    const table = document.querySelector(".tb-sheet");
    const data = new DataTransfer();
    data.setData("text/plain", "\t食品\t日用品\t飲料\n東京\t50\t20\t30\n大阪\t40\t30\t25\n名古屋\t60\t25\t20\n");
    table.dispatchEvent(new ClipboardEvent("paste", { clipboardData: data, bubbles: true, cancelable: true }));
  });
  await page.waitForTimeout(500);
  await page.fill('.tb-sheet-side input[aria-label="グラフのタイトル"]', "地域別の売上");
  await page.waitForTimeout(400);
  await shot("sheet");
  await page.click('.tb-sheet-dialog button:has-text("反映する")');
  await page.waitForTimeout(400);
  const c = await find("chart");
  assert(c.chart.type === "stacked-bar" && c.chart.labels.join() === "東京,大阪,名古屋" && c.chart.series.map((s) => s.name).join() === "食品,日用品,飲料", `chart: ${JSON.stringify(c.chart)}`);
  assert(c.chart.title === "地域別の売上" && c.chart.series[2].values.join() === "30,25,20", "title and values");
  assert(await page.isVisible(".slide-wrap .hs-obj-chart svg"), "the chart is on the slide");
  const tabs = await page.$$eval(".rb-tabs [role=tab]", (els) => els.map((el) => el.textContent).join());
  assert(/グラフのデザイン/.test(tabs), "グラフのデザイン tab");
  await ribbon("グラフの種類の変更");
  await page.click('.tb-chart-grid button[data-chart="multi-line"]');
  await page.waitForTimeout(300);
  assert((await find("chart")).chart.type === "multi-line", "kind changed");
  await shot("chart");
});

await step("a link on words: typed, picked, 挿入 → リンク", async () => {
  await tab("挿入");
  await ribbon("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  await page.mouse.click(...at(120, 900));
  await page.waitForTimeout(200);
  await page.keyboard.type("詳しくは社内ポータル");
  await page.keyboard.down("Shift");
  for (let i = 0; i < 6; i += 1) await page.keyboard.press("ArrowLeft");
  await page.keyboard.up("Shift");
  await ribbon("リンク・動作");
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", "https://example.com/portal");
  await page.click("#askOk");
  await page.waitForTimeout(300);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const t = (await objects()).find((o) => o.kind === "text");
  assert(/<a href="https:\/\/example\.com\/portal"[^>]*>社内ポータル<\/a>/.test(t.text), `link: ${t.text}`);
});

await step("crop a picture on the stage (drag a mark, Enter keeps it)", async () => {
  await tab("挿入");
  await ribbon("画像");
  await menuItem("内蔵の写真");
  await page.waitForSelector(".rb-gallery.photos");
  await page.locator(".rb-gallery.photos button").first().click();
  await page.waitForFunction(() => window.__hpptx.slide().elements.some((o) => o.kind === "image"), null, { timeout: 5000 });
  const before = await find("image");
  await tab("図の形式");
  await ribbon("トリミング");
  await page.waitForSelector(".ed-crop-box");
  const mark = await page.locator(".ed-crop-e").boundingBox();
  await page.mouse.move(mark.x + mark.width / 2, mark.y + mark.height / 2);
  await page.mouse.down();
  await page.mouse.move(mark.x - 80, mark.y + mark.height / 2, { steps: 6 });
  await page.mouse.up();
  await shot("crop");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(400);
  const after = await find("image");
  assert(after.crop?.r > 0.05 && after.w < before.w - 50 && Math.abs(after.x - before.x) < 1, `cropped: ${JSON.stringify(after.crop)} ${before.w} → ${after.w}`);
  await page.keyboard.press("Delete");
  await page.waitForTimeout(300);
});

await step("曲線 and フリーフォーム drawn by hand; a point moved (頂点の編集)", async () => {
  await tab("挿入");
  await page.locator('.rb-btn.big:has-text("図形")').first().click();
  await page.click('.rb-gallery button[data-freeform="polygon"]');
  await page.waitForSelector(".ed-freeform");
  await page.mouse.click(...at(1300, 250));
  await page.mouse.click(...at(1700, 300));
  await page.mouse.click(...at(1600, 600));
  await page.mouse.click(...at(1301, 251));
  await page.waitForTimeout(400);
  let shapes = (await objects()).filter((o) => o.shape === "custom");
  assert(shapes.length === 1 && shapes[0].path.closed && shapes[0].path.pts.length === 3 && shapes[0].fill === await page.evaluate(() => window.SlideEngine.OBJECT_DEFAULTS.shape.fill), `polygon: ${JSON.stringify(shapes[0])}`);
  await page.locator('.rb-btn.big:has-text("図形")').first().click().catch(async () => { await tab("挿入"); await page.locator('.rb-btn.big:has-text("図形")').first().click(); });
  await page.click('.rb-gallery button[data-freeform="curve"]');
  await page.waitForSelector(".ed-freeform");
  await page.mouse.click(...at(1250, 750));
  await page.mouse.click(...at(1450, 650));
  await page.mouse.click(...at(1650, 800));
  await page.mouse.dblclick(...at(1800, 700));
  await page.waitForTimeout(400);
  shapes = (await objects()).filter((o) => o.shape === "custom");
  assert(shapes.length === 2 && !shapes[1].path.closed && shapes[1].path.curve && shapes[1].fill === "none", `curve: ${JSON.stringify(shapes[1])}`);
  // 頂点の編集 on the polygon (right-click menu).
  const poly = shapes[0];
  await page.mouse.click(...at(poly.x + poly.w * 0.6, poly.y + poly.h * 0.4));
  await page.mouse.click(...at(poly.x + poly.w * 0.6, poly.y + poly.h * 0.4), { button: "right" });
  await page.click('.ed-menu button:has-text("頂点の編集")');
  await page.waitForSelector(".ed-vertex");
  const v = await page.locator(".ed-vertex").nth(1).boundingBox();
  await page.mouse.move(v.x + v.width / 2, v.y + v.height / 2);
  await page.mouse.down();
  await page.mouse.move(v.x + 60, v.y + 40, { steps: 5 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  const edited = (await objects()).find((o) => o.id === poly.id);
  assert(edited.w > poly.w + 60, `the box follows the point: ${poly.w} → ${edited.w}`);
  await page.keyboard.press("Escape");
  await shot("freeform");
});

await step("presenting: the table, the chart, the link and the drawn shapes", async () => {
  await page.waitForTimeout(1200);
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player .hs-slide .hs-obj-table");
  await page.waitForTimeout(800);
  assert(await page.locator(".hs-player .hs-obj-chart svg").count() === 1, "chart");
  assert(await page.locator(".hs-player .hs-obj[data-el] .hs-obj-geom path").count() >= 2, "drawn shapes");
  const link = page.locator('.hs-player .hs-obj-tx a[href="https://example.com/portal"]');
  assert(await link.count() === 1, "the link");
  const count = await page.textContent(".hs-player-count");
  const [popup] = await Promise.all([context.waitForEvent("page", { timeout: 5000 }).catch(() => null), link.click()]);
  await page.waitForTimeout(500);
  assert(popup, "the link opens a new tab");
  await popup?.close();
  assert((await page.textContent(".hs-player-count")) === count, "the slide did not move on");
  await shot("present");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

await step("export: the table, chart and drawn shapes in the file", async () => {
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#downloadBtn").then(async () => { if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn"); })]);
  const file = join(outDir, "tables-export.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  assert(/"kind":"table"/.test(html) && /"kind":"chart"/.test(html) && /"shape":"custom"/.test(html), "objects in the file");
  const viewer = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  await viewer.goto(`file://${file}#2`);
  await viewer.waitForSelector(".hs-player .hs-obj-table td");
  const cells = await viewer.locator(".hs-player .hs-obj-table td").count();
  await viewer.close();
  assert(cells >= 20, `cells in the file: ${cells}`);
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
