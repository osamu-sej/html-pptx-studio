// 図形に変換 in a real browser: an AI layout slide (cards with a "詳しく" card and a deep-dive page, a click-by-click
// process, a chart, a table) becomes objects from ホーム → 図形に変換 and from the right-click menu; the pieces of
// an item move together, words are edited in place, undo brings the layout back; the cover cannot be converted;
// the converted slides present with their build (one item per click), the deep-dive page still opens from its item,
// and the exported file carries the objects.
// Usage: node qa/studio-convert.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const shot = async (name) => { const file = join(outDir, `convert-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); if (await page.isVisible(".slide-wrap .hs-slide")) await measure(); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
  for (let i = 0; i < 3 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
};
const slideNow = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide() ?? null)));
const text = (html) => String(html || "").replace(/<[^>]+>/g, "");
const assert = (ok, message) => { if (!ok) throw new Error(message); };
let box = null;
const at = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
const measure = async () => { box = await page.locator(".slide-wrap .hs-slide").boundingBox(); };
const tab = (label) => page.click(`.rb-tabs [role=tab]:has-text("${label}")`);
// A ribbon button, opening its group first when a narrow window has folded the group into one button.
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
const ribbonButton = async (label) => {
  const direct = page.locator(`.rb-body .rb-btn:has-text("${label}")`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct;
  for (const folded of await page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = page.locator(`.rb-pop .rb-fold .rb-btn:has-text("${label}")`).first();
    if (await inside.count()) return inside;
  }
  return null;
};
const shortcut = (key) => page.keyboard.press(process.platform === "darwin" ? `Meta+${key}` : `Control+${key}`);
const goTo = async (n) => { await page.click(`.film-item:nth-child(${n})`); await page.waitForTimeout(500); await measure(); };
const waitBlank = () => page.waitForFunction(() => window.__hpptx.slide()?.type === "blank", null, { timeout: 8000 });

const deck = {
  title: "変換の確認", audience: "社内", purpose: "確認", theme: "clarity", transition: "fade", motion: {}, memo: "",
  slides: [
    { type: "title", title: "変換の確認" },
    { type: "cards", title: "選び方の3つの軸", takeaway: "用途・品質・運用で比べる", notes: "話すこと", source: "社内調査", items: [{ title: "用途", desc: "何に使うか" }, { title: "品質", desc: "正確さと速さ" }, { title: "運用", desc: "安全に使えるか" }], details: [{ target: "items[0]", text: "用途の内訳" }] },
    { type: "content", title: "品質の深掘り", drillOf: "items[1]", points: ["正確さ", "速さ"] },
    { type: "process", title: "進め方", takeaway: "3段階で進める", items: [{ title: "準備", desc: "目的を共有" }, { title: "実行", desc: "手順を実施" }, { title: "確認", desc: "結果を振り返る" }] },
    { type: "imageText", title: "売上の推移", takeaway: "伸びている", image: { chartType: "bar", data: { title: "売上", unit: "億円", items: [{ label: "1月", value: 10 }, { label: "2月", value: 14 }, { label: "3月", value: 18 }] } }, points: ["3か月で1.8倍"] },
    { type: "table", title: "目標", takeaway: "半年で達成", headers: ["項目", "現状", "目標"], rows: [["作業時間", "120h", "80h"], ["廃棄率", "3.0%", "2.5%"]] },
    { type: "closing", title: "おわり" },
  ],
};

await page.goto(base);
await page.evaluate((value) => localStorage.setItem("hpptx-studio-current-v1", JSON.stringify({ deck: value, selected: 1, savedAt: new Date().toISOString() })), deck);
await page.goto(base);
await page.waitForSelector(".film-item");
await page.waitForTimeout(900);

await step("the cover keeps its layout: 図形に変換 is off there", async () => {
  await goTo(1);
  await tab("ホーム");
  const button = await ribbonButton("図形に変換");
  assert(button, "the button is on the ホーム tab");
  assert(await button.isDisabled(), "disabled on the cover");
  await page.keyboard.press("Escape");
});

await step("ホーム → 図形に変換: the cards become shapes and text boxes, grouped per card; title, key message, notes stay", async () => {
  await goTo(2);
  await shot("cards-before");
  await ribbon("図形に変換");
  await waitBlank();
  await page.waitForTimeout(600);
  const slide = await slideNow();
  const list = slide.elements || [];
  assert(list.length >= 6, `objects: ${list.length}`);
  for (const word of ["用途", "品質", "運用", "正確さと速さ"]) assert(list.some((o) => text(o.text).includes(word)), `"${word}" is in a text box`);
  assert(new Set(list.map((o) => o.group).filter(Boolean)).size >= 3, "a group per card");
  assert(slide.title === "選び方の3つの軸" && slide.takeaway === "用途・品質・運用で比べる" && slide.notes === "話すこと" && slide.source === "社内調査", "kept fields");
  assert(!slide.items, "the layout's items are gone");
  assert(list.some((o) => o.item === "items[0]") && list.some((o) => o.item === "items[1]"), "the card with details and the one with a deep-dive page keep their item");
  assert(await page.locator('.slide-wrap .hs-obj[data-item="items[1]"] .hs-drill-badge').count() === 1, "the deep-dive mark sits on the card");
  // Every object stands where its data says (a mark must not push its object out of place).
  const misplaced = await page.evaluate(() => {
    const slide = document.querySelector(".slide-wrap .hs-slide");
    const r = slide.getBoundingClientRect();
    const k = r.width / 1920;
    return window.__hpptx.slide().elements.filter((o) => o.kind !== "line").filter((o) => {
      const el = slide.querySelector(`.hs-obj[data-el="${o.id}"]`);
      const b = el.getBoundingClientRect();
      return Math.abs((b.top - r.top) / k - o.y) > 2 || Math.abs((b.left - r.left) / k - o.x) > 2;
    }).map((o) => o.id);
  });
  assert(!misplaced.length, `objects out of place: ${misplaced.join()}`);
  await shot("cards-after");
});

await step("a card's pieces move together; double-click edits a word in place", async () => {
  const before = await slideNow();
  const word = before.elements.find((o) => text(o.text) === "運用");
  const mates = before.elements.filter((o) => o.group && o.group === word.group);
  assert(mates.length >= 2, "the card is a group");
  await page.mouse.click(...at(word.x + 20, word.y + word.h / 2));
  await page.waitForTimeout(200);
  const sel = await page.evaluate(() => window.__hpptx.selection());
  assert(mates.every((o) => sel.includes(o.id)), "clicking selects the whole card");
  const [sx, sy] = at(word.x + 20, word.y + word.h / 2);
  await page.mouse.move(sx, sy);
  await page.mouse.down();
  await page.mouse.move(sx + 30, sy + 40, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  const after = await slideNow();
  const moved = after.elements.filter((o) => o.group === word.group);
  assert(moved.every((o) => { const was = mates.find((m) => m.id === o.id); return was && (o.kind === "line" ? o.x1 !== was.x1 : Math.abs(o.y - was.y) > 10); }), "every piece moved");
  await shortcut("z");
  await page.waitForTimeout(300);
  await page.keyboard.press("Escape");
  await page.keyboard.press("Escape");
  const t = (await slideNow()).elements.find((o) => text(o.text) === "運用");
  await page.mouse.dblclick(...at(t.x + 20, t.y + t.h / 2));
  await page.waitForTimeout(200);
  await page.mouse.dblclick(...at(t.x + 20, t.y + t.h / 2));
  await page.waitForTimeout(300);
  await shortcut("a");
  await page.keyboard.type("運用と保守");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  assert((await slideNow()).elements.some((o) => text(o.text) === "運用と保守"), "the word was edited");
});

await step("undo brings the layout back; the right-click menu converts again", async () => {
  for (let i = 0; i < 4 && (await slideNow()).type === "blank"; i += 1) { await page.keyboard.press("Escape"); await shortcut("z"); await page.waitForTimeout(300); }
  assert((await slideNow()).type === "cards", "the layout is back");
  await page.keyboard.press("Escape");
  await page.mouse.click(...at(1700, 1000), { button: "right" });
  await page.waitForSelector(".ed-menu");
  await page.click('.ed-menu button:has-text("図形に変換")');
  await waitBlank();
  assert(((await slideNow()).elements || []).length >= 6, "converted from the menu");
});

await step("a click-by-click process keeps its build as animations on the steps", async () => {
  await goTo(4);
  await tab("ホーム");
  await ribbon("図形に変換");
  await waitBlank();
  const slide = await slideNow();
  const clicks = (slide.timeline || []).filter((e) => e.start === "click" && e.cls === "in");
  assert(clicks.length === 3, `click entrances: ${clicks.length}`);
  assert(clicks.every((e) => /^grp:|^o/.test(e.el)), "on the step groups");
  await shot("process-after");
});

await step("a chart becomes a chart object with its data; a table a table object", async () => {
  await goTo(5);
  await ribbon("図形に変換");
  await waitBlank();
  const chart = (await slideNow()).elements.find((o) => o.kind === "chart");
  assert(chart && chart.chart.labels.join() === "1月,2月,3月" && chart.chart.series[0].values.join() === "10,14,18", `chart: ${JSON.stringify(chart?.chart)}`);
  await goTo(6);
  await ribbon("図形に変換");
  await waitBlank();
  const table = (await slideNow()).elements.find((o) => o.kind === "table");
  assert(table && table.cells.length === 3 && text(table.cells[1][2].text) === "80h", `table: ${JSON.stringify(table?.cells)}`);
  await shot("table-after");
});

await step("presenting: one step per click, and the deep-dive page opens from its card", async () => {
  await goTo(4);
  await page.waitForTimeout(1200);
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector("#presenter .hs-player .hs-slide .hs-obj");
  await page.waitForTimeout(900);
  const hidden = () => page.$$eval("#presenter .hs-player-slide .hs-obj", (els) => els.filter((el) => { const fx = el.querySelector(".hs-obj-fx") || el; const cs = getComputedStyle(fx); return cs.visibility === "hidden" || Number(cs.opacity) < 0.05; }).length);
  const start = await hidden();
  assert(start >= 6, `hidden before the clicks: ${start}`);
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(1400);
  const one = await hidden();
  assert(one < start && one > 0, `after one click: ${one} of ${start}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  await goTo(2);
  await page.waitForTimeout(1200);
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector("#presenter .hs-player .hs-slide .hs-obj");
  await page.waitForTimeout(1000);
  await page.locator('#presenter .hs-player-slide .hs-obj[data-drill]').first().click();
  await page.waitForSelector("#presenter .hs-player .hs-slide.hs-drill", { timeout: 5000 });
  await shot("present-drill");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(600);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

await step("export: the converted slides in the file", async () => {
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#downloadBtn").then(async () => { if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn"); })]);
  const file = join(outDir, "convert-export.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  assert(/"type":"blank"/.test(html) && /"kind":"chart"/.test(html) && /"kind":"table"/.test(html), "objects in the file");
  const viewer = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  await viewer.goto(`file://${file}#2`);
  await viewer.waitForSelector(".hs-player .hs-obj");
  const count = await viewer.locator(".hs-player .hs-obj").count();
  await viewer.close();
  assert(count >= 6, `objects in the file: ${count}`);
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
