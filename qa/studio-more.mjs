// More of PowerPoint in a real browser: the chart kinds 面・円・散布図・レーダー・ウォーターフォール・じょうご, スライドの再利用
// (from the library), アウトラインからスライド, フォト アルバム, キャプションの挿入 on a video, 翻訳 (with the AI answered by
// a stand-in) and ビデオの作成 (the show played by itself and recorded; the screen capture is a stand-in stream here).
// Usage: node qa/studio-more.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g|example\.com/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name) => { const file = join(outDir, `more-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible("#presenter .hs-player")) { await page.keyboard.press("Escape"); await page.waitForTimeout(400); }
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
// The stage drawn again from the deck (and the deck saved): away to the first slide and back.
const refresh = async () => { const i = await page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide())); await page.locator('.film-item[data-index="0"]').click(); await page.locator(`.film-item[data-index="${i}"]`).click(); await page.waitForTimeout(300); };
const withFiles = async (paths, open) => { const [chooser] = await Promise.all([page.waitForEvent("filechooser"), open()]); await chooser.setFiles(paths); };

// 翻訳: a stand-in for the AI that returns each string in brackets.
const jobs = new Map();
let seq = 1;
await page.route("**/api/codex/status", (route) => route.fulfill({ json: { available: true, authorized: true, authenticated: true, planType: "test", team: false, fallback: { available: false, model: "x", web: false, reason: "" } } }));
await page.route("**/api/decks/translate", async (route) => {
  const body = JSON.parse(route.request().postData());
  const id = `00000000-0000-4000-8000-${String(seq++).padStart(12, "0")}`;
  jobs.set(id, { id, kind: "translate", status: "completed", stage: "完了", detail: "訳しました。", texts: body.texts.map((t) => `EN[${t}]`) });
  route.fulfill({ status: 202, json: { jobId: id } });
});
await page.route(/\/api\/decks\/[0-9a-f-]{36}(\/events)?$/, (route) => {
  const url = new URL(route.request().url());
  const job = jobs.get(url.pathname.split("/")[3]);
  if (!job) return route.continue();
  if (!url.pathname.endsWith("/events")) return route.fulfill({ json: job });
  return route.fulfill({ status: 200, headers: { "content-type": "text/event-stream" }, body: `event: progress\ndata: ${JSON.stringify(job)}\n\n` });
});

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

await step("グラフ: the kinds in the picker; 円・ウォーターフォール・じょうご・レーダー・散布図・面 inserted and drawn", async () => {
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.waitForSelector(".tb-chart-grid");
  const kinds = await page.$$eval(".tb-chart-grid button", (els) => els.map((el) => el.dataset.chart));
  for (const k of ["area", "pie", "scatter", "radar", "waterfall", "funnel"]) assert(kinds.includes(k), `picker has ${k}: ${kinds}`);
  await page.keyboard.press("Escape");
  const expect = { pie: ".hs-arc", waterfall: "rect.hs-bar[data-neg]", funnel: "rect.hs-bar[data-center]", radar: "polygon.hs-oarea", scatter: ".hs-dot", area: "path.hs-oarea" };
  for (const [kind, sel] of Object.entries(expect)) {
    await tab("挿入");
    await ribbonBtn("グラフ");
    await page.click(`.tb-chart-grid button[data-chart="${kind}"]`);
    await page.waitForTimeout(400);
    if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
    const o = (await slide()).elements.at(-1);
    assert(o.kind === "chart" && o.chart.type === kind, `${kind} inserted: ${o.chart?.type}`);
    assert(await page.locator(`#stageBody .hs-obj[data-el="${o.id}"] ${sel}`).count(), `${kind} drawn with ${sel}`);
  }
  await shot("charts");
  // The kind can be changed afterwards.
  await tab("グラフのデザイン");
  await ribbonBtn("グラフの種類");
  await page.click('.tb-chart-grid button[data-chart="pie"]');
  await page.waitForTimeout(300);
  assert((await slide()).elements.at(-1).chart.type === "pie", "changed to 円");
});

await step("グラフ（追加）: 集合横棒・積み上げ横棒・積み上げ面・バブル・ヒストグラム・箱ひげ図・ツリーマップ・サンバースト drawn; 軸ラベル and 近似曲線", async () => {
  await tab("ホーム");
  await ribbonBtn("新しいスライド");
  await menuItem("白紙");
  await page.waitForTimeout(300);
  const expect = { hbar: "rect.hs-bar.h", "stacked-hbar": "rect.hs-bar.h", "stacked-area": "path.hs-oarea", bubble: "circle.hs-dot", histogram: "rect.hs-bar", boxplot: "rect.hs-bar[data-center]", treemap: "rect.hs-bar[data-center]", sunburst: ".hs-ochart-pie .hs-oslice" };
  for (const [kind, sel] of Object.entries(expect)) {
    await tab("挿入");
    await ribbonBtn("グラフ");
    await page.click(`.tb-chart-grid button[data-chart="${kind}"]`);
    await page.waitForSelector(".tb-sheet-dialog[open]");
    if (["bubble", "histogram", "boxplot", "treemap", "sunburst"].includes(kind)) assert(!(await page.locator(".tb-sheet-dialog .tb-kind-hint").isHidden()), `${kind}: the data sheet says how it is read`);
    await page.locator('.tb-sheet-dialog .dialog-foot button:has-text("反映する")').click();
    await page.waitForTimeout(300);
    const o = (await slide()).elements.at(-1);
    assert(o.kind === "chart" && o.chart.type === kind, `${kind} inserted: ${o.chart?.type}`);
    const el = page.locator(`#stageBody .hs-obj[data-el="${o.id}"]`);
    assert(await el.locator(sel).count(), `${kind} drawn with ${sel}`);
    await el.screenshot({ path: join(outDir, `more-chart-${kind}.png`) });
    await page.keyboard.press("Delete");
    await page.waitForTimeout(200);
  }
  // グラフ要素を追加 → 軸ラベル and 近似曲線 on a line chart.
  await tab("挿入");
  await ribbonBtn("グラフ");
  await page.click('.tb-chart-grid button[data-chart="line"]');
  await page.waitForSelector(".tb-sheet-dialog[open]");
  await page.locator('.tb-sheet-dialog .dialog-foot button:has-text("反映する")').click();
  await tab("グラフのデザイン");
  await ribbonBtn("グラフ要素");
  await menuItem("軸ラベル");
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", "月");
  await page.click("#askOk");
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", "売上（億円）");
  await page.click("#askOk");
  await page.waitForTimeout(300);
  await ribbonBtn("グラフ要素");
  await menuItem("近似曲線");
  await page.waitForTimeout(300);
  const c = (await slide()).elements.at(-1);
  assert(c.chart.opts?.axisX === "月" && c.chart.opts?.axisY === "売上（億円）" && c.chart.opts?.trend === "linear", `opts: ${JSON.stringify(c.chart.opts)}`);
  const el = page.locator(`#stageBody .hs-obj[data-el="${c.id}"]`);
  assert(await el.locator(".hs-axis-title").count() === 2 && await el.locator(".hs-trend").count() === 1, "the axis titles and the trendline drawn");
  await el.screenshot({ path: join(outDir, "more-chart-axis-trend.png") });
});

await step("動作設定ボタン: 進む/次へ drawn with its click set; in the show a click on it goes to the next slide", async () => {
  await tab("挿入");
  await ribbonBtn("図形");
  const names = await page.$$eval(".rb-pop .rb-gallery button[title^='動作設定ボタン']", (els) => els.map((el) => el.title));
  assert(names.length === 12, `twelve action buttons: ${names.length}`);
  await page.click('.rb-pop .rb-gallery button[title="動作設定ボタン: 進む/次へ"]');
  const r = await page.locator(".slide-wrap .hs-slide").first().boundingBox();
  const at = (x, y) => [r.x + (x * r.width) / 1920, r.y + (y * r.height) / 1080];
  await page.mouse.move(...at(1500, 820));
  await page.mouse.down();
  await page.mouse.move(...at(1700, 960), { steps: 5 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  const b = (await slide()).elements.at(-1);
  assert(b.shape === "actionButtonForwardNext" && b.action?.type === "next", `the button: ${JSON.stringify({ shape: b.shape, action: b.action })}`);
  await page.keyboard.press("Escape");
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector("#presenter .hs-player");
  await page.waitForTimeout(800);
  const count = () => page.textContent("#presenter .hs-player-count");
  const before = await count();
  const box = await page.locator(`#presenter .hs-obj[data-el="${b.id}"]`).last().boundingBox();
  await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
  await page.waitForFunction((was) => document.querySelector("#presenter .hs-player-count")?.textContent !== was, before, { timeout: 5000 });
  const pos = (text) => Number(String(text).trim().split("/")[0]);
  assert(pos(await count()) === pos(before) + 1, `moved on to the next slide: ${before} → ${await count()}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  if (await page.isVisible("#presenter .hs-player")) await page.keyboard.press("Escape");
});

let saved = 0;
await step("スライドの再利用: two slides from a deck in the library, after this one (one ⌘Z)", async () => {
  await page.click("#saveDeckBtn");
  await page.waitForTimeout(800);
  saved = (await deck()).slides.length;
  await tab("ホーム");
  await ribbonBtn("新しいスライド");
  await menuItem("スライドの再利用");
  await page.waitForSelector(".st-reuse-dialog[open] .st-source");
  await page.locator(".st-reuse-dialog .st-source").first().click();
  await page.waitForSelector(".st-reuse-dialog .zm-cell");
  await page.locator('.st-reuse-dialog .zm-cell[data-index="3"]').click();
  await page.locator('.st-reuse-dialog .zm-cell[data-index="5"]').click();
  assert((await page.textContent(".st-reuse-dialog .st-count")).includes("2枚"), "two picked");
  await shot("reuse");
  const before = await deck();
  await page.click(".st-reuse-dialog .st-ok");
  await page.waitForTimeout(500);
  const after = await deck();
  assert(after.slides.length === before.slides.length + 2, `two more: ${after.slides.length}`);
  const at = after.slides.indexOf(after.slides.find((s) => s.title === before.slides[3].title && !before.slides.includes(s)));
  assert(after.slides.filter((s) => s.title === before.slides[3].title).length >= 2, "a copy of slide 4");
  assert(new Set(after.slides.map((s) => s.sid).filter(Boolean)).size === after.slides.map((s) => s.sid).filter(Boolean).length, `slide ids stay unique (${at})`);
  await undo();
  assert((await deck()).slides.length === before.slides.length, "⌘Z takes both back");
});

await step("アウトラインからスライド: headings and their points become slides", async () => {
  const before = (await deck()).slides.length;
  await tab("ホーム");
  await ribbonBtn("新しいスライド");
  await menuItem("アウトラインから");
  await page.waitForSelector(".st-outline-dialog[open]");
  await page.fill(".st-outline", "売上の現状\n  前年比112%\n  東日本が伸びている\n来期の打ち手\n- 品揃えの見直し\n- 発注の自動化\nまとめ");
  assert((await page.textContent(".st-outline-count")).includes("3枚"), await page.textContent(".st-outline-count"));
  await page.click(".st-outline-dialog .st-ok");
  await page.waitForTimeout(500);
  const d = await deck();
  assert(d.slides.length === before + 3, `three slides: ${d.slides.length - before}`);
  const s = await slide();
  assert(s.type === "content" && s.title === "売上の現状" && s.points.length === 2, `first: ${JSON.stringify({ type: s.type, title: s.title, points: s.points })}`);
  await undo();
});

await step("フォト アルバム: three pictures, four to a slide with captions → one slide", async () => {
  const before = (await deck()).slides.length;
  const files = ["store-operations.jpg", "customer-experience.jpg", "data-insight.jpg"].map((f) => join(root, "public", "assets", f));
  await tab("挿入");
  await ribbonBtn("画像");
  await withFiles(files, () => menuItem("フォト アルバム"));
  await page.waitForSelector(".st-album-dialog[open]");
  await page.selectOption(".st-album-dialog select", "4");
  await page.check(".st-album-dialog input[type=checkbox] >> nth=0");
  await page.click(".st-album-dialog .st-ok");
  await page.waitForFunction((n) => window.__hpptx.deck().slides.length === n + 1, before, { timeout: 8000 });
  const s = await slide();
  const pics = s.elements.filter((o) => o.kind === "image");
  const caps = s.elements.filter((o) => o.kind === "text");
  assert(pics.length === 3 && caps.length === 3, `3 pictures with captions: ${pics.length} ${caps.length}`);
  assert(pics.every((p) => /^idb:/.test(p.src) && p.alt), "kept in the browser, with alt text");
  assert(s.title.startsWith("フォト アルバム"), s.title);
  await shot("album");
});

await step("キャプションの挿入: an .srt file on a video becomes its subtitles; the checker no longer asks", async () => {
  await tab("挿入");
  await ribbonBtn("ビデオ");
  await menuItem("YouTube・URL");
  await page.waitForSelector("dialog[open] input");
  await page.fill("dialog[open] input", "https://example.com/movie.mp4");
  await page.keyboard.press("Enter");
  await page.waitForSelector('.rb-tabs [role=tab]:text-is("再生")', { timeout: 10000 });
  const vid = (await slide()).elements.find((o) => o.kind === "video");
  assert(vid, "a video");
  await tab("再生");
  const srt = join(outDir, "captions.srt");
  await writeFile(srt, "1\n00:00:00,500 --> 00:00:02,000\n売上の説明です\n");
  await withFiles([srt], () => ribbonBtn("キャプション"));
  await page.waitForTimeout(400);
  const v = (await slide()).elements.find((o) => o.id === vid.id);
  assert(/^WEBVTT/.test(v.captions || "") && /00:00:00\.500 --> 00:00:02\.000/.test(v.captions) && v.captionLang === "ja", `captions: ${v.captions}`);
  await tab("校閲");
  await ribbonBtn("アクセシビリティ");
  await page.waitForSelector("#a11yPane:not([hidden]) .a11y-head");
  const here = await page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide()));
  assert(!(await page.$$eval(`#a11yPane .a11y-item[data-rule="media"][data-slide="${here}"]`, (els) => els.length)), "no caption tip for this slide");
});

await step("翻訳: this slide's words go to the AI and come back in English (one ⌘Z)", async () => {
  await page.evaluate(() => { const s = window.__hpptx.slide(); s.elements.push({ id: "tr1", kind: "text", x: 100, y: 960, w: 600, h: 60, text: "<p>売上の現状</p>" }); });
  await refresh();
  await tab("校閲");
  await ribbonBtn("翻訳");
  await menuItem("英語に");
  await page.waitForFunction(() => /EN\[/.test(JSON.stringify(window.__hpptx.slide())), null, { timeout: 8000 });
  const s = await slide();
  const t = s.elements.find((o) => o.id === "tr1");
  assert(t.text.includes("EN[売上の現状]"), `object text: ${t.text}`);
  assert(s.title.startsWith("EN["), `title: ${s.title}`);
  await undo();
  assert((await slide()).elements.find((o) => o.id === "tr1").text.includes("<p>売上の現状</p>"), "⌘Z puts the Japanese back");
});

await step("ビデオの作成: the show plays by itself once and a .webm is downloaded", async () => {
  // A stand-in for the browser's tab capture: a canvas stream (the real one asks which tab to share).
  await page.evaluate(() => {
    navigator.mediaDevices.getDisplayMedia = async () => { const c = document.createElement("canvas"); c.width = 320; c.height = 180; const g = c.getContext("2d"); let n = 0; setInterval(() => { g.fillStyle = `hsl(${(n += 7) % 360} 60% 60%)`; g.fillRect(0, 0, 320, 180); }, 50); return c.captureStream(15); };
    // Three slides make a short video.
    const d = window.__hpptx.deck();
    d.show = { range: { from: 1, to: 3 } };
  });
  await page.click(".rb-file");
  await page.locator('.rb-pop .rb-menu button:has-text("ビデオの作成")').click();
  await page.waitForSelector(".video-dialog[open]");
  await page.fill(".video-dialog input[type=number]", "1");
  await page.uncheck(".video-dialog input[type=checkbox] >> nth=0");
  assert((await page.textContent(".vx-estimate")).includes("約0分3秒"), await page.textContent(".vx-estimate"));
  const [download] = await Promise.all([page.waitForEvent("download", { timeout: 30000 }), page.click(".video-dialog .vx-go")]);
  // Headless Chromium drops a Japanese file name ("download"); a real browser keeps 「資料名.webm」.
  assert(/\.webm$/.test(download.suggestedFilename()) || download.suggestedFilename() === "download", download.suggestedFilename());
  const file = join(outDir, "more-video.webm");
  await download.saveAs(file);
  const bytes = await import("node:fs/promises").then((fs) => fs.readFile(file));
  assert(bytes.length > 1000, `a video file: ${bytes.length}`);
  assert(bytes.readUInt32BE(0) === 0x1a45dfa3, "a WebM (EBML) file");
  assert(!(await page.isVisible("#presenter .hs-player")), "the show closes at the end");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
