// インタラクション in a real browser — what only HTML does with objects, set from the ribbon and played while
// presenting and in the exported file: a lift and a note under the mouse, a click that opens details, zooms in,
// turns a card over, shows objects as tabs or puts a spotlight on one; motion that keeps going; HTML-only
// animations (typewriter…); a PowerPoint chart that answers the mouse and hides a series from its legend; and
// おまかせ, which puts all of this on a slide (and on a deck brought over from PowerPoint) in one undo step.
// Usage: node qa/studio-interact.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { execFileSync } from "node:child_process";
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
const decks = join(outDir, "import");
await mkdir(decks, { recursive: true });
execFileSync(process.env.PYTHON_BIN || "python3", [join(root, "test", "pptx_fixtures.py"), decks], { stdio: "ignore" });

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const shot = async (name, target = page) => { const file = join(outDir, `interact-${name}.png`); await target.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
  for (let i = 0; i < 3 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const slideNow = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide() ?? null)));
const objOf = async (id) => (await slideNow()).elements.find((o) => o.id === id);
const goTo = async (n) => { await page.click(`.film-item:nth-child(${n})`); await page.waitForTimeout(500); };
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
// A ribbon button by the start of its tooltip (labels like 「動き」 also appear inside other buttons' labels).
const ribbonTitled = async (prefix) => {
  const direct = page.locator(`.rb-body .rb-btn[title^="${prefix}"]`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct.click();
  for (const folded of await page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = page.locator(`.rb-pop .rb-fold .rb-btn[title^="${prefix}"]`).first();
    if (await inside.count()) return inside.click();
  }
  throw new Error(`no ribbon button titled "${prefix}…"`);
};
const menuItem = (label) => page.click(`.rb-pop .rb-menu button:has-text("${label}")`);
const select = async (id) => {
  const el = page.locator(`.slide-wrap .hs-obj[data-el="${id}"]`);
  const b = await el.boundingBox();
  await page.mouse.click(b.x + 10, b.y + b.height - 10);
  await page.waitForTimeout(200);
  const sel = await page.evaluate(() => window.__hpptx.selection());
  assert(sel.includes(id), `selected ${id}: ${sel}`);
};

const card = (id, x, y, text, extra = {}) => ({ id, kind: "shape", shape: "roundRect", x, y, w: 380, h: 170, fill: "#dce4f2", text: `<p>${text}</p>`, fs: 32, ...extra });
const chart = { id: "ch", kind: "chart", x: 960, y: 560, w: 820, h: 400, chart: { type: "clustered-bar", labels: ["4月", "5月", "6月"], series: [{ name: "計画", values: [100, 120, 140] }, { name: "実績", values: [95, 130, 150] }],
  style: { type: "clustered-bar", dir: "col", gap: 80, font: { size: 24, color: "#1a1a1a" }, legend: { pos: "b" }, cat: { line: "#d9d9d9" }, val: { hide: true, line: "none" },
    series: [{ kind: "bar", color: "#b7c3da", label: { val: true } }, { kind: "bar", color: "#1f3864", label: { val: true } }] } } };
const deck = {
  title: "インタラクションの確認", audience: "社内", purpose: "確認", theme: "clarity", transition: "fade", motion: {}, memo: "",
  slides: [
    { type: "title", title: "インタラクションの確認" },
    { type: "blank", title: "手で付ける", elements: [
      card("a", 80, 170, "ホバーで浮く"), card("b", 500, 170, "詳細を開く"), card("c", 920, 170, "裏返す"), card("d", 1340, 170, "タブ1"),
      card("e", 1340, 380, "タブ1の中身"), card("f", 80, 380, "ずっと浮く"),
      { id: "g", kind: "text", x: 500, y: 400, w: 700, h: 100, text: "<p>タイプライターで出る文</p>", fs: 40 },
      chart,
    ] },
    { type: "blank", title: "おまかせの確認", hideTitle: true, elements: [
      { id: "t", kind: "text", x: 70, y: 60, w: 1200, h: 80, text: "<p>おまかせの確認</p>", fs: 56, bold: true },
      card("k1", 80, 200, "成果"), card("k2", 500, 200, "課題"), card("k3", 920, 200, "方針"),
      { id: "num", kind: "text", x: 80, y: 420, w: 700, h: 140, text: "<p>1,440時間</p>", fs: 96, bold: true },
      { id: "ln", kind: "line", x1: 80, y1: 620, x2: 900, y2: 620, stroke: "#1f3864", strokeW: 4, tail: "triangle" },
      { ...chart, id: "ch2", x: 960, y: 420 },
    ] },
    { type: "closing", title: "おわり" },
  ],
};

await page.goto(base);
await page.evaluate((value) => localStorage.setItem("hpptx-studio-current-v1", JSON.stringify({ deck: value, selected: 1, savedAt: new Date().toISOString() })), deck);
await page.goto(base);
await page.waitForSelector(".film-item");
await page.waitForTimeout(900);

await step("the インタラクション tab: under the mouse, a click, motion that keeps going, おまかせ", async () => {
  await goTo(2);
  await tab("インタラクション");
  for (const label of ["HTMLの動き", "反応", "説明を出す", "クリック", "動き"]) assert(await page.locator(`.rb-body .rb-btn:has-text("${label}"), .rb-body .rb-folded`).count(), `"${label}" on the tab`);
  await shot("tab");
});

await step("マウスを乗せたとき: 浮き上がる and a note", async () => {
  await select("a");
  await ribbonTitled("マウスを乗せたとき：");
  await menuItem("浮き上がる");
  await page.waitForTimeout(200);
  assert((await objOf("a")).hover === "lift", "hover set");
  await ribbonTitled("マウスを乗せると吹き出し");
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", "マウスを乗せると出る説明");
  await page.click("#askOk");
  await page.waitForTimeout(200);
  assert((await objOf("a")).tip === "マウスを乗せると出る説明", "tip set");
});

await step("クリックしたとき: details, turning over, tabs; ずっと動く", async () => {
  await select("b");
  await ribbonTitled("クリックしたとき：");
  await menuItem("詳細を開く");
  await page.waitForSelector(".rb-pop .ix-form");
  const inputs = page.locator(".rb-pop .ix-form input[type=text]");
  await inputs.nth(0).fill("内訳");
  await page.fill(".rb-pop .ix-form textarea >> nth=0", "部門ごとの削減時間です");
  await page.fill(".rb-pop .ix-form textarea >> nth=1", "営業,420\n店舗,610");
  await page.click('.rb-pop .ix-form button:has-text("設定する")');
  await page.waitForTimeout(200);
  const b = await objOf("b");
  assert(b.action?.type === "popup" && b.action.rows.length === 2 && b.action.title === "内訳", `popup: ${JSON.stringify(b.action)}`);
  await select("c");
  await ribbonTitled("クリックしたとき：");
  await menuItem("裏返す");
  await page.waitForSelector(".rb-pop .ix-form");
  await page.fill(".rb-pop .ix-form textarea", "裏の答え");
  await page.click('.rb-pop .ix-form button:has-text("設定する")');
  await page.waitForTimeout(200);
  assert((await objOf("c")).action?.type === "flip", "flip set");
  await select("d");
  await ribbonTitled("クリックしたとき：");
  await menuItem("表示・非表示");
  await page.waitForSelector(".rb-pop .ix-targets");
  await page.check('.rb-pop .ix-targets input[value="e"]');
  await page.check(".rb-pop .ix-check input");
  await page.click('.rb-pop .ix-form button:has-text("設定する")');
  await page.waitForTimeout(200);
  const d = await objOf("d");
  assert(d.action?.type === "reveal" && d.action.targets.join() === "e" && d.action.only, `reveal: ${JSON.stringify(d.action)}`);
  await select("f");
  await ribbonTitled("スライドを表示している間ずっと");
  await menuItem("ふわふわ浮く");
  await page.waitForTimeout(200);
  assert((await objOf("f")).loop === "float", "loop set");
  assert(await page.locator(".ed-ix-badge").count() >= 5, "marks on the stage show what each object does");
  await shot("badges");
});

await step("アニメーション → HTMLの効果: typewriter on the text", async () => {
  await select("g");
  await tab("アニメーション");
  await ribbonTitled("PowerPointにはない動き");
  await page.waitForSelector(".rb-pop .an-html-block");
  assert(await page.locator(".rb-pop .an-fx.an-html").count() >= 10, "the HTML-only effects");
  await shot("gallery");
  await page.click('.rb-pop .an-fx[data-fx="in:typewriter"]');
  await page.waitForTimeout(400);
  const slide = await slideNow();
  assert(slide.timeline?.some((e) => e.el === "g" && e.fx === "typewriter"), `timeline: ${JSON.stringify(slide.timeline)}`);
  await page.waitForSelector(".motion-banner", { state: "detached", timeout: 8000 }).catch(() => {});
});

await step("presenting: a lift and a note under the mouse, details without moving on, a card turning over, a tab", async () => {
  await page.keyboard.press("Escape");
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player-stage .hs-slide .hs-obj");
  await page.waitForTimeout(1600);
  const node = (id) => page.locator(`.hs-player-stage .hs-obj[data-el="${id}"]`);
  let b = await node("a").boundingBox();
  await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
  await page.waitForTimeout(500);
  const lift = await page.evaluate(() => getComputedStyle(document.querySelector('.hs-player-stage .hs-obj[data-el="a"] > .hs-obj-move')).transform);
  assert(/matrix\(1, 0, 0, 1, 0, -1\d/.test(lift), `lifted: ${lift}`);
  assert((await page.textContent(".hs-player-stage .hs-tip")).includes("マウスを乗せると出る説明"), "the note");
  await shot("present-hover");
  const count = await page.textContent(".hs-player-count");
  b = await node("b").boundingBox();
  await page.mouse.click(b.x + 20, b.y + 20);
  await page.waitForSelector(".hs-player-stage .hs-evidence");
  assert((await page.textContent(".hs-player-stage .hs-evidence")).includes("営業"), "the breakdown");
  assert(!(await page.textContent(".hs-player-stage .hs-evidence")).includes("<p>"), "words, not markup");
  await shot("present-details");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  assert(await page.textContent(".hs-player-count") === count, "the slide did not move on");
  b = await node("c").boundingBox();
  await page.mouse.click(b.x + 20, b.y + 20);
  await page.waitForTimeout(1000);
  assert(await node("c").evaluate((el) => el.classList.contains("is-flipped")), "turned over");
  assert(await node("e").evaluate((el) => getComputedStyle(el).visibility === "hidden"), "the tab's contents wait hidden");
  b = await node("d").boundingBox();
  await page.mouse.click(b.x + 20, b.y + 20);
  await page.waitForTimeout(700);
  assert(await node("e").evaluate((el) => getComputedStyle(el).visibility === "visible"), "the tab shows its contents");
  assert(await node("f").evaluate((el) => el.getAnimations().some((a) => a.animationName === "hs-ix-float")), "floating all along");
  await shot("present-flip-tab");
});

const present = async () => {
  await page.keyboard.press("Escape");
  await goTo(2);
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player-stage .hs-slide .hs-obj");
  await page.waitForTimeout(1600);
};

await step("presenting: a PowerPoint chart answers the mouse and hides a series from its legend", async () => {
  await present();
  const bar = page.locator(".hs-player-stage .hs-obj[data-el=\"ch\"] .hs-obar").nth(3);
  const b = await bar.boundingBox();
  await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
  await page.waitForTimeout(400);
  assert((await page.textContent(".hs-player-stage .hs-tip")).includes("実績"), "the value under the mouse");
  assert(await page.evaluate(() => document.querySelector(".hs-player-stage .hs-ochart").classList.contains("hs-has-hot")), "the rest dims");
  await shot("present-chart");
  const count = await page.textContent(".hs-player-count");
  await page.click('.hs-player-stage .hs-ochart-key[data-series="0"]');
  await page.waitForTimeout(400);
  const hidden = await page.evaluate(() => [...document.querySelectorAll('.hs-player-stage .hs-ochart [data-s="0"]')].every((el) => el.classList.contains("hs-off")));
  assert(hidden, "series 1 hidden");
  assert(await page.textContent(".hs-player-count") === count, "the legend click did not move on");
});

await step("presenting: the typewriter on its click", async () => {
  await present();
  await page.evaluate(() => { window.__sample = null; document.addEventListener("keydown", () => setTimeout(() => { const us = [...document.querySelectorAll('.hs-player-stage .hs-obj[data-el="g"] .hs-u')]; window.__sample = [us.filter((u) => getComputedStyle(u).visibility === "visible").length, us.length, Boolean(document.querySelector(".hs-player-stage .hs-caret"))]; }, 500), { once: true, capture: true }); });
  await page.keyboard.press("ArrowRight");
  await page.waitForFunction(() => window.__sample != null);
  const [shown, all, caret] = await page.evaluate(() => window.__sample);
  assert(all > 5 && shown > 0 && shown < all && caret, `typing: ${shown}/${all} caret=${caret}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(500);
});

await step("おまかせ: the title rises, the chart grows, the figure counts up, the line is drawn, cards lift; one undo", async () => {
  await goTo(3);
  await tab("インタラクション");
  await ribbonTitled("このスライドに、HTMLならではの動き");
  await page.waitForTimeout(500);
  const slide = await slideNow();
  const fx = Object.fromEntries((slide.timeline || []).map((e) => [e.el, e.fx]));
  assert(fx.t === "maskRise" && fx.ch2 === "chartGrow" && fx.num === "countUp" && fx.ln === "draw", `timeline: ${JSON.stringify(fx)}`);
  assert(["k1", "k2", "k3"].every((id) => slide.elements.find((o) => o.id === id).hover === "lift"), "cards lift");
  await shot("omakase");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(400);
  const back = await slideNow();
  assert(!back.timeline && !back.elements.some((o) => o.hover), "undo takes it all back");
  await page.keyboard.press("Control+Shift+z");
  await page.waitForTimeout(400);
  assert((await slideNow()).timeline?.length === 4, "redo");
});

await step("presenting おまかせ: the figure counts up as the slide arrives", async () => {
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player-stage .hs-obj[data-el=\"num\"]");
  const samples = [];
  for (let i = 0; i < 8; i += 1) { samples.push(await page.evaluate(() => document.querySelector('.hs-player-stage .hs-obj[data-el="num"]')?.textContent || "")); await page.waitForTimeout(220); }
  await page.waitForTimeout(1200);
  const end = await page.evaluate(() => document.querySelector('.hs-player-stage .hs-obj[data-el="num"]').textContent);
  assert(samples.some((t) => /^\d/.test(t) && t !== "1,440時間"), `counting: ${samples.join(" / ")}`);
  assert(end.trim() === "1,440時間", `ends on the figure: ${end}`);
  await shot("present-omakase");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

await step("export: the file carries the interactions and plays them", async () => {
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#downloadBtn").then(async () => { await page.waitForTimeout(400); if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn"); })]);
  const file = join(outDir, "interact-export.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  assert(/"hover":"lift"/.test(html) && /"type":"popup"/.test(html) && /"loop":"float"/.test(html), "in the deck data");
  const viewer = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  await viewer.goto(`file://${file}#2`);
  await viewer.waitForSelector('.hs-player .hs-obj[data-el="b"]');
  await viewer.waitForTimeout(1200);
  const b = await viewer.locator('.hs-player .hs-obj[data-el="b"]').boundingBox();
  await viewer.mouse.click(b.x + 20, b.y + 20);
  await viewer.waitForSelector(".hs-player .hs-evidence", { timeout: 4000 });
  const a = await viewer.locator('.hs-player .hs-obj[data-el="a"]').boundingBox();
  await viewer.keyboard.press("Escape");
  await viewer.mouse.move(a.x + a.width / 2, a.y + a.height / 2);
  await viewer.waitForTimeout(400);
  const tip = await viewer.textContent(".hs-player .hs-tip").catch(() => "");
  await viewer.close();
  assert(tip.includes("マウスを乗せると出る説明"), `the note in the file: ${tip}`);
});

await step("a deck brought over from PowerPoint: おまかせ from the import message", async () => {
  // おまかせ by hand from the message bar (the automatic path is studio-motion-ide.mjs).
  await page.evaluate(() => { localStorage.clear(); localStorage.setItem("hpptx-auto-html", "0"); });
  await page.goto(base);
  await page.waitForSelector("#importDeckBtn");
  await page.setInputFiles("#importDeckFile", join(decks, "template.pptx"));
  await page.waitForSelector("#importModeDialog[open]");
  await page.click('#importModeDialog [data-mode="exact"]');
  await page.waitForSelector('.callout button:has-text("HTMLの動きをおまかせで付ける")', { timeout: 30000 });
  const before = await page.evaluate(() => window.__hpptx.deck().slides.map((s) => (s.timeline || []).length));
  await page.click('.callout button:has-text("HTMLの動きをおまかせで付ける")');
  await page.waitForTimeout(600);
  const after = await page.evaluate(() => window.__hpptx.deck().slides.map((s) => (s.timeline || []).length));
  const hovers = await page.evaluate(() => window.__hpptx.deck().slides.flatMap((s) => s.elements || []).filter((o) => o.hover).length);
  assert(after.reduce((a, b) => a + b, 0) > before.reduce((a, b) => a + b, 0), `animations added: ${before} → ${after}`);
  assert(hovers > 0 || after.some((n, i) => n > before[i]), "interactions added");
  const kept = await page.evaluate(() => window.__hpptx.deck().slides.some((s) => (s.timeline || []).some((e) => e.fx === "flyIn")));
  assert(kept, "PowerPoint's own animations stay");
  // The template's marks (the logo, 社外秘, the slogan) are on every page and stay as they are.
  const touched = await page.evaluate(() => {
    const slides = window.__hpptx.deck().slides;
    const marks = new Set(slides.flatMap((s) => (s.elements || []).filter((o) => o.kind === "image" || /社外秘|未来を共につくる/.test(o.text || "")).map((o) => o.id)));
    const animated = new Set(slides.flatMap((s) => (s.timeline || []).map((e) => e.el)));
    return slides.flatMap((s) => (s.elements || []).filter((o) => marks.has(o.id) && (o.hover || o.action || animated.has(o.id))).map((o) => o.name || o.text || o.kind));
  });
  assert(!touched.length, `the template's marks are left alone: ${touched.join(" / ")}`);
  await shot("import-omakase");
});

await step("one reset removes both PowerPoint and HTML actions from every imported page", async () => {
  const before = await page.evaluate(() => structuredClone(window.__hpptx.deck()));
  await page.click("#animTab");
  await page.locator(".an-reset-actions button:has-text('全ページの動き・操作を削除')").click();
  const cleared = await page.evaluate(() => structuredClone(window.__hpptx.deck()));
  assert(cleared.slides.length === before.slides.length, "all pages remain");
  assert(cleared.slides.every((s, i) => s.elements.length === before.slides[i].elements.length && !s.timeline?.length), "objects remain and all animations are gone");
  assert(cleared.slides.every((s) => s.elements.every((o) => !o.action && !o.hover && !o.tip && typeof o.loop !== "string")), "all HTML interactions are gone");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(350);
  assert(JSON.stringify((await page.evaluate(() => window.__hpptx.deck())).slides) === JSON.stringify(before.slides), "one undo restores every original and HTML action");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
