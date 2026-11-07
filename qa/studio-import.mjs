// PowerPoint import in a real browser: 既存の資料を取り込む → 見た目どおりに取り込む brings a deck
// over as source pages with their own master marks and editable objects (title boxes, shapes, bullets,
// a table, and a chart with its original colours and labels). Animations and transitions are retained;
// presenting plays them on click, and the exported file carries them. HTMLレイアウトに組み直す
// still rebuilds a deck into layouts, and closing the choice imports nothing.
// Usage: node qa/studio-import.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
execFileSync(process.env.PYTHON_BIN || "python3", [join(root, "test", "pptx_fixtures.py"), decks], { stdio: "inherit" });

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const shot = async (name) => { const file = join(outDir, `import-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
  for (let i = 0; i < 3 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deckNow = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck() ?? null)));
const goTo = async (n) => { await page.click(`.film-item:nth-child(${n})`); await page.waitForTimeout(500); };
const text = (html) => String(html || "").replace(/<[^>]+>/g, "");

await page.goto(base);
await page.evaluate(() => { localStorage.clear(); localStorage.setItem("hpptx-auto-html", "0"); });
await page.goto(base);
await page.waitForSelector("#importDeckBtn");
await page.waitForTimeout(600);

await step("choosing a PowerPoint asks how to bring it over; closing the question imports nothing", async () => {
  await page.setInputFiles("#importDeckFile", join(decks, "template.pptx"));
  await page.waitForSelector("#importModeDialog[open]");
  assert((await page.textContent("#importModeName")).includes("template.pptx"), "names the file");
  assert(await page.isVisible('#importModeDialog [data-mode="exact"]') && await page.isVisible('#importModeDialog [data-mode="layout"]'), "two ways");
  await shot("choice");
  await page.click("#importModeDialog [data-close]");
  await page.waitForTimeout(500);
  assert(!(await deckNow()), "no deck after closing");
});

await step("見た目どおりに取り込む: every page and its original master become editable objects", async () => {
  await page.setInputFiles("#importDeckFile", join(decks, "template.pptx"));
  await page.waitForSelector("#importModeDialog[open]");
  await page.click('#importModeDialog [data-mode="exact"]');
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length === 4, null, { timeout: 30000 });
  await page.waitForTimeout(800);
  const deck = await deckNow();
  assert(deck.slides.every((s) => s.type === "blank" && s.hideTitle), "白紙 pages that draw their own title");
  assert(deck.slides.every((s) => s.master === "source"), "the original master is preserved");
  assert(deck.slides[2].hidden && JSON.stringify(deck.slides[2]).includes("隠したページ"), "the hidden slide is retained for editing");
  assert(await page.locator(".film-item.is-hidden").count() === 1, "the hidden slide is marked in the editor");
  const callout = await page.textContent(".callout");
  assert(/見た目どおりに取り込みました/.test(callout) && /アニメーション2件/.test(callout) && /非表示のスライド1枚/.test(callout), `callout: ${callout}`);
  await shot("cover");
});

await step("the page looks as it did: title box, shapes, bullets, table and a chart in its own colours", async () => {
  await goTo(2);
  const slide = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide())));
  assert(await page.locator(".slide-wrap .hs-title").count() === 0, "no studio title over the slide's own");
  const words = slide.elements.map((o) => text(o.text)).join(" ");
  for (const w of ["改革の全体像", "現状の課題", "30分", "端末で発注する", "社内ポータル"]) assert(words.includes(w), `"${w}"`);
  assert(/社外秘|未来を共につくる/.test(words), "the original template marks are editable objects");
  assert(await page.locator('.slide-wrap .hs-obj-tx p[data-bullet="■"]').count() === 3, "bullets");
  assert(await page.locator(".slide-wrap .hs-obj table td").count() === 9, "the table, cell by cell");
  const bars = await page.evaluate(() => [...document.querySelectorAll(".slide-wrap .hs-ochart .hs-obar")].map((b) => b.getAttribute("fill")));
  assert(bars.join() === "#1f3864,#b7c3da,#b7c3da", `bars: ${bars}`);
  const labels = await page.evaluate(() => [...document.querySelectorAll(".slide-wrap .hs-ochart text")].map((t) => t.textContent));
  for (const w of ["42h", "30h", "18h", "A店"]) assert(labels.includes(w), `chart label ${w}`);
  // Each object stands where its data says.
  const misplaced = await page.evaluate(() => {
    const el = document.querySelector(".slide-wrap .hs-slide");
    const r = el.getBoundingClientRect();
    const k = r.width / 1920;
    return window.__hpptx.slide().elements.filter((o) => o.kind !== "line" && !o.rot).filter((o) => {
      const b = el.querySelector(`.hs-obj[data-el="${o.id}"]`).getBoundingClientRect();
      return Math.abs((b.left - r.left) / k - o.x) > 2 || Math.abs((b.top - r.top) / k - o.y) > 2 || Math.abs(b.width / k - o.w) > 2;
    }).map((o) => o.id);
  });
  assert(!misplaced.length, `out of place: ${misplaced}`);
  await shot("page");
});

await step("アニメーションを編集 opens the animation pane with PowerPoint's animations", async () => {
  await page.click('.callout button:has-text("アニメーションを編集")');
  await page.waitForTimeout(500);
  assert(await page.isVisible("#animPane"), "the animation pane");
  assert(await page.locator('.rb-tabs [role=tab][aria-selected="true"]:has-text("アニメーション")').count() === 1, "the animation tab");
  const slide = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide())));
  assert(slide.timeline.length === 2 && slide.timeline[0].fx === "fade" && slide.timeline[1].fx === "flyIn", `timeline: ${JSON.stringify(slide.timeline)}`);
  assert(slide.transition === "fade", "the transition");
  await shot("animations");
});

await step("imported objects edit like any other: select, nudge, undo", async () => {
  const before = await page.evaluate(() => window.__hpptx.slide().elements.find((o) => o.name === "課題"));
  const el = page.locator(`.slide-wrap .hs-obj[data-el="${before.id}"]`);
  const box = await el.boundingBox();
  await page.mouse.click(box.x + 12, box.y + box.height - 12);
  await page.waitForTimeout(200);
  const sel = await page.evaluate(() => window.__hpptx.selection());
  assert(sel.includes(before.id), `selected: ${sel}`);
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(300);
  const moved = await page.evaluate((id) => window.__hpptx.slide().elements.find((o) => o.id === id).x, before.id);
  assert(moved > before.x, `moved: ${before.x} → ${moved}`);
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  const back = await page.evaluate((id) => window.__hpptx.slide().elements.find((o) => o.id === id).x, before.id);
  assert(Math.abs(back - before.x) < 0.01, "undo puts it back");
  await page.keyboard.press("Escape");
});

await step("グラフのデザイン → 書式をリセット gives the chart the studio's look; undo brings PowerPoint's back", async () => {
  const id = await page.evaluate(() => window.__hpptx.slide().elements.find((o) => o.kind === "chart").id);
  const box = await page.locator(`.slide-wrap .hs-obj[data-el="${id}"]`).boundingBox();
  await page.mouse.click(box.x + box.width - 20, box.y + 20);
  await page.waitForTimeout(300);
  await page.click('.rb-tabs [role=tab]:has-text("グラフのデザイン")');
  await page.waitForTimeout(200);
  let reset = page.locator('.rb-body .rb-btn:has-text("リセット")').first();
  if (!(await reset.count()) || !(await reset.isVisible())) {
    for (const folded of await page.locator(".rb-body .rb-folded").all()) {
      await folded.click();
      reset = page.locator('.rb-pop .rb-fold .rb-btn:has-text("リセット")').first();
      if (await reset.count()) break;
    }
  }
  await reset.click();
  await page.waitForTimeout(400);
  assert(await page.locator(`.slide-wrap .hs-obj[data-el="${id}"] .hs-ochart`).count() === 0, "PowerPoint's look is gone");
  assert(await page.locator(`.slide-wrap .hs-obj[data-el="${id}"] .hs-chart svg`).count() === 1, "the studio's chart");
  await page.keyboard.press("Escape");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(400);
  assert(await page.locator(`.slide-wrap .hs-obj[data-el="${id}"] .hs-ochart`).count() === 1, "undo brings it back");
});

await step("presenting plays PowerPoint's animation on its click", async () => {
  await goTo(2);
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player .hs-slide .hs-objects .hs-obj");
  await page.waitForTimeout(1200);
  const id = await page.evaluate(() => window.__hpptx.slide().timeline[0].el);
  const vis = () => page.evaluate((el) => getComputedStyle(document.querySelector(`.hs-player .hs-slide .hs-obj[data-el="${el}"] .hs-obj-fx`)).visibility, id);
  assert(await vis() === "hidden", "hidden before its click");
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(1600);
  assert(await vis() === "visible", "shown after the click");
  await shot("present");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(500);
});

await step("a button from PowerPoint still jumps to its slide while presenting", async () => {
  await goTo(4);
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player .hs-slide .hs-obj[data-action]", { timeout: 8000 }).catch(() => {});
  await page.waitForTimeout(1000);
  const count = () => page.textContent(".hs-player-count");
  assert(/^\s*3\b/.test(await count()), `on slide 3: ${await count()}`);
  const id = await page.evaluate(() => window.__hpptx.slide().elements.find((o) => o.action?.type === "slide")?.id);
  assert(id, "the button carries its jump");
  await page.click(`.hs-player .hs-slide .hs-obj[data-el="${id}"]`);
  await page.waitForTimeout(1500);
  assert(/^\s*2\b/.test(await count()), `went to slide 2: ${await count()}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(500);
});

await step("export: the file carries the objects, the chart and the animation", async () => {
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#downloadBtn").then(async () => { await page.waitForTimeout(400); if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn"); })]);
  const file = join(outDir, "import-export.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  assert(/"hideTitle":true/.test(html) && /"kind":"chart"/.test(html) && /"timeline":\[/.test(html), "in the deck data");
  const viewer = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  await viewer.goto(`file://${file}#2`);
  await viewer.waitForSelector(".hs-player .hs-ochart .hs-obar");
  const count = await viewer.locator(".hs-player .hs-obj").count();
  await viewer.close();
  assert(count >= 8, `objects in the file: ${count}`);
});

await step("one reset removes every imported action but keeps every source page and object", async () => {
  const before = await deckNow();
  await page.click("#animTab");
  await page.locator(".an-reset-actions button:has-text('全ページの動き・操作を削除')").click();
  const after = await deckNow();
  assert(after.slides.length === before.slides.length && after.slides[2].hidden, "hidden source pages remain");
  assert(after.slides.every((slide, i) => slide.master === "source" && slide.elements.length === before.slides[i].elements.length), "all original objects and masters remain");
  assert(after.slides.every((slide) => !slide.timeline && !slide.elements.some((object) => object.action)), "animations and click jumps are removed");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(350);
  assert(JSON.stringify((await deckNow()).slides) === JSON.stringify(before.slides), "one undo restores all imported actions");
});

await step("a 4:3 deck in another template: centred on the page, its pictures kept in the browser, its pie chart drawn", async () => {
  await page.evaluate(() => { localStorage.clear(); localStorage.setItem("hpptx-auto-html", "0"); });
  await page.goto(base);
  await page.waitForSelector("#importDeckBtn");
  await page.setInputFiles("#importDeckFile", join(decks, "plain43.pptx"));
  await page.waitForSelector("#importModeDialog[open]");
  await page.click('#importModeDialog [data-mode="exact"]');
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length === 3, null, { timeout: 30000 });
  await goTo(3);
  await page.waitForTimeout(800);
  const slide = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide())));
  const picture = slide.elements.find((o) => o.kind === "image");
  assert(picture && /^idb:/.test(picture.src), `the picture is stored in the browser: ${picture?.src?.slice(0, 30)}`);
  const loaded = await page.evaluate((id) => document.querySelector(`.slide-wrap .hs-obj[data-el="${id}"] img`)?.naturalWidth || 0, picture.id);
  assert(loaded > 0, "the picture shows");
  assert(picture.x >= 240, `inside the 4:3 page: x=${picture.x}`);
  assert(await page.locator(".slide-wrap .hs-ochart .hs-oslice").count() === 3, "a slice per category");
  await shot("plain");
});

await step("HTMLレイアウトに組み直す still rebuilds a deck into the studio's layouts", async () => {
  await page.evaluate(() => { localStorage.clear(); localStorage.setItem("hpptx-auto-html", "0"); });
  await page.goto(base);
  await page.waitForSelector("#importDeckBtn");
  await page.setInputFiles("#importDeckFile", join(decks, "plain43.pptx"));
  await page.waitForSelector("#importModeDialog[open]");
  await page.click('#importModeDialog [data-mode="layout"]');
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length >= 2, null, { timeout: 30000 });
  await page.waitForTimeout(600);
  const deck = await deckNow();
  assert(deck.slides[0].type === "title", `the cover is a layout: ${deck.slides[0].type}`);
  assert(deck.slides.some((s) => s.type !== "blank" && s.type !== "title" && s.type !== "closing"), "layouts, not 白紙 pages");
  assert(!/見た目どおり/.test(await page.textContent(".callout")), "the rebuild's own message");
  await shot("layout");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
