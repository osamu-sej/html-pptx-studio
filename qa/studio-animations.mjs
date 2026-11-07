// PowerPoint-style animations in a real browser: effects from the アニメーション tab (with the automatic preview),
// another effect added, effect options, timing, the アニメーション ウィンドウ (order, delete, details), a trigger,
// motion paths (a preset whose end is dragged, one drawn by hand), a part of the layout animated, the
// 画面切り替え tab (transition, its length, moving on by itself), copy and paste with animations, undo — then
// the presentation (hidden until their click, a trigger, moving on by itself) and the exported file.
// Usage: node qa/studio-animations.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const shot = async (name) => { const file = join(outDir, `anim-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await page.waitForTimeout(150); if (await page.isVisible(".slide-wrap .hs-slide")) await measure(); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };
const slideNow = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide() ?? null)));
const objects = async () => (await slideNow())?.elements ?? [];
const timeline = async () => (await slideNow())?.timeline ?? [];
const assert = (ok, message) => { if (!ok) throw new Error(message); };
let box = null;
const at = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
const measure = async () => { box = await page.locator(".slide-wrap .hs-slide").boundingBox(); };
const drag = async ([x1, y1], [x2, y2], { steps = 8 } = {}) => { await measure(); await page.mouse.move(...at(x1, y1)); await page.mouse.down(); await page.mouse.move(...at(x2, y2), { steps }); await page.mouse.up(); await page.waitForTimeout(250); };
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
const menuItem = (label) => page.locator(`.rb-pop button:has-text("${label}")`).first().click();
const shortcut = (key) => page.keyboard.press(process.platform === "darwin" ? `Meta+${key}` : `Control+${key}`);
const selectObject = async (index) => {
  if ((await page.getAttribute("#formatTab", "aria-selected")) !== "true") await page.click("#formatTab");
  const list = await objects();
  await page.locator("#formatPane .fp-sel-list li").nth(list.length - 1 - index).click();
  await page.waitForTimeout(200);
};
const drawShape = async (title, from, to) => {
  const before = (await objects()).length;
  for (let attempt = 0; attempt < 2; attempt += 1) {
    await tab("挿入");
    await page.locator('.rb-btn.big:has-text("図形")').first().click();
    await page.click(`.rb-gallery button[title="${title}"]`);
    await drag(from, to);
    if ((await objects()).length > before) return;
  }
  throw new Error(`shape was not inserted: ${title}`);
};

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);

await step("a blank slide with two shapes and a text box", async () => {
  await page.click("#sampleDeckBtn");
  await page.waitForSelector(".film-item");
  const tabs = await page.$$eval(".rb-tabs [role=tab]", (els) => els.map((el) => el.textContent));
  assert(tabs.join() === "ホーム,挿入,描画,デザイン,画面切り替え,アニメーション,インタラクション,スライド ショー,校閲,表示,開発", `tabs: ${tabs}`);
  await ribbon("新しいスライド");
  await menuItem("白紙");
  await page.waitForTimeout(400);
  await measure();
  await drawShape("四角形: 角を丸くする", [300, 300], [700, 520]);
  await drawShape("楕円", [1300, 650], [1500, 850]);
  await tab("挿入");
  await ribbon("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  await page.mouse.click(...at(900, 300));
  await page.waitForTimeout(200);
  await page.keyboard.type("売上を伸ばす三つの柱");
  await page.mouse.click(...at(1700, 950));
  await page.waitForTimeout(300);
  const kinds = (await objects()).map((o) => o.kind).join();
  assert(kinds === "shape,shape,text", `objects: ${kinds}`);
});

await step("アニメーション tab: スライドイン from the gallery, with the automatic preview", async () => {
  await selectObject(0);
  await tab("アニメーション");
  await page.locator('.an-quick .an-fx[data-fx="in:flyIn"]').click();
  await page.waitForSelector(".motion-banner", { timeout: 3000 });
  const [e] = await timeline();
  const [shape] = await objects();
  assert(e && e.cls === "in" && e.fx === "flyIn" && e.el === shape.id && e.start === "click", `entry: ${JSON.stringify(e)}`);
  await page.waitForSelector(".motion-banner", { state: "detached", timeout: 6000 });
  assert(await page.isVisible('.ed-layer .ed-anim-badge:has-text("1")'), "order mark 1 by the shape");
  await shot("tab");
});

await step("アニメーションの追加: パルス after the entrance; 効果のオプション: 左から", async () => {
  await ribbon("アニメーションの追加");
  await page.waitForSelector(".an-gallery");
  await shot("gallery");
  await page.locator('.an-gallery .an-fx[data-fx="em:pulse"]').click();
  await page.waitForTimeout(300);
  let list = await timeline();
  assert(list.length === 2 && list[1].fx === "pulse" && list[1].start === "click", `two: ${JSON.stringify(list)}`);
  // Pick the entrance in the pane and change its direction.
  await page.click("#animTab");
  await page.locator(".an-row").first().click();
  await ribbon("効果のオプション");
  await menuItem("左から");
  await page.waitForTimeout(300);
  list = await timeline();
  assert(list[0].dir === "left", `direction: ${list[0].dir}`);
});

await step("the text box fades in after the shape (start, duration from the ribbon)", async () => {
  await selectObject(2);
  await page.locator('.an-quick .an-fx[data-fx="in:fade"]').click();
  await page.waitForTimeout(300);
  await page.selectOption('.rb-body select[aria-label="開始"]', "after");
  await page.waitForTimeout(200);
  await page.fill('.rb-body input[aria-label="継続時間（秒）"]', "1.5");
  await page.press('.rb-body input[aria-label="継続時間（秒）"]', "Enter");
  await page.waitForTimeout(300);
  const e = (await timeline()).at(-1);
  assert(e.fx === "fade" && e.start === "after" && e.dur === 1500, `text: ${JSON.stringify(e)}`);
});

await step("アニメーション ウィンドウ: order, details, delete and undo", async () => {
  await page.click("#animTab");
  const rows = await page.locator(".an-row").count();
  assert(rows === 3, `rows: ${rows}`);
  await shot("pane");
  // Move the text's fade up one place.
  await page.locator(".an-row").nth(2).click();
  await page.click('.an-actions button[title="順番を前にする"]');
  await page.waitForTimeout(300);
  let list = await timeline();
  assert(list[1].fx === "fade" && list[2].fx === "pulse", `order: ${list.map((e) => e.fx)}`);
  // The details: repeat twice, smooth end.
  await page.selectOption('.an-details select[aria-label="繰り返し"]', "2");
  await page.waitForTimeout(250);
  await page.selectOption('.an-details select[aria-label="滑らかさ"]', "out");
  await page.waitForTimeout(250);
  list = await timeline();
  assert(list[1].repeat === 2 && list[1].ease === "out", `details: ${JSON.stringify(list[1])}`);
  // Delete the pulse, then undo.
  await page.locator(".an-row").nth(2).click();
  await page.click('.an-actions button:has-text("削除")');
  await page.waitForTimeout(250);
  assert((await timeline()).length === 2, "deleted");
  await page.locator(".slide-wrap").click({ position: { x: 5, y: 5 } }).catch(() => {});
  await shortcut("z");
  await page.waitForTimeout(300);
  assert((await timeline()).length === 3, "undo brings it back");
});

await step("a trigger: clicking the circle plays the text's fade", async () => {
  await page.locator(".an-row").nth(1).click();
  await page.selectOption('.an-details select[aria-label="トリガー"]', { index: 2 });
  await page.waitForTimeout(300);
  const [, circle] = await objects();
  const fade = (await timeline()).find((e) => e.fx === "fade");
  assert(fade.trigger === circle.id, `trigger: ${JSON.stringify(fade)} circle ${circle.id}`);
  assert(await page.isVisible(".an-trigger-head"), "the pane shows the trigger's sequence");
  assert(await page.isVisible('.ed-layer .ed-anim-badge:has-text("⚡")'), "a ⚡ mark by the text");
});

await step("motion paths: 直線（右へ）, its end dragged; a path drawn by hand", async () => {
  await selectObject(1);
  await ribbon("アニメーションの追加");
  await page.locator('.an-gallery .an-fx[data-fx="path:lineRight"]').click();
  await page.waitForTimeout(400);
  let e = (await timeline()).at(-1);
  assert(e.cls === "path" && e.path.pts.length === 2 && e.path.pts[1][0] > 200, `path: ${JSON.stringify(e)}`);
  // Pick it so its points can be dragged; drag the end up.
  await page.click("#animTab");
  await page.locator(`.an-row[data-id="${e.id}"]`).click();
  await page.waitForTimeout(300);
  const end = page.locator(".ed-layer .ed-anim-pt[data-handle]").last();
  const b = await end.boundingBox();
  await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
  await page.mouse.down();
  await page.mouse.move(b.x + b.width / 2, b.y - 120, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  e = (await timeline()).at(-1);
  assert(e.path.pts[1][1] < -100, `end dragged up: ${JSON.stringify(e.path.pts)}`);
  await shot("path");
  // A path of one's own: three clicks and a double-click.
  await ribbon("アニメーションの追加");
  await page.locator('.an-gallery .an-fx[data-fx="path:custom"]').click();
  await page.waitForSelector(".ed-pathdraw");
  await page.mouse.click(...at(1200, 600));
  await page.mouse.click(...at(1000, 450));
  await page.mouse.dblclick(...at(800, 650));
  await page.waitForTimeout(400);
  e = (await timeline()).at(-1);
  assert(e.fx === "custom" && e.path.pts.length >= 3, `drawn: ${JSON.stringify(e.path)}`);
});

await step("copy and paste keep the animations (pointed at the copy)", async () => {
  await selectObject(0);
  const before = (await timeline()).length;
  await shortcut("c");
  await shortcut("v");
  await page.waitForTimeout(400);
  const list = await objects();
  const copy = list.at(-1);
  const anims = (await timeline()).filter((e) => e.el === copy.id);
  assert(anims.length === 2 && (await timeline()).length === before + 2, `copied animations: ${anims.length}`);
  await shortcut("z");
  await page.waitForTimeout(300);
  assert((await timeline()).length === before && (await objects()).length === 3, "undo removes both");
});

await step("画面切り替え: 押し上げ, 1.5 s, moving on by itself after 2 s", async () => {
  await tab("画面切り替え");
  await page.click('.an-tr[data-tr="push"]');
  await page.waitForSelector(".transition-preview", { timeout: 3000 });
  await page.waitForTimeout(2600);
  await page.fill('.rb-body input[aria-label="切り替えにかける時間（秒）"]', "1.5");
  await page.press('.rb-body input[aria-label="切り替えにかける時間（秒）"]', "Enter");
  await page.waitForTimeout(200);
  await page.locator('.rb-body label:has-text("自動的に切り替え") input').check();
  await page.waitForTimeout(200);
  await page.fill('.rb-body input[aria-label="自動的に切り替えるまでの秒数"]', "2");
  await page.press('.rb-body input[aria-label="自動的に切り替えるまでの秒数"]', "Enter");
  await page.waitForTimeout(300);
  const s = await slideNow();
  assert(s.transition === "push" && s.transitionDur === 1500 && s.advance === 2, `slide: ${s.transition} ${s.transitionDur} ${s.advance}`);
  await shot("transition");
});

await step("a part of the layout: the next slide's title fades in", async () => {
  await page.locator(".film-item").nth(3).click();
  await page.waitForTimeout(400);
  await tab("アニメーション");
  await page.click("#animTab");
  await page.click('.an-actions button:has-text("レイアウトの部品")');
  await menuItem("タイトル");
  await page.waitForTimeout(400);
  const [e] = await timeline();
  assert(e?.el === "@title" && e.fx === "fade", `layout part: ${JSON.stringify(e)}`);
  // The automatic preview plays it first; then the editor shows its order mark.
  await page.waitForSelector(".motion-banner", { state: "detached", timeout: 6000 });
  assert(await page.isVisible(".ed-layer .ed-anim-badge"), "its mark on the stage");
  await page.locator(".film-item").nth(1).click();
  await page.waitForTimeout(400);
});

await step("presenting: hidden until its click, the trigger plays the text, then the slide moves on by itself", async () => {
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player .hs-slide .hs-objects .hs-obj");
  await page.waitForTimeout(900);
  const hidden = await page.evaluate(() => [...document.querySelectorAll(".hs-player .hs-slide .hs-obj")].map((el) => getComputedStyle(el.querySelector(".hs-obj-fx")).visibility));
  assert(hidden[0] === "hidden" && hidden[2] === "hidden" && hidden[1] === "visible", `before the first click: ${hidden}`);
  assert(await page.locator(".hs-player .hs-obj[data-trigger]").count() === 1, "the circle is a trigger");
  await page.click(".hs-player-stage", { position: { x: 40, y: 40 } });
  await page.waitForTimeout(900);
  const shown = await page.evaluate(() => getComputedStyle(document.querySelector(".hs-player .hs-slide .hs-obj .hs-obj-fx")).visibility);
  assert(shown === "visible", "the shape flew in");
  await page.locator(".hs-player .hs-obj[data-trigger]").click();
  await page.waitForTimeout(1600);
  const text = await page.evaluate(() => getComputedStyle(document.querySelectorAll(".hs-player .hs-slide .hs-obj")[2].querySelector(".hs-obj-fx")).visibility);
  assert(text === "visible", "the trigger showed the text");
  const count0 = await page.textContent(".hs-player-count");
  await shot("present");
  // The remaining clicks play by themselves after 2 s, then the next slide comes.
  await page.waitForFunction((c) => document.querySelector(".hs-player-count")?.textContent !== c, count0, { timeout: 15000 });
  await page.keyboard.press("Escape");
  await page.waitForTimeout(500);
  await page.locator(".film-item").nth(1).click();
  await page.waitForTimeout(300);
});

await step("export: the file carries the animations and plays them", async () => {
  // Without the 2-second 自動的に切り替え (tested above), only the click shows the shape: nothing races the check.
  await tab("画面切り替え");
  await page.locator('.rb-body label:has-text("自動的に切り替え") input').uncheck();
  await page.waitForTimeout(300);
  assert((await slideNow()).advance == null, "no automatic advance on this slide");
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#downloadBtn").then(async () => { if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn"); })]);
  const file = join(outDir, "anim-export.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  assert(/timelinePlan/.test(html), "animate.js inside");
  assert(/"timeline":\[/.test(html), "the animations in the deck data");
  const viewer = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  await viewer.goto(`file://${file}#2`);
  await viewer.waitForSelector(".hs-player .hs-objects .hs-obj");
  await viewer.waitForTimeout(800);
  const before = await viewer.evaluate(() => getComputedStyle(document.querySelector(".hs-player .hs-slide .hs-obj .hs-obj-fx")).visibility);
  await viewer.keyboard.press("ArrowRight");
  await viewer.waitForTimeout(900);
  const after = await viewer.evaluate(() => getComputedStyle(document.querySelector(".hs-player .hs-slide .hs-obj .hs-obj-fx")).visibility);
  await viewer.close();
  assert(before === "hidden" && after === "visible", `exported: ${before} → ${after}`);
});

await step("every page's animations and click actions can be reset in one undo step", async () => {
  const before = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
  assert(before.slides.some((slide) => slide.timeline?.length), "the fixture has effects before reset");
  await page.click("#animTab");
  await page.locator(".an-reset-actions button:has-text('全ページの動き・操作を削除')").click();
  const cleared = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
  assert(cleared.transition === "none", "deck transitions are disabled");
  assert(cleared.slides.every((slide) => !slide.timeline?.length && slide.animation === "none" && slide.transition === "none" && slide.advance == null), "all page actions are cleared");
  assert(cleared.slides.map((slide) => slide.title).join("|") === before.slides.map((slide) => slide.title).join("|"), "content remains");
  await page.locator(".film-item").first().click();
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player .hs-slide");
  const coverMotion = await page.evaluate(() => {
    const slide = document.querySelector(".hs-player .hs-slide");
    return { ambient: slide.dataset.ambient, numbers: slide.classList.contains("hs-numbers"), backdrop: slide.dataset.backdrop || "" };
  });
  assert(coverMotion.ambient === "off" && !coverMotion.numbers && !coverMotion.backdrop, `cover still: ${JSON.stringify(coverMotion)}`);
  await page.keyboard.press("Escape");
  await shortcut("z");
  await page.waitForTimeout(300);
  const restored = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
  assert(JSON.stringify(restored.slides) === JSON.stringify(before.slides), "one undo restores every page");
  await page.locator(".an-reset-actions button:has-text('AIにHTML演出を相談')").click();
  assert((await page.inputValue("#chatInput")).includes("HTMLならでは"), "AI prompt starts from HTML rather than old PowerPoint effects");
});

await step("最近使用した図形; 書式のコピー (ダブルクリックで続けて); アニメーションの後 薄くする・サウンド", async () => {
  await page.locator(".film-item").nth(1).click();
  await tab("ホーム");
  await ribbon("新しいスライド");
  await menuItem("白紙");
  await page.waitForSelector(".slide-wrap .hs-slide");
  await drawShape("楕円", [300, 300], [700, 600]);
  await drawShape("正方形/長方形", [900, 300], [1200, 500]);
  await drawShape("正方形/長方形", [900, 650], [1200, 850]);
  // The gallery shows the shapes used, the last first.
  await tab("挿入");
  await page.locator('.rb-btn.big:has-text("図形")').first().click();
  const recent = await page.$$eval(".rb-gallery .rb-recent-shapes button", (els) => els.map((el) => el.dataset.shape));
  assert(recent[0] === "rect" && recent.includes("ellipse"), `最近使用した図形: ${recent}`);
  await page.keyboard.press("Escape");
  // The ellipse in orange (a standard colour), then its format onto both boxes with one double-click.
  const [ellipse, a, b] = await objects();
  await selectObject(0);
  await tab("図形の書式");
  await page.locator('.rb-body .rb-drop:has-text("塗り")').first().click().catch(async () => { await page.locator('.rb-body .rb-drop[title*="塗りつぶし"]').first().click(); });
  await page.locator('.rb-pop .rb-sw[aria-label="オレンジ"]').first().click();
  await page.waitForTimeout(250);
  assert((await objects())[0].fill === "#ffc000", `the ellipse orange: ${(await objects())[0].fill}`);
  await tab("ホーム");
  await page.locator('.rb-body .rb-btn[title^="書式のコピー/貼り付け"]').first().dblclick();
  await page.waitForTimeout(200);
  const click = async (o) => { await page.mouse.click(...at(o.x + o.w / 2, o.y + o.h / 2)); await page.waitForTimeout(250); };
  await click(a);
  await click(b);
  let list = await objects();
  assert(list[1].fill === "#ffc000" && list[2].fill === "#ffc000", `both boxes: ${list[1].fill} ${list[2].fill}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  // アニメーションの後: the ellipse fades in, then fades back; a sound as it starts.
  await selectObject(0);
  await tab("アニメーション");
  await page.locator('.an-quick .an-fx[data-fx="in:fade"]').click();
  await page.waitForTimeout(300);
  await page.click("#animTab");
  await page.locator(".an-row").first().click();
  await page.selectOption('.an-details select[aria-label="アニメーションの後"]', "dim");
  await page.selectOption('.an-details select[aria-label="サウンド"]', "chime");
  await page.waitForTimeout(300);
  const [e] = await timeline();
  assert(e.after === "dim" && e.sound === "chime", `the entry: ${JSON.stringify(e)}`);
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player .hs-slide .hs-objects .hs-obj");
  await page.waitForTimeout(800);
  await page.click(".hs-player-stage", { position: { x: 40, y: 40 } });
  await page.waitForTimeout(1800);
  const opacity = await page.evaluate((id) => getComputedStyle(document.querySelector(`.hs-player .hs-obj[data-el="${id}"] .hs-obj-fx`)).opacity, ellipse.id);
  assert(Math.abs(Number(opacity) - 0.35) < 0.02, `faded after its entrance: ${opacity}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
