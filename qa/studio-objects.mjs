// The PowerPoint-style editor in a real browser: a blank slide built from objects — draw shapes, type in them,
// move / resize / rotate with the mouse, text boxes, arrows, photos and icons from the ribbon, align and group,
// the clipboard, the 書式 pane, undo, the selection pane, locking, the right-click menu — then reload, present
// and export, and check the objects come along everywhere.
// Usage: node qa/studio-objects.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const shot = async (name) => { const file = join(outDir, `objects-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await page.waitForTimeout(150); if (await page.isVisible(".slide-wrap .hs-slide")) await measure(); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };
const saved = () => page.evaluate(() => JSON.parse(localStorage.getItem("hpptx-studio-current-v1") || "null"));
const objects = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide()?.elements ?? [])));
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const near = (a, b, tolerance = 6) => Math.abs(a - b) <= tolerance;
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
// Pick an object from the 選択ウィンドウ (the photo sits on top of the shape, as PowerPoint centres new pictures).
const pickFromPane = async (kind) => {
  if ((await page.getAttribute("#formatTab", "aria-selected")) !== "true") await page.click("#formatTab");
  const list = await objects();
  const index = list.findIndex((o) => o.kind === kind);
  await page.locator("#formatPane .fp-sel-list li").nth(list.length - 1 - index).click();
  await page.waitForTimeout(200);
};

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);

await step("a blank slide from the ribbon (白紙)", async () => {
  await page.click("#sampleDeckBtn");
  await page.waitForSelector(".film-item");
  await page.waitForSelector("#ribbon:not([hidden]) .rb-tabs");
  const tabs = await page.$$eval(".rb-tabs [role=tab]", (els) => els.map((el) => el.textContent));
  assert(tabs.join() === "ホーム,挿入,描画,デザイン,画面切り替え,アニメーション,インタラクション,スライド ショー,校閲,表示,開発", `tabs: ${tabs}`);
  await ribbon("新しいスライド");
  await menuItem("白紙");
  await page.waitForTimeout(400);
  assert((await page.getAttribute(".slide-wrap .hs-slide", "data-type")) === "blank", "not a blank slide");
  assert(await page.isVisible('.slide-wrap .hs-title[data-placeholder="タイトルを入力"]'), "title placeholder");
  await measure();
});

await step("draw a rounded rectangle from the shape gallery", async () => {
  for (let attempt = 0; attempt < 2 && !(await objects()).length; attempt += 1) {
    await tab("挿入");
    await page.locator('.rb-btn.big:has-text("図形")').first().click();
    await page.waitForSelector(".rb-gallery.shapes");
    if (!attempt) await shot("gallery");
    await page.click('.rb-gallery button[title="四角形: 角を丸くする"]');
    await drag([300, 300], [700, 500]);
  }
  const [o] = await objects();
  assert(o?.kind === "shape" && o.shape === "roundRect", `drawn: ${JSON.stringify(o)}`);
  assert(near(o.x, 300) && near(o.y, 300) && near(o.w, 400) && near(o.h, 200), `box ${o.x},${o.y},${o.w},${o.h}`);
  assert(await page.isVisible(".ed-layer .ed-handle.ed-se"), "selection handles");
  assert(await page.isVisible('.rb-tabs [role=tab]:has-text("図形の書式")'), "contextual tab");
});

await step("double-click to type in the shape", async () => {
  await page.mouse.dblclick(...at(500, 400));
  await page.keyboard.type("売上を伸ばす");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const [o] = await objects();
  assert(o.text === "<p>売上を伸ばす</p>", `text: ${o.text}`);
  assert((await page.textContent('.slide-wrap .hs-obj[data-el] .hs-obj-tx')).includes("売上を伸ばす"), "text on the slide");
});

await step("move by dragging (snaps to the slide centre)", async () => {
  await drag([500, 400], [958, 610]);
  const [o] = await objects();
  assert(near(o.x + o.w / 2, 960, 1), `centred: ${o.x + o.w / 2}`);
  assert(near(o.y, 510, 12), `y ${o.y}`);
});

await step("resize from a corner and rotate with the handle", async () => {
  let [o] = await objects();
  const se = await page.locator(".ed-layer .ed-handle.ed-se").boundingBox();
  await page.mouse.move(se.x + se.width / 2, se.y + se.height / 2);
  await page.mouse.down();
  await page.mouse.move(se.x + 80, se.y + 40, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(250);
  const [r] = await objects();
  assert(r.w > o.w + 50 && r.h > o.h + 20, `resized ${o.w}×${o.h} → ${r.w}×${r.h}`);
  o = r;
  const rot = await page.locator(".ed-layer .ed-rotate").boundingBox();
  await page.mouse.move(rot.x + rot.width / 2, rot.y + rot.height / 2);
  await page.mouse.down();
  await page.mouse.move(rot.x + 120, rot.y + 40, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(250);
  const [turned] = await objects();
  assert(turned.rot && Math.abs(turned.rot) > 5, `rotation ${turned.rot}`);
  await shot("shape");
  await shortcut("z");
  await page.waitForTimeout(250);
  const [back] = await objects();
  assert(!back.rot, "undo the rotation");
});

await step("a text box: click, type, it grows with the text", async () => {
  await tab("挿入");
  await ribbon("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  await page.mouse.click(...at(120, 180));
  await page.waitForTimeout(250);
  await page.keyboard.type("1行目");
  await page.keyboard.press("Enter");
  await page.keyboard.type("2行目");
  await page.mouse.click(...at(1500, 850));
  await page.waitForTimeout(400);
  const text = (await objects()).find((o) => o.kind === "text");
  assert(text && /1行目.*2行目/.test(text.text), `text box: ${JSON.stringify(text)}`);
  assert(text.h > 100, `grows with two lines: ${text.h}`);
});

await step("an arrow, a photo and an icon from the ribbon", async () => {
  await tab("挿入");
  await ribbon("矢印");
  await drag([200, 900], [700, 800]);
  const line = (await objects()).find((o) => o.kind === "line");
  assert(line && line.tail === "triangle", `arrow: ${JSON.stringify(line)}`);
  await tab("挿入");
  await ribbon("画像");
  await menuItem("内蔵の写真");
  await page.waitForSelector(".rb-gallery.photos");
  await page.locator(".rb-gallery.photos button").first().click();
  await page.waitForFunction(() => window.__hpptx.slide().elements.some((o) => o.kind === "image"), null, { timeout: 5000 });
  assert(await page.isVisible('.rb-tabs [role=tab][aria-selected="true"]:has-text("図の形式")'), "the picture tab opens");
  await tab("挿入");
  await page.locator('.rb-btn.big:has-text("アイコン")').click();
  await page.waitForSelector(".rb-gallery.icons");
  await page.locator(".rb-gallery-grid.icons button").first().click();
  await page.waitForTimeout(300);
  const kinds = (await objects()).map((o) => o.kind).sort().join();
  assert(kinds === "icon,image,line,shape,text", `kinds: ${kinds}`);
  await shot("inserted");
});

await step("select with a box, align left, group and ungroup", async () => {
  await page.keyboard.press("Escape");
  await drag([60, 140], [1900, 1000]);
  const selected = await page.$$eval(".ed-layer .ed-outline, .ed-layer .ed-line-outline", (els) => els.length);
  assert(selected === 5, `selected by the box: ${selected}`);
  await tab("ホーム");
  await ribbon("配置");
  await menuItem("左揃え");
  await page.waitForTimeout(300);
  const lefts = (await objects()).filter((o) => o.kind !== "line").map((o) => o.x);
  assert(Math.max(...lefts) - Math.min(...lefts) < 1, `left edges: ${lefts}`);
  await shortcut("z");
  await page.waitForTimeout(250);
  await drag([60, 140], [1900, 1000]);
  await shortcut("g");
  await page.waitForTimeout(250);
  const groups = new Set((await objects()).map((o) => o.group));
  assert(groups.size === 1 && [...groups][0], `one group: ${[...groups]}`);
  await page.keyboard.press("Escape");
  await page.mouse.click(...at(130, 205));
  await page.waitForTimeout(200);
  assert((await page.$$eval(".ed-layer .ed-outline, .ed-layer .ed-line-outline", (els) => els.length)) === 5, "a click selects the whole group");
  await shot("group");
  await page.keyboard.press(process.platform === "darwin" ? "Meta+Shift+G" : "Control+Shift+G");
  await page.waitForTimeout(250);
  assert((await objects()).every((o) => !o.group), "ungrouped");
});

await step("clipboard: copy, paste, duplicate, delete", async () => {
  await page.keyboard.press("Escape");
  await page.keyboard.press("Escape");
  const before = (await objects()).length;
  await pickFromPane("shape");
  await page.locator(".slide-wrap").focus().catch(() => {});
  await shortcut("c");
  await shortcut("v");
  await page.waitForTimeout(300);
  let list = await objects();
  assert(list.length === before + 1 && list.at(-1).kind === "shape", "pasted a shape on top");
  assert(near(list.at(-1).x, list.find((o) => o.kind === "shape").x + 20, 1), "pasted with an offset");
  await shortcut("d");
  await page.waitForTimeout(300);
  list = await objects();
  assert(list.length === before + 2 && list.filter((o) => o.kind === "shape").length === 3, "duplicated");
  await page.keyboard.press("Delete");
  await page.waitForTimeout(300);
  assert((await objects()).length === before + 1, "deleted");
  await pickFromPane("shape");
  const pasted = (await objects()).filter((o) => o.kind === "shape").at(-1).id;
  await page.locator(`#formatPane .fp-sel-list li[data-id="${pasted}"]`).click();
  await page.keyboard.press("Delete");
  await page.waitForTimeout(300);
  assert((await objects()).length === before, "back to where we were");
});

await step("format: fill from the 書式 pane, size and bold from the ribbon, undo", async () => {
  await pickFromPane("shape");
  await page.waitForSelector('#formatPane [data-fp="fill"]');
  await page.click('#formatPane [data-fp="fill"] .fp-sw[aria-label="オレンジ"]');
  await page.waitForTimeout(250);
  assert((await objects()).find((o) => o.kind === "shape").fill === "#ffc000", "fill");
  await tab("ホーム");
  await page.fill(".rb-size", "28");
  await page.press(".rb-size", "Enter");
  await page.waitForTimeout(250);
  assert((await objects()).find((o) => o.kind === "shape").fs === 56, "28pt = 56px");
  await pickFromPane("shape");
  await shortcut("b");
  await page.waitForTimeout(250);
  assert((await objects()).find((o) => o.kind === "shape").bold === true, "bold");
  await shot("formatted");
  await shortcut("z");
  await page.waitForTimeout(250);
  assert(!(await objects()).find((o) => o.kind === "shape").bold, "undo bold");
});

await step("selection pane: hide and show; lock keeps it still", async () => {
  await page.waitForSelector("#formatPane .fp-sel-list li");
  const count = await page.locator("#formatPane .fp-sel-list li").count();
  assert(count === (await objects()).length, `rows ${count}`);
  await page.locator("#formatPane .fp-sel-list li").first().locator(".fp-sel-eye").click();
  await page.waitForTimeout(250);
  assert((await objects()).at(-1).hidden === true, "hidden (the top row is the front object)");
  await page.locator("#formatPane .fp-sel-list li").first().locator(".fp-sel-eye").click();
  await page.waitForTimeout(250);
  assert(!(await objects()).at(-1).hidden, "shown again");
  // The right-click menu on the text box (visible above the photo): lock it, then the arrow keys leave it.
  const text = (await objects()).find((o) => o.kind === "text");
  await page.mouse.click(...at(text.x + 30, text.y + 30), { button: "right" });
  await page.waitForSelector(".ed-menu");
  assert((await page.evaluate(() => window.__hpptx.selection())).join() === text.id, "a right-click selects what is under the mouse");
  await shot("menu");
  await page.click('.ed-menu button:has-text("ロック")');
  await page.waitForTimeout(250);
  assert((await objects()).find((o) => o.kind === "text").locked === true, "locked");
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(200);
  assert((await objects()).find((o) => o.kind === "text").x === text.x, "locked objects do not move");
  await drag([text.x + 30, text.y + 30], [text.x + 300, text.y + 200]);
  assert((await objects()).find((o) => o.kind === "text").x === text.x, "nor with the mouse");
  // Unlock from the 配置 section of the 書式 pane.
  await page.locator('#formatPane [data-fp="arrange"] > summary').click();
  await page.click('#formatPane [data-fp="arrange"] .fp-tool[title="ロックを解除"]');
  await page.waitForTimeout(250);
  assert(!(await objects()).find((o) => o.kind === "text").locked, "unlocked");
});

await step("link: the text box goes to slide 5 when clicked in the presentation", async () => {
  await pickFromPane("text");
  await tab("挿入");
  await ribbon("リンク・動作");
  await page.waitForSelector(".rb-pop .rb-form");
  await page.selectOption('.rb-pop select[aria-label="クリックしたときの動作"]', "slide");
  await page.selectOption('.rb-pop select[aria-label="移動先のスライド"]', "4");
  await page.click('.rb-pop .rb-form button:has-text("設定する")');
  await page.waitForTimeout(300);
  const action = (await objects()).find((o) => o.kind === "text").action;
  const sid = await page.evaluate(() => window.__hpptx.deck().slides[4].sid);
  assert(action?.type === "slide" && action.to && action.to === sid, `action ${JSON.stringify(action)} → ${sid}`);
});

await step("the work comes back after a reload", async () => {
  const before = await objects();
  await page.waitForTimeout(1200);
  await page.reload();
  await page.waitForSelector(".slide-wrap .hs-objects .hs-obj");
  const after = await objects();
  assert(JSON.stringify(after) === JSON.stringify(before), "same objects");
  await measure();
});

await step("present: the objects are on the slide", async () => {
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector(".hs-player .hs-slide .hs-objects .hs-obj");
  await page.waitForTimeout(600);
  const n = await page.locator(".hs-player .hs-slide .hs-objects .hs-obj").count();
  assert(n === (await objects()).filter((o) => !o.hidden).length, `objects in the presentation: ${n}`);
  await shot("present");
  // The linked text box jumps to slide 5 instead of advancing.
  const linked = page.locator(".hs-player .hs-slide .hs-obj[data-action='slide']");
  assert((await linked.count()) === 1, "the link is clickable in the presentation");
  await linked.click();
  await page.waitForTimeout(900);
  const expected = await page.evaluate(() => window.__hpptx.deck().slides.slice(0, 5).filter((slide) => !slide.drillOf).length);
  const count = await page.textContent(".hs-player-count");
  assert(count.startsWith(`${expected} /`), `went to slide 5: ${count}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  await page.locator(".film-item").nth(1).click();
  await page.waitForSelector(".slide-wrap .hs-objects .hs-obj");
  await measure();
});

await step("export: one HTML file with the objects and their engine", async () => {
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#downloadBtn").then(async () => { if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn"); })]);
  const file = join(outDir, "objects-export.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  assert(/objectLayer/.test(html), "objects.js inside");
  assert(/"elements":\[/.test(html), "objects in the deck data");
  assert(/data:image\/jpeg;base64/.test(html), "the photo travels inside the file");
  const viewer = await browser.newPage({ viewport: { width: 1280, height: 720 } });
  await viewer.goto(`file://${file}#2`);
  await viewer.waitForSelector(".hs-player .hs-objects .hs-obj");
  await viewer.screenshot({ path: join(outDir, "objects-export.png") });
  await viewer.close();
});

await step("その他の色: a colour of its own is kept on the shape and is not flagged", async () => {
  await pickFromPane("shape");
  const id = await page.evaluate(() => window.__hpptx.selection()[0]);
  await page.evaluate(() => { const input = document.querySelector('#formatPane [data-fp="fill"] input[type="color"]'); input.value = "#ff8800"; input.dispatchEvent(new Event("change", { bubbles: true })); });
  await page.waitForFunction((objectId) => window.__hpptx.slide().elements.find((o) => o.id === objectId)?.fill === "#ff8800", id, { timeout: 5000 });
  await page.click("#issueSummary");
  await page.waitForSelector("#exportCheckDialog[open]");
  const text = await page.textContent("#exportCheckDialog");
  assert(!/#ff8800/.test(text), `no finding for a colour of its own: ${text.slice(0, 200)}`);
  await page.keyboard.press("Escape");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
