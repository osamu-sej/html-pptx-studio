// The editing window as PowerPoint's, in a real browser: the ribbon across the top with its display options
// (always / tabs only, where a tab opens the ribbon over the slide / auto-hide; ⌘F1; a double-click on a tab),
// the whole slide in the window (fit) and the zoom in the status bar, the notes under the slide, the thumbnails and
// the task pane resized, closed and reopened (and remembered), the デザイン・スライド ショー・校閲・表示 tabs,
// the ribbon in the スライド一覧 view, and narrow windows.
// Usage: node qa/studio-layout.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { execFileSync } from "node:child_process";
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
const decks = join(outDir, "import");
await mkdir(decks, { recursive: true });
execFileSync(process.env.PYTHON_BIN || "python3", [join(root, "test", "pptx_fixtures.py"), decks], { stdio: "ignore" });

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
// The window of the person who asked for this (a MacBook browser window).
const context = await browser.newContext({ viewport: { width: 1904, height: 954 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name) => { const file = join(outDir, `layout-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
  for (let i = 0; i < 3 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const box = (selector) => page.locator(selector).first().boundingBox();
const tab = (label) => page.click(`.rb-tabs [role=tab]:has-text("${label}")`);
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
const menuItem = (label) => page.click(`.rb-pop .rb-menu button:has-text("${label}")`);
// A menu item by its whole label (「50%」 is also inside 「150%」).
const menuExact = (label) => page.click(`.rb-pop .rb-menu button:has(span:text-is("${label}"))`);
const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
/** The slide on the stage lies wholly inside the stage (nothing to scroll to see it). */
async function slideFits() {
  const slide = await box(".slide-wrap .hs-scaler");
  const stage = await box("#stageBody");
  assert(slide && stage, "slide and stage");
  assert(slide.y >= stage.y - 1 && slide.y + slide.height <= stage.y + stage.height + 1 && slide.x >= stage.x - 1 && slide.x + slide.width <= stage.x + stage.width + 1,
    `the slide (${Math.round(slide.y)}–${Math.round(slide.y + slide.height)}) fits the stage (${Math.round(stage.y)}–${Math.round(stage.y + stage.height)})`);
  return slide;
}
const ribbonMode = () => page.getAttribute("#ribbon", "data-mode");

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.waitForSelector("#importDeckBtn");
await page.setInputFiles("#importDeckFile", join(decks, "template.pptx"));
await page.waitForSelector("#importModeDialog[open]");
await page.click('#importModeDialog [data-mode="exact"]');
await page.waitForSelector("#msgBar .callout", { timeout: 30000 });
await page.waitForTimeout(600);

await step("PowerPoint's window: the ribbon across the top, a one-line message bar, the whole slide, the status bar", async () => {
  const ribbon = await box("#ribbon");
  const vw = page.viewportSize().width;
  assert(ribbon.x === 0 && Math.abs(ribbon.width - vw) < 2, `the ribbon spans the window: ${ribbon.x} ${ribbon.width}`);
  const film = await box("#filmstrip");
  assert(film.y >= ribbon.y + ribbon.height - 1, "the thumbnails are under the ribbon");
  const bar = await box("#msgBar");
  assert(bar.height <= 52, `the message bar is one line: ${bar.height}`);
  assert((await page.textContent("#msgBar")).includes("見た目どおりに取り込"), "the import message is in the bar");
  await slideFits();
  assert(/スライド 1 \/ 4/.test(await page.textContent("#slidePos")), "status bar: スライド 1 / 4");
  assert(await page.isVisible("#issueSummary"), "the checks are in the status bar");
  assert(await page.isVisible('#statusBar input[name=view][value=single] + span'), "the views are in the status bar");
  assert(!(await page.locator(".stage-toolbar").count()), "no separate toolbar row above the ribbon");
  await shot("window");
});

await step("the message bar shows its whole text on 詳しく and closes", async () => {
  await page.click("#msgBar .msg-more");
  assert(await page.evaluate(() => document.getElementById("msgBar").classList.contains("open")), "open");
  await page.click("#msgBar .msg-more");
  await page.click('#msgBar button[title="閉じる"]');
  await page.waitForTimeout(300);
  assert(await page.isHidden("#msgBar"), "closed");
  await slideFits();
});

await step("リボンを折りたたむ (˄): only the tabs; the slide grows into the room", async () => {
  const before = await slideFits();
  const room = await box("#stageBody");
  await page.click(".rb-end .rb-pin");
  await page.waitForTimeout(400);
  assert((await ribbonMode()) === "tabs", "tabs only");
  assert(!(await page.locator("#ribbon .rb-body").count()), "no ribbon body");
  const roomAfter = await box("#stageBody");
  assert(roomAfter.height > room.height + 80, `the stage gets the ribbon's room: ${Math.round(room.height)} → ${Math.round(roomAfter.height)}`);
  const after = await slideFits();
  assert(after.height >= before.height, `the slide does not shrink: ${Math.round(before.height)} → ${Math.round(after.height)}`);
  await shot("tabs-only");
});

await step("tabs only: a tab opens the ribbon over the slide (nothing moves); a click on the slide closes it", async () => {
  const before = await box(".slide-wrap .hs-scaler");
  await tab("挿入");
  await page.waitForSelector("#ribbon.rb-floating .rb-body");
  const body = await box("#ribbon .rb-body");
  const after = await box(".slide-wrap .hs-scaler");
  assert(Math.abs(after.y - before.y) < 1 && Math.abs(after.height - before.height) < 1, "the slide stays where it is");
  assert(body.y + body.height > after.y, "the ribbon floats over the top of the stage");
  await shot("tabs-floating");
  await page.mouse.click(after.x + after.width / 2, after.y + after.height - 30);
  await page.waitForTimeout(300);
  assert(!(await page.locator("#ribbon .rb-body").count()), "closed by the click on the slide");
});

await step("tabs only: a command closes the floating ribbon; Esc closes it; the pin keeps it", async () => {
  await tab("ホーム");
  await page.waitForSelector("#ribbon.rb-floating .rb-body");
  await ribbonBtn("選択ウィンドウ");
  await page.waitForTimeout(300);
  assert(!(await page.locator("#ribbon .rb-body").count()), "closed after the command");
  assert(await page.isVisible("#formatPane"), "the command ran (選択ウィンドウ)");
  await tab("表示");
  await page.waitForSelector("#ribbon.rb-floating .rb-body");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  assert(!(await page.locator("#ribbon .rb-body").count()), "Esc closes it");
  await tab("表示");
  await page.click("#ribbon .rb-end .rb-pin");
  await page.waitForTimeout(300);
  assert((await ribbonMode()) === "full" && !(await page.locator("#ribbon.rb-floating").count()), "pinned: always shown");
});

await step("⌘F1 and a double-click on a tab fold and unfold the ribbon", async () => {
  await page.click("#stageBody", { position: { x: 5, y: 5 } });
  await page.keyboard.press("Control+F1");
  await page.waitForTimeout(250);
  assert((await ribbonMode()) === "tabs", "⌘F1 folds");
  await page.keyboard.press("Control+F1");
  await page.waitForTimeout(250);
  assert((await ribbonMode()) === "full", "⌘F1 unfolds");
  await page.dblclick('.rb-tabs [role=tab]:has-text("ホーム")');
  await page.waitForTimeout(250);
  assert((await ribbonMode()) === "tabs", "double-click folds");
  await page.dblclick('.rb-tabs [role=tab]:has-text("ホーム")');
  await page.waitForTimeout(250);
  assert((await ribbonMode()) === "full", "double-click unfolds");
});

await step("リボンの表示オプション: auto-hide behind a thin bar, remembered after a reload", async () => {
  await page.click(".rb-options");
  await menuItem("自動的に非表示");
  await page.waitForTimeout(300);
  assert((await ribbonMode()) === "auto", "auto-hide");
  assert(await page.isVisible("#ribbon .rb-reveal"), "the bar to bring it back");
  assert(await page.isHidden("#ribbon .rb-tabs"), "even the tabs are hidden");
  const big = await slideFits();
  await page.click("#ribbon .rb-reveal");
  await page.waitForSelector("#ribbon .rb-float .rb-body");
  const still = await box(".slide-wrap .hs-scaler");
  assert(Math.abs(still.y - big.y) < 1, "the ribbon comes over the slide");
  await shot("auto-hide-open");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  assert(!(await page.locator("#ribbon .rb-float").count()), "Esc hides it again");
  await page.reload();
  await page.waitForSelector(".slide-wrap .hs-scaler");
  assert((await ribbonMode()) === "auto", "remembered");
  await page.click("#ribbon .rb-reveal");
  await page.click(".rb-options");
  await menuItem("常にリボンを表示");
  await page.waitForTimeout(300);
  assert((await ribbonMode()) === "full", "back to always");
});

await step("zoom: fit by default; − ＋, the slider, 50%, ⌘＋ホイール, and back to fit", async () => {
  assert((await page.getAttribute("#zoomFitBtn", "aria-pressed")) === "true", "fit by default");
  const fit = Number((await page.textContent("#zoomPctBtn")).replace("%", ""));
  assert(fit > 40 && fit < 120, `fit percent ${fit}`);
  await page.click("#zoomInBtn");
  await page.waitForTimeout(250);
  const zoomed = Number((await page.textContent("#zoomPctBtn")).replace("%", ""));
  assert(zoomed > fit && (await page.getAttribute("#zoomFitBtn", "aria-pressed")) === "false", `zoom in: ${fit} → ${zoomed}`);
  await page.click("#zoomPctBtn");
  await menuExact("50%");
  await page.waitForTimeout(250);
  const half = await box(".slide-wrap .hs-scaler");
  assert(Math.abs(half.width - 640) < 2, `50% is 640px (13.33in at 96dpi × 0.5): ${half.width}`);
  await page.fill("#zoomSlider", "150");
  await page.dispatchEvent("#zoomSlider", "input");
  await page.waitForTimeout(250);
  const big = await box(".slide-wrap .hs-scaler");
  assert(Math.abs(big.width - 1920) < 2, `150%: ${big.width}`);
  assert(await page.evaluate(() => { const b = document.getElementById("stageBody"); return b.scrollWidth > b.clientWidth; }), "zoomed past the window, the stage scrolls");
  await page.click("#zoomFitBtn");
  await page.waitForTimeout(250);
  await slideFits();
  const stage = await box("#stageBody");
  await page.mouse.move(stage.x + stage.width / 2, stage.y + stage.height / 2);
  await page.keyboard.down("Control");
  await page.mouse.wheel(0, -120);
  await page.keyboard.up("Control");
  await page.waitForTimeout(250);
  assert(Number((await page.textContent("#zoomPctBtn")).replace("%", "")) > fit, "⌘＋ホイール zooms in");
  await page.click("#zoomFitBtn");
  await page.waitForTimeout(200);
});

await step("ノート: the notes under the slide, typed into the slide, resized; the slide still fits", async () => {
  await page.click(".film-item:nth-child(2)");
  await page.waitForTimeout(400);
  await page.click("#notesToggle");
  await page.waitForSelector("#notesPane:not([hidden])");
  assert((await page.innerText("#notesInput")).includes("全体像"), "the slide's own notes");
  await page.fill("#notesInput", "ノート欄から書いた話す内容");
  await page.waitForTimeout(300);
  assert((await deck()).slides[1].notes === "ノート欄から書いた話す内容", "typed into the slide");
  await slideFits();
  const before = await box("#notesPane");
  const split = await box(".notes-split");
  await page.mouse.move(split.x + split.width / 2, split.y + 2);
  await page.mouse.down();
  await page.mouse.move(split.x + split.width / 2, split.y - 80, { steps: 4 });
  await page.mouse.up();
  await page.waitForTimeout(250);
  const after = await box("#notesPane");
  assert(after.height > before.height + 60, `taller: ${before.height} → ${after.height}`);
  await slideFits();
  await shot("notes");
});

await step("the task pane and the thumbnails: resized, closed, reopened, remembered", async () => {
  const side = await box(".side-panel");
  const split = await box('.pane-split[data-split="side"]');
  await page.mouse.move(split.x + 2, split.y + 200);
  await page.mouse.down();
  await page.mouse.move(split.x - 120, split.y + 200, { steps: 5 });
  await page.mouse.up();
  await page.waitForTimeout(250);
  const wider = await box(".side-panel");
  assert(wider.width > side.width + 100, `wider task pane: ${side.width} → ${wider.width}`);
  await page.click("#sideClose");
  await page.waitForTimeout(300);
  const closed = await box(".side-panel");
  assert(closed.width < 50, `closed task pane: ${closed.width}`);
  await slideFits();
  await shot("side-closed");
  await page.click("#animTab");
  await page.waitForTimeout(300);
  assert((await box(".side-panel")).width > 300 && await page.isVisible("#animPane"), "a tab reopens it on that pane");
  const film = await box('.pane-split[data-split="film"]');
  await page.mouse.move(film.x + 2, film.y + 300);
  await page.mouse.down();
  await page.mouse.move(40, film.y + 300, { steps: 5 });
  await page.mouse.up();
  await page.waitForTimeout(250);
  assert(await page.isHidden("#filmstrip") && await page.isVisible("#filmReopen"), "dragged shut, the thumbnails close");
  await page.reload();
  await page.waitForSelector(".slide-wrap .hs-scaler");
  assert(await page.isHidden("#filmstrip"), "remembered after a reload");
  assert(Math.abs((await box(".side-panel")).width - wider.width) < 2, "the task pane width is remembered");
  await page.click("#filmReopen");
  await page.waitForTimeout(250);
  assert(await page.isVisible("#filmstrip"), "reopened");
});

await step("デザイン: the deck's transition from the ribbon; スライド ショー: hide a slide; 校閲: find and replace", async () => {
  await tab("デザイン");
  await ribbonBtn("切り替え");
  await menuItem("ワイプ");
  await page.waitForTimeout(300);
  assert((await deck()).transition === "wipe", "deck transition");
  await shot("design-tab");
  await page.click(".film-item:nth-child(4)");
  await page.waitForTimeout(300);
  await tab("スライド ショー");
  await ribbonBtn("非表示スライド");
  await page.waitForTimeout(300);
  assert((await deck()).slides[3].hidden === true, "hidden");
  assert(await page.locator('.film-item:nth-child(4) .film-flags span:has-text("非")').count(), "the thumbnail says 非");
  assert((await page.textContent("#slidePos")).includes("非表示スライド"), "the status bar says so");
  await ribbonBtn("非表示スライド");
  await page.waitForTimeout(300);
  assert(!(await deck()).slides[3].hidden, "back in the show");
  await shot("slideshow-tab");
  await tab("校閲");
  await ribbonBtn("検索");
  await page.waitForSelector("#replaceDialog[open]");
  await page.keyboard.press("Escape");
});

await step("表示: notes, thumbnails, task pane and zoom from the ribbon", async () => {
  await tab("表示");
  // (The 表示 tab has two ノート buttons, as in PowerPoint: the notes pane and the ノート表示 page.)
  await page.locator('.rb-body .rb-btn[title="スライドの下にスピーカーノートを表示"]').click();
  await page.waitForTimeout(250);
  assert(await page.isHidden("#notesPane"), "notes off");
  await ribbonBtn("ズーム");
  await menuExact("100%");
  await page.waitForTimeout(250);
  assert(Math.abs((await box(".slide-wrap .hs-scaler")).width - 1280) < 2, "100% is 1280px");
  await ribbonBtn("ウィンドウに");
  await page.waitForTimeout(250);
  await slideFits();
  await ribbonBtn("作業ウィンドウ");
  await page.waitForTimeout(250);
  assert((await box(".side-panel")).width < 50, "task pane closed from the ribbon");
  await ribbonBtn("作業ウィンドウ");
  await page.waitForTimeout(250);
  assert((await box(".side-panel")).width > 300, "and back");
});

await step("スライド一覧: the ribbon stays; slide commands work, object commands wait for the 1枚 view", async () => {
  await page.click('#statusBar label:has-text("一覧")');
  await page.waitForTimeout(400);
  assert(await page.isVisible("#ribbon .rb-tabs"), "ribbon in the slide sorter");
  await tab("ホーム");
  const slideGroup = page.locator('.rb-group:has(.rb-label:text-is("スライド"))');
  assert(!(await slideGroup.evaluate((g) => g.classList.contains("rb-off"))), "スライド group usable");
  const fontGroup = page.locator('.rb-group:has(.rb-label:text-is("フォント"))');
  assert(await fontGroup.evaluate((g) => g.classList.contains("rb-off")), "フォント group waits");
  await shot("sorter");
  await page.click('#statusBar label:has-text("1枚")');
  await page.waitForTimeout(400);
});

await step("thumbnails: the right-click menu, and Delete / ⌘D / ⌘Z on the focused thumbnails", async () => {
  const n = (await deck()).slides.length;
  await page.click(".film-item:nth-child(2)", { button: "right" });
  await page.waitForSelector(".ed-menu");
  for (const label of ["コピー", "新しいスライド", "スライドの複製", "スライドの削除", "非表示スライドに設定", "このスライドから発表"]) assert(await page.locator(`.ed-menu button:has-text("${label}")`).count(), `menu: ${label}`);
  await shot("film-menu");
  await page.click('.ed-menu button:has-text("コピー")');
  await page.click(".film-item:nth-child(2)", { button: "right" });
  await page.click('.ed-menu button:has-text("貼り付け")');
  await page.waitForTimeout(300);
  const pasted = await deck();
  assert(pasted.slides.length === n + 1 && pasted.slides[2].title === pasted.slides[1].title, "pasted after it");
  await page.click(".film-item:nth-child(3)");
  await page.keyboard.press("Delete");
  await page.waitForTimeout(300);
  assert((await deck()).slides.length === n, "Delete on the thumbnails deletes the slide");
  await page.click(".film-item:nth-child(2)");
  await page.keyboard.press("Control+d");
  await page.waitForTimeout(300);
  assert((await deck()).slides.length === n + 1, "⌘D duplicates");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  assert((await deck()).slides.length === n, "⌘Z undoes it");
  // A click on the slide takes the keys back: Delete then removes the selected object, not the slide.
  await page.click(".film-item:nth-child(2)");
  const objectsBefore = (await deck()).slides[1].elements.length;
  const target = page.locator('.slide-wrap .hs-obj[data-kind="shape"]').first();
  const b = await target.boundingBox();
  await page.mouse.click(b.x + b.width / 2, b.y + b.height / 2);
  await page.waitForTimeout(200);
  await page.keyboard.press("Delete");
  await page.waitForTimeout(300);
  const after = await deck();
  assert(after.slides.length === n && after.slides[1].elements.length === objectsBefore - 1, `Delete on the slide removes the object (${objectsBefore} → ${after.slides[1].elements.length}), slides ${after.slides.length}`);
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
});

await step("表示 → ルーラー: centimetres around the slide, the selection shaded on them; the slide keeps clear of them", async () => {
  await page.click(".film-item:nth-child(2)");
  await page.waitForTimeout(300);
  await tab("表示");
  await ribbonBtn("ルーラー");
  await page.waitForTimeout(400);
  assert(await page.isVisible(".ruler-h") && await page.isVisible(".ruler-v"), "both rulers");
  const ruler = await box(".ruler-h");
  const slide = await slideFits();
  assert(slide.y >= ruler.y + ruler.height - 1, "the slide is below the ruler");
  const target = page.locator('.slide-wrap .hs-obj[data-kind="shape"]').first();
  const b = await target.boundingBox();
  await page.mouse.click(b.x + b.width / 2, b.y + b.height / 2);
  await page.waitForTimeout(300);
  const shade = await page.evaluate(({ x }) => {
    const canvas = document.querySelector(".ruler-h");
    const r = canvas.getBoundingClientRect();
    const dpr = canvas.width / r.width;
    return [...canvas.getContext("2d").getImageData(Math.round((x - r.left) * dpr), Math.round(2 * dpr), 1, 1).data];
  }, { x: b.x + b.width / 2 });
  assert(shade[0] < 240 && shade[2] > shade[0], `the selection is shaded on the ruler: ${shade}`);
  await shot("rulers");
  await tab("表示");
  await ribbonBtn("ルーラー");
  await page.waitForTimeout(300);
  assert(await page.isHidden(".ruler-h"), "rulers off");
});

await step("挿入 → 記号と特殊文字 and 日付と時刻; ホーム → 文字の間隔; 表示 → グレースケール", async () => {
  await page.click(".film-item:nth-child(4)");
  await page.waitForTimeout(300);
  const before = (await deck()).slides[3].elements.length;
  await tab("挿入");
  await ribbonBtn("記号と特殊文字");
  await page.click('.rb-pop .rb-symbols button[title="※"]');
  await page.waitForTimeout(300);
  const added = (await deck()).slides[3].elements;
  assert(added.length === before + 1 && /※/.test(added.at(-1).text), "a text box with the symbol");
  const id = added.at(-1).id;
  // While typing in it, the date goes in at the caret.
  await page.keyboard.press("Enter");
  await page.waitForTimeout(200);
  await page.keyboard.press("End");
  await tab("挿入");
  await ribbonBtn("日付と時刻");
  await page.locator(".rb-pop .rb-menu button").nth(0).click();
  await page.waitForTimeout(200);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const typed = (await deck()).slides[3].elements.find((o) => o.id === id).text;
  assert(new RegExp(`※.*${new Date().getFullYear()}年`).test(typed), `the date after the symbol: ${typed}`);
  await page.evaluate((oid) => window.__hpptx, id);
  await tab("ホーム");
  await ribbonBtn("文字の間隔").catch(async () => { await page.click('.rb-body .rb-btn[title="文字の間隔"]'); });
  await page.click('.rb-pop .rb-menu button:has(span:text-is("広く"))');
  await page.waitForTimeout(300);
  assert((await deck()).slides[3].elements.find((o) => o.id === id).ls === 0.05, "character spacing: 広く");
  await tab("表示");
  await ribbonBtn("グレース");
  await page.waitForTimeout(200);
  assert(await page.evaluate(() => document.getElementById("stageBody").classList.contains("view-gray")), "grayscale view");
  await shot("grayscale");
  await ribbonBtn("グレース");
  await page.waitForTimeout(200);
  assert(!(await page.evaluate(() => document.getElementById("stageBody").classList.contains("view-gray"))), "back to colour");
});

await step("a narrow window (1280×720): the slide still fits; the ribbon folds its groups", async () => {
  await page.setViewportSize({ width: 1280, height: 720 });
  await page.waitForTimeout(500);
  await tab("ホーム");
  await slideFits();
  assert(await page.locator(".rb-body .rb-folded, .rb-body .rb-group.compact").count(), "groups fold or compact");
  await shot("narrow");
  await page.setViewportSize({ width: 1904, height: 954 });
  await page.waitForTimeout(300);
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
