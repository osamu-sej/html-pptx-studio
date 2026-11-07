// PowerPoint's スライド ショー in a real browser: スライド ショーの設定 (a range, looping until Esc, the pen colour, a
// kiosk that moves on by itself, no animation), 目的別スライド ショー (choose slides, put them in an order, start
// it), the pen and highlighter while presenting with 「インク注釈を保持しますか？」 (the ink stays on its slide as an
// ink object, one ⌘Z), and 常に字幕を使用.
// Usage: node qa/studio-show.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 } });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const dialogs = [];
page.on("dialog", (dialog) => { dialogs.push(dialog.message()); dialog.accept(); });
const shot = async (name) => { const file = join(outDir, `show-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible("#presenter .hs-player")) { await page.keyboard.press("Escape"); await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
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
const counter = async () => (await page.textContent("#presenter .hs-player-count")).trim();
const shownTitle = async () => (await page.locator("#presenter .hs-player-slide .hs-slide").last().locator(".hs-title").first().textContent().catch(() => "")).trim();
const strip = (t) => String(t || "").replace(/\*\*/g, "").replace(/\s+/g, "").slice(0, 8);
const undo = async () => { await page.keyboard.press("Control+z"); await page.waitForTimeout(300); };

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.click("#sampleDeckBtn");
await page.waitForSelector(".film-item");
await page.waitForSelector("#ribbon:not([hidden]) .rb-tabs");

await step("スライド ショーの設定: a range (2〜4), looping until Esc and a blue pen, saved on the deck in one ⌘Z", async () => {
  await tab("スライド ショー");
  for (const label of ["最初から", "このスライド", "目的別", "スライド ショー", "非表示スライド", "リハーサル", "常に字幕を"]) assert(await page.locator(`.rb-body .rb-btn:has-text("${label}")`).count(), `button ${label}`);
  await ribbonBtn("の設定");
  await page.waitForSelector(".sh-settings[open]");
  await page.check('.sh-settings input[name="slides"][value="range"]');
  await page.fill('.sh-settings input[name="from"]', "2");
  await page.fill('.sh-settings input[name="to"]', "4");
  await page.check('.sh-settings input[name="loop"]');
  await page.selectOption('.sh-settings select[name="penColor"]', "#0070c0");
  await shot("settings");
  await page.click(".sh-settings .sh-ok");
  await page.waitForTimeout(300);
  const show = (await deck()).show;
  assert(show?.range?.from === 2 && show.range.to === 4 && show.loop && show.penColor === "#0070c0", `saved: ${JSON.stringify(show)}`);
  await undo();
  assert(!(await deck()).show, "⌘Z takes the settings back");
  await page.keyboard.press("Control+y");
  await page.waitForTimeout(300);
  assert((await deck()).show?.loop, "⌘Y puts them back");
});

await step("F5 plays only slides 2〜4 and, at the end, starts them again", async () => {
  const d = await deck();
  await page.keyboard.press("F5");
  await page.waitForSelector("#presenter .hs-player");
  await page.waitForTimeout(900);
  assert((await counter()).startsWith("1 / 3"), `three slides, from the first: ${await counter()}`);
  assert(strip(await shownTitle()) === strip(d.slides[1].title), `slide 2 first: ${await shownTitle()}`);
  for (let i = 0; i < 40 && !(await counter()).startsWith("3 / 3"); i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(200); }
  assert((await counter()).startsWith("3 / 3"), `reached the last: ${await counter()}`);
  for (let i = 0; i < 40 && (await counter()).startsWith("3 / 3"); i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(200); }
  assert((await counter()).startsWith("1 / 3"), `and round again: ${await counter()}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

await step("ペン while presenting (Ctrl+P, the blue of the settings) and 蛍光ペン: kept on the slide as ink when asked", async () => {
  const before = (await deck()).slides[1].elements?.length || 0;
  dialogs.length = 0;
  await page.keyboard.press("F5");
  await page.waitForSelector("#presenter .hs-player");
  await page.waitForTimeout(900);
  await page.keyboard.press("Control+p");
  await page.waitForSelector('#presenter .hs-player.pen-on[data-pen="pen"]');
  const r = await page.locator("#presenter .hs-show-ink").boundingBox();
  const at = (x, y) => [r.x + (x * r.width) / 1920, r.y + (y * r.height) / 1080];
  await page.mouse.move(...at(400, 500));
  await page.mouse.down();
  await page.mouse.move(...at(800, 560), { steps: 10 });
  await page.mouse.move(...at(1200, 520), { steps: 10 });
  await page.mouse.up();
  assert((await counter()).startsWith("1 / 3"), "writing does not move on");
  await page.keyboard.press("Control+i");
  await page.mouse.move(...at(400, 700));
  await page.mouse.down();
  await page.mouse.move(...at(1000, 700), { steps: 10 });
  await page.mouse.up();
  const paths = await page.locator("#presenter .hs-show-ink path").count();
  assert(paths === 2, `two strokes on the slide: ${paths}`);
  assert((await page.locator("#presenter .hs-show-ink path").first().getAttribute("stroke")) === "#0070c0", "the blue pen");
  await shot("pen");
  await page.keyboard.press("Escape");
  assert(!(await page.locator("#presenter .hs-player.pen-on").count()), "Esc puts the pen down first");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(600);
  assert(dialogs.some((m) => m.includes("インク")), `asked to keep the ink: ${dialogs}`);
  const slide = (await deck()).slides[1];
  const inks = (slide.elements || []).filter((o) => o.kind === "ink");
  assert((slide.elements?.length || 0) === before + 1 && inks.length === 1 && inks[0].strokes.length === 2, `one ink object with both strokes: ${JSON.stringify((slide.elements || []).map((o) => [o.kind, o.strokes?.length]))}`);
  assert(inks[0].strokes.some((st) => st.highlighter && st.color === "#ffff00"), "the highlighter in yellow");
  assert(Math.abs(inks[0].x - 400) < 30 && Math.abs(inks[0].x + inks[0].w - 1200) < 30, `where it was written: ${inks[0].x} ${inks[0].w}`);
  assert(await page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide())) === 1, "the editor is on that slide");
  assert(await page.locator(`#stageBody .hs-obj[data-el="${inks[0].id}"] .hs-ink path`).count() === 2, "drawn on the stage");
  await undo();
  assert(!((await deck()).slides[1].elements || []).some((o) => o.kind === "ink"), "⌘Z takes the ink away");
});

let showId = null;
await step("目的別スライド ショー: new, two slides in an order of their own, saved; started from the menu", async () => {
  await tab("スライド ショー");
  await ribbonBtn("目的別");
  await menuItem("目的別スライド ショー…");
  await page.waitForSelector(".sh-customs[open]");
  await page.click(".sh-customs .sh-new");
  await page.fill(".sh-customs .sh-name", "役員向け（短縮版）");
  await page.check('.sh-customs input[type=checkbox][data-index="4"]');
  await page.check('.sh-customs input[type=checkbox][data-index="2"]');
  assert((await page.locator(".sh-customs .sh-order li").count()) === 2, "two slides in the order");
  await page.locator(".sh-customs .sh-order li").first().locator('button[title="下へ"]').click();
  await shot("custom-shows");
  await page.click(".sh-customs .sh-ok");
  await page.waitForTimeout(300);
  const d = await deck();
  const cs = d.customShows?.[0];
  showId = cs?.id;
  assert(cs?.name === "役員向け（短縮版）" && cs.sids.length === 2, `saved: ${JSON.stringify(cs)}`);
  assert(cs.sids[0] === d.slides[2].sid && cs.sids[1] === d.slides[4].sid, "in the order chosen (slide 3, then slide 5)");
  await ribbonBtn("目的別");
  await menuItem("役員向け");
  await page.waitForSelector("#presenter .hs-player");
  await page.waitForTimeout(900);
  assert((await counter()).startsWith("1 / 2"), `two slides: ${await counter()}`);
  assert(strip(await shownTitle()) === strip(d.slides[2].title), `slide 3 first: ${await shownTitle()}`);
  await page.keyboard.press("ArrowRight");
  for (let i = 0; i < 20 && (await counter()).startsWith("1 / 2"); i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(200); }
  assert(strip(await shownTitle()) === strip(d.slides[4].title), `then slide 5: ${await shownTitle()}`);
});

await step("自動プレゼンテーション: the custom show as a kiosk — clicks do not move it, it moves on by itself", async () => {
  await tab("スライド ショー");
  await ribbonBtn("の設定");
  await page.waitForSelector(".sh-settings[open]");
  await page.check('.sh-settings input[name="kind"][value="kiosk"]');
  await page.fill('.sh-settings input[name="kioskSeconds"]', "2");
  await page.check('.sh-settings input[name="slides"][value="custom"]');
  await page.check('.sh-settings input[name="noAnimation"]');
  await page.click(".sh-settings .sh-ok");
  await page.waitForTimeout(300);
  const show = (await deck()).show;
  assert(show.kiosk && show.custom === showId && show.kioskSeconds === 2 && show.noAnimation && !show.range, `saved: ${JSON.stringify(show)}`);
  await page.keyboard.press("F5");
  await page.waitForSelector("#presenter .hs-player.kiosk");
  await page.waitForTimeout(500);
  assert((await counter()).startsWith("1 / 2"), `the custom show: ${await counter()}`);
  await page.mouse.click(800, 500);
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(300);
  assert((await counter()).startsWith("1 / 2"), "a click and → do nothing");
  for (let i = 0; i < 24 && !(await counter()).startsWith("2 / 2"); i += 1) await page.waitForTimeout(250);
  assert((await counter()).startsWith("2 / 2"), `it moved on by itself: ${await counter()}`);
  const seen = [];
  for (let i = 0; i < 24 && !(await counter()).startsWith("1 / 2"); i += 1) { await page.waitForTimeout(250); seen.push(await counter()); }
  assert((await counter()).startsWith("1 / 2"), `and loops: ${seen.join(", ")}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  assert(!(await page.isVisible("#presenter .hs-player")), "Esc ends it");
});

await step("常に字幕を使用 and タイミングを使用: toggles on the deck's show settings", async () => {
  await tab("スライド ショー");
  await ribbonBtn("常に字幕を");
  assert((await deck()).show.captions === true, "captions on");
  assert(await page.locator('.rb-body .rb-btn[aria-pressed="true"]:has-text("常に字幕を")').count(), "pressed");
  await ribbonBtn("タイミングを使用");
  assert((await deck()).show.useTimings === false, "timings off");
  await ribbonBtn("タイミングを使用");
  assert((await deck()).show.useTimings === undefined, "timings on again");
  await ribbonBtn("常に字幕を");
});

await step("deleting the custom show the settings use: the settings go back to all slides", async () => {
  await tab("スライド ショー");
  await ribbonBtn("目的別");
  await menuItem("目的別スライド ショー…");
  await page.waitForSelector(".sh-customs[open]");
  await page.click('.sh-customs .sh-actions button:has-text("削除")');
  await page.click(".sh-customs .sh-ok");
  await page.waitForTimeout(300);
  const d = await deck();
  assert(!d.customShows && !d.show?.custom, `gone: ${JSON.stringify([d.customShows, d.show])}`);
  assert(d.show?.kiosk, "the rest of the settings stay");
});

await step("メディア コントロールを表示: off in the settings, the show has no video controls or sound bars; on again", async () => {
  await tab("スライド ショー");
  await ribbonBtn("の設定");
  await page.waitForSelector(".sh-settings[open]");
  assert(await page.isChecked('.sh-settings input[name="mediaControls"]'), "on by default");
  await page.check('.sh-settings input[name="kind"][value="speaker"]');
  await page.uncheck('.sh-settings input[name="loop"]');
  await page.uncheck('.sh-settings input[name="mediaControls"]');
  await page.click(".sh-settings .sh-ok");
  await page.waitForTimeout(300);
  assert((await deck()).show?.noControls === true, `saved: ${JSON.stringify((await deck()).show)}`);
  await ribbonBtn("このスライド");
  await page.waitForSelector("#presenter .hs-player-slide");
  await page.waitForTimeout(400);
  assert(await page.locator("#presenter .hs-player-slide .hs-no-controls, #presenter .hs-player-slide.hs-no-controls").count(), "the slide is told");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  await ribbonBtn("の設定");
  await page.waitForSelector(".sh-settings[open]");
  assert(!(await page.isChecked('.sh-settings input[name="mediaControls"]')), "kept off");
  await page.check('.sh-settings input[name="mediaControls"]');
  await page.click(".sh-settings .sh-ok");
  await page.waitForTimeout(300);
  assert(!(await deck()).show?.noControls, "on again");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
