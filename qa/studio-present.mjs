// 貼り付けのオプション・フォントの置換・ノート表示・画面切り替えのサウンド・オンライン プレゼンテーション in real
// browsers: a copied shape pasted with the destination's style (its colour moved to the deck's palette) and a text box
// pasted as words only (⇧⌘V); one font set in another across the deck; the notes page view typing notes and moving
// between slides; a sound chosen for a slide's way in, heard when the show arrives there and given to every slide;
// and a second browser watching an online presentation: it follows the presenter slide by slide, cannot move the
// show itself, gets the updated deck and hears that it has ended.
// Usage: node qa/studio-present.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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

const browserArgs = ["--autoplay-policy=no-user-gesture-required", ...(process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [])];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const errors = [];
async function open(name, { quiet = null } = {}) {
  const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, permissions: ["clipboard-read", "clipboard-write"] });
  const p = await context.newPage();
  p.on("pageerror", (error) => errors.push(`${name} pageerror: ${error.message}`));
  p.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g|ERR_ABORTED|EventSource/.test(message.text()) && !(quiet && quiet.test(message.text()))) errors.push(`${name} console: ${message.text()}`); });
  p.on("dialog", (dialog) => dialog.accept());
  return p;
}
const page = await open("presenter");
const shot = async (name, p = page) => { const file = join(outDir, `present-${name}.png`); await p.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
const objects = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide()?.elements ?? [])));
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
const menuItem = (label) => page.locator(`.rb-pop button:has-text("${label}")`).first().click();
const shortcut = (key) => page.keyboard.press(`Control+${key}`);
let box = null;
const at = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
const measure = async () => { await page.waitForSelector(".slide-wrap .hs-slide"); box = await page.locator(".slide-wrap .hs-slide").boundingBox(); };
const drag = async ([x1, y1], [x2, y2]) => { await measure(); await page.mouse.move(...at(x1, y1)); await page.mouse.down(); await page.mouse.move(...at(x2, y2), { steps: 8 }); await page.mouse.up(); await page.waitForTimeout(250); };
const clickAt = async (x, y) => { await measure(); await page.mouse.click(...at(x, y)); await page.waitForTimeout(200); };

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.click("#sampleDeckBtn");
await page.waitForSelector(".film-item");
await page.waitForSelector("#ribbon:not([hidden]) .rb-tabs");
await tab("挿入");
await ribbonBtn("新しいスライド");
await menuItem("白紙");
await page.waitForSelector(".slide-wrap .hs-slide");
await page.waitForTimeout(400);
const palette = await page.evaluate(() => window.SlideEngine.PALETTE.fill.map(([c]) => c.toLowerCase()));

await step("貼り付け先のスタイルを使用: a copied orange shape comes back in the nearest colour of the deck's palette", async () => {
  await tab("挿入");
  await page.locator('.rb-btn.big:has-text("図形")').first().click();
  await page.waitForSelector(".rb-gallery.shapes");
  await page.click('.rb-gallery button[title="正方形/長方形"]');
  await drag([200, 300], [600, 560]);
  const [shape] = (await objects()).filter((o) => o.kind === "shape");
  assert(shape, "a rectangle");
  await tab("図形の書式");
  await page.locator('.rb-body .rb-btn[title="図形の塗りつぶし"]:visible').first().click();
  await page.locator(".rb-pop .rb-custom input[type=color]").fill("#ff6600");
  await page.waitForTimeout(300);
  assert((await objects()).find((o) => o.id === shape.id).fill === "#ff6600", "an orange fill (not one of the palette)");
  await shortcut("c");
  await tab("ホーム");
  await ribbonBtn("貼り付け");
  await menuItem("貼り付け先のスタイルを使用");
  await page.waitForTimeout(300);
  const list = (await objects()).filter((o) => o.kind === "shape");
  assert(list.length === 2, `pasted: ${list.length}`);
  const pasted = list.at(-1);
  assert(pasted.fill !== "#ff6600" && palette.includes(pasted.fill.toLowerCase()), `a palette fill: ${pasted.fill}`);
  await shot("paste-theme");
});

await step("テキストのみ保持 (⇧⌘V): a text box's words come back plain, without its formatting", async () => {
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書きテキスト ボックス");
  await clickAt(900, 300);
  await page.keyboard.type("貼り付けの確認");
  await page.keyboard.press("Control+a");
  await page.keyboard.press("Control+b");
  await clickAt(1700, 950);
  const source = (await objects()).find((o) => o.kind === "text");
  const bold = /<b>|<strong>|font-weight/;
  assert(source && bold.test(source.text), `a bold text box: ${source?.text}`);
  await clickAt(source.x + 20, source.y + 20);
  await shortcut("c");
  const before = (await objects()).length;
  await page.keyboard.press("Control+Shift+v");
  await page.waitForTimeout(400);
  let list = await objects();
  if (list.length === before) {
    // No paste event reached the page (the system clipboard is out of reach): the ribbon's option does the same.
    await tab("ホーム");
    await ribbonBtn("貼り付け");
    await menuItem("テキストのみ保持");
    await page.waitForTimeout(300);
    list = await objects();
  }
  const plain = list.at(-1);
  assert(list.length === before + 1 && plain.kind === "text", `a new text box: ${list.length - before}`);
  assert(/貼り付けの確認/.test(plain.text) && !bold.test(plain.text), `plain words: ${plain.text}`);
});

await step("フォントの置換: every text box in the theme's body font set in BIZ UDPゴシック, one ⌘Z back", async () => {
  await tab("校閲");
  await ribbonBtn("フォントの");
  await page.waitForSelector(".font-dialog[open]");
  const options = await page.locator('.font-dialog select[aria-label="置換前のフォント"] option').allTextContents();
  assert(options.some((t) => /本文のフォント/.test(t)), `fonts in use: ${options}`);
  await page.selectOption('.font-dialog select[aria-label="置換後のフォント"]', "ud");
  await page.click(".font-dialog .fr-ok");
  await page.waitForTimeout(300);
  const texts = (await deck()).slides.flatMap((s) => s.elements || []).filter((o) => ["text", "shape"].includes(o.kind) && o.text);
  assert(texts.length >= 2 && texts.every((o) => o.font === "ud"), `all set in ud: ${texts.map((o) => o.font)}`);
  await shortcut("z");
  await page.waitForTimeout(300);
  const back = (await deck()).slides.flatMap((s) => s.elements || []).filter((o) => ["text", "shape"].includes(o.kind) && o.text);
  assert(back.every((o) => o.font !== "ud"), "one ⌘Z takes it back");
});

await step("ノート表示: the slide above its notes; typing writes the notes; ▶ goes to the next slide", async () => {
  await page.locator(".film-item").nth(1).click();
  await tab("表示");
  await page.locator('.rb-body .rb-btn[title^="スライドとノートを1ページにした表示"]').click();
  await page.waitForSelector(".notes-view[open] .nv-notes");
  assert((await page.textContent(".notes-view .nv-count")).startsWith("2 /"), "opens on the slide shown");
  assert(await page.isVisible(".notes-view .nv-page .hs-slide"), "the slide is on the page");
  await page.fill(".notes-view .nv-notes .notes-input", "ノート表示で書いたメモ");
  await page.locator('.notes-view button[aria-label="次のスライド"]').click();
  await page.waitForTimeout(250);
  assert((await page.textContent(".notes-view .nv-count")).startsWith("3 /"), "the next slide");
  await shot("notes-view");
  await page.keyboard.press("Escape");
  assert((await deck()).slides[1].notes === "ノート表示で書いたメモ", "the notes are kept");
});

await step("ノートの書式: 太字・箇条書きを付ける→ノート表示・発表者のノート・印刷に出る; 文字だけの編集では書式が外れる", async () => {
  await page.locator(".film-item").nth(1).click();
  await page.waitForTimeout(300);
  if ((await page.getAttribute("#notesToggle", "aria-pressed")) !== "true") await page.click("#notesToggle");
  await page.waitForSelector("#notesPane:not([hidden]) .notes-tools");
  assert(await page.locator("#notesPane .notes-tool").count() === 8, "a toolbar with eight tools");
  // Typed words with a bullet list under them; the first line made bold.
  await page.fill("#notesInput", "要点を落ち着いて話す");
  await page.click("#notesInput");
  await page.keyboard.press("End");
  await page.keyboard.press("Enter");
  await page.click("#notesPane .tool-bullet");
  await page.keyboard.type("一つ目");
  await page.keyboard.press("Enter");
  await page.keyboard.type("二つ目");
  await page.waitForTimeout(300);
  let slide = (await deck()).slides[1];
  assert((slide.notesRich || "").includes("<ul><li>一つ目</li><li>二つ目</li></ul>"), `a bullet list: ${slide.notesRich}`);
  assert(/^要点を落ち着いて話す\n\n?・一つ目\n・二つ目/.test(slide.notes), `the plain words with the bullets: ${JSON.stringify(slide.notes)}`);
  await page.evaluate(() => { const box = document.getElementById("notesInput"); const first = box.firstChild; const r = document.createRange(); r.selectNodeContents(first.nodeType === 3 ? first : first.firstChild || first); const sel = getSelection(); sel.removeAllRanges(); sel.addRange(r); box.focus(); });
  await page.click("#notesPane .tool-bold");
  await page.waitForTimeout(300);
  slide = (await deck()).slides[1];
  assert((slide.notesRich || "").includes("<b>要点を落ち着いて話す</b>") && (slide.notesRich || "").includes("<li>一つ目</li>"), `bold on the first line only: ${JSON.stringify({ rich: slide.notesRich, html: await page.evaluate(() => document.getElementById("notesInput").innerHTML) })}`);
  assert(/^要点を落ち着いて話す\n\n?・一つ目\n・二つ目/.test(slide.notes), `the plain words do not change with the bold: ${JSON.stringify(slide.notes)}`);
  await shot("notes-rich");
  // The formatting is still there on coming back to the slide.
  await page.locator(".film-item").nth(2).click();
  await page.waitForTimeout(300);
  await page.locator(".film-item").nth(1).click();
  await page.waitForTimeout(300);
  assert(await page.locator("#notesInput b, #notesInput ul li").count() >= 3, "the bold and the list come back");
  // ノート表示 shows the same.
  await tab("表示");
  await page.locator('.rb-body .rb-btn[title^="スライドとノートを1ページにした表示"]').click();
  await page.waitForSelector(".notes-view[open] .nv-notes .notes-input");
  assert(await page.locator(".notes-view .notes-input b").count() === 1 && await page.locator(".notes-view .notes-input li").count() === 2, "the notes page shows the formatting");
  await page.keyboard.press("Escape");
  // The presenter's notes bar (N) shows it too.
  await tab("スライド ショー");
  await ribbonBtn("このスライド");
  await page.waitForSelector("#presenter .hs-player");
  await page.keyboard.press("n");
  await page.waitForTimeout(300);
  assert(await page.locator("#presenter .hs-player-notes b").count() === 1 && await page.locator("#presenter .hs-player-notes li").count() === 2, "the notes bar shows the bold words and the list");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  // The AI's (or any plain) rewrite of the words drops the old formatting from view.
  await page.evaluate(() => { const s = window.__hpptx.slide(); s.notes = "AI が書き直した"; });
  await page.locator(".film-item").nth(2).click();
  await page.locator(".film-item").nth(1).click();
  await page.waitForTimeout(300);
  assert(!(await page.locator("#notesInput b").count()) && (await page.innerText("#notesInput")).includes("AI が書き直した"), "the plain words win over stale formatting");
});

await step("画面切り替え → サウンド: chosen (and heard), rung as the show arrives, applied to all (one ⌘Z)", async () => {
  // Count the oscillators the synthesiser makes (every sound uses some).
  await page.evaluate(() => { window.__osc = 0; const make = AudioContext.prototype.createOscillator; AudioContext.prototype.createOscillator = function () { window.__osc += 1; return make.call(this); }; });
  await page.locator(".film-item").nth(2).click();
  await tab("画面切り替え");
  await ribbonBtn("サウンド");
  await menuItem("チャイム");
  await page.waitForTimeout(300);
  assert((await deck()).slides[2].transitionSound === "chime", "the slide has the chime");
  assert(await page.evaluate(() => window.__osc >= 3), "heard when chosen");
  assert((await page.textContent(".an-sound-name")) === "チャイム", "the ribbon names it");
  // The show: from the slide before, the next click arrives at the slide with the sound.
  await page.evaluate(() => { window.__osc = 0; });
  await page.locator(".film-item").nth(1).click();
  await tab("スライド ショー");
  await ribbonBtn("このスライド");
  await page.waitForSelector("#presenter .hs-player");
  await page.waitForTimeout(500);
  assert(await page.evaluate(() => window.__osc === 0), "no sound on a slide without one");
  for (let k = 0; k < 12 && (await page.textContent("#presenter .hs-player-count")).startsWith("2 /"); k += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(250); }
  await page.waitForTimeout(400);
  assert(await page.evaluate(() => window.__osc >= 3), "rung as the show arrives at the slide");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  await page.locator(".film-item").nth(2).click();
  await tab("画面切り替え");
  await ribbonBtn("すべてに適用");
  await page.waitForTimeout(300);
  assert((await deck()).slides.every((s) => s.transitionSound === "chime"), "every slide has it");
  await shortcut("z");
  await page.waitForTimeout(300);
  assert((await deck()).slides.filter((s) => s.transitionSound).length === 1, "one ⌘Z takes it back");
});

const viewer = await open("viewer");
let link = "";
const viewerCount = () => viewer.textContent("#presenter .hs-player-count");
const presenterCount = () => page.textContent("#presenter .hs-player-count");

await step("オンライン プレゼンテーション: 接続 makes a link; the viewer opens it and waits for the presenter", async () => {
  await page.locator(".film-item").nth(0).click();
  await tab("スライド ショー");
  await ribbonBtn("オンライン");
  await page.waitForSelector(".online-dialog[open]");
  await page.click(".online-dialog .op-start");
  // (接続 measures every slide's text first; a long deck on a busy machine takes a while.)
  await page.waitForSelector(".online-dialog .op-link", { timeout: 90000 });
  link = await page.inputValue(".online-dialog .op-link");
  assert(/[?&]watch=[A-Za-z0-9_-]{20,}/.test(link), `a link: ${link}`);
  await viewer.goto(link);
  await viewer.waitForSelector("#presenter .hs-player .hs-player-live", { timeout: 15000 });
  await viewer.waitForFunction(() => /待っています/.test(document.querySelector(".op-banner")?.textContent || ""), null, { timeout: 8000 });
  await page.waitForFunction(() => /視聴者 1人/.test(document.querySelector(".online-dialog .op-status")?.textContent || ""), null, { timeout: 8000 });
  assert(!(await viewer.isVisible(".hs-player-btn.end")), "the viewer has no 終了 button");
  await shot("online-dialog");
  await shot("online-waiting", viewer);
});

const onlineDialog = async () => { if (!(await page.isVisible(".online-dialog[open]"))) { await tab("スライド ショー"); await ribbonBtn("オンライン"); await page.waitForSelector(".online-dialog[open]"); } };

await step("the viewer follows the presenter, click by click, and cannot move the show itself", async () => {
  await onlineDialog();
  await page.click(".online-dialog .op-start");
  await page.waitForSelector("#presenter .hs-player");
  await viewer.waitForFunction(() => !(document.querySelector(".op-banner")?.textContent), null, { timeout: 8000 });
  for (let k = 0; k < 6; k += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(350); }
  const here = await presenterCount();
  await viewer.waitForFunction((want) => document.querySelector("#presenter .hs-player-count")?.textContent === want, here, { timeout: 8000 });
  const state = await viewer.evaluate(() => ({ index: window.__hpptxWatch.player.index, step: window.__hpptxWatch.player.step, live: window.__hpptxWatch.state.live }));
  assert(state.live, "live");
  await viewer.keyboard.press("ArrowRight");
  await viewer.mouse.click(800, 450);
  await viewer.waitForTimeout(500);
  assert((await viewerCount()) === here, "keys and clicks do not move the viewer's show");
  await page.keyboard.press("ArrowLeft");
  await page.waitForTimeout(300);
  await page.keyboard.press("ArrowLeft");
  const back = await presenterCount();
  await viewer.waitForFunction((want) => document.querySelector("#presenter .hs-player-count")?.textContent === want, back, { timeout: 8000 });
  await shot("online-following", viewer);
});

// A wait on the viewer that says what the viewer was showing when it gave up (a CI run leaves no other trace).
const viewerState = () => viewer.evaluate(() => ({ banner: document.querySelector(".op-banner")?.textContent || "", title: window.__hpptxWatch?.deck?.title, live: window.__hpptxWatch?.state?.live })).catch((error) => ({ error: error.message }));
const viewerUntil = async (what, fn, arg) => {
  try { await viewer.waitForFunction(fn, arg, { timeout: 20000 }); }
  catch (error) { throw new Error(`${what}: ${error.message.split("\n")[0]} - the viewer showed ${JSON.stringify(await viewerState())}`); }
};

await step("the presenter ends the slide show: the viewer waits; 資料の更新 brings the changed deck", async () => {
  await page.keyboard.press("Escape");
  await page.waitForTimeout(500);
  await viewerUntil("the viewer waits after the show ends", () => /待っています/.test(document.querySelector(".op-banner")?.textContent || ""), null);
  await page.fill("#deckTitleInput", "オンライン発表のテスト");
  await page.press("#deckTitleInput", "Enter");
  await onlineDialog();
  await page.waitForSelector(".online-dialog[open] .op-update:not([hidden])");
  assert(/資料の更新/.test(await page.textContent(".online-dialog .op-body")), "says the deck changed");
  await page.click(".online-dialog .op-update");
  await viewerUntil("the viewer gets the changed deck", () => window.__hpptxWatch.deck?.title === "オンライン発表のテスト", null);
  assert(await viewer.isVisible("#presenter .hs-player"), "the viewer still shows the show");
});

await step("while online, an ordinary slide show (このスライドから) goes to the viewers too", async () => {
  if (await page.isVisible(".online-dialog[open]")) await page.keyboard.press("Escape");
  await page.locator(".film-item").nth(2).click();
  await tab("スライド ショー");
  await ribbonBtn("このスライド");
  await page.waitForSelector("#presenter .hs-player");
  await viewer.waitForFunction(() => window.__hpptxWatch.state.live && !(document.querySelector(".op-banner")?.textContent), null, { timeout: 20000 });
  const here = await presenterCount();
  await viewer.waitForFunction((want) => document.querySelector("#presenter .hs-player-count")?.textContent === want, here, { timeout: 20000 });
  await page.keyboard.press("Escape");
  await viewer.waitForFunction(() => /待っています/.test(document.querySelector(".op-banner")?.textContent || ""), null, { timeout: 20000 });
});

await step("オンライン プレゼンテーションの終了: the viewer hears it is over", async () => {
  await onlineDialog();
  await page.click(".online-dialog .op-end");
  await viewer.waitForFunction(() => /終了しました/.test(document.querySelector(".op-banner")?.textContent || ""), null, { timeout: 20000 });
  await page.waitForTimeout(300);
  assert(!(await page.evaluate(() => document.querySelector('.rb-btn[aria-pressed="true"]')?.textContent?.includes("オンライン"))), "the ribbon button is no longer pressed");
  await shot("online-over", viewer);
});

await step("a link to a show that does not exist says so", async () => {
  // (Its stream is refused with 404, which the browser logs.)
  const lost = await open("lost", { quiet: /404/ });
  await lost.goto(`${base}/?watch=${"x".repeat(24)}`);
  await lost.waitForSelector(".op-message", { timeout: 15000 });
  assert(/見つかりません/.test(await lost.textContent(".op-message")), "not found");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
