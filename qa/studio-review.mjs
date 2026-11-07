// PowerPoint's アクセシビリティ チェック, 印刷 and スライド ショーの記録 in a real browser: the checker in the task pane
// (alternative text added or marked decorative, the reading order set to what the eye sees, ↑↓ in the 読み取り順序
// list), 代替テキストを編集 from the right-click menu, the print dialog (layouts, a range, the preview, the pages
// handed to the browser, コメントを印刷する), and recording narration with timings (kept per slide, played unless 「ナレーションを
// 付けない」, cleared again).
// Usage: node qa/studio-review.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: [...browserArgs, "--use-fake-ui-for-media-stream", "--use-fake-device-for-media-stream", "--autoplay-policy=no-user-gesture-required"] });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, permissions: ["microphone"] });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const dialogs = [];
page.on("dialog", (dialog) => { dialogs.push(dialog.message()); dialog.accept(); });
const shot = async (name) => { const file = join(outDir, `review-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible("#presenter .hs-player")) { await page.keyboard.press("Escape"); await page.waitForTimeout(400); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
const slide = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide())));
const index = () => page.evaluate(() => window.__hpptx.deck().slides.indexOf(window.__hpptx.slide()));
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
let box = null;
const measure = async () => { await page.locator("#stageBody .slide-wrap .hs-slide").first().waitFor({ state: "visible", timeout: 5000 }); box = await page.locator("#stageBody .slide-wrap .hs-slide").first().boundingBox(); };
const at = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
const undo = async () => { await page.keyboard.press("Control+z"); await page.waitForTimeout(300); };
const issues = () => page.$$eval("#a11yPane .a11y-item", (els) => els.map((el) => ({ rule: el.dataset.rule, slide: Number(el.dataset.slide) })));

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
const blank = await index();
// Three text boxes written bottom first, and a picture without a description.
await page.evaluate(() => {
  const s = window.__hpptx.slide();
  s.title = "読み取り順序のテスト";
});
for (const [y, words] of [[760, "最後に読む"], [480, "次に読む"], [220, "最初に読む"]]) {
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await menuItem("横書き");
  await measure();
  await page.mouse.click(...at(300, y));
  await page.keyboard.type(words);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(200);
  await page.mouse.click(...at(1850, 1040));
}

await step("アクセシビリティ チェック: the pane opens with what to fix, by level, starting with errors", async () => {
  await tab("校閲");
  await ribbonBtn("アクセシビリティ");
  await page.waitForSelector("#a11yPane:not([hidden]) .a11y-head");
  assert(await page.isVisible("#a11yTab"), "the tab shows");
  const list = await issues();
  assert(list.some((x) => x.rule === "order" && x.slide === blank), `reading order on the new slide: ${JSON.stringify(list.filter((x) => x.slide === blank))}`);
  assert((await page.locator("#a11yPane .a11y-level").first().getAttribute("class")).includes("error") || !list.some((x) => x.rule === "alt"), "errors first");
  await shot("checker");
});

await step("読み取り順序: 見た目の順にそろえる sets the order top to bottom (one ⌘Z), and the warning goes", async () => {
  await page.locator(`#a11yPane .a11y-item[data-rule="order"][data-slide="${blank}"] button:has-text("見た目の順")`).click();
  await page.waitForTimeout(300);
  const s = await slide();
  const words = s.readingOrder.map((id) => s.elements.find((o) => o.id === id).text.replace(/<[^>]+>/g, ""));
  assert(words.join() === "最初に読む,次に読む,最後に読む", `top to bottom: ${words}`);
  assert(!(await issues()).some((x) => x.rule === "order" && x.slide === blank), "the warning is gone");
  const dom = await page.$$eval("#stageBody .hs-objects > .hs-obj", (els) => els.map((el) => el.textContent.trim()));
  assert(dom.join() === "最初に読む,次に読む,最後に読む", `the page reads in that order: ${dom}`);
  await undo();
  assert(!(await slide()).readingOrder, "⌘Z takes it back");
  await page.locator(`#a11yPane .a11y-item[data-rule="order"][data-slide="${blank}"] button:has-text("見た目の順")`).click();
  await page.waitForTimeout(300);
});

await step("読み取り順序 list: ↓ reads the first one later", async () => {
  await page.locator("#a11yPane .a11y-read li").first().locator('button[aria-label="後に読む"]').click();
  await page.waitForTimeout(300);
  const s = await slide();
  const words = s.readingOrder.map((id) => s.elements.find((o) => o.id === id).text.replace(/<[^>]+>/g, ""));
  assert(words.join() === "次に読む,最初に読む,最後に読む", `moved: ${words}`);
  assert((await page.locator("#a11yPane .a11y-read li").first().textContent()).includes("次に読む"), "the list follows");
});

await step("代替テキスト: a picture without a description is an error; 説明を追加 fixes it, 装飾用にする hides another", async () => {
  await page.evaluate(() => { const s = window.__hpptx.slide(); s.elements.push({ id: "pic1", kind: "image", src: "asset:storeOperations", x: 1200, y: 200, w: 480, h: 320 }, { id: "pic2", kind: "image", src: "asset:storeOperations", x: 1200, y: 600, w: 480, h: 320 }); });
  await tab("校閲");
  await ribbonBtn("アクセシビリティ");
  await page.locator('#a11yPane button:has-text("再チェック")').click();
  await page.waitForTimeout(300);
  const alts = (await issues()).filter((x) => x.rule === "alt" && x.slide === blank);
  assert(alts.length === 2, `two pictures without text: ${alts.length}`);
  await page.locator(`#a11yPane .a11y-item[data-rule="alt"][data-slide="${blank}"]`).first().locator('button:has-text("説明を追加")').click();
  await page.waitForSelector(".alt-dialog[open]");
  await page.fill(".alt-dialog .alt-input", "セブン‐イレブンの店舗の外観");
  await page.click(".alt-dialog .alt-ok");
  await page.waitForTimeout(300);
  await page.locator(`#a11yPane .a11y-item[data-rule="alt"][data-slide="${blank}"]`).first().locator('button:has-text("装飾用")').click();
  await page.waitForTimeout(300);
  const s = await slide();
  const [p1, p2] = ["pic1", "pic2"].map((id) => s.elements.find((o) => o.id === id));
  assert([p1, p2].some((o) => o.alt === "セブン‐イレブンの店舗の外観") && [p1, p2].some((o) => o.decorative), `fixed: ${JSON.stringify([p1.alt, p1.decorative, p2.alt, p2.decorative])}`);
  assert(!(await issues()).some((x) => x.rule === "alt" && x.slide === blank), "no more alt errors on the slide");
  const described = [p1, p2].find((o) => o.alt);
  assert((await page.getAttribute(`#stageBody .hs-obj[data-el="${described.id}"] img`, "alt")) === described.alt, "the picture says it");
  await shot("checker-fixed");
});

await step("右クリック →「代替テキストを編集…」 opens the same dialog for the shape", async () => {
  await measure();
  const r = await page.locator("#stageBody .hs-obj").first().boundingBox();
  await page.mouse.click(r.x + r.width / 2, r.y + r.height / 2, { button: "right" });
  await page.locator('.ed-menu button:has-text("代替テキストを編集")').click();
  await page.waitForSelector(".alt-dialog[open]");
  await page.keyboard.press("Escape");
});

await step("印刷 (⌘P): layouts with a preview, a range, and the pages handed to the browser", async () => {
  await page.mouse.click(...at(1850, 1040));
  await page.keyboard.press("Control+p");
  await page.waitForSelector(".print-dialog[open] .pr-sheet");
  const total = (await deck()).slides.filter((s) => !s.hidden).length;
  assert((await page.textContent(".pr-count")).includes(`${total}枚`), `all slides: ${await page.textContent(".pr-count")}`);
  await page.selectOption('.print-dialog select[name="layout"]', "h6");
  await page.waitForTimeout(400);
  assert((await page.textContent(".pr-count")).includes(`${Math.ceil(total / 6)}ページ`), `six a page: ${await page.textContent(".pr-count")}`);
  assert((await page.locator(".pr-preview .pr-sheet").first().locator(".pr-slide").count()) === 6, "six slides on the first page");
  await shot("print-handout");
  await page.selectOption('.print-dialog select[name="range"]', "custom");
  await page.fill('.print-dialog input[name="custom"]', "2-4,6");
  await page.selectOption('.print-dialog select[name="layout"]', "notes");
  await page.waitForTimeout(400);
  assert((await page.textContent(".pr-count")).includes("4枚のスライド・4ページ"), `notes pages: ${await page.textContent(".pr-count")}`);
  await page.fill('.print-dialog input[name="custom"]', "abc");
  await page.waitForTimeout(400);
  assert((await page.textContent(".pr-preview")).includes("1,3,5-8"), "a bad range is explained");
  await page.fill('.print-dialog input[name="custom"]', "1-3");
  await page.selectOption('.print-dialog select[name="layout"]', "h3");
  await page.waitForTimeout(400);
  await page.evaluate(() => { window.__printed = null; window.print = () => { const r = document.getElementById("printRoot"); window.__printed = { pages: r.querySelectorAll(":scope > .pr-page").length, slides: r.querySelectorAll(".pr-slide").length, lines: r.querySelectorAll(".pr-lines").length, page: document.getElementById("printPageStyle")?.textContent }; }; });
  await page.click(".print-dialog .pr-go");
  await page.waitForFunction(() => window.__printed, null, { timeout: 8000 });
  const printed = await page.evaluate(() => window.__printed);
  assert(printed.pages === 1 && printed.slides === 3 && printed.lines === 3 && /A4 portrait/.test(printed.page), `printed: ${JSON.stringify(printed)}`);
});

await step("印刷 → コメントを印刷する: a page of comments (and replies) after the slide that has them", async () => {
  await page.evaluate(() => window.dispatchEvent(new Event("afterprint")));
  await page.evaluate(() => {
    const s = window.__hpptx.deck().slides[1];
    s.comments = [{ id: "qc1", text: "数字の出典を書く", by: "佐藤", at: new Date().toISOString(), replies: [{ id: "qr1", text: "注を足しました", by: "鈴木", at: new Date().toISOString() }] }, { id: "qc2", text: "済み", by: "佐藤", done: true }];
  });
  await page.mouse.click(...at(1850, 1040));
  await page.keyboard.press("Control+p");
  await page.waitForSelector(".print-dialog[open] .pr-sheet");
  await page.selectOption('.print-dialog select[name="range"]', "custom");
  await page.fill('.print-dialog input[name="custom"]', "1-3");
  await page.selectOption('.print-dialog select[name="layout"]', "full");
  await page.waitForTimeout(400);
  assert((await page.textContent(".pr-count")).includes("3枚のスライド・3ページ"), `without comments: ${await page.textContent(".pr-count")}`);
  await page.check('.print-dialog input[name="comments"]');
  await page.waitForTimeout(400);
  assert((await page.textContent(".pr-count")).includes("3枚のスライド・4ページ"), `a comments page: ${await page.textContent(".pr-count")}`);
  const sheets = await page.$$eval(".pr-preview .pr-sheet", (els) => els.map((el) => (el.querySelector(".pr-comments") ? "comments" : "slide")));
  assert(sheets.join() === "slide,slide,comments,slide", `after slide 2: ${sheets}`);
  const words = await page.textContent(".pr-preview .pr-comments");
  assert(words.includes("スライド 2 のコメント") && words.includes("佐藤") && words.includes("数字の出典を書く") && words.includes("注を足しました") && words.includes("解決済み"), `the comments, replies and who wrote them: ${words}`);
  await shot("print-comments");
  await page.selectOption('.print-dialog select[name="layout"]', "h6");
  await page.waitForTimeout(400);
  const handout = await page.$$eval(".pr-preview .pr-sheet", (els) => els.map((el) => (el.querySelector(".pr-comments") ? "comments" : "page")));
  assert(handout.join() === "page,comments", `handouts: the comments at the end: ${handout}`);
  assert(await page.locator(".pr-preview .pr-comments .pr-ph-pageNo").count(), "with the page's footer");
  await page.evaluate(() => { window.__printed = null; window.print = () => { const r = document.getElementById("printRoot"); window.__printed = { comments: r.querySelectorAll(".pr-comments li").length }; }; });
  await page.click(".print-dialog .pr-go");
  await page.waitForFunction(() => window.__printed, null, { timeout: 8000 });
  const printed = await page.evaluate(() => window.__printed);
  assert(printed.comments === 3, `printed: ${JSON.stringify(printed)}`);
  await page.evaluate(() => window.dispatchEvent(new Event("afterprint")));
  await page.evaluate(() => { delete window.__hpptx.deck().slides[1].comments; });
});

await step("スライド ショーの記録: narration and timings per slide, kept when asked; ナレーションの再生 and クリア", async () => {
  await page.evaluate(() => window.dispatchEvent(new Event("afterprint")));
  dialogs.length = 0;
  await tab("スライド ショー");
  await ribbonBtn("記録");
  await menuItem("先頭から記録");
  await page.waitForSelector("#presenter .hs-player");
  await page.waitForSelector(".rehearse-clock.recording");
  await page.waitForTimeout(1600);
  await page.keyboard.press("ArrowRight");
  for (let i = 0; i < 10 && (await page.textContent("#presenter .hs-player-count")).trim().startsWith("1 "); i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(250); }
  await page.waitForTimeout(1600);
  await page.keyboard.press("Escape");
  await page.waitForFunction(() => !document.querySelector("#presenter .hs-player"), null, { timeout: 5000 });
  await page.waitForTimeout(1500);
  assert(dialogs.some((m) => m.includes("ナレーション")), `asked: ${dialogs}`);
  const d = await deck();
  const narr = (i) => (d.slides[i].elements || []).filter((o) => o.narration);
  assert(narr(0).length === 1 && narr(1).length === 1, `one narration on slides 1 and 2: ${narr(0).length} ${narr(1).length}`);
  assert(narr(0)[0].autoplay && narr(0)[0].hideIcon && /^idb:/.test(narr(0)[0].src), `a hidden sound that starts with the slide: ${JSON.stringify(narr(0)[0])}`);
  assert(d.slides[0].advance >= 1 && d.slides[1].advance >= 1, `timings: ${d.slides[0].advance} ${d.slides[1].advance}`);
  await ribbonBtn("ナレーションの再生");
  assert((await deck()).show?.noNarration === true, "ナレーションの再生 off");
  await ribbonBtn("ナレーションの再生");
  await ribbonBtn("記録");
  await menuItem("すべてのスライドのナレーションをクリア");
  await page.waitForTimeout(300);
  assert(!(await deck()).slides.some((s) => (s.elements || []).some((o) => o.narration)), "narration cleared");
  await ribbonBtn("記録");
  await menuItem("すべてのスライドのタイミングをクリア");
  await page.waitForTimeout(300);
  assert(!(await deck()).slides.some((s) => s.advance != null), "timings cleared");
  await undo();
  assert((await deck()).slides[0].advance >= 1, "⌘Z brings the timings back");
});

await step("表記ゆれチェック: ＡＩ／AI and ユーザ／ユーザー found and made one (whole words, one ⌘Z)", async () => {
  await page.evaluate(() => {
    const s = window.__hpptx.slide();
    s.elements.push({ id: "pv1", kind: "text", x: 100, y: 900, w: 900, h: 80, text: "<p>ＡＩでユーザの声を集める</p>" }, { id: "pv2", kind: "text", x: 1000, y: 900, w: 800, h: 80, text: "<p>AIとユーザーとユーザビリティ</p>" });
  });
  await tab("校閲");
  await ribbonBtn("表記ゆれ");
  await page.waitForSelector(".proof-dialog[open]");
  const groups = await page.$$eval(".proof-dialog .proof-group", (els) => els.map((el) => [...el.querySelectorAll(".proof-word")].map((w) => w.textContent).sort().join("/")));
  assert(groups.includes("AI/ＡＩ") && groups.includes("ユーザ/ユーザー"), `found: ${groups}`);
  await shot("proofing");
  for (const word of ["AI", "ユーザー"]) {
    const group = page.locator(".proof-dialog .proof-group", { has: page.locator(`.proof-word:text-is("${word}")`) });
    await group.locator(`label:has(.proof-word:text-is("${word}")) input`).check();
    await group.locator(".proof-fix").click();
    await page.waitForTimeout(300);
  }
  const s = await slide();
  const text = (id) => s.elements.find((o) => o.id === id).text.replace(/<[^>]+>/g, "");
  assert(text("pv1") === "AIでユーザーの声を集める" && text("pv2") === "AIとユーザーとユーザビリティ", `unified: ${text("pv1")} / ${text("pv2")}`);
  await page.keyboard.press("Escape");
  await undo();
  assert((await slide()).elements.find((o) => o.id === "pv1").text.includes("ユーザの"), "⌘Z takes back the last one");
});

await step("スペル チェック: on, the text box being typed in is checked by the browser", async () => {
  await tab("校閲");
  await ribbonBtn("スペル");
  assert(await page.evaluate(() => localStorage.getItem("hpptx-spellcheck")) === "1", "remembered");
  await measure();
  const r = await page.locator('#stageBody .hs-obj[data-el="pv2"]').boundingBox();
  await page.mouse.dblclick(r.x + r.width / 2, r.y + r.height / 2);
  await page.waitForSelector("#stageBody .ed-typing-tx");
  assert(await page.evaluate(() => document.querySelector("#stageBody .ed-typing-tx").spellcheck), "spellcheck on while typing");
  await page.keyboard.press("Escape");
  await tab("校閲");
  await ribbonBtn("スペル");
});

await step("ファイル → 情報: properties saved on the deck, statistics, and ドキュメント検査 removes notes", async () => {
  await page.click(".rb-file");
  await page.locator('.rb-pop .rb-menu button:has-text("情報")').click();
  await page.waitForSelector(".fileinfo-dialog[open]");
  assert((await page.textContent(".fileinfo-dialog .fi-stats")).includes("スライド"), "statistics");
  const notesRow = page.locator('.fileinfo-dialog .fi-row[data-inspect="notes"]');
  assert(await notesRow.locator("button").count(), "notes are found");
  const author = page.locator(".fileinfo-dialog .fi-fields label", { hasText: "作成者" }).locator("input");
  await author.fill("店舗運営部 山田");
  await page.locator(".fileinfo-dialog .fi-fields label", { hasText: "キーワード" }).locator("input").fill("店舗,生成AI");
  await shot("fileinfo");
  await notesRow.locator("button").click();
  await page.waitForTimeout(300);
  assert(!(await deck()).slides.some((s) => String(s.notes || "").trim()), "notes removed");
  assert(!(await notesRow.locator("button").count()), "the row says none");
  await page.click(".fileinfo-dialog .fi-ok");
  await page.waitForTimeout(300);
  const info = (await deck()).info;
  assert(info?.author === "店舗運営部 山田" && info.keywords === "店舗,生成AI", `saved: ${JSON.stringify(info)}`);
  await undo();
  assert(!(await deck()).info, "⌘Z takes the properties back");
  await undo();
  assert((await deck()).slides.some((s) => String(s.notes || "").trim()), "and the notes");
});

await step("ファイル → 情報 → 最終版にする: the message bar says so, keys and the ribbon do not edit; 編集する lets them back", async () => {
  await page.locator(".film-item").nth(1).click();
  await page.click(".rb-file");
  await page.locator('.rb-pop .rb-menu button:has-text("情報")').click();
  await page.waitForSelector(".fileinfo-dialog[open]");
  await page.click(".fileinfo-dialog .fi-final");
  await page.waitForSelector("#msgBar .callout-final");
  assert((await deck()).info?.status === "最終版", "the document's 状態");
  assert(await page.evaluate(() => document.body.classList.contains("is-final")), "the page knows");
  const slides = (await deck()).slides.length;
  await page.locator(".film-item").nth(1).click();
  await page.keyboard.press("Control+d");
  await page.keyboard.press("Delete");
  await page.waitForTimeout(300);
  assert((await deck()).slides.length === slides, "no slide duplicated or deleted");
  await page.click('.rb-tabs [role=tab]:text-is("ホーム")');
  assert(await page.locator(".rb-body > .rb-group.rb-off").count() > 2, "the ホーム tab is off");
  await page.click('.rb-tabs [role=tab]:text-is("スライド ショー")');
  assert(!(await page.locator(".rb-body > .rb-group.rb-off").count()), "the スライド ショー tab still works");
  // The notes, the panes and the search's commands take nothing either.
  const notes = (await deck()).slides[1].notes || "";
  assert(await page.evaluate(() => document.getElementById("notesInput").isContentEditable === false), "the notes are read-only");
  await page.evaluate(() => document.getElementById("notesInput").focus());
  await page.keyboard.type("書き足し");
  assert(((await deck()).slides[1].notes || "") === notes, "nothing typed into the notes");
  assert(await page.evaluate(() => getComputedStyle(document.querySelector("#animPane") || document.body).opacity !== "" && [...document.querySelectorAll("#inspector button, #animPane button, #commentPane button")].every((b) => getComputedStyle(b).pointerEvents === "none")), "the panes' buttons do not answer");
  await page.keyboard.press("Escape");
  await page.keyboard.press("Alt+KeyQ");
  await page.keyboard.type("新しいスライド");
  await page.waitForSelector(".rb-pop .rb-search-item");
  await page.keyboard.press("Enter");
  await page.waitForTimeout(400);
  assert((await deck()).slides.length === slides, "a command from the search does not edit");
  await page.keyboard.press("Escape");
  await shot("final");
  await page.click("#msgBar .btn-final-edit");
  await page.waitForTimeout(300);
  assert(!(await deck()).info?.status && !(await page.evaluate(() => document.body.classList.contains("is-final"))), "編集する: editable again");
  assert(!(await page.locator("#msgBar .callout-final").count()), "the bar goes");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
