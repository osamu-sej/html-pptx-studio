// More of PowerPoint, in a real browser: sections in the thumbnails (added, folded, renamed, swapped, in the slide
// sorter too) and dragging slides in the sorter; review comments (校閲 → 新しいコメント: the writer's name, replies,
// @mentions, editing, pins on objects, resolve, delete, the thumbnail's mark); rehearse timings (スライド ショー → リハーサル: the times become 自動で切り替え); the 開発 tab
// (the slide's JSON edited and applied, checked first; the HTML it becomes); ⌘H for find and replace; the ファイル
// menu; guides (added, dragged with the distance shown, snapped to, dragged off to delete, kept).
// Usage: node qa/studio-powerpoint.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
let lastDialog = "";
page.on("dialog", (dialog) => { lastDialog = dialog.message(); dialog.accept(); });
const shot = async (name) => { const file = join(outDir, `powerpoint-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
  for (let i = 0; i < 3 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
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
const answer = async (text) => {
  await page.waitForSelector("#askDialog[open]");
  await page.fill("#askInput", text);
  await page.click("#askOk");
  await page.waitForTimeout(300);
};
const film = (n) => page.click(`.film-item:nth-child(${n})`);
const filmByIndex = (index) => page.locator(".film-item").nth(index);

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.click("#sampleDeckBtn");
await page.waitForSelector(".slide-wrap .hs-slide");
await page.waitForTimeout(500);

await step("セクション: added from the ribbon and from a thumbnail's menu; the thumbnails show and fold them", async () => {
  await filmByIndex(2).click();
  await tab("ホーム");
  await ribbonBtn("セクション");
  await answer("背景");
  let d = await deck();
  assert(d.slides[2].section === "背景", "section on slide 3");
  const heads = await page.$$eval(".film-section b", (els) => els.map((el) => el.textContent));
  assert(heads.join() === "既定のセクション,背景", `headers: ${heads}`);
  await filmByIndex(6).click({ button: "right" });
  await page.click('.ed-menu button:has-text("セクションの追加")');
  await answer("打ち手");
  d = await deck();
  assert(d.slides[6].section === "打ち手", "section on slide 7");
  const shown = await page.locator(".film-item").count();
  await page.click('.film-section:has(b:text-is("背景"))');
  await page.waitForTimeout(200);
  assert(await page.locator(".film-item").count() === shown - 4, "folded: its 4 slides are out of the way");
  await shot("sections-folded");
  await page.click('.film-section:has(b:text-is("背景"))');
  await page.waitForTimeout(200);
  assert(await page.locator(".film-item").count() === shown, "unfolded");
});

await step("セクション: renamed, swapped with the one before, removed (the slides stay); ⌘Z", async () => {
  await page.click('.film-section:has(b:text-is("背景"))', { button: "right" });
  await page.click('.ed-menu button:has-text("セクション名の変更")');
  await answer("現状と背景");
  assert((await deck()).slides[2].section === "現状と背景", "renamed");
  const before = (await deck()).slides.map((s) => s.title);
  await page.click('.film-section:has(b:text-is("打ち手"))', { button: "right" });
  await page.click('.ed-menu button:has-text("セクションを上へ移動")');
  await page.waitForTimeout(300);
  const after = await deck();
  assert(after.slides[2].section === "打ち手" && after.slides[2].title === before[6], `打ち手 comes first: ${after.slides[2].title}`);
  const back = after.slides.findIndex((s) => s.section === "現状と背景");
  assert(after.slides[back].title === before[2], "現状と背景 follows with its slides");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  assert((await deck()).slides.map((s) => s.title).join() === before.join(), "⌘Z puts them back");
  await page.click('.film-section:has(b:text-is("打ち手"))', { button: "right" });
  await page.click('.ed-menu button:has-text("セクションの削除")');
  await page.waitForTimeout(300);
  const d = await deck();
  assert(!d.slides.some((s) => s.section === "打ち手") && d.slides.length === before.length, "removed, every slide kept");
});

await step("スライド一覧: the sections over their slides; drag a slide to move it", async () => {
  await page.click('#statusBar label:has-text("一覧")');
  await page.waitForTimeout(400);
  assert(await page.locator(".grid-section").count() >= 2, "sections in the sorter");
  const before = (await deck()).slides.map((s) => s.title);
  await page.locator(".grid-item").nth(5).dragTo(page.locator(".grid-item").nth(3));
  await page.waitForTimeout(400);
  const after = (await deck()).slides.map((s) => s.title);
  assert(after[3] === before[5], `slide 6 moved before slide 4: ${after.slice(2, 7)}`);
  await shot("sorter");
  await page.click('#statusBar label:has-text("1枚")');
  await page.waitForTimeout(300);
});

await step("校閲 → 新しいコメント: the writer's name asked once, the thumbnail's mark, resolve, delete", async () => {
  await filmByIndex(1).click();
  await tab("校閲");
  await ribbonBtn("新しい");
  await page.waitForSelector("#commentPane:not([hidden])");
  await page.waitForTimeout(150);
  assert(await page.evaluate(() => document.activeElement?.classList.contains("cm-input")), "the box is ready to type");
  assert(/名前が未設定/.test(await page.textContent("#commentPane .cm-who")), "no name yet");
  await page.keyboard.type("数字の出所を確認してください");
  await page.click('#commentPane button.cm-post');
  await answer("山田 太郎");
  let d = await deck();
  const c = d.slides[1].comments?.[0];
  assert(d.slides[1].comments?.length === 1 && c.text === "数字の出所を確認してください", "posted");
  assert(c.by === "山田 太郎" && /^[a-z0-9]{6,24}$/.test(c.uid) && c.at, `who and when: ${JSON.stringify(c)}`);
  assert((await page.textContent("#commentPane .cm-item .cm-by")) === "山田 太郎", "the name on the thread");
  assert((await page.textContent("#commentPane .cm-item .pp-avatar")) === "山田", "an avatar with the initials");
  assert(await page.evaluate(() => localStorage.getItem("hpptx-user-name")) === "山田 太郎", "the name is kept for next time");
  assert(await filmByIndex(1).locator(".flag-comment").count(), "the thumbnail shows it");
  await page.click('#commentPane .cm-item button:has-text("解決")');
  await page.waitForTimeout(300);
  d = await deck();
  assert(d.slides[1].comments[0].done === true && d.slides[1].comments[0].doneBy === "山田 太郎", "resolved, and by whom");
  assert(/解決済み（山田 太郎）/.test(await page.textContent("#commentPane .cm-item")), "shown as resolved by the name");
  assert(!(await filmByIndex(1).locator(".flag-comment").count()), "a resolved comment drops the mark");
  await page.click('#commentPane .cm-item button:has-text("もう一度開く")');
  await page.waitForTimeout(300);
  assert(!(await deck()).slides[1].comments[0].done, "opened again");
  await filmByIndex(4).click();
  await page.click('#commentPane button:has-text("前へ")');
  await page.waitForTimeout(300);
  assert((await page.textContent("#slidePos")).startsWith("スライド 2 /"), "前へ goes to the slide with comments");
});

await step("コメント: someone else's thread, @ suggestions in a reply, replies, editing your own, the 自分宛て filter", async () => {
  // A colleague's comment (as it arrives from a shared session or a PowerPoint file), added through the slide's JSON.
  await tab("開発");
  await ribbonBtn("JSON");
  await page.waitForSelector("#codeDialog[open]");
  const data = JSON.parse(await page.inputValue("#codeText"));
  data.comments.push({ id: "sato1", text: "@山田 太郎 表の単位は？", by: "佐藤", uid: "sato000001", at: "2026-09-01T09:00:00.000Z" });
  await page.fill("#codeText", JSON.stringify(data));
  await page.click("#codeApply");
  await page.waitForTimeout(300);
  await tab("校閲");
  await ribbonBtn("新しい");
  await page.waitForTimeout(200);
  const sato = page.locator('#commentPane .cm-item[data-id="sato1"]');
  assert(await sato.locator(".pp-mention.me").count(), "a mention of me is marked");
  assert(await sato.evaluate((el) => el.classList.contains("to-me")), "and the thread shows it is for me");
  assert(!(await sato.locator('button:has-text("編集")').count()), "someone else's comment cannot be edited");
  await sato.locator('button:has-text("返信")').click();
  await page.waitForTimeout(150);
  await page.keyboard.type("@佐");
  await page.waitForSelector("#commentPane .cm-mentions li");
  assert(/佐藤/.test(await page.textContent("#commentPane .cm-mentions")), "the names to pick after @");
  await page.keyboard.press("Enter");
  await page.keyboard.type("単位は億円です");
  assert(await page.inputValue("#commentPane .cm-reply-input") === "@佐藤 単位は億円です", "the picked name is put in");
  await page.click("#commentPane .cm-reply-post");
  await page.waitForTimeout(300);
  let d = await deck();
  const reply = d.slides[1].comments.find((x) => x.id === "sato1").replies?.[0];
  assert(reply && reply.text === "@佐藤 単位は億円です" && reply.by === "山田 太郎", `the reply: ${JSON.stringify(reply)}`);
  assert(await page.locator('#commentPane .cm-item[data-id="sato1"] .cm-reply .pp-mention').count(), "the reply's @name is marked");
  await page.locator("#commentPane .cm-item").first().locator('button:has-text("編集")').first().click();
  await page.fill("#commentPane .cm-edit-input", "数字の出所（年度）を確認してください");
  await page.click("#commentPane .cm-save");
  await page.waitForTimeout(300);
  d = await deck();
  assert(d.slides[1].comments[0].text === "数字の出所（年度）を確認してください" && d.slides[1].comments[0].edited, "edited, and marked as edited");
  assert(/編集済み/.test(await page.textContent("#commentPane .cm-item")), "（編集済み）");
  await page.selectOption("#commentPane .cm-filter", "mine");
  await page.waitForTimeout(150);
  assert(await page.locator("#commentPane .cm-item").count() === 1, "自分宛て: only the thread that names me");
  await page.selectOption("#commentPane .cm-filter", "all");
  await shot("comment-threads");
});

await step("コメント: pinned to the selected object, pins on the slide open their thread, shown or hidden", async () => {
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await page.click('.rb-pop .rb-menu button:has-text("横書き")');
  const slide = await page.locator(".slide-wrap .hs-scaler").boundingBox();
  const k = slide.width / 1920;
  await page.mouse.move(slide.x + 1300 * k, slide.y + 700 * k);
  await page.mouse.down();
  await page.mouse.move(slide.x + 1650 * k, slide.y + 790 * k, { steps: 4 });
  await page.mouse.up();
  await page.keyboard.type("注記");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  const box = (await deck()).slides[1].elements.at(-1);
  assert((await page.evaluate(() => window.__hpptx.selection())).join() === box.id, "the box stays selected");
  await tab("校閲");
  await ribbonBtn("新しい");
  await page.waitForTimeout(200);
  assert(/選んでいるテキスト「注記」に付けます/.test(await page.textContent("#commentPane .cm-new")), "the box says where it goes");
  await page.keyboard.type("この注記は不要では？");
  await page.click("#commentPane button.cm-post");
  await page.waitForTimeout(300);
  const d = await deck();
  const pinned = d.slides[1].comments.at(-1);
  assert(pinned.anchor === box.id, `pinned to the box: ${JSON.stringify(pinned)}`);
  assert(/📌 テキスト「注記」/.test(await page.textContent(`#commentPane .cm-item[data-id="${pinned.id}"]`)), "the thread names the object");
  assert(await page.locator(".ed-comment-pin").count() === 3, `a pin per open thread: ${await page.locator(".ed-comment-pin").count()}`);
  const pin = await page.locator(`.ed-comment-pin[data-id="${pinned.id}"]`).boundingBox();
  const b = await page.locator(`.slide-wrap .hs-obj[data-el="${box.id}"]`).boundingBox();
  assert(Math.abs(pin.x - (b.x + b.width)) < 20 && pin.y < b.y + 4, `the pin sits at the box's top right: ${JSON.stringify(pin)} vs ${JSON.stringify(b)}`);
  await shot("comment-pins");
  await page.click(`.ed-comment-pin[data-id="sato1"]`);
  await page.waitForTimeout(150);
  assert(await page.locator('#commentPane .cm-item[data-id="sato1"].flash').count(), "a pin brings its thread forward");
  await ribbonBtn("コメントの表示");
  await page.click('.rb-pop .rb-menu button:has-text("スライド上にコメントを表示")');
  await page.waitForTimeout(200);
  assert(!(await page.locator(".ed-comment-pin").count()), "pins hidden");
  await ribbonBtn("コメントの表示");
  await page.click('.rb-pop .rb-menu button:has-text("スライド上にコメントを表示")');
  await page.waitForTimeout(200);
  assert(await page.locator(".ed-comment-pin").count() === 3, "and shown again");
});

await step("ユーザー名: renamed from ファイル (your comments follow); 校閲 → 削除 clears the slide's comments; ⌘Z", async () => {
  await page.click("#ribbon .rb-file");
  await page.click('.rb-pop .rb-menu button:has-text("ユーザー名")');
  await answer("山田 花子");
  let d = await deck();
  const byName = d.slides[1].comments.map((x) => x.by);
  assert(byName[0] === "山田 花子" && byName.includes("佐藤"), `mine renamed, the colleague's kept: ${byName}`);
  assert(d.slides[1].comments.find((x) => x.id === "sato1").replies[0].by === "山田 花子", "my reply too");
  await tab("校閲");
  await ribbonBtn("削除");
  await page.click('.rb-pop .rb-menu button:has-text("このスライドのコメントをすべて削除")');
  await page.waitForTimeout(300);
  assert(!(await deck()).slides[1].comments, "all deleted");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  assert((await deck()).slides[1].comments.length === 3, "⌘Z brings them back");
  await page.locator('#commentPane .cm-item[data-id="sato1"] button:has-text("スレッドを削除")').click();
  await page.waitForTimeout(300);
  assert((await deck()).slides[1].comments.length === 2, "a thread deleted");
});

await step("スライド ショー → リハーサル: the slides are timed; the times become 自動で切り替え", async () => {
  await tab("スライド ショー");
  await ribbonBtn("リハーサル");
  await page.waitForSelector("#presenter .rehearse-clock", { timeout: 10000 });
  await page.waitForTimeout(1300);
  assert(/リハーサル/.test(await page.textContent(".rehearse-clock")), "the clock");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(800);
  assert(/リハーサルの時間/.test(lastDialog), `asked to keep the times: ${lastDialog}`);
  const d = await deck();
  assert(d.slides[0].advance >= 1, `slide 1 moves on by itself after its rehearsed time: ${d.slides[0].advance}`);
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  assert(!(await deck()).slides[0].advance, "⌘Z takes the timings back");
});

await step("開発 → スライドのJSON: edited and applied (checked first); HTMLを表示", async () => {
  await filmByIndex(3).click();
  await tab("開発");
  await ribbonBtn("JSON");
  await page.waitForSelector("#codeDialog[open]");
  const text = await page.inputValue("#codeText");
  const data = JSON.parse(text);
  assert(data.type && data.title, "the slide's data");
  await page.fill("#codeText", "{ broken");
  await page.click("#codeApply");
  assert(await page.isVisible("#codeDialog[open]") && /JSONとして読めません/.test(await page.textContent("#codeStatus")), "broken JSON is refused");
  await page.fill("#codeText", JSON.stringify({ ...data, type: "nope" }));
  await page.click("#codeApply");
  assert(/未対応のレイアウト/.test(await page.textContent("#codeStatus")), "an unknown layout is refused");
  await page.fill("#codeText", JSON.stringify({ ...data, title: "JSONで直したタイトル" }, null, 2));
  await page.click("#codeApply");
  await page.waitForTimeout(300);
  assert(await page.isHidden("#codeDialog[open]"), "applied and closed");
  assert((await deck()).slides[3].title === "JSONで直したタイトル", "the slide changed");
  assert((await page.textContent(".slide-wrap .hs-slide")).includes("JSONで直したタイトル"), "and shows it");
  await ribbonBtn("HTMLを");
  await page.waitForSelector("#codeDialog[open]");
  assert(await page.evaluate(() => document.getElementById("codeText").readOnly), "read-only HTML");
  assert(/class="hs-slide/.test(await page.inputValue("#codeText")), "the engine's HTML");
  await shot("code-html");
  await page.keyboard.press("Escape");
});

await step("⌘H opens find and replace", async () => {
  await page.click("#stageBody", { position: { x: 5, y: 5 } });
  await page.keyboard.press("Control+h");
  await page.waitForSelector("#replaceDialog[open]");
  await page.keyboard.press("Escape");
});

await step("ファイル: the menu at the start of the ribbon (open the help, back to 作成)", async () => {
  await page.click("#ribbon .rb-file");
  await page.waitForSelector(".rb-pop .rb-menu");
  for (const label of ["新規", "開く", "上書き保存", "名前を付けて保存", "HTMLファイル", "PDF", "JSONで保存", "版の履歴"]) assert(await page.locator(`.rb-pop .rb-menu button:has-text("${label}")`).count(), `menu: ${label}`);
  await page.click('.rb-pop .rb-menu button:has-text("使い方とショートカット")');
  await page.waitForSelector("#helpDialog[open]");
  await page.keyboard.press("Escape");
});

await step("ガイド: added from 表示 and from a right-click, dragged (cm from the middle), objects snap to it, dragged off to delete, kept", async () => {
  await filmByIndex(3).click();
  await page.waitForTimeout(300);
  await tab("表示");
  await ribbonBtn("ガイド");
  await page.click('.rb-pop .rb-menu button:has-text("垂直ガイドを追加")');
  await page.waitForTimeout(300);
  let d = await deck();
  assert(JSON.stringify(d.guides) === JSON.stringify({ x: [960], y: [] }), `a vertical guide in the middle: ${JSON.stringify(d.guides)}`);
  const guide = page.locator(".ed-cguide.v").first();
  const g = await guide.boundingBox();
  const slide = await page.locator(".slide-wrap .hs-scaler").boundingBox();
  const k = slide.width / 1920;
  await page.mouse.move(g.x + g.width / 2, slide.y + 200);
  await page.mouse.down();
  await page.mouse.move(g.x + g.width / 2 - 150 * k, slide.y + 200, { steps: 6 });
  assert(/cm/.test(await page.textContent(".ed-cguide-label")), "the distance from the middle while dragging");
  await shot("guide-drag");
  await page.mouse.up();
  await page.waitForTimeout(300);
  d = await deck();
  assert(Math.abs(d.guides.x[0] - 810) <= 2, `moved to 810: ${d.guides.x}`);
  // A right-click on the slide's background adds a horizontal guide there.
  await page.mouse.click(slide.x + 40 * k, slide.y + 600 * k, { button: "right" });
  await page.click('.ed-menu button:has-text("水平ガイドを追加")');
  await page.waitForTimeout(300);
  d = await deck();
  assert(d.guides.y.length === 1 && Math.abs(d.guides.y[0] - 600) <= 2, `horizontal guide at 600: ${d.guides.y}`);
  // Drag an object so its left edge comes near the guide: it snaps onto it.
  await tab("挿入");
  await ribbonBtn("テキスト ボックス");
  await page.click('.rb-pop .rb-menu button:has-text("横書き")');
  await page.mouse.move(slide.x + 300 * k, slide.y + 300 * k);
  await page.mouse.down();
  await page.mouse.move(slide.x + 600 * k, slide.y + 380 * k, { steps: 4 });
  await page.mouse.up();
  await page.keyboard.type("吸着");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  d = await deck();
  const box = d.slides[3].elements.at(-1);
  const b = await page.locator(`.slide-wrap .hs-obj[data-el="${box.id}"]`).boundingBox();
  const at = d.guides.x[0];
  const dx = (at + 4 - box.x) * k;
  await page.mouse.move(b.x + b.width / 2, b.y + b.height / 2);
  await page.mouse.down();
  await page.mouse.move(b.x + b.width / 2 + dx / 2, b.y + b.height / 2, { steps: 3 });
  await page.mouse.move(b.x + b.width / 2 + dx, b.y + b.height / 2, { steps: 3 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  const snapped = (await deck()).slides[3].elements.find((o) => o.id === box.id);
  assert(Math.abs(snapped.x - at) <= 0.6, `the box's left edge snaps to the guide at ${at}: ${snapped.x}`);
  // Off the slide: deleted; ⌘Z brings it back; a reload keeps the guides.
  const g2 = await page.locator(".ed-cguide.v").first().boundingBox();
  await page.mouse.move(g2.x + g2.width / 2, slide.y + 100);
  await page.mouse.down();
  await page.mouse.move(slide.x - 40, slide.y + 100, { steps: 6 });
  await page.mouse.up();
  await page.waitForTimeout(300);
  assert(!(await deck()).guides?.x?.length, "dragged off the slide: deleted");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  assert((await deck()).guides.x.length === 1, "⌘Z brings it back");
  await page.reload();
  await page.waitForSelector(".slide-wrap .hs-scaler");
  assert(await page.locator(".ed-cguide").count() === 2 || (await deck()).guides.x.length === 1, "kept after a reload");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
