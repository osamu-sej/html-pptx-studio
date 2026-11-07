// Sound and video in a real browser: 挿入 → オーディオ (a file from this device), the 再生 tab (開始, play across
// slides, hide the icon, fades, volume, trimming, the preview), the sound in the motion IDE's list, a slide show
// where the sound starts by itself and keeps playing on the next slides, recording a sound and the screen here,
// and the exported HTML carrying the sound.
// Usage: node qa/studio-media.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir, readFile, writeFile } from "node:fs/promises";
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

// A two-second tone as a WAV file (8 kHz, 16-bit, mono).
const rate = 8000;
const samples = rate * 2;
const wav = Buffer.alloc(44 + samples * 2);
wav.write("RIFF", 0); wav.writeUInt32LE(36 + samples * 2, 4); wav.write("WAVE", 8); wav.write("fmt ", 12);
wav.writeUInt32LE(16, 16); wav.writeUInt16LE(1, 20); wav.writeUInt16LE(1, 22); wav.writeUInt32LE(rate, 24); wav.writeUInt32LE(rate * 2, 28);
wav.writeUInt16LE(2, 32); wav.writeUInt16LE(16, 34); wav.write("data", 36); wav.writeUInt32LE(samples * 2, 40);
for (let i = 0; i < samples; i += 1) wav.writeInt16LE(Math.round(Math.sin((2 * Math.PI * 440 * i) / rate) * 8000), 44 + i * 2);
const wavFile = join(outDir, "media-tone.wav");
await writeFile(wavFile, wav);

const browserArgs = [
  "--use-fake-device-for-media-stream", "--use-fake-ui-for-media-stream", "--autoplay-policy=no-user-gesture-required",
  "--auto-select-desktop-capture-source=Entire screen", "--enable-usermedia-screen-capturing",
  ...(process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : []),
];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, permissions: ["microphone"] });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name) => { const file = join(outDir, `media-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  for (let i = 0; i < 3 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deck = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
const tab = (label) => page.click(`.rb-tabs [role=tab]:text-is("${label}")`);
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
const ribbonField = async (label) => {
  const direct = page.locator(`.rb-body .rb-field:has(span:text-is("${label}"))`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct;
  for (const folded of await page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = page.locator(`.rb-pop .rb-fold .rb-field:has(span:text-is("${label}"))`).first();
    if (await inside.count()) return inside;
  }
  throw new Error(`no ribbon field "${label}"`);
};
const slideIndex = 2;
// Select an object from the 書式 pane's list (as the selection pane does).
const pickObject = async (id) => {
  if ((await page.getAttribute("#formatTab", "aria-selected")) !== "true") await page.click("#formatTab");
  await page.locator(`#formatPane .fp-sel-list li[data-id="${id}"]`).click();
  await page.waitForTimeout(200);
};
const sound = async () => (await deck()).slides[slideIndex].elements?.find((o) => o.kind === "audio");

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.click("#sampleDeckBtn");
await page.waitForSelector(".slide-wrap .hs-slide");
await page.waitForTimeout(500);
await page.locator(".film-item").nth(slideIndex).click();
await page.waitForTimeout(300);

await step("挿入 → オーディオ → このデバイスのオーディオ: a speaker icon in the middle of the slide; the 再生 tab", async () => {
  await tab("挿入");
  await ribbonBtn("オーディオ");
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), page.click('.rb-pop .rb-menu button:has-text("このデバイスのオーディオ")')]);
  await chooser.setFiles(wavFile);
  await page.waitForTimeout(800);
  const o = await sound();
  assert(o && o.src.startsWith("idb:") && o.fileName === "media-tone.wav", `the sound: ${JSON.stringify(o)}`);
  assert(Math.abs(o.x + o.w / 2 - 960) < 2 && Math.abs(o.y + o.h / 2 - 540) < 2, "in the middle of the slide");
  assert(await page.locator(`.slide-wrap .hs-obj[data-el="${o.id}"] .hs-audio-icon`).count(), "the speaker icon on the slide");
  assert(await page.locator('.rb-tabs [role=tab]:text-is("再生")').count(), "the 再生 tab appears");
  assert(await page.locator(".film-item").nth(slideIndex).locator('span[title="オーディオあり"]').count(), "the thumbnail shows a sound");
  await shot("inserted");
});

await step("再生 → 開始: 一連のクリック動作 adds a 再生 to the click order; 自動 plays with the slide; one undo step each", async () => {
  await tab("再生");
  const start = (await ribbonField("開始")).locator("select");
  await start.selectOption("sequence");
  await page.waitForTimeout(300);
  let d = await deck();
  let o = d.slides[slideIndex].elements.find((x) => x.kind === "audio");
  assert((d.slides[slideIndex].timeline || []).some((e) => e.el === o.id && e.cls === "media" && e.fx === "play"), "a 再生 in the click order");
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
  await (await ribbonField("開始")).locator("select").selectOption("auto");
  await page.waitForTimeout(300);
  d = await deck();
  o = d.slides[slideIndex].elements.find((x) => x.kind === "audio");
  assert(o.autoplay === true && !(d.slides[slideIndex].timeline || []).some((e) => e.el === o.id), "自動: plays with the slide, out of the click order");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  d = await deck();
  o = d.slides[slideIndex].elements.find((x) => x.kind === "audio");
  assert(!o.autoplay && (d.slides[slideIndex].timeline || []).some((e) => e.el === o.id), "⌘Z goes back to the click order in one step");
  await page.keyboard.press("Control+y");
  await page.waitForTimeout(300);
  assert((await sound()).autoplay === true, "⌘Y");
});

await step("再生: play across slides, hide the icon, loop, fades, volume", async () => {
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
  await (await ribbonField("スライド切り替え後も再生")).locator("input").check();
  await page.waitForTimeout(250);
  await (await ribbonField("再生中のアイコンを隠す")).locator("input").check();
  await page.waitForTimeout(250);
  await (await ribbonField("停止するまで繰り返す")).locator("input").check();
  await page.waitForTimeout(250);
  const fade = (await ribbonField("フェードイン")).locator("input");
  await fade.fill("0.5");
  await fade.press("Enter");
  await page.waitForTimeout(250);
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
  await ribbonBtn("音量");
  await page.click('.rb-pop .rb-menu button:has-text("小")');
  await page.waitForTimeout(250);
  const o = await sound();
  assert(o.across === 999 && o.hideIcon === true && o.loop === true && o.fadeIn === 0.5 && o.volume === 0.33, `settings: ${JSON.stringify(o)}`);
});

await step("再生 → トリミング: the kept part's start and end", async () => {
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
  await ribbonBtn("トリミング");
  await page.waitForSelector(".trim-dialog[open]");
  await page.waitForFunction(() => /元の長さ 0:02/.test(document.querySelector(".trim-info")?.textContent || ""), null, { timeout: 5000 });
  await page.fill(".trim-dialog input[aria-label='開始時間（秒）']", "0.3");
  await page.press(".trim-dialog input[aria-label='開始時間（秒）']", "Tab");
  await page.fill(".trim-dialog input[aria-label='終了時間（秒）']", "1.6");
  await page.press(".trim-dialog input[aria-label='終了時間（秒）']", "Tab");
  assert(/長さ 0:01\.3/.test(await page.textContent(".trim-info")), `the kept length: ${await page.textContent(".trim-info")}`);
  await shot("trim");
  await page.click(".trim-dialog .trim-ok");
  await page.waitForTimeout(300);
  const o = await sound();
  assert(o.trimStart === 0.3 && o.trimEnd === 1.6, `trimmed: ${o.trimStart}–${o.trimEnd}`);
});

await step("再生 → プレビュー: plays here, trimmed, then stops", async () => {
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
  await ribbonBtn("再生");
  await page.waitForTimeout(400);
  const playing = await page.evaluate(() => [...document.querySelectorAll("body > audio")].some((a) => !a.paused && a.currentTime >= 0.3));
  assert(playing, "the sound plays from the trimmed start");
  assert(await page.locator('.rb-body .rb-btn[aria-pressed="true"]:has-text("再生")').count(), "the button shows it is playing");
  await page.waitForTimeout(1500);
  const t = await page.evaluate(() => [...document.querySelectorAll("body > audio")].map((a) => a.currentTime)[0]);
  assert(t >= 0.3 && t < 1.65, `繰り返す: it loops inside the trimmed part (${t})`);
  await ribbonBtn("再生");
  await page.waitForTimeout(200);
  assert(!(await page.evaluate(() => [...document.querySelectorAll("body > audio")].some((a) => !a.paused))), "a second click stops it");
});

await step("the motion IDE lists the sound with how it starts (and changes it)", async () => {
  await tab("アニメーション");
  await ribbonBtn("ウィンドウ");
  await page.waitForSelector("#animPane:not([hidden])");
  const row = page.locator('#animPane select[data-auto^="media:"]');
  assert(await row.count() === 1, "a row for the sound");
  assert(await row.inputValue() === "auto", "自動で再生");
  assert(/🔊/.test(await page.textContent("#animPane .an-auto")), "marked as a sound");
  await row.selectOption("click");
  await page.waitForTimeout(300);
  assert((await sound()).autoplay === false, "changed from the list");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(300);
  assert((await sound()).autoplay === true, "⌘Z");
});

await step("slide show: the sound starts by itself, its icon hidden, and keeps playing on the next slides; ending stops it", async () => {
  await tab("スライド ショー");
  await ribbonBtn("このスライド");
  await page.waitForSelector("#presenter .hs-player-slide audio", { state: "attached", timeout: 10000 });
  await page.waitForTimeout(900);
  let state = await page.evaluate(() => { const a = document.querySelector("#presenter .hs-player-slide audio"); return { paused: a.paused, t: a.currentTime, vol: a.volume, hidden: getComputedStyle(document.querySelector("#presenter .hs-audio-icon")).visibility }; });
  assert(!state.paused && state.t >= 0.3, `playing from the trimmed start: ${JSON.stringify(state)}`);
  assert(state.vol <= 0.34, `at the 小 volume: ${state.vol}`);
  assert(state.hidden === "hidden", "the icon is hidden while presenting");
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(400);
  for (let i = 0; i < 6 && await page.locator("#presenter .hs-player-slide audio").count(); i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(400); }
  state = await page.evaluate(() => { const a = document.querySelector("#presenter .hs-carry audio"); return a ? { paused: a.paused, t: a.currentTime } : null; });
  assert(state && !state.paused, `still playing on the next slide: ${JSON.stringify(state)}`);
  await page.waitForTimeout(1600);
  state = await page.evaluate(() => { const a = document.querySelector("#presenter .hs-carry audio"); return { paused: a.paused, t: a.currentTime }; });
  assert(!state.paused && state.t < 1.65, `loops inside the trim: ${JSON.stringify(state)}`);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(500);
  assert(!(await page.evaluate(() => [...document.querySelectorAll("audio")].some((a) => !a.paused))), "ending the show stops it");
});

await step("再生 → ブックマーク: add, name, remove; トリガー → ブックマーク時 plays the title as the sound passes it in the show", async () => {
  await page.locator(".film-item").nth(slideIndex).click();
  await page.waitForTimeout(300);
  const id = (await sound()).id;
  await pickObject(id);
  await page.waitForTimeout(200);
  await tab("再生");
  await ribbonBtn("ブックマーク…");
  await page.waitForSelector(".bm-dialog[open]");
  await page.waitForFunction(() => document.querySelector(".bm-dialog audio")?.readyState >= 1, null, { timeout: 8000 });
  await page.evaluate(() => { document.querySelector(".bm-dialog audio").currentTime = 1.4; });
  await page.waitForTimeout(200);
  await page.click(".bm-dialog .bm-add");
  await page.fill(".bm-dialog .bm-name", "サビ");
  await shot("bookmarks");
  await page.click(".bm-dialog .bm-ok");
  await page.waitForTimeout(300);
  let o = await sound();
  assert(o.bookmarks?.length === 1 && Math.abs(o.bookmarks[0].t - 1.4) < 0.05 && o.bookmarks[0].name === "サビ", `the bookmark: ${JSON.stringify(o.bookmarks)}`);
  // While the preview plays, ブックマークの追加 marks that moment; ブックマークの削除 takes one away.
  await ribbonBtn("再生");
  await page.waitForTimeout(500);
  await ribbonBtn("ブックマークの追加");
  await page.waitForTimeout(300);
  await ribbonBtn("再生");
  o = await sound();
  assert(o.bookmarks.length === 2, `a bookmark at the playing moment: ${JSON.stringify(o.bookmarks)}`);
  const extra = o.bookmarks.find((b) => b.name !== "サビ");
  await ribbonBtn("ブックマークの削除");
  await page.locator(`.rb-pop .rb-menu button:has-text("${extra.name}")`).first().click();
  await page.waitForTimeout(300);
  o = await sound();
  assert(o.bookmarks.length === 1 && o.bookmarks[0].name === "サビ", `removed: ${JSON.stringify(o.bookmarks)}`);
  // The title comes in when the sound reaches サビ (アニメーション → トリガー → ブックマーク時).
  await tab("アニメーション");
  if (!(await page.isVisible("#animPane:not([hidden])"))) await ribbonBtn("ウィンドウ");
  await page.waitForSelector("#animPane:not([hidden])");
  await page.click('#animPane button:has-text("＋ レイアウトの部品")');
  await page.locator('.rb-pop .rb-menu button:has-text("タイトル")').first().click();
  await page.waitForTimeout(400);
  const key = `${id}@${o.bookmarks[0].id}`;
  if (!(await page.locator('#animPane .an-details select[aria-label="トリガー"]').count())) await page.locator('#animPane .an-row:has-text("タイトル")').first().click();
  await page.waitForSelector('#animPane .an-details select[aria-label="トリガー"]', { timeout: 5000 });
  await page.locator('#animPane .an-details select[aria-label="トリガー"]').selectOption(key);
  await page.waitForTimeout(400);
  const entry = ((await deck()).slides[slideIndex].timeline || []).find((e) => e.el === "@title");
  assert(entry?.trigger === key, `the title waits for the bookmark: ${JSON.stringify(entry)}`);
  assert(/ブックマーク時：.*「サビ」/.test(await page.textContent("#animPane")), "the pane names the bookmark");
  await tab("スライド ショー");
  await ribbonBtn("このスライド");
  await page.waitForSelector("#presenter .hs-player-slide audio", { state: "attached", timeout: 10000 });
  assert(await page.locator("#presenter .hs-player-slide .hs-audio-track .hs-mark").count() === 1, "the bookmark's mark on the bar");
  const titleSeen = () => page.evaluate(() => getComputedStyle(document.querySelector('#presenter .hs-player-slide [data-field="title"]')).visibility);
  assert(await titleSeen() === "hidden", "the title waits for the bookmark");
  await page.evaluate(() => { const a = document.querySelector("#presenter .hs-player-slide audio"); if (a.paused) window.SlideEngine.mediaPlay(a, { fromStart: true }); });
  await page.waitForFunction(() => getComputedStyle(document.querySelector('#presenter .hs-player-slide [data-field="title"]')).visibility === "visible", null, { timeout: 6000 });
  for (let i = 0; i < 4 && await page.isVisible("#presenter"); i += 1) { await page.keyboard.press("Escape"); await page.waitForTimeout(400); }
  // Taking the bookmark away puts the title back in the click order.
  await page.locator(".film-item").nth(slideIndex).click();
  await page.waitForTimeout(300);
  await pickObject(id);
  await page.waitForSelector('.rb-tabs [role=tab]:text-is("再生")', { timeout: 5000 });
  await tab("再生");
  await ribbonBtn("ブックマークの削除");
  await page.locator('.rb-pop .rb-menu button:has-text("すべて削除")').first().click();
  await page.waitForTimeout(300);
  const after = ((await deck()).slides[slideIndex].timeline || []).find((e) => e.el === "@title");
  assert(after && !after.trigger && !(await sound()).bookmarks, `back to the click order: ${JSON.stringify(after)}`);
});

await step("挿入 → オーディオ → オーディオの録音: record, play back, insert", async () => {
  await page.locator(".film-item").nth(4).click();
  await page.waitForTimeout(300);
  await tab("挿入");
  await ribbonBtn("オーディオ");
  await page.click('.rb-pop .rb-menu button:has-text("オーディオの録音")');
  await page.waitForSelector(".rec-dialog[open]");
  await page.fill(".rec-dialog .rec-name", "ナレーション");
  await page.click(".rec-dialog .rec-start");
  await page.waitForFunction(() => document.querySelector(".rec-dialog")?.classList.contains("recording"), null, { timeout: 5000 });
  await page.waitForTimeout(1300);
  assert(/0:01/.test(await page.textContent(".rec-clock")), "the clock runs");
  await shot("recording");
  await page.click(".rec-dialog .rec-stop");
  await page.waitForFunction(() => !document.querySelector(".rec-dialog .rec-insert")?.disabled, null, { timeout: 5000 });
  assert(/録音しました/.test(await page.textContent(".rec-status")), "recorded");
  await page.click(".rec-dialog .rec-insert");
  await page.waitForTimeout(800);
  const o = (await deck()).slides[4].elements?.find((x) => x.kind === "audio");
  assert(o && o.fileName.startsWith("ナレーション.") && o.src.startsWith("idb:"), `the recording on the slide: ${JSON.stringify(o)}`);
});

await step("挿入 → 画面録画: record the screen, insert it as a video", async () => {
  await ribbonBtn("画面");
  await page.waitForSelector(".rec-dialog[open]");
  await page.click(".rec-dialog .rec-start");
  await page.waitForFunction(() => document.querySelector(".rec-dialog")?.classList.contains("recording") || /録画できません/.test(document.querySelector(".rec-status")?.textContent || ""), null, { timeout: 8000 });
  assert(!/録画できません/.test(await page.textContent(".rec-status")), await page.textContent(".rec-status"));
  await page.waitForTimeout(1200);
  await page.click(".rec-dialog .rec-stop");
  await page.waitForFunction(() => !document.querySelector(".rec-dialog .rec-insert")?.disabled, null, { timeout: 8000 });
  await page.click(".rec-dialog .rec-insert");
  await page.waitForTimeout(1000);
  const o = (await deck()).slides[4].elements?.find((x) => x.kind === "video");
  assert(o && o.fileName.startsWith("画面録画.") && o.src.startsWith("idb:"), `the screen recording: ${JSON.stringify(o)}`);
  assert(await page.locator('.rb-tabs [role=tab]:text-is("再生")').count(), "a video has the 再生 tab too");
  await tab("再生");
  assert(await (await ribbonField("全画面再生")).locator("input").isEnabled(), "全画面再生 for a video");
  assert(!(await (await ribbonField("再生中のアイコンを隠す")).locator("input").isEnabled()), "the icon setting is for sounds");
});

await step("再生 → 表紙画像: a frame of the video becomes its poster (shown before it plays); リセット takes it away", async () => {
  await tab("再生");
  await ribbonBtn("表紙画像");
  await page.locator('.rb-pop button:has-text("ビデオの1コマから")').first().click();
  await page.waitForSelector(".poster-dialog[open]");
  await page.waitForFunction(() => document.querySelector(".poster-dialog video")?.readyState >= 2, null, { timeout: 15000 });
  await page.click(".poster-dialog .poster-take");
  await page.waitForTimeout(300);
  const o = (await deck()).slides[4].elements.find((x) => x.kind === "video");
  assert(/^data:image\/jpeg;base64,/.test(o.poster || ""), `the poster: ${String(o.poster).slice(0, 40)}`);
  assert(await page.locator(`#stageBody .hs-obj[data-el="${o.id}"] video[poster^="data:image/jpeg"]`).count(), "the video shows it");
  await ribbonBtn("表紙画像");
  await page.locator('.rb-pop button:has-text("リセット")').first().click();
  await page.waitForTimeout(300);
  assert(!(await deck()).slides[4].elements.find((x) => x.kind === "video").poster, "reset");
});

await step("HTML出力: the sound travels inside the file, with its settings", async () => {
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#downloadBtn").then(async () => { await page.waitForTimeout(300); if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn"); })]);
  const file = join(outDir, "media-export.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  assert(/data:audio\/(wav|x-wav|wave);base64/.test(html), "the WAV inside the file");
  assert(/"kind":"audio"/.test(html) && /"across":999/.test(html) && /"trimStart":0.3/.test(html), "and its settings");
  // The exported file plays it like the studio does.
  const view = await context.newPage();
  await view.goto(`file://${file}`);
  await view.waitForSelector(".hs-slide");
  assert(await view.evaluate(() => typeof window.SlideEngine?.mediaPlay === "function"), "the player knows sounds");
  await view.close();
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
