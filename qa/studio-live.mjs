// PowerPoint Live in real browsers: a poll (アンケート) inserted and edited; presented online with two viewers — the
// QR code, the viewers' answers filling the presenter's bars (a changed answer replacing the old one), reactions
// floating up on the presenter's screen, the presenter's captions on the viewers' screens, the join card (Q); and
// the same poll answered by clicks in an ordinary show.
// Usage: node qa/studio-live.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
const errors = [];
async function open(name) {
  const context = await browser.newContext({ viewport: { width: 1600, height: 1000 } });
  // The presenter's microphone is stood in for: window.__say("…") hands the recognition a finished sentence.
  await context.addInitScript(() => {
    class FakeRecognition { constructor() { window.__rec = this; } start() {} stop() {} }
    window.SpeechRecognition = FakeRecognition;
    window.__say = (text) => { const result = [{ transcript: text }]; result.isFinal = true; window.__rec?.onresult?.({ resultIndex: 0, results: [result] }); };
  });
  const p = await context.newPage();
  p.on("pageerror", (error) => errors.push(`${name} pageerror: ${error.message}`));
  p.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g|ERR_ABORTED|EventSource/.test(message.text())) errors.push(`${name} console: ${message.text()}`); });
  p.on("dialog", (dialog) => dialog.accept());
  return p;
}
const page = await open("presenter");
const shot = async (name, p = page) => { const file = join(outDir, `live-${name}.png`); await p.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
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
const counts = (p) => p.evaluate(() => [...document.querySelectorAll("#presenter .hs-poll .hs-poll-count")].map((el) => Number(el.textContent)));

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

await step("挿入 → アンケート: a question and choices (one per line), edited again by a double-click", async () => {
  await tab("挿入");
  await ribbonBtn("アンケート");
  await page.waitForSelector(".poll-dialog[open]");
  await page.fill(".poll-dialog .pl-question", "来期はどこから始めますか？");
  await page.fill(".poll-dialog .pl-options", "・店舗から\n・本部から\n・同時に");
  await page.click(".poll-dialog .pl-ok");
  const [poll] = (await objects()).filter((o) => o.kind === "poll");
  assert(poll && poll.question === "来期はどこから始めますか？" && poll.options.join("|") === "店舗から|本部から|同時に", `poll: ${JSON.stringify(poll)}`);
  // The stage draws the new poll a moment after the deck has it: wait for it before reading where it is.
  await page.waitForSelector(`#stageBody .hs-obj[data-el="${poll.id}"]`);
  const b = await page.locator(`#stageBody .hs-obj[data-el="${poll.id}"]`).boundingBox();
  await page.mouse.dblclick(b.x + b.width / 2, b.y + b.height / 2);
  await page.waitForSelector(".poll-dialog[open]");
  await page.fill(".poll-dialog .pl-options", "店舗から\n本部から");
  await page.click(".poll-dialog .pl-ok");
  assert((await objects()).find((o) => o.id === poll.id).options.length === 2, "two choices now");
  await shot("poll-edit");
});

const viewers = [await open("viewer1"), await open("viewer2")];
let pollSlide = 0;
await step("online: the QR code; two viewers answer the poll and the presenter's bars fill (a changed answer replaces the old)", async () => {
  pollSlide = await page.evaluate(() => window.__hpptx.deck().slides.findIndex((s) => (s.elements || []).some((o) => o.kind === "poll")));
  await tab("スライド ショー");
  await ribbonBtn("オンライン");
  await page.waitForSelector(".online-dialog[open]");
  await page.click(".online-dialog .op-start");
  await page.waitForSelector(".online-dialog .op-link", { timeout: 90000 });
  assert(await page.locator(".online-dialog .op-qr svg").count(), "a QR code");
  const link = await page.inputValue(".online-dialog .op-link");
  for (const v of viewers) { await v.goto(link); await v.waitForSelector("#presenter .hs-player", { timeout: 20000 }); }
  await page.click(".online-dialog .op-start");
  await page.waitForSelector("#presenter .hs-player");
  // The show starts on the slide shown (the poll's).
  for (let k = 0; k < 60 && !(await page.locator("#presenter .hs-poll").count()); k += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(150); }
  for (const v of viewers) await v.waitForSelector("#presenter .hs-poll .hs-poll-opt", { timeout: 10000 });
  await viewers[0].click("#presenter .hs-poll-opt >> nth=0");
  await viewers[1].click("#presenter .hs-poll-opt >> nth=1");
  await page.waitForFunction(() => [...document.querySelectorAll("#presenter .hs-poll .hs-poll-count")].map((el) => el.textContent).join(",") === "1,1", null, { timeout: 8000 });
  await viewers[0].click("#presenter .hs-poll-opt >> nth=1");
  await page.waitForFunction(() => [...document.querySelectorAll("#presenter .hs-poll .hs-poll-count")].map((el) => el.textContent).join(",") === "0,2", null, { timeout: 8000 });
  assert(JSON.stringify(await counts(viewers[1])) === "[0,2]", "the viewers see the results too");
  assert(await viewers[0].locator("#presenter .hs-poll-opt.is-mine").count() === 1, "the viewer's own answer is marked");
  await shot("poll-results");
  await shot("poll-viewer", viewers[0]);
});

await step("reactions float up on the presenter's screen; the presenter's captions reach the viewers; Q shows the join card", async () => {
  // The viewer's bar comes back when the pointer moves (or the screen is tapped).
  await viewers[0].mouse.move(800, 500);
  await viewers[0].mouse.move(800, 940);
  await viewers[0].click("#presenter .hs-react-btn >> nth=2");
  await page.waitForSelector("#presenter .hs-react", { timeout: 6000 });
  assert((await page.textContent("#presenter .hs-react")).includes("👏"), "a clap");
  await page.keyboard.press("j");
  await page.waitForTimeout(300);
  await page.evaluate(() => window.__say("来期の始め方をみなさんに聞きます"));
  await viewers[1].waitForFunction(() => /来期の始め方/.test(document.querySelector("#presenter .hs-captions")?.textContent || ""), null, { timeout: 8000 });
  await page.keyboard.press("q");
  await page.waitForSelector("#presenter .hs-join:not([hidden]) svg", { timeout: 3000 });
  await shot("join-card");
  await page.keyboard.press("q");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
  if (!(await page.isVisible(".online-dialog[open]"))) { await tab("スライド ショー"); await ribbonBtn("オンライン"); }
  assert(/アンケートの回答 2件/.test(await page.textContent(".online-dialog .op-status")), "the dialog counts the answers");
  await page.click(".online-dialog .op-end");
});

await step("an ordinary show: clicks on the choices count the answers here", async () => {
  await page.locator(".film-item").nth(pollSlide).click();
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector("#presenter .hs-poll .hs-poll-opt");
  await page.click("#presenter .hs-poll-opt >> nth=0");
  await page.click("#presenter .hs-poll-opt >> nth=0");
  await page.click("#presenter .hs-poll-opt >> nth=1");
  assert(JSON.stringify(await counts(page)) === "[2,1]", `counted: ${JSON.stringify(await counts(page))}`);
  await page.keyboard.press("Escape");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
