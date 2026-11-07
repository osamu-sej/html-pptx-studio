// 画面切り替えの効果のオプション and コーチによるリハーサル in a real browser: 押し上げ set to come from the right (the
// preview and the show move the slide sideways), the option dropped when the transition has none, すべてに適用 carrying
// it; and a rehearsal with the coach (the browser's speech recognition stood in for): the live pace and filler count,
// then the report with the fillers, the word said twice and the slide whose text was read out.
// Usage: node qa/studio-rehearse.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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
// The speech recognition stands in for a microphone: window.__say("…") hands it a finished sentence.
await context.addInitScript(() => {
  class FakeRecognition {
    constructor() { window.__rec = this; this.running = false; }
    start() { this.running = true; }
    stop() { this.running = false; }
  }
  window.SpeechRecognition = FakeRecognition;
  window.__say = (text) => {
    const result = [{ transcript: text }];
    result.isFinal = true;
    window.__rec?.onresult?.({ resultIndex: 0, results: [result] });
  };
});
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name) => { const file = join(outDir, `rehearse-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible("#presenter .hs-player")) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
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
const menuItem = (label) => page.locator(`.rb-pop button:has-text("${label}")`).first().click();

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.click("#sampleDeckBtn");
await page.waitForSelector(".film-item");
await page.waitForSelector("#ribbon:not([hidden]) .rb-tabs");

await step("効果のオプション: 押し上げ from the right; the preview moves the slide sideways", async () => {
  await page.locator(".film-item").nth(2).click();
  await tab("画面切り替え");
  await page.click('.an-tr[data-tr="push"]');
  await page.waitForTimeout(1400);
  await ribbonBtn("効果の");
  const options = await page.locator(".rb-pop .rb-menu button").allTextContents();
  assert(["下から", "上から", "右から", "左から"].every((o) => options.some((t) => t.includes(o))), `options: ${options}`);
  await menuItem("右から");
  await page.waitForSelector(".transition-preview .tr-dir", { timeout: 3000 });
  const dx = await page.evaluate(() => getComputedStyle(document.querySelector(".transition-preview .tr-dir")).getPropertyValue("--tr-dx").trim());
  assert(dx === "100%", `from the right: ${dx}`);
  assert((await deck()).slides[2].transitionDir === "right", "the slide keeps the direction");
  await page.waitForTimeout(1200);
});

await step("the show: arriving at the slide, the push comes from the right", async () => {
  await page.locator(".film-item").nth(1).click();
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector("#presenter .hs-player");
  await page.waitForTimeout(600);
  for (let k = 0; k < 12 && !(await page.locator("#presenter .hs-tr-in-push.tr-dir").count()); k += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(120); }
  assert(await page.locator("#presenter .hs-tr-in-push.tr-dir").count(), "a sideways push");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

await step("すべてに適用 carries the direction; a transition without options drops it", async () => {
  await page.locator(".film-item").nth(2).click();
  await tab("画面切り替え");
  await ribbonBtn("すべてに適用");
  await page.waitForTimeout(300);
  assert((await deck()).slides.every((s) => s.transition === "push" && s.transitionDir === "right"), "every slide");
  await page.click('.an-tr[data-tr="fade"]');
  await page.waitForTimeout(1200);
  const s = (await deck()).slides[2];
  assert(s.transition === "fade" && !s.transitionDir, `fade has no direction: ${JSON.stringify([s.transition, s.transitionDir])}`);
  assert(await page.locator('.rb-body .rb-btn:has-text("効果の")').first().isDisabled(), "効果のオプション is off for fade");
});

await step("コーチによるリハーサル: the live meter, then the report (fillers, a repeat, a slide read out)", async () => {
  await tab("スライド ショー");
  await ribbonBtn("コーチによる");
  await page.waitForSelector("#presenter .coach-meter");
  await page.evaluate(() => window.__say("えーと、本日は、えー、生成AIの、あのー、その その成果をご報告します"));
  await page.waitForTimeout(500);
  assert(/つなぎ言葉 3/.test(await page.textContent(".coach-meter")), `live count: ${await page.textContent(".coach-meter")}`);
  await shot("meter");
  await page.waitForTimeout(5200);
  // The next slide: its title and message read out word for word.
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(400);
  const index = await page.evaluate(() => [...document.querySelectorAll("#presenter .hs-player-count")].pop()?.textContent);
  const n = Number(index.split("/")[0]) - 1;
  const words = await page.evaluate((i) => { const s = window.__hpptx.deck().slides[i]; return [s.title, s.takeaway, s.subhead, s.message, s.text].filter(Boolean).join("、").replace(/<[^>]*>|\*\*/g, ""); }, n);
  await page.evaluate((w) => window.__say(w), words);
  await page.waitForTimeout(5200);
  await page.keyboard.press("Escape");
  await page.waitForSelector(".coach-dialog[open]", { timeout: 5000 });
  const text = await page.textContent(".coach-dialog");
  assert(/つなぎ言葉/.test(text) && /3 回/.test(text), "three fillers in the report");
  assert(/その/.test(text), "the word said twice");
  assert(new RegExp(`${n + 1}枚目は、スライドの文字をそのまま読み上げて`).test(text), `slide ${n + 1} read out: ${text.slice(0, 300)}`);
  assert((await page.locator(".coach-slides tbody tr").count()) >= 2, "a row a slide");
  await shot("report");
  await page.keyboard.press("Escape");
  assert(await page.evaluate(() => !window.__rec.running), "listening stops with the show");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
