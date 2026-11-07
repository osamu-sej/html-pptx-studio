// 共同編集 in two real browsers at once: sharing makes a link; the other person joins with it; each sees the other
// (the share button, a mark on the thumbnail of the slide they are on, an outline around what they have selected);
// changes cross over (a text box, a shape on the same slide at the same time, a comment with its writer, the order of
// the slides, the deck's title, a sound with its file); ⌘Z takes back only one's own change; leaving is seen.
// Usage: node qa/studio-coedit.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir, writeFile } from "node:fs/promises";
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
async function person(name) {
  const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, permissions: ["clipboard-read", "clipboard-write"] });
  const page = await context.newPage();
  page.on("pageerror", (error) => errors.push(`${name} pageerror: ${error.message}`));
  page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g|ERR_ABORTED|EventSource/.test(`${message.text()} ${message.location().url}`)) errors.push(`${name} console: ${message.text()} ${message.location().url}`.trim()); });
  page.on("dialog", (dialog) => dialog.accept());
  await page.goto(base);
  await page.evaluate((n) => { localStorage.clear(); localStorage.setItem("hpptx-user-name", n); }, name);
  return { page, context, name };
}
const step = async (label, fn) => {
  try { await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); }
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const deck = (p) => p.page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
const until = async (p, fn, arg, what, timeout = 8000) => {
  try { await p.page.waitForFunction(fn, arg, { timeout, polling: 150 }); } catch { throw new Error(`${p.name}: ${what}`); }
};
const tab = (p, label) => p.page.click(`.rb-tabs [role=tab]:text-is("${label}")`);
const ribbonBtn = async (p, label) => {
  const direct = p.page.locator(`.rb-body .rb-btn:has-text("${label}")`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct.click();
  for (const folded of await p.page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = p.page.locator(`.rb-pop .rb-fold .rb-btn:has-text("${label}")`).first();
    if (await inside.count()) return inside.click();
  }
  throw new Error(`no ribbon button "${label}"`);
};
async function drawBox(p, from, to) {
  const box = await p.page.locator(".slide-wrap .hs-slide").boundingBox();
  const at = (x, y) => [box.x + (x * box.width) / 1920, box.y + (y * box.height) / 1080];
  await p.page.mouse.move(...at(...from));
  await p.page.mouse.down();
  await p.page.mouse.move(...at(...to), { steps: 6 });
  await p.page.mouse.up();
  await p.page.waitForTimeout(250);
}
const film = (p, i) => p.page.locator(".film-item").nth(i);

const a = await person("山田");
const b = await person("佐藤");
let link = "";

await step("共有: the deck becomes a room; the link is shown and copied", async () => {
  await a.page.click("#sampleDeckBtn");
  await a.page.waitForSelector(".slide-wrap .hs-slide");
  await a.page.click("#shareBtn");
  await a.page.waitForSelector("#shareDialog[open]");
  link = await a.page.inputValue("#shareLink");
  assert(/[?&]room=[A-Za-z0-9_-]{20,}/.test(link), `a link: ${link}`);
  await a.page.click("#shareDialog [data-close]");
  await until(a, () => /(^|[^0-9])1人$/.test(document.getElementById("shareBtn").textContent), null, "shared");
  const d = await deck(a);
  assert(d.slides.every((s) => s.sid), "every slide has its id");
});

await step("the other person opens the link: the same deck; both see two people", async () => {
  await b.page.goto(link);
  await b.page.waitForSelector(".slide-wrap .hs-slide");
  await until(b, () => /(^|[^0-9])2人$/.test(document.getElementById("shareBtn").textContent), null, "B sees 2");
  await until(a, () => /(^|[^0-9])2人$/.test(document.getElementById("shareBtn").textContent), null, "A sees 2");
  const [da, db] = [await deck(a), await deck(b)];
  assert(db.slides.length === da.slides.length && db.title === da.title, "the same deck");
  assert(await a.page.locator('#shareBtn .pp-avatar[title^="佐藤"]').count(), "A sees 佐藤's avatar");
});

await step("presence: the thumbnail of the slide the other is on; the outline of what they selected", async () => {
  await film(b, 2).click();
  await until(a, () => document.querySelectorAll('.film-item[data-index="2"] .co-film .pp-avatar').length > 0, null, "A sees B on slide 3");
  await film(a, 1).click();
  await film(b, 1).click();
  await tab(a, "挿入");
  await ribbonBtn(a, "テキスト ボックス");
  await a.page.click('.rb-pop .rb-menu button:has-text("横書き")');
  await drawBox(a, [300, 700], [800, 800]);
  await a.page.keyboard.type("山田のメモ");
  await a.page.keyboard.press("Escape");
  await until(b, () => (window.__hpptx.deck().slides[1].elements || []).some((o) => /山田のメモ/.test(o.text || "")), null, "B gets A's text box");
  await until(b, () => document.querySelector(".co-sel .co-sel-name")?.textContent === "山田", null, "B sees A's selection outline");
  await b.page.screenshot({ path: join(outDir, "coedit-presence.png") });
});

await step("both on one slide at once: A's text and B's shape both survive", async () => {
  await tab(b, "挿入");
  await b.page.locator('.rb-btn.big:has-text("図形")').first().click();
  await b.page.waitForSelector(".rb-gallery.shapes");
  await b.page.click('.rb-gallery button[title="楕円"]');
  // A types into its box while B draws.
  const typing = (async () => {
    const d = await deck(a);
    const id = d.slides[1].elements.find((o) => /山田のメモ/.test(o.text || "")).id;
    const node = await a.page.locator(`.slide-wrap .hs-obj[data-el="${id}"]`).boundingBox();
    await a.page.mouse.dblclick(node.x + node.width / 2, node.y + node.height / 2);
    await a.page.keyboard.press("End");
    await a.page.keyboard.type("（追記）");
    await a.page.keyboard.press("Escape");
  })();
  await drawBox(b, [1100, 600], [1400, 850]);
  await typing;
  for (const p of [a, b]) {
    await until(p, () => { const els = window.__hpptx.deck().slides[1].elements || []; return els.some((o) => o.shape === "ellipse") && els.some((o) => /（追記）/.test(o.text || "")); }, null, "both changes on both sides", 10000);
  }
});

await step("⌘Z takes back only one's own change", async () => {
  await b.page.click("#stageBody", { position: { x: 5, y: 5 } });
  await b.page.keyboard.press("Control+z");
  for (const p of [a, b]) await until(p, () => { const els = window.__hpptx.deck().slides[1].elements || []; return !els.some((o) => o.shape === "ellipse") && els.some((o) => /（追記）/.test(o.text || "")); }, null, "B's shape gone, A's text kept");
});

await step("a comment with its writer; the deck's title; the order of the slides", async () => {
  await tab(b, "校閲");
  await ribbonBtn(b, "新しい");
  await b.page.waitForSelector("#commentPane:not([hidden]) .cm-input");
  await b.page.fill("#commentPane .cm-input", "@山田 この数字の出典は？");
  await b.page.click("#commentPane .cm-post");
  await until(a, () => (window.__hpptx.deck().slides[1].comments || []).some((c) => c.by === "佐藤" && /出典/.test(c.text)), null, "A gets B's comment with the name");
  await a.page.fill("#deckTitleInput", "共同編集のテスト資料");
  await a.page.press("#deckTitleInput", "Enter");
  await until(b, () => window.__hpptx.deck().title === "共同編集のテスト資料", null, "B gets the title");
  const before = (await deck(a)).slides.map((s) => s.sid);
  await film(a, 3).click({ button: "right" });
  await a.page.click('.ed-menu button:has-text("上へ移動")');
  await a.page.waitForTimeout(400);
  const after = (await deck(a)).slides.map((s) => s.sid);
  assert(after.join() !== before.join(), `A moved a slide: ${after.slice(0, 5)}`);
  await until(b, (sids) => window.__hpptx.deck().slides.map((s) => s.sid).join() === sids, after.join(), "B gets the new order");
});

await step("a sound inserted here plays for the other person too (the file goes through the room)", async () => {
  const rate = 8000;
  const n = rate;
  const wav = Buffer.alloc(44 + n * 2);
  wav.write("RIFF", 0); wav.writeUInt32LE(36 + n * 2, 4); wav.write("WAVE", 8); wav.write("fmt ", 12); wav.writeUInt32LE(16, 16); wav.writeUInt16LE(1, 20); wav.writeUInt16LE(1, 22); wav.writeUInt32LE(rate, 24); wav.writeUInt32LE(rate * 2, 28); wav.writeUInt16LE(2, 32); wav.writeUInt16LE(16, 34); wav.write("data", 36); wav.writeUInt32LE(n * 2, 40);
  for (let i = 0; i < n; i += 1) wav.writeInt16LE(Math.round(Math.sin(i / 3) * 6000), 44 + i * 2);
  const file = join(outDir, "coedit-tone.wav");
  await writeFile(file, wav);
  await film(a, 1).click();
  await tab(a, "挿入");
  await ribbonBtn(a, "オーディオ");
  const [chooser] = await Promise.all([a.page.waitForEvent("filechooser"), a.page.click('.rb-pop .rb-menu button:has-text("このデバイスのオーディオ")')]);
  await chooser.setFiles(file);
  await until(b, () => (window.__hpptx.deck().slides[1].elements || []).some((o) => o.kind === "audio"), null, "B gets the sound object");
  await film(b, 1).click();
  await tab(b, "スライド ショー");
  await ribbonBtn(b, "このスライド");
  await b.page.waitForSelector("#presenter .hs-player-slide audio", { state: "attached", timeout: 10000 });
  const src = await b.page.getAttribute("#presenter .hs-player-slide audio", "src");
  assert(/^blob:/.test(src || ""), `B has the file itself: ${src}`);
  await b.page.keyboard.press("Escape");
  await b.page.waitForTimeout(300);
});

await step("leaving: the other side sees one person again", async () => {
  await b.page.click("#shareBtn");
  await b.page.click('.rb-pop .rb-menu button:has-text("共同編集から抜ける")');
  await until(a, () => /(^|[^0-9])1人$/.test(document.getElementById("shareBtn").textContent), null, "A sees B left");
  assert(!/room=/.test(b.page.url()), "B's address no longer has the room");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
