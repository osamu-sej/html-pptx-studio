// The AI side of the studio with the server's AI answers stubbed in the browser (no Codex login needed):
// outline → generation with live slides (in the theme chosen on the create view) → chat proposal → three variants.
// Usage: node qa/studio-ai-mock.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
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

const sample = JSON.parse(await readFile(join(root, "public", "samples", "ai-rollout.json"), "utf8"));
const first = (type) => sample.slideData.find((slide) => slide.type === type);
const pick = ["title", "executiveSummary", "kpi", "waterfall", "imageText", "process", "roadmap", "statement", "closing"].map(first);
const generated = { deckTitle: sample.deckTitle, purpose: sample.purpose, audience: sample.audience, slideData: pick };
const outline = {
  deckTitle: sample.deckTitle,
  reply: "結論を2枚目に置き、数字→内訳→展開計画の順にしました。",
  slides: pick.map((slide) => ({ type: slide.type, title: slide.title || "次のアクション", takeaway: slide.takeaway || "", content: "" })),
};
const revisedSlides = pick.map((slide, i) => (i === 2 ? { ...slide, title: "削減効果（簡潔版）", takeaway: "1人あたり**月12時間**を取り戻した" } : slide));
const chat = {
  reply: "3枚目の見出しとキーメッセージを短くしました。",
  suggestions: ["工程をクリックで1つずつ出して", "数字を目立たせて"],
  // An older server could still send a theme; the theme stays the user's choice anyway.
  deckTitle: null, theme: "midnight", transition: null, changed: true, summary: "1枚を修正",
  slides: revisedSlides, items: revisedSlides.map((_, i) => ({ from: i, changed: i === 2 })), deleted: [], moved: false, issues: [],
};
const variants = [
  { label: "数字を主役に", slide: { ...pick[2], title: "削減効果", takeaway: "**1,440時間**を取り戻した" }, issues: [], unsourced: [] },
  { label: "ひと言で", slide: { type: "statement", title: "削減効果", text: "毎月**1,440時間**が、考える時間に変わった", takeaway: "5部門120人・3か月の試行結果" }, issues: [], unsourced: [] },
  { label: "図解で", slide: { type: "cards", title: "削減効果", takeaway: "3つの業務で大きく減った", items: [{ title: "資料作成", desc: "月620時間", icon: "document" }, { title: "議事録", desc: "月380時間", icon: "mic" }, { title: "データ集計", desc: "月290時間", icon: "chartBar" }] }, issues: [], unsourced: [] },
];

const jobs = new Map();
const imageRequests = [];
let next = 1;
const newJob = (result, partials = []) => {
  const id = `00000000-0000-4000-8000-${String(next++).padStart(12, "0")}`;
  jobs.set(id, { id, status: "completed", stage: "完了", detail: "できました。", ...result, partials });
  return id;
};

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const page = await browser.newPage({ viewport: { width: 1600, height: 1000 } });
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });

await page.route("**/api/codex/status", (route) => route.fulfill({ json: { available: true, authorized: true, authenticated: true, planType: "test", team: false, fallback: { available: false, model: "gemma4:12b", web: false, reason: "" } } }));
await page.route("**/api/decks/outline", (route) => route.fulfill({ status: 202, json: { jobId: newJob({ kind: "outline", outline }) } }));
await page.route(/\/api\/decks$/, (route) => route.fulfill({ status: 202, json: { jobId: newJob({ kind: "deck", deck: generated, detail: `${pick.length}枚の構成が完成しました。` }, [3, 6]) } }));
await page.route("**/api/decks/image", async (route) => {
  const request = route.request().postDataJSON();
  imageRequests.push(request.index);
  const jobId = newJob({ kind: "image", image: { url: "/api/decks/mock/image", mime: "image/jpeg" } });
  await route.fulfill({ status: 202, json: { jobId } });
});
await page.route("**/api/decks/mock/image", async (route) => route.fulfill({ status: 200, contentType: "image/jpeg", body: await readFile(join(root, "public", "assets", "data-insight.jpg")) }));
await page.route("**/api/decks/chat", (route) => route.fulfill({ status: 202, json: { jobId: newJob({ kind: "chat", chat }) } }));
await page.route("**/api/decks/variants", (route) => route.fulfill({ status: 202, json: { jobId: newJob({ kind: "variants", variants }) } }));
await page.route("**/api/decks/revise", (route) => route.fulfill({ status: 503, json: { error: "テストでは作り直しません" } }));
await page.route(/\/api\/decks\/[0-9a-f-]{36}(\/events)?$/, (route) => {
  const url = new URL(route.request().url());
  const id = url.pathname.split("/")[3];
  const job = jobs.get(id);
  if (!job) return route.fulfill({ status: 404, json: { error: "Not found" } });
  const { partials, ...done } = job;
  if (!url.pathname.endsWith("/events")) return route.fulfill({ json: done });
  const events = [
    ...partials.map((n) => ({ ...done, status: "running", stage: "書き出し中", detail: `${n}枚目まで書きました`, partial: generated.slideData.slice(0, n) })),
    done,
  ].map((event) => `event: progress\ndata: ${JSON.stringify(event)}\n\n`).join("");
  return route.fulfill({ status: 200, headers: { "content-type": "text/event-stream" }, body: events });
});

const shot = async (name) => { const file = join(outDir, `ai-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);

await step("outline", async () => {
  await page.fill("#briefInput", "生成AIの試行結果を役員に報告し、全社展開の予算承認をもらいたい。");
  await page.click("#askChatGptBtn");
  await page.waitForSelector("#outlinePanel:not(.hidden) .outline-item", { timeout: 10000 });
  // The mock outline puts a process and a roadmap side by side: both are lines across the slide.
  const note = await page.textContent("#outlineVariety:not(.hidden)", { timeout: 3000 });
  if (!/6枚目と7枚目が、どちらも「横に並ぶ流れ」/.test(note)) throw new Error(`variety note: ${note}`);
  const groups = await page.$eval(".outline-item select", (select) => [...select.querySelectorAll("optgroup")].map((group) => group.label));
  if (!groups.includes("横に並ぶ流れ") || !groups.includes("図形")) throw new Error(`layout groups: ${groups}`);
  await shot("outline");
  // Choosing a different look for the roadmap clears the note.
  await page.selectOption(".outline-item:nth-child(7) select", "checklist");
  await page.waitForSelector("#outlineVariety.hidden", { state: "attached", timeout: 3000 });
  await page.selectOption(".outline-item:nth-child(7) select", "roadmap");
});

await step("generate from the outline", async () => {
  await page.click("#outlineBuildBtn");
  await page.waitForFunction((n) => document.querySelectorAll(".film-item").length === n, pick.length, { timeout: 15000 });
  await page.waitForTimeout(1500);
  const theme = await page.getAttribute(".slide-wrap .hs-slide", "data-theme");
  if (theme !== "clarity") throw new Error(`theme was ${theme}`);
  await page.waitForFunction(() => document.querySelector("#autoImageStatus")?.textContent?.includes("画像 3/9枚"), null, { timeout: 15000 });
  if (imageRequests.length !== 3) throw new Error(`automatic images: ${imageRequests.length} instead of 3`);
  await shot("generated");
  await page.click(".film-item:nth-child(2)");
  await shot("generated-evidence");
  await page.click(".film-item:nth-child(3)");
  await shot("generated-kpi");
  await page.click(".film-item:nth-child(1)");
  await page.reload();
  await page.waitForFunction(() => document.querySelector("#autoImageStatus")?.textContent?.includes("画像 3/9枚"), null, { timeout: 10000 });
  const stored = await page.evaluate(() => JSON.parse(localStorage.getItem("hpptx-studio-current-v1"))?.deck?.slides?.filter((slide) => slide.media?.kind === "image").length);
  if (stored !== 3) throw new Error(`images not saved: ${stored}`);
});

await step("chat proposal: text, and the theme stays the user's", async () => {
  await page.click("#chatTab");
  await page.fill("#chatInput", "@3 をもっと簡潔に。テーマも画面共有向けに");
  await page.keyboard.press("Enter");
  await page.waitForSelector(".proposal .proposal-actions", { timeout: 10000 });
  await page.waitForTimeout(800);
  await shot("proposal");
  await page.click('.proposal-actions button:has-text("選んだ変更を採用")');
  await page.waitForTimeout(900);
  const title = await page.textContent(".slide-wrap .hs-title");
  const theme = await page.getAttribute(".slide-wrap .hs-slide", "data-theme");
  if (!title.includes("簡潔版") || theme !== "clarity") throw new Error(`title=${title} theme=${theme}`);
  await shot("applied");
});

await step("three variants", async () => {
  await page.click("#formTab");
  await page.click('.inspector-ai button:has-text("別案を3つ")');
  await page.waitForSelector(".variant-grid .variant", { timeout: 10000 });
  await page.waitForTimeout(900);
  await shot("variants");
  await page.click('.variant:has-text("ひと言で") button:has-text("この案にする")');
  await page.waitForTimeout(700);
  const type = await page.getAttribute(".slide-wrap .hs-slide", "data-type");
  if (type !== "statement") throw new Error(`type=${type}`);
  await shot("variant-chosen");
});

await writeFile(join(outDir, "ai-errors.txt"), errors.join("\n"));
console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
