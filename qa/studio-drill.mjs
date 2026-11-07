// Deep-dive pages in a real browser: make one from an item, find it in the filmstrip, jump to it while
// presenting and come back to the same step, keep it with its slide when moving and deleting, and use it in
// the exported file.
// Usage: node qa/studio-drill.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
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
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
const shot = async (name, target = page) => { const file = join(outDir, `drill-${name}.png`); await target.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };
const films = () => page.$$eval(".film-item", (items) => items.map((item) => (item.classList.contains("is-drill") ? "↳" : item.querySelector(".hs-slide")?.dataset.type)));
const expect = (value, wanted, what) => { if (JSON.stringify(value) !== JSON.stringify(wanted)) throw new Error(`${what}: ${JSON.stringify(value)}`); };

// The slide from the request: four stages, where "02" should open a page that digs into it.
const deck = {
  title: "生成AIの歴史・選び方・将来", audience: "社内", purpose: "研修", theme: "kinari", transition: "fade", motion: {}, memo: "",
  slides: [
    { type: "title", title: "生成AIの歴史・選び方・将来" },
    { type: "stepUp", title: "生成AIへ至る発展段階", takeaway: "AIは**規則の実行から内容の生成へ**、担う役割を広げてきた。", items: [
      { title: "規則による処理", desc: "人が定めた条件に沿い、決められた判断や処理を実行する。" },
      { title: "データから学習", desc: "蓄積された例から傾向を捉え、分類や予測に活用する。" },
      { title: "複雑な特徴の把握", desc: "大量の情報から特徴を捉え、認識や理解の範囲を広げる。" },
      { title: "内容の生成", desc: "指示と文脈を基に、文章や画像などの新しい出力を組み立てる。" },
    ] },
    { type: "cards", title: "選び方の3つの軸", takeaway: "用途・品質・運用で比べる", items: [{ title: "用途", desc: "何に使うか" }, { title: "品質", desc: "正確さ" }, { title: "運用", desc: "安全に使えるか" }] },
    { type: "statement", title: "これから", text: "まず**小さく試す**" },
    { type: "closing", title: "次のアクション", message: "来週までに1つの業務で試す" },
  ],
};

await page.goto(base);
await page.evaluate((value) => localStorage.setItem("hpptx-studio-current-v1", JSON.stringify({ deck: value, selected: 1, savedAt: new Date().toISOString() })), deck);
await page.goto(base);
await page.waitForSelector(".film-item");
await page.waitForTimeout(800);

await step("make a deep-dive page for item 02 from the inspector", async () => {
  await page.click("#formTab");
  await page.click(".film-item:nth-child(2)");
  await page.waitForSelector("#drillSection .drill-add select");
  await page.selectOption("#drillSection .drill-add select", "items[1]");
  await page.click('#drillSection button:has-text("白紙で作る")');
  await page.waitForSelector(".film-item.is-drill");
  expect(await films(), ["title", "stepUp", "↳", "cards", "statement", "closing"], "filmstrip");
  await page.waitForSelector("#inspector .drill-info");
  const caption = await page.textContent("#slidePos");
  if (!/2枚目の深掘りページ/.test(caption)) throw new Error(`status bar: ${caption}`);
  // The page number counts the story (like ‹#›): the deep-dive page keeps its slide's number, says it is a deep
  // dive, and names that slide in the line above its key message.
  const pageNo = await page.textContent(".slide-wrap .hs-page");
  if (!/^02 ・ 深掘り$/.test(pageNo.trim())) throw new Error(`page number on the deep-dive page: ${pageNo}`);
  const eyebrow = await page.textContent(".slide-wrap .hs-drill .hs-eyebrow");
  if (!/^↳ /.test(eyebrow)) throw new Error(`the deep-dive page should name its slide: ${eyebrow}`);
  const meta = await page.textContent("#deckMeta");
  if (!/5枚＋深掘り1枚/.test(meta)) throw new Error(`deck meta: ${meta}`);
  await page.fill("#f-title", "データから学習する仕組み");
  await page.waitForTimeout(300);
  await shot("studio-page");
});

await step("the slide shows an arrow on 02, and the arrow opens the page in the editor", async () => {
  await page.click('#inspector button:has-text("元のスライドへ")');
  await page.waitForSelector('.slide-wrap .hs-slide[data-type="stepUp"] .hs-drill-badge');
  const badges = await page.$$eval('.slide-wrap [data-drill]', (els) => [...new Set(els.map((el) => el.dataset.item))]);
  expect(badges, ["items[1]"], "items with a deep-dive page");
  const numbers = await page.$$eval(".film-item", (items) => items.map((item) => item.querySelector(".hs-page")?.textContent ?? ""));
  expect(numbers.filter(Boolean), ["02 / 05", "02 ・ 深掘り", "03 / 05", "04 / 05"], "page numbers count the story only");
  await shot("studio-slide");
  await page.click(".slide-wrap .hs-drill-badge");
  await page.waitForSelector(".film-item.is-drill.selected");
  await page.click('#inspector button:has-text("元のスライドへ")');
});

await step("present: click 02, see the page, Esc comes back to the same step", async () => {
  await page.click("#presentBtn");
  await page.waitForSelector('#presenter .hs-player .hs-slide[data-type="stepUp"]');
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(500);
  await page.locator('#presenter .hs-player-slide [data-drill]:not(svg *)').first().click();
  await page.waitForSelector("#presenter .hs-player .hs-slide.hs-drill", { timeout: 5000 });
  await page.waitForTimeout(800);
  const counter = await page.textContent("#presenter .hs-player-count");
  if (!/2 \/ 5 ・ 深掘り/.test(counter)) throw new Error(`counter: ${counter}`);
  if (!(await page.isVisible("#presenter .hs-player-back"))) throw new Error("no way back shown");
  await shot("present-page");
  await page.keyboard.press("Escape");
  await page.waitForFunction(() => !document.querySelector("#presenter .hs-slide.hs-drill") && document.querySelector('#presenter .hs-slide[data-type="stepUp"]'), null, { timeout: 5000 });
  await page.waitForTimeout(700);
  if (await page.isHidden("#presenter")) throw new Error("Esc ended the presentation instead of going back");
  const hidden = await page.$$eval("#presenter .hs-player-slide [data-g].hs-hidden", (els) => new Set(els.map((el) => el.dataset.g)).size);
  if (hidden !== 2) throw new Error(`expected the last 2 stages still hidden, ${hidden} are`);
  await shot("present-back");
});

await step("present: the story skips the page, the end of the page returns, numbers jump by the story", async () => {
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowRight");
  await page.waitForSelector('#presenter .hs-slide[data-type="cards"]', { timeout: 5000 });
  await page.keyboard.press("ArrowLeft");
  await page.waitForSelector('#presenter .hs-slide[data-type="stepUp"]', { timeout: 5000 });
  await page.waitForTimeout(700);
  await page.locator('#presenter .hs-player-slide [data-drill]:not(svg *)').first().click();
  await page.waitForSelector("#presenter .hs-slide.hs-drill", { timeout: 5000 });
  await page.waitForTimeout(700);
  for (let i = 0; i < 6 && await page.$("#presenter .hs-slide.hs-drill"); i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(250); }
  await page.waitForSelector('#presenter .hs-slide[data-type="stepUp"]:not(.hs-drill)', { timeout: 5000 });
  await page.waitForTimeout(700);
  await page.keyboard.press("4");
  await page.keyboard.press("Enter");
  await page.waitForSelector('#presenter .hs-slide[data-type="statement"]', { timeout: 5000 });
  const counter = await page.textContent("#presenter .hs-player-count");
  if (!/^4 \/ 5$/.test(counter.trim())) throw new Error(`counter after jumping to 4: ${counter}`);
  await page.keyboard.press("Escape");
  await page.waitForSelector("#presenter.hidden", { state: "attached", timeout: 5000 });
});

await step("moving the slide takes its page along; deleting it deletes the page, undo brings both back", async () => {
  await page.dragAndDrop(".film-item:nth-child(2)", ".film-item:nth-child(5)");
  await page.waitForTimeout(400);
  expect(await films(), ["title", "cards", "stepUp", "↳", "statement", "closing"], "filmstrip after the move");
  await page.click(".film-item:nth-child(3)");
  await page.click('#inspector .inspector-foot button:has-text("削除")');
  await page.waitForTimeout(300);
  expect(await films(), ["title", "cards", "statement", "closing"], "filmstrip after deleting the slide");
  await page.click(".topbar");
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(400);
  expect(await films(), ["title", "cards", "stepUp", "↳", "statement", "closing"], "filmstrip after undo");
});

let exported = null;
await step("export HTML", async () => {
  const download = page.waitForEvent("download", { timeout: 60000 });
  await page.click("#downloadBtn");
  await page.waitForTimeout(800);
  if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn");
  const file = await download;
  exported = join(outDir, "drill-export.html");
  await file.saveAs(exported);
});

if (exported) {
  await step("the exported file opens the page and comes back", async () => {
    const view = await context.newPage();
    view.on("pageerror", (error) => errors.push(`export pageerror: ${error.message}`));
    await view.goto(pathToFileURL(exported).href);
    await view.waitForSelector(".hs-player .hs-slide", { timeout: 20000 });
    await view.keyboard.press("ArrowRight");
    await view.keyboard.press("ArrowRight");
    await view.waitForSelector('.hs-player .hs-slide[data-type="stepUp"]', { timeout: 5000 });
    for (let i = 0; i < 4; i += 1) { await view.keyboard.press("ArrowRight"); await view.waitForTimeout(150); }
    await view.waitForTimeout(600);
    await view.locator('.hs-player-slide [data-drill]:not(svg *)').first().click();
    await view.waitForSelector(".hs-player .hs-slide.hs-drill", { timeout: 5000 });
    await view.waitForTimeout(700);
    await shot("export-page", view);
    await view.keyboard.press("Escape");
    await view.waitForFunction(() => !document.querySelector(".hs-slide.hs-drill") && document.querySelector('.hs-slide[data-type="stepUp"]'), null, { timeout: 5000 });
    await view.close();
  });
}

await step("details: the item is marked 「＋ 詳しく」 on its text line, and × closes the card", async () => {
  // The slide from the report: a pyramid whose bottom level opens a card.
  const pyramidDeck = { title: "生成AIの歩みとこれから", theme: "editorial", transition: "fade", motion: {}, memo: "", slides: [
    { type: "title", title: "生成AIの歩みとこれから" },
    { type: "pyramid", title: "将来の活用を支える層", takeaway: "将来の競争力は、技術よりも**運用を積み上げる力**で差がつく。",
      levels: [{ title: "継続改善", description: "利用結果を振り返り、更新し続ける。" }, { title: "運用設計", description: "目的、担当、確認手順を決める。" }, { title: "ルール整備", description: "入力してよい情報を定める。" }],
      details: [{ target: "levels[2]", title: "ルール整備の範囲", text: "入力してよい情報、確認すべき出力、承認が必要な用途を明確にします。" }] },
    { type: "closing", message: "次のアクション" },
  ] };
  await page.evaluate((value) => localStorage.setItem("hpptx-studio-current-v1", JSON.stringify({ deck: value, selected: 1, savedAt: new Date().toISOString() })), pyramidDeck);
  await page.goto(base);
  await page.waitForSelector(".film-item");
  await page.click(".film-item:nth-child(2)");
  await page.click("#presentBtn");
  await page.waitForSelector('#presenter .hs-slide[data-type="pyramid"]');
  for (let i = 0; i < 3; i += 1) { await page.keyboard.press("ArrowRight"); await page.waitForTimeout(200); }
  await page.waitForSelector("#presenter .hs-player-flash", { timeout: 4000 });
  const hint = await page.textContent("#presenter .hs-player-flash");
  if (!/「＋ 詳しく」の付いた項目はクリックできます/.test(hint)) throw new Error(`hint: ${hint}`);
  await page.waitForTimeout(800);
  const mark = page.locator("#presenter .hs-detail-badge");
  const box = await mark.boundingBox();
  if (!box || box.width < 80) throw new Error(`the mark is not shown in full: ${JSON.stringify(box)}`);
  if ((await mark.evaluate((el) => el.parentElement.tagName)) !== "LI") throw new Error("the mark is not on the text line");
  // Marks let clicks through to their item; for this look only, let the browser find the mark itself.
  const hit = await page.evaluate(({ x, y }) => {
    const mark = document.querySelector("#presenter .hs-detail-badge");
    mark.style.pointerEvents = "auto";
    const found = document.elementFromPoint(x, y)?.closest(".hs-detail-badge") === mark;
    mark.style.pointerEvents = "";
    return found;
  }, { x: box.x + box.width / 2, y: box.y + box.height / 2 });
  if (!hit) throw new Error("the mark is covered or clipped");
  await shot("details-mark");
  await page.locator('#presenter li[data-detail]').click();
  await page.waitForSelector("#presenter .hs-popover");
  await page.click("#presenter .hs-popover-close");
  await page.waitForSelector("#presenter .hs-popover", { state: "detached", timeout: 3000 });
  if ((await page.getAttribute("#presenter .hs-player-slide .hs-slide", "data-type")) !== "pyramid") throw new Error("× moved to another slide");
  await page.keyboard.press("Escape");
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
