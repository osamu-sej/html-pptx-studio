// Motion graphics in a real browser: moving backdrops, kinetic type, strokes that draw themselves, entrances,
// emphasis styles, hover effects, the spotlight build, every transition (deck-wide and per slide), photo motion,
// and a Lottie animation from upload to the exported file.
// Usage: node qa/studio-motion.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir, readFile, writeFile } from "node:fs/promises";
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
const lottieFile = join(root, "qa", "fixtures", "spin.json");
await mkdir(outDir, { recursive: true });
const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
const shot = async (name, target = page) => { const file = join(outDir, `motion-${name}.png`); await target.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => { try { await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); } };
const film = (n) => page.click(`.film-item:nth-child(${n})`);

/** Lay slides out in a grid over the page (rendered by the engine itself) and start them. */
async function gallery(specs, { cols = 3, play = true } = {}) {
  await page.evaluate(({ specs, cols, play }) => {
    const E = window.SlideEngine;
    document.getElementById("qaGallery")?.remove();
    const box = document.createElement("div");
    box.id = "qaGallery";
    Object.assign(box.style, { position: "fixed", inset: "0", zIndex: "9999", background: "#11131a", display: "grid", gridTemplateColumns: `repeat(${cols}, 1fr)`, gap: "10px", padding: "10px", alignContent: "start", overflow: "auto" });
    document.body.append(box);
    for (const spec of specs) {
      const cell = document.createElement("div");
      cell.style.cssText = "color:#c8cfdd;font:600 12px sans-serif;display:grid;gap:4px";
      const el = E.render(spec.slide, { deck: spec.deck, index: spec.index ?? 1, mode: spec.mode || "present", assetBase: "/assets/" });
      cell.append(E.mount(el), spec.label);
      box.append(cell);
      if (play) requestAnimationFrame(() => E.play(el));
    }
  }, { specs, cols, play });
}
const clearGallery = () => page.evaluate(() => document.getElementById("qaGallery")?.remove());

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);

await step("the sample deck opens in its theme", async () => {
  await page.click("#sampleDeckBtn");
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length >= 40, null, { timeout: 20000 });
  await page.waitForSelector('.slide-wrap .hs-slide[data-theme="clarity"][data-type="title"]', { state: "attached", timeout: 5000 });
  await page.waitForTimeout(500);
  await shot("edit-cover");
});

await step("design dialog: live previews of every backdrop, kinetic and emphasis style", async () => {
  await page.click("#designBtn");
  await page.waitForFunction(() => document.querySelectorAll("#backdropGrid .motion-card").length === 13 && document.querySelectorAll("#kineticGrid .motion-card").length === 10 && document.querySelectorAll("#emphasisGrid .motion-card").length === 6);
  const options = await page.evaluate(() => ["transitionSelect", "entranceSelect", "hoverSelect"].map((id) => document.getElementById(id).options.length));
  if (options.join() !== "12,10,6") throw new Error(`the deck's motion menus are incomplete: ${options}`);
  await page.locator("#backdropGrid").scrollIntoViewIfNeeded();
  await page.waitForTimeout(1600);
  await shot("design-backdrops");
  const moving = await page.evaluate(() => [...document.querySelectorAll("#backdropGrid .hs-bd *")].some((el) => el.getAnimations().some((a) => a.playState === "running")));
  if (!moving) throw new Error("the backdrop previews do not move");
  await page.locator("#kineticGrid").scrollIntoViewIfNeeded();
  await page.waitForTimeout(3700);
  await page.waitForTimeout(450);
  await shot("design-kinetic");
  await page.locator("#emphasisGrid").scrollIntoViewIfNeeded();
  await page.evaluate(() => { for (const slide of document.querySelectorAll("#emphasisGrid .hs-slide")) window.SlideEngine.play(slide); });
  await page.waitForTimeout(2400);
  await shot("design-emphasis");
  await page.click('#emphasisGrid .motion-card[data-value="circle"]');
  await page.waitForTimeout(200);
  if ((await page.getAttribute('#emphasisGrid .motion-card[data-value="circle"]', "aria-checked")) !== "true") throw new Error("the circle card is not selected");
  await page.click('#emphasisGrid .motion-card[data-value="marker"]');
  await page.click('#backdropGrid .motion-card:has-text("グリッド")');
  await page.waitForTimeout(300);
  if ((await page.getAttribute('#backdropGrid .motion-card[data-value="grid"]', "aria-checked")) !== "true") throw new Error("the grid card is not selected");
  if (!(await page.locator('.slide-wrap .hs-slide[data-backdrop="grid"]').count())) throw new Error("the cover did not take the grid backdrop");
  await page.click('#backdropGrid .motion-card[data-value="orbits"]');
  await page.click("#designDialog .dialog-foot [data-close]");
});

await step("presenting the cover: kinetic title over moving orbits", async () => {
  await film(1);
  await page.click("#presentBtn");
  await page.waitForSelector(".hs-player .hs-slide", { timeout: 10000 });
  await page.waitForTimeout(260);
  await shot("present-cover-mid");
  const state = await page.evaluate(() => ({
    units: document.querySelectorAll(".hs-player .hs-kin .hs-k").length,
    masks: document.querySelectorAll(".hs-player .hs-kin .hs-km").length,
    orbit: [...document.querySelectorAll(".hs-player .hs-bd-sat")].some((el) => el.getAnimations().length),
  }));
  if (state.units < 4 || !state.masks) throw new Error(`the title was not split for the mask motion (${JSON.stringify(state)})`);
  if (!state.orbit) throw new Error("the satellites do not orbit");
  await page.waitForTimeout(1800);
  await shot("present-cover-end");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
});

await step("kinetic styles (statement slide)", async () => {
  const slide = { type: "statement", title: "私たちが目指すこと", text: "毎月**1,440時間**を、考える仕事に使える会社へ", takeaway: "5部門120人・3か月の試行から" };
  const styles = await page.evaluate(() => Object.keys(window.SlideEngine.KINETIC));
  const specs = styles.map((style) => ({ slide: { ...slide, kinetic: style }, deck: { title: "QA", theme: "midnight", motion: { backdrop: "gradient" }, slides: [{ type: "title", title: "QA" }, slide, { type: "closing" }] }, label: style }));
  await gallery(specs, { cols: 3 });
  await page.waitForTimeout(420);
  const mid = await page.evaluate(() => ({ on: document.querySelectorAll("#qaGallery .hs-k.on").length }));
  await shot("kinetic-mid");
  await page.waitForTimeout(2800);
  await shot("kinetic-end");
  const done = await page.evaluate(() => [...document.querySelectorAll("#qaGallery .hs-slide")].map((el) => el.querySelector(".hs-statement-text").textContent));
  await clearGallery();
  if (!mid.on) throw new Error("the decode style did not scramble");
  if (done.some((text) => text !== "毎月1,440時間を、考える仕事に使える会社へ")) throw new Error(`text changed after the motion: ${JSON.stringify(done)}`);
});

await step("backdrops in a light and a dark theme", async () => {
  const kinds = await page.evaluate(() => Object.keys(window.SlideEngine.BACKDROPS));
  for (const theme of ["clarity", "midnight", "kinari"]) {
    const slide = { type: "section", title: "展開計画", takeaway: "10月から3段階で全社へ" };
    await gallery([...kinds.map((kind) => ({ slide: { ...slide, backdrop: kind }, deck: { title: "QA", theme, slides: [{ type: "title", title: "QA" }, slide, { type: "closing" }] }, label: `${theme} / ${kind}`, mode: "preview" })),
      { slide: { type: "cards", title: "本文スライドにも薄く", takeaway: "カードの後ろで動く", backdrop: "particles", items: [{ title: "資料作成", desc: "月620時間", icon: "document" }, { title: "議事録", desc: "月380時間", icon: "mic" }, { title: "集計", desc: "月290時間", icon: "chartBar" }] }, deck: { title: "QA", theme, slides: [{ type: "title", title: "QA" }, {}, { type: "closing" }] }, label: `${theme} / cards + particles`, mode: "preview" }], { cols: 5, play: false });
    await page.waitForTimeout(1400);
    await shot(`backdrops-${theme}`);
  }
  await clearGallery();
});

await step("icons draw themselves and the marker sweeps in", async () => {
  const index = await page.evaluate(() => [...document.querySelectorAll(".film-item")].findIndex((el) => el.querySelector('.hs-slide[data-type="cards"], .hs-slide[data-type="headerCards"], .hs-slide[data-type="bulletCards"]')));
  if (index < 0) throw new Error("no card slide in the sample");
  await film(index + 1);
  await page.click('button:has-text("▶ 動きを確認")');
  await page.waitForTimeout(900);
  const mid = await page.evaluate(() => [...document.querySelectorAll(".slide-wrap .hs-icon > *")].map((el) => parseFloat(getComputedStyle(el).strokeDashoffset)));
  await shot("draw-mid");
  await page.waitForTimeout(2600);
  const end = await page.evaluate(() => [...document.querySelectorAll(".slide-wrap .hs-icon > *")].map((el) => parseFloat(getComputedStyle(el).strokeDashoffset)));
  await shot("draw-end");
  await page.keyboard.press("Escape");
  if (!mid.length) throw new Error("no icons on the slide");
  if (!mid.some((value) => value > 0.02)) throw new Error(`icons were not drawing at 0.9s: ${mid.slice(0, 6)}`);
  if (end.some((value) => value > 0.001)) throw new Error(`icons did not finish drawing: ${end.slice(0, 6)}`);
});


const cardsSlide = { type: "cards", title: "3つの打ち手", takeaway: "まず**資料作成**から始める", items: [{ title: "資料作成", desc: "月620時間を削減", icon: "document" }, { title: "議事録", desc: "月380時間を削減", icon: "mic" }, { title: "集計", desc: "月290時間を削減", icon: "chartBar" }] };
const qaDeck = (slide, extra = {}) => ({ title: "QA", theme: "clarity", slides: [{ type: "title", title: "QA" }, slide, { type: "closing" }], ...extra });

await step("entrances: every style lands on a readable slide", async () => {
  const kinds = await page.evaluate(() => Object.keys(window.SlideEngine.ENTRANCES));
  await gallery(kinds.map((kind) => ({ slide: cardsSlide, deck: qaDeck(cardsSlide, { motion: { entrance: kind } }), label: kind })), { cols: 3 });
  await page.waitForTimeout(240);
  await shot("entrances-mid");
  await page.waitForTimeout(2200);
  await shot("entrances-end");
  const stuck = await page.evaluate(() => [...document.querySelectorAll("#qaGallery .hs-slide")].flatMap((el) => [...el.querySelectorAll("[data-g], .hs-head")].filter((part) => Number(getComputedStyle(part).opacity) < 0.99).map(() => el.dataset.entrance)));
  await clearGallery();
  if (stuck.length) throw new Error(`parts stayed faded after the entrance: ${[...new Set(stuck)]}`);
});

await step("emphasis styles draw around the phrase", async () => {
  const styles = await page.evaluate(() => Object.keys(window.SlideEngine.EMPHASES));
  const statement = { type: "statement", title: "私たちが目指すこと", text: "毎月**1,440時間**を、考える仕事へ" };
  await gallery([
    ...styles.map((style) => ({ slide: { ...statement, kinetic: "none" }, deck: qaDeck(statement, { motion: { emphasis: style } }), label: `statement / ${style}` })),
    ...styles.map((style) => ({ slide: cardsSlide, deck: qaDeck(cardsSlide, { theme: "midnight", motion: { emphasis: style } }), label: `takeaway (midnight) / ${style}` })),
  ], { cols: 4 });
  await page.waitForTimeout(700);
  await shot("emphasis-mid");
  await page.waitForTimeout(2400);
  await shot("emphasis-end");
  const drawn = await page.evaluate(() => [...document.querySelectorAll("#qaGallery .hs-slide")].map((el) => ({ style: el.dataset.emphasis, running: el.querySelector(".hs-em").getAnimations({ subtree: true }).filter((a) => a.playState === "running").length })));
  await clearGallery();
  const still = drawn.filter((entry) => entry.running);
  if (still.length) throw new Error(`emphasis still animating after 3s: ${JSON.stringify(still)}`);
});

await step("photo motion: every style on a full-bleed photo", async () => {
  const motions = await page.evaluate(() => Object.keys(window.SlideEngine.PHOTO_MOTIONS));
  const hero = { type: "hero", title: "現場の時間を、お客様へ", takeaway: "AIで事務作業を減らす", visualAsset: "storeOperations" };
  await gallery(motions.map((motion) => ({ slide: { ...hero, photoMotion: motion }, deck: qaDeck(hero, { theme: "midnight" }), label: motion })), { cols: 4 });
  await page.waitForTimeout(500);
  await shot("photo-mid");
  const moving = await page.evaluate(() => [...document.querySelectorAll("#qaGallery .hs-media")].filter((el) => el.dataset.motion !== "parallax").map((el) => ({ motion: el.dataset.motion, running: el.getAnimations({ subtree: true }).length > 0 })));
  await page.waitForTimeout(2500);
  await shot("photo-end");
  await clearGallery();
  const idle = moving.filter((entry) => !entry.running).map((entry) => entry.motion);
  if (idle.length) throw new Error(`photo motion does not play: ${idle}`);
});

await step("hover: tilt, glow and zoom react to the pointer in the presentation", async () => {
  const index = await page.evaluate(() => [...document.querySelectorAll(".film-item")].findIndex((el) => el.querySelector('.hs-slide[data-type="cards"], .hs-slide[data-type="headerCards"]')));
  if (index < 0) throw new Error("no card slide in the sample");
  for (const hover of ["tilt", "glow", "zoom"]) {
    await page.click("#designBtn");
    await page.selectOption("#hoverSelect", hover);
    await page.click("#designDialog .dialog-foot [data-close]");
    await film(index + 1);
    await page.click("#presentBtn");
    await page.waitForSelector(".hs-player .hs-slide", { timeout: 10000 });
    await page.waitForTimeout(1600);
    const box = await page.locator(".hs-player .hs-slide .hs-item").first().boundingBox();
    await page.mouse.move(box.x + box.width * 0.85, box.y + box.height * 0.2);
    await page.waitForTimeout(500);
    const look = await page.evaluate(() => { const el = document.querySelector(".hs-player .hs-item.hs-hot"); const css = el && getComputedStyle(el); return el && { hover: el.closest(".hs-slide").dataset.hover, transform: css.transform, shadow: css.boxShadow }; });
    await shot(`hover-${hover}`);
    await page.keyboard.press("Escape");
    await page.waitForTimeout(300);
    if (!look) throw new Error(`${hover}: no item under the pointer`);
    if (look.hover !== hover) throw new Error(`${hover}: the slide says ${look.hover}`);
    if (hover !== "glow" && look.transform === "none") throw new Error(`${hover}: the item did not move`);
    if (hover === "glow" && look.shadow === "none") throw new Error("glow: no light around the item");
  }
  await page.click("#designBtn");
  await page.selectOption("#hoverSelect", "lift");
  await page.click("#designDialog .dialog-foot [data-close]");
});

await step("spotlight: every item stays, each click puts the next in focus", async () => {
  const index = await page.evaluate(() => [...document.querySelectorAll(".film-item")].findIndex((el) => el.querySelector('.hs-slide[data-type="cards"], .hs-slide[data-type="headerCards"]')));
  await film(index + 1);
  await page.click("#formTab");
  await page.click('#inspector .motion-grid button:has-text("順に注目")');
  await page.waitForTimeout(200);
  await page.locator("#inspector .motion-grid").scrollIntoViewIfNeeded();
  await shot("spotlight-inspector");
  await page.click("#presentBtn");
  await page.waitForSelector(".hs-player .hs-slide[data-build=spotlight]", { timeout: 10000 });
  await page.waitForTimeout(1500);
  const before = await page.evaluate(() => ({ hidden: document.querySelectorAll(".hs-player .hs-hidden").length, spot: document.querySelectorAll(".hs-player .hs-spot").length }));
  await page.keyboard.press("ArrowRight");
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(700);
  const after = await page.evaluate(() => ({ spot: [...document.querySelectorAll(".hs-player .hs-spot")].map((el) => el.dataset.g), faded: [...document.querySelectorAll(".hs-player [data-g]:not(.hs-spot)")].filter((el) => Number(getComputedStyle(el).opacity) < 0.5).length }));
  await shot("spotlight-present");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  await page.keyboard.press("Control+z");
  if (before.hidden || before.spot) throw new Error(`spotlight should start with everything plainly visible: ${JSON.stringify(before)}`);
  if (!after.spot.length || after.spot.some((g) => g !== "1")) throw new Error(`the second item is not in focus: ${JSON.stringify(after)}`);
  if (!after.faded) throw new Error("the other items did not fade back");
});

await step("transitions: every kind, and no leftovers", async () => {
  for (const [transition, how] of [["wipe", "key"], ["circle", "click"], ["push", "key"], ["flip", "key"], ["dive", "key"], ["blinds", "key"], ["curtain", "click"]]) {
    await page.click("#designBtn");
    await page.selectOption("#transitionSelect", transition);
    await page.click("#designDialog .dialog-foot [data-close]");
    await film(2);
    await page.click("#presentBtn");
    await page.waitForSelector(".hs-player .hs-slide", { timeout: 10000 });
    await page.waitForTimeout(900);
    if (how === "key") await page.keyboard.press("ArrowRight");
    else await page.mouse.click(1200, 300);
    await page.waitForTimeout(transition === "push" ? 280 : 400);
    const mid = await page.evaluate((t) => ({ band: document.querySelectorAll(".hs-tr-band").length, entering: document.querySelectorAll(`.hs-tr-in-${t}`).length }), transition);
    await shot(`transition-${transition}`);
    await page.waitForTimeout(1000);
    const after = await page.evaluate(() => ({ band: document.querySelectorAll(".hs-tr-band").length, slides: document.querySelectorAll(".hs-player-stage > .hs-scaler").length }));
    await page.keyboard.press("Escape");
    await page.waitForTimeout(300);
    if (!mid.entering) throw new Error(`${transition}: the next slide was not entering`);
    if (transition === "wipe" && !mid.band) throw new Error("wipe: no colour band");
    if (after.band || after.slides !== 1) throw new Error(`${transition}: leftovers after the transition ${JSON.stringify(after)}`);
  }
  await page.click("#designBtn");
  await page.selectOption("#transitionSelect", "fade");
  await page.click("#designDialog .dialog-foot [data-close]");
});

await step("one slide's own transition overrides the deck's (forward and back)", async () => {
  await film(3);
  await page.click("#formTab");
  await page.selectOption('#inspector select[data-mg="transition"]', "curtain");
  await film(2);
  await page.click("#presentBtn");
  await page.waitForSelector(".hs-player .hs-slide", { timeout: 10000 });
  await page.waitForTimeout(900);
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(300);
  const forward = await page.evaluate(() => document.querySelectorAll(".hs-tr-in-curtain").length);
  await page.waitForTimeout(1200);
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(200);
  const next = await page.evaluate(() => ({ curtain: document.querySelectorAll(".hs-tr-in-curtain").length, fade: document.querySelectorAll(".hs-tr-in-fade, .hs-tr-in").length }));
  await page.waitForTimeout(1000);
  await page.keyboard.press("ArrowLeft");
  await page.waitForTimeout(1200);
  await page.keyboard.press("ArrowLeft");
  await page.waitForTimeout(300);
  const back = await page.evaluate(() => document.querySelectorAll(".hs-tr-in-curtain").length);
  await page.waitForTimeout(1000);
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  await film(3);
  await page.keyboard.press("Control+z");
  if (!forward) throw new Error("the slide's own curtain did not play on the way in");
  if (next.curtain) throw new Error("the curtain leaked onto the next slide");
  if (!back) throw new Error("going back out of the slide should close with its own curtain");
});

await step("one slide's own motion graphics from the edit tab", async () => {
  await film(4);
  await page.click("#formTab");
  await page.selectOption('#inspector select[data-mg="backdrop"]', "waves");
  await page.selectOption('#inspector select[data-mg="kinetic"]', "chars");
  await page.selectOption('#inspector select[data-mg="entrance"]', "drop");
  await page.selectOption('#inspector select[data-mg="emphasis"]', "underline");
  await page.waitForTimeout(400);
  const attrs = await page.evaluate(() => { const el = document.querySelector(".slide-wrap .hs-slide"); return { backdrop: el.dataset.backdrop, kinetic: el.dataset.kinetic, entrance: el.dataset.entrance, emphasis: el.dataset.emphasis }; });
  await page.locator('#inspector [data-key="motion-graphics"]').scrollIntoViewIfNeeded();
  await shot("inspector");
  if (attrs.backdrop !== "waves" || attrs.kinetic !== "chars" || attrs.entrance !== "drop" || attrs.emphasis !== "underline") throw new Error(`the slide did not take its own settings: ${JSON.stringify(attrs)}`);
  for (let i = 0; i < 4; i += 1) await page.keyboard.press("Control+z");
});

let lottieIndex = 5;
await step("a Lottie animation: one frame in the editor, playing in the presentation", async () => {
  await film(lottieIndex);
  await page.click("#formTab");
  await page.setInputFiles('#inspector input[type=file][accept^=".json"]', lottieFile);
  await page.waitForSelector(".slide-wrap .hs-lottie-host svg", { timeout: 10000 });
  await page.waitForTimeout(400);
  await shot("lottie-edit");
  const still = await page.evaluate(() => { const anim = document.querySelector(".slide-wrap .hs-lottie-host").hsAnim; return { paused: anim.isPaused, frame: Math.round(anim.currentFrame) }; });
  if (!still.paused) throw new Error("the editor should hold one frame");
  await page.click("#presentBtn");
  await page.waitForSelector(".hs-player .hs-lottie-host svg", { timeout: 10000 });
  const a = await page.evaluate(() => document.querySelector(".hs-player .hs-lottie-host").hsAnim.currentFrame);
  await page.waitForTimeout(500);
  const b = await page.evaluate(() => ({ frame: document.querySelector(".hs-player .hs-lottie-host").hsAnim.currentFrame, paused: document.querySelector(".hs-player .hs-lottie-host").hsAnim.isPaused }));
  await shot("lottie-present");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(300);
  if (b.paused || b.frame === a) throw new Error(`the animation does not play (${a} → ${b.frame})`);
});

await step("the exported file carries the Lottie player and plays the animation", async () => {
  const download = page.waitForEvent("download", { timeout: 60000 });
  await page.click("#downloadBtn");
  await page.waitForTimeout(800);
  if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn");
  const file = join(outDir, "motion-export.html");
  await (await download).saveAs(file);
  const html = await readFile(file, "utf8");
  if (!/loadAnimation/.test(html)) throw new Error("the Lottie player is not in the file");
  const view = await context.newPage();
  view.on("pageerror", (error) => errors.push(`export pageerror: ${error.message}`));
  await view.goto(`${pathToFileURL(file).href}#${lottieIndex}`);
  await view.waitForSelector(".hs-player .hs-lottie-host svg", { timeout: 20000 });
  await view.waitForTimeout(700);
  const playing = await view.evaluate(() => { const anim = document.querySelector(".hs-player .hs-lottie-host").hsAnim; return anim && !anim.isPaused; });
  await shot("export-lottie", view);
  await view.goto(`${pathToFileURL(file).href}#1`);
  await view.reload();
  await view.waitForSelector(".hs-player .hs-kin", { timeout: 20000 });
  await view.waitForTimeout(300);
  await shot("export-cover", view);
  await view.close();
  if (!playing) throw new Error("the animation does not play in the exported file");
});

await writeFile(join(outDir, "motion-errors.txt"), errors.join("\n"));
console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
