import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { allLayoutSlides } from "./fixtures.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));

/** Load engine.js + motion.js into a DOM, exactly as a browser page (and an exported file) does. */
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const engine = (await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons);
  const objects = await readFile(join(root, "public", "engine", "objects.js"), "utf8");
  const animate = await readFile(join(root, "public", "engine", "animate.js"), "utf8");
  const motion = await readFile(join(root, "public", "engine", "motion.js"), "utf8");
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  vm.runInContext(engine, context, { filename: "engine.js" });
  vm.runInContext(objects, context, { filename: "objects.js" });
  vm.runInContext(animate, context, { filename: "animate.js" });
  vm.runInContext(motion, context, { filename: "motion.js" });
  return { E: window.SlideEngine, window };
}

const deckOf = (slides, extra = {}) => ({ title: "検証デッキ", audience: "役員", purpose: "意思決定", theme: "clarity", transition: "fade", motion: {}, slides, ...extra });

test("the engine renders every layout in every theme with editable, animatable parts", async () => {
  const { E } = await loadEngine();
  assert.equal(E.THEMES.length, 8);
  assert.equal(E.THEMES[0].id, E.DEFAULT_THEME, "the default theme comes first");
  for (const theme of E.THEMES) {
    assert.equal(theme.colors.accents.length, 6, `${theme.id}: PowerPoint's six accents`);
    assert.ok(theme.officeFonts.heading && theme.officeFonts.body, `${theme.id}: its theme fonts`);
  }
  // 44 layouts the AI writes, and 白紙 (blank) for pages people build from objects.
  assert.equal(Object.keys(E.TYPE_LABELS).length, 45);
  assert.equal(E.TYPE_LABELS.blank, "白紙（自由配置）");
  const types = new Set(allLayoutSlides.map((slide) => slide.type));
  assert.equal(types.size, 44, "the fixture covers every layout the AI writes");
  for (const theme of E.THEMES) {
    const deck = deckOf(allLayoutSlides, { theme: theme.id });
    deck.slides.forEach((slide, index) => {
      const el = E.render(slide, { deck, index, mode: "edit" });
      assert.ok(el.classList.contains("hs-slide"), `${theme.id} #${index + 1}`);
      assert.equal(el.dataset.theme, theme.id);
      assert.equal(el.dataset.type, slide.type);
      assert.ok(el.querySelector("[data-field]"), `${theme.id} ${slide.type}: nothing editable`);
      if (!["title", "section", "closing", "hero", "statement"].includes(slide.type)) {
        assert.ok(el.querySelector(".hs-title[data-field=title]"), `${slide.type}: title`);
        // Pages are numbered "03 / 44".
        assert.match(el.querySelector(".hs-foot .hs-page").textContent, new RegExp(`${String(index + 1).padStart(2, "0")} / ${String(deck.slides.length).padStart(2, "0")}`));
      }
    });
  }
});

test("builds: steps appear on click, parallel items cascade, covers stay still", async () => {
  const { E } = await loadEngine();
  const deck = deckOf(allLayoutSlides);
  const byType = (type) => allLayoutSlides.findIndex((slide) => slide.type === type);
  const process = E.render(allLayoutSlides[byType("process")], { deck, index: byType("process"), mode: "present" });
  assert.equal(process.dataset.build, "click");
  assert.equal(Number(process.dataset.steps), 3, "one click per step");
  assert.equal(process.querySelectorAll("[data-g]").length >= 3, true);
  const cards = E.render(allLayoutSlides[byType("cards")], { deck, index: byType("cards"), mode: "present" });
  assert.equal(cards.dataset.build, "cascade");
  assert.equal(cards.dataset.steps, "0");
  const cover = E.render(allLayoutSlides[0], { deck, index: 0, mode: "present" });
  assert.equal(cover.dataset.build, "none");
  const forced = E.render({ ...allLayoutSlides[byType("cards")], animation: "click" }, { deck, index: 3, mode: "present" });
  assert.equal(Number(forced.dataset.steps), 3);
  E.play(process, { step: 0, animate: false });
  assert.equal(process.querySelectorAll("[data-g].hs-hidden").length, process.querySelectorAll("[data-g]").length, "nothing shown before the first click");
  E.reveal(process, 1);
  assert.equal(process.querySelectorAll('[data-g="0"].hs-hidden').length, 0);
});

test("numbers count up, charts are drawn as SVG and details get a badge", async () => {
  const { E } = await loadEngine();
  const sample = JSON.parse(await readFile(join(root, "public", "samples", "ai-rollout.json"), "utf8"));
  const deck = deckOf(sample.slideData);
  const kpiIndex = sample.slideData.findIndex((slide) => slide.type === "kpi" && slide.details?.length);
  const kpi = E.render(sample.slideData[kpiIndex], { deck, index: kpiIndex, mode: "present" });
  assert.ok(kpi.querySelector(".hs-count[data-count]"), "figures are marked for counting up");
  assert.ok(kpi.querySelector("[data-detail] .hs-detail-badge"), "a detail badge marks the clickable item");
  const chartIndex = sample.slideData.findIndex((slide) => slide.type === "imageText" && slide.image?.chartType === "bar");
  const chart = E.render(sample.slideData[chartIndex], { deck, index: chartIndex, mode: "present" });
  assert.ok(chart.querySelector("svg .hs-mark[data-tip]"), "bars carry a hover tooltip");
  assert.equal(E.numParts("1,440h").num, "1,440");
  assert.equal(E.numParts("4.3点").unit, "点");
});

test("photos and videos: slots, placement, YouTube and browser-kept files", async () => {
  const { E } = await loadEngine();
  const slides = [
    { type: "title", title: "表紙" },
    { type: "hero", title: "全面写真", visualAsset: "ai", photoMotion: "zoom" },
    { type: "cards", title: "動画付き", takeaway: "結論", items: [{ title: "A" }, { title: "B" }], media: { src: "https://www.youtube.com/watch?v=dQw4w9WgXcQ", kind: "video" } },
    { type: "content", title: "録画", takeaway: "結論", points: ["A"], media: { src: "idb:abc", kind: "video", autoplay: true, muted: true, placement: { x: 0.1, y: 0.2, w: 0.3, h: 0.3 } } },
    { type: "closing", message: "以上" },
  ];
  const deck = deckOf(slides);
  const hero = E.render(slides[1], { deck, index: 1, mode: "present", assetBase: "/assets/" });
  assert.equal(hero.querySelector(".hs-media").dataset.motion, "zoom");
  assert.match(hero.querySelector(".hs-media img").getAttribute("src"), /\/assets\/ai-executive-hero\.jpg$/);
  const exported = E.render(slides[1], { deck, index: 1, mode: "present", assetMap: { "ai-executive-hero.jpg": "data:image/jpeg;base64,AAAA" } });
  assert.equal(exported.querySelector(".hs-media img").getAttribute("src"), "data:image/jpeg;base64,AAAA", "exports carry their photos");
  const youtubeLive = E.render(slides[2], { deck, index: 2, mode: "present" });
  assert.match(youtubeLive.querySelector(".hs-placed iframe").getAttribute("src"), /youtube-nocookie\.com\/embed\/dQw4w9WgXcQ/);
  // YouTube refuses to play (error 153) unless the player is told which site embeds it.
  assert.equal(youtubeLive.querySelector(".hs-placed iframe").getAttribute("referrerpolicy"), "strict-origin-when-cross-origin");
  const youtubeEdit = E.render(slides[2], { deck, index: 2, mode: "edit" });
  assert.equal(youtubeEdit.querySelector("iframe"), null, "the editor shows a still picture, not a player");
  const video = E.render(slides[3], { deck, index: 3, mode: "present", mediaUrls: { "idb:abc": "blob:video" } });
  const placed = video.querySelector(".hs-placed");
  assert.equal(placed.style.left, "10%");
  assert.equal(placed.querySelector("video").getAttribute("src"), "blob:video");
  assert.ok(placed.querySelector("video").hasAttribute("data-autoplay"));
  assert.equal(E.youtubeId("https://youtu.be/dQw4w9WgXcQ"), "dQw4w9WgXcQ");
});

test("motion graphics: kinetic type and backdrops by default, per deck and per slide", async () => {
  const { E } = await loadEngine();
  const slides = [
    { type: "title", title: "表紙" },
    { type: "section", title: "章" },
    { type: "content", title: "本文", takeaway: "結論", points: ["A", "B"] },
    { type: "hero", title: "全面写真", visualAsset: "ai" },
    { type: "closing", message: "以上" },
  ];
  const plain = deckOf(slides);
  const cover = E.render(slides[0], { deck: plain, index: 0, mode: "present" });
  assert.equal(cover.dataset.kinetic, "mask", "big lines move by default");
  assert.equal(cover.dataset.backdrop, undefined, "no backdrop unless chosen");
  assert.ok(cover.querySelector(".hs-cover-art"), "the theme's cover art stays");
  assert.equal(cover.dataset.draw, "on");
  assert.equal(E.render(slides[2], { deck: plain, index: 2, mode: "present" }).dataset.kinetic, undefined, "body slides keep their titles still");

  const moving = deckOf(slides, { motion: { kinetic: "type", backdrop: "orbits", draw: false } });
  const title = E.render(slides[0], { deck: moving, index: 0, mode: "present" });
  assert.equal(title.dataset.kinetic, "type");
  assert.ok(title.querySelector('.hs-decor .hs-bd[data-bd="orbits"] .hs-bd-sat'), "satellites on their orbits");
  assert.equal(title.querySelector(".hs-cover-art"), null, "the backdrop replaces the cover art");
  assert.equal(title.dataset.draw, "off");
  assert.equal(E.render(slides[1], { deck: moving, index: 1, mode: "present" }).dataset.backdrop, "orbits");
  assert.equal(E.render(slides[2], { deck: moving, index: 2, mode: "present" }).dataset.backdrop, undefined, "the deck's backdrop stays on stage slides");
  assert.equal(E.render(slides[3], { deck: moving, index: 3, mode: "present" }).dataset.backdrop, undefined, "a full-bleed photo is its own backdrop");

  const own = E.render({ ...slides[2], kinetic: "chars", backdrop: "waves" }, { deck: moving, index: 2, mode: "present" });
  assert.equal(own.dataset.kinetic, "chars");
  assert.equal(own.dataset.backdrop, "waves");
  const off = E.render({ ...slides[1], kinetic: "none", backdrop: "none" }, { deck: moving, index: 1, mode: "present" });
  assert.equal(off.dataset.kinetic, undefined);
  assert.equal(off.querySelector(".hs-bd"), null);

  for (const kind of Object.keys(E.BACKDROPS)) {
    const a = E.render({ ...slides[1], backdrop: kind }, { deck: plain, index: 1, mode: "thumb" });
    const b = E.render({ ...slides[1], backdrop: kind }, { deck: plain, index: 1, mode: "present" });
    assert.ok(a.querySelector(`.hs-bd[data-bd="${kind}"]`).children.length > 0, `${kind} draws something`);
    assert.equal(a.querySelector(".hs-bd").outerHTML, b.querySelector(".hs-bd").outerHTML, `${kind}: thumbnails match the presentation`);
  }
  const icon = E.render({ type: "cards", title: "t", takeaway: "k", items: [{ title: "A", icon: "rocket" }] }, { deck: plain, index: 2, mode: "present" }).querySelector(".hs-icon");
  assert.ok([...icon.children].every((shape) => shape.getAttribute("pathLength") === "1"), "icon strokes can draw themselves");
});

test("kinetic type splits the big lines without changing their text or line-break rules", async () => {
  const { E } = await loadEngine();
  const slide = { type: "statement", title: "目指すこと", text: "毎月**1,440時間**を、考える仕事へ。AI活用" };
  const deck = deckOf([{ type: "title", title: "t" }, slide, { type: "closing" }]);
  for (const mode of ["mask", "words", "chars", "type", "wave", "zoom", "flip", "slide"]) {
    const el = E.render({ ...slide, kinetic: mode }, { deck, index: 1, mode: "present" });
    const text = el.querySelector(".hs-statement-text");
    const before = text.textContent;
    E.play(el);
    const units = [...text.querySelectorAll(".hs-k")];
    assert.ok(text.classList.contains("hs-kin"), mode);
    assert.equal(text.textContent, before, `${mode}: the words stay the same`);
    assert.ok(units.length > 3, `${mode}: split into units`);
    assert.ok(!units.some((unit) => /^[、。]/.test(unit.textContent)), `${mode}: no unit starts with 、 or 。`);
    assert.ok(text.querySelector(".hs-em .hs-k"), `${mode}: the emphasis keeps its marker`);
    assert.match(text.style.getPropertyValue("--kst"), /^\d+ms$/);
    assert.equal(units.at(-1).classList.contains("hs-k-last"), true);
    if (mode === "mask") assert.ok(text.querySelector(".hs-km > .hs-k"), "mask: each unit rises from behind its own edge");
    if (["chars", "wave", "flip"].includes(mode)) assert.deepEqual([...text.querySelectorAll(".hs-kw")].map((word) => word.textContent), ["1,440", "AI"], `${mode}: Latin words and figures do not break`);
    if (["zoom", "slide"].includes(mode)) assert.ok(units.length < [...before].length / 2, `${mode}: moves word by word`);
    E.play(el);
    assert.equal(text.querySelectorAll(".hs-k").length, units.length, `${mode}: playing again does not split twice`);
  }
});

test("more motion types: entrances, hovers, emphasis, spotlight and photo motion, per deck and per slide", async () => {
  const { E } = await loadEngine();
  const schemas = await import("../server/schemas.mjs");
  // The AI chooses from the same lists the engine can play.
  assert.deepEqual(Object.keys(E.KINETIC), schemas.KINETIC_STYLES);
  assert.deepEqual(Object.keys(E.BACKDROPS), schemas.BACKDROP_KINDS);
  assert.deepEqual(Object.keys(E.ENTRANCES), schemas.ENTRANCES);
  assert.deepEqual(Object.keys(E.HOVERS), schemas.HOVERS);
  assert.deepEqual(Object.keys(E.EMPHASES), schemas.EMPHASES);
  assert.deepEqual(Object.keys(E.TRANSITIONS), schemas.TRANSITIONS);
  assert.deepEqual(Object.keys(E.PHOTO_MOTIONS), schemas.PHOTO_MOTIONS);
  assert.deepEqual([...E.BUILDS], schemas.BUILDS);

  const slides = [
    { type: "title", title: "表紙" },
    { type: "cards", title: "選択肢", takeaway: "**B案**を選ぶ", items: [{ title: "A" }, { title: "B" }, { title: "C" }] },
    { type: "hero", title: "全面写真", visualAsset: "ai", photoMotion: "drift" },
    { type: "closing", message: "以上" },
  ];
  const deck = deckOf(slides, { motion: { entrance: "zoom", hover: "tilt", emphasis: "circle" } });
  const cards = E.render(slides[1], { deck, index: 1, mode: "present" });
  assert.equal(cards.dataset.entrance, "zoom");
  assert.equal(cards.dataset.hover, "tilt");
  assert.equal(cards.dataset.emphasis, "circle");
  assert.ok(cards.querySelector(".hs-em"), "the phrase to emphasise is marked");
  const own = E.render({ ...slides[1], entrance: "drop", emphasis: "box" }, { deck, index: 1, mode: "present" });
  assert.equal(own.dataset.entrance, "drop", "a slide may pick its own entrance");
  assert.equal(E.render({ ...slides[1], entrance: "spin" }, { deck, index: 1, mode: "present" }).dataset.entrance, "spin", "the showpiece entrances are entrances");
  assert.equal(own.dataset.emphasis, "box");
  assert.equal(E.render({ ...slides[1], entrance: "none" }, { deck, index: 1, mode: "present" }).dataset.entrance, "none");
  const unknown = E.render({ ...slides[1], entrance: "tornado", emphasis: "sparkle" }, { deck: deckOf(slides, { motion: { hover: "wobble" } }), index: 1, mode: "present" });
  assert.equal(unknown.dataset.entrance, "rise", "unknown values fall back to the defaults");
  assert.equal(unknown.dataset.emphasis, "marker");
  assert.equal(unknown.dataset.hover, "lift");

  // Spotlight: everything stays on screen, and each click puts one item in focus.
  const spot = E.render({ ...slides[1], animation: "spotlight" }, { deck, index: 1, mode: "present" });
  assert.equal(spot.dataset.build, "spotlight");
  assert.equal(Number(spot.dataset.steps), 3, "one click per item");
  E.play(spot, { step: 0, animate: false });
  assert.equal(spot.querySelectorAll(".hs-hidden").length, 0, "nothing is hidden");
  assert.ok(!spot.classList.contains("hs-spotting"));
  E.reveal(spot, 2);
  assert.ok(spot.classList.contains("hs-spotting"));
  assert.deepEqual([...spot.querySelectorAll(".hs-spot")].map((el) => el.dataset.g), [...spot.querySelectorAll('[data-g="1"]')].map((el) => el.dataset.g));
  assert.ok(spot.querySelector('[data-g="1"].hs-spot'), "the second item is in focus");
  assert.equal(spot.querySelector('[data-g="0"].hs-spot'), null);
  E.play(spot, { step: 3, animate: false });
  assert.ok(spot.querySelector('[data-g="2"].hs-spot'), "stepping back lands on the right item");

  for (const motion of Object.keys(E.PHOTO_MOTIONS)) {
    const hero = E.render({ ...slides[2], photoMotion: motion }, { deck, index: 2, mode: "present" });
    assert.equal(hero.querySelector(".hs-media").dataset.motion, motion);
  }
  assert.equal(E.render({ ...slides[2], photoMotion: "shake" }, { deck, index: 2, mode: "present" }).querySelector(".hs-media").dataset.motion, undefined, "an unknown photo motion stays still");
});

test("formulas: arithmetic on the conditions a, b, c, never code", async () => {
  const { E } = await loadEngine();
  assert.equal(E.evalFormula("a × b × c ÷ 100", { a: 120, b: 12, c: 60 }), 864);
  assert.equal(E.evalFormula("(a + b) * 2 - c / 4", { a: 1, b: 2, c: 8 }), 4);
  assert.equal(E.evalFormula("－a＋１０", { a: 3 }), 7, "full-width signs and digits work");
  assert.equal(E.evalFormula("-(a - b)", { a: 1, b: 3 }), 2);
  for (const bad of ["a +", "a b", "(a", "alert(1)", "a ** 2", "d + 1", "", "1/0"]) assert.ok(Number.isNaN(E.evalFormula(bad, { a: 1, b: 2, c: 3 })), bad);
  assert.equal(E.formulaTokens("a × 2").length, 3);
  assert.equal(E.formulaTokens("window"), null);
});

test("pages the audience works with: evidence, rankings, before/after, simulators, gaps and 'now'", async () => {
  const { E } = await loadEngine();
  const byTitle = (title) => allLayoutSlides.find((slide) => slide.title === title);
  const deck = deckOf(allLayoutSlides);
  const at = (slide) => allLayoutSlides.indexOf(slide);

  // The source of the figures sits at the foot of the page and goes with the chart tooltips.
  const rankSlide = byTitle("切り口で並び替え");
  const rank = E.render({ ...rankSlide, details: [{ target: "takeaway", text: "1人あたりでは営業", rows: [{ label: "営業", value: "8時間" }], source: "社内集計" }] }, { deck, index: at(rankSlide), mode: "present" });
  assert.equal(rank.dataset.source, "社内集計 2026年");
  assert.match(rank.querySelector(".hs-foot .hs-source").textContent, /^出所：社内集計 2026年$/);
  const row = (label) => rank.querySelector(`.hs-rank-row[data-label="${label}"]`);
  assert.equal(row("企画").style.getPropertyValue("--r"), "0", "the first view ranks by its own values");
  assert.equal(row("営業").style.getPropertyValue("--r"), "1");
  assert.ok(row("営業").classList.contains("is-hot"), "the highlighted item stays marked in every view");
  E.rankShow(rank.querySelector(".hs-rank"), 1);
  assert.equal(row("営業").style.getPropertyValue("--r"), "0", "switching the view re-sorts");
  assert.match(row("営業").dataset.tip, /営業：8時間（1位）/);
  assert.equal(rank.querySelectorAll(".hs-rank-views [data-view]").length, 2);
  assert.ok(rank.querySelector(".hs-rank-views").classList.contains("hs-control"));
  // Evidence for the key message: a badge beside the sentence, which stays editable on its own.
  const takeaway = rank.querySelector('.hs-takeaway-row[data-detail="takeaway"]');
  assert.ok(takeaway.querySelector(".hs-takeaway[data-field=takeaway]"));
  assert.equal(takeaway.querySelector(".hs-takeaway").textContent, rankSlide.takeaway);
  assert.match(takeaway.querySelector(".hs-detail-badge").textContent, /根拠/);

  const shiftSlide = byTitle("前後の差");
  const shift = E.render(shiftSlide, { deck, index: at(shiftSlide), mode: "present" });
  const rows = [...shift.querySelectorAll(".hs-shift-row")];
  assert.equal(rows.length, 2);
  assert.ok(rows[0].classList.contains("is-hot"), "the biggest change is the one in colour");
  assert.equal(rows[0].querySelector(".hs-shift-delta").textContent, "−55");
  assert.match(rows[0].querySelector(".hs-shift-bar").style.getPropertyValue("--from"), /%$/);

  const simSlide = byTitle("試算");
  const sim = E.render(simSlide, { deck, index: at(simSlide), mode: "present" });
  assert.equal(sim.querySelector(".hs-sim-out").textContent, "1,440");
  assert.match(sim.querySelector(".hs-sim-diff").textContent, /目標まで あと 1,560時間/);
  assert.equal(sim.querySelectorAll("input[type=range][data-sim]").length, 2);
  assert.match(sim.querySelector(".hs-sim-formula").textContent, /月の削減時間 ＝ 人数 × 1人あたり/);
  const range = sim.querySelector('input[data-sim="0"]');
  range.setAttribute("value", "300");
  range.value = "300";
  assert.equal(E.simUpdate(sim), 3600, "moving a condition recalculates the result");
  assert.match(sim.querySelector(".hs-sim-diff").textContent, /目標を 600時間 上回る/);
  assert.ok(sim.querySelector(".hs-sim").classList.contains("is-met"));

  const gapSlide = byTitle("不足と打ち手");
  const gap = E.render(gapSlide, { deck, index: at(gapSlide), mode: "present" });
  assert.equal(gap.dataset.build, "click", "measures come on one per click");
  assert.equal(gap.dataset.steps, "3");
  assert.equal(gap.querySelector(".hs-gap-num").textContent, "+80", "a finished slide shows every measure on");
  E.play(gap, { step: 0, animate: false });
  assert.equal(gap.querySelectorAll("[data-measure].is-on").length, 0);
  assert.equal(gap.querySelectorAll(".hs-hidden").length, 0, "nothing is hidden: the measures are switched, not revealed");
  assert.equal(gap.querySelector(".hs-gap-num").textContent, "1,560");
  E.reveal(gap, 1);
  assert.deepEqual([...gap.querySelectorAll("[data-measure].is-on")].map((el) => el.dataset.measure), ["0"]);
  assert.equal(gap.querySelector(".hs-gap-num").textContent, "940");
  E.play(gap, { step: Infinity, animate: false });
  assert.equal(gap.querySelector(".hs-gap-state").textContent, "目標を上回る");

  const gantt = allLayoutSlides.find((slide) => slide.type === "gantt");
  assert.equal(E.render(gantt, { deck, index: at(gantt), mode: "present" }).querySelector(".hs-gnow"), null, "no 'now' unless given");
  const now = E.render({ ...gantt, now: 1.5 }, { deck, index: at(gantt), mode: "present" }).querySelector(".hs-gnow");
  assert.equal(now.style.getPropertyValue("--at"), "1.5");
});

test("Lottie animations: a player box when shown, a badge in thumbnails", async () => {
  const { E } = await loadEngine();
  const slides = [
    { type: "title", title: "表紙" },
    { type: "cards", title: "動き", takeaway: "結論", items: [{ title: "A" }], media: { src: "idb:lot", kind: "lottie", name: "rocket.json", placement: { x: 0.5, y: 0.2, w: 0.3, h: 0.4 } } },
    { type: "closing" },
  ];
  const deck = deckOf(slides);
  const desc = E.mediaOf(slides[1]);
  assert.equal(desc.kind, "lottie");
  assert.equal(desc.fit, "contain", "animations are shown whole by default");
  assert.equal(E.mediaOf({ media: { src: "https://lottie.host/a/b.json" } }).kind, "lottie");
  const live = E.render(slides[1], { deck, index: 1, mode: "present", mediaUrls: { "idb:lot": "blob:lottie" } });
  const host = live.querySelector(".hs-placed.hs-lottie .hs-lottie-host");
  assert.equal(host.dataset.src, "blob:lottie");
  assert.ok(host.hasAttribute("data-autoplay") && host.hasAttribute("data-loop"));
  const thumb = E.render(slides[1], { deck, index: 1, mode: "thumb", mediaUrls: { "idb:lot": "blob:lottie" } });
  assert.equal(thumb.querySelector(".hs-lottie-host"), null);
  assert.match(thumb.querySelector(".hs-lottie-badge").textContent, /rocket/);
  assert.equal(typeof E.mountLottie, "function");
});

test("deck-wide design: accent colour, motion switches and theme fonts", async () => {
  const { E } = await loadEngine();
  const deck = deckOf(allLayoutSlides.slice(0, 3), { theme: "editorial", accent: "#123456", motion: { entrance: "blur", hover: "focus", numbers: false, ambient: false } });
  const el = E.render(deck.slides[2], { deck, index: 2, mode: "present" });
  assert.equal(el.style.getPropertyValue("--accent"), "#123456");
  assert.equal(el.dataset.entrance, "blur");
  assert.equal(el.dataset.hover, "focus");
  assert.equal(el.dataset.ambient, "off");
  assert.ok(!el.classList.contains("hs-numbers"));
  assert.match(E.fontHref(["editorial"]), /^https:\/\/fonts\.googleapis\.com\/css2\?family=Shippori\+Mincho\+B1/);
  const fitted = E.render(deck.slides[2], { deck, index: 2, mode: "thumb", fit: { fs: 0.85, ts: 0.9 } });
  assert.equal(fitted.style.getPropertyValue("--fs"), "0.85");
  assert.ok(fitted.classList.contains("hs-static"), "thumbnails never animate");
});

test("deep-dive pages: outside the story, numbered by their slide, opened from an item with an arrow", async () => {
  const { E } = await loadEngine();
  const slides = [
    { type: "title", title: "表紙" },
    { type: "stepUp", title: "発展段階", takeaway: "役割を広げてきた", items: [{ title: "規則", desc: "a" }, { title: "学習", desc: "b" }, { title: "生成", desc: "c" }], details: [{ target: "items[1]", text: "補足" }, { target: "items[2]", text: "生成の補足" }] },
    { type: "content", title: "学習の仕組み", takeaway: "例から傾向をつかむ", points: ["背景", "具体例"], drillOf: "items[1]" },
    { type: "cards", title: "選び方", takeaway: "3つの軸", items: [{ title: "用途" }, { title: "品質" }] },
    { type: "closing", message: "以上" },
  ];
  // The engine runs in its own context: compare plain copies of what it returns.
  const story = JSON.parse(JSON.stringify(E.storyMap(slides)));
  assert.deepEqual(story.order, [0, 1, 3, 4]);
  assert.deepEqual(story.parent, { 2: 1 });
  assert.deepEqual(story.drills[1], [{ index: 2, target: "items[1]" }]);
  assert.deepEqual([story.no[1], story.no[2], story.no[3]], [2, 2, 3]);
  assert.deepEqual([...E.storyMap([{ type: "title", drillOf: "items[0]" }, { type: "closing" }]).order], [0, 1], "the cover is never a deep-dive page");

  const deck = deckOf(slides);
  const parent = E.render(slides[1], { deck, index: 1, mode: "present" });
  assert.equal(parent.querySelector(".hs-page").textContent, "02 / 04", "numbers count the story only");
  const opener = [...parent.querySelectorAll('[data-item="items[1]"]')].find((el) => el.dataset.drill);
  assert.equal(opener?.dataset.drill, "2");
  assert.equal(parent.querySelectorAll(".hs-drill-badge").length, 1);
  assert.equal(parent.querySelector('[data-item="items[1]"] .hs-detail-badge'), null, "the arrow replaces the details mark on that item");
  assert.ok(parent.querySelector('[data-item="items[2]"] .hs-detail-badge'), "other items keep their details");

  const page = E.render(slides[2], { deck, index: 2, mode: "present" });
  assert.ok(page.classList.contains("hs-drill"));
  assert.equal(page.dataset.drillOf, "1");
  assert.equal(page.querySelector(".hs-page").textContent, "02 ・ 深掘り");
  assert.match(page.querySelector(".hs-eyebrow").textContent, /↳ 発展段階/);
  assert.equal(page.querySelectorAll("[data-drill]").length, 0, "one level: a deep-dive page opens nothing further");
  assert.equal(E.render(slides[3], { deck, index: 3, mode: "thumb" }).querySelector(".hs-page").textContent, "03 / 04");
});

test("clickable items say so: labelled marks on the text line, not clipped inside a shape", async () => {
  const { E } = await loadEngine();
  const pyramid = { type: "pyramid", title: "支える層", takeaway: "運用で差がつく", levels: [{ title: "改善", description: "a" }, { title: "設計", description: "b" }, { title: "ルール", description: "c" }], details: [{ target: "levels[2]", title: "範囲", text: "入力してよい情報" }] };
  const slides = [{ type: "title", title: "表紙" }, pyramid, { type: "content", title: "設計の中身", takeaway: "担当を決める", points: ["a"], drillOf: "levels[1]" }, { type: "closing", message: "以上" }];
  const el = E.render(pyramid, { deck: deckOf(slides), index: 1, mode: "present" });
  const detail = el.querySelector(".hs-detail-badge");
  assert.equal(detail.parentElement.tagName, "LI", "the mark sits on the text line, not on the clipped pyramid layer");
  assert.equal(detail.textContent, "詳しく");
  const drill = el.querySelector(".hs-drill-badge");
  assert.equal(drill.parentElement.tagName, "LI");
  assert.equal(drill.textContent, "深掘り");
});
