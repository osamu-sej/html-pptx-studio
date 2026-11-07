// What only HTML can do: objects that answer the mouse (hover, notes), clicks that open details, zoom, turn a card
// over, show other objects or put a spotlight on one, objects that keep moving, the HTML-only animation effects,
// PowerPoint charts that answer the mouse, and おまかせ (public/editor/htmlfx.mjs) putting them on a slide.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { chromeOf, describeAdded, enhanceSlide } from "../public/editor/htmlfx.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));

async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  for (const name of ["engine.js", "objects.js", "animate.js", "motion.js"]) {
    const code = await readFile(join(root, "public", "engine", name), "utf8");
    vm.runInContext(name === "engine.js" ? code.replace("/*__ICONS__*/{}", () => icons) : code, context, { filename: name });
  }
  return { E: window.SlideEngine, window };
}

const plain = (value) => JSON.parse(JSON.stringify(value));
const deckWith = (slide) => ({ title: "検証", theme: "clarity", transition: "fade", motion: {}, slides: [{ type: "title", title: "表紙" }, slide, { type: "closing" }] });
const box = (id, extra = {}) => ({ id, kind: "shape", shape: "roundRect", x: 200, y: 300, w: 400, h: 200, fill: "#dce4f2", text: `<p>${id}</p>`, ...extra });

test("hover, notes and loops are kept only when they are known", async () => {
  const { E } = await loadEngine();
  const [a, b] = E.normalizeObjects([
    box("a", { hover: "lift", loop: "float", tip: "  マウスを乗せると出る説明  " }),
    box("b", { hover: "explode", loop: 3, tip: "   " }),
  ]);
  assert.deepEqual([a.hover, a.loop, a.tip], ["lift", "float", "マウスを乗せると出る説明"]);
  assert.deepEqual([b.hover, b.loop, b.tip], [undefined, undefined, undefined]);
  assert.equal(E.normalizeObjects([box("c", { tip: "あ".repeat(500) })])[0].tip.length, 200, "a note is short");
  for (const list of [E.IX_HOVERS, E.IX_LOOPS, E.IX_CLICKS]) assert.ok(Object.values(list).every((label) => typeof label === "string" && label));
});

test("click actions: details, zoom, spotlight, turning over and showing other objects", async () => {
  const { E } = await loadEngine();
  const list = E.normalizeObjects([
    box("pop", { action: { type: "popup", title: " 内訳 ", text: '<p>本文<script>alert(1)</script></p>', rows: [...Array(12)].map((_, i) => ({ label: `項目${i}`, value: i })), source: "社内調査" } }),
    box("empty", { action: { type: "popup", text: "<p> </p>" } }),
    box("zoom", { action: { type: "zoom" } }),
    box("spot", { action: { type: "spot" } }),
    box("flip", { action: { type: "flip", back: "<p>裏の答え</p>", fill: "#F5F0EA" } }),
    box("flip2", { action: { type: "flip", back: "" } }),
    box("tab", { action: { type: "reveal", targets: ["pop", "ghost", "tab", "pop", "zoom"], only: true } }),
    box("none", { action: { type: "reveal", targets: ["ghost"] } }),
  ]);
  const by = Object.fromEntries(list.map((o) => [o.id, o]));
  const pop = by.pop.action;
  assert.equal(pop.title, "内訳");
  assert.doesNotMatch(pop.text, /script/, "the words are sanitized");
  assert.equal(pop.rows.length, 8, "eight rows at most");
  assert.deepEqual(plain(pop.rows[1]), { label: "項目1", value: "1" });
  assert.equal(pop.source, "社内調査");
  assert.equal(by.empty.action, undefined, "a card with nothing to say is dropped");
  assert.deepEqual(plain([by.zoom.action, by.spot.action]), [{ type: "zoom" }, { type: "spot" }]);
  assert.deepEqual(plain(by.flip.action), { type: "flip", back: "<p>裏の答え</p>", fill: "#f5f0ea" });
  assert.match(by.flip2.action.back, /裏の文字/, "a card always has a back");
  assert.deepEqual(plain(by.tab.action), { type: "reveal", targets: ["pop", "zoom"], only: true }, "only objects on the slide, never itself");
  assert.equal(by.none.action, undefined, "nothing to show, no action");
});

test("presenting: the objects carry their interactions; the editor draws them still", async () => {
  const { E } = await loadEngine();
  const elements = E.normalizeObjects([
    box("card", { hover: "lift", tip: "説明", group: "g1" }),
    { id: "mark", kind: "text", x: 220, y: 320, w: 100, h: 60, text: "<p>A</p>", loop: "pulse", group: "g1" },
    box("pop", { x: 700, action: { type: "popup", title: "内訳", text: "<p><b>太字</b>の説明</p>", rows: [{ label: "A", value: "1" }] } }),
    box("flip", { x: 1200, action: { type: "flip", back: "<p>裏の答え</p>" } }),
    box("tab", { y: 700, action: { type: "reveal", targets: ["panel"] } }),
    box("panel", { x: 700, y: 700 }),
  ]);
  const slide = { type: "blank", title: "検証", elements };
  const live = E.render(slide, { deck: deckWith(slide), index: 1, mode: "present" });
  const obj = (id) => live.querySelector(`.hs-obj[data-el="${id}"]`);
  assert.equal(obj("card").dataset.hover, "lift");
  assert.equal(obj("card").dataset.tip, "説明");
  assert.equal(obj("mark").dataset.loop, "pulse");
  assert.ok(obj("mark").style.getPropertyValue("--ix-ox"), "a group's loop turns about the group's centre");
  assert.equal(obj("pop").dataset.action, "popup");
  assert.equal(obj("pop").dataset.detail, "obj:pop", "the player opens the card like a layout's details");
  assert.ok(obj("flip").querySelector(".hs-obj-flip > .hs-obj-back"), "a card that turns over has a back");
  assert.match(obj("flip").querySelector(".hs-obj-back").textContent, /裏の答え/);
  assert.ok(obj("panel").classList.contains("hs-ix-wait"), "what a reveal button shows waits hidden");
  assert.ok(!obj("tab").classList.contains("hs-ix-wait"));
  const edit = E.render(slide, { deck: deckWith(slide), index: 1, mode: "edit" });
  assert.equal(edit.querySelectorAll("[data-hover], [data-loop], [data-tip], .hs-ix-wait").length, 0, "the editor shows everything, still");
  const [detail] = E.objectDetails(slide);
  assert.deepEqual(plain(detail), { target: "obj:pop", title: "内訳", text: "<p><b>太字</b>の説明</p>", html: true, rows: [{ label: "A", value: "1" }], source: "" });
  const nodes = E.richNodes("<p><b>太字</b><img src=x onerror=alert(1)></p>");
  const holder = live.ownerDocument.createElement("div");
  holder.append(nodes);
  assert.match(holder.innerHTML, /<b>太字<\/b>/);
  assert.doesNotMatch(holder.innerHTML, /onerror|<img/, "the card's words pass the sanitizer");
});

test("a PowerPoint chart answers the mouse: every mark has a note, and its legend shows or hides a series", async () => {
  const { E } = await loadEngine();
  const chart = { type: "bar", labels: ["A店", "B店"], series: [{ name: "今年", values: [42, 30] }, { name: "昨年", values: [35, 28] }],
    style: { type: "bar", dir: "col", legend: { pos: "b" }, series: [{ color: "#1f3864" }, { color: "#b7c3da" }] } };
  const elements = E.normalizeObjects([{ id: "ch", kind: "chart", x: 200, y: 200, w: 1000, h: 600, chart }]);
  const slide = { type: "blank", title: "検証", elements };
  const live = E.render(slide, { deck: deckWith(slide), index: 1, mode: "present" });
  const bars = [...live.querySelectorAll(".hs-obar.hs-mark")];
  assert.equal(bars.length, 4);
  assert.ok(bars.every((b) => b.getAttribute("data-tip") && b.getAttribute("data-s") != null));
  assert.match(bars[0].getAttribute("data-tip"), /A店・今年：42/);
  assert.deepEqual([...live.querySelectorAll(".hs-ochart-key[data-series]")].map((k) => k.getAttribute("data-series")), ["0", "1"]);
});

test("the HTML-only effects play on the timeline; those that cannot play backwards have no exit", async () => {
  const { E } = await loadEngine();
  const elements = [box("t"), { id: "ch", kind: "chart", x: 0, y: 0, w: 400, h: 300, chart: { type: "bar", labels: ["A"], series: [{ name: "s", values: [1] }] } }];
  const list = E.normalizeTimeline([
    { el: "t", cls: "in", fx: "typewriter", by: "char" },
    { el: "t", cls: "in", fx: "maskRise" },
    { el: "t", cls: "in", fx: "countUp" },
    { el: "t", cls: "out", fx: "countUp" },
    { el: "t", cls: "out", fx: "decode" },
    { el: "t", cls: "out", fx: "blurIn" },
    { el: "ch", cls: "in", fx: "chartGrow" },
    { el: "t", cls: "em", fx: "shine" },
    { el: "t", cls: "em", fx: "spotlight" },
    { el: "t", cls: "em", fx: "marker", color: "#FFE699" },
  ], { type: "blank", elements });
  assert.deepEqual(plain(list.map((e) => `${e.cls}:${e.fx}`)), ["in:typewriter", "in:maskRise", "in:countUp", "out:blurIn", "in:chartGrow", "em:shine", "em:spotlight", "em:marker"]);
  assert.equal(list[0].by, undefined, "an HTML effect works on the whole object");
  for (const [cls, fx] of [["in", "typewriter"], ["in", "decode"], ["in", "draw"], ["em", "ripple"], ["out", "blurIn"]]) assert.ok(E.animIsHtml(cls, fx), `${cls}:${fx}`);
  assert.ok(!E.animIsHtml("in", "fade") && !E.animIsHtml("em", "pulse"), "PowerPoint's own effects are not");
});

const E0 = { richToText: (html) => String(html).replace(/<[^>]+>/g, " ") };

test("おまかせ puts HTML's moves on a slide by what its objects are, and keeps what the slide has", () => {
  const slide = {
    type: "blank", title: "改革の全体像",
    elements: [
      { id: "ttl", kind: "text", x: 80, y: 40, w: 1400, h: 100, fs: 56, text: "<p>改革の全体像</p>" },
      { id: "c1", kind: "shape", x: 100, y: 300, w: 400, h: 200, fill: "#dce4f2", text: "<p>現状の課題</p>", group: "g1" },
      { id: "c1i", kind: "icon", x: 120, y: 320, w: 60, h: 60, icon: "check", group: "g1" },
      { id: "c2", kind: "shape", x: 600, y: 300, w: 400, h: 200, fill: "#dce4f2", text: "<p>打ち手</p>" },
      { id: "num", kind: "shape", x: 1100, y: 300, w: 400, h: 200, fill: "none", fs: 96, text: "<p>30分</p>" },
      { id: "ln", kind: "line", x1: 100, y1: 600, x2: 900, y2: 600 },
      { id: "ch", kind: "chart", x: 100, y: 650, w: 800, h: 380, chart: {} },
      { id: "pic", kind: "image", x: 1100, y: 600, w: 600, h: 400, src: "data:image/png;base64,AAA" },
      { id: "own", kind: "shape", x: 1100, y: 100, w: 300, h: 120, fill: "#f5f0ea", text: "<p>自前</p>", action: { type: "next" } },
    ],
    timeline: [{ id: "k1", el: "c2", cls: "in", fx: "flyIn", start: "click" }],
  };
  const before = JSON.stringify(slide);
  const { elements, timeline, added, count } = enhanceSlide(slide, E0);
  assert.equal(JSON.stringify(slide), before, "the slide itself is not changed");
  const fxOf = (id) => timeline.filter((e) => e.el === id).map((e) => e.fx);
  assert.deepEqual(fxOf("ttl"), ["maskRise"], "the title rises out of a mask");
  assert.deepEqual(fxOf("ch"), ["chartGrow"]);
  assert.deepEqual(fxOf("num"), ["countUp"], "a big figure counts up");
  assert.deepEqual(fxOf("ln"), ["draw"]);
  assert.deepEqual(fxOf("c2"), ["flyIn"], "the slide's own animation stays and nothing is added to it");
  assert.equal(timeline.at(-1).id, "k1", "the slide's own clicks come after what plays as it arrives");
  assert.equal(timeline[0].start, "after");
  assert.ok(timeline.slice(1, -1).every((e) => e.start === "with"));
  const obj = Object.fromEntries(elements.map((o) => [o.id, o]));
  assert.deepEqual([obj.c1.hover, obj.c1i.hover, obj.c2.hover], ["lift", "lift", "lift"], "cards lift with what is grouped with them");
  assert.equal(obj.own.hover, undefined, "an object with its own click stays as it is");
  assert.deepEqual(plain(obj.pic.action), { type: "zoom" });
  assert.equal(obj.pic.hover, "zoom");
  assert.deepEqual(added, { title: 1, charts: 1, numbers: 1, lines: 1, cards: 2, pictures: 1 });
  assert.equal(count, 7);
  assert.equal(describeAdded(added), "タイトル・グラフ・数字・線・カード2・写真");
  // Run again: nothing more to add.
  assert.equal(enhanceSlide({ ...slide, elements, timeline }, E0).count, 0);
});

test("おまかせ leaves the template's furniture alone: the same object in the same place on slide after slide", () => {
  const logo = { id: "logo", kind: "image", x: 1700, y: 20, w: 180, h: 140, src: "data:image/png;base64,LOGO" };
  const secret = { id: "sec", kind: "text", x: 20, y: 20, w: 300, h: 60, fs: 60, text: "<p>秘（B）2024</p>" };
  const page = (n) => ({ type: "blank", title: `ページ${n}`, elements: [
    { ...logo, id: `logo${n}` }, { ...secret, id: `sec${n}` },
    { id: `pic${n}`, kind: "image", x: 300, y: 300, w: 600, h: 400, src: `data:image/png;base64,P${n}` },
  ] });
  const slides = [page(1), page(2), page(3), { ...page(4), hidden: true }];
  const chrome = chromeOf(slides);
  assert.equal(chrome.size, 2, "the logo and 秘 are furniture; each slide's own picture is not");
  const { elements, timeline, added } = enhanceSlide(slides[0], E0, { chrome });
  assert.deepEqual(added.pictures, 1);
  assert.deepEqual(elements.filter((o) => o.action).map((o) => o.id), ["pic1"]);
  assert.ok(!timeline.some((e) => e.el === "sec1"), "the big 秘 does not count up");
  assert.equal(chromeOf([page(1)]).size, 0, "one slide has no furniture");
  assert.equal(chromeOf([page(1), { ...page(2), hidden: true }]).size, 0, "hidden slides do not count");
});

test("動作設定: a click plays a sound with its jump, マウスの通過 jumps or plays a sound (never opens a link)", async () => {
  const { E } = await loadEngine();
  const [a, b, c, d, e] = E.normalizeObjects([
    box("a", { action: { type: "next", sound: "chime" }, overAction: { type: "slide", to: "s3", sound: "nope" } }),
    box("b", { action: { type: "sound", sound: "applause" }, overAction: { type: "url", href: "https://example.com/" } }),
    box("c", { action: { type: "sound" }, overAction: { type: "sound", sound: "bell" } }),
    box("d", { action: { type: "url", href: "https://example.com/", sound: "click" }, overAction: { type: "end" } }),
    box("e", { overAction: "next" }),
  ]);
  assert.deepEqual(plain(a.action), { type: "next", sound: "chime" });
  assert.deepEqual(plain(a.overAction), { type: "slide", to: "s3" }, "an unknown sound is left out");
  assert.deepEqual(plain(b.action), { type: "sound", sound: "applause" });
  assert.equal(b.overAction, undefined, "a browser opens a page only for a click");
  assert.equal(c.action, undefined, "a sound action needs a sound");
  assert.deepEqual(plain(c.overAction), { type: "sound", sound: "bell" });
  assert.deepEqual(plain(d.action), { type: "url", href: "https://example.com/", sound: "click" });
  assert.deepEqual(plain(d.overAction), { type: "end" });
  assert.equal(e.overAction, undefined);
  const slide = { type: "blank", title: "検証", elements: [a, c] };
  const live = E.render(slide, { deck: deckWith(slide), index: 1, mode: "present" });
  assert.equal(live.querySelector('.hs-obj[data-el="a"]').dataset.over, "slide");
  assert.equal(live.querySelector('.hs-obj[data-el="c"]').dataset.over, "sound");
  const edit = E.render(slide, { deck: deckWith(slide), index: 1, mode: "edit" });
  assert.equal(edit.querySelectorAll("[data-over]").length, 0, "the editor does not act on the pointer");
});
