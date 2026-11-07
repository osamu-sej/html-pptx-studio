import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { deckShape } from "../server/schemas.mjs";
import { withoutImageData } from "../server/codex-app-server.mjs";

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
const objects = [
  { id: "box1", kind: "shape", shape: "roundRect", x: 200, y: 300, w: 400, h: 200, text: "<p>売上</p>" },
  { id: "txt1", kind: "text", x: 900, y: 300, w: 500, h: 120, text: "<p>一言</p>", group: "g1" },
  { id: "pic1", kind: "image", x: 900, y: 500, w: 300, h: 200, src: "asset:office", group: "g1" },
  { id: "vid1", kind: "video", x: 1300, y: 600, w: 400, h: 225, src: "https://example.com/a.mp4" },
  { id: "btn1", kind: "shape", shape: "rect", x: 100, y: 800, w: 200, h: 80, text: "<p>押す</p>" },
];

test("the catalogue matches PowerPoint's: entrances and their exits, emphasis, motion paths, media", async () => {
  const { E } = await loadEngine();
  assert.ok(Object.keys(E.ANIM_IN).length >= 18, "entrances");
  assert.ok(Object.keys(E.ANIM_EM).length >= 16, "emphasis");
  assert.ok(Object.keys(E.ANIM_PATHS).length >= 20, "motion paths");
  assert.deepEqual(Object.keys(E.ANIM_MEDIA), ["play", "pause", "stop"]);
  for (const [key, def] of Object.entries(E.ANIM_IN)) {
    // HTML-only effects that cannot play backwards (counting, decoding, a chart growing) have no exit.
    if (def.noExit) assert.ok(def.label && def.html && !def.out, `${key}: an HTML entrance without an exit`);
    else assert.ok(def.label && def.out, `${key} has an entrance and an exit name`);
    if (def.dirs) assert.ok(def.dirs.some(([k]) => k === def.dir), `${key}'s default direction is one of its own`);
    if (def.outDirs) assert.deepEqual(def.outDirs.map(([k]) => k), def.dirs.map(([k]) => k), `${key}: the exit names the same directions`);
  }
  assert.equal(E.animLabel("out", "flyIn"), "スライドアウト");
  assert.equal(E.animLabel("in", "flyIn"), "スライドイン");
  assert.deepEqual(plain(E.animDirs("out", "flyIn")[0]), ["bottom", "下へ"]);
});

test("a timeline keeps only what it can play, with sensible numbers", async () => {
  const { E } = await loadEngine();
  const slide = { type: "blank", elements: objects };
  const list = E.normalizeTimeline([
    { id: "a1", el: "box1", cls: "in", fx: "flyIn", dir: "left", dur: 800 },
    { id: "a1", el: "box1", cls: "em", fx: "spin", amount: 99999, dur: -5 },
    { el: "nobody", cls: "in", fx: "fade" },
    { el: "box1", cls: "in", fx: "explode" },
    { el: "box1", cls: "wobble", fx: "fade" },
    { el: "box1", cls: "media", fx: "play" },
    { el: "vid1", cls: "media", fx: "play" },
    { el: "grp:g1", cls: "in", fx: "wipe", dir: "sideways", start: "after", delay: 500 },
    { el: "grp:nope", cls: "in", fx: "fade" },
    { el: "@title", cls: "in", fx: "fade" },
    { el: "@g2", cls: "out", fx: "zoom" },
    { el: "@body", cls: "in", fx: "fade" },
    { el: "box1", cls: "path", fx: "custom", path: { pts: [[0, 0]] } },
    { el: "box1", cls: "path", fx: "lineDown", path: { pts: [[0, 0], [0, 300], ["x", 1]], curve: true } },
    { el: "txt1", cls: "in", fx: "fade", by: "char", trigger: "btn1", repeat: 3, rewind: true, ease: "bounce", color: "#zzz" },
    { el: "pic1", cls: "in", fx: "fade", by: "char", trigger: "ghost" },
  ], slide);
  assert.equal(list.length, 9, JSON.stringify(list.map((e) => `${e.el}:${e.fx}`)));
  assert.equal(new Set(list.map((e) => e.id)).size, list.length, "ids are unique");
  assert.deepEqual(plain(list[0]), { id: "a1", el: "box1", cls: "in", fx: "flyIn", start: "click", dur: 800, delay: 0, dir: "left" });
  assert.equal(list[1].amount, 3600, "spin is capped");
  assert.equal(list[1].dur, 1, "a duration is at least 1 ms");
  assert.equal(list[2].el, "vid1", "media effects only on videos and animations");
  const group = list.find((e) => e.el === "grp:g1");
  assert.equal(group.dir, "bottom", "an unknown direction falls back to the effect's default");
  assert.equal(group.start, "after");
  assert.equal(group.delay, 500);
  assert.ok(list.some((e) => e.el === "@title") && list.some((e) => e.el === "@g2"), "layout parts can be animated");
  const path = list.find((e) => e.cls === "path");
  assert.deepEqual(plain(path.path), { pts: [[0, 0], [0, 300]], curve: true });
  const text = list.find((e) => e.el === "txt1");
  assert.equal(text.by, "char");
  assert.equal(text.trigger, "btn1");
  assert.equal(text.repeat, 3);
  assert.equal(text.rewind, true);
  assert.equal(text.ease, "bounce");
  const pic = list.find((e) => e.el === "pic1");
  assert.equal(pic.by, undefined, "pictures have no words");
  assert.equal(pic.trigger, undefined, "a trigger must be an object on the slide");
});

test("clicks, with and after: groups and their timing, and triggers on their own", async () => {
  const { E } = await loadEngine();
  const slide = {
    type: "blank", elements: objects,
    timeline: [
      { id: "t0", el: "box1", cls: "in", fx: "fade", start: "with", dur: 500 },
      { id: "t1", el: "txt1", cls: "in", fx: "fade", start: "after", dur: 400, delay: 100 },
      { id: "c1", el: "box1", cls: "em", fx: "pulse", start: "click", dur: 500 },
      { id: "c1b", el: "txt1", cls: "em", fx: "spin", start: "with", dur: 1000, delay: 200 },
      { id: "c1c", el: "pic1", cls: "in", fx: "zoom", start: "after", dur: 500 },
      { id: "c2", el: "box1", cls: "out", fx: "fade", start: "click", dur: 500, repeat: 2 },
      { id: "tr1", el: "pic1", cls: "em", fx: "pulse", start: "with", trigger: "btn1" },
      { id: "tr2", el: "pic1", cls: "em", fx: "teeter", start: "click", trigger: "btn1" },
    ],
  };
  const plan = E.timelinePlan(slide);
  assert.equal(plan.clicks, 2);
  assert.deepEqual(plain(plan.main.map((g) => g.items.map((i) => i.e.id))), [["t0", "t1"], ["c1", "c1b", "c1c"], ["c2"]]);
  assert.deepEqual(plain(plan.main[0].items.map((i) => [i.begin, i.end])), [[0, 500], [600, 1000]]);
  // "with" starts with the previous one (plus its own delay); "after" waits for the previous one to end.
  assert.deepEqual(plain(plan.main[1].items.map((i) => [i.begin, i.end])), [[0, 500], [200, 1200], [1200, 1700]]);
  assert.equal(plan.main[1].total, 1700);
  assert.equal(plan.main[2].total, 1000, "repeats count");
  assert.deepEqual(plain([...plan.triggers.keys()]), ["btn1"]);
  assert.deepEqual(plain(plan.triggers.get("btn1").map((g) => g.items.map((i) => i.e.id))), [["tr1"], ["tr2"]]);
  assert.equal(E.timelineTakesLayout(slide), false);
  assert.equal(E.timelineTakesLayout({ timeline: [{ el: "@g0" }] }), true);
  assert.equal(E.timelineTakesLayout({ timeline: [{ el: "@title" }] }), false, "the title does not stop the layout's build");
});

test("a rendered slide counts the animations' clicks after the layout's and hides what enters later", async () => {
  const { E } = await loadEngine();
  const slide = {
    type: "content", title: "三つの柱", takeaway: "柱は三つ", points: ["一つ目", "二つ目", "三つ目"], animation: "click", elements: objects,
    timeline: [
      { el: "box1", cls: "in", fx: "fade", start: "click" },
      { el: "txt1", cls: "em", fx: "pulse", start: "click" },
      { el: "pic1", cls: "in", fx: "wipe", start: "click", trigger: "btn1" },
    ],
  };
  const deck = { title: "検証", theme: "clarity", motion: {}, slides: [{ type: "title", title: "表紙" }, slide, { type: "closing" }] };
  const live = E.render(slide, { deck, index: 1, mode: "present" });
  assert.equal(live.dataset.lsteps, "3", "the layout's three items first");
  assert.equal(live.dataset.steps, "5", "then the two clicks of the animations");
  const fx = (id) => live.querySelector(`.hs-obj[data-el="${id}"] .hs-obj-fx`);
  assert.ok(fx("box1").classList.contains("hs-anim-hide"), "an object that enters later starts hidden");
  assert.ok(fx("pic1").classList.contains("hs-anim-hide"), "so does one a trigger brings in");
  assert.ok(!fx("txt1").classList.contains("hs-anim-hide"), "emphasis leaves it on the slide");
  assert.ok(live.querySelector('.hs-obj[data-el="btn1"]').hasAttribute("data-trigger"), "the trigger can be clicked");
  const edit = E.render(slide, { deck, index: 1, mode: "edit" });
  assert.equal(edit.querySelectorAll(".hs-anim-hide").length, 0, "the editor shows everything");
  // Animating the layout's own items takes over from its click build.
  const takeover = { ...slide, timeline: [{ el: "@g1", cls: "in", fx: "fade", start: "click" }] };
  const el = E.render(takeover, { deck, index: 1, mode: "present" });
  assert.equal(el.dataset.build, "none");
  assert.equal(el.dataset.lsteps, "0");
  assert.equal(el.dataset.steps, "1");
  assert.ok(el.querySelectorAll("[data-g]").length >= 3, "the items keep their numbers");
  assert.ok([...el.querySelectorAll('[data-g="1"]')].every((node) => node.classList.contains("hs-anim-hide")));
  assert.equal(el.querySelectorAll(".hs-hidden").length, 0, "no build hides anything");
  assert.deepEqual(plain(E.layoutTargets(el).map((t) => t.el)).slice(0, 3), ["@title", "@takeaway", "@g0"]);
});

test("motion paths: presets fit the object, points come evenly along the curve, closed paths return", async () => {
  const { E } = await loadEngine();
  const down = E.pathPreset("lineDown", { w: 200, h: 100 });
  assert.deepEqual(plain(down), { pts: [[0, 0], [0, 280]] });
  const big = E.pathPreset("lineRight", { w: 2000, h: 100 });
  assert.equal(big.pts[1][0], 640, "a preset is never longer than 640 px");
  const pts = E.pathPoints({ pts: [[0, 0], [300, 0], [300, 300]] }, 7);
  assert.deepEqual(plain(pts[0]), [0, 0]);
  assert.deepEqual(plain(pts.at(-1)), [300, 300]);
  const steps = pts.slice(1).map((p, i) => Math.hypot(p[0] - pts[i][0], p[1] - pts[i][1]));
  assert.ok(Math.max(...steps) - Math.min(...steps) < 1, `even spacing: ${steps}`);
  const circle = E.pathPreset("circle", { w: 100, h: 100 });
  assert.equal(circle.closed, true);
  const around = E.pathPoints(circle, 33);
  assert.ok(Math.hypot(...around.at(-1)) < 1, "a closed path ends where it started");
  assert.ok(Math.max(...around.map((p) => p[1])) > 200, "and goes round");
  assert.deepEqual(plain(E.pathEnd(circle)), [0, 0]);
  assert.match(E.pathD(down), /^M0 0 L/);
});

test("the AI never writes animations: they stay out of its view and every change keeps them", () => {
  const timeline = [{ id: "a", el: "box1", cls: "in", fx: "fade" }, { id: "b", el: "box1", cls: "out", fx: "fade" }];
  const shown = withoutImageData({ slides: [{ type: "blank", title: "x", elements: objects, timeline, sid: "s1" }] });
  assert.equal(shown.slides[0].timeline, "[アニメーション2件]");
  assert.equal(shown.slides[0].sid, undefined);
  const parsed = deckShape.parse({ title: "t", purpose: "", audience: "", theme: "clarity", slides: [{ type: "title", title: "表紙" }, { type: "blank", title: "x", elements: objects, timeline }] });
  assert.equal(parsed.slides[1].timeline.length, 2, "the deck keeps them");
});

test("アニメーションの後 and an effect's sound: kept where they fit; a faded / hidden end after the effect, a hide at the next click", async () => {
  const { E, window } = await loadEngine();
  const slide = { type: "blank", elements: objects, timeline: [
    { id: "a1", el: "box1", cls: "in", fx: "fade", start: "click", after: "dim", sound: "chime" },
    { id: "a2", el: "txt1", cls: "in", fx: "fade", start: "click", after: "hideNext" },
    { id: "a3", el: "pic1", cls: "out", fx: "fade", start: "click", after: "dim", sound: "BAD!" },
    { id: "a4", el: "btn1", cls: "em", fx: "pulse", start: "click", after: "hide" },
    { id: "a5", el: "vid1", cls: "media", fx: "play", start: "click", after: "dim" },
  ] };
  const list = plain(E.normalizeTimeline(slide.timeline, slide));
  assert.equal(list[0].after, "dim");
  assert.equal(list[0].sound, "chime");
  assert.equal(list[1].after, "hideNext");
  assert.equal(list[2].after, undefined, "an exit has no after");
  assert.equal(list[2].sound, undefined, "only sound names");
  assert.equal(list[3].after, "hide");
  assert.equal(list[4].after, undefined, "media has no after");
  assert.ok(E.ANIM_AFTERS.dim && E.ANIM_AFTERS.hide && E.ANIM_AFTERS.hideNext);
  // Playing: seeking to the end of click 2 hides what was to go at the next click (the 2nd click came).
  const el = E.render(slide, { mode: "present", index: 0, deck: { slides: [slide], theme: "clarity" } });
  window.document.body.append(el);
  const fxOf = (id) => el.querySelector(`.hs-obj[data-el="${id}"] .hs-obj-fx`) || el.querySelector(`.hs-obj[data-el="${id}"]`);
  const seen = [];
  const animate = window.Element.prototype.animate;
  window.Element.prototype.animate = function (frames, opts) { seen.push([this.closest(".hs-obj")?.dataset.el, JSON.stringify(frames.at(-1))]); return animate ? animate.call(this, frames, opts) : { finished: Promise.resolve(), finish() {}, cancel() {}, effect: { updateTiming() {} } }; };
  try {
    assert.ok(fxOf("txt1"), "the text drawn");
    E.animSeek(el, 3);
    await new Promise((r) => setTimeout(r, 20));
    assert.ok(seen.some(([id, f]) => id === "txt1" && /hidden/.test(f)), `the text hidden when the next click came: ${JSON.stringify(seen.filter(([id]) => id === "txt1"))}`);
    assert.ok(seen.some(([id, f]) => id === "box1" && /0\.35/.test(f)), `the box faded after its entrance: ${JSON.stringify(seen.filter(([id]) => id === "box1"))}`);
  } finally { window.Element.prototype.animate = animate; }
});

test("ブックマーク時: a trigger on a media's bookmark is kept, plays as the media passes it, and the media stays a player", async () => {
  const { E, window } = await loadEngine();
  const slide = { type: "blank", elements: [
    E.normalizeObject({ id: "snd", kind: "audio", x: 0, y: 0, w: 120, h: 120, src: "https://example.com/a.mp3", bookmarks: [{ id: "m1", t: 2, name: "サビ" }, { t: 5 }] }),
    E.normalizeObject({ id: "box", kind: "shape", shape: "rect", x: 300, y: 100, w: 200, h: 100, fill: "#dce4f2" }),
  ], timeline: [
    { id: "a1", el: "box", cls: "in", fx: "fade", start: "click", trigger: "snd@m1" },
    { id: "a2", el: "box", cls: "em", fx: "pulse", start: "click", trigger: "snd@nope" },
    { id: "a3", el: "box", cls: "em", fx: "pulse", start: "click", trigger: "ghost@m1" },
  ] };
  const list = plain(E.normalizeTimeline(slide.timeline, slide));
  assert.equal(list[0].trigger, "snd@m1");
  assert.equal(list[1].trigger, undefined, "only a bookmark the media has");
  assert.equal(list[2].trigger, undefined, "only media on the slide");
  assert.deepEqual(plain([...E.timelinePlan(slide).triggers.keys()]), ["snd@m1"]);
  const el = E.render(slide, { mode: "present", index: 0, deck: { slides: [slide], theme: "clarity" } });
  window.document.body.append(el);
  assert.equal(el.querySelector('.hs-obj[data-el="snd"]').hasAttribute("data-trigger"), false, "a bookmark trigger leaves the sound a player");
  const audio = el.querySelector("audio");
  assert.deepEqual(JSON.parse(audio.getAttribute("data-bookmarks")).map((b) => b.t), [2, 5]);
  assert.equal(el.querySelectorAll(".hs-audio-track .hs-mark").length, 2, "a mark per bookmark on the bar");
  let now = 0;
  Object.defineProperty(audio, "currentTime", { get: () => now, set: (v) => { now = v; }, configurable: true });
  const seen = [];
  const animate = window.Element.prototype.animate;
  window.Element.prototype.animate = function (frames, opts) { seen.push(this.closest(".hs-obj")?.dataset.el); return animate ? animate.call(this, frames, opts) : { finished: Promise.resolve(), finish() {}, cancel() {}, effect: { updateTiming() {} } }; };
  try {
    E.animStart(el, 0);
    const tick = (t) => { now = t; audio.dispatchEvent(new window.Event("timeupdate")); };
    tick(0.5); tick(1.8);
    await new Promise((r) => setTimeout(r, 10));
    assert.equal(seen.filter((id) => id === "box").length, 0, "nothing before the bookmark");
    tick(2.1);
    await new Promise((r) => setTimeout(r, 10));
    assert.ok(seen.includes("box"), "passing the bookmark plays the entrance");
    const count = seen.length;
    now = 1; audio.dispatchEvent(new window.Event("seeked"));
    tick(4.5);
    await new Promise((r) => setTimeout(r, 10));
    assert.equal(seen.length, count, "a jump past the bookmark does not play it");
  } finally { window.Element.prototype.animate = animate; }
});
