// 画面切り替え「変形」: which parts of two slides are the same thing (objects.js morphPairs), and that the player carries them across.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  for (const file of ["objects.js", "animate.js", "motion.js"]) vm.runInContext(await readFile(join(root, "public", "engine", file), "utf8"), context, { filename: file });
  return { E: window.SlideEngine, window };
}
const plain = (v) => JSON.parse(JSON.stringify(v));
const shape = (id, extra = {}) => ({ id, kind: "shape", shape: "rect", x: 100, y: 100, w: 300, h: 200, ...extra });
const text = (id, words, extra = {}) => ({ id, kind: "text", x: 100, y: 400, w: 600, h: 80, text: `<p>${words}</p>`, ...extra });

test("morphPairs: a duplicated slide keeps its ids, so the parts moved on the copy are the same parts", async () => {
  const { E } = await loadEngine();
  const a = [shape("s1"), text("t1", "売上"), shape("gone")];
  const b = [shape("s1", { x: 900, y: 300, w: 500, rot: 20 }), text("t1", "売上は伸びた"), shape("new")];
  const pairs = E.morphPairs(a, b);
  assert.deepEqual(plain(pairs.map(([x, y]) => [x.id, y.id])), [["s1", "s1"], ["t1", "t1"]], "the same id and kind, whatever the words and place");
  assert.equal(pairs[0][0], a[0], "the objects themselves come back, not copies");
});

test("morphPairs: the same \"!!\" name is the same part even with another id", async () => {
  const { E } = await loadEngine();
  const a = [shape("x1", { name: "!!円" }), shape("x2", { name: "!!角" })];
  const b = [shape("y2", { name: "!!角" }), shape("y1", { name: "!!円" })];
  assert.deepEqual(plain(E.morphPairs(a, b).map(([x, y]) => [x.id, y.id]).sort()), [["x1", "y1"], ["x2", "y2"]]);
  // A name without the two exclamation marks is only a name; a "!!" name does not pair with an id alone.
  assert.deepEqual(plain(E.morphPairs([shape("k", { name: "!!A" })], [shape("k", { name: "!!B" })])), [], "different !! names are different parts");
  assert.equal(E.morphPairs([shape("k", { name: "題" })], [shape("k", { name: "別の名前" })]).length, 1, "an ordinary name does not matter");
});

test("morphPairs: the same words, picture or icon pair up when there is only one of each", async () => {
  const { E } = await loadEngine();
  const a = [text("a1", "目標"), { id: "p1", kind: "image", src: "idb:logo", x: 0, y: 0, w: 100, h: 100 }, { id: "i1", kind: "icon", icon: "store", x: 0, y: 0, w: 60, h: 60 }];
  const b = [text("b1", "目標"), { id: "p9", kind: "image", src: "idb:logo", x: 800, y: 0, w: 200, h: 200 }, { id: "i9", kind: "icon", icon: "store", x: 500, y: 0, w: 60, h: 60 }];
  assert.deepEqual(plain(E.morphPairs(a, b).map(([x, y]) => `${x.id}>${y.id}`).sort()), ["a1>b1", "i1>i9", "p1>p9"]);
  // Two with the same words leave no way to tell which is which.
  assert.deepEqual(plain(E.morphPairs([text("a1", "同じ"), text("a2", "同じ")], [text("b1", "同じ")])), []);
  // Nothing in common.
  assert.deepEqual(plain(E.morphPairs([text("a1", "前")], [text("b1", "後")])), []);
});

test("morphPairs: every part is used once, order of rules is name, id, look; what plays or hides is left out", async () => {
  const { E } = await loadEngine();
  const a = [shape("s", { name: "!!主役" }), text("t", "同じ")];
  const b = [text("t", "同じ", { name: "!!主役" }), shape("s")];
  // "!!主役" first: a.s ⇄ b.t; then ids: a.t ⇄ ... b.t is taken, b.s has no a.s left, so by words: a.t (同じ) has no partner either.
  const pairs = E.morphPairs(a, b).map(([x, y]) => `${x.id}>${y.id}`);
  assert.deepEqual(plain(pairs), ["s>t"]);
  assert.deepEqual(plain(E.morphPairs([{ id: "v", kind: "video", x: 0, y: 0, w: 10, h: 10 }, shape("h", { hidden: true })], [{ id: "v", kind: "video", x: 0, y: 0, w: 10, h: 10 }, shape("h", { hidden: true })])), [], "videos and hidden parts do not morph");
  assert.deepEqual(plain(E.morphPairs(null, undefined)), []);
  assert.deepEqual(plain(E.morphPairs([shape("s")], [{ kind: "shape" }])), [], "a part with no id is skipped");
});

test("morphIn: the leaving parts travel to the new ones in a layer over the new slide, then the layer goes", async () => {
  const { E, window } = await loadEngine();
  const slide = (elements) => E.render({ type: "blank", hideTitle: true, title: "", elements: E.normalizeObjects(elements) }, { mode: "present", index: 1 });
  const a = [shape("s1"), shape("old", { x: 1200 })];
  const b = [shape("s1", { x: 900, y: 500, w: 500, h: 100, rot: 30 }), shape("fresh", { x: 50 })];
  const prev = slide(a);
  const next = slide(b);
  window.document.body.append(prev, next);
  const pairs = E.morphPairs(E.normalizeObjects(a), E.normalizeObjects(b));
  assert.equal(pairs.length, 1);
  const calls = [];
  for (const el of [...prev.querySelectorAll(".hs-obj"), ...next.querySelectorAll(".hs-obj")]) el.animate = (frames, options) => { calls.push({ el, frames, options }); return { cancel() {} }; };
  const done = E.morphIn(prev, next, pairs, { dur: 100 });
  const layer = next.querySelector(":scope > .hs-morph-layer");
  assert.ok(layer, "a layer over the new slide");
  assert.equal(layer.querySelectorAll(".hs-obj").length, 2, "both the old parts are taken over");
  assert.equal(prev.querySelectorAll(":scope > .hs-objects > .hs-obj").length, 0, "the old slide has none left");
  assert.equal(layer.querySelector("[data-el]"), null, "they no longer answer to their ids");
  const mover = calls.find((c) => c.frames.some((f) => /translate/.test(f.transform || "")) && c.el.parentNode === layer);
  assert.ok(mover, "the old part moves");
  assert.match(mover.frames.at(-1).transform, /translate\(\d+(\.\d+)?px, \d+(\.\d+)?px\) rotate\(30deg\) scale\(1\.66/, `towards the new place: ${mover.frames.at(-1).transform}`);
  assert.equal(mover.frames.at(-1).opacity, 0, "and gives way to the new part");
  const arriving = calls.find((c) => c.el.dataset.el === "s1" && c.el.parentNode !== layer);
  assert.ok(arriving && /translate\(-/.test(arriving.frames[0].transform), "the new part starts from where the old one was");
  assert.equal(arriving.options.fill, "backwards");
  const fresh = calls.find((c) => c.el.dataset.el === "fresh");
  assert.deepEqual(plain(fresh.frames.map((f) => f.opacity)), [0, 1], "a part with no partner fades in");
  await done;
  assert.equal(next.querySelector(".hs-morph-layer"), null, "the layer is gone when it is over");
  assert.equal((await E.morphIn(null, next, [], {})), undefined, "nothing to do with no old slide");
});
