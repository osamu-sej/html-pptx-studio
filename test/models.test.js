// 3D モデル: the object as kept in a deck (its file, view and poster), how it is drawn before the live view, the
// 3D animations (models.js reads --m3d-*), and the views and built-in models models.js offers.
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
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  for (const name of ["engine.js", "objects.js", "animate.js", "motion.js", "models.js"]) {
    const code = await readFile(join(root, "public", "engine", name), "utf8");
    vm.runInContext(name === "engine.js" ? code.replace("/*__ICONS__*/{}", () => icons) : code, context, { filename: name });
  }
  return { E: window.SlideEngine, window };
}
const plain = (value) => JSON.parse(JSON.stringify(value));
const POSTER = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

test("a 3D model keeps its file (a GLB, a glTF, a built-in model), a tidy view and its poster", async () => {
  const { E } = await loadEngine();
  const base = { id: "m1", kind: "model", x: 100, y: 100, w: 400, h: 300 };
  const cube = E.normalizeObject({ ...base, src: "builtin:cube", view: { yaw: 400, pitch: -120, zoom: 9, panX: 0, roll: 0 }, poster: POSTER });
  assert.equal(cube.src, "builtin:cube");
  assert.deepEqual(plain(cube.view), { yaw: 360, pitch: -89, zoom: 4 }, "clamped; the defaults left out");
  assert.equal(cube.poster, POSTER);
  assert.equal(E.normalizeObject({ ...base, src: "data:model/gltf-binary;base64,Z2xURg==" }).src.slice(0, 26), "data:model/gltf-binary;bas");
  assert.ok(E.normalizeObject({ ...base, src: "idb:abc123" }));
  assert.equal(E.normalizeObject({ ...base, src: "javascript:alert(1)" }), null, "no other kind of address");
  assert.equal(E.normalizeObject({ ...base, src: "builtin:cube", poster: "data:text/html,<b>x</b>" }).poster, undefined, "a poster is a picture");
  assert.equal(E.normalizeObject({ ...base, src: "builtin:cube", view: { yaw: 0, zoom: 1 } }).view, undefined, "the front view needs nothing kept");
  assert.equal(E.KIND_LABELS?.model ?? E.objectName({ ...cube, name: undefined }, 0).replace(/\s*\d+$/, ""), "3D モデル");
});

test("before the live view, the slide shows the poster (or a placeholder) and carries what the drawing needs", async () => {
  const { E } = await loadEngine();
  const slide = { type: "blank", title: "", elements: [
    { id: "m1", kind: "model", x: 100, y: 100, w: 400, h: 300, src: "builtin:store", view: { yaw: 30 }, poster: POSTER, alt: "店舗" },
    { id: "m2", kind: "model", x: 600, y: 100, w: 300, h: 300, src: "builtin:cube" },
  ] };
  const el = E.render(slide, { mode: "present", index: 0, deck: { theme: "clarity", slides: [slide] } });
  const [a, b] = el.querySelectorAll(".hs-model");
  assert.equal(a.dataset.src, "builtin:store");
  assert.deepEqual(JSON.parse(a.dataset.view), { yaw: 30 });
  assert.equal(a.querySelector(".hs-model-poster").getAttribute("src"), POSTER);
  assert.equal(a.querySelector(".hs-model-poster").getAttribute("alt"), "店舗");
  assert.ok(b.querySelector(".hs-model-placeholder"), "no poster yet: a placeholder");
});

test("the 3D animations: 到着 / 退出, ターンテーブル, スイング, ジャンプしてターン — for 3D models, turning by --m3d-yaw", async () => {
  const { E } = await loadEngine();
  assert.equal(E.ANIM_IN.arrive3d.model, true);
  assert.equal(E.animLabel("in", "arrive3d"), "到着");
  assert.equal(E.animLabel("out", "arrive3d"), "退出");
  for (const fx of ["turntable3d", "swing3d", "jump3d"]) assert.equal(E.ANIM_EM[fx].model, true, fx);
  assert.match(JSON.stringify(E.ANIM_IN.arrive3d.kf({})), /--m3d-yaw/);
  // A deck's timeline keeps them.
  const slide = { type: "blank", elements: [{ id: "m1", kind: "model", x: 0, y: 0, w: 300, h: 300, src: "builtin:cube" }] };
  const timeline = E.normalizeTimeline([{ id: "a1", el: "m1", cls: "em", fx: "turntable3d", start: "click", amount: 720, dir: "ccw" }, { id: "a2", el: "m1", cls: "in", fx: "arrive3d", start: "click" }], slide);
  assert.deepEqual(plain(timeline.map((e) => e.fx)), ["turntable3d", "arrive3d"]);
  assert.equal(timeline[0].amount, 720);
});

test("models.js offers the 3D モデル ビュー presets and the built-in models, and loads three.js from where it is told", async () => {
  const { E } = await loadEngine();
  assert.deepEqual(plain(E.MODEL_VIEWS.front), ["正面", 0, 0]);
  assert.equal(E.MODEL_VIEWS.back[1], 180);
  assert.ok(Object.keys(E.MODEL_BUILTINS).length >= 8 && E.MODEL_BUILTINS.store === "店舗");
  assert.equal(E.threeUrls.three, "/vendor/three/three.module.js");
  assert.equal(typeof E.mountModels, "function");
  assert.equal(typeof E.modelPoster, "function");
});

test("おまかせ gives a 3D model its 到着 (it arrives turning) and says so", async () => {
  const { E } = await loadEngine();
  const { enhanceSlide, describeAdded } = await import("../public/editor/htmlfx.mjs");
  const slide = { type: "blank", elements: [{ id: "m1", kind: "model", x: 100, y: 100, w: 400, h: 400, src: "builtin:cube" }] };
  const { timeline, added } = enhanceSlide(slide, E);
  assert.deepEqual(plain(timeline.map((e) => [e.el, e.cls, e.fx])), [["m1", "in", "arrive3d"]]);
  assert.equal(describeAdded(added), "3D モデル");
});
