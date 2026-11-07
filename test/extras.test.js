// More of PowerPoint's 挿入 in the engine: スライド ズーム, カメオ, 数式, fields (slide number, date) and 段組み.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { objectSchema } from "../server/schemas.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  return { E: window.SlideEngine };
}
const deckOf = (slides) => ({ theme: "clarity", slides });

test("スライド ズーム shows the target slide; a missing target says so; it never draws itself again", async () => {
  const { E } = await loadEngine();
  const target = { type: "content", title: "ズーム先の結論", sid: "t1", points: ["a"] };
  const zoom = E.normalizeObject({ id: "z", kind: "zoom", x: 100, y: 100, w: 480, h: 270, target: "t1" });
  assert.equal(zoom.target, "t1");
  assert.equal(E.normalizeObject({ id: "z2", kind: "zoom", x: 0, y: 0, w: 10, h: 10, target: "<bad>" }), null);
  const home = { type: "blank", sid: "h1", elements: [zoom, { ...zoom, id: "self", target: "h1" }] };
  const deck = deckOf([{ type: "title", title: "表紙", sid: "c" }, home, target]);
  const el = E.render(home, { deck, index: 1, mode: "present" });
  const z = el.querySelector('.hs-obj[data-el="z"] .hs-zoom');
  assert.ok(z.querySelector(".hs-slide"), "the target drawn inside");
  assert.ok(z.textContent.includes("ズーム先の結論"));
  const missing = E.render({ ...home, elements: [{ ...zoom, target: "nope" }] }, { deck, index: 1, mode: "edit" });
  assert.match(missing.textContent, /リンク先のスライドがありません/);
  assert.ok(objectSchema.safeParse({ id: "z", kind: "zoom" }).success && objectSchema.safeParse({ id: "c", kind: "camera" }).success && objectSchema.safeParse({ id: "e", kind: "equation" }).success);
});

test("カメオ: a placeholder while editing, a video while presenting, in a circle", async () => {
  const { E } = await loadEngine();
  const cam = E.normalizeObject({ id: "c", kind: "camera", x: 1500, y: 700, w: 300, h: 300, mask: "ellipse" });
  const slide = { type: "blank", elements: [cam] };
  const edit = E.render(slide, { deck: deckOf([slide]), index: 0, mode: "edit" });
  assert.ok(edit.querySelector(".hs-camera-placeholder") && !edit.querySelector("video"));
  const live = E.render(slide, { deck: deckOf([slide]), index: 0, mode: "present" });
  assert.ok(live.querySelector(".hs-camera video.hs-camera-feed"));
  assert.match(live.querySelector(".hs-camera").getAttribute("style"), /border-radius: ?50%/);
});

test("数式: TeX becomes MathML (fractions, roots, limits, Greek), nothing else gets in", async () => {
  const { E } = await loadEngine();
  const math = E.texToMathML("\\sum_{i=1}^{n} i = \\frac{n(n+1)}{2} \\le \\sqrt[3]{\\alpha}<img src=x onerror=alert(1)>");
  const html = math.outerHTML;
  assert.match(html, /<munderover><mo largeop="true">∑<\/mo>/);
  assert.match(html, /<mfrac>/);
  assert.match(html, /<mroot>/);
  assert.match(html, /α/);
  assert.ok(!/<img|onerror=/.test(html.replace(/<mo>[^<]*<\/mo>|<mi>[^<]*<\/mi>/g, "")), "markup in the source is only characters");
  const eq = E.normalizeObject({ id: "e", kind: "equation", x: 0, y: 0, w: 800, h: 200, tex: "E=mc^2", fs: 80 });
  const slide = { type: "blank", elements: [eq] };
  const el = E.render(slide, { deck: deckOf([slide]), index: 0, mode: "present" });
  assert.ok(el.querySelector(".hs-equation math msup"));
  assert.equal(E.normalizeObject({ id: "e2", kind: "equation", x: 0, y: 0, w: 10, h: 10, tex: "  " }), null);
});

test("fields fill in the slide number, the count and the date; text can run in columns", async () => {
  const { E } = await loadEngine();
  const text = E.normalizeObject({ id: "t", kind: "text", x: 0, y: 0, w: 800, h: 200, cols: 2, colGap: 40, text: '<p>No.<span data-field="slideno">#</span> / <span data-field="total">#</span> <span data-field="date">日付</span> <span data-field="evil" onclick="x">z</span></p>' });
  assert.equal(text.cols, 2);
  assert.match(text.text, /data-field="slideno"/);
  assert.ok(!/evil|onclick/.test(text.text), "only known fields");
  const slides = [{ type: "title", title: "表紙", sid: "a" }, { type: "blank", sid: "b", elements: [text] }, { type: "closing", sid: "c" }];
  const el = E.render(slides[1], { deck: deckOf(slides), index: 1, mode: "present" });
  const tx = el.querySelector('.hs-obj[data-el="t"] .hs-obj-tx');
  assert.match(tx.textContent, /No\.2 \/ 3 \d{4}年\d{1,2}月\d{1,2}日/);
  assert.equal(tx.style.columnCount, "2");
});
