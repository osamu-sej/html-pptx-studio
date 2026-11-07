// 描画 (ink.mjs, objects.js): strokes kept as fractions of their box, drawn as paths; ink turned into shapes.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { inkObject, recognize, simplify, strokesOf } from "../public/editor/ink.mjs";
import { objectSchema } from "../server/schemas.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "animate.js"), "utf8"), context, { filename: "animate.js" });
  return { E: window.SlideEngine };
}
const circle = (cx, cy, r, n = 60) => Array.from({ length: n + 1 }, (_, i) => [cx + r * Math.cos((i / n) * 2 * Math.PI), cy + r * 0.8 * Math.sin((i / n) * 2 * Math.PI)]);
const polyline = (pts, per = 12) => pts.slice(1).flatMap((p, i) => Array.from({ length: per }, (_, k) => [pts[i][0] + ((p[0] - pts[i][0]) * k) / per, pts[i][1] + ((p[1] - pts[i][1]) * k) / per])).concat([pts.at(-1)]);

test("strokes become an ink object (fractions of its box) and come back in slide pixels", async () => {
  const { E } = await loadEngine();
  const o = inkObject([{ pts: [[100, 100], [300, 200]], color: "#1f3864", width: 6 }], "i1");
  assert.equal(o.kind, "ink");
  assert.ok(o.x < 100 && o.y < 100 && o.x + o.w > 300, "the box holds the stroke and its width");
  const back = strokesOf(o)[0].pts;
  assert.ok(Math.abs(back[0][0] - 100) < 0.1 && Math.abs(back[1][1] - 200) < 0.1);
  const n = E.normalizeObject({ ...o, strokes: [...o.strokes, { pts: [], color: "#ff00ff" }, { pts: [[0.5, 0.5]], color: "#ff00ff", width: 999, highlighter: true }] });
  assert.equal(n.strokes.length, 2, "an empty stroke goes");
  assert.equal(n.strokes[1].color, "#ff00ff", "a pen of any colour");
  assert.equal(n.strokes[1].width, 80);
  assert.equal(E.normalizeObject({ id: "x", kind: "ink", x: 0, y: 0, w: 10, h: 10, strokes: [] }), null, "ink without strokes is nothing");
  assert.ok(objectSchema.safeParse({ id: "i", kind: "ink" }).success);
  const slide = { type: "blank", elements: [n] };
  const el = E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  assert.equal(el.querySelectorAll('.hs-obj[data-kind="ink"] .hs-ink path').length, 2);
  assert.ok(el.querySelector(".hs-ink path.hs-ink-hl"), "a highlighter is marked");
  assert.equal(E.normalizeTimeline([{ el: "i1", cls: "in", fx: "draw" }], slide).length, 1, "描画で再生 is the draw effect");
});

test("インクを図形に変換: lines, ellipses, rectangles and triangles; handwriting stays ink", () => {
  assert.equal(recognize([[100, 100], [140, 102], [200, 101], [400, 104]]).kind, "line");
  const ellipse = recognize(circle(500, 400, 200));
  assert.equal(ellipse.shape, "ellipse");
  assert.ok(Math.abs(ellipse.w - 400) < 6 && Math.abs(ellipse.h - 320) < 6);
  const rect = recognize(polyline([[100, 100], [500, 100], [500, 300], [100, 300], [102, 104]]));
  assert.equal(rect.shape, "rect");
  assert.deepEqual([Math.round(rect.x), Math.round(rect.y), Math.round(rect.w), Math.round(rect.h)], [100, 100, 400, 200]);
  const tri = recognize(polyline([[300, 100], [500, 400], [100, 400], [301, 104]]));
  assert.equal(tri.shape, "custom");
  assert.equal(tri.path.pts.length, 3);
  // A scribble (handwriting) that does not close is no shape.
  assert.equal(recognize(polyline([[100, 100], [140, 160], [180, 100], [220, 160], [260, 100], [300, 160]])), null);
  assert.ok(simplify(polyline([[0, 0], [100, 0], [100, 100]]), 1).length <= 4);
});
