// 図形の結合 (merge.mjs): outlines to polygons, the five ways to merge, and the hand-drawn shape they become.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { area, canMerge, mergeObjects, pathRings, shapeRegion } from "../public/editor/merge.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));

async function load() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  const module = { exports: {} };
  new Function("module", "exports", "define", await readFile(join(root, "node_modules", "polygon-clipping", "dist", "polygon-clipping.umd.min.js"), "utf8"))(module, module.exports, undefined);
  return { E: window.SlideEngine, pc: module.exports };
}
const polyArea = (multi) => multi.reduce((sum, poly) => sum + poly.reduce((s, ring, k) => s + (k ? -1 : 1) * Math.abs(area(ring)), 0), 0);
const shapeArea = (o) => {
  const rings = [o.path.pts, ...(o.path.parts || [])].map((r) => r.map(([x, y]) => [o.x + x * o.w, o.y + y * o.h]));
  // Even-odd: the rings of one shape never overlap after merging, so holes inside the outer ring subtract.
  const sorted = rings.map((r) => Math.abs(area(r))).sort((a, b) => b - a);
  return sorted;
};
let n = 0;
const makeId = () => `m${(n += 1)}`;
const square = (id, x, y, s = 100, extra = {}) => ({ id, kind: "shape", shape: "rect", x, y, w: s, h: s, fill: "#dce4f2", ...extra });

test("path outlines: lines, arcs and curves become closed rings with the right area", async () => {
  const rect = pathRings("M0 0 L100 0 L100 50 L0 50 Z");
  assert.equal(rect.length, 1);
  assert.equal(Math.abs(area(rect[0])), 5000);
  const circle = pathRings("M100 50 A50 50 0 0 1 0 50 A50 50 0 0 1 100 50 Z");
  assert.ok(Math.abs(Math.abs(area(circle[0])) - Math.PI * 2500) < 30, `circle area ${area(circle[0])}`);
  const two = pathRings("M0 0 L10 0 L10 10 Z m20 0 l10 0 l0 10 z");
  assert.equal(two.length, 2, "relative commands and several rings");
  assert.equal(two[1][0][0], 20);
  const curve = pathRings("M0 0 C0 100 100 100 100 0 Z");
  assert.ok(curve[0].length > 8, "a curve is flattened into many steps");
});

test("a shape's region: turned and flipped as drawn; a frame keeps its hole", async () => {
  const { E, pc } = await load();
  const turned = shapeRegion(E.withDefaults({ id: "a", kind: "shape", shape: "rect", x: 0, y: 0, w: 200, h: 100, rot: 90 }), E, pc);
  const xs = turned[0][0].map((p) => p[0]);
  assert.ok(Math.abs(Math.min(...xs) - 50) < 0.01 && Math.abs(Math.max(...xs) - 150) < 0.01, "turned 90°: 100 wide around the same centre");
  const frame = shapeRegion(E.withDefaults({ id: "f", kind: "shape", shape: "frame", x: 0, y: 0, w: 300, h: 200 }), E, pc);
  assert.equal(frame[0].length, 2, "outer ring and hole");
  assert.ok(Math.abs(polyArea(frame) - (300 * 200 - 250 * 150)) < 1);
});

test("接合・型抜き/合成・重なり抽出・単純型抜き・切り出し", async () => {
  const { E, pc } = await load();
  const a = square("a", 0, 0, 100, { text: "<p>A</p>", fill: "#f2eadf" });
  const b = square("b", 50, 50, 100);
  assert.ok(canMerge([a, b]) && !canMerge([a]) && !canMerge([a, { id: "l", kind: "line" }]));
  const [union] = mergeObjects([a, b], "union", E, pc, makeId);
  assert.equal(union.id, "a", "keeps the first selected shape's id (its animations stay)");
  assert.equal(union.shape, "custom");
  assert.equal(union.fill, "#f2eadf", "and its format");
  assert.equal(union.text, "<p>A</p>", "and its text");
  assert.deepEqual([union.x, union.y, union.w, union.h], [0, 0, 150, 150]);
  const near = (value, want) => assert.ok(Math.abs(value - want) < 5, `${value} ≈ ${want}`);
  near(shapeArea(union)[0], 17500);
  const [combine] = mergeObjects([a, b], "combine", E, pc, makeId);
  assert.equal(1 + (combine.path.parts?.length || 0), 2, "two pieces around the cut-out overlap");
  const [inter] = mergeObjects([a, b], "intersect", E, pc, makeId);
  assert.deepEqual([inter.x, inter.y, inter.w, inter.h], [50, 50, 50, 50]);
  const [sub] = mergeObjects([a, b], "subtract", E, pc, makeId);
  near(shapeArea(sub)[0], 7500);
  const [subOther] = mergeObjects([b, a], "subtract", E, pc, makeId);
  assert.equal(subOther.id, "b", "the order of selection decides which shape stays");
  const pieces = mergeObjects([a, b], "fragment", E, pc, makeId);
  assert.equal(pieces.length, 3, "A only, the overlap, B only");
  const sizes = pieces.map((p) => shapeArea(p)[0]).sort((x, y) => x - y);
  [2500, 7500, 7500].forEach((want, k) => near(sizes[k], want));
  assert.equal(pieces.filter((p) => p.text).length, 1, "the text stays on one piece");
  assert.equal(mergeObjects([a, square("far", 500, 500)], "intersect", E, pc, makeId), null, "nothing in common");
  const hole = mergeObjects([square("big", 0, 0, 300), square("small", 100, 100, 100)], "subtract", E, pc, makeId)[0];
  assert.equal(hole.path.parts.length, 1, "a hole is a second ring");
  const back = E.normalizeObject(hole);
  assert.equal(back.path.parts.length, 1, "kept by the engine");
  const g = E.geometry("custom", back.w, back.h, null, back.path);
  assert.equal(g.rule, "evenodd", "drawn even-odd");
  assert.equal(g.paths[0].match(/M/g).length, 2);
});
