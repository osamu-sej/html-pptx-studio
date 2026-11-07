import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import * as ops from "../public/editor/ops.mjs";
import { buildTimeline, convertedSlide, cssLength, KEEPS_LAYOUT, parseColor, parsePolygon, pathSubpaths, patternOf, simplify, snapToPalette } from "../public/editor/convert.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));

async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "animate.js"), "utf8"), context, { filename: "animate.js" });
  return { E: window.SlideEngine, window };
}
const deckWith = (slide) => ({ title: "検証", theme: "clarity", transition: "fade", motion: {}, slides: [{ type: "title", title: "表紙" }, slide, { type: "closing" }] });
const words = (html) => String(html || "").replace(/<[^>]+>/g, "");

// ---------------------------------------------------------------- tables

test("a table keeps its cells, merges inside the table, shares that add up, and its words' settings", async () => {
  const { E } = await loadEngine();
  const o = E.normalizeObject({
    kind: "table", x: 0, y: 0, w: 800, h: 300, fs: 32, color: "#1F3864", lh: 1.5, font: "ud", style: "light1", header: false,
    cols: [2, 1, 1], rows: [1, "x", 1],
    cells: [
      [{ text: "<p>見出し</p><script>x</script>", cs: 9, fill: "#DCE4F2", bold: true }, {}, {}],
      [{ text: "<p>A</p>", rs: 2 }, { text: "<p>B</p>", align: "center", valign: "bottom" }, { evil: 1 }],
      [{ text: "covered" }, {}, { text: "<p>C</p>" }],
    ],
  });
  assert.equal(o.fs, 32, "the table's word size survives (it was lost on every edit)");
  assert.equal(o.color, "#1f3864");
  assert.equal(o.lh, 1.5);
  assert.equal(o.font, "ud");
  assert.equal(o.style, "light1");
  assert.equal(o.header, false);
  assert.deepEqual(o.cols.map((v) => Math.round(v * 100)), [50, 25, 25]);
  assert.deepEqual(o.rows.map((v) => Math.round(v * 1000)), [333, 333, 333], "unusable shares are equal");
  assert.equal(o.cells[0][0].cs, 3, "a span stays inside the table");
  assert.ok(o.cells[0][1].merged && o.cells[0][2].merged);
  assert.equal(o.cells[0][0].text, "<p>見出し</p>");
  assert.equal(o.cells[0][0].fill, "#dce4f2");
  assert.equal(o.cells[1][0].rs, 2);
  assert.ok(o.cells[2][0].merged, "covered cells are marked, their words dropped");
  assert.equal(o.cells[2][0].text, undefined);
  assert.equal(o.cells[1][2].evil, undefined);
  assert.equal(E.objectText(o), "見出し\nA\tB\t\n\tC");
  // Drawn: a real table, merged cells spanning, the class of its style.
  const slide = { type: "blank", title: "", elements: [o] };
  const el = E.render(slide, { deck: deckWith(slide), index: 1, mode: "thumb" });
  const table = el.querySelector(".hs-otable");
  assert.ok(table.classList.contains("ts-light1"));
  assert.ok(!table.classList.contains("has-header"));
  assert.equal(table.querySelector('td[data-r="0"][data-c="0"]').getAttribute("colspan"), "3");
  assert.equal(table.querySelector('td[data-r="1"][data-c="0"]').getAttribute("rowspan"), "2");
  assert.equal(table.querySelectorAll("td").length, 1 + 3 + 2);
  assert.match(table.getAttribute("style"), /font-size:\s*calc\(32px/);
});

test("table operations: insert and delete rows and columns, merge and split, settings, lines", () => {
  let t = { ...ops.makeTable(3, 3, { x: 0, y: 0, w: 600, h: 300 }) };
  t = ops.tableCells(t, 0, 0, 0, 0, { text: "<p>a</p>" });
  t = ops.tableCells(t, 0, 1, 0, 1, { text: "<p>b</p>" });
  t = ops.tableMerge(t, 0, 0, 0, 1);
  assert.equal(t.cells[0][0].cs, 2);
  assert.equal(t.cells[0][0].text, "<p>a</p><p>b</p>", "merged cells keep every word");
  assert.ok(t.cells[0][1].merged);
  assert.deepEqual(ops.tableOrigin(t, 0, 1), [0, 0]);
  // A column inserted inside a merge widens the merge; the table grows.
  t = ops.tableInsertCol(t, 1);
  assert.equal(t.cells[0][0].cs, 3);
  assert.equal(t.cells[0].length, 4);
  assert.equal(Math.round(t.w), 800);
  assert.equal(Math.round(t.cols.reduce((a, b) => a + b, 0) * 1000), 1000);
  t = ops.tableInsertRow(t, 3);
  assert.equal(t.cells.length, 4);
  assert.equal(Math.round(t.h), 400);
  t = ops.tableSplit(t, 0, 2);
  assert.ok(!t.cells[0][0].cs && !t.cells[0][1].merged, "split back into single cells");
  assert.equal(t.cells[0][0].text, "<p>a</p><p>b</p>", "the words stay in the first");
  assert.equal(ops.tableDeleteRows(t, 0, 3), null, "a table never loses every row");
  t = ops.tableDeleteRows(t, 1, 2);
  assert.equal(t.cells.length, 2);
  t = ops.tableDeleteCols(t, 3, 3);
  assert.equal(t.cells[0].length, 3);
  t = ops.tableCells(t, 0, 0, 1, 2, { fill: "#f2f2f2", bold: true });
  assert.ok(t.cells.flat().every((cell) => cell.fill === "#f2f2f2" && cell.bold));
  t = ops.tableCells(t, 0, 0, 0, 0, { bold: undefined });
  assert.equal(t.cells[0][0].bold, undefined, "undefined takes a setting off");
  const before = t.cols.map((f) => f * t.w);
  t = ops.tableResizeLine(t, "cols", 0, 40);
  const after = t.cols.map((f) => f * t.w);
  assert.ok(Math.abs(after[0] - before[0] - 40) < 0.5 && Math.abs(before[1] - after[1] - 40) < 0.5, "a line moves between its two columns");
  t = ops.tableDistribute(t, "cols");
  assert.ok(t.cols.every((f) => Math.abs(f - t.cols[0]) < 1e-4));
  const pasted = ops.tableFromText("A\tB\n1\t2\t3\n");
  assert.equal(pasted.length, 2);
  assert.equal(pasted[1].length, 3);
  assert.equal(words(pasted[1][2].text), "3");
});

// ---------------------------------------------------------------- charts and drawn shapes

test("a chart keeps its kind, labels and numbers, and draws like the layouts' charts", async () => {
  const { E } = await loadEngine();
  const o = E.normalizeObject({ kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { type: "combo", title: " 売上 ", unit: "億円", labels: ["1月", "2月", 3], series: [{ name: "売上", values: [1, "2", "x"] }, { values: [5, 6, 7, 8] }] } });
  assert.equal(o.chart.type, "combo");
  assert.deepEqual([...o.chart.labels], ["1月", "2月", "3"]);
  assert.deepEqual([...o.chart.series[0].values], [1, 2, 0], "numbers only, one per label");
  assert.equal(o.chart.series[1].name, "系列2");
  assert.equal(o.chart.title, "売上");
  assert.equal(E.normalizeObject({ kind: "chart", chart: { type: "pie", labels: [], series: [] } }), null, "a chart needs labels and a series");
  assert.equal(E.normalizeObject({ kind: "chart", chart: { type: "contour", labels: ["a"], series: [{ values: [1] }] } }).chart.type, "bar", "unknown kinds draw as bars");
  // The object's data and the engine's chart model are the same thing, both ways.
  for (const type of Object.keys(E.CHART_KINDS)) {
    const chart = E.normalizeObject({ kind: "chart", chart: { type, labels: ["A", "B"], series: [{ name: "一", values: [1, 2] }, { name: "二", values: [3, 4] }] } }).chart;
    const model = E.chartModel(E.chartSpec(chart));
    // The pareto chart is the one that draws its first series from the largest (the data itself stays as it was typed).
    assert.deepEqual([...model.labels], type === "pareto" ? ["B", "A"] : ["A", "B"], type);
    assert.deepEqual([...model.series[0].values], type === "pareto" ? [2, 1] : [1, 2], type);
  }
  const slide = { type: "blank", title: "", elements: [o] };
  const el = E.render(slide, { deck: deckWith(slide), index: 1, mode: "thumb" });
  assert.equal(el.querySelector(".hs-obj-chart-title").textContent, "売上");
  assert.ok(el.querySelector(".hs-obj-chart svg"));
});

test("a shape drawn by hand keeps its points (fractions of its box), open or closed, straight or curved", async () => {
  const { E } = await loadEngine();
  const closed = E.normalizeObject({ kind: "shape", shape: "custom", x: 10, y: 10, w: 100, h: 50, path: { pts: [[0, 0], [1, 0], [0.5, 1], ["x", 2]], closed: true, curve: 1 } });
  assert.equal(closed.shape, "custom");
  assert.deepEqual(closed.path.pts.map((p) => [...p]), [[0, 0], [1, 0], [0.5, 1]], "unusable points go");
  assert.equal(closed.path.closed, true);
  assert.equal(closed.path.curve, true);
  assert.equal(E.normalizeObject({ kind: "shape", shape: "custom", x: 0, y: 0, w: 10, h: 10, path: { pts: [[0, 0]] } }).shape, "rect", "one point is no drawing");
  assert.equal(E.normalizeObject({ kind: "text", shape: "custom", x: 0, y: 0, w: 10, h: 10, path: { pts: [[0, 0], [1, 1]] } }).shape, "rect", "text boxes are not drawn by hand");
  const straight = E.freeformD({ pts: [[0, 0], [1, 0], [1, 1]], closed: true }, 200, 100);
  assert.equal(straight, "M0 0 L200 0 L200 100 Z");
  assert.match(E.freeformD({ pts: [[0, 0], [0.5, 1], [1, 0]], curve: true }, 100, 100), /^M0 0 C/);
  const g = E.geometry("custom", 100, 50, null, { pts: [[0, 0], [1, 1]] });
  assert.equal(g.open, true, "an open drawing has no fill");
  const slide = { type: "blank", title: "", elements: [closed] };
  const el = E.render(slide, { deck: deckWith(slide), index: 1, mode: "thumb" });
  assert.equal(el.querySelector(".hs-obj").dataset.fill, E.themeColors(E.DEFAULT_THEME).accent1, "a shape's fill is アクセント 1 (PowerPoint's default)");
  assert.equal(E.objectName(closed, 0), "フリーフォーム 1");
});

test("a picture's crop and a text link are kept safely", async () => {
  const { E } = await loadEngine();
  const pic = E.normalizeObject({ kind: "image", src: "asset:ai", x: 0, y: 0, w: 300, h: 200, crop: { l: 0.1, t: 0.2, r: 0, b: 0.05 } });
  assert.deepEqual({ ...pic.crop }, { l: 0.1, t: 0.2, r: 0, b: 0.05 });
  const slide = { type: "blank", title: "", elements: [pic, { id: "t", kind: "text", x: 0, y: 300, w: 400, h: 60, text: '<p><a href="https://example.com/a">リンク</a> と <a href="javascript:alert(1)">危ない</a></p>' }] };
  const el = E.render(slide, { deck: deckWith(slide), index: 1, mode: "present" });
  const img = el.querySelector(".hs-obj-img img");
  assert.match(img.getAttribute("style"), /width:\s*111\.1111%/);
  assert.match(img.getAttribute("style"), /left:\s*-11\.1111%/);
  const links = [...el.querySelectorAll(".hs-obj-tx a")];
  assert.equal(links.length, 1, "only web and mail links are links");
  assert.equal(links[0].getAttribute("href"), "https://example.com/a");
  assert.equal(links[0].getAttribute("rel"), "noopener noreferrer");
});

// ---------------------------------------------------------------- 図形に変換

test("図形に変換: colours come back as the palette's (a tint is its colour, faded), with how far they were", () => {
  assert.deepEqual(parseColor("rgb(31, 56, 100)"), { r: 31, g: 56, b: 100, a: 1 });
  assert.deepEqual(parseColor("rgba(0, 0, 0, 0)"), null);
  assert.deepEqual(parseColor("color(srgb 0.717647 0.764706 0.854902)"), { r: 183, g: 195, b: 218, a: 1 });
  assert.deepEqual(parseColor("color(srgb 0.121569 0.219608 0.392157 / 0.3)"), { r: 31, g: 56, b: 100, a: 0.3 });
  assert.equal(parseColor("transparent"), null);
  const fills = ["#ffffff", "#f1f5fb", "#dce4f2", "#b7c3da", "#f2f2f2", "#d9d9d9", "#f5f0ea", "#d6c9b8", "#1f3864", "#808080"];
  assert.deepEqual(snapToPalette(parseColor("#dce4f2"), fills), { color: "#dce4f2", opacity: 1, err: 0 });
  // color-mix(in srgb, #b7c3da 86%, #f1f5fb): close to 青灰, a little lighter.
  const mixed = snapToPalette(parseColor("color(srgb 0.758588 0.8 0.878196)"), fills);
  assert.equal(mixed.color, "#b7c3da");
  assert.ok(mixed.opacity > 0.75 && mixed.opacity < 0.95, `opacity ${mixed.opacity}`);
  // The accent at 7% over white is navy at 7%.
  const faint = snapToPalette({ r: 31, g: 56, b: 100, a: 0.07 }, fills);
  assert.deepEqual([faint.color, faint.opacity], ["#1f3864", 0.07]);
  assert.ok(faint.err < 1);
  // Text takes the nearest palette colour as it is.
  assert.equal(snapToPalette(parseColor("rgb(30, 30, 30)"), ["#1a1a1a", "#1f3864", "#808080"], { fade: false }).color, "#1a1a1a");
});

test("図形に変換: clip-path polygons and SVG paths become points", () => {
  assert.equal(cssLength("calc(100% - 44px)", 200), 156);
  assert.equal(cssLength("50%", 80), 40);
  assert.equal(cssLength("12px", 80), 12);
  assert.equal(cssLength("calc(100% + 10px)", 100), 110);
  assert.deepEqual(parsePolygon("polygon(0px 0px, calc(100% - 44px) 0px, 100% 50%, calc(100% - 44px) 100%, 0px 100%)", 200, 100), [[0, 0], [0.78, 0], [1, 0.5], [0.78, 1], [0, 1]]);
  assert.equal(parsePolygon("circle(50%)", 10, 10), null);
  const [tri] = pathSubpaths("M-16,-22 L14,0 L-16,22 Z");
  assert.ok(tri.closed && tri.straight);
  assert.deepEqual(tri.pts, [[-16, -22], [14, 0], [-16, 22]]);
  const subs = pathSubpaths("M0 0 h10 v10 z m 20 0 l 5 5 M100 100 A 10 10 0 0 1 120 100");
  assert.equal(subs.length, 3);
  assert.deepEqual(subs[0].pts, [[0, 0], [10, 0], [10, 10]]);
  assert.deepEqual(subs[1].pts, [[20, 0], [25, 5]], "a relative move starts from where the closed path began");
  assert.equal(subs[2].straight, false);
  assert.match(subs[2].d, /^M100 100 A10 10 0 0 1 120 100$/);
  const smooth = pathSubpaths("M0 0 C 10 0 20 10 30 10 S 50 20 60 20");
  assert.match(smooth[0].d, /C10 0 20 10 30 10 C40 10 50 20 60 20/, "S reflects the last control point");
  // A ruled pattern (a Gantt chart's month lines) becomes lines; a hatch its average tone.
  const ruled = patternOf("repeating-linear-gradient(90deg, rgba(0, 0, 0, 0) 0px, rgba(0, 0, 0, 0) calc(25% - 1px), rgb(217, 217, 217) calc(25% - 1px), rgb(217, 217, 217) 25%)", 800, 60);
  assert.equal(ruled.lines.length, 4);
  assert.deepEqual(ruled.lines.map((l) => Math.round(l.x1)), [200, 400, 600, 800]);
  assert.ok(ruled.lines.every((l) => l.y1 === 0 && l.y2 === 60 && Math.abs(l.w - 1) < 0.01));
  const hatch = patternOf("repeating-linear-gradient(135deg, rgba(0, 0, 0, 0) 0px, rgba(0, 0, 0, 0) 10px, rgba(128, 128, 128, 0.14) 10px, rgba(128, 128, 128, 0.14) 20px)", 200, 40);
  assert.ok(hatch.fill && Math.abs(hatch.fill.a - 0.07) < 0.001, "half the stripes at 14%");
  assert.equal(patternOf("linear-gradient(red, blue)", 10, 10), null);
  // Points on a straight line collapse to its ends.
  assert.deepEqual(simplify([[0, 0], [5, 0.1], [10, 0], [10, 10]], 0.5), [[0, 0], [10, 0], [10, 10]]);
});

test("図形に変換: the layout's build becomes animations on the item groups; the slide keeps what it should", () => {
  const groups = [{ g: 1, target: "grp:b" }, { g: 0, target: "grp:a" }, { g: 2, target: "o3" }];
  const clicks = buildTimeline({ build: "click", entrance: "rise", groups });
  assert.deepEqual(clicks.map((e) => [e.el, e.fx, e.dir, e.start]), [["grp:a", "floatIn", "up", "click"], ["grp:b", "floatIn", "up", "click"], ["o3", "floatIn", "up", "click"]]);
  const cascade = buildTimeline({ build: "cascade", entrance: "slide", groups });
  assert.deepEqual(cascade.map((e) => [e.fx, e.dir, e.start, e.delay]), [["flyIn", "left", "with", 0], ["flyIn", "left", "with", 170], ["flyIn", "left", "with", 340]]);
  assert.deepEqual(buildTimeline({ build: "fade", entrance: "rise", groups }).map((e) => e.fx), ["fade", "fade", "fade"]);
  assert.deepEqual(buildTimeline({ build: "none", entrance: "rise", groups }), []);
  // Animations on the layout's items follow them; the rest stay.
  const kept = buildTimeline({ build: "none", entrance: "rise", groups, timeline: [{ id: "x", el: "@g1", cls: "em", fx: "pulse", by: "word" }, { id: "y", el: "@title", cls: "in", fx: "fade" }, { id: "z", el: "@g9", cls: "in", fx: "fade" }] });
  assert.deepEqual(kept.map((e) => e.el), ["grp:b", "@title"], "an item that is gone takes its animation with it");
  assert.equal(kept[0].by, undefined, "by paragraph/word applies to a layout part, not a group");
  const slide = { type: "process", title: "T", takeaway: "K", notes: "N", source: "S", sid: "s1", items: [{ title: "a" }], animation: "click", media: { src: "asset:ai" }, elements: [{ id: "mine", kind: "shape" }], transitionDur: 800, details: [{ target: "items[0]", text: "d" }] };
  const next = convertedSlide(slide, [{ id: "c1", kind: "text" }], clicks);
  assert.equal(next.type, "blank");
  assert.deepEqual(next.elements.map((o) => o.id), ["c1", "mine"], "converted pieces under the ones placed by hand");
  for (const key of ["title", "takeaway", "notes", "source", "sid", "transitionDur", "details"]) assert.deepEqual(next[key], slide[key], key);
  for (const key of ["items", "animation", "media"]) assert.equal(next[key], undefined, key);
  assert.equal(next.timeline.length, 3);
  assert.ok(KEEPS_LAYOUT.has("title") && KEEPS_LAYOUT.has("section") && !KEEPS_LAYOUT.has("process"));
});

test("an object made from a layout item opens the item's card and deep-dive page", async () => {
  const { E } = await loadEngine();
  const o = E.normalizeObject({ kind: "shape", x: 100, y: 300, w: 400, h: 200, item: "items[1]" });
  assert.equal(o.item, "items[1]");
  assert.equal(E.normalizeObject({ kind: "shape", x: 0, y: 0, w: 1, h: 1, item: "<script>" }).item, undefined);
  const slide = { type: "blank", title: "カード", elements: [o], details: [{ target: "items[1]", text: "根拠の数字" }] };
  const deck = { ...deckWith(slide), slides: [{ type: "title", title: "表紙" }, slide, { type: "content", title: "深掘り", drillOf: "items[1]", points: ["x"] }, { type: "closing" }] };
  const el = E.render(slide, { deck, index: 1, mode: "present" });
  const node = el.querySelector('.hs-obj[data-item="items[1]"]');
  assert.equal(node.dataset.detail, "items[1]");
  assert.equal(node.dataset.drill, "2", "the deep-dive page opens from it");
  assert.ok(node.querySelector(".hs-drill-badge"));
  assert.doesNotMatch(node.getAttribute("style"), /position/, "the mark never takes the object out of its place");
});
