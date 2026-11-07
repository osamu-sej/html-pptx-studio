// A PowerPoint deck brought over as it looks (tools/pptx_exact.py): source pages draw their own title box
// and master marks as editable objects, bullets as PowerPoint draws them, table borders cell by cell,
// and charts drawn with their own formatting (chart.style → officeChart).
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { applyChatOperations } from "../server/chat.mjs";
import { codexChatSchema, codexDeckSchema, deckShape } from "../server/schemas.mjs";

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

const deckWith = (...slides) => ({ title: "検証", theme: "clarity", transition: "fade", motion: {}, slides });

test("a source page keeps its master, viewport, gradient, shadow, font and hidden slide state", async () => {
  const { E } = await loadEngine();
  const [shape] = E.normalizeObjects([{ id: "g", kind: "shape", shape: "rect", x: 300, y: 200, w: 400, h: 200,
    fill: "none", text: "<p>Source</p>", fontFace: "Arial", gradient: { angle: 90,
      stops: [{ at: 0, color: "#ff0000", opacity: 1 }, { at: 1, color: "#0000ff", opacity: 0.5 }] },
    shadow: { dx: 5, dy: 6, blur: 10, color: "#000000", opacity: 0.5 } }]);
  const page = { type: "blank", title: "Source", hideTitle: true, master: "source", sourceViewport: { x: 240, y: 0, w: 1440, h: 1080 }, elements: [shape] };
  const hidden = { ...page, sid: "hidden", hidden: true };
  const deck = deckWith(page, hidden);
  assert.deepEqual([...E.storyMap(deck.slides).order], [0], "hidden source pages remain in the deck but outside the show");
  const root = E.render(page, { deck, index: 0, mode: "present" });
  assert.equal(root.querySelector(".hs-foot"), null);
  assert.match(root.querySelector(".hs-objects").getAttribute("style"), /clip-path/);
  assert.equal(root.querySelectorAll("linearGradient stop").length, 2);
  assert.match(root.querySelector(".hs-obj-rot").getAttribute("style"), /drop-shadow/);
  assert.match(root.querySelector(".hs-obj-tx").getAttribute("style"), /Arial/);
  assert.equal(deckShape.parse(deck).slides[1].hidden, true);
});

test("objects and table cells beyond the old import limits are retained", async () => {
  const { E } = await loadEngine();
  const objects = Array.from({ length: 301 }, (_, i) => ({ id: `part${i}`, kind: "shape", shape: "rect", x: i, y: 0, w: 10, h: 10, fill: "#ffffff" }));
  assert.equal(E.normalizeObjects(objects).length, 301);
  const cells = Array.from({ length: 41 }, () => Array.from({ length: 21 }, () => ({ text: "<p>value</p>" })));
  const [table] = E.normalizeObjects([{ id: "wide", kind: "table", x: 0, y: 0, w: 800, h: 600, cells }]);
  assert.equal(table.cells.length, 41);
  assert.equal(table.cells[0].length, 21);
});

test("an imported page draws its own title box on the master PowerPoint used", async () => {
  const { E } = await loadEngine();
  const cover = { type: "blank", title: "発注業務の改革", hideTitle: true, master: "title", elements: [{ id: "t", kind: "text", x: 145, y: 445, w: 1700, h: 60, text: "<p>発注業務の改革</p>" }] };
  const page = { type: "blank", title: "改革の全体像", hideTitle: true, elements: [] };
  const deck = deckWith(cover, page);
  const first = E.render(cover, { deck, index: 0, mode: "present" });
  assert.equal(first.querySelector(".hs-title"), null, "no second title over the slide's own title box");
  const second = E.render(page, { deck, index: 1, mode: "edit" });
  assert.equal(second.querySelector(".hs-title"), null);
  const plain = E.render({ type: "blank", title: "白紙" }, { deck, index: 1, mode: "edit" });
  assert.ok(plain.querySelector(".hs-title"), "a 白紙 page people make still has its title");
});

test("the server keeps an imported page's title setting and master through AI changes", () => {
  for (const schema of [codexDeckSchema, codexChatSchema]) assert.doesNotMatch(JSON.stringify(schema), /"(hideTitle|master)":/, "the AI never writes them");
  const deck = deckShape.parse({ title: "t", slides: [{ type: "blank", title: "表紙", hideTitle: true, master: "title", elements: [] }, { type: "blank", title: "中扉", hideTitle: true, master: "title", elements: [] }, { type: "closing" }] });
  assert.equal(deck.slides[0].master, "title");
  assert.equal(deck.slides[1].hideTitle, true);
  const applied = applyChatOperations(deck.slides, { operations: [{ op: "replace", slide: 2, content: { type: "blank", title: "中扉（改）" } }] });
  assert.equal(applied.slides[1].hideTitle, true);
  assert.equal(applied.slides[1].master, "title");
  const relaid = applyChatOperations(deck.slides, { operations: [{ op: "replace", slide: 2, content: { type: "content", title: "整理", points: ["A"] } }] });
  assert.equal(relaid.slides[1].hideTitle, undefined, "a layout page shows its title again");
  assert.throws(() => deckShape.parse({ title: "t", slides: [{ type: "blank", master: "bogus" }, { type: "closing" }] }));
});

test("bullets as PowerPoint draws them: a safe mark per paragraph, never on list items", async () => {
  const { E } = await loadEngine();
  const html = E.sanitizeRich('<p data-bullet="■">一</p><p data-bullet="1.">二</p><p data-bullet="&quot;><script>">三</p><p data-bullet="abcdef">四</p><ul><li data-bullet="●">五</li></ul><p data-indent="1" data-bullet="–">六</p>');
  assert.match(html, /<p data-bullet="■">一<\/p>/);
  assert.match(html, /<p data-bullet="1\.">二<\/p>/);
  assert.match(html, /<p>三<\/p>/, "a mark that could break out is dropped");
  assert.match(html, /<p>四<\/p>/, "at most four characters");
  assert.match(html, /<li>五<\/li>/);
  assert.match(html, /<p data-bullet="–" data-indent="1">六<\/p>/);
  assert.equal(E.sanitizeRich(html), html);
});

test("table borders cell by cell, as the deck's table style drew them", async () => {
  const { E } = await loadEngine();
  const [table] = E.normalizeObjects([{ id: "t", kind: "table", x: 0, y: 0, w: 600, h: 200, cols: [0.5, 0.5], rows: [0.5, 0.5],
    cells: [[{ text: "<p>A</p>", bt: { c: "#1F3864", w: 4 }, bb: "none" }, { text: "<p>B</p>", br: { c: "red", w: 999 } }], [{ text: "<p>C</p>", bl: { c: "nope", w: 2 } }, { text: "<p>D</p>" }]] }]);
  const [a, b] = table.cells[0];
  assert.deepEqual({ ...a.bt }, { c: "#1f3864", w: 4 });
  assert.equal(a.bb, "none");
  assert.ok(b.br.w <= 40, "a border's width is clamped");
  assert.equal(table.cells[1][0].bl, undefined, "a border with no colour is dropped");
  const el = E.render({ type: "blank", title: "", elements: [table] }, { deck: deckWith({ type: "blank" }), index: 0, mode: "present" });
  const td = el.querySelector('td[data-r="0"][data-c="0"]');
  assert.match(td.getAttribute("style"), /border-top:\s*4px solid #1f3864/);
  assert.match(td.getAttribute("style"), /border-bottom:\s*none/);
});

test("Excel number formats for data labels and axes", async () => {
  const { E } = await loadEngine();
  const cases = [
    [1234.5, "General", "1234.5"], [0.25, "0%", "25%"], [0.125, "0.0%", "12.5%"], [1234.5, "#,##0", "1,235"], [100, '0"%"', "100%"],
    [42, '0"h"', "42h"], [-5, "0;(0)", "(5)"], [-5, "0", "-5"], [3.14159, "0.00", "3.14"], [1500, '"¥"#,##0', "¥1,500"], [2500000, '#,##0,"千"', "2,500千"],
    [12, "[Red]0", "12"], [0.5, "#.##", ".5"], [7, "00", "07"],
  ];
  for (const [value, code, want] of cases) assert.equal(E.numFormat(value, code), want, `${value} in ${code}`);
});

test("a chart's PowerPoint formatting is kept, checked and limited", async () => {
  const { E } = await loadEngine();
  const labels = Array.from({ length: 70 }, (_, i) => `${i + 1}月`);
  const [o] = E.normalizeObjects([{ id: "c", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: {
    type: "bar", labels, series: [{ name: "売上", values: labels.map((_, i) => i) }],
    style: { type: "bar", dir: "bar", gap: 9999, font: { size: 28, color: "#1A1A1A", evil: 1 }, legend: { pos: "b" }, plot: { x: 0.1, y: 0.1, w: 0.8, h: 0.8 },
      cat: { reverse: true, line: "#D9D9D9", font: { size: 40 } }, val: { hide: true, min: 0, max: 118, grid: "none", format: "0%" },
      series: [{ kind: "bar", color: "#B7C3DA", points: [{ i: 0, color: "#1F3864" }, { i: 99, color: "#000000" }, { i: 1, color: "javascript:x" }], label: { val: true, pos: "outEnd", format: '0"%"', font: { size: 48, bold: true } },
        pointLabels: [{ i: 0, val: true, runs: [{ t: "最上位  ", font: { size: 40 } }, { t: "100%", font: { bold: true } }] }, { i: 3, show: false }], marker: { s: "bogus", z: 999 } }],
      bogus: true } } }]);
  assert.equal(o.chart.labels.length, 70, "all 70 categories survive normalization");
  const st = o.chart.style;
  assert.equal(st.gap, 500);
  assert.deepEqual({ ...st.font }, { size: 28, color: "#1a1a1a" });
  assert.equal(st.bogus, undefined);
  assert.equal(st.val.format, "0%");
  assert.equal(st.val.grid, "none");
  const [series] = st.series;
  assert.deepEqual(series.points.map((p) => ({ ...p })), [{ i: 0, color: "#1f3864" }], "points past the categories and bad colours are dropped");
  assert.equal(series.marker.s, "none");
  assert.equal(series.marker.z, 80);
  assert.equal(series.pointLabels[0].runs[1].t, "100%");
  assert.equal(series.pointLabels[1].show, false);
});

test("a PowerPoint chart is drawn with its own colours, labels, axes and legend", async () => {
  const { E } = await loadEngine();
  const chart = {
    type: "bar", labels: ["オペレーション支援部", "商品本部", "人事部"], series: [{ name: "活用率", values: [100, 91, 33] }],
    style: { type: "bar", dir: "bar", gap: 55, font: { size: 40, color: "#1a1a1a" }, cat: { reverse: true, line: "#d9d9d9" }, val: { hide: true, min: 0, max: 118, line: "none" },
      series: [{ kind: "bar", color: "#b7c3da", points: [{ i: 0, color: "#1f3864" }], label: { val: true, format: '0"%"', pos: "outEnd" }, pointLabels: [{ i: 0, val: true, format: '0"%"', font: { size: 48, bold: true, color: "#1f3864" } }] }] },
  };
  const slide = { type: "blank", title: "", elements: E.normalizeObjects([{ id: "c", kind: "chart", x: 100, y: 200, w: 900, h: 500, chart }]) };
  const el = E.render(slide, { deck: deckWith(slide), index: 0, mode: "present" });
  const svg = el.querySelector('[data-el="c"] svg.hs-ochart');
  assert.ok(svg, "drawn as PowerPoint draws it, not with the layouts' look");
  assert.equal(svg.getAttribute("viewBox"), "0 0 900 500", "in slide pixels");
  const bars = [...svg.querySelectorAll(".hs-obar")];
  assert.deepEqual(bars.map((b) => b.getAttribute("fill")), ["#1f3864", "#b7c3da", "#b7c3da"], "the first bar keeps its own colour");
  // A reversed category axis puts the first category at the top.
  assert.ok(Number(bars[0].getAttribute("y")) < Number(bars[2].getAttribute("y")));
  // Bars are as long as their values on the 0–118 axis.
  const widths = bars.map((b) => Number(b.getAttribute("width")));
  assert.ok(Math.abs(widths[0] / widths[2] - 100 / 33) < 0.01);
  const texts = [...svg.querySelectorAll("text")].map((t) => t.textContent);
  for (const words of ["100%", "91%", "33%", "オペレーション支援部", "人事部"]) assert.ok(texts.includes(words), words);
  assert.ok(!texts.some((t) => /^(0|20|40)$/.test(t)), "the hidden value axis has no labels");
  const first = [...svg.querySelectorAll("text")].find((t) => t.textContent === "100%");
  assert.match(first.getAttribute("style"), /font-size:\s*48px;\s*font-weight:\s*700/);
  assert.equal(first.getAttribute("fill"), "#1f3864");
  // A chart without PowerPoint formatting keeps the layouts' look.
  const plain = E.render({ type: "blank", title: "", elements: E.normalizeObjects([{ id: "p", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { type: "bar", labels: ["A", "B"], series: [{ name: "s", values: [1, 2] }] } }]) }, { deck: deckWith({ type: "blank" }), index: 0, mode: "present" });
  assert.equal(plain.querySelector(".hs-ochart"), null);
  assert.ok(plain.querySelector(".hs-chart svg"));
});

test("line, pie, stacked and combo charts in PowerPoint's way", async () => {
  const { E } = await loadEngine();
  const draw = (chart) => E.officeChart(E.normalizeObjects([{ id: "c", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart }])[0].chart, 900, 500);
  // A line with a marker and a two-run label only on its last point; gridlines; a legend at the bottom.
  const line = draw({ type: "multi-line", labels: ["4月", "5月", "6月"], series: [{ name: "全体", values: [45, 58, 78] }, { name: "下位", values: [22, 22, 33] }],
    style: { type: "multi-line", legend: { pos: "b" }, val: { grid: "#f2f2f2" }, series: [{ kind: "line", color: "#1f3864", width: 5, marker: { s: "none", at: [2] }, pointLabels: [{ i: 2, val: true, runs: [{ t: "全体  " }, { t: "78%", font: { size: 48, bold: true } }] }] }, { kind: "line", color: "#b7c3da", width: 3, marker: { s: "none" } }] } });
  assert.equal(line.querySelectorAll(".hs-oline").length, 2);
  assert.equal(line.querySelectorAll(".hs-omark").length, 1, "a marker only where PowerPoint put one");
  assert.equal(line.querySelector(".hs-oline").getAttribute("stroke-width"), "5");
  assert.ok([...line.querySelectorAll("tspan")].some((t) => t.textContent === "78%"));
  assert.ok(line.querySelector(".hs-ochart-legend"));
  assert.ok(line.querySelectorAll(".hs-ochart-grid line").length >= 3);
  // A pie: a slice per category in its own colour, percentages, no hole.
  const pie = draw({ type: "donut", labels: ["A", "B", "C"], series: [{ name: "構成", values: [50, 30, 20] }],
    style: { type: "donut", hole: 0, series: [{ points: [{ i: 0, color: "#1f3864" }, { i: 1, color: "#b7c3da" }, { i: 2, color: "#d6c9b8" }], label: { pct: true } }] } });
  assert.deepEqual([...pie.querySelectorAll(".hs-oslice")].map((p) => p.getAttribute("fill")), ["#1f3864", "#b7c3da", "#d6c9b8"]);
  assert.ok([...pie.querySelectorAll("text")].map((t) => t.textContent).includes("50%"));
  assert.match(pie.querySelector(".hs-oslice").getAttribute("d"), /^M450 250 L/, "a pie slice starts at the centre");
  // 100% stacked bars: each category fills the axis; the axis reads 0%–100%.
  const stacked = draw({ type: "100-stacked-bar", labels: ["20代", "30代"], series: [{ name: "満足", values: [55, 60] }, { name: "不満", values: [45, 20] }], style: { type: "100-stacked-bar", stack: "percent", dir: "bar" } });
  const texts = [...stacked.querySelectorAll("text")].map((t) => t.textContent);
  assert.ok(texts.includes("0%") && texts.includes("100%") && !texts.includes("120%"), texts.join(" "));
  const [b1, , b3] = [...stacked.querySelectorAll(".hs-obar")];
  assert.ok(Math.abs(Number(b1.getAttribute("width")) / Number(b3.getAttribute("width")) - 55 / 45) < 0.01);
  // Bars and a line together.
  const combo = draw({ type: "combo", labels: ["4月", "5月"], series: [{ name: "売上", values: [120, 135] }, { name: "利益率", values: [8, 9] }], style: { type: "combo", series: [{ kind: "bar" }, { kind: "line" }] } });
  assert.equal(combo.querySelectorAll(".hs-obar").length, 2);
  assert.equal(combo.querySelectorAll(".hs-oline").length, 1);
  // After the kind is changed in the editor, the old per-series kinds no longer apply.
  const changed = draw({ type: "line", labels: ["4月", "5月"], series: [{ name: "売上", values: [120, 135] }], style: { type: "bar", dir: "bar", series: [{ kind: "bar", color: "#1f3864" }] } });
  assert.equal(changed.querySelectorAll(".hs-obar").length, 0);
  assert.equal(changed.querySelector(".hs-oline").getAttribute("stroke"), "#1f3864", "the colour stays");
});

test("an imported chart's marks say their colour (data-paint)", async () => {
  const { E } = await loadEngine();
  const chart = { type: "bar", labels: ["A"], series: [{ name: "s", values: [1] }], style: { type: "bar", series: [{ color: "#4472c4" }] } };
  const svg = E.officeChart(E.normalizeObjects([{ id: "c", kind: "chart", x: 0, y: 0, w: 300, h: 200, chart }])[0].chart, 300, 200);
  assert.equal(svg.querySelector(".hs-obar").getAttribute("data-paint"), "#4472c4", "the mark's colour");
});
