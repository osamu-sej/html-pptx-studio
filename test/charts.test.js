// The studio's chart kinds (objects.js CHART_KINDS, drawn by engine.js): 面・円・散布図・レーダー・ウォーターフォール・じょうご
// as well as the bars, lines and donut — each drawn with the parts the グラフが伸びる animation moves.
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
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  return window.SlideEngine;
}
const draw = (E, chart) => {
  const o = E.normalizeObject({ id: "c", kind: "chart", x: 100, y: 100, w: 960, h: 560, chart });
  const slide = { type: "blank", elements: [o] };
  return { o, el: E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } }) };
};

test("every PowerPoint chart kind the studio offers is kept and drawn", async () => {
  const E = await loadEngine();
  for (const kind of ["area", "pie", "scatter", "radar", "waterfall", "funnel"]) assert.ok(E.CHART_KINDS[kind], kind);
  const two = { labels: ["A", "B", "C", "D"], series: [{ name: "今年", values: [3, 5, 4, 6] }, { name: "前年", values: [2, 4, 3, 5] }] };

  let { o, el } = draw(E, { type: "area", ...two });
  assert.equal(o.chart.type, "area");
  assert.equal(el.querySelectorAll(".hs-chart .hs-oarea").length, 2, "an area per series");
  assert.equal(el.querySelectorAll(".hs-chart .hs-draw").length, 2);
  assert.equal(el.querySelectorAll(".hs-legend span").length, 2, "a legend for two series");

  ({ el } = draw(E, { type: "pie", labels: ["来店", "アプリ", "宅配"], series: [{ name: "構成比", values: [60, 25, 15] }] }));
  const arcs = [...el.querySelectorAll(".hs-chart .hs-arc")];
  assert.equal(arcs.length, 3, "a slice each");
  assert.equal(Number(arcs[0].getAttribute("stroke-width")), Number(arcs[0].getAttribute("r")) * 2, "no hole");

  ({ el } = draw(E, { type: "scatter", labels: ["10", "20", "30"], series: [{ name: "y", values: [1, 4, 9] }] }));
  const dots = [...el.querySelectorAll(".hs-chart .hs-dot")];
  assert.equal(dots.length, 3);
  assert.ok(Number(dots[2].getAttribute("cx")) > Number(dots[0].getAttribute("cx")) && Number(dots[2].getAttribute("cy")) < Number(dots[0].getAttribute("cy")), "X from the labels, Y from the values");

  ({ el } = draw(E, { type: "radar", labels: ["a", "b", "c", "d", "e"], series: two.series.map((s) => ({ ...s, values: [...s.values, 2] })) }));
  assert.equal(el.querySelectorAll(".hs-chart polygon.hs-oarea").length, 2);
  assert.equal(el.querySelectorAll(".hs-chart polygon.hs-grid").length, 4, "four rings");

  ({ el } = draw(E, { type: "waterfall", labels: ["前年", "増", "減", "合計"], series: [{ name: "利益", values: [100, 30, -20, 0] }] }));
  const bars = [...el.querySelectorAll(".hs-chart rect.hs-bar")];
  assert.equal(bars.length, 4);
  assert.ok(bars[2].hasAttribute("data-neg"), "a decrease");
  assert.equal(bars[3].getAttribute("fill"), "var(--accent)", "the total stands on the axis in navy");
  // The decrease runs from 130 down to 110; the total's top is at 110, where the decrease ends.
  const bottomOf = (b) => Number(b.getAttribute("y")) + Number(b.getAttribute("height"));
  assert.ok(Math.abs(Number(bars[3].getAttribute("y")) - bottomOf(bars[2])) < 1, "the total reaches the running sum (110)");

  ({ el } = draw(E, { type: "funnel", labels: ["認知", "来店", "購入"], series: [{ name: "人", values: [1000, 400, 100] }] }));
  const widths = [...el.querySelectorAll(".hs-chart rect.hs-bar")].map((b) => Number(b.getAttribute("width")));
  assert.ok(widths[0] > widths[1] && widths[1] > widths[2], "narrowing stages");
  assert.ok([...el.querySelectorAll(".hs-chart text")].some((t) => t.textContent.includes("40%")), "the share of the first stage");
});

test("outline to slides, album boxes and fitting pictures (slidetools.mjs)", async () => {
  const { outlineToSlides, albumBoxes, fitIn } = await import("../public/editor/slidetools.mjs");
  assert.deepEqual(outlineToSlides("売上の現状\n  前年比112%\n  - 東日本が伸びている\n来期の打ち手\n・品揃え\n# まとめ"), [
    { type: "content", title: "売上の現状", points: ["前年比112%", "東日本が伸びている"] },
    { type: "content", title: "来期の打ち手", points: ["品揃え"] },
    { type: "section", title: "まとめ" },
  ]);
  assert.deepEqual(outlineToSlides("## A\n1. one\n2) two"), [{ type: "content", title: "A", points: ["one", "two"] }]);
  for (const per of [1, 2, 4]) {
    const boxes = albumBoxes(per, { captions: true });
    assert.equal(boxes.length, per);
    for (const b of boxes) assert.ok(b.x >= 120 && b.x + b.w <= 1800 && b.y >= 230 && b.y + b.h + 56 <= 950, `${per}: ${JSON.stringify(b)}`);
  }
  const fit = fitIn({ x: 0, y: 0, w: 400, h: 400 }, 800, 400);
  assert.deepEqual(fit, { x: 0, y: 100, w: 400, h: 200 });
});

test("video plan: saved timings or the default seconds, slide by slide (video.mjs); SRT to WebVTT (media.mjs)", async () => {
  const { videoPlan } = await import("../public/editor/video.mjs");
  const slides = [{ advance: 2 }, {}, { advance: 0.5 }];
  assert.deepEqual(videoPlan([0, 1, 2], slides, { seconds: 4 }), { plan: [{ index: 0, ms: 2000 }, { index: 1, ms: 4000 }, { index: 2, ms: 500 }], total: 6500 });
  assert.equal(videoPlan([0, 1, 2], slides, { seconds: 4, useTimings: false }).total, 12000);
  const { toVtt } = await import("../public/editor/media.mjs");
  assert.equal(toVtt("1\r\n00:00:01,000 --> 00:00:02,500\r\nこんにちは\r\n"), "WEBVTT\n\n1\n00:00:01.000 --> 00:00:02.500\nこんにちは");
  assert.equal(toVtt("﻿WEBVTT\n\n00:01.000 --> 00:02.000\nhi"), "WEBVTT\n\n00:01.000 --> 00:02.000\nhi");
  const E = await loadEngine();
  const o = E.normalizeObject({ id: "v", kind: "video", src: "https://example.com/a.mp4", x: 0, y: 0, w: 640, h: 360, captions: toVtt("1\n00:00:01,000 --> 00:00:02,000\n字幕"), captionLang: "ja" });
  assert.match(o.captions, /^WEBVTT/);
  const slide = { type: "blank", elements: [o] };
  const el = E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  const track = el.querySelector("video track");
  assert.ok(track && track.getAttribute("kind") === "captions" && track.getAttribute("src").startsWith("data:text/vtt"), "a captions track");
  assert.equal(E.normalizeObject({ id: "v", kind: "video", src: "https://example.com/a.mp4", x: 0, y: 0, w: 640, h: 360, captions: "not vtt" }).captions, undefined);
});

test("chart elements: no data labels, the legend below / at the right / none, no gridlines", async () => {
  const E = await loadEngine();
  const two = { labels: ["A", "B", "C"], series: [{ name: "今年", values: [3, 5, 4] }, { name: "前年", values: [2, 4, 3] }] };
  const kept = E.normalizeObject({ id: "c", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { type: "clustered-bar", ...two, opts: { labels: false, legend: "right", grid: false, junk: 1 } } }).chart.opts;
  assert.deepEqual({ ...kept }, { labels: false, legend: "right", grid: false });
  let { el } = draw(E, { type: "clustered-bar", ...two });
  assert.ok(el.querySelectorAll(".hs-chart .hs-val").length > 0, "values by default");
  assert.ok(el.querySelector(".hs-chart-wrap > .hs-legend:first-child"), "the legend on top by default");
  ({ el } = draw(E, { type: "clustered-bar", ...two, opts: { labels: false, legend: "bottom" } }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-val").length, 0, "no data labels");
  assert.ok(el.querySelector(".hs-chart-wrap > .hs-legend:last-child.at-bottom"), "the legend below");
  ({ el } = draw(E, { type: "area", ...two, opts: { legend: "none", grid: false } }));
  assert.equal(el.querySelectorAll(".hs-legend").length, 0, "no legend");
  assert.equal(el.querySelectorAll(".hs-chart line.hs-grid").length, 0, "no gridlines");
  assert.ok(el.querySelector(".hs-chart line.hs-axisline"), "the axis stays");
});

const plain = (v) => (v === undefined ? undefined : JSON.parse(JSON.stringify(v)));

test("histogram bins, box statistics, treemap layout and 親/子 labels", async () => {
  const E = await loadEngine();
  const bins = E.histogramBins([2, 3, 3, 4, 4, 4, 5, 5, 5, 5, 6, 6, 6, 7, 7, 7, 8, 8, 9, 10, 11, 12, 14, 18]);
  assert.equal(bins.reduce((a, b) => a + b.count, 0), 24, "every value counted once");
  const width = bins[0].to - bins[0].from;
  assert.ok(bins.every((b) => Math.abs(b.to - b.from - width) < 1e-9), "bins of one width");
  assert.ok([1, 2, 2.5, 5].some((m) => Math.abs(width / 10 ** Math.floor(Math.log10(width)) - m) < 1e-9), `a round width: ${width}`);
  assert.ok(bins[0].from <= 2 && bins.at(-1).to > 18, "the range covered");
  assert.deepEqual(plain(E.histogramBins([])), []);

  const st = E.boxStats([62, 68, 70, 71, 74, 77, 80, 95, 200]);
  assert.equal(st.median, 74);
  assert.equal(st.q1, 70);
  assert.equal(st.q3, 80);
  assert.deepEqual(plain(st.outliers), [200], "beyond 1.5 IQR");
  assert.equal(st.max, 95, "the whisker stops at the last value inside");
  assert.equal(E.boxStats([]), null);

  const cells = E.squarify([{ name: "a", value: 6 }, { name: "b", value: 3 }, { name: "c", value: 1 }, { name: "z", value: 0 }], 0, 0, 100, 50);
  assert.equal(cells.length, 3, "zero leaves no cell");
  const area = (n) => { const c = cells.find((x) => x.name === n); return c.w * c.h; };
  assert.ok(Math.abs(area("a") - 3000) < 1e-6 && Math.abs(area("b") - 1500) < 1e-6 && Math.abs(area("c") - 500) < 1e-6, "areas in proportion");
  assert.ok(cells.every((c) => c.x >= -1e-9 && c.y >= -1e-9 && c.x + c.w <= 100 + 1e-9 && c.y + c.h <= 50 + 1e-9), "inside the box");

  const tree = E.hierarchy(["食品/おにぎり", "食品／弁当", "飲料/お茶", "その他"], [4, 3, 2, 1]);
  assert.deepEqual(plain(tree.map((n) => [n.name, n.value, n.children.length])), [["食品", 7, 2], ["飲料", 2, 1], ["その他", 1, 0]]);
});

test("the new chart kinds are kept and drawn, with 軸ラベル and 近似曲線", async () => {
  const E = await loadEngine();
  const two = { labels: ["東", "中", "西"], series: [{ name: "今期", values: [3, 5, 4] }, { name: "前期", values: [2, 4, 3] }] };
  let { o, el } = draw(E, { type: "hbar", ...two });
  assert.equal(o.chart.type, "hbar");
  assert.equal(el.querySelectorAll(".hs-chart .hs-bar.h").length, 6, "a horizontal bar per label and series");
  ({ el } = draw(E, { type: "stacked-hbar", ...two }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-bar.h").length, 6);
  assert.equal(el.querySelectorAll(".hs-legend span").length, 2, "a legend for the stacked series");
  ({ el } = draw(E, { type: "stacked-area", ...two }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-oarea").length, 2);
  ({ el } = draw(E, { type: "bubble", labels: ["10", "20", "30"], series: [{ name: "Y", values: [5, 9, 7] }, { name: "大きさ", values: [1, 9, 4] }] }));
  const rs = [...el.querySelectorAll(".hs-chart circle.hs-dot")].map((c) => Number(c.getAttribute("r")));
  assert.equal(rs.length, 3);
  assert.ok(Math.max(...rs) > Math.min(...rs) * 2, "bubbles sized by the second series");
  assert.equal(el.querySelectorAll(".hs-legend").length, 0, "no legend");
  ({ el } = draw(E, { type: "histogram", labels: ["1", "2", "3", "4", "5", "6"], series: [{ name: "分", values: [1, 2, 2, 3, 3, 9] }] }));
  assert.ok(el.querySelectorAll(".hs-chart .hs-bar").length >= 2, "bins as bars");
  ({ el } = draw(E, { type: "boxplot", labels: ["1", "2", "3", "4", "5"], series: [{ name: "A", values: [1, 2, 3, 4, 5] }, { name: "B", values: [2, 3, 4, 5, 9] }] }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-bar[data-center]").length, 2, "a box per series");
  assert.equal(el.querySelectorAll(".hs-chart .hs-median").length, 2);
  const hier = { labels: ["食品/おにぎり", "食品/弁当", "飲料/お茶"], series: [{ name: "売上", values: [4, 3, 2] }] };
  ({ el } = draw(E, { type: "treemap", ...hier }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-bar[data-center]").length, 3, "a cell per member");
  assert.equal(el.querySelectorAll(".hs-chart .hs-tm-group").length, 2, "the groups named");
  ({ el } = draw(E, { type: "sunburst", ...hier }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-ochart-pie .hs-oslice").length, 5, "2 groups inside, 3 members outside");

  ({ o, el } = draw(E, { type: "line", labels: ["1", "2", "3", "4"], series: [{ name: "Y", values: [1, 3, 2, 5] }], opts: { axisX: "月", axisY: "売上（億円）", trend: "linear", bogus: 1 } }));
  assert.deepEqual(plain(o.chart.opts), { axisX: "月", axisY: "売上（億円）", trend: "linear" });
  assert.deepEqual([...el.querySelectorAll(".hs-chart .hs-axis-title")].map((t) => t.textContent), ["売上（億円）", "月"]);
  assert.equal(el.querySelectorAll(".hs-chart .hs-trend").length, 1, "the trendline");
  ({ el } = draw(E, { type: "pie", labels: ["a", "b"], series: [{ name: "x", values: [1, 2] }], opts: { axisX: "月" } }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-axis-title").length, 0, "a pie has no axes");
});

test("動作設定ボタン: twelve of them, the ones that move through the show with their click set", async () => {
  const E = await loadEngine();
  const group = E.SHAPE_GROUPS.find(([name]) => name === "動作設定ボタン");
  assert.equal(group?.[1].length, 12);
  assert.deepEqual(plain(E.SHAPES.actionButtonForwardNext.action), { type: "next" });
  assert.deepEqual(plain(E.SHAPES.actionButtonBeginning.action), { type: "first" });
  assert.equal(E.SHAPES.actionButtonHelp.action, undefined);
  for (const key of group[1]) assert.ok(E.geometry(key, 100, 100).paths.length, key);
  assert.ok(E.geometry("actionButtonForwardNext", 100, 100).extras.some((x) => x.tone === "dark"), "its sign drawn darker");
});

test("軸の書式: the value axis's own minimum and maximum; values beyond it stay on the edge", async () => {
  const E = await loadEngine();
  const data = { labels: ["4月", "5月", "6月"], series: [{ name: "売上", values: [60, 80, 140] }] };
  const { o, el } = draw(E, { type: "line", ...data, opts: { axisMin: 50, axisMax: 100 } });
  assert.equal(o.chart.opts.axisMin, 50);
  assert.equal(o.chart.opts.axisMax, 100);
  const ticks = [...el.querySelectorAll(".hs-tick")].map((t) => t.textContent);
  assert.equal(ticks[0], "50", `from the minimum: ${ticks}`);
  assert.equal(ticks.at(-1), "100", `to the maximum: ${ticks}`);
  const top = Math.min(...[...el.querySelectorAll("line.hs-grid, line.hs-axisline")].map((l) => Number(l.getAttribute("y1"))));
  const dots = [...el.querySelectorAll("circle")].map((c) => Number(c.getAttribute("cy")));
  assert.ok(dots.length && Math.min(...dots) >= top - 0.01, `140 stays on the top edge (${Math.min(...dots)} vs ${top})`);
  // A maximum not above the minimum is dropped; bars keep their zero line and stop at the maximum.
  assert.equal(draw(E, { type: "line", ...data, opts: { axisMin: 100, axisMax: 50 } }).o.chart.opts.axisMax, undefined);
  const bars = draw(E, { type: "bar", ...data, opts: { axisMax: 100 } }).el;
  // A bar's top is its value label's place (16 above it); the zero line is the axis line.
  const y0 = Number(bars.querySelector("line.hs-axisline").getAttribute("y1"));
  const heights = [...bars.querySelectorAll("text.hs-val")].map((t) => y0 - Number(t.getAttribute("y")) - 16);
  assert.ok(heights.length === 3 && heights[1] / heights[2] > 0.79 && heights[1] / heights[2] < 0.81 && heights[0] / heights[2] > 0.59 && heights[0] / heights[2] < 0.61, `bars on a 0–100 axis (140 stops at the top): ${heights}`);
});

test("色の変更: a chart takes one of PowerPoint's colour sets, made from the theme's accents", async () => {
  const E = await loadEngine();
  const data = { labels: ["A", "B"], series: [{ name: "一", values: [1, 2] }, { name: "二", values: [2, 1] }] };
  const { o, el } = draw(E, { type: "clustered-bar", ...data, colors: "colorful2" });
  assert.equal(o.chart.colors, "colorful2");
  const box = el.querySelector(".hs-obj-rot > .hs-obj-chart");
  assert.equal(box.getAttribute("data-colors"), "colorful2");
  const accents = E.themeColors(E.DEFAULT_THEME);
  assert.match(box.getAttribute("style") || "", new RegExp(`--c1: ?${accents.accent2}`), "カラフル starts at アクセント 2");
  assert.equal(draw(E, { type: "bar", ...data, colors: "rainbow" }).o.chart.colors, undefined, "only the sets offered");
  // Every colour of every set is a theme colour or one of its tones, so it follows a change of theme.
  for (const theme of E.THEMES) {
    for (const key of Object.keys(E.CHART_COLORS)) {
      for (const [name, value] of Object.entries(E.chartColorVars(key, theme.id))) {
        if (name === "--c-muted") continue;
        assert.ok(E.themeColorRef(value, theme.id), `${theme.id} ${key} ${name} ${value} is a theme colour`);
      }
    }
  }
  const mono = E.chartColorVars("mono1", E.DEFAULT_THEME);
  assert.equal(mono["--c1"], accents.accent1, "モノクロ starts at the accent itself");
  assert.equal(new Set(Object.entries(mono).filter(([k]) => /^--c\d$/.test(k)).map(([, v]) => v)).size, 6, "six different tones");
});

test("クイック レイアウト: a set of chart elements at once; the title, unit and axis bounds stay", async () => {
  const E = await loadEngine();
  const norm = (c) => E.normalizeObject({ id: "c", kind: "chart", x: 0, y: 0, w: 960, h: 560, chart: c }).chart;
  const chart = { type: "bar", title: "売上", unit: "億円", labels: ["4月", "5月"], series: [{ name: "売上", values: [3, 5] }, { name: "目標", values: [4, 4] }], opts: { axisMax: 10, legend: "right", grid: false } };
  assert.equal(E.chartLayoutOf(norm(chart)), "layout2", "data labels with the legend on the right");
  const four = norm(E.applyChartLayout(chart, "layout4"));
  assert.equal(JSON.stringify(four.opts), JSON.stringify({ legend: "none", grid: false, axisMax: 10 }));
  assert.equal(four.title, "売上");
  assert.equal(E.chartLayoutOf(four), "layout4");
  const five = norm(E.applyChartLayout(four, "layout5"));
  assert.equal(five.opts.axisX, "項目", "axis titles come with the layout");
  assert.equal(five.opts.axisY, "億円", "named after the unit");
  assert.equal(five.opts.labels, false);
  assert.equal(five.opts.legend, undefined, "the legend on top is the default");
  assert.equal(E.chartLayoutOf(five), "layout5");
  const one = norm(E.applyChartLayout({ ...five, opts: { ...five.opts, axisX: "月" } }, "layout1"));
  assert.equal(one.opts?.axisX, undefined, "a layout without axis titles takes them away");
  assert.equal(one.opts?.axisMax, 10);
  assert.equal(E.chartLayoutOf(one), "layout1");
  assert.equal(E.applyChartLayout(chart, "nope"), chart, "an unknown layout changes nothing");
  for (const key of Object.keys(E.CHART_LAYOUTS)) assert.equal(E.chartLayoutOf(norm(E.applyChartLayout(chart, key))), key, `${key} is found again`);
});

test("グラフ スタイル: colours, elements and the bars' gap or the line's curve at once; everything else of the chart stays", async () => {
  const E = await loadEngine();
  const norm = (c) => E.normalizeObject({ id: "c", kind: "chart", x: 0, y: 0, w: 960, h: 560, chart: c }).chart;
  const bar = { type: "bar", title: "売上", unit: "億円", labels: ["4月", "5月"], series: [{ name: "売上", values: [3, 5] }, { name: "目標", values: [4, 4] }], opts: { axisMax: 10, axisX: "月", trend: "linear", gap: 33, hideSeries: [1] } };
  assert.equal(E.chartStyleOf(norm(bar)), "", "a gap of its own is no style");
  const three = norm(E.applyChartStyle(bar, "style3"));
  assert.equal(three.colors, "mono1");
  assert.equal(three.opts.gap, 20);
  assert.equal(three.opts.legend, undefined, "the legend on top is the default");
  assert.equal(three.title, "売上");
  assert.equal(three.unit, "億円");
  assert.equal(three.opts.axisMax, 10, "the axis bounds stay");
  assert.equal(three.opts.axisX, "月", "the axis titles stay");
  assert.equal(three.opts.trend, "linear", "the trendline stays");
  assert.equal(JSON.stringify(three.opts.hideSeries), "[1]", "the filter stays");
  assert.equal(E.chartStyleOf(three), "style3");
  const four = norm(E.applyChartStyle(three, "style4"));
  assert.equal(four.colors, "gray");
  assert.equal(four.opts.labels, false);
  assert.equal(four.opts.legend, "bottom");
  assert.equal(four.opts.gap, 200);
  const one = norm(E.applyChartStyle(four, "style1"));
  assert.equal(one.colors, undefined, "the theme's own order again");
  assert.equal(one.opts.gap, undefined, "the standard gap again");
  assert.equal(one.opts.labels, undefined);
  assert.equal(E.chartStyleOf(one), "style1");
  for (const key of Object.keys(E.CHART_STYLES)) assert.equal(E.chartStyleOf(norm(E.applyChartStyle(bar, key))), key, `${key} is found again on bars`);
  // A line has no gap but a curve and markers; a pie has neither.
  const line = { type: "multi-line", labels: ["a", "b", "c"], series: [{ name: "s", values: [1, 3, 2] }, { name: "t", values: [2, 1, 4] }] };
  const eight = norm(E.applyChartStyle(line, "style8"));
  assert.equal(eight.opts.smooth, true);
  assert.equal(eight.opts.marker, "diamond");
  assert.equal(eight.opts.gap, undefined, "no gap on a line");
  assert.equal(E.chartStyleOf(eight), "style8");
  const two = norm(E.applyChartStyle(eight, "style2"));
  assert.equal(two.opts.smooth, undefined, "a style without the curve takes it away");
  assert.equal(two.opts.marker, undefined);
  for (const key of Object.keys(E.CHART_STYLES)) assert.equal(E.chartStyleOf(norm(E.applyChartStyle(line, key))), key, `${key} is found again on lines`);
  const pie = norm(E.applyChartStyle({ type: "pie", labels: ["a", "b"], series: [{ name: "s", values: [1, 2] }] }, "style5"));
  assert.equal(pie.colors, "gray");
  assert.equal(pie.opts.legend, "right");
  assert.equal(pie.opts.gap, undefined);
  // An imported chart keeps its own formatting; an unknown style changes nothing.
  const office = { ...bar, style: { x: 1 } };
  assert.equal(E.applyChartStyle(office, "style2"), office);
  assert.equal(E.chartStyleOf(office), "");
  assert.equal(E.applyChartStyle(bar, "nope"), bar);
  // Every colour set a style names exists, and every legend place is one the chart knows.
  for (const s of Object.values(E.CHART_STYLES)) {
    assert.ok(s.colors === "" || E.CHART_COLORS[s.colors], s.label);
    assert.ok(["top", "bottom", "right", "none"].includes(s.legend), s.label);
    assert.ok(s.gap == null || (s.gap >= 0 && s.gap <= 500), s.label);
    assert.ok(!s.marker || E.MARKER_SHAPES.includes(s.marker), s.label);
  }
});

test("パレート図: the first series from the largest, with a line of the running share on a right axis that ends at 100%", async () => {
  const E = await loadEngine();
  const chart = { type: "pareto", labels: ["弁当", "おにぎり", "パン", "サラダ"], series: [{ name: "廃棄数", values: [30, 50, 10, 10] }, { name: "使わない", values: [1, 2, 3, 4] }] };
  const { o, el } = draw(E, chart);
  assert.equal(o.chart.type, "pareto");
  const spec = E.chartSpec(o.chart);
  assert.equal(spec.chartType, "combo", "drawn as bars and a line");
  assert.deepEqual(plain(spec.data.items.map((i) => i.label)), ["おにぎり", "弁当", "パン", "サラダ"], "from the largest (a tie keeps its order)");
  assert.deepEqual(plain(spec.data.items.map((i) => i.barValue)), [50, 30, 10, 10]);
  assert.deepEqual(plain(spec.data.items.map((i) => i.value)), [50, 80, 90, 100], "the running share, 100 at the end");
  assert.deepEqual(plain(spec.data.legendLabels), ["廃棄数", "累積比率"]);
  assert.equal(spec.opts.axis2, true);
  assert.equal(spec.opts.axis2Pct, true);
  assert.equal(o.chart.opts, undefined, "what makes it a pareto is not kept on the chart itself");
  // Drawn: the labels in the new order, the right axis in %, the left axis ending at the total (100).
  const texts = [...el.querySelectorAll(".hs-chart svg text")].map((t) => t.textContent);
  assert.deepEqual(texts.filter((t) => ["おにぎり", "弁当", "パン", "サラダ"].includes(t)), ["おにぎり", "弁当", "パン", "サラダ"]);
  const right = [...el.querySelectorAll(".hs-tick2")].map((t) => t.textContent);
  assert.deepEqual(right, ["0%", "25%", "50%", "75%", "100%"]);
  assert.ok([...el.querySelectorAll(".hs-chart svg text")].some((t) => t.textContent === "100%" && t.getAttribute("class")?.includes("hs-val")), "the last point says 100%");
  assert.ok(texts.includes("100"), `the left axis ends at the total: ${texts.join(",")}`);
  assert.equal(el.querySelectorAll(".hs-bar").length, 4);
  // Negative values count as nothing; an empty total does not divide by zero.
  const odd = E.chartSpec(E.normalizeObject({ id: "d", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { type: "pareto", labels: ["a", "b"], series: [{ name: "s", values: [-5, 0] }] } }).chart);
  assert.deepEqual(plain(odd.data.items.map((i) => i.value)), [0, 0]);
  // The gap works on its bars like any bar chart's; the kind is one only the studio draws (not a PowerPoint chart's own look).
  assert.ok(E.BAR_GAP_CHARTS.has("pareto"));
  assert.equal(E.normalizeObject({ id: "e", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { ...chart, opts: { gap: 40 } } }).chart.opts.gap, 40);
});

test("グラフ フィルター hides series and categories without losing them; データ テーブル shows the values under the chart", async () => {
  const E = await loadEngine();
  const chart = { type: "waterfall", labels: ["期首", "増", "減", "期末"], series: [{ name: "売上", values: [10, 5, -3, 12] }, { name: "目標", values: [9, 4, -2, 11] }],
    opts: { totals: [0, 3], hideSeries: [1], hideLabels: [1, 99], table: true } };
  const { o, el } = draw(E, chart);
  assert.equal(JSON.stringify(o.chart.opts), JSON.stringify({ totals: [0, 3], hideSeries: [1], hideLabels: [1], table: true }), "only places that exist");
  assert.equal(o.chart.series.length, 2, "the data stays");
  const shown = E.visibleChart(o.chart);
  assert.equal(JSON.stringify(shown.labels), JSON.stringify(["期首", "減", "期末"]));
  assert.equal(JSON.stringify(shown.series.map((s) => s.values)), JSON.stringify([[10, -3, 12]]));
  assert.equal(JSON.stringify(shown.opts.totals), JSON.stringify([0, 2]), "the totals follow the shown categories");
  const table = el.querySelector(".hs-obj-chart .hs-chart-table");
  assert.ok(table, "the data table");
  assert.equal([...table.querySelectorAll("thead th")].map((th) => th.textContent).join("|"), "|期首|減|期末");
  assert.equal([...table.querySelectorAll("tbody tr")].map((tr) => tr.textContent).join("|"), "売上10-312");
  const all = E.normalizeObject({ id: "c2", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { ...chart, opts: { hideSeries: [0, 1], hideLabels: [0, 1, 2, 3] } } });
  assert.equal(all.chart.opts, undefined, "everything hidden is not a filter");
  const pie = draw(E, { type: "pie", labels: ["a", "b"], series: [{ name: "s", values: [1, 2] }], opts: { table: true } });
  assert.equal(pie.el.querySelector(".hs-chart-table"), null, "no table for a pie");
});

test("第2軸: a combination chart's line reads on a scale of its own, with ticks down the right side", async () => {
  const E = await loadEngine();
  const data = { type: "combo", labels: ["4月", "5月", "6月"], series: [{ name: "売上", values: [12000, 15000, 18000] }, { name: "前年比", values: [96, 102, 108] }] };
  assert.equal(E.normalizeObject({ id: "a", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { ...data, opts: { axis2: true } } }).chart.opts.axis2, true);
  assert.equal(E.normalizeObject({ id: "b", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { ...data, type: "line", opts: { axis2: true } } }).chart.opts, undefined, "only a combination chart");
  assert.equal(E.normalizeObject({ id: "c", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { ...data, series: [data.series[0]], opts: { axis2: true } } }).chart.opts, undefined, "it takes two series");
  const heights = (el) => [...el.querySelectorAll(".hs-chart circle.hs-mark")].map((c) => Number(c.getAttribute("cy")));
  const { el: shared } = draw(E, data);
  const { el: second } = draw(E, { ...data, opts: { axis2: true } });
  // On the bars' scale (12000…18000) a line of 96…108 sits on the floor; on its own scale it climbs.
  const flat = heights(shared);
  const climbs = heights(second);
  assert.ok(Math.max(...flat) - Math.min(...flat) < 2, `flat on a shared axis: ${flat}`);
  assert.ok(Math.max(...climbs) - Math.min(...climbs) > 40, `a line of its own: ${climbs}`);
  const ticks = [...second.querySelectorAll(".hs-chart .hs-tick2")].map((t) => t.textContent);
  assert.equal(ticks.length, 5, "five marks down the right");
  assert.ok(Number(ticks.at(-1)) >= 108 && Number(ticks.at(-1)) < 200, `the line's scale: ${ticks}`);
  assert.ok(second.querySelectorAll(".hs-chart .hs-tick:not(.hs-tick2)").length >= 5, "the bars' scale down the left");
  assert.equal(shared.querySelectorAll(".hs-chart .hs-tick2").length, 0);
});

test("数値の書式: decimals, a display unit and signs reach the labels, the tips and the data table; an axis takes the unit only", async () => {
  const E = await loadEngine();
  const base = { type: "line", labels: ["4月", "5月", "6月"], series: [{ name: "売上", values: [12000, 15000, 24000] }] };
  const norm = (numFmt) => E.normalizeObject({ id: "n", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { ...base, opts: { numFmt } } }).chart.opts?.numFmt;
  assert.deepEqual(plain(norm({ decimals: "1", scale: "10000", prefix: "¥<b>", suffix: " 円 " })), { decimals: 1, scale: 10000, prefix: "¥b", suffix: "円" }, "numbers and clean text only");
  assert.deepEqual(plain(norm({ decimals: 9, scale: 7, prefix: "", suffix: "   " })), undefined, "unknown values are dropped");
  assert.equal(norm({}), undefined);
  const { el } = draw(E, { ...base, opts: { numFmt: { decimals: 1, scale: 10000, prefix: "¥", suffix: "円" } } });
  const tips = [...el.querySelectorAll(".hs-chart circle.hs-mark")].map((c) => c.getAttribute("data-tip"));
  assert.ok(tips[0].endsWith("¥1.2万円"), `a tip: ${tips[0]}`);
  assert.ok(tips[2].endsWith("¥2.4万円"), tips[2]);
  const labels = [...el.querySelectorAll(".hs-chart .hs-val")].map((t) => t.textContent);
  assert.ok(labels.includes("¥2.4万円"), `a label: ${labels}`);
  const ticks = [...el.querySelectorAll(".hs-chart .hs-tick")].map((t) => t.textContent).filter((t) => /\d/.test(t));
  assert.ok(ticks.length >= 4 && ticks.every((t) => !/[¥円]/.test(t) && /万$|^0/.test(t)), `ticks carry the unit, not the signs: ${ticks}`);
  // The data table follows it too; without a format everything stays as it was.
  const { el: withTable } = draw(E, { ...base, type: "bar", opts: { table: true, numFmt: { scale: 1000, suffix: "円" } } });
  assert.ok([...withTable.querySelectorAll(".hs-chart-table td")].some((td) => td.textContent === "12千円"), "the table");
  const { el: plainChart } = draw(E, base);
  assert.ok([...plainChart.querySelectorAll(".hs-chart .hs-val")].some((t) => t.textContent === "24,000"), "as before");
});

test("誤差範囲: a bar and two caps on each mark, a fixed amount or a share of the value", async () => {
  const E = await loadEngine();
  const data = { labels: ["A", "B", "C"], series: [{ name: "値", values: [100, 200, 300] }] };
  const kept = (type, opts) => E.normalizeObject({ id: "e", kind: "chart", x: 0, y: 0, w: 900, h: 500, chart: { ...data, type, opts } }).chart.opts;
  assert.deepEqual(plain(kept("bar", { errorBars: { type: "fixed", amount: "20" } }).errorBars), { type: "fixed", amount: 20 });
  assert.equal(kept("pie", { errorBars: { type: "fixed", amount: 20 } }), undefined, "only bars and lines");
  assert.equal(kept("bar", { errorBars: { type: "percent", amount: 150 } }), undefined, "a share of at most 100");
  assert.equal(kept("bar", { errorBars: { type: "fixed", amount: 0 } }), undefined);
  assert.equal(kept("bar", { errorBars: { type: "other", amount: 5 } }), undefined);
  const lines = (type, opts, series = data.series) => draw(E, { ...data, type, series, opts }).el.querySelectorAll(".hs-chart line.hs-err");
  assert.equal(lines("bar", { errorBars: { type: "fixed", amount: 20 } }).length, 9, "3 marks × (bar + 2 caps)");
  assert.equal(lines("bar").length, 0);
  assert.equal(lines("line", { errorBars: { type: "percent", amount: 10 } }).length, 9);
  const two = [...data.series, { name: "別", values: [90, 180, 270] }];
  assert.equal(lines("clustered-bar", { errorBars: { type: "fixed", amount: 10 } }, two).length, 18, "a bar for each of two series");
  assert.equal(lines("multi-line", { errorBars: { type: "fixed", amount: 10 } }, two).length, 18);
  // The value label sits above the error bar, not on it.
  const labelY = (opts) => Number(draw(E, { ...data, type: "bar", opts }).el.querySelector(".hs-chart .hs-val").getAttribute("y"));
  assert.ok(labelY({ errorBars: { type: "fixed", amount: 40 } }) < labelY(undefined), "the label rides above the error bar");
  // The bar's length follows the amount: ±20 on a 100 scale is longer than ±10.
  const span = (amount) => { const bar = lines("bar", { errorBars: { type: "fixed", amount } })[0]; return Math.abs(Number(bar.getAttribute("y2")) - Number(bar.getAttribute("y1"))); };
  assert.ok(span(40) > span(20) * 1.9 && span(40) < span(20) * 2.1, "double the amount, double the bar");
  // A share: 10% of 300 is longer than 10% of 100.
  const bars = [...lines("bar", { errorBars: { type: "percent", amount: 10 } })].filter((_, i) => i % 3 === 0).map((l) => Math.abs(Number(l.getAttribute("y2")) - Number(l.getAttribute("y1"))));
  assert.ok(bars[2] > bars[0] * 2.9 && bars[2] < bars[0] * 3.1, `a share of each value: ${bars}`);
});

test("近似曲線: the kinds PowerPoint offers are fitted on the data, with their equation and R²", async () => {
  const E = await loadEngine();
  assert.deepEqual(Object.keys(plain(E.TREND_KINDS)), ["linear", "exp", "log", "poly", "power", "movavg"]);
  const xs = [1, 2, 3, 4, 5];
  let fit = E.trendFit("linear", xs, xs.map((x) => 2 * x + 1));
  assert.equal(fit.eq, "y = 2x + 1");
  assert.ok(fit.r2 > 0.9999);
  assert.equal(fit.pts.length, 2, "a straight line needs two points");
  fit = E.trendFit("exp", xs, xs.map((x) => 3 * Math.exp(0.5 * x)));
  assert.equal(fit.eq, "y = 3e^0.5x");
  assert.equal(fit.pts.length, 49);
  fit = E.trendFit("log", xs, xs.map((x) => 2 * Math.log(x) + 1));
  assert.equal(fit.eq, "y = 2ln(x) + 1");
  fit = E.trendFit("power", xs, xs.map((x) => 2 * x ** 1.5));
  assert.equal(fit.eq, "y = 2x^1.5");
  fit = E.trendFit("poly", [0, 1, 2, 3, 4], [2, 0, 0, 2, 6]);
  assert.equal(fit.eq, "y = x² - 3x + 2");
  assert.ok(fit.r2 > 0.9999);
  fit = E.trendFit("movavg", xs, [1, 2, 3, 4, 5], { period: 3 });
  assert.deepEqual(plain(fit.pts), [[3, 2], [4, 3], [5, 4]]);
  assert.equal(fit.r2, null);
  assert.equal(E.trendFit("exp", xs, [1, 2, 0, 4, 5]), null, "指数 needs y above 0");
  assert.equal(E.trendFit("log", [0, 1, 2], [1, 2, 3]), null, "対数 needs x above 0");
  assert.equal(E.trendFit("poly", [1, 2], [1, 2]), null, "多項式 needs three points");
  assert.equal(E.trendFit("nope", xs, xs), null);

  const data = { labels: ["1", "2", "3", "4", "5"], series: [{ name: "Y", values: [3, 5, 9, 15, 25] }] };
  let { o, el } = draw(E, { type: "line", ...data, opts: { trend: "exp", trendEq: true, trendR2: true } });
  assert.deepEqual(plain(o.chart.opts), { trend: "exp", trendEq: true, trendR2: true });
  assert.equal(el.querySelectorAll(".hs-chart .hs-trend").length, 1);
  const words = el.querySelector(".hs-chart .hs-trend-eq").textContent;
  assert.ok(/^y = .*e\^.*x/.test(words) && words.includes("R² = "), words);
  ({ o, el } = draw(E, { type: "bar", ...data, opts: { trend: "poly" } }));
  assert.equal(o.chart.opts.trend, "poly");
  assert.equal(el.querySelectorAll(".hs-chart .hs-trend").length, 1);
  assert.equal(el.querySelectorAll(".hs-chart .hs-trend-eq").length, 0, "no equation unless asked");
  ({ o, el } = draw(E, { type: "line", ...data, opts: { trend: "movavg", trendPeriod: 3, trendEq: true } }));
  assert.deepEqual(plain(o.chart.opts), { trend: "movavg", trendPeriod: 3 }, "a moving average has no equation");
  ({ o } = draw(E, { type: "line", ...data, opts: { trendEq: true, trend: "wavy" } }));
  assert.equal(o.chart.opts, undefined, "an unknown kind, or an equation without a line, is dropped");
  ({ o } = draw(E, { type: "pie", labels: ["a", "b"], series: [{ name: "x", values: [1, 2] }], opts: { trend: "linear" } }));
  assert.equal(o.chart.opts, undefined, "a pie takes no trendline");
  ({ o } = draw(E, { type: "bar", ...data, opts: { trend: true } }));
  assert.equal(o.chart.opts.trend, "linear", "the earlier `true` still reads as a straight line");
  ({ el } = draw(E, { type: "scatter", labels: ["1", "2", "4", "8"], series: [{ name: "Y", values: [2, 3, 5, 9] }], opts: { trend: "log" } }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-trend").length, 1, "scatter, on its own X values");
});

test("棒の間隔の幅・系列の重なり: bars laid out as PowerPoint does, the built-in widths kept when unset", async () => {
  const E = await loadEngine();
  assert.deepEqual(plain(E.barLayout({}, 200, 2, 120)), { bar: 60, pitch: 60, group: 120 });
  assert.deepEqual(plain(E.barLayout({ gap: 100 }, 200, 1, 99)), { bar: 100, pitch: 100, group: 100 });
  const l = E.barLayout({ gap: 50, overlap: -20 }, 300, 2, 1);
  assert.ok(Math.abs(l.bar - 300 / 2.7) < 1e-9 && Math.abs(l.pitch - (300 / 2.7) * 1.2) < 1e-9);
  const same = E.barLayout({ overlap: 100 }, 300, 3, 1);
  assert.equal(same.pitch, 0, "100% overlap puts the bars on one another");
  const two = { labels: ["東", "中", "西"], series: [{ name: "今期", values: [3, 5, 4] }, { name: "前期", values: [2, 4, 3] }] };
  const xOf = (el) => [...el.querySelectorAll(".hs-chart .hs-bar")].map((b) => Number(b.getAttribute("d").match(/^M([-\d.]+),/)[1]));
  let { o, el } = draw(E, { type: "clustered-bar", ...two, opts: { gap: 20, overlap: 100 } });
  assert.deepEqual(plain(o.chart.opts), { gap: 20, overlap: 100 });
  let xs = xOf(el);
  assert.equal(xs[0], xs[1], "the two series of a group on one another");
  ({ el } = draw(E, { type: "clustered-bar", ...two }));
  xs = xOf(el);
  assert.ok(xs[1] > xs[0], "side by side without it");
  ({ o, el } = draw(E, { type: "bar", labels: ["A", "B", "C"], series: [{ name: "x", values: [1, 2, 3] }], opts: { gap: 0 } }));
  xs = xOf(el);
  assert.ok(Math.abs(xs[1] - xs[0] - (960 / 3)) < 1, "no gap: a bar fills its slot");
  assert.equal(o.chart.opts.gap, 0);
  ({ o } = draw(E, { type: "bar", labels: ["A", "B"], series: [{ name: "x", values: [1, 2] }], opts: { gap: 600, overlap: 10 } }));
  assert.equal(o.chart.opts, undefined, "a gap beyond 500% and an overlap on a single series are dropped");
  ({ o } = draw(E, { type: "stacked-bar", ...two, opts: { gap: 40, overlap: 30 } }));
  assert.deepEqual(plain(o.chart.opts), { gap: 40 }, "stacked bars always overlap");
  ({ o } = draw(E, { type: "hbar", ...two, opts: { gap: 40, overlap: -30 } }));
  assert.deepEqual(plain(o.chart.opts), { gap: 40, overlap: -30 });
  ({ o } = draw(E, { type: "pie", labels: ["a", "b"], series: [{ name: "x", values: [1, 2] }], opts: { gap: 40 } }));
  assert.equal(o.chart.opts, undefined);
});

test("項目を逆順: the categories turn round in the chart and in its data table", async () => {
  const E = await loadEngine();
  const data = { labels: ["A", "B", "C"], series: [{ name: "x", values: [1, 2, 3] }, { name: "y", values: [4, 5, 6] }] };
  const shown = E.visibleChart(E.normalizeObject({ id: "c", kind: "chart", x: 0, y: 0, w: 400, h: 300, chart: { type: "clustered-bar", ...data, opts: { reverse: true } } }).chart);
  assert.deepEqual(plain(shown.labels), ["C", "B", "A"]);
  assert.deepEqual(plain(shown.series.map((s) => s.values)), [[3, 2, 1], [6, 5, 4]]);
  const both = E.visibleChart(E.normalizeObject({ id: "c", kind: "chart", x: 0, y: 0, w: 400, h: 300, chart: { type: "clustered-bar", ...data, opts: { reverse: true, hideLabels: [1] } } }).chart);
  assert.deepEqual(plain(both.labels), ["C", "A"], "the filter and the order work together");
  assert.deepEqual(plain(both.series[0].values), [3, 1]);
  let { o, el } = draw(E, { type: "bar", labels: ["A", "B", "C"], series: [{ name: "x", values: [1, 2, 3] }], opts: { reverse: true, table: true } });
  assert.equal(o.chart.opts.reverse, true);
  assert.deepEqual([...el.querySelectorAll(".hs-chart-table thead th")].map((t) => t.textContent).filter(Boolean), ["C", "B", "A"]);
  ({ o } = draw(E, { type: "waterfall", labels: ["前", "増", "合計"], series: [{ name: "x", values: [1, 2, 3] }], opts: { reverse: true } }));
  assert.equal(o.chart.opts, undefined, "a waterfall keeps its running order");
  const plainChart = E.normalizeObject({ id: "c", kind: "chart", x: 0, y: 0, w: 400, h: 300, chart: { type: "bar", ...data } }).chart;
  assert.equal(E.visibleChart(plainChart), plainChart, "nothing to change: the same chart back");
});

test("折れ線の滑らかさ・マーカー、円の角度・ドーナツの穴・切り出し、目盛間隔", async () => {
  const E = await loadEngine();
  const data = { labels: ["1", "2", "3", "4"], series: [{ name: "Y", values: [1, 3, 2, 5] }] };
  let { o, el } = draw(E, { type: "line", ...data, opts: { smooth: true, marker: "diamond" } });
  assert.deepEqual(plain(o.chart.opts), { smooth: true, marker: "diamond" });
  const draws = el.querySelectorAll(".hs-chart .hs-draw");
  assert.equal(draws.length, 1);
  assert.ok(/ C/.test(draws[0].getAttribute("d")), "the line is a curve");
  assert.equal(el.querySelectorAll(".hs-chart polygon.hs-marker").length, 4, "a diamond on each point");
  ({ el } = draw(E, { type: "line", ...data }));
  assert.equal(el.querySelectorAll(".hs-chart polyline.hs-draw").length, 1, "straight segments by default");
  assert.equal(el.querySelectorAll(".hs-chart .hs-marker").length, 0);
  ({ el } = draw(E, { type: "line", ...data, opts: { marker: "none" } }));
  assert.equal(el.querySelectorAll(".hs-chart circle[fill=transparent]").length, 4, "no mark, but the tip stays where the point is");
  ({ o } = draw(E, { type: "line", ...data, opts: { marker: "circle" } }));
  assert.equal(o.chart.opts, undefined, "the round mark is the default");
  ({ o } = draw(E, { type: "bar", ...data, opts: { smooth: true, marker: "square" } }));
  assert.equal(o.chart.opts, undefined, "only lines take them");

  const pie = { labels: ["a", "b", "c"], series: [{ name: "x", values: [50, 30, 20] }] };
  let start = (e) => [...e.querySelectorAll(".hs-chart .hs-arc")].map((a) => a.getAttribute("transform"));
  ({ o, el } = draw(E, { type: "donut", ...pie }));
  assert.ok(start(el)[0].startsWith("rotate(-90 "), "from the top");
  const wideDefault = Number(el.querySelector(".hs-chart .hs-arc").getAttribute("stroke-width"));
  ({ o, el } = draw(E, { type: "donut", ...pie, opts: { angle: 90, hole: 30, explode: 10 } }));
  assert.deepEqual(plain(o.chart.opts), { angle: 90, explode: 10, hole: 30 });
  assert.ok(start(el)[0].startsWith("translate(") && start(el)[0].includes("rotate(0 "), start(el)[0]);
  assert.ok(Number(el.querySelector(".hs-chart .hs-arc").getAttribute("stroke-width")) > wideDefault, "a smaller hole: a thicker ring");
  ({ o, el } = draw(E, { type: "pie", ...pie, opts: { hole: 30, explode: 20, angle: 400 } }));
  assert.deepEqual(plain(o.chart.opts), { explode: 20 }, "a pie has no hole; an angle beyond 359 is dropped");
  ({ el } = draw(E, { type: "donut", ...pie }));
  assert.equal(el.querySelectorAll(".hs-chart text.hs-val").length, 4, "the centre figure with the usual hole");
  ({ el } = draw(E, { type: "donut", ...pie, opts: { hole: 10 } }));
  assert.equal(el.querySelectorAll(".hs-chart text.hs-val").length, 3, "the centre figure goes when the hole is too small for it");

  assert.deepEqual(plain(E.axisTicks({ axisStep: 25 }, 0, 100)), [0, 25, 50, 75, 100]);
  assert.deepEqual(plain(E.axisTicks({}, 0, 100)), [0, 25, 50, 75, 100]);
  assert.deepEqual(plain(E.axisTicks({ axisStep: 0.1 }, 0, 0.3)), [0, 0.1, 0.2, 0.3], "no floating-point crumbs");
  assert.equal(E.axisTicks({ axisStep: 1 }, 0, 100).length, 5, "a step that would draw more than 24 lines is ignored");
  const ticks = (e) => [...e.querySelectorAll(".hs-chart .hs-tick")].map((t) => t.textContent);
  ({ o, el } = draw(E, { type: "line", ...data, opts: { axisMin: 0, axisMax: 10, axisStep: 2 } }));
  assert.equal(o.chart.opts.axisStep, 2);
  assert.deepEqual(ticks(el), ["0", "2", "4", "6", "8", "10"]);
  ({ el } = draw(E, { type: "area", ...data, opts: { axisMin: 0, axisMax: 6, axisStep: 3 } }));
  assert.deepEqual(ticks(el), ["0", "3", "6"]);
  ({ el } = draw(E, { type: "scatter", ...data, opts: { axisMin: 0, axisMax: 8, axisStep: 4 } }));
  assert.ok(ticks(el).slice(0, 3).join() === "0,4,8", ticks(el).join());
  ({ o } = draw(E, { type: "bar", ...data, opts: { axisStep: 5 } }));
  assert.equal(o.chart.opts, undefined, "bars show their values, not a scale");
  ({ o } = draw(E, { type: "line", ...data, opts: { axisStep: -1 } }));
  assert.equal(o.chart.opts, undefined);
});

test("データ ラベル: every point labelled, where a line's labels sit, and a pie's labels outside its slices", async () => {
  const E = await loadEngine();
  const data = { labels: ["1", "2", "3", "4"], series: [{ name: "Y", values: [10, 30, 20, 50] }] };
  const vals = (el) => [...el.querySelectorAll(".hs-chart text.hs-val")].map((t) => t.textContent);
  let { o, el } = draw(E, { type: "line", ...data });
  assert.equal(vals(el).length, 1, "the last point, which is the highest too, as before");
  ({ el } = draw(E, { type: "line", labels: ["1", "2", "3"], series: [{ name: "Y", values: [10, 50, 20] }] }));
  assert.equal(vals(el).length, 2, "the last point and the highest, as before");
  ({ o, el } = draw(E, { type: "line", ...data, opts: { labelAll: true } }));
  assert.deepEqual(plain(o.chart.opts), { labelAll: true });
  assert.deepEqual(vals(el), ["10", "30", "20", "50"], "a value on every point");
  const yOf = (e) => [...e.querySelectorAll(".hs-chart text.hs-val")].map((t) => Number(t.getAttribute("y")));
  const above = yOf(el);
  ({ el } = draw(E, { type: "line", ...data, opts: { labelAll: true, labelPos: "below" } }));
  assert.ok(yOf(el).every((y, i) => y > above[i]), "below the points");
  ({ o, el } = draw(E, { type: "line", ...data, opts: { labelAll: true, labelPos: "right" } }));
  assert.ok([...el.querySelectorAll(".hs-chart text.hs-val")].every((t) => t.getAttribute("text-anchor") === "start"), "to the right of the points");
  assert.equal(o.chart.opts.labelPos, "right");
  const two = { labels: ["1", "2", "3"], series: [{ name: "A", values: [1, 2, 3] }, { name: "B", values: [3, 2, 1] }] };
  ({ el } = draw(E, { type: "multi-line", ...two, opts: { labelAll: true } }));
  assert.equal(vals(el).length, 4, "the last point of a named line shows with its name, so it is not written twice");
  assert.equal(el.querySelectorAll(".hs-chart tspan.hs-val").length, 2, "...in its name's label");
  ({ el } = draw(E, { type: "area", ...data, opts: { labelAll: true } }));
  assert.deepEqual(vals(el), ["10", "30", "20", "50"]);
  const many = { labels: Array.from({ length: 10 }, (_, i) => `${i}`), series: [{ name: "A", values: Array.from({ length: 10 }, (_, i) => i + 1) }, { name: "B", values: Array.from({ length: 10 }, (_, i) => 10 - i) }] };
  ({ el } = draw(E, { type: "clustered-bar", ...many }));
  assert.equal(vals(el).length, 0, "twenty bars: no values unless asked");
  ({ el } = draw(E, { type: "clustered-bar", ...many, opts: { labelAll: true } }));
  assert.equal(vals(el).length, 20);
  ({ el } = draw(E, { type: "line", ...data, opts: { labelAll: true, labels: false } }));
  assert.equal(vals(el).length, 0, "データ ラベル なし wins");
  ({ o } = draw(E, { type: "bar", ...data, opts: { labelAll: true, labelPos: "below" } }));
  assert.equal(o.chart.opts, undefined, "a bar chart already shows every value");
  ({ o } = draw(E, { type: "clustered-bar", ...two, opts: { labelPos: "below" } }));
  assert.equal(o.chart.opts, undefined, "the position is for lines");

  const pie = { labels: ["来店", "アプリ", "宅配"], series: [{ name: "x", values: [50, 30, 20] }] };
  ({ o, el } = draw(E, { type: "donut", ...pie }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-slice-label").length, 0);
  assert.equal(el.querySelectorAll(".hs-chart rect").length, 3, "the list at the right");
  ({ o, el } = draw(E, { type: "donut", ...pie, opts: { sliceLabels: ["percent", "category", "bogus", "category"] } }));
  assert.deepEqual(plain(o.chart.opts), { sliceLabels: ["category", "percent"] }, "in a fixed order, without repeats");
  const labels = [...el.querySelectorAll(".hs-chart .hs-slice-label")];
  assert.equal(labels.length, 3, "a label each");
  assert.deepEqual([...labels[0].querySelectorAll("tspan")].map((t) => t.textContent), ["来店", "50%"], "name and share, a line each");
  assert.equal(el.querySelectorAll(".hs-chart .hs-leader").length, 3, "a leader line each");
  assert.equal(el.querySelectorAll(".hs-chart rect").length, 0, "the list gives way: the slices carry their names");
  assert.ok(labels.some((t) => t.getAttribute("text-anchor") === "start") && labels.some((t) => t.getAttribute("text-anchor") === "end"), "on both sides of the ring");
  ({ el } = draw(E, { type: "pie", ...pie, opts: { sliceLabels: ["value"] } }));
  assert.deepEqual([...el.querySelectorAll(".hs-chart .hs-slice-label")].map((t) => t.textContent), ["50", "30", "20"]);
  assert.equal(el.querySelectorAll(".hs-chart rect").length, 3, "values only: the list stays so the names are still shown");
  ({ el } = draw(E, { type: "pie", ...pie, opts: { sliceLabels: ["category"], labels: false } }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-slice-label, .hs-chart .hs-leader").length, 0, "データ ラベル なし takes them away");
  ({ o } = draw(E, { type: "line", ...data, opts: { sliceLabels: ["value"] } }));
  assert.equal(o.chart.opts, undefined, "only pies");
  ({ el } = draw(E, { type: "donut", ...pie, opts: { sliceLabels: ["category", "value", "percent"], explode: 15, angle: 45 } }));
  assert.equal(el.querySelectorAll(".hs-chart .hs-slice-label").length, 3, "with the slices pulled out and turned");
});
