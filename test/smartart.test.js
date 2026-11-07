// SmartArt (objects.js): items with levels → a layout of shapes and connectors in the theme's colours; 図形に変換; one by one.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { objectsSummary } from "../server/objects.mjs";
import { objectSchema } from "../server/schemas.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "animate.js"), "utf8"), context, { filename: "animate.js" });
  return { E: window.SlideEngine };
}
/** WCAG contrast of two #rrggbb colours. */
function contrast(a, b) {
  const lum = (hex) => {
    const [r, g, bl] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * bl;
  };
  const [x, y] = [lum(a), lum(b)].sort((m, n) => n - m);
  return (x + 0.05) / (y + 0.05);
}

test("items are kept clean: levels never jump, the first is at the top, layouts and colours are known", async () => {
  const { E } = await loadEngine();
  const o = E.normalizeObject({ id: "s", kind: "smartart", x: 0, y: 0, w: 1200, h: 600, smartart: { layout: "nope", color: "red", items: [{ text: "a", level: 2 }, { text: "b\nc", level: 4 }, { text: "<b>d</b>", level: 1 }] } });
  assert.equal(o.smartart.layout, "blocks");
  assert.equal(o.smartart.color, undefined);
  assert.deepEqual(Array.from(o.smartart.items, (it) => `${it.text}:${it.level}`), ["a:0", "b c:1", "<b>d</b>:1"]);
  const empty = E.normalizeObject({ id: "e", kind: "smartart", x: 0, y: 0, w: 100, h: 100, smartart: { items: [] } });
  assert.equal(empty.smartart.items.length, 1, "never without an item");
  assert.ok(objectSchema.safeParse({ id: "s", kind: "smartart" }).success);
  assert.match(objectsSummary([o]), /SmartArt1／文字「a b c d」/);
});

test("every layout lays its items inside the box, in the theme's colours, with words that read on their fill", async () => {
  const { E } = await loadEngine();
  for (const layout of Object.keys(E.SMARTART_LAYOUTS)) {
    for (const color of Object.keys(E.SMARTART_COLORS)) {
      const o = E.normalizeObject({ id: "s", kind: "smartart", x: 0, y: 0, w: 1500, h: 700, smartart: { layout, color, items: E.smartartSample(layout) } });
      const parts = E.smartartParts(o);
      assert.ok(parts.length >= 3, `${layout}: parts`);
      for (const p of parts) {
        if (p.kind === "line") { assert.ok([p.x1, p.y1, p.x2, p.y2].every(Number.isFinite), `${layout} line`); continue; }
        assert.ok([p.x, p.y, p.w, p.h, p.fs].every(Number.isFinite), `${layout} ${color}: numbers`);
        assert.ok(p.x >= -2 && p.y >= -2 && p.x + p.w <= 1502 && p.y + p.h <= 702, `${layout}: inside the box (${p.x},${p.y},${p.w},${p.h})`);
        assert.ok(p.fill === "none" || E.themeColorRef(p.fill, E.DEFAULT_THEME), `${layout} ${color}: fill ${p.fill} is a theme colour`);
        assert.ok(E.themeColorRef(p.color, E.DEFAULT_THEME), `${layout} ${color}: words ${p.color} in a theme colour`);
        if (p.text && p.fill !== "none") assert.ok(contrast(p.color, p.fill) >= 3, `${layout} ${color}: words ${p.color} read on ${p.fill}`);
        assert.ok(p.fs >= 12 && p.fs <= 60, `${layout}: size ${p.fs}`);
      }
      // Every item's words appear somewhere.
      const words = parts.map((p) => p.text || "").join("");
      for (const it of o.smartart.items) assert.ok(words.includes(it.text), `${layout}: ${it.text}`);
    }
  }
});

test("the hierarchy follows the levels; right to left mirrors; 図形に変換 gives grouped objects", async () => {
  const { E } = await loadEngine();
  const o = E.normalizeObject({ id: "s", kind: "smartart", x: 100, y: 100, w: 1500, h: 700, smartart: { layout: "hierarchy", items: [{ text: "社長", level: 0 }, { text: "営業", level: 1 }, { text: "管理", level: 1 }] } });
  const parts = E.smartartParts(o);
  const boxes = Object.fromEntries(parts.filter((p) => p.kind === "shape").map((p) => [p.text.replace(/<[^>]+>/g, ""), p]));
  assert.ok(boxes["社長"].y < boxes["営業"].y && boxes["営業"].y === boxes["管理"].y, "children a row below");
  assert.ok(Math.abs(boxes["社長"].x + boxes["社長"].w / 2 - 750) < 1, "the top sits over its children");
  assert.equal(parts.filter((p) => p.kind === "line").length, 2, "an elbow to each child");
  const rtl = E.smartartParts({ ...o, smartart: { layout: "process", rtl: true, items: [{ text: "a", level: 0 }, { text: "b", level: 0 }, { text: "c", level: 0 }] } });
  const first = rtl.find((p) => p.item === 0);
  assert.ok(first.x > 750, "right to left: the first step on the right");
  const objs = E.smartartObjects(o);
  assert.equal(objs.length, parts.length);
  assert.ok(objs.every((x) => x.group && x.group === objs[0].group), "one group");
  assert.ok(objs.filter((x) => x.kind === "shape").every((x) => x.x >= 100 && x.y >= 100), "placed on the slide");
  assert.ok(objs.some((x) => x.kind === "line" && x.route === "elbow"));
});

test("drawn with a box per item; an animation can bring the items in one by one", async () => {
  const { E } = await loadEngine();
  const slide = { type: "blank", elements: [{ id: "s", kind: "smartart", x: 100, y: 100, w: 1500, h: 700, smartart: { layout: "process", items: [{ text: "計画", level: 0 }, { text: "実行", level: 0 }, { text: "評価", level: 0 }] } }, { id: "t", kind: "text", x: 0, y: 0, w: 100, h: 50, text: "<p>x</p>" }] };
  const el = E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  assert.equal(el.querySelectorAll('.hs-obj[data-kind="smartart"] .hs-sa-step:not(.hs-sa-fixed)').length, 3);
  assert.ok(el.querySelector('.hs-sa-node[data-item="1"]').textContent.includes("実行"));
  const tl = E.normalizeTimeline([{ el: "s", cls: "in", fx: "fade", by: "item" }, { el: "t", cls: "in", fx: "fade", by: "item" }, { el: "s", cls: "in", fx: "fade", by: "para" }], slide);
  assert.equal(tl[0].by, "item", "one by one on a SmartArt");
  assert.equal(tl[1].by, undefined, "not on a text box");
  assert.equal(tl[2].by, undefined, "paragraphs are for text");
});
