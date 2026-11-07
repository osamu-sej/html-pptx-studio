// アクセシビリティ チェック (a11y.mjs), 読み取り順序 (objects.js) and 印刷 (print.mjs): what is reported and how to read
// and print a deck.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { accessibilityIssues, contrast, visualOrder } from "../public/editor/a11y.mjs";
import { commentChunks, commentLines, handoutBoxes, PAGE, paginate, parseRange, printedSlides } from "../public/editor/print.mjs";
import { pictureFileName, rotatedBox } from "../public/editor/picsave.mjs";
import { deckShape } from "../server/schemas.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "animate.js"), "utf8"), context, { filename: "animate.js" });
  return window.SlideEngine;
}
const rules = (issues, i) => issues.filter((x) => x.slide === i).map((x) => x.rule).sort();

test("contrast follows WCAG", () => {
  assert.equal(Math.round(contrast("#000000", "#ffffff")), 21);
  assert.ok(contrast("#1a1a1a", "#dce4f2") > 12);
  assert.ok(contrast("#808080", "#ffffff") < 4.5);
});

test("visual order: rows from the top, then left to right", () => {
  const o = (id, x, y, w = 200, h = 100) => ({ id, kind: "text", x, y, w, h });
  assert.deepEqual(visualOrder([o("c", 100, 600), o("b", 900, 205), o("a", 100, 200)]).map((x) => x.id), ["a", "b", "c"]);
});

test("the checker finds missing alt text, titles, header rows, vague links, low contrast, reading order and duplicates", async () => {
  const E = await loadEngine();
  const deck = { slides: [
    { type: "blank", title: "", elements: [
      { id: "p1", kind: "image", src: "asset:x", x: 100, y: 100, w: 400, h: 300 },
      { id: "p2", kind: "image", src: "asset:x", x: 600, y: 100, w: 400, h: 300, alt: "店舗の外観" },
      { id: "p3", kind: "image", src: "asset:x", x: 1100, y: 100, w: 400, h: 300, decorative: true },
      { id: "t1", kind: "table", x: 100, y: 500, w: 800, h: 300, cells: [["a", "b"], ["c", "d"]], header: false },
    ] },
    { type: "blank", title: "売上", elements: [
      { id: "z", kind: "text", x: 100, y: 800, w: 600, h: 100, text: '<p>詳しくは<a href="https://example.com">こちら</a></p>' },
      { id: "y", kind: "text", x: 100, y: 400, w: 600, h: 100, text: "<p>薄い文字</p>", color: "#808080", fs: 24 },
      { id: "x", kind: "text", x: 100, y: 120, w: 600, h: 100, text: "<p>見出し</p>" },
    ] },
    { type: "content", title: "売上", points: ["a"] },
  ] };
  const issues = accessibilityIssues(deck, E);
  assert.deepEqual(rules(issues, 0), ["alt", "header", "title"]);
  assert.equal(issues.find((x) => x.rule === "alt").el, "p1", "only the picture without a description (not the decorative one)");
  assert.deepEqual(rules(issues, 1), ["contrast", "duplicate", "link", "order"]);
  assert.deepEqual(rules(issues, 2), ["duplicate"]);
  assert.equal(issues[0].level, "error", "errors come first");
  // Setting the reading order to what the eye sees clears that warning.
  deck.slides[1].readingOrder = ["x", "y", "z"];
  assert.ok(!accessibilityIssues(deck, E).some((x) => x.rule === "order"));
});

test("objects keep alternative text or decorative, are drawn for screen readers, and follow the reading order", async () => {
  const E = await loadEngine();
  assert.equal(E.normalizeObject({ id: "a", kind: "shape", x: 0, y: 0, w: 10, h: 10, alt: "  矢印  " }).alt, "矢印");
  const deco = E.normalizeObject({ id: "b", kind: "shape", x: 0, y: 0, w: 10, h: 10, alt: "x", decorative: true });
  assert.equal(deco.decorative, true);
  assert.equal(deco.alt, undefined, "decorative wins");
  const slide = { type: "blank", readingOrder: ["s2", "s1"], elements: [
    { id: "s1", kind: "shape", x: 0, y: 0, w: 100, h: 100, alt: "四角" },
    { id: "s2", kind: "icon", icon: Object.keys(E.icons)[0], x: 200, y: 0, w: 100, h: 100, alt: "店舗" },
    { id: "s3", kind: "line", x1: 0, y1: 0, x2: 10, y2: 10, decorative: true },
  ] };
  assert.deepEqual(Array.from(E.readingOrderOf(slide)), ["s2", "s1", "s3"]);
  const el = E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  const nodes = [...el.querySelectorAll(".hs-objects > .hs-obj")];
  assert.deepEqual(nodes.map((n) => n.getAttribute("data-el")), ["s2", "s1", "s3"], "the DOM follows the reading order");
  assert.equal(nodes[1].style.zIndex, "1", "the stacking order stays");
  assert.equal(nodes[0].style.zIndex, "2");
  assert.equal(nodes[0].getAttribute("role"), "img");
  assert.equal(nodes[0].getAttribute("aria-label"), "店舗");
  assert.equal(nodes[2].getAttribute("aria-hidden"), "true");
  // Without a reading order nothing changes.
  const plain = E.render({ ...slide, readingOrder: undefined }, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  assert.deepEqual([...plain.querySelectorAll(".hs-objects > .hs-obj")].map((n) => n.getAttribute("data-el")), ["s1", "s2", "s3"]);
  assert.equal(plain.querySelector(".hs-obj").style.zIndex, "");
  assert.ok(deckShape.parse({ title: "t", slides: [{ type: "blank", title: "", readingOrder: ["s2", "s1"] }] }).slides[0].readingOrder);
});

test("print: ranges, hidden slides, pages and handout boxes inside the page", () => {
  assert.deepEqual(parseRange("1,3,5-7", 8), [0, 2, 4, 5, 6]);
  assert.deepEqual(parseRange("２", 3), null, "full-width digits are not numbers here");
  assert.deepEqual(parseRange("2、4〜5", 6), [1, 3, 4]);
  assert.equal(parseRange("0-2", 5), null);
  assert.equal(parseRange("3-9", 5), null);
  const deck = { slides: [{}, { hidden: true }, {}, {}] };
  assert.deepEqual(printedSlides(deck, {}), [0, 2, 3]);
  assert.deepEqual(printedSlides(deck, { hidden: true }), [0, 1, 2, 3]);
  assert.deepEqual(printedSlides(deck, { range: "current", current: 1 }), [1], "the current slide prints even when hidden");
  assert.deepEqual(printedSlides(deck, { range: "custom", custom: "2-4" }), [2, 3]);
  assert.deepEqual(paginate([0, 1, 2, 3, 4, 5, 6], "h6").map((p) => p.length), [6, 1]);
  assert.deepEqual(paginate([0, 1, 2], "notes").map((p) => p.length), [1, 1, 1]);
  for (const layout of ["h1", "h2", "h3", "h4", "h6", "h9"]) {
    const n = { h1: 1, h2: 2, h3: 3, h4: 4, h6: 6, h9: 9 }[layout];
    const boxes = handoutBoxes(layout, n);
    assert.equal(boxes.length, n);
    for (const b of boxes) {
      assert.ok(b.x >= PAGE.margin - 1 && b.x + b.w <= PAGE.w - PAGE.margin + 1 && b.y >= PAGE.margin && b.y + b.h <= PAGE.h - PAGE.margin, `${layout} inside the page: ${JSON.stringify(b)}`);
      assert.ok(Math.abs(b.w / b.h - 16 / 9) < 0.02, "16:9");
    }
    // No two slides overlap.
    for (const a of boxes) for (const b of boxes) if (a !== b) assert.ok(a.x + a.w <= b.x || b.x + b.w <= a.x || a.y + a.h <= b.y || b.y + b.h <= a.y, `${layout} overlap`);
  }
  // In columns, the second slide is under the first.
  const cols = handoutBoxes("h6", 6, { order: "columns" });
  assert.equal(cols[1].x, cols[0].x);
  assert.ok(cols[1].y > cols[0].y);
  assert.ok(handoutBoxes("h3", 3)[0].lines, "the 3-slide handout has lines for notes");
});

test("print: コメントを印刷する — each comment, then its replies; empty ones left out; pages of a sensible length", () => {
  const slide = { comments: [
    { id: "c1", text: "数字を確認\n（前年比）", by: "佐藤", at: "2026-10-01T01:00:00Z", done: true, replies: [{ id: "r1", text: "確認しました", by: "鈴木", at: "2026-10-01T02:00:00Z" }, { id: "r2", text: "  " }] },
    { id: "c2", text: "   " },
    { id: "c3", text: "図を差し替え", at: "not a date" },
  ] };
  const lines = commentLines(slide);
  assert.deepEqual(JSON.parse(JSON.stringify(lines.map(({ level, by, text, done }) => ({ level, by, text, done })))), [
    { level: 0, by: "佐藤", text: "数字を確認\n（前年比）", done: true },
    { level: 1, by: "鈴木", text: "確認しました", done: false },
    { level: 0, by: "（名前なし）", text: "図を差し替え", done: false },
  ]);
  assert.ok(lines[0].at.includes("2026"), "when it was written");
  assert.equal(lines[2].at, "", "an unreadable time is left out");
  assert.deepEqual(commentLines({}), []);
  assert.deepEqual(commentLines({ comments: "x" }), []);
  const many = Array.from({ length: 30 }, (_, i) => ({ level: 0, by: "A", at: "", text: "あ".repeat(60), done: false, i }));
  const chunks = commentChunks(many, { perRow: 52, rows: 44 });
  assert.equal(chunks.flat().length, 30, "nothing lost");
  assert.ok(chunks.length > 1 && chunks.every((c) => c.length && 4 * c.length <= 44), "split into pages of at most 44 rows (each takes 4)");
  assert.equal(commentChunks([{ text: "あ".repeat(5000) }]).length, 1, "a long comment still gets a page");
  assert.deepEqual(commentChunks([]), []);
});

test("図として保存: the box a turned object needs, and a file name without the characters files cannot have", () => {
  assert.deepEqual(rotatedBox(200, 100, 0), { w: 200, h: 100 });
  assert.deepEqual(rotatedBox(200, 100, 90), { w: 100, h: 200 });
  assert.deepEqual(rotatedBox(200, 100, -270), { w: 100, h: 200 });
  const b = rotatedBox(100, 100, 45);
  assert.ok(Math.abs(b.w - 141.42) < 0.01 && Math.abs(b.h - 141.42) < 0.01);
  assert.equal(pictureFileName("図 3", "png"), "図 3.png");
  assert.equal(pictureFileName('売上/2026: "上期"?', "svg"), "売上_2026_ _上期__.svg");
  assert.equal(pictureFileName("", "jpg"), "図.jpg");
  assert.equal(pictureFileName("  \u0001 ", "png"), "_.png");
  assert.equal(pictureFileName("あ".repeat(200), "png").length, 84);
});
