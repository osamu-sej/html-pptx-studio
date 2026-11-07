// 段落のタブ設定: tab stops on a paragraph (left / centre / right, px from the text area's left), the words between tab characters
// laid out in boxes at those stops (nothing measured), inches by default as in PowerPoint, and kept through the sanitizer.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import * as ops from "../public/editor/ops.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  return { E: window.SlideEngine, window };
}
const plain = (v) => JSON.parse(JSON.stringify(v));

test("parseTabs: sorted, one to a place, valid kinds and places only, at most eight", async () => {
  const { E } = await loadEngine();
  assert.deepEqual(plain(E.parseTabs("r900, l300,c600")), [{ type: "l", x: 300 }, { type: "c", x: 600 }, { type: "r", x: 900 }]);
  assert.equal(E.tabsText(E.parseTabs("r900,l300")), "l300,r900");
  assert.deepEqual(plain(E.parseTabs("l300,r300")), [{ type: "l", x: 300 }], "one stop to a place");
  assert.deepEqual(plain(E.parseTabs("x300,l0,l99999,l-5,,l12.34,l12.345")), [{ type: "l", x: 12.34 }], "junk is dropped (and places of more than two decimals)");
  assert.equal(E.parseTabs(Array.from({ length: 12 }, (_, i) => `l${(i + 1) * 100}`).join(",")).length, 8);
  assert.deepEqual(plain(E.parseTabs(null)), []);
  assert.deepEqual(plain(E.parseTabs(undefined)), []);
});

test("tabBoxes: a left stop starts the next word, a right stop ends it, a centre stop centres it; inches past the last", async () => {
  const { E } = await loadEngine();
  const boxes = (count, stops, start) => plain(E.tabBoxes(count, E.parseTabs(stops), start));
  // name  |  item  | price (left at 300, right at 900)
  assert.deepEqual(boxes(3, "l300,r900", 0), [{ w: 300, align: "left" }, { w: 300, align: "left" }, { w: 300, align: "right" }], "the price ends at 900");
  // one left stop: the last word runs on from it
  assert.deepEqual(boxes(2, "l300", 0), [{ w: 300, align: "left" }, { w: null, align: "left" }]);
  // one right stop: the word before it gets the first half
  assert.deepEqual(boxes(2, "r500", 0), [{ w: 250, align: "left" }, { w: 250, align: "right" }]);
  // one centre stop: centred on it
  assert.deepEqual(boxes(2, "c600", 0), [{ w: 300, align: "left" }, { w: 600, align: "center" }], "300 → 900, centred on 600");
  // no stops: every inch (144 px)
  assert.deepEqual(boxes(3, "", 0), [{ w: 144, align: "left" }, { w: 144, align: "left" }, { w: null, align: "left" }]);
  // more tabs than stops: inches go on after the last stop
  assert.deepEqual(boxes(3, "l200", 0).map((b) => b.w), [200, 88, null], "the next inch after 200 is 288");
  // a first line that starts further in (or hangs out)
  assert.deepEqual(boxes(2, "l300", 40).map((b) => b.w), [260, null]);
  assert.deepEqual(boxes(2, "l300", -40).map((b) => b.w), [340, null], "a hanging first line starts left of the margin");
  assert.deepEqual(boxes(1, "l300", 0), [{ w: null, align: "left" }], "no tab, no box");
});

test("data-tabs is kept on a paragraph (valid stops only) and nowhere else", async () => {
  const { E } = await loadEngine();
  assert.equal(E.sanitizeRich('<p data-tabs="r900,l300">a\tb</p>'), '<p data-tabs="l300,r900">a\tb</p>');
  assert.equal(E.sanitizeRich('<p data-tabs="nonsense">a</p>'), "<p>a</p>");
  assert.equal(E.sanitizeRich('<ul><li data-tabs="l300">a</li></ul>').includes("data-tabs"), false);
  assert.ok(E.sanitizeRich("<p>a\tb</p>").includes("\t"), "a tab character stays");
});

test("tabs are drawn: the words between tab characters sit in boxes at the stops; plain text is left alone", async () => {
  const { E } = await loadEngine();
  const draw = (text, extra = {}) => {
    const o = E.normalizeObject({ id: "t", kind: "text", x: 0, y: 0, w: 1200, h: 300, text, fs: 36, ...extra });
    const slide = { type: "blank", elements: [o] };
    return E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  };
  const el = draw('<p data-tabs="l300,r900">品名\t数量\t<b>金額</b></p><p>ふつう</p>');
  const boxes = [...el.querySelectorAll(".hs-obj-tx p")[0].children];
  assert.equal(boxes.length, 3);
  assert.deepEqual(boxes.map((b) => b.getAttribute("style")), ["width: 300px; text-align: left", "width: 300px; text-align: left", "width: 300px; text-align: right"]);
  assert.equal(boxes[2].textContent, "金額");
  assert.ok(boxes[2].querySelector("b"), "the bold stays inside its word");
  assert.equal(el.querySelectorAll(".hs-obj-tx p")[1].textContent, "ふつう");
  assert.equal(el.querySelectorAll(".hs-obj-tx p")[1].querySelector(".hs-tabbox"), null, "a paragraph without tabs is not touched");
  // A tab inside bold text is split with the tag repeated on both sides.
  const split = draw("<p><b>前\t後</b></p>");
  const parts = [...split.querySelectorAll(".hs-obj-tx p .hs-tabbox")];
  assert.deepEqual(parts.map((p) => p.textContent), ["前", "後"]);
  assert.ok(parts.every((p) => p.querySelector("b")), "bold on both sides of the tab");
  // Default stops (no data-tabs): an inch each; the last word is not boxed in.
  const dflt = draw("<p>a\tb\tc</p>");
  const widths = [...dflt.querySelectorAll(".hs-tabbox")].map((b) => b.getAttribute("style"));
  assert.deepEqual(widths, ["width: 144px; text-align: left", "width: 144px; text-align: left", null]);
  assert.ok(dflt.querySelectorAll(".hs-tabbox")[2].classList.contains("is-end"));
  // A hanging indent with a tab (the "1.\ttext" way): the first word's box reaches the stop from where the line starts.
  const hang = draw('<p data-tabs="l120" style="margin-left: 120px; text-indent: -120px">1.\t本文</p>');
  assert.equal(hang.querySelector(".hs-tabbox").getAttribute("style"), "width: 120px; text-align: left");
  // Vertical text and text on its side keep their tab characters as they are.
  assert.equal(draw("<p>a\tb</p>", { vertical: true }).querySelector(".hs-tabbox"), null);
  assert.equal(draw("<p>a\tb</p>", { textRot: 90 }).querySelector(".hs-tabbox"), null);
  // Drawn twice, the same (the source html is not changed).
  const again = draw('<p data-tabs="l300">a\tb</p>');
  assert.equal(again.querySelectorAll(".hs-tabbox").length, 2);
});

test("ops.setTabs / tabsOf: stops on every paragraph, read back from the first; none takes them away", async () => {
  const { E, window } = await loadEngine();
  globalThis.document = window.document;
  const html = "<p>一行目\t値</p><p>二行目\t値</p>";
  const set = ops.setTabs(E, html, [{ type: "r", x: 600 }, { type: "l", x: 200 }]);
  assert.equal((set.match(/data-tabs="l200,r600"/g) || []).length, 2, "on both paragraphs");
  assert.deepEqual(plain(ops.tabsOf(E, set)), [{ type: "l", x: 200 }, { type: "r", x: 600 }]);
  const cleared = ops.setTabs(E, set, []);
  assert.ok(!cleared.includes("data-tabs"));
  assert.deepEqual(plain(ops.tabsOf(E, cleared)), []);
  assert.ok(cleared.includes("\t"), "the tab characters stay");
  assert.equal(ops.setTabs(E, "ただの文字", [{ type: "l", x: 100 }]), '<p data-tabs="l100">ただの文字</p>', "words without a paragraph get one");
  delete globalThis.document;
});
