// The eyedropper's colours (public/editor/eyedropper.mjs) and what 音声読み上げ reads (public/editor/readaloud.mjs):
// a slide's words in blocks and in order, a bold run kept in its sentence, decoration and hidden words
// left out.
import assert from "node:assert/strict";
import test from "node:test";
import { parseHTML } from "linkedom";

import { cssToHex } from "../public/editor/eyedropper.mjs";
import { readingChunks } from "../public/editor/readaloud.mjs";

test("cssToHex: rgb(), rgba(), short and long hex; transparent and none give nothing", () => {
  assert.equal(cssToHex("rgb(31, 56, 100)"), "#1f3864");
  assert.equal(cssToHex("rgba(214, 201, 184, 0.5)"), "#d6c9b8");
  assert.equal(cssToHex("rgb(31 56 100 / 80%)"), "#1f3864");
  assert.equal(cssToHex("#ABC"), "#aabbcc");
  assert.equal(cssToHex("#1A1A1A"), "#1a1a1a");
  assert.equal(cssToHex("#1a1a1a00"), null);
  assert.equal(cssToHex("rgba(0, 0, 0, 0)"), null);
  assert.equal(cssToHex("transparent"), null);
  assert.equal(cssToHex("none"), null);
  assert.equal(cssToHex("url(#g)"), null);
});

test("readingChunks: blocks in order, a bold run kept in its sentence, decoration and hidden words left out", () => {
  const { document } = parseHTML(`<div class="hs-slide">
    <div class="hs-decor"><span>飾りの文字</span></div>
    <h2 class="hs-title">結論：全社展開へ</h2>
    <p class="hs-message">試行で<b>月1,440時間</b>を削減</p>
    <ul><li>第1段階 店舗</li><li>第2段階 本部</li></ul>
    <div class="hs-objects"><div class="hs-obj" data-el="t1"><div class="hs-obj-tx"><p>Enter the <b>next</b> phase</p></div></div></div>
    <p hidden>隠れた文</p>
    <svg aria-hidden="true"><text>12</text></svg>
  </div>`);
  const chunks = readingChunks(document.querySelector(".hs-slide"));
  assert.deepEqual(chunks.map((c) => c.text), ["結論：全社展開へ", "試行で月1,440時間を削減", "第1段階 店舗", "第2段階 本部", "Enter the next phase"]);
  assert.equal(chunks[1].el.tagName, "P", "the sentence's block is lit");
  assert.ok(chunks[4].el.closest('[data-el="t1"]'), "an object's words know their object");
});
