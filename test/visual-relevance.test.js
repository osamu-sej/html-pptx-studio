import assert from "node:assert/strict";
import test from "node:test";

import { reconcileChatVisuals, slideMeaning } from "../server/visual-relevance.mjs";

const deck = {
  title: "生成AI研修の報告",
  slides: [
    { type: "title", title: "表紙" },
    { type: "content", title: "次回訪店から実行", takeaway: "推進役が店長と改善計画を作る", points: ["毎週火曜までに担当店の改善ストーリーを完成させる"] },
    { type: "closing", message: "以上" },
  ],
};

const replace = (content) => ({ reply: "画像を変更しました", operations: [{ op: "replace", slide: 2, content }], suggestions: [] });

test("slide meaning uses the actual body, not notes or hidden image data", () => {
  const text = slideMeaning({ ...deck.slides[1], notes: "研修の写真", customImage: "data:image/png;base64,AAAA" });
  assert.match(text, /改善ストーリー/);
  assert.doesNotMatch(text, /研修の写真|base64/);
});

test("unrequested stock photos are removed from AI revisions", () => {
  const { answer, warnings } = reconcileChatVisuals(deck, replace({ ...deck.slides[1], visualAsset: "businessWorkshop" }), "文章を簡潔に");
  assert.equal(answer.operations.length, 0);
  assert.match(warnings[0], /依頼対象ではない/);
});

test("a generic photo request cannot use a deck-wide topic to justify an unrelated image", () => {
  const { answer, warnings } = reconcileChatVisuals(deck, replace({ ...deck.slides[1], visualAsset: "businessWorkshop" }), "この1枚に写真を選んで");
  assert.equal(answer.operations.length, 0);
  assert.match(warnings[0], /内容との関連/);
});

test("a photo that depicts the slide's subject is allowed", () => {
  const { answer, warnings } = reconcileChatVisuals(deck, replace({ ...deck.slides[1], visualAsset: "storeOperations" }), "この1枚に写真を選んで");
  assert.equal(answer.operations[0].content.visualAsset, "storeOperations");
  assert.deepEqual(warnings, []);
});

test("existing photos survive text-only edits", () => {
  const current = { ...deck, slides: [deck.slides[0], { ...deck.slides[1], visualAsset: "storeOperations" }, deck.slides[2]] };
  const next = { ...deck.slides[1], takeaway: "改善計画を来週から実行" };
  delete next.visualAsset;
  const { answer } = reconcileChatVisuals(current, replace(next), "結論を強く");
  assert.equal(answer.operations[0].content.visualAsset, "storeOperations");
});
