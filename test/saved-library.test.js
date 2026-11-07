import assert from "node:assert/strict";
import test from "node:test";

import { buildSearchIndex, searchSavedDecks, slideExcerpt } from "../public/saved-library.js";

const deck = {
  title: "店舗改革の提案",
  purpose: "意思決定",
  memo: "人件費は試算値",
  slides: [
    { type: "title", title: "店舗改革の提案", subtitle: "2026年" },
    { type: "table", title: "作業時間の削減", headers: ["部門", "削減時間"], rows: [["発注", "120時間"]], details: [{ target: "rows[0]", text: "POS集計に基づく" }] },
    { type: "content", title: "実施計画", points: ["10月に試行"], notes: "担当は店舗運営部", media: { src: "idb:secret-video", kind: "video", name: "confidential.mp4" } },
  ],
};
const record = { id: "one", title: deck.title, updatedAt: 10, deck, searchIndex: buildSearchIndex(deck) };

test("saved deck search includes body, tables, figures, details and speaker notes", () => {
  assert.deepEqual(searchSavedDecks([record], "１２０時間")[0].slideMatches, [1]);
  assert.deepEqual(searchSavedDecks([record], "pos 集計")[0].slideMatches, [1]);
  assert.deepEqual(searchSavedDecks([record], "店舗運営部")[0].slideMatches, [2]);
  assert.deepEqual(searchSavedDecks([record], "店舗改革")[0].slideMatches, [0]);
  assert.deepEqual(searchSavedDecks([record], "改革 試行")[0].slideMatches, [0, 1, 2]);
  assert.deepEqual(searchSavedDecks([record], "意思決定")[0].slideMatches, [0, 1, 2]);
  assert.equal(searchSavedDecks([record], "secret-video").length, 0);
  assert.equal(searchSavedDecks([record], "confidential.mp4").length, 0);
  assert.match(slideExcerpt(deck.slides[1], "pos"), /POS集計/);
});
