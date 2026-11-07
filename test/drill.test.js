import assert from "node:assert/strict";
import test from "node:test";

import { applyChatOperations, drillParents } from "../server/chat.mjs";
import { buildChatPrompt, deckNarrativeLines } from "../server/codex-app-server.mjs";
import { unsourcedNumbers } from "../server/facts.mjs";
import { deckShape } from "../server/schemas.mjs";
import { varietyIssues } from "../public/layout-looks.mjs";

const slides = () => [
  { type: "title", title: "表紙" },
  { type: "stepUp", title: "発展段階", takeaway: "役割を広げてきた", items: [{ title: "規則", desc: "条件で処理" }, { title: "学習", desc: "例から学ぶ" }] },
  { type: "content", title: "学習の仕組み", takeaway: "例から傾向をつかむ", points: ["背景"], drillOf: "items[1]" },
  { type: "cards", title: "選び方", takeaway: "3つの軸", items: [{ title: "用途" }] },
  { type: "statement", title: "これから", text: "まず試す" },
  { type: "closing", message: "以上" },
];
const titles = (list) => list.map((slide) => (slide.drillOf ? `↳${slide.title}` : slide.title ?? slide.message));

test("the deck schema keeps a deep-dive page's link", () => {
  const deck = deckShape.parse({ title: "資料", slides: slides() });
  assert.equal(deck.slides[2].drillOf, "items[1]");
  assert.deepEqual(drillParents(deck.slides), { 2: 1 });
});

test("chat edits keep deep-dive pages with their slide", () => {
  // A story slide inserted "after 2" lands after 2's deep-dive page, not between them.
  const inserted = applyChatOperations(slides(), { operations: [{ op: "insert", slide: 2, content: { type: "kpi", title: "数字", takeaway: "伸びた", items: [{ label: "A", value: "1" }] } }] });
  assert.deepEqual(titles(inserted.slides), ["表紙", "発展段階", "↳学習の仕組み", "数字", "選び方", "これから", "以上"]);
  // A new deep-dive page joins the slide's own.
  const drilled = applyChatOperations(slides(), { operations: [{ op: "insert", slide: 2, content: { type: "content", title: "規則の例", takeaway: "条件で動く", points: ["例"], drillOf: "items[0]" } }] });
  assert.deepEqual(titles(drilled.slides), ["表紙", "発展段階", "↳学習の仕組み", "↳規則の例", "選び方", "これから", "以上"]);
  // Deleting the slide deletes its deep-dive page.
  const deleted = applyChatOperations(slides(), { operations: [{ op: "delete", slide: 2 }] });
  assert.deepEqual(titles(deleted.slides), ["表紙", "選び方", "これから", "以上"]);
  assert.deepEqual(deleted.deleted, [1, 2]);
  // Reordering moves the slide with its page, whether or not the AI listed the page.
  for (const order of [[1, 4, 5, 2, 3, 6], [1, 4, 5, 2, 6]]) {
    const moved = applyChatOperations(slides(), { operations: [], order });
    assert.deepEqual(titles(moved.slides), ["表紙", "選び方", "これから", "発展段階", "↳学習の仕組み", "以上"]);
    assert.equal(moved.moved, true);
  }
  // Rewriting a deep-dive page keeps it one, even when the AI leaves the link out.
  const rewritten = applyChatOperations(slides(), { operations: [{ op: "replace", slide: 3, content: { type: "cards", title: "学習の3要素", takeaway: "データ・特徴・予測", items: [{ title: "データ" }] } }] });
  assert.equal(rewritten.slides[2].drillOf, "items[1]");
});

test("the AI sees which slides are deep-dive pages, and link names are not figures or looks", () => {
  const deck = { title: "資料", slides: slides() };
  assert.match(deckNarrativeLines(deck)[2], /^3枚目 \[深掘り（2枚目の items\[1\] から開く。本編外）\/content\]/);
  const prompt = buildChatPrompt({ deck, message: "2枚目の「学習」を深掘りするページを作って" });
  assert.match(prompt, /drillOf（深掘りページ）/);
  assert.match(prompt, /content に drillOf/);
  assert.deepEqual(unsourcedNumbers(slides(), "表紙 発展段階"), [], "items[1] is a link, not a figure");
  // The page between two card-like slides does not hide that they follow each other.
  const between = [slides()[0], { type: "cards", title: "A", items: [{ title: "a" }] }, { type: "content", title: "深掘り", points: ["x"], drillOf: "items[0]" }, { type: "grid2x2", title: "B", items: [{ title: "b" }] },
    { type: "kpi", title: "C", items: [{ label: "c", value: "1" }] }, { type: "statement", title: "D", text: "d" }, { type: "logicTree", title: "E", root: "e", branches: [{ title: "f" }, { title: "g" }] }, slides()[5]];
  assert.deepEqual(varietyIssues(between).map((issue) => [issue.rule, issue.slides]), [["adjacent", [1, 3]]]);
});
