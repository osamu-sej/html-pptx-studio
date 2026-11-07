// 校閲 → 比較 (compare.mjs): slides matched by id (else title); added, removed and changed slides; taking them over.
import assert from "node:assert/strict";
import test from "node:test";

import { accept, compareDecks, slideDiff } from "../public/editor/compare.mjs";

const mine = { slides: [
  { sid: "a", type: "title", title: "表紙" },
  { sid: "b", type: "content", title: "現状", points: ["売上は横ばい"], comments: [{ id: "c1", text: "確認" }] },
  { sid: "c", type: "content", title: "課題", points: ["人手不足"] },
  { sid: "d", type: "closing", title: "次のアクション" },
] };
const theirs = { slides: [
  { sid: "a", type: "title", title: "表紙" },
  { sid: "b", type: "content", title: "現状", points: ["売上は前年比112%"], notes: "ここで数字を強調" },
  { sid: "x", type: "content", title: "打ち手", points: ["発注の自動化"] },
  { sid: "d", type: "closing", title: "次のアクション" },
] };

test("compare finds the changed, added and removed slides in deck order", () => {
  const changes = compareDecks(mine, theirs);
  assert.deepEqual(changes.map((c) => c.kind), ["changed", "added", "removed"]);
  assert.deepEqual(changes[0].what, ["本文", "ノート"]);
  assert.equal(changes[1].theirs, 2);
  assert.equal(changes[2].mine, 2);
  assert.deepEqual(compareDecks(mine, mine), [], "no differences with itself");
  // Comments and the reading order are not differences.
  assert.deepEqual(compareDecks(mine, { slides: mine.slides.map((s) => ({ ...s, comments: undefined })) }), []);
});

test("slides without ids are matched by their title", () => {
  const plain = { slides: theirs.slides.map(({ sid, ...s }) => s) };
  const changes = compareDecks(mine, plain);
  assert.equal(changes.find((c) => c.kind === "changed")?.mine, 1);
  assert.equal(changes.filter((c) => c.kind === "added").length, 1);
});

test("taking over: one change, or all of them, in the right places; ids and our comments kept", () => {
  const changes = compareDecks(mine, theirs);
  const one = accept(mine, theirs, [changes[0]]);
  assert.equal(one[1].points[0], "売上は前年比112%");
  assert.equal(one[1].sid, "b");
  assert.equal(one[1].comments[0].text, "確認", "our comments stay");
  assert.equal(one.length, 4);
  const all = accept(mine, theirs, changes);
  assert.deepEqual(all.map((s) => s.title), ["表紙", "現状", "打ち手", "次のアクション"]);
  const added = accept(mine, theirs, [changes[1]]);
  assert.deepEqual(added.map((s) => s.title), ["表紙", "現状", "打ち手", "課題", "次のアクション"], "the new slide after the one it follows in their version");
  assert.deepEqual(slideDiff({ title: "A", type: "content" }, { title: "B", type: "cards", elements: [{}] }), ["タイトル", "レイアウト", "部品（0→1個）"]);
});
