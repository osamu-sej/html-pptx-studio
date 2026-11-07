// 表記ゆれチェック (proof.mjs) and ファイル → 情報・ドキュメント検査 (fileinfo.mjs).
import assert from "node:assert/strict";
import test from "node:test";

import { unify, variantGroups } from "../public/editor/proof.mjs";
import { deckStats, infoOf, inspect, strip } from "../public/editor/fileinfo.mjs";
import { deckShape } from "../server/schemas.mjs";

test("variants: full and half width, katakana long vowels and half-width kana, okurigana", () => {
  const groups = variantGroups(["ＡＩを使う。AIの活用", "ユーザとユーザーの声", "ｻｰﾋﾞｽとサービス", "打合せと打ち合わせを行う。行なう日", "<p>2026年と２０２６年</p>", "ユーザビリティ"]);
  const byKey = Object.fromEntries(groups.map((g) => [`${g.kind}:${g.key}`, g.forms.map((f) => f.text).sort()]));
  assert.deepEqual(byKey["width:AI"], ["AI", "ＡＩ"]);
  assert.deepEqual(byKey["width:2026"], ["2026", "２０２６"]);
  assert.deepEqual(byKey["kana:ユーザ"], ["ユーザ", "ユーザー"]);
  assert.deepEqual(byKey["kana:サービス"], ["サービス", "ｻｰﾋﾞｽ"]);
  assert.deepEqual(byKey["okurigana:打ち合わせ"], ["打ち合わせ", "打合せ"]);
  assert.deepEqual(byKey["okurigana:行う"], ["行う", "行なう"]);
  // A word written one way only is not a variant.
  assert.ok(!groups.some((g) => g.key === "ユーザビリティ"));
  assert.deepEqual(variantGroups(["売上は好調", "売上の推移"]), []);
});

test("unify writes whole words only, and the longer form is not cut", () => {
  const kana = variantGroups(["ユーザ", "ユーザー"]).find((g) => g.kind === "kana");
  assert.equal(unify("ユーザとユーザーとユーザビリティ", kana, "ユーザー"), "ユーザーとユーザーとユーザビリティ");
  assert.equal(unify("ユーザとユーザー", kana, "ユーザ"), "ユーザとユーザ");
  const width = variantGroups(["ＡＩ", "AI"]).find((g) => g.kind === "width");
  assert.equal(unify("ＡＩとAIとAIM", width, "AI"), "AIとAIとAIM");
  const ok = variantGroups(["取組", "取組み", "取り組み"]).find((g) => g.kind === "okurigana");
  assert.equal(unify("取組と取組みと取り組み", ok, "取り組み"), "取り組みと取り組みと取り組み");
});

test("file info: properties checked, statistics, document inspection and removal", () => {
  assert.equal(infoOf({ author: "  ", other: "x" }), null);
  assert.deepEqual(infoOf({ author: " 山田 ", keywords: "売上,<b>店舗</b>" }), { author: "山田", keywords: "売上,b店舗/b" });
  const deck = { title: "t", info: { author: "山田" }, slides: [
    { type: "title", title: "表紙", notes: "あいさつ" },
    { type: "blank", title: "", hidden: true, comments: [{ id: "c", text: "確認", done: false }], elements: [
      { id: "a", kind: "text", x: 100, y: 100, w: 200, h: 50, text: "<p>見える</p>" },
      { id: "b", kind: "shape", x: 3000, y: 100, w: 100, h: 100 },
      { id: "c", kind: "shape", x: 100, y: 100, w: 100, h: 100, hidden: true },
    ], timeline: [{ id: "t1", el: "b", cls: "in", fx: "fade" }] },
    { type: "content", title: "まとめ", points: ["一つ"] },
  ] };
  const st = deckStats(deck, { storyMap: () => ({ parent: [] }) });
  assert.equal(st.slides, 3);
  assert.equal(st.hidden, 1);
  assert.equal(st.objects, 3);
  assert.equal(st.comments, 1);
  assert.equal(st.open, 1);
  assert.equal(st.notes, 1);
  assert.ok(st.chars >= 9);
  assert.deepEqual(inspect(deck), { comments: 1, notes: 1, hiddenSlides: 1, hiddenObjects: 1, offSlide: 1, info: 1 });
  strip(deck, "offSlide");
  assert.deepEqual(deck.slides[1].elements.map((o) => o.id), ["a", "c"]);
  assert.equal(deck.slides[1].timeline, undefined, "its animation goes with it");
  strip(deck, "hiddenObjects");
  strip(deck, "comments");
  strip(deck, "notes");
  strip(deck, "info");
  strip(deck, "hiddenSlides");
  assert.deepEqual(inspect(deck), { comments: 0, notes: 0, hiddenSlides: 0, hiddenObjects: 0, offSlide: 0, info: 0 });
  assert.equal(deck.slides.length, 2);
  assert.deepEqual(deckShape.parse({ title: "t", slides: [{ type: "blank", title: "" }], info: { author: "山田" } }).info, { author: "山田" });
});
