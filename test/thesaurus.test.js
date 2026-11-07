// 校閲 → 類義語: the built-in thesaurus (public/editor/thesaurus.mjs) finds a word's senses whatever its width, script or
// ending, and its data is tidy.
import assert from "node:assert/strict";
import test from "node:test";

import { SYNONYM_GROUPS, lookupKey, synonymsOf } from "../public/editor/thesaurus.mjs";

test("a word's senses, each with the other words of the sense", () => {
  const senses = synonymsOf("課題");
  assert.ok(senses.length >= 1);
  const words = senses.flatMap((s) => s.words);
  assert.ok(words.includes("問題") && words.includes("論点"), words.join());
  assert.ok(!words.includes("課題"), "never the word itself");
  assert.equal(senses[0].label, "問題・課題", "the sense it heads comes first");
});

test("width, katakana/hiragana and endings do not matter; an unknown word has none", () => {
  const a = synonymsOf("ユーザー").flatMap((s) => s.words);
  assert.ok(a.includes("顧客") && a.includes("消費者"));
  assert.deepEqual(synonymsOf("ゆーざー"), synonymsOf("ユーザー"), "hiragana finds the same");
  assert.deepEqual(synonymsOf("ＫＰＩ"), [], "an unknown word");
  assert.ok(synonymsOf("強化する").flatMap((s) => s.words).includes("拡充"), "an ending taken off");
  assert.ok(synonymsOf("簡単な").flatMap((s) => s.words).includes("容易"));
  assert.deepEqual(synonymsOf(""), []);
  assert.deepEqual(synonymsOf(null), []);
  assert.deepEqual(synonymsOf("する"), [], "an ending alone is not a word");
});

test("a word in two senses gives both", () => {
  const labels = synonymsOf("実績").map((s) => s.label);
  assert.ok(labels.length >= 2, labels.join());
  assert.ok(synonymsOf("まとめ").length >= 2);
});

test("the data is tidy: every group a name and at least three words, none twice in a group", () => {
  assert.ok(SYNONYM_GROUPS.length >= 100);
  for (const [label, ...words] of SYNONYM_GROUPS) {
    assert.ok(label && label.length <= 20, label);
    assert.ok(words.length >= 3, `${label}: ${words.length}`);
    const keys = words.map(lookupKey);
    assert.equal(new Set(keys).size, keys.length, `${label} has a word twice`);
  }
});
