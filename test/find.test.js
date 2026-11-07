// 検索・置換 (public/editor/find.mjs): case, width and whole-word options on one text, with ranges in the original text.
import assert from "node:assert/strict";
import test from "node:test";

import { foldText, matcher, snippet } from "../public/editor/find.mjs";

test("matcher: no options means not caring about case or width, like Word", () => {
  const m = matcher({ needle: "dx" });
  assert.equal(m.count("DX化とＤＸとdx"), 3);
  assert.deepEqual(m.ranges("DX化とＤＸ"), [[0, 2], [4, 6]]);
  assert.equal(matcher({ needle: "ＤＸ" }).count("DX と dx"), 2, "full-width needle finds the half-width text");
  assert.equal(matcher({ needle: "" }), null);
  assert.equal(matcher({}), null);
});

test("matcher: 大文字と小文字を区別する and 全角と半角を区別する", () => {
  assert.equal(matcher({ needle: "dx", caseSensitive: true }).count("DX dx Dx"), 1);
  assert.equal(matcher({ needle: "DX", width: false }).count("DX ＤＸ"), 1);
  assert.equal(matcher({ needle: "DX", width: false, caseSensitive: true }).count("DX dx ＤＸ"), 1);
});

test("matcher: half-width kana and their marks are one letter", () => {
  const m = matcher({ needle: "ガイド" });
  assert.deepEqual(m.ranges("ｶﾞｲﾄﾞ"), [[0, 5]], "ｶﾞ ｲ ﾄﾞ are five characters of the text, three letters");
  assert.equal(m.replace("ｶﾞｲﾄﾞと ガイド", "案内"), "案内と 案内");
  assert.equal(matcher({ needle: "ｶﾞｲﾄﾞ" }).count("ガイド"), 1, "and the other way round");
  assert.equal(matcher({ needle: "ガイド", width: false }).count("ｶﾞｲﾄﾞ"), 0);
});

test("matcher: 完全に一致する単語だけ stands for words of ASCII letters and digits", () => {
  const m = matcher({ needle: "DX", wholeWord: true });
  assert.equal(m.count("DX化 ADX DX2 DX。 (DX) dx_x"), 3, "DX化, DX。 and (DX): the ones with no letter or digit beside them");
  assert.equal(matcher({ needle: "売上", wholeWord: true }).count("売上高と売上"), 2, "Japanese has no spaces to stand alone by");
  assert.equal(matcher({ needle: "A-1", wholeWord: true }).count("A-1 A-12 XA-1"), 1);
});

test("replace: the replacement is plain text, every match, and the original is left alone otherwise", () => {
  const m = matcher({ needle: "$&" });
  assert.equal(m.replace("a $& b", "$1"), "a $1 b", "no replacement patterns");
  assert.equal(matcher({ needle: "柱" }).replace("柱は柱", "施策"), "施策は施策");
  assert.equal(matcher({ needle: "x" }).replace("abc", "y"), "abc");
  assert.equal(matcher({ needle: "a" }).replace("aaa", "aa"), "aaaaaa", "matches are not looked for in the new text");
  assert.equal(matcher({ needle: "ＡＢ" }).replace("ab AB ａｂ", "ー"), "ー ー ー");
});

test("foldText maps each folded letter back to its original span; snippet shows the words round a match", () => {
  const f = foldText("ｶﾞA", { width: true });
  assert.equal(f.folded, "がa".replace("が", "ガ").toLowerCase());
  assert.deepEqual([f.from[0], f.to[0], f.from[1], f.to[1]], [0, 2, 2, 3]);
  assert.equal(foldText("㈱ABC", { width: true }).folded, "(株)abc");
  assert.equal(snippet("この資料では柱を三つ立てて説明します。", [6, 7], 4), "…資料では柱を三つ立…");
  assert.equal(snippet("短い", [0, 1], 10), "短い");
  assert.equal(snippet("a\n\n b", null), "a b");
});

test("matcher: characters outside the BMP, a lone voiced mark and an empty replacement do not trip it", () => {
  const m = matcher({ needle: "a" });
  assert.deepEqual(m.ranges("😀A"), [[2, 3]], "an emoji is two units of the text; the match after it starts in the right place");
  assert.equal(m.replace("😀A😀a", "b"), "😀b😀b");
  assert.equal(matcher({ needle: "😀" }).count("😀😀"), 2);
  assert.equal(matcher({ needle: "x" }).replace("axb", ""), "ab", "replacing with nothing deletes the match");
  assert.equal(matcher({ needle: "ﾞ" }).count("ｶﾞ ﾞ"), 1, "a voiced mark with no kana before it stays a letter of its own");
  assert.equal(matcher({ needle: "  " }).count("a  b"), 1, "spaces are searched for as they are");
  assert.equal(matcher({ needle: "ｱ", width: false }).count("アｱ"), 1);
});
