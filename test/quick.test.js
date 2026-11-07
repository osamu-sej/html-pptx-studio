// 検索 (⌥Q) and the クイック アクセス ツール バー (public/editor/quick.mjs): matching across scripts and widths, the
// order of the results, and a stored toolbar tidied (known commands only, no repeats, safe ribbon references).
import assert from "node:assert/strict";
import test from "node:test";

import { QAT_DEFAULT_ITEMS, fold, normalizeQat, qatKey, searchCommands } from "../public/editor/quick.mjs";

const commands = [
  { ref: "picture|調整|透明度", label: "透明度", title: "図の透明度", tabLabel: "図の形式", group: "調整" },
  { ref: "animation|アニメーションの詳細設定|アニメーション ウィンドウ", label: "アニメーション ウィンドウ", title: "アニメーション ウィンドウを開く", tabLabel: "アニメーション", group: "アニメーションの詳細設定" },
  { ref: "view|表示|ガイド", label: "ガイド", title: "スライドに垂直・水平のガイドを置く", tabLabel: "表示", group: "表示" },
  { ref: "shape|図形のスタイル|透明度", label: "透明度", title: "図形の透明度", tabLabel: "図形の書式", group: "図形のスタイル", disabled: true },
  { ref: "home|図形描画|配置", label: "配置", title: "整列・前面へ・背面へ（透明度ではない）", tabLabel: "ホーム", group: "図形描画" },
];

test("fold: full-width and half-width, katakana and hiragana, case, spaces are matched alike", () => {
  assert.equal(fold("ＡＩ　Ａｎｉｍａｔｉｏｎ"), "ai animation");
  assert.equal(fold("アニメーション"), "あにめーしょん");
  assert.equal(fold("  ガイド \n 表示 "), "がいど 表示");
});

test("searchCommands: every word must match; names first, a command that cannot be used now after the rest", () => {
  assert.deepEqual(searchCommands("あにめ", commands).map((c) => c.ref), ["animation|アニメーションの詳細設定|アニメーション ウィンドウ"]);
  // The name holding the words comes before a description holding them; the unusable twin after the usable one.
  assert.deepEqual(searchCommands("透明度", commands).map((c) => c.ref), ["picture|調整|透明度", "shape|図形のスタイル|透明度", "home|図形描画|配置"]);
  // Words in any order, found in the name, the description or the tab and group.
  assert.deepEqual(searchCommands("表示 ガイド", commands).map((c) => c.label), ["ガイド"]);
  assert.equal(searchCommands("ガイド 透明度", commands).length, 0);
  assert.equal(searchCommands("  ", commands).length, 0);
  assert.equal(searchCommands("透明度", commands, 1).length, 1);
});

test("normalizeQat: known commands and safe ribbon references only, no repeats, the defaults when nothing is stored", () => {
  assert.deepEqual(normalizeQat(null), { show: false, below: false, items: QAT_DEFAULT_ITEMS });
  const q = normalizeQat({ show: 1, below: "yes", items: ["save", "save", "bogus", { ref: "home|フォント|太字", label: "太字", icon: "bold" }, { ref: "home|フォント|太字" }, { ref: "no-pipes" }, { ref: "home|x|<script>", icon: "<img>" }, 42] });
  assert.equal(q.show, true);
  assert.equal(q.below, true);
  assert.deepEqual(q.items.map(qatKey), ["save", "home|フォント|太字", "home|x|<script>"]);
  assert.equal(q.items[2].icon, "", "an icon name is a plain word");
  assert.equal(q.items[2].label, "<script>", "a label is only ever text (set with textContent)");
  assert.equal(normalizeQat({ items: Array.from({ length: 50 }, (_, i) => ({ ref: `home|g|c${i}` })) }).items.length, 30);
  assert.deepEqual(normalizeQat({ show: true, items: [] }).items, [], "an emptied toolbar stays empty");
});

test("a toolbar saved with a ribbon button's old tooltip still finds it", () => {
  const qat = normalizeQat({ show: true, items: [{ ref: "view|表示/非表示|1cmごとの線を表示", label: "グリッド線" }, { ref: "view|表示/非表示|0.25cmごとに吸着", label: "吸着" }] });
  assert.deepEqual(qat.items.map((it) => it.ref), ["view|表示/非表示|グリッド線を表示（間隔はグリッドとガイドの設定で）", "view|表示/非表示|グリッド線に吸着（間隔はグリッドとガイドの設定で）"]);
  assert.deepEqual(qat.items.map((it) => it.label), ["グリッド線", "吸着"], "its label stays");
  assert.equal(normalizeQat({ items: [{ ref: "picture|調整|トリミング・修整・枠線を元に戻す" }] }).items[0].ref, "picture|調整|トリミング・修整・枠線・影を元に戻す", "図のリセット");
});
