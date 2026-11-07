import assert from "node:assert/strict";
import test from "node:test";

import {
  buildDeckPrompt,
  buildChatPrompt,
  buildRepairPrompt,
  buildRevisePrompt,
  normalizeCodexOutputSchema,
  omitNullObjectValues,
  parseDeckResult,
  withoutImageData,
} from "../server/codex-app-server.mjs";

test("image data is never sent to the AI", () => {
  const deck = { title: "x", slides: [{ type: "title", title: "x", customImage: "data:image/jpeg;base64,AAAA" }, { type: "closing" }] };
  assert.deepEqual(withoutImageData(deck.slides)[0].customImage, "[画像あり]");
  assert.doesNotMatch(buildRevisePrompt({ deck, slideIndex: 0 }), /base64/);
});

test("revision prompt targets one slide and carries the deck context", () => {
  const prompt = buildRevisePrompt({
    deck: { title: "中間報告", audience: "役員", slides: [{ type: "title", title: "表紙" }, { type: "content", title: "成果", takeaway: "削減した" }, { type: "closing", message: "承認を依頼" }] },
    slideIndex: 1,
    instruction: "数字を強調",
    issues: [{ field: "takeaway", message: "文字が枠に収まりません" }],
  });
  assert.match(prompt, /2枚目のスライドだけ/);
  assert.match(prompt, /数字を強調/);
  assert.match(prompt, /1枚目 \[導入\/title\] 表紙/);
  assert.match(prompt, /takeaway: 文字が枠に収まりません/);
  assert.match(prompt, /"title":"成果"/);
});

test("chat and revision prompts use slide claims, body and neighboring context", () => {
  const deck = { title: "研修の報告", audience: "役員", purpose: "実行判断", slides: [
    { type: "title", title: "研修の報告" },
    { type: "content", title: "現場で実行する", takeaway: "店長と推進役が改善策を合意", points: ["火曜までに改善ストーリーを完成"] },
    { type: "content", title: "効果を測る", takeaway: "翌月の訪店で定着を確認", points: ["週次で進捗を確認"] },
    { type: "closing", message: "次回会議で判断" },
  ] };
  const chat = buildChatPrompt({ deck, message: "2枚目を直して", current: 1, focus: [1] });
  const revise = buildRevisePrompt({ deck, slideIndex: 1, instruction: "見やすく" });
  for (const prompt of [chat, revise]) {
    assert.match(prompt, /火曜までに改善ストーリー/);
    assert.match(prompt, /翌月の訪店で定着を確認/);
    assert.match(prompt, /写真/);
  }
  assert.match(chat, /資料全体の骨子/);
});

test("repair prompt names the overflowing fields or the validation error", () => {
  const overflow = buildRepairPrompt({ issues: [{ slide: 2, field: "items[1].desc", message: "約8文字超過" }], slides: [{}, {}, { title: "要因" }] });
  assert.match(overflow, /3枚目「要因」の items\[1\]\.desc: 約8文字超過/);
  const invalid = buildRepairPrompt({ error: "指定は8枚ですが、回答は7枚でした。" });
  assert.match(invalid, /回答は7枚/);
});

test("Codex prompt preserves the standalone slide-generation contract", () => {
  const prompt = buildDeckPrompt({
    brief: "生成AI研修の成果を役員向けに報告する。受講者120名、利用率72%。",
    audience: "役員",
    purpose: "意思決定・報告",
    tone: "役員向け・結論先行",
    settings: { slideCount: 5, textDensity: "rich", primaryColor: "#007F40" },
  });

  assert.match(prompt, /指定枚数: 5枚/);
  assert.match(prompt, /1枚目はtitle、最後はclosing/);
  assert.match(prompt, /入力素材にない数値や事実を作らない/);
  assert.match(prompt, /statsCompare/);
  assert.match(prompt, /具体的な次のアクション/);
  assert.match(prompt, /文字量: 多め/);
  assert.match(prompt, /customerExperience/);
});

test("Codex JSON result accepts a fenced structured response", () => {
  const result = parseDeckResult('```json\n{"deckTitle":"成果報告","purpose":"報告","audience":"役員","slideData":[]}\n```');
  assert.equal(result.deckTitle, "成果報告");
  assert.deepEqual(result.slideData, []);
});

test("Codex optional null values are omitted before deck validation", () => {
  const result = omitNullObjectValues({
    slideData: [{ type: "title", date: null, subtitle: null }, { type: "closing", message: "次のアクション" }],
  });
  assert.deepEqual(result, {
    slideData: [{ type: "title" }, { type: "closing", message: "次のアクション" }],
  });
});

test("Codex output schema converts tuple items into Codex-compatible item schemas", () => {
  const outputSchema = normalizeCodexOutputSchema({
    type: "object",
    properties: {
      columns: {
        type: "array",
        items: [
          { type: "array", items: { type: "string" } },
          { type: "array", items: { type: "string" } },
        ],
      },
    },
  });

  assert.equal(outputSchema.properties.columns.items.type, "array");
  assert.equal(outputSchema.properties.columns.items.items.type, "string");
  assert.equal(outputSchema.properties.columns.minItems, 2);
  assert.equal(outputSchema.properties.columns.maxItems, 2);
  assert.equal("prefixItems" in outputSchema.properties.columns, false);
});

test("partialItems returns only the slides that have fully arrived", async () => {
  const { partialItems } = await import("../server/codex-app-server.mjs");
  const text = '```json\n{"deckTitle":"x","slideData":[{"type":"title","title":"a{b}\\"c"},{"type":"content","points":["p]"]},{"type":"clo';
  assert.deepEqual(partialItems(text), [{ type: "title", title: 'a{b}"c' }, { type: "content", points: ["p]"] }]);
  assert.deepEqual(partialItems('{"slideData":'), []);
});

test("partialString shows a JSON string value while it is still being written", async () => {
  const { partialString } = await import("../server/codex-app-server.mjs");
  assert.equal(partialString('{"reply":"3枚目を\\n直し\\"まし', "reply"), '3枚目を\n直し"まし');
  assert.equal(partialString('{"reply":"完了","operations":[]}', "reply"), "完了");
  assert.equal(partialString('{"reply":"\\u3042', "reply"), "あ");
  assert.equal(partialString('{"rep', "reply"), "");
});
