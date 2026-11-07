// 校閲 → 翻訳: the deck's words in another language. The studio sends the distinct strings of the slides, their
// objects and notes; the AI returns one translation per string, in the same order (numbers, names and the **強調**
// marks kept). The studio puts each one back where its original was.

import { z } from "zod";

export const LANGS = { en: "英語", zh: "中国語（簡体字）", ko: "韓国語", ja: "日本語" };
export const translateRequestSchema = z.object({
  texts: z.array(z.string().max(4000)).min(1).max(600),
  to: z.enum(Object.keys(LANGS)),
});
export const translateResultSchema = z.object({ texts: z.array(z.string().max(8000)).max(600) });

export function buildTranslatePrompt({ texts, to }) {
  return [
    `あなたはビジネス資料の翻訳者です。次の文字列をそれぞれ${LANGS[to]}に翻訳してください。`,
    "決まり：",
    "- 1つの文字列に1つの訳。数も順番も入力と同じにする（空の文字列は空のまま）。",
    "- 数値・単位・日付・記号、社名・商品名・システム名などの固有名詞はそのまま。",
    "- **語句** の強調の印、改行（\\n）はそのまま残す。",
    "- スライドの見出し・箇条書きにふさわしい短く自然な言い方にする。説明や注釈は付けない。",
    "",
    "入力（JSONの配列）:",
    JSON.stringify(texts),
    "",
    'texts に訳の配列を入れて返してください。',
  ].join("\n");
}

/** The AI's answer checked against what was asked: one string each, the same count. */
export function checkTranslation(request, result) {
  const { texts } = translateResultSchema.parse(result);
  if (texts.length !== request.texts.length) throw new Error(`訳の数が合いません（${request.texts.length}個に対して${texts.length}個）。1つの文字列に1つの訳を、同じ順で返してください。`);
  return texts.map((t, i) => (request.texts[i].trim() ? t : request.texts[i]));
}
