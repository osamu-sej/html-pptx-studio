// 校閲 → 類義語: words that can stand for the selected one in a slide. The studio has a small thesaurus of its own
// (public/editor/thesaurus.mjs); this asks the AI for the words that are not there. Short, usable words only.

import { z } from "zod";

export const synonymsRequestSchema = z.object({
  word: z.string().trim().min(1).max(40),
  // The sentence the word sits in, so a word with several meanings is taken in the right one.
  context: z.string().max(400).optional(),
});
export const synonymsResultSchema = z.object({
  senses: z.array(z.object({ label: z.string().max(40), words: z.array(z.string().max(30)).max(20) })).max(6),
});

export function buildSynonymsPrompt({ word, context }) {
  return [
    "あなたはビジネス資料の言葉選びを助ける編集者です。次の語の類義語（言い換え）を、意味ごとに挙げてください。",
    "決まり：",
    "- 意味（sense）は多くても3つ。それぞれに短い名前（label。例：「問題・課題」）と、言い換えに使える語（words）を3〜8個。",
    "- スライドの見出し・箇条書きに置ける、短く自然な語だけ（長い言い回しや説明は入れない）。",
    "- 語そのものは入れない。専門用語・社名・商品名は言い換えない（その語しか無いときは senses を空にする）。",
    context ? `- 語が使われている文は「${context.replace(/[\r\n]+/g, " ").slice(0, 200)}」。この文の意味を先に挙げる。` : "",
    "",
    `語: ${word}`,
    "",
    "senses に [{ label, words }] を入れて返してください。",
  ].filter(Boolean).join("\n");
}

/** The AI's answer tidied: the word itself, repeats, empty senses and over-long words taken out. */
export function checkSynonyms(request, result) {
  const { senses } = synonymsResultSchema.parse(result);
  const seen = new Set([request.word.trim()]);
  const out = [];
  for (const sense of senses) {
    const words = [];
    for (const raw of sense.words) {
      const word = raw.trim();
      if (!word || word.length > 20 || seen.has(word)) continue;
      seen.add(word);
      words.push(word);
    }
    if (words.length) out.push({ label: sense.label.trim() || "類義語", words });
  }
  return out;
}
