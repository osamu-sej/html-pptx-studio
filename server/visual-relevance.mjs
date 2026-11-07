import { isDeepStrictEqual } from "node:util";

import { objectWords } from "./objects.mjs";

// A stock photo is a claim about what this slide depicts. Require evidence in the
// slide's message (or in an explicit user instruction), not just the deck title.
const PHOTO_TOPICS = {
  ai: /生成AI|人工知能|\bAI\b|LLM|チャットGPT|ChatGPT/i,
  aiWorkflow: /生成AI|人工知能|\bAI\b|自動化|ワークフロー|業務フロー/i,
  promptDesign: /プロンプト|入力文|指示文/i,
  businessWorkshop: /研修|会議|勉強会|講習|受講|学習|教育|ワークショップ/i,
  businessEtiquette: /接遇|接客マナー|礼儀|応対|ビジネスマナー/i,
  executiveDecision: /役員|経営|意思決定|投資|承認|決裁|戦略/i,
  storeOperations: /店舗|売場|店長|加盟店|現場|訪店|巡回|発注|在庫|品揃え|オペレーション/i,
  dataInsight: /データ|分析|KPI|指標|数値|前年比|実績|効果測定/i,
  transformationRoadmap: /改革|変革|ロードマップ|展開|導入|計画|移行|定着/i,
  customerExperience: /顧客|お客様|接客|購買体験|CX|満足度/i,
};

export function slideMeaning(slide) {
  const skip = new Set(["type", "notes", "customImage", "imagePlacement", "image", "visualAsset", "icon", "animation", "photoMotion", "kinetic", "backdrop", "entrance", "emphasis", "transition", "formula", "media", "details", "date", "drillOf", "timeline", "sid"]);
  const parts = [];
  const visit = (value, key = "") => {
    // Objects placed by hand count for their words only.
    if (key === "elements") { const words = objectWords(value); if (words) parts.push(words); return; }
    if (skip.has(key) || value == null) return;
    if (typeof value === "string") { if (!value.startsWith("data:")) parts.push(value); }
    else if (Array.isArray(value)) value.forEach((item) => visit(item));
    else if (typeof value === "object") Object.entries(value).forEach(([childKey, child]) => visit(child, childKey));
  };
  visit(slide);
  return parts.join(" ").slice(0, 2000);
}

export function reconcileChatVisuals(deck, answer, instruction) {
  const wantsAsset = /写真|画像|イラスト|挿絵|ビジュアル|素材/.test(instruction);
  const warnings = [];
  const operations = [];
  for (const op of answer.operations ?? []) {
    if (!op.content || !["replace", "insert"].includes(op.op)) { operations.push(op); continue; }
    const original = op.op === "replace" ? deck.slides[op.slide - 1] : null;
    const content = { ...op.content };
    for (const field of ["visualAsset"]) {
      if (!wantsAsset) {
        if (Object.hasOwn(content, field) && content[field] !== original?.[field]) {
          warnings.push(`${op.slide}枚目の写真は依頼対象ではないため変更しませんでした。`);
        }
        if (original?.[field] === undefined) delete content[field];
        else content[field] = original[field];
        continue;
      }
      const value = content[field];
      if (!value || value === "none" || value === original?.[field]) continue;
      const topic = PHOTO_TOPICS[value];
      const meaning = `${slideMeaning(content)} ${instruction}`;
      if (!topic?.test(meaning)) {
        if (original?.[field] === undefined) delete content[field];
        else content[field] = original[field];
        warnings.push(`${op.slide}枚目の写真は内容との関連を確認できず、変更しませんでした。`);
      }
    }
    if (original && isDeepStrictEqual(content, original)) continue;
    operations.push({ ...op, content });
  }
  return { answer: { ...answer, operations }, warnings };
}
