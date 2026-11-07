/** Turn a slide's data structure into a visual brief; title alone is never enough. */
import { slideMeaning } from "./visual-relevance.mjs";

export function imageBrief(slide) {
  const claim = slide.takeaway || slide.message || slide.title || "";
  const body = slideMeaning(slide).replace(claim, "").trim().slice(0, 500);
  if (slide.type === "gantt") {
    const periods = slide.periods || [];
    const steps = (slide.items || []).map((item) => {
      const start = Math.max(0, item.start || 0);
      const end = Math.min(periods.length, start + Math.max(1, item.span || 1));
      return `${item.title}: ${periods.slice(start, end).join("→") || "時期未指定"}${item.desc ? `／${item.desc}` : ""}`;
    });
    const scene = "具体的構図: 最重要の作業をしている手元・道具・場所を1つの連続した場面にする。工程にない行為や成果は描かない。";
    return [
      `このページは工程表。最重要の主張は「${claim}」。`,
      `時間軸: ${periods.join(" → ") || "未指定"}。`,
      `工程と時期: ${steps.join("。 ")}。`,
      scene,
      "写真・挿絵は工程表そのものの代替ではなく、最重要の工程を1場面で補強する。最重要の時間帯を視覚の中心に置く。",
      "複数の小さなパネル、無関係な店舗アイコン群、抽象的な3D業務カード、読めない擬似文字は禁止。",
    ].join("\n");
  }
  if (slide.type === "table") {
    return `このページは表。主張「${claim}」。列: ${(slide.headers || []).join("／")}。行の要点: ${(slide.rows || []).slice(0, 5).map((row) => row.join("／")).join("。 ")}。表の意味を補強する単一の場面を描き、表を画像で作り直さない。`;
  }
  if (slide.type === "timeline") {
    return `このページは時系列。主張「${claim}」。順序: ${(slide.milestones || []).map((m) => `${m.date}: ${m.label}`).join(" → ")}。重要な変化を1場面で表現する。`;
  }
  return `このページの主張は「${claim}」。本文の具体的内容: ${body || "主張を表す一場面"}。本文に書かれた対象・場所・行動から、主張が伝わる一場面を選ぶ。図表や数字を画像内に描き直さない。本文にない人物・会議・商品・成果を足さない。`;
}
