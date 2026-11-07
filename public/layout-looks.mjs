/*
 * Layouts grouped by how they look on the screen. Different layouts can look the same — a timeline,
 * a process and a roadmap are all "a line of points across the slide" — so a deck's variety is judged
 * by these looks, not by type names. Shared by the server (prompts and automatic repair of the AI's
 * choices) and the studio (the check list and the outline editor).
 */

export const LOOKS = {
  flow: { label: "横に並ぶ流れ", examples: "工程・年表・ロードマップ・レーン図", types: ["process", "timeline", "roadmap", "flowChart", "gantt", "diagram"], alike: true },
  rows: { label: "縦に並ぶ行", examples: "箇条書き・要約・FAQ・チェックリスト", types: ["content", "processList", "executiveSummary", "bulletCards", "faq", "checklist"], alike: true },
  cards: { label: "カードの並び", examples: "カード・2×2・3列まとめ・SWOT", types: ["cards", "headerCards", "grid2x2", "headerTwoColumn", "headerThreeSummary", "swot"], alike: true },
  compare: { label: "左右の対比", examples: "比較・前後・数値比較", types: ["compare", "beforeAfter", "statsCompare"], alike: true },
  numbers: { label: "数字・グラフ", examples: "KPI・グラフ・ダッシュボード・ウォーターフォール", types: ["kpi", "dashboard", "waterfall", "imageText"] },
  interactive: { label: "動かして確かめる", examples: "試算（スライダー）・不足と打ち手", types: ["simulator", "gap"] },
  shape: { label: "図形", examples: "サイクル・ピラミッド・ベン図・ロジックツリー・マトリクス", types: ["cycle", "pyramid", "funnel", "triangle", "venn", "stepUp", "logicTree", "orgChart", "matrix"] },
  table: { label: "表", examples: "表", types: ["table"] },
  impact: { label: "大きな一文・写真", examples: "ステートメント・全面写真・引用", types: ["statement", "hero", "quote"] },
};

const LOOK_OF_TYPE = Object.fromEntries(Object.entries(LOOKS).flatMap(([look, { types }]) => types.map((type) => [type, look])));

/** The look of a slide (or an outline row), or null for the cover, the agenda, chapter dividers, the close and deep-dive pages. */
export function lookOf(slide) {
  const type = slide?.type;
  if (slide?.drillOf) return null;
  // A chart with text reads as numbers; a photo with text reads as a photo slide.
  if (type === "imageText" && !(slide.image && typeof slide.image === "object")) {
    const photo = slide.image || slide.customImage || slide.media || (slide.visualAsset && slide.visualAsset !== "none");
    if (photo) return "impact";
  }
  return LOOK_OF_TYPE[type] ?? null;
}

/** How many body slides may share one look: one in four, rounded up (6 body slides → 2, 10 → 3). */
export const maxSameLook = (bodyCount) => Math.max(1, Math.ceil(bodyCount / 4));

/** How many different looks the body slides should use (6 or more body slides → 4). */
export const minLooks = (bodyCount) => (bodyCount < 4 ? 1 : Math.min(4, bodyCount - 2));

const numbers = (indices) => indices.map((i) => i + 1).join("・");

/**
 * Where a deck (or an outline) looks monotonous: too many slides of one look, the same look side by side,
 * or too few looks overall. Each issue names the slides involved (0-based) and a sentence for people.
 */
export function varietyIssues(slides) {
  // Deep-dive pages are not part of the story: the slides on either side of them follow each other.
  const story = (slides ?? []).map((slide, index) => ({ slide, index })).filter(({ slide, index }) => !(slide?.drillOf && index > 0));
  const body = story.map(({ slide, index }, pos) => ({ index, pos, look: lookOf(slide) })).filter((entry) => entry.look);
  const issues = [];
  const limit = maxSameLook(body.length);
  const crowded = new Set();
  for (const [look, info] of Object.entries(LOOKS)) {
    if (!info.alike) continue;
    const indices = body.filter((entry) => entry.look === look).map((entry) => entry.index);
    if (indices.length <= limit) continue;
    crowded.add(look);
    issues.push({
      kind: "variety", severity: "warning", rule: "crowded", look, slides: indices, slide: indices[0], keep: limit,
      message: `「${info.label}」（${info.examples}）の見た目が${indices.length}枚あります（${numbers(indices)}枚目）。本文${body.length}枚なら${limit}枚までにすると単調になりません`,
    });
  }
  body.forEach((entry, k) => {
    const previous = body[k - 1];
    if (!previous || previous.look !== entry.look || !LOOKS[entry.look].alike || crowded.has(entry.look)) return;
    if (entry.pos !== previous.pos + 1) return;
    issues.push({
      kind: "variety", severity: "warning", rule: "adjacent", look: entry.look, slides: [previous.index, entry.index], slide: entry.index, keep: 1,
      message: `${previous.index + 1}枚目と${entry.index + 1}枚目が、どちらも「${LOOKS[entry.look].label}」の見た目で続いています`,
    });
  });
  const used = [...new Set(body.map((entry) => entry.look))];
  const need = minLooks(body.length);
  if (body.length && used.length < need) {
    issues.push({
      kind: "variety", severity: "warning", rule: "few", look: null, slides: body.map((entry) => entry.index), slide: body[0].index, keep: null,
      message: `本文${body.length}枚の見た目が${used.length}種類（${used.map((look) => LOOKS[look].label).join("・")}）だけです。${need}種類以上にすると、数字・図形・大きな一文などで変化がつきます`,
    });
  }
  return issues;
}

/** Which look to use instead, by what the slide actually says. The same wording guides the AI and people. */
export const LOOK_ADVICE = [
  "時期のない「やること・備え・ポイント・特徴」→ cards・checklist・grid2x2・bulletCards",
  "層・段階・レベル → pyramid・stepUp ／ 重なり・3要素の関係 → venn・triangle ／ 繰り返す取り組み → cycle",
  "原因・論点の分解 → logicTree ／ 2軸での位置づけ → matrix ／ 体制・役割 → orgChart",
  "数字が主役 → kpi・imageText（グラフ） ／ 決意・問い・転換点 → statement ／ 声 → quote",
  "条件しだいで結果が変わる試算 → simulator ／ 目標との差を打ち手で埋める → gap",
];
