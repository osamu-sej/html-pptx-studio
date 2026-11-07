// What the server (and the AI) needs to know about objects people placed by hand on a slide
// (public/engine/objects.js draws them): their words, and a one-line summary for prompts.

const ENTITIES = { amp: "&", lt: "<", gt: ">", quot: '"', "#39": "'", nbsp: " " };

/** Rich text (the small HTML subset objects keep) → plain text, a line per paragraph. */
export function plainText(html) {
  return String(html ?? "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|li|div)>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(/&(amp|lt|gt|quot|#39|nbsp);/g, (_, name) => ENTITIES[name])
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{2,}/g, "\n")
    .trim();
}

const KIND_NAMES = { shape: "図形", text: "テキスト", image: "画像", line: "線", icon: "アイコン", video: "動画", audio: "オーディオ", lottie: "アニメーション", table: "表", chart: "グラフ", smartart: "SmartArt", ink: "インク", zoom: "ズーム", camera: "カメオ", equation: "数式", model: "3D モデル", poll: "アンケート" };

/** The words an object holds: its text, a table's cells, a chart's title and labels. */
function wordsOf(o) {
  if (o?.kind === "table" && Array.isArray(o.cells)) return o.cells.flat().map((cell) => plainText(cell?.text)).filter(Boolean).join(" ");
  if (o?.kind === "chart" && o.chart) return [o.chart.title, ...(Array.isArray(o.chart.labels) ? o.chart.labels : [])].filter(Boolean).join(" ");
  if (o?.kind === "poll") return [o.question, ...(Array.isArray(o.options) ? o.options : [])].filter(Boolean).join(" ");
  if (o?.kind === "smartart" && Array.isArray(o.smartart?.items)) return o.smartart.items.map((item) => plainText(item?.text)).filter(Boolean).join(" ");
  return plainText(o?.text);
}

/** "[自由配置: 図形2・テキスト1／文字「…」「…」]" — what the AI sees instead of the objects themselves. */
export function objectsSummary(list) {
  const objects = Array.isArray(list) ? list.filter((o) => o && typeof o === "object") : [];
  if (!objects.length) return undefined;
  const counts = new Map();
  for (const o of objects) counts.set(KIND_NAMES[o.kind] ?? "オブジェクト", (counts.get(KIND_NAMES[o.kind] ?? "オブジェクト") ?? 0) + 1);
  const words = objects.map((o) => wordsOf(o).replace(/\s+/g, " ")).filter(Boolean).map((text) => `「${text.slice(0, 60)}」`);
  const kinds = [...counts].map(([name, n]) => `${name}${n}`).join("・");
  return `[自由配置: ${kinds}${words.length ? `／文字${words.join("").slice(0, 400)}` : ""}]`;
}

/** The words of a slide's objects, for search and for judging what the slide is about. */
export function objectWords(list) {
  return (Array.isArray(list) ? list : []).map((o) => wordsOf(o)).filter(Boolean).join(" ");
}
