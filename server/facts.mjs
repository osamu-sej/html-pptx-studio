// Figures the AI writes are checked against the material it was given: a number that appears nowhere in
// the source is most likely invented (or mis-copied), which is the one mistake an executive deck cannot
// afford. Counts and ordinals up to 10 ("3つ", "STEP 2") are ordinary words and are not checked.

const SKIP_KEYS = new Set(["type", "date", "sectionNo", "start", "span", "columns", "icon", "illustration", "visualAsset",
  "customImage", "image64", "animation", "status", "state", "trend", "chartType", "imagePosition", "visual", "done", "twoColumn",
  "target", "media", "photoMotion", "imagePlacement", "kinetic", "backdrop", "entrance", "emphasis", "transition", "drillOf",
  // A simulator's slider ranges and formula constants are settings, not claims; a gantt's "いま" is a position.
  "min", "max", "step", "digits", "formula", "now"]);
const NUMBER = /\d+(?:[.,]\d+)*/g;

/** "１，２００" → "1200", "3.50" → "3.5": one spelling per value, so "1,200店" in the material matches "1200". */
function normalize(token) {
  let value = token.replace(/,(?=\d{3}(?:\D|$))/g, "");
  if (value.includes(".")) value = value.replace(/\.?0+$/, "");
  return value.replace(/^0+(?=\d)/, "");
}

function halfWidth(text) {
  return String(text ?? "").replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0)).replace(/[．]/g, ".").replace(/[，]/g, ",");
}

/** Every value written in the text, normalized. */
export function numbersIn(text) {
  return new Set((halfWidth(text).match(NUMBER) ?? []).map(normalize));
}

const WITH_UNIT = /\d+(?:[.,]\d+)*\s*(?:%|％|pt|ポイント|倍|[万億千]?円|時間|か月|ヶ月|分|秒|店|人|件|回|点|年|月|日|h\/月|h)?/g;

/** The values as written, with their unit ("43%"), for messages. */
function spelled(text, values) {
  const out = [];
  for (const token of halfWidth(text).match(WITH_UNIT) ?? []) {
    const value = normalize(token.match(NUMBER)[0]);
    if (values.includes(value) && !out.some((seen) => seen.startsWith(token.trim()))) out.push(token.trim());
  }
  return out.length ? out : values;
}

function worth(value) {
  return !(Number.isInteger(Number(value)) && Number(value) <= 10);
}

/** [path, text] for every string and number a slide shows (image data and layout settings excluded). */
function entries(value, path = "", out = []) {
  if (typeof value === "number") out.push([path, String(value)]);
  else if (typeof value === "string") { if (!value.startsWith("data:")) out.push([path, value]); }
  else if (Array.isArray(value)) value.forEach((item, i) => entries(item, `${path}[${i}]`, out));
  else if (value && typeof value === "object") {
    for (const [key, child] of Object.entries(value)) if (!SKIP_KEYS.has(key)) entries(child, path ? `${path}.${key}` : key, out);
  }
  return out;
}

/**
 * Figures on the slides that the source does not contain, one issue per field:
 * { slide, field, kind: "unsourced", severity, values, message }. `only` limits the check to some slides.
 */
export function unsourcedNumbers(slides, source, { only = null } = {}) {
  const known = numbersIn(source);
  const issues = [];
  slides.forEach((slide, index) => {
    if (only && !only.has(index)) return;
    for (const [field, text] of entries(slide)) {
      const missing = [...numbersIn(text)].filter((value) => worth(value) && !known.has(value));
      if (missing.length) {
        const values = spelled(text, missing);
        issues.push({ slide: index, field, kind: "unsourced", severity: "warning", values,
          message: `素材にない数値があります（${values.slice(0, 3).map((value) => `「${value}」`).join("")}）` });
      }
    }
  });
  return issues;
}

/** The retry request for invented figures, phrased like the overflow repair. */
export function factRepairLines(issues, slides, where = (index) => `${index + 1}枚目`) {
  if (!issues.length) return [];
  return [
    "次の数値は素材のどこにもありません。素材にある数値に直すか、数値を使わない表現にしてください（計算で求めた値なら、元になった素材の数値も並べて書く）。",
    ...issues.map((issue) => `- ${where(issue.slide)}「${slides[issue.slide]?.title ?? ""}」の ${issue.field}: ${issue.values.join("、")}`),
  ];
}

/** "3枚目「43%」ほか2か所" for progress messages and toasts. */
export function describeUnsourced(issues, where = (index) => `${index + 1}枚目`) {
  if (!issues.length) return "";
  const first = issues[0];
  return `素材にない数値が${issues.length}か所あります（${where(first.slide)}「${first.values[0]}」${issues.length > 1 ? "ほか" : ""}）。確認してください`;
}
