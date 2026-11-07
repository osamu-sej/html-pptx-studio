// How much text each part of a slide can hold before the layout has to shrink it (1920×1080 canvas,
// body text 30px). The browser fits text exactly; this catches what is clearly too long on the server,
// so the AI can shorten it before the deck reaches the user.

const len = (value) => [...String(value ?? "").replace(/\*\*/g, "")].reduce((sum, ch) => sum + (/[\x20-\x7e]/.test(ch) ? 0.55 : 1), 0);

const LIMITS = {
  title: 30, takeaway: 60, subtitle: 60, message: 90, text: 110, conclusion: 70, action: 70, summary: 60,
  point: 60, step: 50, card: 70, faq: 70, cell: 22, leaf: 24, lane: 24, milestone: 26,
};

const NOT_ON_SLIDE = new Set(["customImage", "media", "imagePlacement", "notes", "details", "image", "visualAsset", "photoMotion", "animation", "kinetic", "backdrop", "entrance", "emphasis", "transition", "drillOf"]);

/** A few words the AI can act on for every field that is clearly too long (1.3× the guide). */
export function capacityIssues(slides, { only = null } = {}) {
  const issues = [];
  const over = (slide, field, value, limit) => {
    const n = Math.ceil(len(value));
    if (n > Math.ceil(limit * 1.3)) issues.push({ kind: "overflow", severity: "error", slide, field, message: `${n}字あります。${limit}字以内にしてください` });
  };
  slides.forEach((slide, i) => {
    if (only && !only.has(i)) return;
    const titleLimit = slide.type === "title" ? 36 : slide.type === "hero" ? 40 : LIMITS.title;
    if (slide.title) over(i, "title", slide.title, titleLimit);
    if (slide.takeaway) over(i, "takeaway", slide.takeaway, LIMITS.takeaway);
    if (slide.subtitle) over(i, "subtitle", slide.subtitle, LIMITS.subtitle);
    if (slide.message) over(i, "message", slide.message, LIMITS.message);
    if (slide.text) over(i, "text", slide.text, slide.type === "statement" ? 60 : LIMITS.text);
    if (slide.conclusion) over(i, "conclusion", slide.conclusion, LIMITS.conclusion);
    if (slide.action) over(i, "action", slide.action, LIMITS.action);
    if (slide.summary) over(i, "summary", slide.summary, LIMITS.summary);
    (slide.points ?? []).forEach((point, k) => over(i, `points[${k}]`, point, LIMITS.point));
    (slide.steps ?? []).forEach((step, k) => over(i, `steps[${k}]`, step, slide.type === "process" && (slide.steps?.length ?? 0) >= 5 ? 34 : LIMITS.step));
    const cols = Math.max(1, Math.min(4, slide.columns || (slide.items?.length === 4 ? 2 : slide.items?.length || 1)));
    (slide.items ?? []).forEach((it, k) => {
      if (typeof it === "string") { over(i, `items[${k}]`, it, LIMITS.point); return; }
      if (it.title) over(i, `items[${k}].title`, it.title, 24);
      if (it.desc) over(i, `items[${k}].desc`, it.desc, cols >= 3 ? 48 : LIMITS.card);
      if (it.q) over(i, `items[${k}].q`, it.q, 36);
      if (it.a) over(i, `items[${k}].a`, it.a, LIMITS.faq);
      if (it.value) over(i, `items[${k}].value`, it.value, 10);
    });
    (slide.rows ?? []).forEach((row, r) => row.forEach((cell, c) => over(i, `rows[${r}][${c}]`, cell, LIMITS.cell)));
    (slide.branches ?? []).forEach((branch, b) => (branch.items ?? []).forEach((leaf, k) => over(i, `branches[${b}].items[${k}]`, leaf, LIMITS.leaf)));
    (slide.milestones ?? []).forEach((mile, k) => over(i, `milestones[${k}].label`, mile.label, LIMITS.milestone));
    (slide.measures ?? []).forEach((m, k) => { if (m.title) over(i, `measures[${k}].title`, m.title, 16); if (m.desc) over(i, `measures[${k}].desc`, m.desc, 30); });
    (slide.inputs ?? []).forEach((input, k) => { if (input.label) over(i, `inputs[${k}].label`, input.label, 14); });
    if (slide.source) over(i, "source", slide.source, 60);
    // Pictures, speaker notes and click-to-open details are not text on the slide.
    const bodyChars = JSON.stringify(slide, (key, value) => (NOT_ON_SLIDE.has(key) ? undefined : value)).length;
    if (bodyChars > 2600 && !["table"].includes(slide.type)) issues.push({ kind: "overflow", severity: "warning", slide: i, field: "body", message: "1枚の文字が多すぎます。要素を絞るか2枚に分けてください" });
  });
  return issues;
}
