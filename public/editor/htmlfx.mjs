// HTMLの動きをおまかせで付ける: what only HTML can do, put on a slide's objects by what they are — the title rises
// out of a mask, charts grow from their axis (and answer the mouse), big figures count up, lines are drawn, cards
// lift under the mouse, pictures zoom in on a click. What the slide already has (its own animations, a click
// action, a hover, a loop) stays as it is, so a deck brought over from PowerPoint keeps its own animations.
// The template's furniture (a logo, 秘（B）, a slogan, a rule: the same object in the same place on slide after slide,
// as a deck brought over from PowerPoint carries it) is left alone.
// Pure: (slide, E, { chrome }) → { elements, timeline, added } — the app makes it one undo step.

const SLIDE_AREA = 1920 * 1080;

/** Plain words of an object. */
const wordsOf = (o, E) => (o.text ? E.richToText(o.text).trim() : "");

/** The biggest text size an object uses (its own size or a run's). */
function biggestSize(o) {
  let size = Number(o.fs) || 0;
  for (const m of String(o.text || "").matchAll(/font-size:\s*([\d.]+)px/g)) size = Math.max(size, Number(m[1]));
  return size;
}

/** The object that is the slide's title: the words of slide.title, else the biggest words near the top. */
function titleOf(slide, list, E) {
  const title = String(slide.title || "").trim();
  const texts = list.filter((o) => ["text", "shape"].includes(o.kind) && wordsOf(o, E));
  const same = title ? texts.find((o) => wordsOf(o, E).replace(/\s+/g, "") === title.replace(/\s+/g, "")) : null;
  if (same) return same;
  const top = texts.filter((o) => o.y < 220 && (o.fill === "none" || o.kind === "text")).sort((a, b) => biggestSize(b) - biggestSize(a));
  return top[0] && biggestSize(top[0]) >= 40 ? top[0] : null;
}

/** What makes an object the same on another slide: its kind, box and content. */
export function signature(o) {
  const box = [o.x, o.y, o.w, o.h, o.x1, o.y1, o.x2, o.y2].map((v) => (v == null ? "" : Math.round(Number(v)))).join(",");
  const content = o.kind === "image" ? `${String(o.src || "").length}:${String(o.src || "").slice(-48)}` : String(o.text || o.icon || o.shape || "");
  return `${o.kind}|${box}|${content}`;
}
/** The deck's furniture: objects repeated, unchanged and in place, on several of its slides. */
export function chromeOf(slides) {
  const withObjects = (Array.isArray(slides) ? slides : []).filter((slide) => slide?.elements?.length && !slide.hidden);
  const counts = new Map();
  for (const slide of withObjects) for (const sig of new Set(slide.elements.map(signature))) counts.set(sig, (counts.get(sig) || 0) + 1);
  const least = Math.max(2, Math.ceil(withObjects.length * 0.4));
  return new Set([...counts].filter(([, n]) => n >= least).map(([sig]) => sig));
}

let seq = 0;
const newId = () => `x${Date.now().toString(36).slice(-5)}${(seq++).toString(36)}`;

/**
 * Suggestions for one slide. Returns the slide's new objects and animations, and what was added
 * ({ title, charts, numbers, lines, cards, pictures }).
 */
export function enhanceSlide(slide, E, { chrome = new Set() } = {}) {
  const list = (Array.isArray(slide?.elements) ? slide.elements : []).map((o) => ({ ...o }));
  const timeline = Array.isArray(slide?.timeline) ? slide.timeline.map((e) => ({ ...e })) : [];
  const added = { title: 0, charts: 0, numbers: 0, lines: 0, cards: 0, pictures: 0 };
  if (!list.length) return { elements: list, timeline, added, count: 0 };
  // What already comes in by an animation keeps it (a group's animation counts for its members).
  const animated = new Set();
  for (const e of timeline) {
    if (e.el?.startsWith("grp:")) for (const o of list) { if (o.group === e.el.slice(4)) animated.add(o.id); }
    else animated.add(e.el);
  }
  const furniture = new Set(list.filter((o) => chrome.has(signature(o))).map((o) => o.id));
  const free = (o) => !o.hidden && !animated.has(o.id) && !furniture.has(o.id);
  // As the slide arrives (before its first click): the title, then charts, figures and lines one after another.
  const arrive = [];
  const push = (o, fx, dur) => {
    arrive.push({ id: newId(), el: o.id, cls: "in", fx, start: arrive.length ? "with" : "after", dur, delay: arrive.length ? Math.min(900, 160 * arrive.length) : 0 });
    animated.add(o.id);
  };
  const title = titleOf(slide, list.filter((o) => !furniture.has(o.id)), E);
  if (title && free(title)) { push(title, "maskRise", 900); added.title += 1; }
  for (const o of list) {
    if (!free(o)) continue;
    if (o.kind === "chart") { push(o, "chartGrow", 1300); added.charts += 1; continue; }
    // A SmartArt's items float in one after another.
    if (o.kind === "smartart") { push(o, "floatIn", 700); arrive[arrive.length - 1].by = "item"; added.diagrams = (added.diagrams || 0) + 1; continue; }
    if (o.kind === "line" && list.length <= 60) { push(o, "draw", 900); added.lines += 1; continue; }
    // A 3D model arrives turning (到着).
    if (o.kind === "model") { push(o, "arrive3d", 1600); added.models = (added.models || 0) + 1; continue; }
    if (["text", "shape"].includes(o.kind) && o !== title && /\d/.test(wordsOf(o, E)) && biggestSize(o) >= 56 && wordsOf(o, E).length <= 24) { push(o, "countUp", 1400); added.numbers += 1; }
  }
  // Under the mouse and on a click (what PowerPoint cannot do at all).
  const cards = list.filter((o) => o.kind === "shape" && o !== title && !furniture.has(o.id) && o.fill && o.fill !== "none" && !o.hover && !o.action && wordsOf(o, E) && o.w * o.h < SLIDE_AREA * 0.3);
  const lift = new Set();
  if (cards.length >= 2) {
    for (const o of cards) {
      lift.add(o.id);
      // The words and marks grouped with a card move with it.
      if (o.group) for (const m of list) if (m.group === o.group && !m.hover && !m.action) lift.add(m.id);
    }
  }
  for (const o of list) {
    if (lift.has(o.id)) { o.hover = "lift"; added.cards += o.kind === "shape" ? 1 : 0; }
    if (o.kind === "image" && !o.action && !o.hidden && !furniture.has(o.id) && o.w < 1500 && o.h < 900 && o.w * o.h > 120 * 120) { o.action = { type: "zoom" }; if (!o.hover) o.hover = "zoom"; added.pictures += 1; }
  }
  // Animations that come as the slide arrives go first; the slide's own clicks follow.
  const next = arrive.length ? [...arrive, ...timeline.map((e, i) => (i === 0 && e.start !== "click" ? { ...e, start: "after" } : e))] : timeline;
  const count = Object.values(added).reduce((a, b) => a + b, 0);
  return { elements: list, timeline: next, added, count };
}

/** In words: what おまかせ added ("タイトル・グラフ2・数字1"). */
export function describeAdded(added) {
  const names = { title: "タイトル", charts: "グラフ", diagrams: "SmartArt", numbers: "数字", lines: "線", cards: "カード", pictures: "写真", models: "3D モデル" };
  return Object.entries(added).filter(([, n]) => n > 0).map(([k, n]) => (n > 1 ? `${names[k]}${n}` : names[k])).join("・");
}
