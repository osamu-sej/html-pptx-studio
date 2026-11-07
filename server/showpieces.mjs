// 見せ場（showpieces）: a generated deck always has a few moments that only HTML can give — parts turning in 3D,
// whirling in, big words that zoom or flip, a moving graphic behind the cover, a dive into the next chapter.
// The AI is asked to place them; whatever it left out is placed here on the slides that matter most (figures,
// the one-line message, the turning points, the cover and the close), spread through the deck and never the
// same twice in a row. What the AI (or the user) chose for a slide stays as it is.
// Pure: (slides) → { slides, added } where added lists [index, what] for the job's detail.

/** Motion that counts as a showpiece on a slide. */
const DYNAMIC_ENTRANCES = new Set(["spin", "swirl", "flip", "zoom", "drop", "pop"]);
const DYNAMIC_TRANSITIONS = new Set(["flip", "dive", "curtain", "zoom", "circle", "morph"]);
const DYNAMIC_KINETIC = new Set(["chars", "type", "scramble", "wave", "zoom", "flip", "slide"]);

const set = (value) => value != null && value !== "auto" && value !== "";

/** Does the slide already have a showpiece? */
export function isShowpiece(slide) {
  return DYNAMIC_ENTRANCES.has(slide?.entrance) || DYNAMIC_TRANSITIONS.has(slide?.transition) || DYNAMIC_KINETIC.has(slide?.kinetic)
    || (set(slide?.backdrop) && slide.backdrop !== "none");
}

const FIGURES = new Set(["kpi", "statsCompare", "waterfall", "dashboard", "funnel", "simulator", "gap"]);
const CARDS = new Set(["cards", "headerCards", "bulletCards", "triangle", "pyramid", "stepUp", "cycle", "grid2x2", "checklist",
  "compare", "beforeAfter", "swot", "venn", "matrix", "headerThreeSummary", "executiveSummary", "logicTree"]);
const hasChart = (slide) => slide?.type === "imageText" && slide.image && typeof slide.image === "object";

/** How much a slide is a moment to make the most of (0: never). */
function weight(slide, index, count) {
  if (!slide || slide.drillOf) return 0;
  const type = slide.type;
  if (type === "hero" || type === "statement") return 6;
  if (FIGURES.has(type) || hasChart(slide)) return 5;
  if (type === "executiveSummary") return 4;
  if (type === "closing") return 4;
  if (index === 0 || type === "title") return 4;
  if (type === "section") return 3;
  if (type === "beforeAfter" || type === "logicTree" || type === "roadmap") return 3;
  if (CARDS.has(type)) return 2;
  return index > 0 && index < count - 1 ? 1 : 0;
}

/** How many showpieces a deck of `count` slides has: 2 for a short deck, about one in three for a long one (at most 6). */
export function showpieceCount(count) {
  if (count <= 2) return count >= 2 ? 1 : 0;
  if (count <= 5) return 2;
  if (count <= 9) return 3;
  return Math.min(6, Math.ceil(count / 3));
}

// The turns used one after another, so two showpieces are never the same.
const ENTRANCE_TURNS = ["spin", "swirl", "flip", "spin", "drop", "swirl"];

/** The motion a showpiece slide gets (only the fields still empty are filled). */
function bundle(slide, turn) {
  const type = slide.type;
  const entrance = ENTRANCE_TURNS[turn % ENTRANCE_TURNS.length];
  if (type === "title") return { kinetic: "flip", backdrop: "rays" };
  if (type === "closing") return { kinetic: "chars", backdrop: "rays", entrance: "swirl" };
  if (type === "section") return { transition: "dive", backdrop: "orbits", kinetic: "slide" };
  if (type === "statement") return { kinetic: "zoom", backdrop: "ripple", transition: "zoom" };
  if (type === "hero") return { kinetic: "zoom", photoMotion: "tilt", transition: "dive" };
  if (FIGURES.has(type) || hasChart(slide)) return { entrance: entrance === "drop" ? "spin" : entrance, transition: "zoom" };
  return { entrance, transition: turn % 2 ? "flip" : "zoom" };
}

export function stageShowpieces(slides) {
  const list = Array.isArray(slides) ? slides.map((slide) => ({ ...slide })) : [];
  const count = list.length;
  const added = [];
  const want = showpieceCount(count);
  let have = list.filter(isShowpiece).length;
  if (have >= want) return { slides: list, added };
  // The strongest moments first; among equals, the ones further from a showpiece already placed.
  const placed = new Set(list.map((slide, index) => (isShowpiece(slide) ? index : -1)).filter((index) => index >= 0));
  const candidates = list.map((slide, index) => ({ index, w: weight(slide, index, count) })).filter((c) => c.w > 0 && !placed.has(c.index));
  let turn = 0;
  while (have < want && candidates.length) {
    const distance = (index) => (placed.size ? Math.min(...[...placed].map((p) => Math.abs(p - index))) : count);
    candidates.sort((a, b) => b.w + Math.min(distance(b.index), 3) * 0.6 - (a.w + Math.min(distance(a.index), 3) * 0.6) || a.index - b.index);
    const { index } = candidates.shift();
    const slide = list[index];
    const fields = Object.entries(bundle(slide, turn)).filter(([key]) => !set(slide[key]));
    // Photo motion is for a slide with a photo; a hero without one keeps its words' motion only.
    const usable = fields.filter(([key]) => key !== "photoMotion" || set(slide.visualAsset) && slide.visualAsset !== "none");
    if (!usable.length) continue;
    for (const [key, value] of usable) slide[key] = value;
    added.push([index, Object.fromEntries(usable)]);
    placed.add(index);
    have += 1;
    turn += 1;
  }
  return { slides: list, added };
}
