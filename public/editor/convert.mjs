// 図形に変換: a slide the AI laid out becomes objects people edit one by one, the way PowerPoint turns a SmartArt
// graphic into shapes. The slide is drawn off screen at full size and read back: every box with a fill or a border
// becomes a shape (holding its text when it has some), every run of text a text box, rules become lines, pictures,
// icons, charts and tables their own objects, and an SVG drawing its shapes, lines and words. The pieces of one item
// are grouped, and the layout's build (click by click, cascade) becomes animations on those groups, so the slide
// still presents the way it did. The slide turns into 白紙, keeping its title, key message, notes and source.
// Colours that are (nearly) the theme's become the theme's own, so they follow a change of theme; any other
// colour is kept as it was drawn.

import { distToSegment } from "./ops.mjs";

const SVG_NS = "http://www.w3.org/2000/svg";

/** Pages that are layouts of their own (the cover, chapters, closing, full-bleed pages) keep their layout. */
export const KEEPS_LAYOUT = new Set(["title", "section", "closing", "hero", "statement", "blank"]);
/** What a converted slide keeps; the layout's own fields (points, items, chart data…) go. */
export const KEPT_FIELDS = ["title", "subhead", "takeaway", "source", "details", "drillOf", "notes", "notesRich", "sid", "transition", "transitionDur", "transitionSound", "transitionDir", "advance", "kinetic", "backdrop", "entrance", "emphasis"];

// ---------------------------------------------------------------- colours

const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
const r2 = (n) => Math.round(n * 100) / 100;

/** A computed CSS colour (rgb(), rgba(), color(srgb …), #hex) as { r, g, b, a }; null when nothing shows. */
export function parseColor(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text || text === "none" || text === "transparent") return null;
  const alpha = (v) => (v == null ? 1 : v.endsWith("%") ? Number(v.slice(0, -1)) / 100 : Number(v));
  const make = (r, g, b, a) => {
    const c = { r: clamp(Math.round(Number(r)), 0, 255), g: clamp(Math.round(Number(g)), 0, 255), b: clamp(Math.round(Number(b)), 0, 255), a: clamp(Number(a), 0, 1) };
    return [c.r, c.g, c.b, c.a].every(Number.isFinite) && c.a > 0.004 ? c : null;
  };
  let m = /^rgba?\(\s*([\d.]+)[\s,]+([\d.]+)[\s,]+([\d.]+)(?:[\s,/]+([\d.]+%?))?\s*\)$/.exec(text);
  if (m) return make(m[1], m[2], m[3], alpha(m[4]));
  m = /^color\(srgb\s+([-\d.e]+)\s+([-\d.e]+)\s+([-\d.e]+)(?:\s*\/\s*([\d.]+%?))?\s*\)$/.exec(text);
  if (m) return make(m[1] * 255, m[2] * 255, m[3] * 255, alpha(m[4]));
  m = /^#([0-9a-f]{6})([0-9a-f]{2})?$/.exec(text);
  if (m) return make(parseInt(m[1].slice(0, 2), 16), parseInt(m[1].slice(2, 4), 16), parseInt(m[1].slice(4, 6), 16), m[2] ? parseInt(m[2], 16) / 255 : 1);
  if (text === "white") return make(255, 255, 255, 1);
  if (text === "black") return make(0, 0, 0, 1);
  return null;
}
const rgbOf = (hex) => [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
export const hexOf = (c) => `#${[c.r, c.g, c.b].map((v) => v.toString(16).padStart(2, "0")).join("")}`;
/** What a see-through colour looks like over the white page. */
const overWhite = (c) => [c.r, c.g, c.b].map((v) => v * c.a + 255 * (1 - c.a));

/**
 * The palette colour (with an opacity, over white) that looks like `c`: a tint of a theme colour mixed with white
 * comes back as its colour, faded. Returns { color, opacity, err } (opacity 1 when it is the colour itself; err is
 * how far, 0–255, the look is from `c`).
 */
export function snapToPalette(c, palette, { fade = true } = {}) {
  if (!c) return null;
  const seen = overWhite(c);
  let best = null;
  for (const hex of palette) {
    const p = rgbOf(hex);
    let alpha = 1;
    if (fade) {
      let num = 0;
      let den = 0;
      for (let i = 0; i < 3; i += 1) { num += (255 - p[i]) * (255 - seen[i]); den += (255 - p[i]) ** 2; }
      alpha = den ? clamp(num / den, 0.02, 1) : 1;
    }
    const err = Math.max(...[0, 1, 2].map((i) => Math.abs(p[i] * alpha + 255 * (1 - alpha) - seen[i])));
    // The closest look wins; a faded colour pays a little for its fading, so a pale palette colour that matches
    // nearly as well is taken as it is.
    const score = err + (1 - alpha) * 2;
    if (!best || score < best.score) best = { color: hex, opacity: Math.round(alpha * 100) / 100, score, err };
  }
  return best && { color: best.color, opacity: best.opacity >= 0.98 ? 1 : best.opacity, err: best.err };
}

// ---------------------------------------------------------------- CSS geometry

/** "calc(100% - 44px)", "50%", "12px" → pixels of `size`. */
export function cssLength(token, size) {
  const text = String(token || "").trim();
  const body = /^calc\((.*)\)$/s.exec(text)?.[1] ?? text;
  const parts = body.match(/[+-](?=\s)|-?[\d.]+(?:e-?\d+)?(?:%|px)?/g) || [];
  let total = 0;
  let sign = 1;
  for (const part of parts) {
    if (part === "+" || part === "-") { sign = part === "-" ? -1 : 1; continue; }
    const n = parseFloat(part);
    if (!Number.isFinite(n)) continue;
    total += sign * (part.endsWith("%") ? (n / 100) * size : n);
    sign = 1;
  }
  return total;
}

/** A computed `clip-path: polygon(…)` as points (fractions of the box). */
export function parsePolygon(value, w, hh) {
  const m = /^polygon\((.*)\)$/s.exec(String(value || "").trim());
  if (!m || !(w > 0) || !(hh > 0)) return null;
  const body = m[1].replace(/^\s*(nonzero|evenodd)\s*,/, "");
  const points = [];
  let depth = 0;
  let cur = "";
  for (const ch of body) {
    if (ch === "(") depth += 1;
    if (ch === ")") depth -= 1;
    if (ch === "," && depth === 0) { points.push(cur); cur = ""; } else cur += ch;
  }
  if (cur.trim()) points.push(cur);
  const out = [];
  for (const point of points) {
    const pair = [];
    depth = 0;
    cur = "";
    for (const ch of point.trim()) {
      if (ch === "(") depth += 1;
      if (ch === ")") depth -= 1;
      if (/\s/.test(ch) && depth === 0) { if (cur) pair.push(cur); cur = ""; } else cur += ch;
    }
    if (cur) pair.push(cur);
    if (pair.length !== 2) return null;
    out.push([cssLength(pair[0], w) / w, cssLength(pair[1], hh) / hh]);
  }
  return out.length > 2 ? out : null;
}

/**
 * A computed `repeating-linear-gradient(…)` (a ruled pattern or a hatch) over a w × h box: the thin coloured bands
 * of a horizontal or vertical pattern as lines, anything else as its average colour.
 * Returns { lines: [{ x1, y1, x2, y2, w, color }] } or { fill: { r, g, b, a } } or null.
 */
export function patternOf(image, w, hh) {
  const m = /^repeating-linear-gradient\((.*)\)$/s.exec(String(image || "").trim());
  if (!m || !(w > 0) || !(hh > 0)) return null;
  const parts = [];
  let depth = 0;
  let cur = "";
  for (const ch of m[1]) {
    if (ch === "(") depth += 1;
    if (ch === ")") depth -= 1;
    if (ch === "," && depth === 0) { parts.push(cur.trim()); cur = ""; } else cur += ch;
  }
  if (cur.trim()) parts.push(cur.trim());
  let angle = 180;
  if (/^(-?[\d.]+deg|to )/.test(parts[0])) {
    const head = parts.shift();
    angle = { "to right": 90, "to left": 270, "to bottom": 180, "to top": 0 }[head] ?? parseFloat(head);
  }
  const across = Math.abs(((angle % 180) + 180) % 180 - 90) < 0.5; // bands run left to right: vertical lines
  const along = Math.abs(((angle % 180) + 180) % 180) < 0.5; // top to bottom: horizontal lines
  const size = across ? w : along ? hh : Math.abs(w * Math.sin((angle * Math.PI) / 180)) + Math.abs(hh * Math.cos((angle * Math.PI) / 180));
  const stops = [];
  for (const part of parts) {
    const colorText = (part.match(/^(rgba?\([^)]*\)|color\([^)]*\)|#[0-9a-f]+|[a-z]+)/i) || [])[0];
    if (!colorText) return null;
    const color = parseColor(colorText);
    const positions = part.slice(colorText.length).trim();
    const tokens = [];
    depth = 0;
    cur = "";
    for (const ch of positions) {
      if (ch === "(") depth += 1;
      if (ch === ")") depth -= 1;
      if (/\s/.test(ch) && depth === 0) { if (cur) tokens.push(cur); cur = ""; } else cur += ch;
    }
    if (cur) tokens.push(cur);
    const at = tokens.map((t) => cssLength(t, size));
    if (!at.length) at.push(stops.length ? stops.at(-1).at : 0);
    for (const p of at) stops.push({ color, at: p });
  }
  const period = stops.at(-1)?.at - stops[0]?.at;
  if (!(period > 0.5)) return null;
  // Coloured bands within one period.
  const bands = [];
  for (let i = 0; i < stops.length - 1; i += 1) {
    const [a, b] = [stops[i], stops[i + 1]];
    if (b.at - a.at <= 0) continue;
    const color = a.color && b.color ? { ...a.color, a: (a.color.a + b.color.a) / 2 } : a.color || b.color;
    const alpha = ((a.color?.a || 0) + (b.color?.a || 0)) / 2;
    if (color && alpha > 0.01) bands.push({ from: a.at - stops[0].at, to: b.at - stops[0].at, color: { ...color, a: alpha } });
  }
  if (!bands.length) return null;
  const covered = bands.reduce((sum, band) => sum + band.to - band.from, 0) / period;
  if ((across || along) && covered < 0.2) {
    const lines = [];
    for (let k = 0; k * period < size + 0.5 && lines.length < 60; k += 1) {
      for (const band of bands) {
        const mid = k * period + (band.from + band.to) / 2;
        if (mid > size + 0.5) continue;
        const width = band.to - band.from;
        lines.push(across ? { x1: mid, y1: 0, x2: mid, y2: hh, w: width, color: band.color } : { x1: 0, y1: mid, x2: w, y2: mid, w: width, color: band.color });
      }
    }
    return { lines };
  }
  // A hatch: its average tone.
  const strongest = bands.reduce((a, b) => (b.color.a > a.color.a ? b : a)).color;
  return { fill: { ...strongest, a: strongest.a * covered } };
}

// ---------------------------------------------------------------- SVG paths

const ARGS = { m: 2, l: 2, h: 1, v: 1, c: 6, s: 4, q: 4, t: 2, a: 7, z: 0 };

/**
 * An SVG path's subpaths in absolute commands: [{ d, closed, straight, pts }] — `pts` are the corners of a path
 * made of straight lines only (curves are measured along the path instead).
 */
export function pathSubpaths(d) {
  const tokens = String(d || "").match(/[a-df-z]|[-+]?(?:\d*\.\d+|\d+\.?\d*)(?:e[-+]?\d+)?/gi) || [];
  const subs = [];
  let sub = null;
  let cur = [0, 0];
  let start = [0, 0];
  let lastCtrl = null;
  let lastCmd = "";
  let i = 0;
  const f = (n) => r2(n);
  const begin = (p) => { sub = { d: `M${f(p[0])} ${f(p[1])}`, closed: false, straight: true, pts: [p] }; subs.push(sub); };
  while (i < tokens.length) {
    let cmd = tokens[i];
    if (/[a-z]/i.test(cmd)) i += 1;
    else if (lastCmd) cmd = lastCmd.toLowerCase() === "m" ? (lastCmd === "m" ? "l" : "L") : lastCmd;
    else break;
    const lower = cmd.toLowerCase();
    const rel = cmd !== cmd.toUpperCase();
    const n = ARGS[lower];
    if (n === undefined) break;
    if (lower === "z") {
      if (sub) { sub.d += " Z"; sub.closed = true; }
      cur = [...start];
      lastCtrl = null;
      lastCmd = cmd;
      continue;
    }
    const args = tokens.slice(i, i + n).map(Number);
    if (args.length < n || args.some((v) => !Number.isFinite(v))) break;
    i += n;
    const pt = (x, y) => (rel ? [cur[0] + x, cur[1] + y] : [x, y]);
    if (lower === "m") {
      cur = pt(args[0], args[1]);
      start = [...cur];
      begin(cur);
      lastCtrl = null;
    } else {
      if (!sub) begin(cur);
      if (lower === "l" || lower === "h" || lower === "v") {
        const next = lower === "l" ? pt(args[0], args[1]) : lower === "h" ? [rel ? cur[0] + args[0] : args[0], cur[1]] : [cur[0], rel ? cur[1] + args[0] : args[0]];
        sub.d += ` L${f(next[0])} ${f(next[1])}`;
        sub.pts.push(next);
        cur = next;
        lastCtrl = null;
      } else if (lower === "c" || lower === "s") {
        const c1 = lower === "c" ? pt(args[0], args[1]) : lastCtrl && /[cs]/i.test(lastCmd) ? [2 * cur[0] - lastCtrl[0], 2 * cur[1] - lastCtrl[1]] : [...cur];
        const c2 = lower === "c" ? pt(args[2], args[3]) : pt(args[0], args[1]);
        const end = lower === "c" ? pt(args[4], args[5]) : pt(args[2], args[3]);
        sub.d += ` C${f(c1[0])} ${f(c1[1])} ${f(c2[0])} ${f(c2[1])} ${f(end[0])} ${f(end[1])}`;
        sub.straight = false;
        lastCtrl = c2;
        cur = end;
      } else if (lower === "q" || lower === "t") {
        const c = lower === "q" ? pt(args[0], args[1]) : lastCtrl && /[qt]/i.test(lastCmd) ? [2 * cur[0] - lastCtrl[0], 2 * cur[1] - lastCtrl[1]] : [...cur];
        const end = lower === "q" ? pt(args[2], args[3]) : pt(args[0], args[1]);
        sub.d += ` Q${f(c[0])} ${f(c[1])} ${f(end[0])} ${f(end[1])}`;
        sub.straight = false;
        lastCtrl = c;
        cur = end;
      } else if (lower === "a") {
        const end = pt(args[5], args[6]);
        sub.d += ` A${f(args[0])} ${f(args[1])} ${f(args[2])} ${args[3] ? 1 : 0} ${args[4] ? 1 : 0} ${f(end[0])} ${f(end[1])}`;
        sub.straight = false;
        lastCtrl = null;
        cur = end;
      }
    }
    lastCmd = cmd;
  }
  return subs.filter((s) => s.d.includes(" "));
}

/** Fewer points along a line that looks the same (Douglas–Peucker). */
export function simplify(pts, tolerance = 0.6) {
  if (pts.length < 3) return pts;
  const keep = new Array(pts.length).fill(false);
  keep[0] = keep[pts.length - 1] = true;
  const stack = [[0, pts.length - 1]];
  while (stack.length) {
    const [a, b] = stack.pop();
    let far = -1;
    let dist = 0;
    for (let i = a + 1; i < b; i += 1) { const d = distToSegment(pts[i], pts[a], pts[b]); if (d > dist) { dist = d; far = i; } }
    if (far >= 0 && dist > tolerance) { keep[far] = true; stack.push([a, far], [far, b]); }
  }
  return pts.filter((_, i) => keep[i]);
}

/** Points with the repeats taken out. */
const distinct = (pts) => pts.filter((p, i) => i === 0 || Math.hypot(p[0] - pts[i - 1][0], p[1] - pts[i - 1][1]) > 0.05);

// ---------------------------------------------------------------- the layout's build as animations

// The layout's way in (data-entrance) as the nearest PowerPoint entrance.
const ENTRANCE_FX = { rise: ["floatIn", "up"], fade: ["fade"], blur: ["fade"], pop: ["zoom", "in"], slide: ["flyIn", "left"], zoom: ["zoom", "in"], flip: ["swivel", "h"], wipe: ["wipe", "left"], drop: ["bounce"], spin: ["swivel", "h"], swirl: ["pinwheel"], none: ["appear"] };

/**
 * The layout's build as timeline entries on the converted groups: click by click (click, spotlight), one after
 * another as the slide arrives (cascade), all together (fade). Animations the timeline already put on the layout's
 * items ("@g1") follow their item to its group; the rest stay as they were.
 * `groups`: [{ g, target }] in the order of the items.
 */
export function buildTimeline({ build, entrance, groups, fields = {}, timeline = [] }) {
  const [fx, dir] = ENTRANCE_FX[entrance] || ENTRANCE_FX.rise;
  const out = [];
  let n = 0;
  const id = () => `cv${Date.now().toString(36).slice(-4)}${(n++).toString(36)}${Math.random().toString(36).slice(2, 4)}`;
  const ordered = [...groups].sort((a, b) => a.g - b.g);
  if (build === "click" || build === "spotlight") {
    for (const { target } of ordered) out.push({ id: id(), el: target, cls: "in", fx, ...(dir ? { dir } : {}), start: "click" });
  } else if (build === "cascade") {
    ordered.forEach(({ target }, i) => out.push({ id: id(), el: target, cls: "in", fx, ...(dir ? { dir } : {}), start: "with", delay: i * 170 }));
  } else if (build === "fade") {
    for (const { target } of ordered) out.push({ id: id(), el: target, cls: "in", fx: "fade", start: "with", delay: 0 });
  }
  const byG = new Map(groups.map(({ g, target }) => [g, target]));
  for (const entry of Array.isArray(timeline) ? timeline : []) {
    const m = /^@g(\d+)$/.exec(entry?.el || "");
    if (!m) { out.push(fields[entry?.el] ? { ...entry, el: fields[entry.el] } : entry); continue; }
    const target = byG.get(Number(m[1]));
    if (target) out.push({ ...entry, el: target, ...(entry.by ? { by: undefined } : {}) });
  }
  return out.map((entry) => Object.fromEntries(Object.entries(entry).filter(([, v]) => v !== undefined)));
}

/** The slide after conversion: 白紙 with what it keeps, its objects first and the ones placed by hand above them. */
export function convertedSlide(slide, objects, timeline) {
  const out = { type: "blank" };
  for (const key of KEPT_FIELDS) if (slide[key] !== undefined) out[key] = slide[key];
  const elements = [...objects, ...(Array.isArray(slide.elements) ? slide.elements : [])];
  if (elements.length) out.elements = elements;
  if (timeline?.length) out.timeline = timeline;
  return out;
}

// ---------------------------------------------------------------- reading a drawn slide back

export function createConverter(app) {
  const { E } = app;
  const fills = () => E.PALETTE.fill.map(([c]) => c);
  const lines = () => E.PALETTE.line.map(([c]) => c);
  const inks = () => E.PALETTE.text.map(([c]) => c);
  const highlights = () => E.PALETTE.highlight.map(([c]) => c);

  /** Why a slide cannot be converted (or null when it can). */
  function refusal(slide) {
    if (!slide) return "スライドがありません";
    if (slide.type === "blank") return "このスライドはすでに白紙（自由配置）です";
    if (KEEPS_LAYOUT.has(slide.type)) return "表紙・章扉・クロージング・全面写真・ひと言メッセージのページはそのままのレイアウトで使います（白紙のスライドを足すと自由に作れます）";
    return null;
  }

  /** The slide at `index` as objects and animations: { objects, timeline, slide } (the converted slide). */
  /**
   * `chrome`: also read what the studio draws around the layout (the title, the key message and the line above them,
   * the foot with the source and the page number) as text boxes — for a PowerPoint, where nothing redraws them.
   */
  async function convert(index, { chrome = false, any = false } = {}) {
    const deck = app.deck();
    const slide = deck?.slides[index];
    // `any`: the pages the editor keeps as layouts (a cover, a chapter…) are read too (a PowerPoint has no layouts to keep).
    const why = any && slide && slide.type !== "blank" ? null : refusal(slide);
    if (why) throw new Error(why);
    const takeover = Boolean(E.timelineTakesLayout?.(slide));
    const build = takeover ? "none" : E.BUILDS.includes(slide.animation) && slide.animation !== "auto" ? slide.animation : E.recommendedBuild(slide.type);
    const entrance = E.ENTRANCES[slide.entrance] ? slide.entrance : E.ENTRANCES[deck.motion?.entrance] ? deck.motion.entrance : slide.entrance === "none" || deck.motion?.entrance === "none" ? "none" : "rise";
    // Drawn with an animation on the first item, so every item carries its number (data-g) whatever the build.
    const probe = { ...slide, elements: undefined, timeline: [{ id: "cvprobe", el: "@g0", cls: "in", fx: "appear" }] };
    const opts = { ...app.renderOptions(), index, mode: "thumb", fit: app.fitFor?.(index) ?? undefined };
    const el = E.render(probe, opts);
    const host = document.createElement("div");
    host.className = "cv-host";
    host.setAttribute("aria-hidden", "true");
    host.style.cssText = "position:fixed;left:-30000px;top:0;width:1920px;height:1080px;overflow:hidden;pointer-events:none;contain:layout style;";
    host.append(el);
    document.body.append(host);
    try {
      await document.fonts?.ready;
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      if (!opts.fit) E.fit(el);
      materializePseudos(el);
      const reader = createReader(el, { slide, deck, opts, chrome });
      const atoms = reader.read();
      const { objects, groups, fields, counts } = finish(atoms, slide, deck, index);
      const timeline = buildTimeline({ build, entrance, groups, fields, timeline: slide.timeline });
      return { objects, timeline, fields, counts, slide: convertedSlide(slide, objects, timeline) };
    } finally {
      host.remove();
    }
  }

  /** ::before and ::after that show something become real elements, so they are measured like the rest. */
  function materializePseudos(slideEl) {
    const made = [];
    for (const el of slideEl.querySelectorAll(".hs-frame *, .hs-overlay *")) {
      if (el instanceof SVGElement) continue;
      for (const which of ["::before", "::after"]) {
        const cs = getComputedStyle(el, which);
        const content = cs.content;
        if (!content || content === "none" || content === "normal") continue;
        const text = /^"(.*)"$/s.exec(content)?.[1]?.replace(/\\"/g, '"').replace(/\\a\s?/g, "\n") ?? "";
        const shows = text.trim() || parseColor(cs.backgroundColor) || cs.backgroundImage !== "none" || ["Top", "Right", "Bottom", "Left"].some((s) => parseFloat(cs[`border${s}Width`]) > 0 && cs[`border${s}Style`] !== "none");
        if (!shows) continue;
        const span = document.createElement("span");
        span.className = "cv-pseudo";
        for (let i = 0; i < cs.length; i += 1) {
          const name = cs[i];
          if (name === "content" || name.startsWith("animation") || name.startsWith("transition")) continue;
          span.style.setProperty(name, cs.getPropertyValue(name));
        }
        span.textContent = text;
        made.push([el, which, span]);
      }
    }
    // Every pseudo-element is read before any is replaced (a new child would change :first-child and the like).
    for (const [el, which, span] of made) { if (which === "::before") el.prepend(span); else el.append(span); }
    const style = document.createElement("style");
    style.textContent = ".cv-host *::before, .cv-host *::after { content: none !important; }";
    slideEl.parentElement.append(style);
  }

  // How close (0–255 per channel) a colour must be to a theme colour to become it.
  const NEAR = 6;
  /** A colour as it was drawn (its opacity kept). */
  const exact = (c) => ({ color: hexOf({ r: Math.round(c.r), g: Math.round(c.g), b: Math.round(c.b) }), opacity: c.a >= 0.98 ? 1 : Math.round(c.a * 100) / 100 });
  /** The theme's colour when what was read is (nearly) one of them, else the colour as it was drawn. */
  const near = (c, palette, opts) => {
    if (!c) return null;
    // A see-through colour lies over whatever is under it (not only the white page): kept as it is, with its opacity.
    if (c.a < 0.98) return exact(c);
    const snap = snapToPalette(c, palette, opts);
    return snap && snap.err <= NEAR ? snap : exact(c);
  };
  /** What was read in the theme's colours where it is one of them: fills fade to their tint, lines and text likewise. */
  function fillOf(c, extraOpacity = 1) {
    const snap = near(c, fills());
    if (!snap) return null;
    const opacity = Math.round(snap.opacity * extraOpacity * 100) / 100;
    return { fill: snap.color, ...(opacity < 0.99 ? { fillOpacity: Math.max(0.02, opacity) } : {}) };
  }
  const lineOf = (c) => near(c, lines(), { fade: false })?.color ?? null;
  const inkOf = (c) => near(c, inks(), { fade: false })?.color ?? E.OBJECT_DEFAULTS.text.color;
  // A marker is see-through (the accent at a sixth): it is the colour it shows over the white page.
  const highlightOf = (c) => {
    if (!c) return null;
    const [r, g, b] = overWhite(c).map((v) => Math.round(v));
    return near({ r, g, b, a: 1 }, highlights(), { fade: false })?.color ?? null;
  };

  /** Everything the slide shows (its frame and placed media), in drawing order, as atoms { o, g, item, pair }. */
  /**
   * The title, the key message and the foot of any slide (a 白紙 page too) as text boxes, read from the slide as drawn
   * (for a PowerPoint, where nothing redraws them). [] when the slide draws none.
   */
  async function chromeObjects(index) {
    const deck = app.deck();
    const slide = deck?.slides[index];
    if (!slide) return [];
    const opts = { ...app.renderOptions(), index, mode: "thumb", fit: app.fitFor?.(index) ?? undefined };
    const el = E.render({ ...slide, elements: undefined }, opts);
    const host = document.createElement("div");
    host.setAttribute("aria-hidden", "true");
    host.style.cssText = "position:fixed;left:-30000px;top:0;width:1920px;height:1080px;overflow:hidden;pointer-events:none;contain:layout style;";
    host.append(el);
    document.body.append(host);
    try {
      await document.fonts?.ready;
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      if (!opts.fit) E.fit(el);
      const atoms = createReader(el, { slide, deck, opts, chrome: true, only: true }).read();
      // Each keeps which part it is ("@title", "@takeaway"), for the animations on those parts (a PowerPoint export).
      return atoms.map((atom) => { const o = E.normalizeObject(atom.o); if (o && atom.field) o.field = atom.field; return o; }).filter(Boolean);
    } finally {
      host.remove();
    }
  }

  function createReader(slideEl, { slide, opts, chrome = false, only = false }) {
    const origin = slideEl.getBoundingClientRect();
    const atoms = [];
    let pairs = 0;
    const rectOf = (node) => { const r = node.getBoundingClientRect(); return { x: r.left - origin.left, y: r.top - origin.top, w: r.width, h: r.height }; };
    const add = (o, ctx, pair = null) => {
      if (!o) return;
      const c = ctx.clip;
      if (c && o.kind !== "line" && !o.rot && (o.kind === "text" || (o.kind === "shape" && o.shape === "rect"))) {
        // Cut to the clipping box (a text box keeps its words where they were: only its frame shrinks).
        const x1 = Math.max(o.x, c.x1);
        const y1 = o.kind === "text" ? o.y : Math.max(o.y, c.y1);
        const x2 = Math.min(o.x + o.w, c.x2);
        const y2 = Math.min(o.y + o.h, c.y2);
        if (x2 - x1 < 1 || y2 - y1 < 1) return;
        if (o.kind === "text") { if (o.valign === "top") o.h = r2(y2 - y1); }
        else Object.assign(o, { x: r2(x1), y: r2(y1), w: r2(x2 - x1), h: r2(y2 - y1) });
      }
      atoms.push({ o, g: ctx.g, item: ctx.item, field: ctx.field, count: Boolean(ctx.count), pair });
    };
    const ctxFor = (el, ctx) => {
      const next = { ...ctx };
      if (next.g == null && el.dataset?.g != null && el.dataset.g !== "") next.g = Number(el.dataset.g);
      if (next.item == null && el.dataset?.item) next.item = el.dataset.item;
      // A statement/closing animates its large message, not its small title or eyebrow.
      if (next.field == null && el.classList?.contains("hs-statement-text")) next.field = "@kinetic";
      if (next.field == null && el.classList?.contains("hs-closing-message")) next.field = "@kinetic";
      // The title and the key message (read with the chrome): an animation on "@title" follows them.
      if (next.field == null && (el.dataset?.field === "title" || el.dataset?.field === "takeaway")) next.field = `@${el.dataset.field}`;
      // A figure the show counts up as it appears (data-count).
      if (!next.count && el.dataset?.count) next.count = true;
      return next;
    };
    const opacityOf = (a) => (a < 0.98 ? { opacity: r2(Math.max(0.05, a)) } : {});
    const escapeHtml = (text) => text.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
    const weightOf = (cs) => (Number(cs.fontWeight) >= 600 || cs.fontWeight === "bold" ? 700 : 400);
    const dashOf = (style) => (style === "dashed" ? "dash" : style === "dotted" ? "roundDot" : undefined);
    const gradientColor = (image) => (String(image).match(/rgba?\([^)]*\)|color\([^)]*\)/g) || []).map(parseColor).filter(Boolean).sort((a, b) => b.a - a.a)[0] || null;

    /**
     * Adds a filled shape; a tinted one with an outline becomes the fill and, above it, the outline on its own
     * (so each can be changed on its own), the two grouped so they move together.
     */
    function addShape(o, ctx) {
      if (!o) return;
      const tinted = o.fill && o.fill !== "none" && o.fill !== "#ffffff";
      if (!tinted || !o.stroke || o.stroke === "none") add(o, ctx);
      else {
        const pair = `p${(pairs += 1)}`;
        const { text, ...shape } = o;
        add({ ...o, stroke: "none", strokeW: undefined, dash: undefined }, ctx, pair);
        add({ ...shape, fill: "none", fillOpacity: undefined }, ctx, pair);
      }
      for (const line of o.after || []) add(line, ctx);
    }

    // ---- boxes

    function boxOf(cs) {
      const sides = ["Top", "Right", "Bottom", "Left"].map((s) => {
        const w = parseFloat(cs[`border${s}Width`]) || 0;
        const style = cs[`border${s}Style`];
        const color = parseColor(cs[`border${s}Color`]);
        return w > 0.2 && color && !["none", "hidden"].includes(style) ? { w, style, color } : null;
      });
      let fill = parseColor(cs.backgroundColor);
      // A ruled pattern or a hatch is worked out once the box's size is known (boxObject).
      const pattern = /^repeating-linear-gradient/.test(cs.backgroundImage || "") ? cs.backgroundImage : null;
      if (!fill && !pattern && cs.backgroundImage && cs.backgroundImage !== "none" && !/url\(/.test(cs.backgroundImage)) {
        // A gradient (a wash): its strongest colour.
        fill = gradientColor(cs.backgroundImage);
      }
      // A ring drawn with a shadow's spread (no offset, no blur) is an outline.
      let ring = null;
      for (const part of String(cs.boxShadow || "none").split(/,(?![^(]*\))/)) {
        if (part.trim() === "none") break;
        const color = parseColor((part.match(/rgba?\([^)]*\)|color\([^)]*\)/) || [])[0]);
        const nums = part.replace(/rgba?\([^)]*\)|color\([^)]*\)/g, "").match(/-?[\d.]+(?=px)/g)?.map(Number) || [];
        const [x = 0, y = 0, blur = 0, spread = 0] = nums;
        if (color && !x && !y && !blur && spread > 0.2) { ring = { w: spread, color, style: "solid", inset: /inset/.test(part) }; break; }
      }
      return { fill, sides, ring, pattern, any: Boolean(fill || ring || pattern || sides.some(Boolean)) };
    }

    function shapeOfBox(cs, rect) {
      const clip = cs.clipPath && cs.clipPath !== "none" ? cs.clipPath : "";
      const pts = clip.startsWith("polygon") ? parsePolygon(clip, rect.w, rect.h) : null;
      if (pts) return { shape: "custom", path: { pts: pts.map(([x, y]) => [Math.round(x * 10000) / 10000, Math.round(y * 10000) / 10000]), closed: true } };
      if (/^(circle|ellipse)\(/.test(clip)) return { shape: "ellipse" };
      const min = Math.min(rect.w, rect.h);
      const radius = (v, size) => { const first = String(v || "0").split(" ")[0]; return first.endsWith("%") ? (parseFloat(first) / 100) * size : parseFloat(first) || 0; };
      const pct = String(cs.borderTopLeftRadius).trim().endsWith("%") && parseFloat(cs.borderTopLeftRadius) >= 50;
      const radii = [cs.borderTopLeftRadius, cs.borderTopRightRadius, cs.borderBottomRightRadius, cs.borderBottomLeftRadius].map((v) => radius(v, min));
      const adj = (r) => [Math.round(clamp(r / min, 0, 0.5) * 10000) / 10000];
      if (radii.every((r) => r < 0.75)) return { shape: "rect" };
      if (pct || (Math.abs(rect.w - rect.h) < 1 && radii.every((r) => r >= min / 2 - 0.5))) return { shape: "ellipse" };
      if (radii.every((r) => Math.abs(r - radii[0]) < 0.5)) return { shape: "roundRect", adj: adj(radii[0]) };
      if (radii[0] > 0.5 && Math.abs(radii[0] - radii[1]) < 0.5 && radii[2] < 0.5 && radii[3] < 0.5) return { shape: "round2SameRect", adj: adj(radii[0]) };
      return { shape: "roundRect", adj: adj(Math.max(...radii)) };
    }

    /** Rotation of an element's own transform (degrees), 0 when it only moves or scales. */
    function rotationOf(cs) {
      const m = /^matrix\(([^)]*)\)$/.exec(cs.transform || "");
      if (!m) return 0;
      const [a, b] = m[1].split(",").map(Number);
      const deg = (Math.atan2(b, a) * 180) / Math.PI;
      return Math.abs(deg) > 0.3 ? r2(deg) : 0;
    }
    /** An element's box on the slide, upright with its rotation (getBoundingClientRect is the box around it). */
    function placeOf(el, cs) {
      const rect = rectOf(el);
      const rot = rotationOf(cs);
      if (!rot || !(el instanceof HTMLElement)) return { ...rect, rot: 0 };
      const w = el.offsetWidth;
      const hh = el.offsetHeight;
      return { x: rect.x + rect.w / 2 - w / 2, y: rect.y + rect.h / 2 - hh / 2, w, h: hh, rot };
    }

    /** Borders drawn on some sides only, as lines along those sides. */
    function sideLines(place, sides, ctx, opacity) {
      const [top, right, bottom, left] = sides;
      const { x, y, w, h: hh } = place;
      const line = (side, x1, y1, x2, y2) => {
        const color = lineOf(side.color);
        if (!color) return;
        add({ kind: "line", x1: r2(x1), y1: r2(y1), x2: r2(x2), y2: r2(y2), stroke: color, strokeW: r2(clamp(side.w, 0.5, 200)), ...(dashOf(side.style) ? { dash: dashOf(side.style) } : {}), ...opacityOf(side.color.a * opacity) }, ctx);
      };
      if (top) line(top, x, y + top.w / 2, x + w, y + top.w / 2);
      if (bottom) line(bottom, x, y + hh - bottom.w / 2, x + w, y + hh - bottom.w / 2);
      if (left) line(left, x + left.w / 2, y, x + left.w / 2, y + hh);
      if (right) line(right, x + w - right.w / 2, y, x + w - right.w / 2, y + hh);
    }

    /**
     * A box (a fill and/or an outline all round) as a shape — a thin bar without words as a line — and the borders
     * on some sides only as lines. Returns the shape (not yet added) or null.
     */
    function boxObject(el, cs, box, ctx, opacity, { words = false } = {}) {
      const place = placeOf(el, cs);
      if (place.w < 0.5 || place.h < 0.5) return null;
      const [t, r, b, l] = box.sides;
      const all = Boolean(t && r && b && l && [r, b, l].every((s) => Math.abs(s.w - t.w) < 0.5 && hexOf(s.color) === hexOf(t.color) && s.style === t.style));
      if (!all && box.sides.some(Boolean) && !place.rot) sideLines(place, box.sides, ctx, opacity);
      let fill = box.fill ? fillOf(box.fill) : null;
      const pattern = box.pattern ? patternOf(box.pattern, place.w, place.h) : null;
      if (pattern?.fill && !fill) fill = fillOf(pattern.fill);
      // Ruled lines (a Gantt chart's month lines) go above the box's own fill.
      const ruled = pattern?.lines && !place.rot ? pattern.lines.map((l) => {
        const color = snapToPalette(l.color, lines());
        return { kind: "line", x1: r2(place.x + l.x1), y1: r2(place.y + l.y1), x2: r2(place.x + l.x2), y2: r2(place.y + l.y2), stroke: color.color, strokeW: r2(clamp(l.w, 0.5, 200)), ...opacityOf(opacity * color.opacity) };
      }) : [];
      const outline = all ? t : box.ring;
      if (!fill && !outline) { for (const l of ruled) add(l, ctx); return null; }
      const outlineOf = outline ? { stroke: lineOf(outline.color) || "#d9d9d9", strokeW: r2(clamp(outline.w, 0.5, 60)), ...(dashOf(outline.style) ? { dash: dashOf(outline.style) } : {}) } : { stroke: "none" };
      const geometry = shapeOfBox(cs, place);
      // A shape's outline runs along its edge; a ring sits outside the box (or inside, inset): the shape grows to match.
      if (!all && box.ring) {
        const d = (box.ring.inset ? -1 : 1) * (box.ring.w / 2);
        Object.assign(place, { x: place.x - d, y: place.y - d, w: place.w + 2 * d, h: place.h + 2 * d });
      }
      const thin = Math.min(place.w, place.h);
      if (fill && !outline && !words && !place.rot && geometry.shape !== "custom" && thin <= 6 && Math.max(place.w, place.h) >= thin * 3) {
        // A rule drawn as a thin box.
        const across = place.w >= place.h;
        const color = lineOf(box.fill);
        add({ kind: "line", ...(across ? { x1: r2(place.x), y1: r2(place.y + place.h / 2), x2: r2(place.x + place.w), y2: r2(place.y + place.h / 2) } : { x1: r2(place.x + place.w / 2), y1: r2(place.y), x2: r2(place.x + place.w / 2), y2: r2(place.y + place.h) }),
          stroke: color, strokeW: r2(clamp(thin, 0.5, 60)), ...opacityOf(opacity * (fill.fillOpacity ?? 1) * box.fill.a) }, ctx);
        return null;
      }
      const shape = {
        kind: "shape", x: r2(place.x), y: r2(place.y), w: r2(place.w), h: r2(place.h), ...(place.rot ? { rot: place.rot } : {}),
        ...geometry, ...(fill || { fill: "none" }), ...outlineOf, ...opacityOf(opacity),
      };
      if (ruled.length) Object.defineProperty(shape, "after", { value: ruled, enumerable: false });
      return shape;
    }

    // ---- text

    const OWN_TAGS = new Set(["IMG", "VIDEO", "IFRAME", "CANVAS", "INPUT", "BUTTON", "SELECT", "TEXTAREA", "TABLE", "AUDIO", "OBJECT"]);
    /** Is a child part of its parent's line of words, or a thing of its own (a block, a pill, an icon, something placed)? */
    function flowOf(node) {
      if (node.nodeType === 3) return node.data.trim() ? "text" : "space";
      if (node.nodeType !== 1) return "skip";
      if (node.tagName === "BR") return "br";
      if (node instanceof SVGElement || OWN_TAGS.has(node.tagName)) return "own";
      const cs = getComputedStyle(node);
      if (cs.display === "none") return "skip";
      if (cs.position === "absolute" || cs.position === "fixed" || cs.float !== "none") return "own";
      if (cs.display === "contents" || cs.display === "inline") return "inline";
      return cs.display.startsWith("inline") ? "own" : "block";
    }
    /** An element whose content is one flow of words (plain inline children only): { texts, own } or null. */
    function leafOf(el) {
      const texts = [];
      const own = [];
      const scan = (parent) => {
        for (const node of parent.childNodes) {
          const kind = flowOf(node);
          if (kind === "block") return false;
          if (kind === "text") texts.push(node);
          else if (kind === "own") own.push(node);
          else if (kind === "inline" && scan(node) === false) return false;
        }
        return true;
      };
      if (scan(el) === false || !texts.length) return null;
      return { texts, own };
    }
    /** Where words stand: the box around their lines, and the height of one line's letters (to find the line boxes). */
    function textRect(nodes) {
      let box = null;
      let glyph = 0;
      for (const node of nodes) {
        const range = document.createRange();
        range.selectNodeContents(node);
        for (const r of range.getClientRects()) {
          if (r.width < 0.1 && r.height < 0.1) continue;
          glyph = Math.max(glyph, r.height);
          const rr = { x1: r.left - origin.left, y1: r.top - origin.top, x2: r.right - origin.left, y2: r.bottom - origin.top };
          box = box ? { x1: Math.min(box.x1, rr.x1), y1: Math.min(box.y1, rr.y1), x2: Math.max(box.x2, rr.x2), y2: Math.max(box.y2, rr.y2) } : rr;
        }
      }
      return box && { x: box.x1, y: box.y1, w: box.x2 - box.x1, h: box.y2 - box.y1, glyph };
    }

    /** The words under `roots` (children of `el`) as rich text, with the formatting that differs from the box's own. */
    function richOf(el, roots, own) {
      const base = getComputedStyle(el);
      const keepSpaces = /^(pre|pre-wrap|break-spaces)$/.test(base.whiteSpace);
      const lines = base.whiteSpace === "pre-line";
      const out = [];
      const walk = (node, parentCs) => {
        if (node.nodeType === 3) {
          let text = node.data;
          if (!keepSpaces) text = lines ? text.replace(/[ \t]+/g, " ") : text.replace(/\s+/g, " ");
          out.push(escapeHtml(text).replace(/\n/g, "<br>"));
          return;
        }
        if (node.nodeType !== 1 || own.has(node)) return;
        if (node.tagName === "BR") { out.push("<br>"); return; }
        if (node instanceof SVGElement || OWN_TAGS.has(node.tagName)) return;
        const cs = getComputedStyle(node);
        if (cs.display === "none") return;
        const style = [];
        const color = parseColor(cs.webkitTextFillColor) || parseColor(cs.color);
        const parentColor = parseColor(parentCs.webkitTextFillColor) || parseColor(parentCs.color);
        if (color && (!parentColor || inkOf(color) !== inkOf(parentColor))) style.push(`color: ${inkOf(color)}`);
        if (Math.abs(parseFloat(cs.fontSize) - parseFloat(parentCs.fontSize)) > 0.4) style.push(`font-size: ${r2(parseFloat(cs.fontSize))}px`);
        // A marker under the words (the layouts' emphasis) is a highlight.
        const bg = parseColor(cs.backgroundColor) || (cs.backgroundImage !== "none" ? gradientColor(cs.backgroundImage) : null);
        if (bg) { const mark = highlightOf(bg); if (mark) style.push(`background-color: ${mark}`); }
        const open = [];
        const close = [];
        const wrap = (o, c) => { open.push(o); close.unshift(c); };
        if (style.length) wrap(`<span style="${style.join("; ")}">`, "</span>");
        const w = weightOf(cs);
        if (w !== weightOf(parentCs)) { if (w === 700) wrap("<b>", "</b>"); else wrap('<span style="font-weight: 400">', "</span>"); }
        if (cs.fontStyle === "italic" && parentCs.fontStyle !== "italic") wrap("<i>", "</i>");
        const deco = cs.textDecorationLine || "";
        const parentDeco = parentCs.textDecorationLine || "";
        if (deco.includes("underline") && !parentDeco.includes("underline")) wrap("<u>", "</u>");
        if (deco.includes("line-through") && !parentDeco.includes("line-through")) wrap("<s>", "</s>");
        if (cs.verticalAlign === "super") wrap("<sup>", "</sup>");
        if (cs.verticalAlign === "sub") wrap("<sub>", "</sub>");
        out.push(...open);
        for (const child of node.childNodes) walk(child, cs);
        out.push(...close);
      };
      for (const node of roots) walk(node, base);
      let html = out.join("");
      if (!keepSpaces) html = html.replace(/^((?:<[^>]+>)*)\s+/, "$1").replace(/\s+((?:<\/[^>]+>)*)$/, "$1");
      // A list item's bullet or number, which the page drew as its marker.
      if (base.display === "list-item" && base.listStyleType && base.listStyleType !== "none") {
        const index = [...el.parentElement.children].filter((c) => getComputedStyle(c).display === "list-item").indexOf(el) + 1;
        html = `${/decimal/.test(base.listStyleType) ? `${index}. ` : "• "}${html}`;
      }
      return E.sanitizeRich(html);
    }

    /** Text settings of a box of words. */
    function textStyle(cs) {
      const fs = parseFloat(cs.fontSize) || 28;
      const lh = cs.lineHeight === "normal" ? 1.3 : parseFloat(cs.lineHeight) / fs;
      const flex = /flex|grid/.test(cs.display);
      const column = /column/.test(cs.flexDirection);
      let align = { start: "left", left: "left", center: "center", right: "right", end: "right", justify: "justify", "-webkit-center": "center" }[cs.textAlign] || "left";
      let valign = "top";
      if (flex) {
        const grid = /grid/.test(cs.display);
        const across = grid ? `${cs.justifyItems} ${cs.justifyContent}` : column ? cs.alignItems : cs.justifyContent;
        const down = grid ? `${cs.alignItems} ${cs.alignContent}` : column ? cs.justifyContent : cs.alignItems;
        if (/center/.test(across)) align = "center"; else if (/(flex-)?end|right/.test(across)) align = "right";
        if (/center/.test(down)) valign = "middle"; else if (/(flex-)?end/.test(down)) valign = "bottom";
      } else if (cs.display === "table-cell") valign = cs.verticalAlign === "middle" ? "middle" : cs.verticalAlign === "bottom" ? "bottom" : "top";
      // Letters drawn as outlines only (a transparent fill and a stroke): the stroke is their outline.
      const strokeW = parseFloat(cs.webkitTextStrokeWidth) || 0;
      const hollow = strokeW > 0 && /^(transparent|rgba\([^)]*,\s*0\))$/.test(cs.webkitTextFillColor || "");
      const color = hollow ? parseColor(cs.webkitTextStrokeColor) || parseColor(cs.color) : parseColor(cs.webkitTextFillColor) || parseColor(cs.color);
      const family = cs.fontFamily || "";
      // The face the words are set in: the theme's heading font, its body font (the default), or a face of its own.
      const first = family.split(",")[0].trim().replace(/^["']|["']$/g, "");
      const theme = E.THEMES.find((t) => t.id === app.deck()?.theme);
      const font = /^"?BIZ UDP/i.test(family) ? "ud" : /教科書|Kyokasho|Klee/i.test(family) ? "kyokasho" : first && first === theme?.officeFonts?.heading && first !== theme?.officeFonts?.body ? "heading" : undefined;
      const face = !font && first && first !== theme?.officeFonts?.body && /^[\p{L}\p{N}\s._+-]{1,60}$/u.test(first) ? { fontFace: first } : {};
      const ls = cs.letterSpacing === "normal" ? 0 : parseFloat(cs.letterSpacing) / fs;
      const deco = cs.textDecorationLine || "";
      return {
        fs: r2(fs), color: inkOf(color), lh: r2(clamp(Number.isFinite(lh) ? lh : 1.3, 0.8, 4)), align, valign,
        ...(weightOf(cs) === 700 ? { bold: true } : {}), ...(cs.fontStyle === "italic" ? { italic: true } : {}),
        ...(deco.includes("underline") ? { underline: true } : {}), ...(deco.includes("line-through") ? { strike: true } : {}),
        ...(font ? { font } : {}), ...face, ...(hollow ? { toutline: { w: r2(clamp(strokeW, 0.25, 24)), color: inkOf(color) }, hollow: true } : {}),
        ...(Math.abs(ls) > 0.004 ? { ls: r2(clamp(ls, -0.2, 1)) } : {}),
        ...(/nowrap|^pre$/.test(cs.whiteSpace) ? { wrap: false } : {}), ...(/^vertical/.test(cs.writingMode) ? { vertical: true } : {}),
      };
    }

    /**
     * A text box for words as the page set them. Its lines start where the page's lines start (the line boxes, not
     * the letters), wrapping words keep the width they wrapped in, balanced headings the width of their longest line.
     * `run`: words between blocks (no box of their own) — the box is their lines.
     */
    function textObject(el, cs, roots, texts, own, opacity, { run = false } = {}) {
      const text = richOf(el, roots, own);
      if (!text || !E.richToText(text).trim()) return null;
      const st = textStyle(cs);
      // A button centres its words up and down by itself.
      if (el.tagName === "BUTTON" && st.valign === "top") st.valign = "middle";
      const tr = textRect(texts);
      if (!tr) return null;
      const lineBox = cs.lineHeight === "normal" ? tr.glyph : st.lh * st.fs;
      const half = Math.max(0, (lineBox - tr.glyph) / 2);
      let box;
      // One line of words in a box that just fits them stays one line (a label, a chip, a button).
      const single = tr.h <= lineBox * 1.4;
      if (run) { box = { x: tr.x, y: tr.y - half, w: tr.w, h: tr.h + half * 2 }; if (single) st.wrap = false; }
      else {
        const rect = rectOf(el);
        const pad = ["Top", "Right", "Bottom", "Left"].map((s) => (parseFloat(cs[`padding${s}`]) || 0) + (parseFloat(cs[`border${s}Width`]) || 0));
        const content = { x: rect.x + pad[3], y: rect.y + pad[0], w: rect.w - pad[1] - pad[3], h: rect.h - pad[0] - pad[2] };
        const balanced = /balance|pretty/.test(cs.textWrap || cs.textWrapStyle || "");
        const fitText = balanced || content.w < tr.w - 1;
        if (single && content.w < tr.w + 8) st.wrap = false;
        const x = fitText ? (st.align === "center" ? tr.x : st.align === "right" ? tr.x : tr.x) : content.x;
        const w = fitText ? tr.w : content.w;
        if (st.valign !== "top" && content.h >= tr.h - 1) box = { x, y: content.y, w, h: content.h };
        else box = { x, y: content.y, w, h: Math.max(tr.y + tr.h + half - content.y, lineBox) };
      }
      // A little room so the last word of a line never wraps on its own.
      const grow = 3;
      const x = st.align === "center" ? box.x - grow / 2 : st.align === "right" ? box.x - grow : box.x;
      return {
        kind: "text", x: r2(x), y: r2(box.y), w: r2(Math.max(4, box.w + grow)), h: r2(Math.max(4, box.h)), text,
        fill: "none", stroke: "none", pad: [0, 0, 0, 0], autofit: "none", ...st, ...opacityOf(opacity),
      };
    }
    /** All the words in an element, a paragraph per block of words (table cells). */
    function wordsOf(el) {
      const leaf = leafOf(el);
      if (leaf) return richOf(el, [...el.childNodes], new Set(leaf.own));
      return [...el.children].map(wordsOf).filter(Boolean).join("");
    }

    // ---- pictures, charts, tables, icons

    function mediaObject(el, opacity) {
      const kind = el.dataset.kind || "image";
      const rect = rectOf(el);
      if (rect.w < 2 || rect.h < 2) return null;
      const media = E.mediaOf(slide, opts) || null;
      const shown = el.querySelector("img, video")?.getAttribute("src") || "";
      const picture = typeof slide.image === "string" && slide.image.startsWith("data:image/") ? slide.image : null;
      const src = kind === "image" ? [picture, media?.src].filter(Boolean).find((c) => E.resolveSrc(c, opts) === shown) || (/^(data:image|https?:)/.test(shown) ? shown : null) : media?.src;
      if (!src) return null;
      const mask = shapeOfBox(getComputedStyle(el), rect);
      const fit = el.classList.contains("fit-contain") ? "contain" : "cover";
      const base = { x: r2(rect.x), y: r2(rect.y), w: r2(rect.w), h: r2(rect.h), ...opacityOf(opacity) };
      if (kind === "image") return { kind: "image", ...base, src, fit, ...(mask.shape !== "rect" && mask.shape !== "custom" ? { mask: mask.shape, ...(mask.adj ? { adj: mask.adj } : {}) } : {}) };
      if (kind === "lottie") return { kind: "lottie", ...base, src, fit: "contain", autoplay: media?.autoplay !== false, loop: media?.loop !== false };
      return { kind: "video", ...base, src, fit, autoplay: media?.autoplay !== false, loop: media?.loop !== false, muted: media?.muted !== false };
    }

    function chartObject(el, opacity) {
      const spec = slide.image && typeof slide.image === "object" ? slide.image : null;
      if (!spec || !E.CHART_KINDS[spec.chartType || "bar"]) return null;
      const model = E.chartModel(spec);
      const chart = E.normalizeObject({ kind: "chart", chart: { type: model.type, labels: model.labels, series: model.series, unit: spec.data?.unit } })?.chart;
      if (!chart) return null;
      let rect = rectOf(el);
      const svg = el.querySelector(".hs-chart svg");
      const vb = svg?.viewBox?.baseVal;
      if (svg && vb?.width && vb?.height) {
        // The drawing keeps its proportions inside its holder (xMidYMid meet): the chart is the drawn part.
        const box = rectOf(svg);
        const k = Math.min(box.w / vb.width, box.h / vb.height);
        const drawn = { x: box.x + (box.w - vb.width * k) / 2, y: box.y + (box.h - vb.height * k) / 2, w: vb.width * k, h: vb.height * k };
        const legend = el.querySelector(".hs-legend");
        const top = legend ? Math.min(rectOf(legend).y, drawn.y) : drawn.y;
        rect = { x: drawn.x, y: top, w: drawn.w, h: drawn.y + drawn.h - top };
      }
      // A chart of one series can colour its bars one by one (the one that matters in the accent): their colours as drawn.
      const bars = model.series.length === 1 && svg ? [...svg.querySelectorAll(".hs-bar")] : [];
      const paints = bars.length === model.labels.length ? bars.map((bar) => { const c = parseColor(getComputedStyle(bar).fill); return c ? hexOf({ r: Math.round(c.r), g: Math.round(c.g), b: Math.round(c.b) }) : null; }) : [];
      const pointColors = paints.length && paints.every(Boolean) && new Set(paints).size > 1 ? { pointColors: paints } : {};
      return { kind: "chart", x: r2(rect.x), y: r2(rect.y), w: r2(rect.w), h: r2(rect.h), chart, ...pointColors, ...opacityOf(opacity) };
    }

    function tableObject(table, opacity) {
      const rect = rectOf(table);
      const rows = [...table.rows];
      if (!rows.length || rect.w < 4) return null;
      const grid = [];
      rows.forEach((tr, r) => {
        grid[r] = grid[r] || [];
        let c = 0;
        for (const cell of tr.cells) {
          while (grid[r][c]) c += 1;
          const rs = Math.max(1, cell.rowSpan || 1);
          const span = Math.max(1, cell.colSpan || 1);
          const style = getComputedStyle(cell);
          const html = wordsOf(cell);
          const bg = parseColor(style.backgroundColor);
          const st = textStyle(style);
          // Colours are written into every cell, so the table style's own header colours do not replace them.
          const entry = {
            ...(html ? { text: html } : {}), ...(bg ? { fill: fillOf(bg)?.fill } : {}), color: st.color,
            ...(st.bold ? { bold: true } : {}), ...(st.align !== "left" ? { align: st.align } : {}), ...(st.valign !== "top" ? { valign: st.valign } : {}),
            ...(rs > 1 ? { rs } : {}), ...(span > 1 ? { cs: span } : {}),
          };
          for (let dr = 0; dr < rs; dr += 1) for (let dc = 0; dc < span; dc += 1) {
            grid[r + dr] = grid[r + dr] || [];
            grid[r + dr][c + dc] = dr || dc ? { merged: true } : entry;
          }
          c += span;
        }
      });
      const nCols = Math.max(...grid.map((row) => row.length));
      const cells = grid.map((row) => Array.from({ length: nCols }, (_, c) => row[c] || {}));
      // Column widths from cells that span one column; row heights from the rows.
      const cols = Array.from({ length: nCols }, (_, c) => {
        for (const tr of rows) {
          let at = 0;
          for (const cell of tr.cells) { if (at === c && (cell.colSpan || 1) === 1) return cell.getBoundingClientRect().width; at += cell.colSpan || 1; }
        }
        return rect.w / nCols;
      });
      const first = table.querySelector("td, th");
      const fs = first ? parseFloat(getComputedStyle(first).fontSize) : 28;
      const header = [...rows[0].cells].every((cell) => cell.tagName === "TH");
      return {
        kind: "table", x: r2(rect.x), y: r2(rect.y), w: r2(rect.w), h: r2(rect.h), cells, cols, rows: rows.map((tr) => tr.getBoundingClientRect().height),
        style: "rows", header, banded: false, fs: r2(fs), ...opacityOf(opacity),
      };
    }

    function iconObject(svg, opacity) {
      const name = svg.dataset.icon;
      if (!name || !E.icons[name]) return null;
      const rect = rectOf(svg);
      if (rect.w < 1 || rect.h < 1) return null;
      const cs = getComputedStyle(svg);
      const color = lineOf(parseColor(cs.color) || parseColor(cs.stroke)) || E.OBJECT_DEFAULTS.line.stroke;
      const strokeW = parseFloat(svg.getAttribute("stroke-width") || cs.strokeWidth) || 1.75;
      return { kind: "icon", icon: name, x: r2(rect.x), y: r2(rect.y), w: r2(rect.w), h: r2(rect.h), color, strokeW: r2(clamp(strokeW, 0.5, 6)), ...opacityOf(opacity) };
    }

    // ---- SVG drawings

    function svgObjects(svg, ctx, opacity) {
      const svgRect = svg.getBoundingClientRect();
      if (svgRect.width < 1 || svgRect.height < 1) return;
      const scratch = document.createElementNS(SVG_NS, "path");
      svg.append(scratch);
      const toSlide = (m, [x, y]) => [m.a * x + m.c * y + m.e - origin.left, m.b * x + m.d * y + m.f - origin.top];
      const paint = (value) => {
        const ref = /url\("?#([^")]+)"?\)/.exec(value || "");
        if (ref) {
          const stop = svg.ownerDocument.getElementById(ref[1])?.querySelector("stop");
          return stop ? parseColor(getComputedStyle(stop).stopColor) : null;
        }
        return parseColor(value);
      };
      const arrowsOf = (cs) => ({ ...(cs.markerEnd && cs.markerEnd !== "none" ? { tail: "triangle" } : {}), ...(cs.markerStart && cs.markerStart !== "none" ? { head: "triangle" } : {}) });
      for (const el of svg.querySelectorAll("*")) {
        if (el === scratch || el.closest("defs, clipPath, mask, marker, pattern, symbol, title, desc")) continue;
        const isText = el.tagName === "text";
        if (!isText && !(el instanceof SVGGeometryElement)) continue;
        const cs = getComputedStyle(el);
        if (cs.display === "none" || cs.visibility === "hidden") continue;
        let alpha = opacity;
        for (let p = el; p && p !== svg; p = p.parentElement) alpha *= Number(getComputedStyle(p).opacity);
        if (alpha < 0.02) continue;
        const holder = el.closest("[data-g], [data-item]");
        const local = ctxFor(holder && svg.contains(holder) ? holder : svg, ctx);
        const m = el.getScreenCTM();
        if (!m) continue;
        const scale = Math.sqrt(Math.abs(m.a * m.d - m.b * m.c)) || 1;
        if (isText) { svgText(el, cs, m, local, alpha); continue; }
        const fillC = paint(cs.fill);
        const fillA = fillC ? { ...fillC, a: fillC.a * Number(cs.fillOpacity || 1) } : null;
        const strokeC = paint(cs.stroke);
        const strokeW = strokeC ? (parseFloat(cs.strokeWidth) || 1) * (cs.vectorEffect === "non-scaling-stroke" ? 1 : scale) : 0;
        const strokeShows = Boolean(strokeC && strokeW > 0.1 && Number(cs.strokeOpacity || 1) > 0.02);
        // pathLength="1" marks the strokes that draw themselves: their dashes are the drawing, not a dashed line.
        const dash = cs.strokeDasharray && cs.strokeDasharray !== "none" && !el.hasAttribute("pathLength") ? "dash" : undefined;
        const stroke = strokeShows ? { stroke: lineOf(strokeC) || E.OBJECT_DEFAULTS.line.stroke, strokeW: r2(clamp(strokeW, 0.5, 200)), ...(dash ? { dash } : {}) } : { stroke: "none" };
        const strokeAlpha = strokeShows ? strokeC.a * Number(cs.strokeOpacity || 1) : 1;
        // A line, or an outline with no fill, can fade: its tint comes back as a palette colour with an opacity.
        const faded = strokeShows ? snapToPalette({ ...strokeC, a: strokeAlpha }, lines()) : null;
        const lone = strokeShows ? { ...stroke, stroke: faded.color } : stroke;
        const upright = Math.abs(m.b) < 1e-6 && Math.abs(m.c) < 1e-6;
        const tag = el.tagName;
        // A circle stroked as one dash (a gauge's track and value): PowerPoint's 円弧, from where the dash starts
        // (clockwise from three o'clock, as the circle is turned) for as far as it runs.
        if (tag === "circle" && strokeShows && cs.strokeDasharray && cs.strokeDasharray !== "none") {
          const dashes = cs.strokeDasharray.split(/[\s,]+/).map(parseFloat).filter(Number.isFinite);
          const r = el.r.baseVal.value;
          const circ = 2 * Math.PI * r;
          if (dashes.length === 2 && dashes[0] > 0 && dashes[0] < circ * 0.999 && dashes[1] >= circ - dashes[0] - 1) {
            const offset = parseFloat(cs.strokeDashoffset) || 0;
            const turn = (Math.atan2(m.b, m.a) * 180) / Math.PI;
            const start = (((turn - (offset / circ) * 360) % 360) + 360) % 360;
            const end = (start + (dashes[0] / circ) * 360) % 360;
            const [cx, cy] = toSlide(m, [el.cx.baseVal.value, el.cy.baseVal.value]);
            const R = r * scale;
            const cap = cs.strokeLinecap === "round" ? { cap: "round" } : {};
            add({ kind: "shape", shape: "arc", adj: [Math.round(start * 100) / 100, Math.round(end * 100) / 100], x: r2(cx - R), y: r2(cy - R), w: r2(2 * R), h: r2(2 * R), fill: "none", ...lone, ...cap, dash: undefined, ...opacityOf(strokeAlpha < 0.99 ? strokeAlpha : 1) }, local);
            continue;
          }
        }
        if (tag === "line") {
          if (!strokeShows) continue;
          const p1 = toSlide(m, [el.x1.baseVal.value, el.y1.baseVal.value]);
          const p2 = toSlide(m, [el.x2.baseVal.value, el.y2.baseVal.value]);
          add({ kind: "line", x1: r2(p1[0]), y1: r2(p1[1]), x2: r2(p2[0]), y2: r2(p2[1]), ...lone, ...arrowsOf(cs), ...opacityOf(alpha * faded.opacity) }, local);
          continue;
        }
        const filled = fillA ? fillOf(fillA) : null;
        if (!filled && !strokeShows) continue;
        if (upright && (tag === "rect" || tag === "circle" || tag === "ellipse")) {
          let x; let y; let w; let hh; let shape = { shape: "ellipse" };
          if (tag === "rect") {
            [x, y] = toSlide(m, [el.x.baseVal.value, el.y.baseVal.value]);
            w = el.width.baseVal.value * m.a;
            hh = el.height.baseVal.value * m.d;
            const rx = (el.rx.baseVal.value || el.ry.baseVal.value) * Math.abs(m.a);
            shape = rx > 0.5 ? { shape: "roundRect", adj: [Math.round(clamp(rx / Math.min(Math.abs(w), Math.abs(hh)), 0, 0.5) * 10000) / 10000] } : { shape: "rect" };
          } else {
            const rx = tag === "circle" ? el.r.baseVal.value : el.rx.baseVal.value;
            const ry = tag === "circle" ? el.r.baseVal.value : el.ry.baseVal.value;
            [x, y] = toSlide(m, [el.cx.baseVal.value - rx, el.cy.baseVal.value - ry]);
            w = 2 * rx * m.a;
            hh = 2 * ry * m.d;
          }
          if (w < 0) { x += w; w = -w; }
          if (hh < 0) { y += hh; hh = -hh; }
          if (w < 0.5 || hh < 0.5) continue;
          addShape({ kind: "shape", x: r2(x), y: r2(y), w: r2(w), h: r2(hh), ...shape, ...(filled || { fill: "none" }), ...stroke, ...opacityOf(alpha) }, local);
          continue;
        }
        // Anything else: its outline as points (corners exactly; curves measured along the way).
        let subs;
        if (tag === "path") subs = pathSubpaths(el.getAttribute("d"));
        else if (tag === "polyline" || tag === "polygon") {
          const pts = [...el.points].map((p) => [p.x, p.y]);
          subs = pts.length ? [{ pts, straight: true, closed: tag === "polygon" }] : [];
        } else subs = [{ el, straight: false, closed: true }];
        for (const sub of subs) {
          let pts;
          if (sub.straight) pts = sub.pts.map((p) => toSlide(m, p));
          else {
            const target = sub.el || scratch;
            if (!sub.el) scratch.setAttribute("d", sub.d);
            const total = target.getTotalLength();
            const n = clamp(Math.ceil(total / Math.max(0.25, 2.5 / scale)), 8, 1500);
            pts = simplify(Array.from({ length: n + 1 }, (_, i) => { const p = target.getPointAtLength((total * i) / n); return toSlide(m, [p.x, p.y]); }), 0.6);
          }
          pts = distinct(pts);
          if (sub.closed && pts.length > 2 && Math.hypot(pts[0][0] - pts.at(-1)[0], pts[0][1] - pts.at(-1)[1]) < 0.5) pts.pop();
          if (pts.length < 2) continue;
          const closed = sub.closed || (Boolean(filled) && !strokeShows);
          if (pts.length === 2 || (!closed && pts.every((p) => distToSegment(p, pts[0], pts.at(-1)) < 0.5))) {
            const [a, b] = [pts[0], pts.at(-1)];
            add({ kind: "line", x1: r2(a[0]), y1: r2(a[1]), x2: r2(b[0]), y2: r2(b[1]), ...(strokeShows ? lone : { stroke: lineOf(fillA) || E.OBJECT_DEFAULTS.line.stroke, strokeW: 1 }), ...arrowsOf(cs), ...opacityOf(alpha * (strokeShows ? faded.opacity : 1)) }, local);
            continue;
          }
          let tolerance = 0.6;
          while (pts.length > 380 && tolerance < 12) { tolerance *= 1.6; pts = simplify(pts, tolerance); }
          const withFill = closed && filled;
          addShape(customShape(pts, closed, { ...(withFill ? filled : { fill: "none" }), ...(withFill ? stroke : lone) }, alpha * (withFill || !strokeShows ? 1 : faded.opacity)), local);
        }
      }
      scratch.remove();

      function svgText(el, cs, m, local, alpha) {
        const spans = [...el.querySelectorAll("tspan")];
        const content = spans.length > 1 ? spans.map((t) => t.textContent.trim()).filter(Boolean) : [el.textContent.trim()];
        if (!content.join("").trim()) return;
        let bb;
        try { bb = el.getBBox(); } catch { return; }
        const [p1, p2] = [toSlide(m, [bb.x, bb.y]), toSlide(m, [bb.x + bb.width, bb.y + bb.height])];
        const x = Math.min(p1[0], p2[0]);
        const y = Math.min(p1[1], p2[1]);
        const w = Math.abs(p2[0] - p1[0]);
        const hh = Math.abs(p2[1] - p1[1]);
        const fs = (parseFloat(cs.fontSize) || 16) * Math.abs(m.d);
        const align = { middle: "center", end: "right" }[cs.textAnchor] || "left";
        const grow = Math.max(6, fs * 0.2);
        add({
          kind: "text", x: r2(align === "center" ? x - grow / 2 : align === "right" ? x - grow : x), y: r2(y), w: r2(w + grow), h: r2(hh),
          text: E.sanitizeRich(content.map(escapeHtml).join("<br>")), fill: "none", stroke: "none", pad: [0, 0, 0, 0], autofit: "none", wrap: false,
          fs: r2(fs), lh: r2(clamp(hh / content.length / fs, 0.8, 4)), color: inkOf(paint(cs.fill) || { r: 26, g: 26, b: 26, a: 1 }), align, valign: "top",
          ...(weightOf(cs) === 700 ? { bold: true } : {}), ...opacityOf(alpha),
        }, local);
      }
    }

    /** A drawn outline (slide points) as a shape whose points are fractions of its box. */
    function customShape(pts, closed, paint, alpha) {
      const xs = pts.map((p) => p[0]);
      const ys = pts.map((p) => p[1]);
      const x = Math.min(...xs);
      const y = Math.min(...ys);
      const w = Math.max(1, Math.max(...xs) - x);
      const hh = Math.max(1, Math.max(...ys) - y);
      const frac = (v) => Math.round(v * 10000) / 10000;
      return { kind: "shape", shape: "custom", x: r2(x), y: r2(y), w: r2(w), h: r2(hh), path: { pts: pts.map(([px, py]) => [frac((px - x) / w), frac((py - y) / hh)]), ...(closed ? { closed: true } : {}) }, ...paint, ...opacityOf(alpha) };
    }

    /** A slider (試算) as its track and knob. */
    function rangeObjects(el, ctx, opacity) {
      const rect = rectOf(el);
      const min = Number(el.min || 0);
      const max = Number(el.max || 100);
      const at = max > min ? (Number(el.value) - min) / (max - min) : 0;
      const trackH = Math.max(6, Math.min(14, rect.h / 3));
      const knob = Math.min(rect.h, 40);
      const cx = rect.x + knob / 2 + at * (rect.w - knob);
      const cy = rect.y + rect.h / 2;
      add({ kind: "shape", shape: "roundRect", adj: [0.5], x: r2(rect.x), y: r2(cy - trackH / 2), w: r2(rect.w), h: r2(trackH), fill: "#f2f2f2", stroke: "none", ...opacityOf(opacity) }, ctx);
      if (cx - rect.x > trackH) add({ kind: "shape", shape: "roundRect", adj: [0.5], x: r2(rect.x), y: r2(cy - trackH / 2), w: r2(cx - rect.x), h: r2(trackH), fill: E.OBJECT_DEFAULTS.shape.fill, stroke: "none", ...opacityOf(opacity) }, ctx);
      add({ kind: "shape", shape: "ellipse", x: r2(cx - knob / 2 + 3), y: r2(cy - knob / 2 + 3), w: r2(knob - 6), h: r2(knob - 6), fill: "#ffffff", stroke: E.OBJECT_DEFAULTS.shape.fill, strokeW: 6, ...opacityOf(opacity) }, ctx);
    }

    // ---- the walk

    function visit(el, ctx) {
      if (el.nodeType !== 1 || el.matches(chrome ? ".hs-detail-badge, .hs-drill-badge, script, style, template" : ".hs-head, .hs-detail-badge, .hs-drill-badge, .hs-takeaway-row, script, style, template")) return;
      const cs = getComputedStyle(el);
      if (cs.display === "none") return;
      const opacity = ctx.opacity * Number(cs.opacity || 1);
      if (opacity < 0.02) return;
      const here = { ...ctxFor(el, ctx), opacity };
      // A box that clips what overflows it (a card, a track): text boxes and plain boxes inside stay within it.
      if (/hidden|clip/.test(`${cs.overflowX} ${cs.overflowY}`) && !(el instanceof SVGElement)) {
        const r = rectOf(el);
        const c = ctx.clip;
        here.clip = c ? { x1: Math.max(c.x1, r.x), y1: Math.max(c.y1, r.y), x2: Math.min(c.x2, r.x + r.w), y2: Math.min(c.y2, r.y + r.h) } : { x1: r.x, y1: r.y, x2: r.x + r.w, y2: r.y + r.h };
      }
      if (el instanceof SVGSVGElement) {
        if (el.classList.contains("hs-icon")) add(iconObject(el, opacity), here);
        else svgObjects(el, here, opacity);
        return;
      }
      if (el instanceof SVGElement) return;
      if (el.classList.contains("hs-chart-wrap")) {
        const chart = chartObject(el, opacity);
        if (chart) { add(chart, here); return; }
      }
      if (el.classList.contains("hs-media") && el.dataset.kind) {
        const media = mediaObject(el, opacity);
        if (media) { add(media, here); return; }
      }
      if (el.tagName === "TABLE") { add(tableObject(el, opacity), here); return; }
      if (el.tagName === "IMG") {
        const src = el.getAttribute("src") || "";
        const rect = rectOf(el);
        if (/^(data:image|https?:)/.test(src) && rect.w > 2) add({ kind: "image", x: r2(rect.x), y: r2(rect.y), w: r2(rect.w), h: r2(rect.h), src, fit: cs.objectFit === "contain" ? "contain" : cs.objectFit === "cover" ? "cover" : "fill", ...opacityOf(opacity) }, here);
        return;
      }
      if (el.tagName === "INPUT" && el.type === "range") { rangeObjects(el, here, opacity); return; }
      const hidden = cs.visibility === "hidden";
      const box = hidden ? { any: false, sides: [] } : boxOf(cs);
      let leaf = hidden ? null : leafOf(el);
      // A line of words with a chip, an icon or a picture in it: each run of words is its own box (below), so the
      // words after the chip stay after it.
      if (leaf && leaf.own.some((node) => node.nodeType === 1 && !/absolute|fixed/.test(getComputedStyle(node).position))) leaf = null;
      if (leaf) {
        const own = new Set(leaf.own);
        const filled = Boolean(box.fill || box.ring) || box.sides.every(Boolean);
        // Words in a box of their own: one shape holding them, as PowerPoint makes it.
        const shape = box.any && filled && !rotationOf(cs) ? boxObject(el, cs, box, here, opacity, { words: true }) : null;
        if (shape) {
          const words = textObject(el, cs, [...el.childNodes], leaf.texts, own, 1);
          if (words) {
            const st = textStyle(cs);
            // The words sit in the element's content box, wherever the shape's edge ended up.
            const rect = rectOf(el);
            const inset = ["Top", "Right", "Bottom", "Left"].map((s) => (parseFloat(cs[`padding${s}`]) || 0) + (parseFloat(cs[`border${s}Width`]) || 0));
            // Across, the words keep the width the text box found for them (their wrapping, a little slack).
            const pad = [rect.y + inset[0] - shape.y, shape.x + shape.w - (words.x + words.w), shape.y + shape.h - (rect.y + rect.h - inset[2]), words.x - shape.x].map((v) => r2(Math.max(0, v)));
            const { kind, x, y, w, h, fill, stroke, pad: _pad, autofit, opacity: _o, ...rest } = words;
            Object.assign(shape, rest, { pad, autofit: "none", valign: el.tagName === "BUTTON" && st.valign === "top" ? "middle" : st.valign });
            // The shape's own text area is inset (rounded corners, an ellipse): the padding makes up for it, so
            // the words have the room they had on the page.
            const area = E.geometry(shape.shape, shape.w, shape.h, shape.adj, shape.path).text;
            const insets = [area[1], shape.w - area[2], shape.h - area[3], area[0]];
            shape.pad = shape.pad.map((v, i) => r2(Math.max(0, v - insets[i])));
          }
          addShape(shape, here);
        } else {
          if (box.any) addShape(boxObject(el, cs, box, here, opacity), here);
          add(textObject(el, cs, [...el.childNodes], leaf.texts, own, opacity), here);
        }
        for (const node of leaf.own) visit(node, here);
        return;
      }
      if (box.any) addShape(boxObject(el, cs, box, here, opacity), here);
      // Mixed content: the words between blocks are text boxes of their own.
      let run = [];
      const later = [];
      const flush = () => {
        const texts = [];
        const own = [];
        const roots = [];
        for (const node of run) {
          if (node.nodeType === 3 || node.tagName === "BR") { roots.push(node); if (node.nodeType === 3 && node.data.trim()) texts.push(node); continue; }
          const inner = leafOf(node);
          if (!inner) { later.push(node); continue; }
          roots.push(node);
          texts.push(...inner.texts);
          own.push(...inner.own);
        }
        if (texts.length && !hidden) add(textObject(el, cs, roots, texts, new Set(own), opacity, { run: true }), here);
        for (const node of [...own, ...later.splice(0)]) visit(node, here);
        run = [];
      };
      for (const node of [...el.childNodes]) {
        const kind = flowOf(node);
        if (kind === "text" || kind === "inline" || kind === "br" || (kind === "space" && run.length)) { run.push(node); continue; }
        if (kind === "space" || kind === "skip") continue;
        flush();
        visit(node, here);
      }
      flush();
    }

    function read() {
      const base = { g: null, item: null, opacity: 1 };
      if (only) {
        // The chrome alone: the header inside the frame, the foot and a vertical label on the slide.
        for (const node of slideEl.querySelectorAll(":scope > .hs-frame > .hs-head, :scope > .hs-foot, :scope > .hs-vlabel")) visit(node, base);
        return atoms;
      }
      for (const part of slideEl.querySelectorAll(":scope > .hs-frame, :scope > .hs-overlay")) for (const child of part.children) visit(child, base);
      if (chrome) for (const node of slideEl.querySelectorAll(":scope > .hs-foot, :scope > .hs-vlabel")) visit(node, base);
      return atoms;
    }
    return { read };
  }

  /** Clean objects, their groups (one per item), the links details and deep-dive pages need, and the groups to animate. */
  function finish(atoms, slide, deck, index) {
    const story = E.storyMap(deck.slides);
    const targets = new Set([...(Array.isArray(slide.details) ? slide.details.map((d) => d?.target) : []), ...(story.drills[index] || []).map((d) => d.target)].filter(Boolean));
    const objects = [];
    const byG = new Map();
    const pairs = new Map();
    const seed = Math.random().toString(36).slice(2, 6);
    const byItem = new Map();
    const byField = new Map();
    const counts = [];
    for (const atom of atoms) {
      const o = E.normalizeObject(atom.o);
      if (!o) continue;
      // What only a PowerPoint export uses (the slide's object normalizer does not keep it).
      if (atom.o.pointColors) o.pointColors = atom.o.pointColors;
      if (atom.o.hollow) o.hollow = true;
      objects.push(o);
      if (atom.item && targets.has(atom.item) && o.kind !== "line") { if (!byItem.has(atom.item)) byItem.set(atom.item, []); byItem.get(atom.item).push(o); }
      if (atom.field) { if (!byField.has(atom.field)) byField.set(atom.field, []); byField.get(atom.field).push(o); }
      if (atom.count && (o.kind === "text" || o.kind === "shape") && o.text) counts.push(o.id);
      if (atom.g != null) { if (!byG.has(atom.g)) byG.set(atom.g, []); byG.get(atom.g).push(o); }
      else if (atom.pair) { if (!pairs.has(atom.pair)) pairs.set(atom.pair, []); pairs.get(atom.pair).push(o); }
      if (objects.length >= 300) break;
    }
    const groups = [];
    for (const [g, list] of [...byG].sort((a, b) => a[0] - b[0])) {
      if (list.length > 1) {
        const id = `cv${seed}${g}`;
        for (const o of list) o.group = id;
        groups.push({ g, target: `grp:${id}` });
      } else groups.push({ g, target: list[0].id });
    }
    // An item with a "詳しく" card or a deep-dive page opens it from its largest piece (the others let clicks through).
    for (const [item, list] of byItem) list.reduce((a, b) => (b.w * b.h > a.w * a.h ? b : a)).item = item;
    // A tinted box and its outline (drawn apart, see addShape) move together.
    for (const [pair, list] of pairs) if (list.length > 1) for (const o of list) o.group = `cv${seed}${pair}`;
    // The title and the key message: one text box (or a group of its pieces) each, for "@title" and "@takeaway".
    const fields = {};
    for (const [field, list] of byField) {
      if (list.length === 1) { fields[field] = list[0].id; continue; }
      const id = `cv${seed}${field.slice(1)}`;
      for (const o of list) if (!o.group) o.group = id;
      fields[field] = list.every((o) => o.group === id) ? `grp:${id}` : list[0].id;
    }
    return { objects, groups, fields, counts };
  }

  return { convert, refusal, chromeObjects };
}
