// Operations on a slide's objects (slide.elements), PowerPoint style: new objects, moving, resizing a rotated box,
// snapping to guides, aligning, distributing, grouping, ordering, duplicating, and formatting the text inside.
// Plain data in, plain data out (rich text helpers use the page's DOM through the engine).

import { deleteRefs, insertRefs, mapFormulas } from "./sheet.mjs";

export const W = 1920;
export const H = 1080;
export const PX_PER_CM = 144 / 2.54;
export const PX_PER_PT = 2;

const clone = (value) => JSON.parse(JSON.stringify(value));
const rad = (deg) => (deg * Math.PI) / 180;
const round2 = (n) => Math.round(n * 100) / 100;

/** Guides on a page: the themes' content area (120 px in from the sides, 96 px from the top, 104 px from the foot) and the middle. */
export const SLIDE_GUIDES = { x: [120, 960, 1800], y: [96, 540, 976] };

// ---------------------------------------------------------------- new objects

// What a click (without dragging) inserts, by shape: PowerPoint-like proportions in slide pixels.
const TALL = new Set(["upArrow", "downArrow", "upDownArrow", "leftBracket", "rightBracket", "leftBrace", "rightBrace", "downArrowCallout"]);
const WIDE = new Set(["rightArrow", "leftArrow", "leftRightArrow", "homePlate", "chevron", "notchedRightArrow", "stripedRightArrow", "rightArrowCallout", "wave", "doubleWave", "flowTerminator", "flowPunchedTape", "flowDocument", "flowMultidocument", "flowProcess", "flowAlternate", "flowData", "flowPredefined", "flowPreparation", "flowManualInput", "flowManualOperation", "flowStoredData", "flowDisplay", "flowDirectAccess", "rect", "roundRect", "snip1Rect", "snip2SameRect", "snip2DiagRect", "round1Rect", "round2SameRect", "round2DiagRect", "parallelogram", "trapezoid", "wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout", "lineCallout", "bracketPair", "bracePair", "plaque", "foldedCorner", "bevel"]);
export function defaultSize(shape) {
  if (TALL.has(shape)) return [144, 288];
  if (WIDE.has(shape)) return [360, 200];
  return [240, 240];
}

let seed = 0;
export function newId() {
  seed += 1;
  return `o${Date.now().toString(36).slice(-5)}${seed.toString(36)}${Math.random().toString(36).slice(2, 5)}`;
}

// 既定の図形に設定・既定のテキスト ボックスに設定・既定の線に設定 (kept on the deck as deck.objectDefaults): which
// parts of an object's look a new shape, text box or line starts with — never its words, place or size.
export const DEFAULT_STYLE_KEYS = {
  shape: ["fill", "gradient", "pattern", "fillOpacity", "stroke", "strokeW", "dash", "fs", "color", "bold", "italic", "underline", "uline", "strike", "sline", "caps", "tshadow", "tglow", "toutline", "align", "valign", "font", "lh", "ls", "psp", "pad", "autofit", "opacity"],
  text: ["fill", "gradient", "pattern", "fillOpacity", "stroke", "strokeW", "dash", "fs", "color", "bold", "italic", "underline", "uline", "strike", "sline", "caps", "tshadow", "tglow", "toutline", "align", "valign", "font", "lh", "ls", "psp", "pad", "autofit", "wrap", "opacity"],
  line: ["stroke", "strokeW", "dash", "head", "tail", "headSize", "tailSize", "route", "opacity"],
};

/** The look of an object (with its defaults filled in) to keep as the default for `kind`. */
export function defaultStyleOf(kind, o) {
  const keys = DEFAULT_STYLE_KEYS[kind];
  if (!keys || !o) return null;
  const out = {};
  for (const key of keys) if (o[key] !== undefined && o[key] !== null) out[key] = clone(o[key]);
  return Object.keys(out).length ? out : null;
}

/** deck.objectDefaults made safe: only the three kinds, only their look keys, each value passed by `check`
 *  (kind, style) → style (the studio runs it through the object normalizer). null when empty. */
export function objectDefaultsOf(value, check = (kind, style) => style) {
  if (!value || typeof value !== "object" || Array.isArray(value)) return null;
  const out = {};
  for (const kind of Object.keys(DEFAULT_STYLE_KEYS)) {
    const raw = value[kind];
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) continue;
    const picked = defaultStyleOf(kind, raw);
    const checked = picked && check(kind, picked);
    const kept = checked && defaultStyleOf(kind, checked);
    if (kept) out[kind] = kept;
  }
  return Object.keys(out).length ? out : null;
}

/** A new object of a kind at a box (slide pixels); lines take x1,y1,x2,y2. */
export function makeObject(kind, box, extra = {}) {
  const base = { id: newId(), kind };
  if (kind === "line") return { ...base, x1: round2(box.x1), y1: round2(box.y1), x2: round2(box.x2), y2: round2(box.y2), ...extra };
  return { ...base, x: round2(box.x), y: round2(box.y), w: round2(Math.max(4, box.w)), h: round2(Math.max(4, box.h)), ...extra };
}

// ---------------------------------------------------------------- geometry

export function corners(o) {
  if (o.kind === "line") return [[o.x1, o.y1], [o.x2, o.y2]];
  const cx = o.x + o.w / 2;
  const cy = o.y + o.h / 2;
  const a = rad(o.rot || 0);
  const [c, s] = [Math.cos(a), Math.sin(a)];
  return [[-1, -1], [1, -1], [1, 1], [-1, 1]].map(([kx, ky]) => {
    const dx = (kx * o.w) / 2;
    const dy = (ky * o.h) / 2;
    return [cx + dx * c - dy * s, cy + dx * s + dy * c];
  });
}

/** The upright box around objects (lines attached to others are measured where they are drawn). */
export function bounds(objects, ends = null) {
  const pts = objects.flatMap((o) => (o.kind === "line" && ends ? ends(o) : corners(o)));
  if (!pts.length) return null;
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  return { x, y, w: Math.max(...xs) - x, h: Math.max(...ys) - y };
}

/** Is the point inside the (rotated) box of an object? Lines: within `slack` of the line. */
export function hits(o, [px, py], slack = 8) {
  if (o.kind === "line") return distToSegment([px, py], [o.x1, o.y1], [o.x2, o.y2]) <= slack;
  const cx = o.x + o.w / 2;
  const cy = o.y + o.h / 2;
  const a = rad(-(o.rot || 0));
  const dx = px - cx;
  const dy = py - cy;
  const lx = dx * Math.cos(a) - dy * Math.sin(a);
  const ly = dx * Math.sin(a) + dy * Math.cos(a);
  return Math.abs(lx) <= o.w / 2 + slack && Math.abs(ly) <= o.h / 2 + slack;
}

export function distToSegment([px, py], [x1, y1], [x2, y2]) {
  const dx = x2 - x1;
  const dy = y2 - y1;
  const len = dx * dx + dy * dy;
  const t = len ? Math.max(0, Math.min(1, ((px - x1) * dx + (py - y1) * dy) / len)) : 0;
  return Math.hypot(px - (x1 + t * dx), py - (y1 + t * dy));
}

const HANDLES = { nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0], se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0] };
export const HANDLE_NAMES = Object.keys(HANDLES);

/**
 * Resize a (possibly rotated) box by dragging one of its handles to `p` (slide pixels): the opposite side stays
 * put. keepRatio (Shift, or pictures by their corners) keeps the proportions; fromCenter (Ctrl/Alt) grows both
 * ways. Dragging past the opposite side flips the object, as PowerPoint does.
 */
export function resizeBox(o, handle, [px, py], { keepRatio = false, fromCenter = false, min = 4 } = {}) {
  const [hx, hy] = HANDLES[handle];
  const a = rad(o.rot || 0);
  const [c, s] = [Math.cos(a), Math.sin(a)];
  const toWorld = ([lx, ly]) => [lx * c - ly * s, lx * s + ly * c];
  const toLocal = ([wx, wy]) => [wx * c + wy * s, -wx * s + wy * c];
  const cx = o.x + o.w / 2;
  const cy = o.y + o.h / 2;
  const anchorLocal = fromCenter ? [0, 0] : [(-hx * o.w) / 2, (-hy * o.h) / 2];
  const [ax, ay] = toWorld(anchorLocal);
  const anchor = [cx + ax, cy + ay];
  const [dx, dy] = toLocal([px - anchor[0], py - anchor[1]]);
  const k = fromCenter ? 2 : 1;
  let w = hx ? Math.abs(dx) * k : o.w;
  let h = hy ? Math.abs(dy) * k : o.h;
  let flipX = hx && dx * hx < 0;
  let flipY = hy && dy * hy < 0;
  if (keepRatio && o.w > 0 && o.h > 0) {
    const ratio = o.w / o.h;
    if (hx && hy) {
      const scale = Math.max(w / o.w, h / o.h);
      w = o.w * scale;
      h = o.h * scale;
    } else if (hx) h = w / ratio;
    else w = h * ratio;
  }
  w = Math.max(min, w);
  h = Math.max(min, h);
  // The new centre: from the anchor, half the new size towards the dragged side (in the box's own frame).
  const sideX = hx ? (flipX ? -hx : hx) : 0;
  const sideY = hy ? (flipY ? -hy : hy) : 0;
  const centerLocal = fromCenter ? [0, 0] : [hx ? (sideX * w) / 2 : 0, hy ? (sideY * h) / 2 : 0];
  const [ox, oy] = toWorld(centerLocal);
  const center = fromCenter ? [cx, cy] : [anchor[0] + ox, anchor[1] + oy];
  if (!fromCenter && !hx) { const [sx, sy] = toWorld([0, 0]); center[0] += sx; center[1] += sy; }
  const out = { x: round2(center[0] - w / 2), y: round2(center[1] - h / 2), w: round2(w), h: round2(h) };
  if (flipX) out.flipH = !o.flipH;
  if (flipY) out.flipV = !o.flipV;
  return out;
}

/** The angle (degrees) from the centre of a box to a point, with 0 straight up (the rotate handle's rest). */
export function angleTo(o, [px, py], snap = false) {
  const cx = o.x + o.w / 2;
  const cy = o.y + o.h / 2;
  let deg = (Math.atan2(py - cy, px - cx) * 180) / Math.PI + 90;
  if (snap) deg = Math.round(deg / 15) * 15;
  deg = ((deg + 180) % 360 + 360) % 360 - 180;
  return round2(deg);
}

/** Rotate a point around a centre. */
export function rotatePoint([x, y], [cx, cy], deg) {
  const a = rad(deg);
  const dx = x - cx;
  const dy = y - cy;
  return [cx + dx * Math.cos(a) - dy * Math.sin(a), cy + dx * Math.sin(a) + dy * Math.cos(a)];
}

// ---------------------------------------------------------------- snapping (smart guides)

/**
 * Snap a moving box to the edges and centres of the other objects, the slide and the guides. Returns the
 * adjustment and the guide lines to draw. `tolerance` is in slide pixels.
 */
export function snapBox(box, targets, { tolerance = 8, guides = SLIDE_GUIDES, grid = 0 } = {}) {
  const xs = [];
  const ys = [];
  const add = (list, value, from, to, kind) => list.push({ value, from, to, kind });
  for (const t of targets) {
    for (const v of [t.x, t.x + t.w / 2, t.x + t.w]) add(xs, v, t.y, t.y + t.h, "object");
    for (const v of [t.y, t.y + t.h / 2, t.y + t.h]) add(ys, v, t.x, t.x + t.w, "object");
  }
  for (const v of [0, W / 2, W]) add(xs, v, 0, H, "slide");
  for (const v of [0, H / 2, H]) add(ys, v, 0, W, "slide");
  for (const v of guides?.x ?? []) add(xs, v, 0, H, "guide");
  for (const v of guides?.y ?? []) add(ys, v, 0, W, "guide");
  const best = (edges, cands) => {
    let found = null;
    for (const edge of edges) {
      for (const cand of cands) {
        const d = cand.value - edge;
        if (Math.abs(d) <= tolerance && (!found || Math.abs(d) < Math.abs(found.d) - 0.01)) found = { d, cand, edge };
      }
    }
    return found;
  };
  const sx = best([box.x, box.x + box.w / 2, box.x + box.w], xs);
  const sy = best([box.y, box.y + box.h / 2, box.y + box.h], ys);
  let dx = sx ? sx.d : 0;
  let dy = sy ? sy.d : 0;
  if (grid > 0) {
    if (!sx) dx = Math.round(box.x / grid) * grid - box.x;
    if (!sy) dy = Math.round(box.y / grid) * grid - box.y;
  }
  const lines = [];
  const moved = { x: box.x + dx, y: box.y + dy, w: box.w, h: box.h };
  if (sx) {
    const v = sx.cand.value;
    const same = xs.filter((c) => Math.abs(c.value - v) < 0.5);
    const lo = Math.min(moved.y, ...same.map((c) => c.from));
    const hi = Math.max(moved.y + moved.h, ...same.map((c) => c.to));
    lines.push({ axis: "x", at: v, from: lo, to: hi, kind: sx.cand.kind });
  }
  if (sy) {
    const v = sy.cand.value;
    const same = ys.filter((c) => Math.abs(c.value - v) < 0.5);
    const lo = Math.min(moved.x, ...same.map((c) => c.from));
    const hi = Math.max(moved.x + moved.w, ...same.map((c) => c.to));
    lines.push({ axis: "y", at: v, from: lo, to: hi, kind: sy.cand.kind });
  }
  // Equal spacing: the gap to the neighbour on one side matches the gap on the other side.
  if (!sx) {
    const row = targets.filter((t) => t.y < moved.y + moved.h && t.y + t.h > moved.y);
    const left = row.filter((t) => t.x + t.w <= box.x + tolerance).sort((a, b) => b.x + b.w - (a.x + a.w))[0];
    const right = row.filter((t) => t.x >= box.x + box.w - tolerance).sort((a, b) => a.x - b.x)[0];
    if (left && right) {
      const target = (left.x + left.w + right.x - box.w) / 2;
      if (Math.abs(target - box.x) <= tolerance) {
        dx = target - box.x;
        const y = moved.y + moved.h / 2;
        lines.push({ axis: "gap", from: left.x + left.w, to: target, at: y }, { axis: "gap", from: target + box.w, to: right.x, at: y });
      }
    }
  }
  if (!sy) {
    const col = targets.filter((t) => t.x < moved.x + moved.w && t.x + t.w > moved.x);
    const above = col.filter((t) => t.y + t.h <= box.y + tolerance).sort((a, b) => b.y + b.h - (a.y + a.h))[0];
    const below = col.filter((t) => t.y >= box.y + box.h - tolerance).sort((a, b) => a.y - b.y)[0];
    if (above && below) {
      const target = (above.y + above.h + below.y - box.h) / 2;
      if (Math.abs(target - box.y) <= tolerance) {
        dy = target - box.y;
        const x = moved.x + moved.w / 2;
        lines.push({ axis: "vgap", from: above.y + above.h, to: target, at: x }, { axis: "vgap", from: target + box.h, to: below.y, at: x });
      }
    }
  }
  return { dx: round2(dx), dy: round2(dy), lines };
}

/** Snap one value (an edge being resized) to nearby edges. */
export function snapValue(value, cands, tolerance) {
  let best = null;
  for (const c of cands) if (Math.abs(c - value) <= tolerance && (best == null || Math.abs(c - value) < Math.abs(best - value))) best = c;
  return best;
}

// ---------------------------------------------------------------- selection and groups

/** Selecting one member of a group selects the whole group. */
export function withGroups(list, ids) {
  const set = new Set(ids);
  const groups = new Set(list.filter((o) => set.has(o.id) && o.group).map((o) => o.group));
  return list.filter((o) => set.has(o.id) || (o.group && groups.has(o.group))).map((o) => o.id);
}

/**
 * 図のレイアウト: the chosen pictures set out evenly with a caption each (their alternative text or name), as
 * PowerPoint's picture layouts do; the pictures fill their frames (cover), and pictures and captions are grouped.
 */
export const PICTURE_LAYOUTS = { row: "キャプション付きの図（横に並べる）", grid: "図のグリッド（2列・キャプション付き）", side: "図と説明（左に図・右に説明）" };
export function pictureLayout(list, ids, kind) {
  const pics = list.filter((o) => ids.includes(o.id) && o.kind === "image");
  if (!pics.length || !PICTURE_LAYOUTS[kind]) return { list, ids: [] };
  const b = bounds(pics);
  // At least a good part of the slide, around where the pictures were, inside the slide.
  const w = Math.min(W - 160, Math.max(b.w, 960));
  const h = Math.min(H - 300, Math.max(b.h, 500));
  const x = Math.min(Math.max(80, b.x + b.w / 2 - w / 2), W - 80 - w);
  const y = Math.min(Math.max(200, b.y + b.h / 2 - h / 2), H - 100 - h);
  const gap = 24;
  const capH = 56;
  const n = pics.length;
  const g = `g${newId().slice(1)}`;
  const caption = (o) => String(o.alt || o.name || (o.fileName || "").replace(/\.[a-z0-9]+$/i, "") || "説明を入力").replace(/[<>&]/g, "").slice(0, 80);
  const placed = new Map();
  const captions = [];
  pics.forEach((o, i) => {
    let pic;
    let cap;
    if (kind === "side") {
      const rowH = (h - gap * (n - 1)) / n;
      const picW = Math.min(rowH * (4 / 3), w * 0.45);
      const top = y + i * (rowH + gap);
      pic = { x, y: top, w: picW, h: rowH };
      cap = { x: x + picW + gap, y: top, w: w - picW - gap, h: rowH, align: "left", valign: "middle" };
    } else {
      const cols = kind === "grid" ? Math.min(2, n) : n;
      const rows = Math.ceil(n / cols);
      const cellW = (w - gap * (cols - 1)) / cols;
      const cellH = (h - gap * (rows - 1)) / rows;
      const [cx, cy] = [x + (i % cols) * (cellW + gap), y + Math.floor(i / cols) * (cellH + gap)];
      pic = { x: cx, y: cy, w: cellW, h: Math.max(40, cellH - capH - 8) };
      cap = { x: cx, y: cy + pic.h + 8, w: cellW, h: capH, align: "center", valign: "top" };
    }
    placed.set(o.id, { ...o, ...Object.fromEntries(Object.entries(pic).map(([k, v]) => [k, round2(v)])), fit: "cover", crop: undefined, rot: undefined, group: g });
    captions.push({ id: newId(), kind: "text", x: round2(cap.x), y: round2(cap.y), w: round2(cap.w), h: round2(cap.h), text: `<p>${caption(o)}</p>`, fs: 28, align: cap.align, valign: cap.valign, group: g });
  });
  return { list: [...list.map((o) => placed.get(o.id) || o), ...captions], ids: [...pics.map((o) => o.id), ...captions.map((o) => o.id)] };
}

export function group(list, ids) {
  if (ids.length < 2) return list;
  const g = `g${newId().slice(1)}`;
  const set = new Set(ids);
  // A group draws together: its members move to the front-most member's place, in their order.
  const members = list.filter((o) => set.has(o.id)).map((o) => ({ ...o, group: g }));
  const last = Math.max(...list.map((o, i) => (set.has(o.id) ? i : -1)));
  const out = [];
  list.forEach((o, i) => {
    if (!set.has(o.id)) out.push(o);
    if (i === last) out.push(...members);
  });
  return out;
}

export function ungroup(list, ids) {
  const set = new Set(ids);
  const groups = new Set(list.filter((o) => set.has(o.id) && o.group).map((o) => o.group));
  return list.map((o) => (o.group && groups.has(o.group) ? (({ group: _, ...rest }) => rest)(o) : o));
}

// ---------------------------------------------------------------- order

/** Bring to front / send to back / one step forward / one step backward (PowerPoint's 前面へ・背面へ). */
export function reorder(list, ids, how) {
  const set = new Set(ids);
  const chosen = list.filter((o) => set.has(o.id));
  const rest = list.filter((o) => !set.has(o.id));
  if (!chosen.length) return list;
  if (how === "front") return [...rest, ...chosen];
  if (how === "back") return [...chosen, ...rest];
  const out = [...list];
  if (how === "forward") {
    for (let i = out.length - 2; i >= 0; i -= 1) {
      if (set.has(out[i].id) && !set.has(out[i + 1].id)) [out[i], out[i + 1]] = [out[i + 1], out[i]];
    }
  } else if (how === "backward") {
    for (let i = 1; i < out.length; i += 1) {
      if (set.has(out[i].id) && !set.has(out[i - 1].id)) [out[i], out[i - 1]] = [out[i - 1], out[i]];
    }
  }
  return out;
}

/** Move one object to a place in the drawing order (the selection pane's drag and drop). */
export function moveTo(list, id, index) {
  const from = list.findIndex((o) => o.id === id);
  if (from < 0) return list;
  const out = [...list];
  const [item] = out.splice(from, 1);
  out.splice(Math.max(0, Math.min(out.length, index)), 0, item);
  return out;
}

// ---------------------------------------------------------------- moving, aligning, distributing

function translate(o, dx, dy) {
  if (o.kind === "line") {
    const out = { ...o, x1: round2(o.x1 + dx), y1: round2(o.y1 + dy), x2: round2(o.x2 + dx), y2: round2(o.y2 + dy) };
    // A connector moved on its own lets go of the objects it was attached to.
    delete out.from;
    delete out.to;
    return out;
  }
  return { ...o, x: round2(o.x + dx), y: round2(o.y + dy) };
}

export function moveBy(list, ids, dx, dy) {
  const set = new Set(ids);
  return list.map((o) => (set.has(o.id) ? translate(o, dx, dy) : o));
}

/**
 * Align the selection (PowerPoint's 配置): left / center / right / top / middle / bottom, to the selection's
 * own box (two or more objects) or to the slide.
 */
export function align(list, ids, how, { toSlide = false, ends = null } = {}) {
  const set = new Set(ids);
  const chosen = list.filter((o) => set.has(o.id));
  if (!chosen.length) return list;
  const area = toSlide || chosen.length === 1 ? { x: 0, y: 0, w: W, h: H } : bounds(chosen, ends);
  // A group aligns as one block.
  const units = new Map();
  for (const o of chosen) {
    const key = o.group || o.id;
    if (!units.has(key)) units.set(key, []);
    units.get(key).push(o);
  }
  let out = list;
  for (const members of units.values()) {
    const b = bounds(members, ends);
    let dx = 0;
    let dy = 0;
    if (how === "left") dx = area.x - b.x;
    if (how === "center") dx = area.x + area.w / 2 - (b.x + b.w / 2);
    if (how === "right") dx = area.x + area.w - (b.x + b.w);
    if (how === "top") dy = area.y - b.y;
    if (how === "middle") dy = area.y + area.h / 2 - (b.y + b.h / 2);
    if (how === "bottom") dy = area.y + area.h - (b.y + b.h);
    out = moveBy(out, members.map((o) => o.id), dx, dy);
  }
  return out;
}

/** Equal gaps between three or more objects (左右に整列 / 上下に整列). */
export function distribute(list, ids, axis, { ends = null } = {}) {
  const set = new Set(ids);
  const units = new Map();
  for (const o of list.filter((item) => set.has(item.id))) {
    const key = o.group || o.id;
    if (!units.has(key)) units.set(key, []);
    units.get(key).push(o);
  }
  const blocks = [...units.values()].map((members) => ({ members, b: bounds(members, ends) }));
  if (blocks.length < 3) return list;
  const pos = axis === "x" ? (b) => b.x : (b) => b.y;
  const size = axis === "x" ? (b) => b.w : (b) => b.h;
  blocks.sort((p, q) => pos(p.b) - pos(q.b));
  const first = blocks[0].b;
  const last = blocks[blocks.length - 1].b;
  const span = pos(last) + size(last) - pos(first);
  const used = blocks.reduce((sum, block) => sum + size(block.b), 0);
  const gap = (span - used) / (blocks.length - 1);
  let at = pos(first);
  let out = list;
  for (const block of blocks) {
    const d = at - pos(block.b);
    out = moveBy(out, block.members.map((o) => o.id), axis === "x" ? d : 0, axis === "x" ? 0 : d);
    at += size(block.b) + gap;
  }
  return out;
}

/** Rotate the selection by a number of degrees (a group turns about its own centre). */
export function rotateBy(list, ids, deg) {
  const set = new Set(ids);
  const chosen = list.filter((o) => set.has(o.id));
  const center = (() => { const b = bounds(chosen); return [b.x + b.w / 2, b.y + b.h / 2]; })();
  const grouped = chosen.length > 1 && chosen.every((o) => o.group && o.group === chosen[0].group);
  return list.map((o) => {
    if (!set.has(o.id)) return o;
    if (o.kind === "line") {
      const c = grouped ? center : [(o.x1 + o.x2) / 2, (o.y1 + o.y2) / 2];
      const [x1, y1] = rotatePoint([o.x1, o.y1], c, deg);
      const [x2, y2] = rotatePoint([o.x2, o.y2], c, deg);
      const out = { ...o, x1: round2(x1), y1: round2(y1), x2: round2(x2), y2: round2(y2) };
      delete out.from;
      delete out.to;
      return out;
    }
    const rot = normAngle((o.rot || 0) + deg);
    if (!grouped) return { ...o, rot };
    const [cx, cy] = rotatePoint([o.x + o.w / 2, o.y + o.h / 2], center, deg);
    return { ...o, rot, x: round2(cx - o.w / 2), y: round2(cy - o.h / 2) };
  });
}

export const normAngle = (deg) => round2((((deg + 180) % 360) + 360) % 360 - 180);

/** Flip left-right or upside down (a group flips about its own centre). */
export function flip(list, ids, axis) {
  const set = new Set(ids);
  const chosen = list.filter((o) => set.has(o.id));
  const b = bounds(chosen);
  const grouped = chosen.length > 1;
  return list.map((o) => {
    if (!set.has(o.id)) return o;
    if (o.kind === "line") {
      const mirror = (x, y) => (axis === "x" ? [b.x + b.w - (x - b.x), y] : [x, b.y + b.h - (y - b.y)]);
      const ends = grouped ? [mirror(o.x1, o.y1), mirror(o.x2, o.y2)] : axis === "x" ? [[o.x2, o.y1], [o.x1, o.y2]] : [[o.x1, o.y2], [o.x2, o.y1]];
      return { ...o, x1: round2(ends[0][0]), y1: round2(ends[0][1]), x2: round2(ends[1][0]), y2: round2(ends[1][1]) };
    }
    const out = { ...o };
    if (axis === "x") { out.flipH = !o.flipH; if (o.rot) out.rot = normAngle(-o.rot); }
    else { out.flipV = !o.flipV; if (o.rot) out.rot = normAngle(-o.rot); }
    if (!out.flipH) delete out.flipH;
    if (!out.flipV) delete out.flipV;
    if (!out.rot) delete out.rot;
    if (grouped) {
      if (axis === "x") out.x = round2(b.x + b.w - (o.x + o.w - b.x));
      else out.y = round2(b.y + b.h - (o.y + o.h - b.y));
    }
    return out;
  });
}

/** Copies of the selection with new ids (and new groups), shifted by `offset`. */
export function duplicate(list, ids, offset = 24) {
  const set = new Set(ids);
  const groupMap = new Map();
  const idMap = new Map();
  const copies = list.filter((o) => set.has(o.id)).map((o) => {
    const copy = translate(clone(o), offset, offset);
    copy.id = newId();
    idMap.set(o.id, copy.id);
    if (o.group) {
      if (!groupMap.has(o.group)) groupMap.set(o.group, `g${newId().slice(1)}`);
      copy.group = groupMap.get(o.group);
    }
    delete copy.name;
    return copy;
  });
  // Connectors copied with both ends keep them, on the copies.
  for (const [i, o] of list.filter((item) => set.has(item.id)).entries()) {
    if (o.kind !== "line") continue;
    for (const end of ["from", "to"]) if (o[end] && idMap.has(o[end].id)) copies[i][end] = { ...o[end], id: idMap.get(o[end].id) };
  }
  return { list: [...list, ...copies], ids: copies.map((o) => o.id), idMap, groupMap };
}

/** Objects copied from elsewhere (the clipboard), with new ids, placed at an offset. */
export function paste(list, objects, offset = 0) {
  const fake = [...objects];
  const { list: merged, ids, idMap, groupMap } = duplicate(fake, fake.map((o) => o.id), offset);
  const added = merged.slice(fake.length);
  return { list: [...list, ...added], ids, idMap, groupMap };
}

/** Remove objects; connectors that pointed at them stay where they were drawn. */
export function remove(list, ids, ends = null) {
  const set = new Set(ids);
  return list.filter((o) => !set.has(o.id)).map((o) => {
    if (o.kind !== "line" || !((o.from && set.has(o.from.id)) || (o.to && set.has(o.to.id)))) return o;
    const out = { ...o };
    if (ends) {
      const [[x1, y1], [x2, y2]] = ends(o);
      Object.assign(out, { x1: round2(x1), y1: round2(y1), x2: round2(x2), y2: round2(y2) });
    }
    if (o.from && set.has(o.from.id)) delete out.from;
    if (o.to && set.has(o.to.id)) delete out.to;
    return out;
  });
}

/** Apply settings to the selection (a function gets each object and returns its changes). */
export function update(list, ids, patch) {
  const set = new Set(ids);
  return list.map((o) => {
    if (!set.has(o.id)) return o;
    const changes = typeof patch === "function" ? patch(o) : patch;
    if (!changes) return o;
    const out = { ...o, ...changes };
    for (const [key, value] of Object.entries(changes)) if (value === undefined || value === null) delete out[key];
    return out;
  });
}

// ---------------------------------------------------------------- text formatting (whole objects)

/**
 * Remove one kind of inline formatting from rich text, so a setting made for the whole object shows everywhere
 * (as PowerPoint does when a shape, not a word, is selected). kind: bold, italic, underline, strike, color,
 * size, highlight, align.
 */
export function clearInline(E, html, kind) {
  if (!html) return html;
  const box = document.createElement("div");
  box.append(E.richFragment(html));
  const unwrap = (el) => el.replaceWith(...el.childNodes);
  const tags = { bold: "b", italic: "i", underline: "u", strike: "s" }[kind];
  if (tags) for (const el of [...box.querySelectorAll(tags)]) unwrap(el);
  const props = { bold: ["font-weight"], italic: ["font-style"], underline: ["text-decoration"], strike: ["text-decoration"], color: ["color"], size: ["font-size"], highlight: ["background-color"] }[kind] ?? [];
  for (const el of [...box.querySelectorAll("span")]) {
    for (const prop of props) el.style.removeProperty(prop);
    if (!el.getAttribute("style")?.trim()) unwrap(el);
  }
  if (kind === "align") for (const el of box.querySelectorAll("p, li, div")) el.removeAttribute("style");
  return E.sanitizeRich(box.innerHTML);
}

/** Turn every paragraph into a bulleted or numbered list, or every list back into paragraphs. */
export function setList(E, html, type, style = null) {
  const box = document.createElement("div");
  box.append(E.richFragment(html || "<p><br></p>"));
  const lines = [];
  const collect = (node) => {
    for (const child of [...node.childNodes]) {
      if (child.nodeName === "P") lines.push(child.innerHTML);
      else if (child.nodeName === "UL" || child.nodeName === "OL") for (const li of child.children) { const copy = li.cloneNode(true); copy.querySelectorAll("ul, ol").forEach((n) => n.remove()); lines.push(copy.innerHTML); for (const sub of li.children) if (sub.nodeName === "UL" || sub.nodeName === "OL") collect({ childNodes: [sub] }); }
    }
  };
  collect(box);
  if (!type) return E.sanitizeRich(lines.map((line) => `<p>${line || "<br>"}</p>`).join(""));
  const tag = type === "number" ? "ol" : "ul";
  const kept = style && E.LIST_STYLES?.[tag]?.[style] ? ` data-style="${style}"` : "";
  const made = E.sanitizeRich(`<${tag}${kept}>${lines.map((line) => `<li>${line || "<br>"}</li>`).join("")}</${tag}>`);
  // The marks' colour and size (and a numbered list's first number) stay through a change of style.
  const was = listPropsOf(html);
  return was ? setListProps(E, made, { start: type === "number" && was.type === "number" ? was.start : undefined, mark: was.mark, msize: was.msize }) : made;
}

/** The list's marker (箇条書き・段落番号の種類): its data-style, or the plain one. */
export function listStyleOf(html) {
  const m = String(html || "").match(/^<(ul|ol)\b([^>]*)>/);
  if (!m) return null;
  return m[2].match(/data-style="([\w-]+)"/)?.[1] || (m[1] === "ul" ? "disc" : "decimal");
}

/** 箇条書きと段落番号 of the first list in the text: its first number, marks' colour and size. */
export function listPropsOf(html) {
  const m = String(html || "").match(/^<(ul|ol)\b([^>]*)>/);
  if (!m) return null;
  const attr = (name) => m[2].match(new RegExp(`${name}="([^"]*)"`))?.[1] ?? null;
  return { type: m[1] === "ol" ? "number" : "bullet", start: Number(attr("start")) || 1, mark: attr("data-mark") || "", msize: attr("data-msize") || "100" };
}

/** Every list in the text given a first number (numbered lists), marks' colour and size; the rest kept. */
export function setListProps(E, html, { start, mark, msize } = {}) {
  const box = document.createElement("div");
  box.append(E.richFragment(html || ""));
  for (const list of box.querySelectorAll("ul, ol")) {
    if (list.parentElement?.closest("ul, ol")) continue;
    if (start !== undefined && list.nodeName === "OL") { if (start > 1) list.setAttribute("start", String(start)); else list.removeAttribute("start"); }
    if (mark !== undefined) { if (mark) list.setAttribute("data-mark", mark); else list.removeAttribute("data-mark"); }
    if (msize !== undefined) { if (msize && msize !== "100") list.setAttribute("data-msize", msize); else list.removeAttribute("data-msize"); }
  }
  return E.sanitizeRich(box.innerHTML);
}

/**
 * 段落 → インデント: every paragraph indented `left` px before its text, its first line moved by `first` px
 * (negative: ぶら下げ, the first line out to the left of the rest). 0 takes an indent away.
 */
export function setIndent(E, html, { left = 0, first = 0 } = {}) {
  const box = document.createElement("div");
  box.append(E.richFragment(html || ""));
  const blocks = box.querySelectorAll("p, li, div");
  if (!blocks.length && box.textContent) { const p = document.createElement("p"); p.append(...box.childNodes); box.append(p); }
  for (const block of box.querySelectorAll("p, li, div")) {
    block.style.marginLeft = left ? `${Math.round(left * 100) / 100}px` : "";
    block.style.textIndent = first ? `${Math.round(first * 100) / 100}px` : "";
  }
  return E.sanitizeRich(box.innerHTML);
}

/** 段落 → タブ設定: the tab stops ([{ type: "l" | "c" | "r", x }], px) on every paragraph of the text (none takes them away). */
export function setTabs(E, html, stops) {
  const text = E.tabsText(E.parseTabs(Array.isArray(stops) ? E.tabsText(stops) : ""));
  const box = document.createElement("div");
  box.append(E.richFragment(html || ""));
  if (!box.querySelector("p, li, div") && box.textContent) { const p = document.createElement("p"); p.append(...box.childNodes); box.append(p); }
  for (const block of box.querySelectorAll("p")) { if (text) block.setAttribute("data-tabs", text); else block.removeAttribute("data-tabs"); }
  return E.sanitizeRich(box.innerHTML);
}
/** The first paragraph's tab stops: [{ type, x }]. */
export function tabsOf(E, html) {
  const m = /^<p\b[^>]*\bdata-tabs="([^"]*)"/.exec(String(html || ""));
  return m ? E.parseTabs(m[1]) : [];
}

/** The first paragraph's indents (px): { left, first }. */
export function indentOf(html) {
  const tag = /^<(p|li|div)\b[^>]*style="([^"]*)"/.exec(String(html || "").replace(/^<(ul|ol)\b[^>]*>/, ""));
  const read = (name) => { const m = new RegExp(`${name}:\\s*(-?[\\d.]+)px`).exec(tag?.[2] || ""); return m ? Number(m[1]) : 0; };
  return { left: read("margin-left"), first: read("text-indent") };
}

/** Which list the text is (all bullets / all numbers / none). */
export function listOf(html) {
  const text = String(html || "");
  if (/^<ul[ >]/.test(text) && !/<p[ >]/.test(text)) return "bullet";
  if (/^<ol[ >]/.test(text) && !/<p[ >]/.test(text)) return "number";
  return null;
}

// ---------------------------------------------------------------- units

export const toCm = (px) => Math.round((px / PX_PER_CM) * 100) / 100;
export const fromCm = (cm) => round2(Number(cm) * PX_PER_CM);
export const toPt = (px) => Math.round((px / PX_PER_PT) * 10) / 10;
export const fromPt = (pt) => round2(Number(pt) * PX_PER_PT);

// ---------------------------------------------------------------- animations (slide.timeline, engine/animate.js)

/**
 * Keep a slide's animations in step with its objects: animations of deleted objects go, a group's animation
 * becomes one per former member when the group is undone (the first keeps its start, the rest play with it),
 * and a trigger that is gone leaves its animations to play on the slide's clicks.
 */
export function reconcileTimeline(timeline, before, after) {
  if (!Array.isArray(timeline) || !timeline.length) return timeline;
  const ids = new Set(after.map((o) => o.id));
  const groups = new Set(after.map((o) => o.group).filter(Boolean));
  const out = [];
  for (const e of timeline) {
    if (!e || typeof e.el !== "string") continue;
    let entries = [e];
    if (e.el.startsWith("grp:") && !groups.has(e.el.slice(4))) {
      const members = before.filter((o) => o.group === e.el.slice(4) && ids.has(o.id));
      entries = members.map((o, i) => ({ ...e, id: i ? `${e.id}-${i}`.slice(0, 32) : e.id, el: o.id, ...(i ? { start: "with", delay: e.delay || 0 } : {}) }));
    } else if (!e.el.startsWith("@") && !e.el.startsWith("grp:") && !ids.has(e.el)) entries = [];
    for (const entry of entries) {
      // A trigger may be an object ("id") or one of a media's bookmarks ("id@mark"): gone with its object.
      if (entry.trigger && !ids.has(String(entry.trigger).split("@")[0])) { const { trigger: _, ...rest } = entry; out.push(rest); } else out.push(entry);
    }
  }
  return out;
}

/** Move animations up or down the list (PowerPoint's 順番を前にする / 後にする). */
export function moveAnimations(timeline, ids, step) {
  const list = [...timeline];
  const set = new Set(ids);
  const order = step < 0 ? list.map((_, i) => i) : list.map((_, i) => list.length - 1 - i);
  for (const i of order) {
    if (!set.has(list[i].id)) continue;
    const j = i + (step < 0 ? -1 : 1);
    if (j < 0 || j >= list.length || set.has(list[j].id)) continue;
    [list[i], list[j]] = [list[j], list[i]];
  }
  return list;
}

// ---------------------------------------------------------------- tables (kind "table")
// A table is edited as a grid where every position points at the cell that covers it, so a merged cell is one
// object in several positions: inserting, deleting, merging and splitting then keep merges right by themselves.

const copyCell = (cell) => JSON.parse(JSON.stringify(cell || {}));
/** A sheet's formulas follow its rows and columns (a table that is not a sheet keeps its words as they are). */
const sheetRefs = (o, fn) => (o.sheet ? mapFormulas(o, fn) : o);
function expand(cells) {
  const grid = cells.map((row) => row.map(() => null));
  cells.forEach((row, r) => row.forEach((cell, c) => {
    if (cell.merged || grid[r][c]) return;
    const own = copyCell(cell);
    delete own.rs; delete own.cs; delete own.merged;
    for (let i = r; i < Math.min(cells.length, r + (cell.rs || 1)); i += 1) for (let j = c; j < Math.min(row.length, c + (cell.cs || 1)); j += 1) grid[i][j] = own;
  }));
  return grid.map((row) => row.map((cell) => cell || {}));
}
function collapse(grid) {
  const done = new Set();
  return grid.map((row, r) => row.map((cell, c) => {
    if (done.has(cell)) return { merged: true };
    done.add(cell);
    let rs = 1;
    let cs = 1;
    while (r + rs < grid.length && grid[r + rs][c] === cell) rs += 1;
    while (c + cs < row.length && row[c + cs] === cell) cs += 1;
    const out = copyCell(cell);
    if (rs > 1) out.rs = rs;
    if (cs > 1) out.cs = cs;
    return out;
  }));
}
/** 罫線 (テーブル デザイン → 罫線): which sides each kind draws, for a cell at the range's edges or inside it. */
export const TABLE_BORDERS = [
  ["bottom", "下罫線"], ["top", "上罫線"], ["left", "左罫線"], ["right", "右罫線"], ["none", "枠なし"], ["all", "格子"],
  ["outside", "外枠"], ["inside", "内側"], ["insideH", "横罫線（内側）"], ["insideV", "縦罫線（内側）"],
];
/**
 * The cells r0..r1 × c0..c1 given borders of a kind drawn with the pen ({ c: colour, w: width px }), or none. The
 * cell across each edge gets the same line (borders collapse into one); a merged cell takes a line only on its own edges.
 */
export function tableBorders(o, r0, c0, r1, c1, kind, pen) {
  const grid = expand(o.cells);
  const rows = grid.length;
  const cols = grid[0]?.length || 0;
  const [ra, rb] = [Math.max(0, Math.min(r0, r1)), Math.min(rows - 1, Math.max(r0, r1))];
  const [ca, cb] = [Math.max(0, Math.min(c0, c1)), Math.min(cols - 1, Math.max(c0, c1))];
  const line = kind === "none" ? "none" : { c: pen.c, w: pen.w };
  const at = (i, j) => grid[i]?.[j];
  for (let i = ra; i <= rb; i += 1) for (let j = ca; j <= cb; j += 1) {
    const cell = at(i, j);
    const top = i === ra;
    const bottom = i === rb;
    const left = j === ca;
    const right = j === cb;
    const want = {
      bottom: { bb: bottom }, top: { bt: top }, left: { bl: left }, right: { br: right },
      outside: { bt: top, bb: bottom, bl: left, br: right },
      inside: { bt: !top, bb: !bottom, bl: !left, br: !right },
      insideH: { bt: !top, bb: !bottom }, insideV: { bl: !left, br: !right },
      all: { bt: true, bb: true, bl: true, br: true }, none: { bt: true, bb: true, bl: true, br: true },
    }[kind] || {};
    // A side is the cell's own only where the next position is another cell (a merge covers the rest).
    const across = { bt: [i - 1, j, "bb"], bb: [i + 1, j, "bt"], bl: [i, j - 1, "br"], br: [i, j + 1, "bl"] };
    for (const [side, on] of Object.entries(want)) {
      if (!on) continue;
      const [ni, nj, facing] = across[side];
      const next = at(ni, nj);
      if (next === cell) continue;
      cell[side] = line === "none" ? "none" : { ...line };
      if (next) next[facing] = line === "none" ? "none" : { ...line };
    }
  }
  return { ...o, cells: collapse(grid) };
}

/** The cell (top-left of a merge) that covers row r, column c. */
export function tableOrigin(o, r, c) {
  const grid = expand(o.cells);
  const cell = grid[r]?.[c];
  for (let i = 0; i < grid.length; i += 1) for (let j = 0; j < grid[i].length; j += 1) if (grid[i][j] === cell) return [i, j];
  return [r, c];
}
const px = (shares, total) => shares.map((f) => f * total);
const toShares = (values) => { const sum = values.reduce((a, b) => a + b, 0) || 1; return values.map((v) => Math.round((v / sum) * 100000) / 100000); };

/** A new table: rows × cols of empty cells (the first row is its header). */
export function makeTable(rows, cols, box) {
  return { kind: "table", ...box, cells: Array.from({ length: rows }, () => Array.from({ length: cols }, () => ({}))), cols: Array.from({ length: cols }, () => 1 / cols), rows: Array.from({ length: rows }, () => 1 / rows) };
}
// ---------------------------------------------------------------- 塗りつぶし → グラデーション, 図のスタイル

// Gradients in the theme's colours: [key, label, from, to], each end a slot ("accent1") or a slot made lighter
// ("accent1+80" = 80 % lighter, as the colour sheet's rows). makeGradient resolves them in the deck's theme.
export const GRADIENT_SETS = [
  ["accent1", "アクセント 1 → 明るく", "accent1", "accent1+60"],
  ["accent1-soft", "アクセント 1（淡）→ 背景 1", "accent1+80", "bg1"],
  ["accent2", "アクセント 2 → 明るく", "accent2", "accent2+60"],
  ["accent1-2", "アクセント 1 → アクセント 2", "accent1", "accent2"],
  ["bg", "背景 2 → 背景 1", "bg2", "bg1"],
  ["tx", "テキスト 2 → テキスト 1", "tx2", "tx1"],
];
// Office's default theme (when no theme is given).
const OFFICE_COLORS = { bg1: "#ffffff", tx1: "#000000", bg2: "#e7e6e6", tx2: "#44546a", accent1: "#4472c4", accent2: "#ed7d31", accent3: "#a5a5a5", accent4: "#ffc000", accent5: "#5b9bd5", accent6: "#70ad47" };
/** A colour made lighter by pct % the way Office does (HSL lightness × (1 − pct) + pct). */
function lighter(hex, pct) {
  const [r, g, b] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let l = (max + min) / 2;
  const d = max - min;
  const sat = d === 0 ? 0 : l > 0.5 ? d / (2 - max - min) : d / (max + min);
  const hue = d === 0 ? 0 : (max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4) / 6;
  l = Math.min(1, l * (1 - pct / 100) + pct / 100);
  const f = (n) => { const k = (n + hue * 12) % 12; const a = sat * Math.min(l, 1 - l); return l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1)); };
  return `#${[0, 8, 4].map((n) => Math.round(Math.max(0, Math.min(1, f(n))) * 255).toString(16).padStart(2, "0")).join("")}`;
}
/** One end of a gradient set in a theme's colours ({ bg1, tx1, …, accent6 }). */
function gradientEnd(spec, colors) {
  const [slot, pct] = spec.split("+");
  const base = (colors?.[slot] || OFFICE_COLORS[slot] || "#808080").toLowerCase();
  return pct ? lighter(base, Number(pct)) : base;
}
// Where the first colour is: angle 0 = left, 90 = top (the colour runs towards the opposite side).
export const GRADIENT_DIRECTIONS = [["down", "下方向", 90], ["up", "上方向", -90], ["right", "右方向", 0], ["left", "左方向", 180], ["diagDown", "右下方向", 45], ["diagUp", "右上方向", -45]];

/** A linear gradient of one of the sets in a theme's colours, running in one of the directions (the object's `gradient`). */
export function makeGradient(setKey, directionKey = "down", colors = OFFICE_COLORS) {
  const set = GRADIENT_SETS.find(([key]) => key === setKey) || GRADIENT_SETS[0];
  const dir = GRADIENT_DIRECTIONS.find(([key]) => key === directionKey) || GRADIENT_DIRECTIONS[0];
  return { angle: dir[2], stops: [{ at: 0, color: gradientEnd(set[2], colors), opacity: 1 }, { at: 1, color: gradientEnd(set[3], colors), opacity: 1 }] };
}

/** Which set and direction a gradient is (null when it is none of them, as one brought over from PowerPoint). */
export function gradientChoice(g, colors = OFFICE_COLORS) {
  if (!g?.stops || g.stops.length !== 2) return null;
  const set = GRADIENT_SETS.find(([, , from, to]) => gradientEnd(from, colors) === g.stops[0].color && gradientEnd(to, colors) === g.stops[1].color);
  const dir = GRADIENT_DIRECTIONS.find(([, , angle]) => angle === g.angle);
  return set && dir ? { set: set[0], direction: dir[0] } : null;
}

// 図のスタイル (the quick styles of a picture): a frame (a mask) and an outline, as PowerPoint's (white and black
// frames, so they suit any theme) — [label, change].
const NO_LOOK = { mask: undefined, adj: undefined, stroke: undefined, strokeW: undefined, dash: undefined, cmpd: undefined, cap: undefined, join: undefined };
export const PICTURE_STYLES = [
  ["標準（枠なし）", { ...NO_LOOK }],
  ["細い枠線（黒）", { ...NO_LOOK, stroke: "#000000", strokeW: 3 }],
  ["太い枠線（グレー）", { ...NO_LOOK, stroke: "#808080", strokeW: 12 }],
  ["シンプルな枠（白）", { ...NO_LOOK, stroke: "#ffffff", strokeW: 24 }],
  ["二重線の枠（黒）", { ...NO_LOOK, stroke: "#000000", strokeW: 18, cmpd: "dbl" }],
  ["角丸", { ...NO_LOOK, mask: "roundRect" }],
  ["角丸・細い枠線（黒）", { ...NO_LOOK, mask: "roundRect", stroke: "#000000", strokeW: 3 }],
  ["角丸・太い枠線（白）", { ...NO_LOOK, mask: "roundRect", stroke: "#ffffff", strokeW: 18 }],
  ["楕円", { ...NO_LOOK, mask: "ellipse" }],
  ["楕円・太い枠線（グレー）", { ...NO_LOOK, mask: "ellipse", stroke: "#808080", strokeW: 12 }],
  ["斜めカット", { ...NO_LOOK, mask: "snip2DiagRect" }],
  ["斜めカット・細い枠線（黒）", { ...NO_LOOK, mask: "snip2DiagRect", stroke: "#000000", strokeW: 3 }],
];

// ---------------------------------------------------------------- tables pasted from Excel or a web page

/** Tab-separated text as rows of strings (Excel's own quoting: "…" around a cell with a tab, a newline or a quote;
 *  a doubled quote inside). */
export function parseTsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  const src = String(text ?? "").replace(/\r\n?/g, "\n");
  for (let i = 0; i < src.length; i += 1) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"' && src[i + 1] === '"') { cell += '"'; i += 1; } else if (ch === '"') quoted = false; else cell += ch;
    } else if (ch === '"' && cell === "") quoted = true;
    else if (ch === "\t") { row.push(cell); cell = ""; }
    else if (ch === "\n") { row.push(cell); rows.push(row); row = []; cell = ""; }
    else cell += ch;
  }
  if (cell !== "" || row.length) { row.push(cell); rows.push(row); }
  return rows;
}

const escapeHtml = (t) => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const NUMERIC = /^[-+−▲△]?[¥￥$€£]?\s?[-+−]?\d[\d,，]*(\.\d+)?\s?[%％円万億千]*$/;
/** A pasted cell: its words as paragraphs, numbers aligned right (as the Excel sheet shows them). */
function pastedCell(raw, extra = {}) {
  const lines = String(raw ?? "").replace(/\u00a0/g, " ").split(/\n/).map((line) => line.replace(/[ \t]+/g, " ").trim());
  while (lines.length && !lines[lines.length - 1]) lines.pop();
  const words = lines.join("\n").trim();
  if (!words) return { ...extra };
  return { text: lines.map((line) => `<p>${escapeHtml(line) || "<br>"}</p>`).join(""), ...(lines.length === 1 && NUMERIC.test(words) ? { align: "right" } : {}), ...extra };
}

/** The first table of some HTML (Excel's and web pages' clipboard) as a grid of cells, merged spans kept; null
 *  when there is none. `parse` turns HTML into a document (the browser's DOMParser by default). */
export function tableGridFromHtml(html, parse = (source) => new DOMParser().parseFromString(source, "text/html")) {
  if (!/<table[\s>]/i.test(String(html || ""))) return null;
  const table = parse(String(html)).querySelector("table");
  if (!table) return null;
  const grid = [];
  const trs = [...table.querySelectorAll("tr")].filter((tr) => tr.closest("table") === table);
  trs.forEach((tr, r) => {
    grid[r] = grid[r] || [];
    let c = 0;
    for (const td of [...tr.children].filter((el) => /^(td|th)$/i.test(el.tagName))) {
      while (grid[r][c]) c += 1;
      const cs = Math.min(30, Math.max(1, parseInt(td.getAttribute("colspan"), 10) || 1));
      const rs = Math.min(100, Math.max(1, parseInt(td.getAttribute("rowspan"), 10) || 1));
      for (const br of td.querySelectorAll("br")) br.replaceWith("\n");
      const style = String(td.getAttribute("style") || "");
      const bold = /^th$/i.test(td.tagName) ? false : /font-weight:\s*(bold|[6-9]00)/i.test(style) || (td.children.length === 1 && /^(b|strong)$/i.test(td.children[0].tagName) && td.children[0].textContent.trim() === td.textContent.trim());
      const aligned = /text-align:\s*(left|center|right)/i.exec(style) || /^(left|center|right)$/i.exec(td.getAttribute("align") || "");
      const cell = pastedCell(td.textContent, { ...(bold ? { bold: true } : {}), ...(aligned ? { align: aligned[1].toLowerCase() } : {}), ...(rs > 1 ? { rs } : {}), ...(cs > 1 ? { cs } : {}) });
      for (let i = 0; i < rs; i += 1) for (let j = 0; j < cs; j += 1) {
        grid[r + i] = grid[r + i] || [];
        grid[r + i][c + j] = i === 0 && j === 0 ? cell : { merged: true };
      }
      c += cs;
    }
  });
  const width = Math.max(0, ...grid.map((row) => row.length));
  const rows = grid.filter(Boolean).map((row) => Array.from({ length: width }, (_, c) => row[c] || {}));
  return rows.length && width ? rows : null;
}

/** Cells on the clipboard, if what is there is a sheet of them: a table in the HTML, or tab-separated text with more
 *  than one cell. Null for plain words. */
export function clipboardGrid({ text = "", html = "" } = {}, parse) {
  let rows = null;
  try { rows = tableGridFromHtml(html, parse); } catch { rows = null; }
  if (!rows && /\t/.test(text || "")) {
    const sheet = parseTsv(text).filter((row, i, all) => row.length > 1 || row[0]?.trim() || i < all.length - 1);
    const width = Math.max(0, ...sheet.map((row) => row.length));
    if (sheet.length && width > 1) rows = sheet.map((row) => Array.from({ length: width }, (_, c) => pastedCell(row[c])));
  }
  if (!rows || (rows.length === 1 && rows[0].length === 1)) return null;
  return { rows: rows.slice(0, 100).map((row) => row.slice(0, 30)) };
}

/** A new table object from pasted cells, sized for the slide (smaller words for a big sheet). */
export function tableFromGrid(grid) {
  const rows = grid.rows.length;
  const cols = grid.rows[0].length;
  const w = Math.min(1640, Math.max(480, cols * 240));
  const rowH = Math.max(36, Math.min(76, Math.floor(820 / rows)));
  const o = makeTable(rows, cols, { x: Math.round((W - w) / 2), y: 200, w, h: rows * rowH });
  o.cells = grid.rows.map((row) => row.map((cell) => ({ ...cell })));
  if (rows > 10 || cols > 7) o.fs = 18; else if (rows > 7 || cols > 5) o.fs = 22; else if (rows > 5) o.fs = 24;
  return o;
}

/** Paste a sheet into a table starting at row r0, column c0: the table grows by rows and columns as it needs, the
 *  words go into the cells (a merged cell takes the words of its first position). */
export function tableFill(o, r0, c0, grid) {
  let t = o;
  const needRows = r0 + grid.rows.length;
  const needCols = c0 + Math.max(...grid.rows.map((row) => row.length));
  while (t.cells.length < Math.min(500, needRows)) t = tableInsertRow(t, t.cells.length);
  while (t.cells[0].length < Math.min(100, needCols)) t = tableInsertCol(t, t.cells[0].length);
  const cells = t.cells.map((row) => row.map((cell) => ({ ...cell })));
  grid.rows.forEach((row, i) => row.forEach((cell, j) => {
    const target = cells[r0 + i]?.[c0 + j];
    if (!target || target.merged) return;
    const { text, align, bold } = cell;
    if (text) target.text = text; else delete target.text;
    if (align) target.align = align;
    if (bold) target.bold = true;
  }));
  return { ...t, cells };
}

/** Insert a row before `at` (a merge crossing that line grows over the new row); the table grows by one row. */
export function tableInsertRow(o, at) {
  const grid = expand(o.cells);
  at = Math.max(0, Math.min(grid.length, at));
  const row = grid[0].map((_, c) => (at > 0 && at < grid.length && grid[at - 1][c] === grid[at][c] ? grid[at][c] : {}));
  grid.splice(at, 0, row);
  const heights = px(o.rows, o.h);
  const added = heights[Math.min(at, heights.length - 1)] || o.h / heights.length;
  heights.splice(at, 0, added);
  return sheetRefs({ ...o, cells: collapse(grid), rows: toShares(heights), h: o.h + added }, (f) => insertRefs(f, "row", at));
}
/** Insert a column before `at`; the table grows by that column's width. */
export function tableInsertCol(o, at) {
  const grid = expand(o.cells);
  at = Math.max(0, Math.min(grid[0].length, at));
  grid.forEach((row) => row.splice(at, 0, at > 0 && at < row.length && row[at - 1] === row[at] ? row[at] : {}));
  const widths = px(o.cols, o.w);
  const added = widths[Math.min(at, widths.length - 1)] || o.w / widths.length;
  widths.splice(at, 0, added);
  return sheetRefs({ ...o, cells: collapse(grid), cols: toShares(widths), w: o.w + added }, (f) => insertRefs(f, "col", at));
}
/** Delete rows r0…r1 (the table shrinks); null when nothing would be left. */
export function tableDeleteRows(o, r0, r1) {
  const [a, b] = [Math.min(r0, r1), Math.max(r0, r1)];
  if (b - a + 1 >= o.cells.length) return null;
  const grid = expand(o.cells).filter((_, r) => r < a || r > b);
  const heights = px(o.rows, o.h);
  const removed = heights.slice(a, b + 1).reduce((s, v) => s + v, 0);
  return sheetRefs({ ...o, cells: collapse(grid), rows: toShares(heights.filter((_, r) => r < a || r > b)), h: Math.max(20, o.h - removed) }, (f) => deleteRefs(f, "row", a, b));
}
/** Delete columns c0…c1 (the table narrows); null when nothing would be left. */
export function tableDeleteCols(o, c0, c1) {
  const [a, b] = [Math.min(c0, c1), Math.max(c0, c1)];
  if (b - a + 1 >= o.cells[0].length) return null;
  const grid = expand(o.cells).map((row) => row.filter((_, c) => c < a || c > b));
  const widths = px(o.cols, o.w);
  const removed = widths.slice(a, b + 1).reduce((s, v) => s + v, 0);
  return sheetRefs({ ...o, cells: collapse(grid), cols: toShares(widths.filter((_, c) => c < a || c > b)), w: Math.max(20, o.w - removed) }, (f) => deleteRefs(f, "col", a, b));
}
/** Merge a block of cells into one (the words of each come along, one paragraph after another). */
export function tableMerge(o, r0, c0, r1, c1) {
  const grid = expand(o.cells);
  const [ra, rb, ca, cb] = [Math.min(r0, r1), Math.max(r0, r1), Math.min(c0, c1), Math.max(c0, c1)];
  // Grow the block to whole merges it cuts through.
  let [top, bottom, left, right] = [ra, rb, ca, cb];
  for (let grew = true; grew;) {
    grew = false;
    const inside = new Set();
    for (let i = top; i <= bottom; i += 1) for (let j = left; j <= right; j += 1) inside.add(grid[i][j]);
    grid.forEach((row, i) => row.forEach((cell, j) => {
      if (!inside.has(cell)) return;
      if (i < top) { top = i; grew = true; } if (i > bottom) { bottom = i; grew = true; }
      if (j < left) { left = j; grew = true; } if (j > right) { right = j; grew = true; }
    }));
  }
  const first = grid[top][left];
  const texts = [];
  const seen = new Set();
  for (let i = top; i <= bottom; i += 1) for (let j = left; j <= right; j += 1) {
    const cell = grid[i][j];
    if (seen.has(cell)) continue;
    seen.add(cell);
    if (cell.text && cell.text.replace(/<[^>]+>/g, "").trim()) texts.push(cell.text);
  }
  const merged = { ...first, ...(texts.length ? { text: texts.join("") } : {}) };
  for (let i = top; i <= bottom; i += 1) for (let j = left; j <= right; j += 1) grid[i][j] = merged;
  return { ...o, cells: collapse(grid) };
}
/** Split a merged cell back into single cells (its words stay in the first). */
export function tableSplit(o, r, c) {
  const grid = expand(o.cells);
  const cell = grid[r]?.[c];
  if (!cell) return o;
  let first = true;
  grid.forEach((row, i) => row.forEach((x, j) => {
    if (x !== cell) return;
    if (first) { first = false; return; }
    grid[i][j] = { ...(cell.fill ? { fill: cell.fill } : {}) };
  }));
  return { ...o, cells: collapse(grid) };
}
/** Settings for a block of cells (fill, colour, bold, alignment…); undefined removes a setting. */
export function tableCells(o, r0, c0, r1, c1, patch) {
  const grid = expand(o.cells);
  const done = new Set();
  for (let i = Math.min(r0, r1); i <= Math.max(r0, r1); i += 1) for (let j = Math.min(c0, c1); j <= Math.max(c0, c1); j += 1) {
    const cell = grid[i]?.[j];
    if (!cell || done.has(cell)) continue;
    done.add(cell);
    const changes = typeof patch === "function" ? patch(cell) : patch;
    for (const [k, v] of Object.entries(changes || {})) { if (v === undefined || v === null) delete cell[k]; else cell[k] = v; }
  }
  return { ...o, cells: collapse(grid) };
}
/** The same height for rows r0…r1, or the same width for columns c0…c1 (高さを揃える・幅を揃える). */
export function tableDistribute(o, axis, a = 0, b = Infinity) {
  const key = axis === "rows" ? "rows" : "cols";
  const list = [...o[key]];
  const [lo, hi] = [Math.max(0, Math.min(a, b)), Math.min(list.length - 1, Math.max(a, b))];
  const share = list.slice(lo, hi + 1).reduce((s, v) => s + v, 0) / (hi - lo + 1);
  for (let i = lo; i <= hi; i += 1) list[i] = share;
  return { ...o, [key]: toShares(list) };
}
/**
 * セルのサイズ (高さ・幅): rows a..b (or columns) each `size` slide pixels; the table grows or shrinks by the
 * difference and the other rows keep their own size.
 */
export function tableSetSize(o, axis, a, b, size) {
  const key = axis === "rows" ? "rows" : "cols";
  const total = key === "rows" ? o.h : o.w;
  const sizes = px(o[key], total);
  const [lo, hi] = [Math.max(0, Math.min(a, b)), Math.min(sizes.length - 1, Math.max(a, b))];
  const each = Math.max(12, Math.min(4000, Number(size) || 0));
  for (let i = lo; i <= hi; i += 1) sizes[i] = each;
  const next = sizes.reduce((sum, v) => sum + v, 0);
  return { ...o, [key]: toShares(sizes), [key === "rows" ? "h" : "w"]: round2(next) };
}
/** Move the line between column i and i+1 (or row) by d slide pixels, keeping both at least `min` wide. */
export function tableResizeLine(o, axis, i, d, min = 24) {
  const key = axis === "rows" ? "rows" : "cols";
  const total = key === "rows" ? o.h : o.w;
  const sizes = px(o[key], total);
  if (i < 0 || i >= sizes.length - 1) return o;
  const move = Math.max(min - sizes[i], Math.min(sizes[i + 1] - min, d));
  sizes[i] += move;
  sizes[i + 1] -= move;
  return { ...o, [key]: toShares(sizes) };
}
/** The text of a table as tab-separated lines (and back), for copying cells as text. */
export function tableFromText(text) {
  const lines = String(text || "").replace(/\r/g, "").split("\n").filter((line, i, all) => line || i < all.length - 1).slice(0, 40);
  const rows = lines.map((line) => line.split("\t").slice(0, 20));
  const cols = Math.max(...rows.map((row) => row.length), 1);
  return rows.map((row) => Array.from({ length: cols }, (_, c) => (row[c] ? { text: `<p>${row[c].replace(/[&<>]/g, (ch) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[ch])}</p>` } : {})));
}

/** The colour of `list` nearest to `hex` (RGB distance). */
export function nearestColor(hex, list) {
  const rgb = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const [r, g, b] = rgb(hex);
  let best = list[0];
  let d = Infinity;
  for (const c of list) { const [x, y, z] = rgb(c); const e = (r - x) ** 2 + (g - y) ** 2 + (b - z) ** 2; if (e < d) { d = e; best = c; } }
  return best;
}

/** 貼り付け先のスタイルを使用: an object's colours moved to the nearest of the deck's palette (fills, lines, words, markers). */
export function snapToPalette(o, P) {
  const hex = (c) => typeof c === "string" && /^#[0-9a-f]{6}$/i.test(c);
  const pick = (c, key) => nearestColor(c.toLowerCase(), P[key].map(([x]) => x));
  const out = { ...o };
  if (hex(out.fill)) out.fill = pick(out.fill, "fill");
  if (hex(out.stroke)) out.stroke = pick(out.stroke, "line");
  if (hex(out.color)) out.color = pick(out.color, out.kind === "icon" ? "line" : "text");
  if (out.gradient?.stops) out.gradient = { ...out.gradient, stops: out.gradient.stops.map((s) => (hex(s.color) ? { ...s, color: pick(s.color, "fill") } : s)) };
  if (out.pattern) out.pattern = { ...out.pattern, fg: hex(out.pattern.fg) ? pick(out.pattern.fg, "fill") : out.pattern.fg, bg: hex(out.pattern.bg) ? pick(out.pattern.bg, "fill") : out.pattern.bg };
  if (typeof out.text === "string") {
    out.text = out.text
      .replace(/background-color:\s*(#[0-9a-f]{6})/gi, (_, c) => `background-color: ${pick(c, "highlight")}`)
      .replace(/(^|[^-])color:\s*(#[0-9a-f]{6})/gi, (_, pre, c) => `${pre}color: ${pick(c, "text")}`);
  }
  return out;
}

// ---------------------------------------------------------------- 図形の効果 (影・反射・光彩・ぼかし)
// What PowerPoint's 図形の効果 ▾ offers, as plain data and pure functions: the studio stores `shadow`, `reflect`, `glow`
// and `soft` on a shape, text box or picture (objects.js normalizes and draws them). Sizes are pt in the screens (2 px each).

export const SHADOW_DIRECTIONS = [["br", "右下", 1, 1], ["b", "下", 0, 1], ["bl", "左下", -1, 1], ["r", "右", 1, 0], ["l", "左", -1, 0], ["tr", "右上", 1, -1], ["t", "上", 0, -1], ["tl", "左上", -1, -1]];
export const SHADOW_COLORS = [["#000000", "黒"], ["#808080", "グレー"], ["#002060", "濃い青"]];
export const REFLECTIONS = [["small", "小（近接）", 0.25], ["medium", "中", 0.5], ["large", "大（全反射）", 0.9]];
export const GLOW_SIZES = [[10, "5 pt"], [16, "8 pt"], [22, "11 pt"], [36, "18 pt"]];
export const GLOW_COLORS = [["#808080", "グレー"], ["#00b050", "緑"], ["#0070c0", "青"], ["#ffc000", "オレンジ"]];
export const SOFT_SIZES = [[2, "1 pt"], [5, "2.5 pt"], [10, "5 pt"], [20, "10 pt"], [50, "25 pt"], [100, "50 pt"]];

/** An outer shadow towards one of the eight directions (the first column of the 影 gallery). */
export function shadowPreset(key, { distance = 8, blur = 10, color = "#000000", opacity = 0.35 } = {}) {
  const dir = SHADOW_DIRECTIONS.find(([k]) => k === key);
  return dir ? { dx: dir[2] * distance, dy: dir[3] * distance, blur, color, opacity } : null;
}
/** Which of the eight directions a shadow is closest to (for the tick in the gallery), or "" when it has none. */
export function shadowDirection(shadow) {
  if (!shadow || (!shadow.dx && !shadow.dy)) return "";
  const sx = Math.abs(shadow.dx) < 1 ? 0 : Math.sign(shadow.dx);
  const sy = Math.abs(shadow.dy) < 1 ? 0 : Math.sign(shadow.dy);
  return SHADOW_DIRECTIONS.find(([, , x, y]) => x === sx && y === sy)?.[0] || "";
}
/** The dialog's way of writing a shadow: a distance (pt) and an angle (degrees, 0 = right, 90 = down). */
export function shadowPolar(shadow) {
  if (!shadow) return { distance: 4, angle: 45, blur: 5, opacity: 0.35, color: "#000000" };
  return { distance: Math.round((Math.hypot(shadow.dx, shadow.dy) / 2) * 10) / 10, angle: Math.round((Math.atan2(shadow.dy, shadow.dx) * 180) / Math.PI), blur: Math.round((shadow.blur / 2) * 10) / 10, opacity: shadow.opacity, color: shadow.color };
}
export function shadowFromPolar({ distance, angle, blur, opacity, color }) {
  const d = (Number(distance) || 0) * 2;
  const a = ((Number(angle) || 0) * Math.PI) / 180;
  return { dx: Math.round(Math.cos(a) * d * 10) / 10, dy: Math.round(Math.sin(a) * d * 10) / 10, blur: Math.max(0, Math.round((Number(blur) || 0) * 2)), color: color || "#000000", opacity: Math.min(1, Math.max(0, Number(opacity) ?? 0.35)) };
}
export function reflectionPreset(key) {
  const r = REFLECTIONS.find(([k]) => k === key);
  return r ? { size: r[2], opacity: 0.4, gap: 2 } : null;
}
// 文字の効果 (図形の書式 → ワードアートのスタイル → 文字の効果): the same 影 and 光彩, on the letters (smaller than a shape's).
export const TEXT_GLOW_SIZES = [[6, "3 pt"], [10, "5 pt"], [16, "8 pt"], [24, "12 pt"]];
/** A gentle shadow of the letters towards one of the eight directions. */
export function textShadowPreset(key) {
  return shadowPreset(key, { distance: 3, blur: 4, color: "#000000", opacity: 0.4 });
}
// 文字の輪郭: the line round the letters (px; the screens say pt, 2 px each).
export const TEXT_OUTLINE_WIDTHS = [[1, "0.5 pt"], [2, "1 pt"], [3, "1.5 pt"], [4, "2 pt"], [6, "3 pt"]];
export const TEXT_OUTLINE_COLORS = [["#000000", "黒"], ["#0070c0", "青"], ["#ffffff", "白"], ["#808080", "グレー"]];
/** A text outline of a colour (the width kept from the one it has, else 1 pt). */
export function textOutline(color, current = null) {
  return { w: current?.w ?? 2, color };
}
/** The text effects an object has, as words ("影・光彩・輪郭"), or "". */
export function textEffectsWords(o) {
  return [o?.tshadow ? "影" : "", o?.tglow ? "光彩" : "", o?.toutline ? "輪郭" : ""].filter(Boolean).join("・");
}
// 3-D 回転: the gallery's turns (degrees about the x axis — the top goes back when positive — and the y axis — the right goes back
// when positive) and the camera's field of view (45° is a natural view, 0° a flat isometric one).
export const ROT3D_PRESETS = [
  ["right", "右を奥に", { x: 0, y: 28 }], ["left", "左を奥に", { x: 0, y: -28 }], ["top", "上を奥に", { x: 28, y: 0 }], ["bottom", "下を奥に", { x: -28, y: 0 }],
  ["topRight", "右上を奥に", { x: 22, y: 22 }], ["topLeft", "左上を奥に", { x: 22, y: -22 }], ["bottomRight", "右下を奥に", { x: -22, y: 22 }], ["bottomLeft", "左下を奥に", { x: -22, y: -22 }],
  ["isoTopLeft", "等角図法：左上", { x: 35, y: -45, p: 0 }], ["isoTopRight", "等角図法：右上", { x: 35, y: 45, p: 0 }], ["isoBottomLeft", "等角図法：左下", { x: -35, y: -45, p: 0 }], ["isoBottomRight", "等角図法：右下", { x: -35, y: 45, p: 0 }],
];
/** A preset's turn as an object's `rot3d` (null for an unknown name). */
export function rot3dPreset(key) {
  const found = ROT3D_PRESETS.find(([k]) => k === key);
  return found ? { ...found[2] } : null;
}
/** Which preset a turn is (or ""). */
export function rot3dKey(rot3d) {
  if (!rot3d) return "";
  return ROT3D_PRESETS.find(([, , t]) => t.x === rot3d.x && t.y === rot3d.y && (t.p ?? 45) === (rot3d.p ?? 45))?.[0] || "";
}
/** The effects an object has, as one line for a button's hint ("影・反射"), or "" when it has none. */
export function effectsWords(o) {
  return [o?.shadow ? "影" : "", o?.reflect ? "反射" : "", o?.glow ? "光彩" : "", o?.soft ? "ぼかし" : "", o?.rot3d ? "3-D 回転" : ""].filter(Boolean).join("・");
}
