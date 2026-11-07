// 図形の結合 (PowerPoint's Merge Shapes): 接合 (union), 型抜き/合成 (combine: the overlaps cut out), 切り出し
// (fragment: every piece on its own), 重なり抽出 (intersect) and 単純型抜き (subtract the others from the first).
// Each shape's outline (its geometry, turned and flipped as drawn) becomes polygons; polygon-clipping (MIT, served
// as /vendor/polygon-clipping.js) does the boolean work; the result is a hand-drawn shape (shape: "custom") whose
// path may hold several rings (holes and separate parts, drawn even-odd). The result takes the first selected
// shape's format, text, animations and interactions, as PowerPoint does.

export const MERGE_MODES = {
  union: "接合",
  combine: "型抜き/合成",
  fragment: "切り出し",
  intersect: "重なり抽出",
  subtract: "単純型抜き",
};
export const MERGE_HINTS = {
  union: "重なりも含めて1つの図形にする",
  combine: "重なった部分をくり抜いて1つの図形にする",
  fragment: "重なりで区切られた部分をそれぞれ別の図形にする",
  intersect: "全部が重なっている部分だけを残す",
  subtract: "最初に選んだ図形から、ほかの図形と重なる部分を取り除く",
};
const MERGEABLE = new Set(["shape", "text"]);
export const canMerge = (list) => list.filter((o) => MERGEABLE.has(o.kind) && !o.locked).length >= 2;

let loading = null;
/** The polygon-clipping library (loaded once, on first use). */
export function clippingLib() {
  if (globalThis.polygonClipping) return Promise.resolve(globalThis.polygonClipping);
  loading ||= new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "/vendor/polygon-clipping.js";
    script.onload = () => (globalThis.polygonClipping ? resolve(globalThis.polygonClipping) : reject(new Error("図形の結合を読み込めません")));
    script.onerror = () => { loading = null; reject(new Error("図形の結合を読み込めません")); };
    document.head.append(script);
  });
  return loading;
}

// ---------------------------------------------------------------- SVG path → polygons

/** The closed rings of an SVG path (curves and arcs flattened to short straight steps of about `step` px). */
export function pathRings(d, step = 4) {
  const rings = [];
  let ring = null;
  let x = 0;
  let y = 0;
  let sx = 0;
  let sy = 0;
  let prevC = null; // last cubic control point (for S)
  let prevQ = null; // last quadratic control point (for T)
  const close = () => { if (ring && ring.length > 2) rings.push(ring); ring = null; };
  const lineTo = (nx, ny) => { if (!ring) ring = [[x, y]]; ring.push([nx, ny]); x = nx; y = ny; };
  const segs = (len) => Math.max(2, Math.min(96, Math.ceil(len / step)));
  const cubic = (x1, y1, x2, y2, ex, ey) => {
    const n = segs(Math.hypot(x1 - x, y1 - y) + Math.hypot(x2 - x1, y2 - y1) + Math.hypot(ex - x2, ey - y2));
    const [x0, y0] = [x, y];
    for (let i = 1; i <= n; i += 1) {
      const t = i / n;
      const u = 1 - t;
      lineTo(u * u * u * x0 + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t * t * t * ex, u * u * u * y0 + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t * t * t * ey);
    }
    prevC = [x2, y2];
  };
  const quad = (x1, y1, ex, ey) => {
    const n = segs(Math.hypot(x1 - x, y1 - y) + Math.hypot(ex - x1, ey - y1));
    const [x0, y0] = [x, y];
    for (let i = 1; i <= n; i += 1) {
      const t = i / n;
      const u = 1 - t;
      lineTo(u * u * x0 + 2 * u * t * x1 + t * t * ex, u * u * y0 + 2 * u * t * y1 + t * t * ey);
    }
    prevQ = [x1, y1];
  };
  // An elliptical arc, from the endpoint form (SVG implementation notes F.6).
  const arc = (rx, ry, rot, large, sweep, ex, ey) => {
    if (!rx || !ry) { lineTo(ex, ey); return; }
    rx = Math.abs(rx); ry = Math.abs(ry);
    const phi = (rot * Math.PI) / 180;
    const [cos, sin] = [Math.cos(phi), Math.sin(phi)];
    const dx = (x - ex) / 2;
    const dy = (y - ey) / 2;
    const x1 = cos * dx + sin * dy;
    const y1 = -sin * dx + cos * dy;
    const lambda = (x1 * x1) / (rx * rx) + (y1 * y1) / (ry * ry);
    if (lambda > 1) { rx *= Math.sqrt(lambda); ry *= Math.sqrt(lambda); }
    const num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1;
    const den = rx * rx * y1 * y1 + ry * ry * x1 * x1;
    let k = Math.sqrt(Math.max(0, num / den));
    if (large === sweep) k = -k;
    const cx1 = (k * rx * y1) / ry;
    const cy1 = (-k * ry * x1) / rx;
    const cx = cos * cx1 - sin * cy1 + (x + ex) / 2;
    const cy = sin * cx1 + cos * cy1 + (y + ey) / 2;
    const angle = (ux, uy, vx, vy) => { const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy); return a; };
    const t1 = angle(1, 0, (x1 - cx1) / rx, (y1 - cy1) / ry);
    let dt = angle((x1 - cx1) / rx, (y1 - cy1) / ry, (-x1 - cx1) / rx, (-y1 - cy1) / ry);
    if (!sweep && dt > 0) dt -= 2 * Math.PI;
    if (sweep && dt < 0) dt += 2 * Math.PI;
    const n = segs(Math.abs(dt) * Math.max(rx, ry));
    for (let i = 1; i <= n; i += 1) {
      const t = t1 + (dt * i) / n;
      const px = rx * Math.cos(t);
      const py = ry * Math.sin(t);
      lineTo(i === n ? ex : cos * px - sin * py + cx, i === n ? ey : sin * px + cos * py + cy);
    }
  };
  const tokens = String(d).match(/[MmLlHhVvCcSsQqTtAaZz]|-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?/g) || [];
  let i = 0;
  let cmd = "";
  const nextNum = () => Number(tokens[i++]);
  const hasNum = () => i < tokens.length && !/^[A-Za-z]$/.test(tokens[i]);
  while (i < tokens.length) {
    if (/^[A-Za-z]$/.test(tokens[i])) cmd = tokens[i++];
    else if (!cmd) { i += 1; continue; }
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();
    const ox = rel ? x : 0;
    const oy = rel ? y : 0;
    if (C === "Z") { if (ring) { x = sx; y = sy; } close(); prevC = prevQ = null; continue; }
    if (!hasNum()) { i += 1; continue; }
    if (C === "M") {
      close();
      x = nextNum() + ox; y = nextNum() + oy;
      sx = x; sy = y;
      ring = [[x, y]];
      cmd = rel ? "l" : "L"; // further pairs are lines
      prevC = prevQ = null;
      continue;
    }
    if (C === "L") { lineTo(nextNum() + ox, nextNum() + oy); prevC = prevQ = null; continue; }
    if (C === "H") { lineTo(nextNum() + ox, y); prevC = prevQ = null; continue; }
    if (C === "V") { lineTo(x, nextNum() + oy); prevC = prevQ = null; continue; }
    if (C === "C") { const a = [nextNum() + ox, nextNum() + oy, nextNum() + ox, nextNum() + oy, nextNum() + ox, nextNum() + oy]; cubic(...a); prevQ = null; continue; }
    if (C === "S") { const r = prevC ? [2 * x - prevC[0], 2 * y - prevC[1]] : [x, y]; const a = [nextNum() + ox, nextNum() + oy, nextNum() + ox, nextNum() + oy]; cubic(r[0], r[1], ...a); prevQ = null; continue; }
    if (C === "Q") { const a = [nextNum() + ox, nextNum() + oy, nextNum() + ox, nextNum() + oy]; quad(...a); prevC = null; continue; }
    if (C === "T") { const r = prevQ ? [2 * x - prevQ[0], 2 * y - prevQ[1]] : [x, y]; quad(r[0], r[1], nextNum() + ox, nextNum() + oy); prevC = null; continue; }
    if (C === "A") { const [rx, ry, rot, large, sweep] = [nextNum(), nextNum(), nextNum(), nextNum(), nextNum()]; arc(rx, ry, rot, large ? 1 : 0, sweep ? 1 : 0, nextNum() + ox, nextNum() + oy); prevC = prevQ = null; continue; }
    i += 1;
  }
  close();
  return rings.map(dedupe).filter((r) => r.length > 2 && Math.abs(area(r)) > 0.01);
}
const dedupe = (ring) => ring.filter((p, k) => k === 0 || Math.hypot(p[0] - ring[k - 1][0], p[1] - ring[k - 1][1]) > 1e-6);
/** Signed area (positive when counter-clockwise in a y-up system). */
export const area = (ring) => ring.reduce((sum, [x1, y1], k) => { const [x2, y2] = ring[(k + 1) % ring.length]; return sum + (x1 * y2 - x2 * y1); }, 0) / 2;
const closeRing = (ring) => (ring.length && (ring[0][0] !== ring.at(-1)[0] || ring[0][1] !== ring.at(-1)[1]) ? [...ring, ring[0]] : ring);

// ---------------------------------------------------------------- a shape's region on the slide

/** The area an object covers, in slide pixels, as a polygon-clipping MultiPolygon (turns and flips applied). */
export function shapeRegion(o, E, pc) {
  const w = Math.max(1, o.w);
  const hh = Math.max(1, o.h);
  const g = E.geometry(o.shape || "rect", w, hh, o.adj, o.path);
  const rings = (g.paths || []).flatMap((d) => pathRings(d));
  if (!rings.length) return [];
  const cx = w / 2;
  const cy = hh / 2;
  const a = ((o.rot || 0) * Math.PI) / 180;
  const [cos, sin] = [Math.cos(a), Math.sin(a)];
  const place = ([px, py]) => {
    let dx = px - cx;
    let dy = py - cy;
    if (o.flipH) dx = -dx;
    if (o.flipV) dy = -dy;
    return [o.x + cx + dx * cos - dy * sin, o.y + cy + dx * sin + dy * cos];
  };
  const placed = rings.map((r) => r.map(place));
  // Even-odd shapes (a frame, a donut): every ring toggles. Non-zero: rings turning like the biggest one fill, the
  // others cut holes.
  if (g.rule === "evenodd") return placed.slice(1).reduce((acc, r) => pc.xor(acc, [[closeRing(r)]]), [[closeRing(placed[0])]]);
  const biggest = placed.reduce((best, r) => (Math.abs(area(r)) > Math.abs(area(best)) ? r : best), placed[0]);
  const sign = Math.sign(area(biggest));
  const fills = placed.filter((r) => Math.sign(area(r)) === sign);
  const holes = placed.filter((r) => Math.sign(area(r)) !== sign);
  let region = fills.slice(1).reduce((acc, r) => pc.union(acc, [[closeRing(r)]]), [[closeRing(fills[0])]]);
  for (const r of holes) region = pc.difference(region, [[closeRing(r)]]);
  return region;
}

// ---------------------------------------------------------------- merging

/** The merged region(s): one MultiPolygon, or (切り出し) a list of separate pieces. */
export function mergeRegions(regions, mode, pc) {
  const [first, ...rest] = regions;
  if (!first?.length || !rest.length) return [];
  if (mode === "union") return [pc.union(first, ...rest)];
  if (mode === "combine") return [rest.reduce((acc, r) => pc.xor(acc, r), first)];
  if (mode === "intersect") return [pc.intersection(first, ...rest)];
  if (mode === "subtract") return [pc.difference(first, ...rest)];
  if (mode === "fragment") {
    // Every region splits what is there into the part inside it and the part outside; what it adds is new.
    let pieces = [first];
    let covered = first;
    for (const r of rest) {
      const next = [];
      for (const p of pieces) {
        const inside = pc.intersection(p, r);
        const outside = pc.difference(p, r);
        if (inside.length) next.push(inside);
        if (outside.length) next.push(outside);
      }
      const fresh = pc.difference(r, covered);
      if (fresh.length) next.push(fresh);
      covered = pc.union(covered, r);
      pieces = next;
    }
    // Each separate polygon is a shape of its own.
    return pieces.flatMap((multi) => multi.map((poly) => [poly]));
  }
  return [];
}

const r4 = (v) => Math.round(v * 10000) / 10000;
/** A MultiPolygon as a hand-drawn shape: its box, and its rings as fractions of the box. */
export function regionShape(multi, base, makeId) {
  const rings = multi.flatMap((poly) => poly.map((ring) => (ring.length > 1 && ring[0][0] === ring.at(-1)[0] && ring[0][1] === ring.at(-1)[1] ? ring.slice(0, -1) : ring))).filter((r) => r.length > 2);
  if (!rings.length) return null;
  const xs = rings.flat().map((p) => p[0]);
  const ys = rings.flat().map((p) => p[1]);
  const x = Math.min(...xs);
  const y = Math.min(...ys);
  const w = Math.max(1, Math.max(...xs) - x);
  const hh = Math.max(1, Math.max(...ys) - y);
  const frac = (ring) => ring.map(([px, py]) => [r4((px - x) / w), r4((py - y) / hh)]);
  const [pts, ...parts] = rings.map(frac);
  const o = { ...base, id: makeId(), kind: "shape", shape: "custom", x: Math.round(x * 100) / 100, y: Math.round(y * 100) / 100, w: Math.round(w * 100) / 100, h: Math.round(hh * 100) / 100, path: { pts, closed: true, ...(parts.length ? { parts } : {}) } };
  for (const key of ["rot", "flipH", "flipV", "adj", "lockRatio"]) delete o[key];
  return o;
}

/**
 * Merge objects (in the order they were selected): the new objects replacing them, or null when nothing is left
 * (two shapes that do not touch have no 重なり).
 */
export function mergeObjects(chosen, mode, E, pc, makeId) {
  const shapes = chosen.filter((o) => MERGEABLE.has(o.kind) && !o.locked);
  if (shapes.length < 2 || !MERGE_MODES[mode]) return null;
  const first = E.withDefaults(shapes[0]);
  const regions = shapes.map((o) => shapeRegion(E.withDefaults(o), E, pc));
  const results = mergeRegions(regions, mode, pc).filter((m) => m.length);
  if (!results.length) return null;
  // The first selected shape's look (a text box becomes a shape with its own fill and line).
  const base = { ...shapes[0] };
  if (base.kind === "text") base.fill = first.fill === "none" ? E.OBJECT_DEFAULTS.shape.fill : first.fill;
  const out = results.map((multi, k) => regionShape(multi, k === 0 ? base : { ...base, text: undefined, action: undefined, hover: undefined, tip: undefined, loop: undefined, name: undefined }, () => (k === 0 ? shapes[0].id : makeId())));
  return out.filter(Boolean);
}
