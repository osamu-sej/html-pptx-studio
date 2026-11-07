/*
 * HTML PPTX Studio — objects placed on a slide by hand, the PowerPoint way: shapes (with text in them),
 * text boxes, pictures, lines and arrows, icons, videos and Lottie animations. Loaded after engine.js and before
 * motion.js; exported presentations carry it too, so it has no dependencies.
 *
 * A slide's objects are `slide.elements`, in drawing order (the last one is in front), measured in slide pixels
 * (1920 × 1080, PowerPoint's 16:9 slide of 13.333 × 7.5 in at 144 px per inch): a point is 2 px, a centimetre 56.69 px.
 * The AI never writes them; people do, in the studio's editor (public/editor.mjs).
 */
(function (root) {
  "use strict";
  const E = root.SlideEngine;
  if (!E) throw new Error("engine.js must load before objects.js");
  const { h, s, W, H } = E;

  const PX_PER_PT = 2;
  const PX_PER_CM = 144 / 2.54;

  // ---------------------------------------------------------------- theme colours (PowerPoint's colour picker)
  // The picker is PowerPoint's: the theme's ten colours (背景 1・テキスト 1・背景 2・テキスト 2・アクセント 1〜6), each
  // with five lighter or darker variants made the way Office makes them (HSL lightness × lumMod + lumOff), then the ten
  // standard colours. An exported deck writes a colour that is one of these as a theme colour (schemeClr with
  // lumMod/lumOff), so it follows the theme in PowerPoint too. Any other colour stays as it is.

  const SLOTS = [["bg1", "背景 1"], ["tx1", "テキスト 1"], ["bg2", "背景 2"], ["tx2", "テキスト 2"],
    ["accent1", "アクセント 1"], ["accent2", "アクセント 2"], ["accent3", "アクセント 3"], ["accent4", "アクセント 4"], ["accent5", "アクセント 5"], ["accent6", "アクセント 6"]];
  const STANDARD_COLORS = [["#c00000", "濃い赤"], ["#ff0000", "赤"], ["#ffc000", "オレンジ"], ["#ffff00", "黄"], ["#92d050", "薄い緑"],
    ["#00b050", "緑"], ["#00b0f0", "薄い青"], ["#0070c0", "青"], ["#002060", "濃い青"], ["#7030a0", "紫"]];
  // 蛍光ペンの色 (PowerPoint's highlighter).
  const HIGHLIGHTS = [["#ffff00", "黄"], ["#00ff00", "明るい緑"], ["#00ffff", "水色"], ["#ff00ff", "ピンク"], ["#0000ff", "青"], ["#ff0000", "赤"],
    ["#000080", "濃い青"], ["#008080", "青緑"], ["#008000", "緑"], ["#800080", "紫"], ["#800000", "濃い赤"], ["#808000", "濃い黄"], ["#808080", "灰色 50%"], ["#c0c0c0", "灰色 25%"]];

  const hexRgb = (hex) => [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  function rgbHsl([r, g, b]) {
    const max = Math.max(r, g, b);
    const min = Math.min(r, g, b);
    const l = (max + min) / 2;
    if (max === min) return [0, 0, l];
    const d = max - min;
    const sat = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    const hue = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [hue / 6, sat, l];
  }
  function hslHex([hue, sat, l]) {
    const f = (n) => {
      const k = (n + hue * 12) % 12;
      const a = sat * Math.min(l, 1 - l);
      return l - a * Math.max(-1, Math.min(k - 3, 9 - k, 1));
    };
    // Halves round down, as Office does (black 50 % lighter is #7f7f7f).
    return `#${[0, 8, 4].map((n) => Math.round(Math.max(0, Math.min(1, f(n))) * 255 - 1e-6).toString(16).padStart(2, "0")).join("")}`;
  }
  /** A colour with PowerPoint's lumMod / lumOff (in 1/1000 %, as the file stores them: 75000 = 75 %). */
  function lumColor(hex, mod = 100000, off = 0) {
    const [hue, sat, l] = rgbHsl(hexRgb(hex));
    return hslHex([hue, sat, Math.max(0, Math.min(1, l * (mod / 100000) + off / 100000))]);
  }
  /** The five variants under a theme colour, as Office picks them from the colour's lightness. */
  function lumVariants(hex) {
    const l = rgbHsl(hexRgb(hex))[2];
    const lighter = (pairs) => pairs.map(([pct]) => [100000 - pct * 1000, pct * 1000, `明るく ${pct}%`]);
    const darker = (pcts) => pcts.map((pct) => [100000 - pct * 1000, 0, `暗く ${pct}%`]);
    if (l <= 0.001) return lighter([[50], [35], [25], [15], [5]]);
    if (l >= 0.999) return darker([5, 15, 25, 35, 50]);
    if (l < 0.2) return lighter([[90], [75], [50], [25], [10]]);
    if (l > 0.8) return darker([10, 25, 50, 75, 90]);
    return [...lighter([[80], [60], [40]]), ...darker([25, 50])];
  }
  /** WCAG contrast of two #rrggbb colours (1 to 21). */
  function contrastRatio(a, b) {
    const lum = (hex) => {
      const [r, g, bl] = hexRgb(hex).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
      return 0.2126 * r + 0.7152 * g + 0.0722 * bl;
    };
    const [x, y] = [lum(a), lum(b)].sort((m, n) => n - m);
    return (x + 0.05) / (y + 0.05);
  }
  const THEME_BY_ID = new Map((E.THEMES || []).map((theme) => [theme.id, theme]));
  /** A theme's ten colours by slot (bg1 … accent6), plus its hyperlink colours. */
  function themeColors(themeId) {
    const c = (THEME_BY_ID.get(themeId) || THEME_BY_ID.get(E.DEFAULT_THEME) || {}).colors || {};
    const out = { bg1: c.bg1 || "#ffffff", tx1: c.tx1 || "#000000", bg2: c.bg2 || "#e7e6e6", tx2: c.tx2 || "#44546a", hlink: c.hlink || "#0563c1", folHlink: c.folHlink || "#954f72" };
    (c.accents || []).forEach((v, i) => { out[`accent${i + 1}`] = v; });
    for (let i = 1; i <= 6; i += 1) out[`accent${i}`] ||= "#808080";
    return out;
  }
  /** PowerPoint's colour sheet for a theme: ten columns (a theme colour and its five variants) and the standard row. */
  function paletteSheet(themeId) {
    const colors = themeColors(themeId);
    const columns = SLOTS.map(([slot, label]) => {
      const hex = colors[slot].toLowerCase();
      return { slot, label, hex, variants: lumVariants(hex).map(([mod, off, how]) => ({ hex: lumColor(hex, mod, off), mod, off, label: `${label}、${how}` })) };
    });
    return { themeId, columns, standard: STANDARD_COLORS };
  }
  /**
   * Where a colour sits in a theme's sheet: { slot, mod, off } (mod/off only when it is a variant), else null.
   * The exporter writes such a colour as a theme colour.
   */
  function themeColorRef(hex, themeId) {
    const want = String(hex || "").toLowerCase();
    if (!/^#[0-9a-f]{6}$/.test(want)) return null;
    for (const col of paletteSheet(themeId).columns) {
      if (col.hex === want) return { slot: col.slot };
      const variant = col.variants.find((v) => v.hex === want);
      if (variant) return { slot: col.slot, mod: variant.mod, off: variant.off };
    }
    return null;
  }

  /**
   * テーマの変更: colours that are theme colours of `from` (a slot or one of its variants) become the same place in `to`,
   * wherever a deck keeps colours (objects and their rich text, tables, backgrounds, animations' colours, the master's
   * objects, the default looks), as PowerPoint recolours theme colours. Other colours stay. The deck changes in place.
   */
  function retheme(deck, from, to) {
    if (!deck || from === to) return deck;
    const a = paletteSheet(from).columns;
    const b = paletteSheet(to).columns;
    const map = new Map();
    a.forEach((col, i) => { if (!map.has(col.hex)) map.set(col.hex, b[i].hex); });
    a.forEach((col, i) => col.variants.forEach((v, j) => { if (!map.has(v.hex)) map.set(v.hex, b[i].variants[j].hex); }));
    const swap = (value) => (value == null ? value : JSON.parse(JSON.stringify(value).replace(/#[0-9a-fA-F]{6}\b/g, (hex) => map.get(hex.toLowerCase()) || hex)));
    for (const slide of deck.slides || []) {
      for (const key of ["elements", "background", "timeline"]) if (slide && slide[key]) slide[key] = swap(slide[key]);
    }
    for (const key of ["masterObjects", "objectDefaults"]) if (deck[key]) deck[key] = swap(deck[key]);
    return deck;
  }

  // The lists the editor offers (and pastes to with 貼り付け先のスタイルを使用): the current theme's colours, their
  // variants and the standard colours. useTheme() switches them when a deck with another theme is opened.
  const PALETTE = { text: [], fill: [], line: [], highlight: HIGHLIGHTS, sheet: null };
  function useTheme(themeId) {
    const sheet = paletteSheet(themeId);
    const flat = [...sheet.columns.map((col) => [col.hex, col.label]), ...sheet.columns.flatMap((col) => col.variants.map((v) => [v.hex, v.label])), ...STANDARD_COLORS];
    const unique = [...new Map(flat.map(([c, label]) => [c, label])).entries()];
    PALETTE.sheet = sheet;
    PALETTE.text = unique;
    PALETTE.fill = unique;
    PALETTE.line = unique;
    themeDefaults(themeId);
    return PALETTE;
  }
  // Fonts an object may use: the theme's two (見出し・本文, written as PowerPoint's +mj / +mn theme fonts) and fonts
  // that come with Windows and macOS (so a PowerPoint opens the same on either). [label, CSS stack, PowerPoint name].
  const FONTS = {
    body: ["本文のフォント（テーマ）", "var(--font-body)", "+mn-ea"],
    heading: ["見出しのフォント（テーマ）", "var(--font-display)", "+mj-ea"],
    yugothic: ["游ゴシック", '"Yu Gothic", YuGothic, "游ゴシック", "游ゴシック体", sans-serif', "游ゴシック"],
    meiryo: ["メイリオ", 'Meiryo, "メイリオ", "Hiragino Sans", sans-serif', "メイリオ"],
    ud: ["BIZ UDPゴシック", '"BIZ UDPGothic", "BIZ UDPゴシック", "Yu Gothic", sans-serif', "BIZ UDPゴシック"],
    mincho: ["游明朝", '"Yu Mincho", YuMincho, "游明朝", "游明朝体", serif', "游明朝"],
    kyokasho: ["UD デジタル 教科書体", '"UD デジタル 教科書体 NP-R", "UD Digi Kyokasho NP-R", "Klee One", "BIZ UDPGothic", serif', "UD デジタル 教科書体 NP-R"],
  };

  // ---------------------------------------------------------------- geometry helpers

  const r2 = (n) => Math.round(n * 100) / 100;
  const P = (x, y) => `${r2(x)} ${r2(y)}`;
  const poly = (points) => `M${points.map(([x, y]) => P(x, y)).join(" L")} Z`;
  const rad = (deg) => (deg * Math.PI) / 180;
  const deg = (r) => (r * 180) / Math.PI;
  const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
  const mod = (a, n) => ((a % n) + n) % n;
  const on = (cx, cy, rx, ry, a) => [cx + rx * Math.cos(rad(a)), cy + ry * Math.sin(rad(a))];

  /** An elliptical arc to angle a1 from a0 (degrees, increasing = clockwise on screen), from the current point. */
  function arcTo(cx, cy, rx, ry, a0, a1) {
    const sweep = a1 >= a0 ? 1 : 0;
    const span = Math.abs(a1 - a0);
    if (span >= 359.99) {
      const mid = on(cx, cy, rx, ry, a0 + (a1 - a0) / 2);
      const end = on(cx, cy, rx, ry, a1);
      return `A${r2(rx)} ${r2(ry)} 0 0 ${sweep} ${P(...mid)} A${r2(rx)} ${r2(ry)} 0 0 ${sweep} ${P(...end)}`;
    }
    return `A${r2(rx)} ${r2(ry)} 0 ${span > 180 ? 1 : 0} ${sweep} ${P(...on(cx, cy, rx, ry, a1))}`;
  }
  /** The clockwise end angle for a sweep from a0 to a1 (a1 after a0, going clockwise). */
  const cw = (a0, a1) => a0 + (mod(a1 - a0, 360) || 360);
  function ellipse(cx, cy, rx, ry) {
    return `M${P(...on(cx, cy, rx, ry, 0))} ${arcTo(cx, cy, rx, ry, 0, 360)} Z`;
  }
  /** A rectangle with its own radius at each corner: [top-left, top-right, bottom-right, bottom-left]. */
  function rrect(x, y, w, hh, radii) {
    const [a, b, c, d] = radii.map((r) => clamp(r, 0, Math.min(w, hh) / 2));
    const arc = (r, tx, ty) => (r ? `A${r2(r)} ${r2(r)} 0 0 1 ${P(tx, ty)}` : `L${P(tx, ty)}`);
    return `M${P(x + a, y)} L${P(x + w - b, y)} ${arc(b, x + w, y + b)} L${P(x + w, y + hh - c)} ${arc(c, x + w - c, y + hh)} L${P(x + d, y + hh)} ${arc(d, x, y + hh - d)} L${P(x, y + a)} ${arc(a, x + a, y)} Z`;
  }
  /** Points of a polygon given in unit space, stretched to fill the box (PowerPoint's regular shapes touch every edge). */
  function fitted(points, w, hh) {
    const xs = points.map((p) => p[0]);
    const ys = points.map((p) => p[1]);
    const [x0, x1, y0, y1] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
    return points.map(([x, y]) => [((x - x0) / (x1 - x0 || 1)) * w, ((y - y0) / (y1 - y0 || 1)) * hh]);
  }
  const regular = (n, w, hh) => fitted(Array.from({ length: n }, (_, k) => on(0, 0, 1, 1, -90 + (k * 360) / n)), w, hh);
  function starUnit(n, inner) {
    const pts = [];
    for (let k = 0; k < n; k += 1) {
      pts.push(on(0, 0, 1, 1, -90 + (k * 360) / n), on(0, 0, inner, inner, -90 + ((k + 0.5) * 360) / n));
    }
    return pts;
  }
  /** Maps unit-space points into the box the same way fitted() maps the outer points. */
  function fitTransform(outer, w, hh) {
    const xs = outer.map((p) => p[0]);
    const ys = outer.map((p) => p[1]);
    const [x0, x1, y0, y1] = [Math.min(...xs), Math.max(...xs), Math.min(...ys), Math.max(...ys)];
    return {
      to: ([x, y]) => [((x - x0) / (x1 - x0 || 1)) * w, ((y - y0) / (y1 - y0 || 1)) * hh],
      from: ([x, y]) => [x0 + (x / (w || 1)) * (x1 - x0), y0 + (y / (hh || 1)) * (y1 - y0)],
    };
  }
  const scale21600 = (points, w, hh) => points.map(([x, y]) => [(x / 21600) * w, (y / 21600) * hh]);
  const flipX = (points, w) => points.map(([x, y]) => [w - x, y]);
  const flipY = (points, hh) => points.map(([x, y]) => [x, hh - y]);
  const swap = (points) => points.map(([x, y]) => [y, x]);

  // ---------------------------------------------------------------- shapes
  // Each shape: label (PowerPoint's name), d(w, h, adj) → an SVG path (or several, drawn in order), its
  // adjustments (adj: defaults, lim: ranges) with handles to drag them, where text goes (text → [l, t, r, b]),
  // extra details (shaded faces, lines) and whether it is an open line (no fill).

  const ss = (w, hh) => Math.min(w, hh);
  const handleX = (pos, set) => ({ pos, set });

  function rightArrow(w, hh, t, hl) {
    const y1 = hh / 2 - t / 2;
    const y2 = hh / 2 + t / 2;
    const x = w - hl;
    return [[0, y1], [x, y1], [x, 0], [w, hh / 2], [x, hh], [x, y2], [0, y2]];
  }
  /** A horizontally defined block arrow turned to any of the four directions. */
  function turned(dir, w, hh, build) {
    if (dir === "right") return build(w, hh);
    if (dir === "left") return flipX(build(w, hh), w);
    if (dir === "down") return swap(build(hh, w));
    return flipY(swap(build(hh, w)), hh);
  }
  const arrowShape = (label, dir) => ({
    label, adj: [0.5, 0.5], lim: [[0.05, 1], [0, 2]],
    d: (w, hh, [a, b]) => poly(turned(dir, w, hh, (ww, hhh) => rightArrow(ww, hhh, a * hhh, Math.min(b * ss(w, hh), ww)))),
    handles: [
      handleX((w, hh, [a, b]) => turned(dir, w, hh, (ww, hhh) => [[ww - Math.min(b * ss(w, hh), ww), 0]])[0],
        (x, y, w, hh, [a]) => [a, (dir === "right" ? w - x : dir === "left" ? x : dir === "down" ? hh - y : y) / ss(w, hh)]),
      handleX((w, hh, [a]) => turned(dir, w, hh, (ww, hhh) => [[0, hhh / 2 - (a * hhh) / 2]])[0],
        (x, y, w, hh, [, b]) => [dir === "right" || dir === "left" ? (hh - 2 * y) / hh : (w - 2 * x) / w, b]),
    ],
    text: (w, hh, [a, b]) => {
      const box = turned(dir, w, hh, (ww, hhh) => [[0, hhh / 2 - (a * hhh) / 2], [ww - Math.min(b * ss(w, hh), ww) * 0.5, hhh / 2 + (a * hhh) / 2]]);
      return [Math.min(box[0][0], box[1][0]), Math.min(box[0][1], box[1][1]), Math.max(box[0][0], box[1][0]), Math.max(box[0][1], box[1][1])];
    },
  });

  function doubleArrow(w, hh, t, hl) {
    hl = Math.min(hl, w / 2);
    const y1 = hh / 2 - t / 2;
    const y2 = hh / 2 + t / 2;
    return [[0, hh / 2], [hl, 0], [hl, y1], [w - hl, y1], [w - hl, 0], [w, hh / 2], [w - hl, hh], [w - hl, y2], [hl, y2], [hl, hh]];
  }

  function noSmoking(w, hh, t) {
    const [cx, cy, rx, ry] = [w / 2, hh / 2, w / 2, hh / 2];
    const r = clamp(1 - 2 * t, 0.05, 0.95);
    const b = Math.min(t, r * 0.9);
    const k = Math.sqrt(r * r - b * b);
    const n = [Math.SQRT1_2, -Math.SQRT1_2];
    const u = [Math.SQRT1_2, Math.SQRT1_2];
    const holes = [1, -1].map((side) => {
      const p1 = [side * b * n[0] + k * u[0], side * b * n[1] + k * u[1]];
      const p2 = [side * b * n[0] - k * u[0], side * b * n[1] - k * u[1]];
      const a1 = deg(Math.atan2(p1[1], p1[0]));
      const a2 = deg(Math.atan2(p2[1], p2[0]));
      const via = side > 0 ? -45 : 135;
      // Go from p1 to p2 the way that passes the middle of the hole.
      const forward = mod(via - a1, 360) < mod(a2 - a1, 360);
      const end = forward ? a1 + mod(a2 - a1, 360) : a1 - mod(a1 - a2, 360);
      return `M${P(cx + p1[0] * rx, cy + p1[1] * ry)} ${arcTo(cx, cy, r * rx, r * ry, a1, end)} Z`;
    });
    return `${ellipse(cx, cy, rx, ry)} ${holes.join(" ")}`;
  }

  // A cloud: puffs (arcs bulging outwards) between points of a smaller ellipse, a little uneven like PowerPoint's.
  const CLOUD = [[-96, 0.5], [-58, 0.58], [-22, 0.52], [12, 0.56], [44, 0.6], [84, 0.54], [118, 0.58], [152, 0.52], [190, 0.6], [228, 0.55]];
  function cloudPath(w, hh) {
    const [cx, cy] = [w / 2, hh / 2];
    const pts = CLOUD.map(([a]) => on(cx, cy, w * 0.36, hh * 0.34, a));
    let d = `M${P(...pts[0])}`;
    for (let k = 0; k < pts.length; k += 1) {
      const [x1, y1] = pts[k];
      const [x2, y2] = pts[(k + 1) % pts.length];
      const r = Math.hypot(x2 - x1, y2 - y1) * CLOUD[k][1];
      d += ` A${r2(r)} ${r2(r)} 0 1 1 ${P(x2, y2)}`;
    }
    return `${d} Z`;
  }

  function wavePath(w, hh, a, periods) {
    const A = a * hh;
    const seg = w / periods;
    let top = `M${P(0, A)}`;
    for (let i = 0; i < periods; i += 1) {
      const x = i * seg;
      top += ` C${P(x + seg / 3, -A)} ${P(x + (2 * seg) / 3, 3 * A)} ${P(x + seg, A)}`;
    }
    let bottom = ` L${P(w, hh - A)}`;
    for (let i = periods - 1; i >= 0; i -= 1) {
      const x = i * seg;
      bottom += ` C${P(x + (2 * seg) / 3, hh + A)} ${P(x + seg / 3, hh - 3 * A)} ${P(x, hh - A)}`;
    }
    return `${top}${bottom} Z`;
  }

  /** Callouts: a box with a tail whose tip is at (w/2 + ax·w, h/2 + ay·h). */
  function wedgeTip(w, hh, [ax, ay]) { return [w / 2 + ax * w, hh / 2 + ay * hh]; }
  function wedgeRect(w, hh, adj, r = 0) {
    const [tx, ty] = wedgeTip(w, hh, adj);
    const inside = tx > 0 && tx < w && ty > 0 && ty < hh;
    if (inside) return r ? rrect(0, 0, w, hh, [r, r, r, r]) : poly([[0, 0], [w, 0], [w, hh], [0, hh]]);
    const vertical = Math.abs(ty - hh / 2) / hh >= Math.abs(tx - w / 2) / w;
    const bw = (vertical ? w : hh) * 0.12;
    if (vertical) {
      const base = clamp(w / 2 + (tx - w / 2) * 0.5, r + bw, w - r - bw);
      const y = ty > hh / 2 ? hh : 0;
      const pts = y ? [[base + bw, hh], [tx, ty], [base - bw, hh]] : [[base - bw, 0], [tx, ty], [base + bw, 0]];
      return y
        ? `M${P(r, 0)} L${P(w - r, 0)} ${r ? `A${r} ${r} 0 0 1 ${P(w, r)}` : ""} L${P(w, hh - r)} ${r ? `A${r} ${r} 0 0 1 ${P(w - r, hh)}` : ""} L${pts.map(([x, yy]) => P(x, yy)).join(" L")} L${P(r, hh)} ${r ? `A${r} ${r} 0 0 1 ${P(0, hh - r)}` : ""} L${P(0, r)} ${r ? `A${r} ${r} 0 0 1 ${P(r, 0)}` : ""} Z`
        : `M${P(r, 0)} L${pts.map(([x, yy]) => P(x, yy)).join(" L")} L${P(w - r, 0)} ${r ? `A${r} ${r} 0 0 1 ${P(w, r)}` : ""} L${P(w, hh - r)} ${r ? `A${r} ${r} 0 0 1 ${P(w - r, hh)}` : ""} L${P(r, hh)} ${r ? `A${r} ${r} 0 0 1 ${P(0, hh - r)}` : ""} L${P(0, r)} ${r ? `A${r} ${r} 0 0 1 ${P(r, 0)}` : ""} Z`;
    }
    const base = clamp(hh / 2 + (ty - hh / 2) * 0.5, r + bw, hh - r - bw);
    const x = tx > w / 2 ? w : 0;
    return x
      ? `M${P(r, 0)} L${P(w - r, 0)} ${r ? `A${r} ${r} 0 0 1 ${P(w, r)}` : ""} L${P(w, base - bw)} L${P(tx, ty)} L${P(w, base + bw)} L${P(w, hh - r)} ${r ? `A${r} ${r} 0 0 1 ${P(w - r, hh)}` : ""} L${P(r, hh)} ${r ? `A${r} ${r} 0 0 1 ${P(0, hh - r)}` : ""} L${P(0, r)} ${r ? `A${r} ${r} 0 0 1 ${P(r, 0)}` : ""} Z`
      : `M${P(r, 0)} L${P(w - r, 0)} ${r ? `A${r} ${r} 0 0 1 ${P(w, r)}` : ""} L${P(w, hh - r)} ${r ? `A${r} ${r} 0 0 1 ${P(w - r, hh)}` : ""} L${P(r, hh)} ${r ? `A${r} ${r} 0 0 1 ${P(0, hh - r)}` : ""} L${P(0, base + bw)} L${P(tx, ty)} L${P(0, base - bw)} L${P(0, r)} ${r ? `A${r} ${r} 0 0 1 ${P(r, 0)}` : ""} Z`;
  }
  const tailHandle = handleX((w, hh, adj) => wedgeTip(w, hh, adj), (x, y, w, hh) => [(x - w / 2) / w, (y - hh / 2) / hh]);

  const EXPLOSION1 = [[10800, 5800], [14522, 0], [14155, 5325], [18380, 4457], [16702, 7315], [21097, 8137], [17607, 10475], [21600, 13290], [16837, 12942], [18145, 18095], [14020, 14457], [13247, 19737], [10532, 14935], [8485, 21600], [7715, 15627], [4762, 17617], [5667, 13937], [135, 14587], [3722, 11775], [0, 8615], [4627, 7617], [370, 2295], [7312, 6320], [8352, 2295]];
  const EXPLOSION2 = [[11462, 4342], [14790, 0], [14525, 5777], [18007, 3172], [16380, 6532], [21600, 6645], [16985, 9402], [18270, 11290], [16380, 12310], [18877, 15632], [14640, 14350], [14942, 17370], [12180, 15935], [11612, 18842], [9872, 17370], [8700, 19712], [7527, 18125], [4917, 21600], [4805, 18240], [1285, 17825], [3330, 15370], [0, 12877], [3935, 11592], [1172, 8270], [5372, 7817], [4502, 3625], [8550, 6382], [9722, 1887]];
  const LIGHTNING = [[8458, 0], [12079, 5596], [10838, 6022], [15468, 10627], [14280, 11148], [21600, 21600], [10900, 13190], [12124, 12620], [5976, 7955], [7172, 7451], [0, 3890]];

  function starShape(label, n, inner) {
    return {
      label, adj: [inner], lim: [[0.05, 0.95]],
      d: (w, hh, [a]) => {
        const pts = starUnit(n, a);
        const fit = fitTransform(pts.filter((_, i) => i % 2 === 0), w, hh);
        return poly(pts.map(fit.to));
      },
      handles: [handleX((w, hh, [a]) => {
        const pts = starUnit(n, a);
        return fitTransform(pts.filter((_, i) => i % 2 === 0), w, hh).to(pts[1]);
      }, (x, y, w, hh, [a]) => {
        const pts = starUnit(n, a);
        const [ux, uy] = fitTransform(pts.filter((_, i) => i % 2 === 0), w, hh).from([x, y]);
        return [clamp(Math.hypot(ux, uy), 0.05, 0.95)];
      })],
      text: (w, hh, [a]) => [w * (0.5 - a * 0.36), hh * (0.5 - a * 0.36), w * (0.5 + a * 0.36), hh * (0.5 + a * 0.36)],
    };
  }

  // 動作設定ボタン (PowerPoint's action buttons): a square with its sign drawn darker; the ones that move through the
  // show come with that click already set (`action`).
  const ACTION_BUTTONS = (() => {
    const sign = (draw) => (w, hh) => {
      const s0 = Math.min(w, hh) * 0.62;
      const x0 = (w - s0) / 2;
      const y0 = (hh - s0) / 2;
      const q = (x, y) => [x0 + x * s0, y0 + y * s0];
      return draw(q, s0, x0, y0);
    };
    const shapeOf = (points) => (q) => [{ d: poly(points.map(([x, y]) => q(x, y))), tone: "dark" }];
    const line = (q, pts) => `M${pts.map(([x, y]) => P(...q(x, y))).join(" L")}`;
    const ring = (q, s0) => { const [cx, cy] = q(0.5, 0.5); return ellipse(cx, cy, s0 * 0.48, s0 * 0.48); };
    const make = (label, extras, action = null) => ({ label: `動作設定ボタン: ${label}`, d: (w, hh) => poly([[0, 0], [w, 0], [w, hh], [0, hh]]), extras: sign(extras), ...(action ? { action } : {}) });
    return {
      actionButtonBackPrevious: make("戻る/前へ", shapeOf([[0.05, 0.5], [0.95, 0.05], [0.95, 0.95]]), { type: "prev" }),
      actionButtonForwardNext: make("進む/次へ", shapeOf([[0.95, 0.5], [0.05, 0.05], [0.05, 0.95]]), { type: "next" }),
      actionButtonBeginning: make("最初へ", (q) => [{ d: `${poly([[0.25, 0.5], [0.95, 0.05], [0.95, 0.95]].map(([x, y]) => q(x, y)))} ${poly([[0.05, 0.05], [0.18, 0.05], [0.18, 0.95], [0.05, 0.95]].map(([x, y]) => q(x, y)))}`, tone: "dark" }], { type: "first" }),
      actionButtonEnd: make("最後へ", (q) => [{ d: `${poly([[0.75, 0.5], [0.05, 0.05], [0.05, 0.95]].map(([x, y]) => q(x, y)))} ${poly([[0.82, 0.05], [0.95, 0.05], [0.95, 0.95], [0.82, 0.95]].map(([x, y]) => q(x, y)))}`, tone: "dark" }], { type: "last" }),
      actionButtonHome: make("ホーム", shapeOf([[0.5, 0.05], [0.95, 0.45], [0.82, 0.45], [0.82, 0.95], [0.18, 0.95], [0.18, 0.45], [0.05, 0.45]]), { type: "first" }),
      actionButtonInformation: make("情報", (q, s0) => { const [cx, cy] = q(0.5, 0.24); return [{ d: ring(q, s0), tone: "dark" }, { d: `${ellipse(cx, cy, s0 * 0.07, s0 * 0.07)} ${poly([[0.43, 0.4], [0.57, 0.4], [0.57, 0.82], [0.43, 0.82]].map(([x, y]) => q(x, y)))}`, tone: "light" }]; }),
      actionButtonReturn: make("戻る", (q) => [{ d: line(q, [[0.8, 0.15], [0.8, 0.6], [0.6, 0.8], [0.3, 0.8], [0.15, 0.62]]), tone: "line" }, { d: poly([[0.02, 0.62], [0.28, 0.48], [0.28, 0.76]].map(([x, y]) => q(x, y))), tone: "dark" }], { type: "prev" }),
      actionButtonMovie: make("ビデオ", (q) => [{ d: `${poly([[0.05, 0.25], [0.62, 0.25], [0.62, 0.75], [0.05, 0.75]].map(([x, y]) => q(x, y)))} ${poly([[0.62, 0.45], [0.95, 0.25], [0.95, 0.75], [0.62, 0.55]].map(([x, y]) => q(x, y)))}`, tone: "dark" }]),
      actionButtonDocument: make("ドキュメント", (q) => [{ d: poly([[0.18, 0.05], [0.62, 0.05], [0.82, 0.25], [0.82, 0.95], [0.18, 0.95]].map(([x, y]) => q(x, y))), tone: "dark" }, { d: line(q, [[0.62, 0.05], [0.62, 0.25], [0.82, 0.25]]), tone: "line" }]),
      actionButtonSound: make("サウンド", (q) => [{ d: poly([[0.05, 0.35], [0.3, 0.35], [0.58, 0.08], [0.58, 0.92], [0.3, 0.65], [0.05, 0.65]].map(([x, y]) => q(x, y))), tone: "dark" }, { d: `${line(q, [[0.72, 0.3], [0.8, 0.5], [0.72, 0.7]])} ${line(q, [[0.84, 0.15], [0.97, 0.5], [0.84, 0.85]])}`, tone: "line" }]),
      actionButtonHelp: make("ヘルプ", (q, s0) => { const [cx, cy] = q(0.5, 0.88); return [{ d: `${line(q, [[0.3, 0.3], [0.38, 0.12], [0.62, 0.12], [0.7, 0.3], [0.5, 0.48], [0.5, 0.66]])}`, tone: "line" }, { d: ellipse(cx, cy, s0 * 0.07, s0 * 0.07), tone: "dark" }]; }),
      actionButtonBlank: make("空白", () => []),
    };
  })();

  const SHAPES = {
    // ---- 四角形
    rect: { label: "正方形/長方形", d: (w, hh) => poly([[0, 0], [w, 0], [w, hh], [0, hh]]) },
    roundRect: {
      label: "四角形: 角を丸くする", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => rrect(0, 0, w, hh, Array(4).fill(a * ss(w, hh))),
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), 0], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const i = a * ss(w, hh) * 0.29; return [i, i, w - i, hh - i]; },
    },
    snip1Rect: {
      label: "四角形: 1つの角を切り取る", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const c = a * ss(w, hh); return poly([[0, 0], [w - c, 0], [w, c], [w, hh], [0, hh]]); },
      handles: [handleX((w, hh, [a]) => [w - a * ss(w, hh), 0], (x, y, w, hh) => [(w - x) / ss(w, hh)])],
    },
    snip2SameRect: {
      label: "四角形: 上の2つの角を切り取る", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const c = a * ss(w, hh); return poly([[c, 0], [w - c, 0], [w, c], [w, hh], [0, hh], [0, c]]); },
      handles: [handleX((w, hh, [a]) => [w - a * ss(w, hh), 0], (x, y, w, hh) => [(w - x) / ss(w, hh)])],
    },
    snip2DiagRect: {
      label: "四角形: 対角を切り取る", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const c = a * ss(w, hh); return poly([[c, 0], [w, 0], [w, hh - c], [w - c, hh], [0, hh], [0, c]]); },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), 0], (x, y, w, hh) => [x / ss(w, hh)])],
    },
    round1Rect: {
      label: "四角形: 1つの角を丸める", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => rrect(0, 0, w, hh, [0, a * ss(w, hh), 0, 0]),
      handles: [handleX((w, hh, [a]) => [w - a * ss(w, hh), 0], (x, y, w, hh) => [(w - x) / ss(w, hh)])],
    },
    round2SameRect: {
      label: "四角形: 上の2つの角を丸める", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const r = a * ss(w, hh); return rrect(0, 0, w, hh, [r, r, 0, 0]); },
      handles: [handleX((w, hh, [a]) => [w - a * ss(w, hh), 0], (x, y, w, hh) => [(w - x) / ss(w, hh)])],
    },
    round2DiagRect: {
      label: "四角形: 対角を丸める", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const r = a * ss(w, hh); return rrect(0, 0, w, hh, [r, 0, r, 0]); },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), 0], (x, y, w, hh) => [x / ss(w, hh)])],
    },

    // ---- 基本図形
    ellipse: { label: "楕円", d: (w, hh) => ellipse(w / 2, hh / 2, w / 2, hh / 2), text: (w, hh) => [w * 0.146, hh * 0.146, w * 0.854, hh * 0.854] },
    triangle: {
      label: "二等辺三角形", adj: [0.5], lim: [[0, 1]],
      d: (w, hh, [a]) => poly([[a * w, 0], [w, hh], [0, hh]]),
      handles: [handleX((w, hh, [a]) => [a * w, 0], (x, y, w) => [x / w])],
      text: (w, hh, [a]) => [(a * w) / 2, hh / 2, (a * w + w) / 2, hh],
    },
    rtTriangle: { label: "直角三角形", d: (w, hh) => poly([[0, 0], [w, hh], [0, hh]]), text: (w, hh) => [w / 12, (7 * hh) / 12, (7 * w) / 12, (11 * hh) / 12] },
    parallelogram: {
      label: "平行四辺形", adj: [0.25], lim: [[0, 1]],
      d: (w, hh, [a]) => { const o = Math.min(a * ss(w, hh), w); return poly([[o, 0], [w, 0], [w - o, hh], [0, hh]]); },
      handles: [handleX((w, hh, [a]) => [Math.min(a * ss(w, hh), w), 0], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const o = Math.min(a * ss(w, hh), w) * 0.6; return [o, 0, w - o, hh]; },
    },
    trapezoid: {
      label: "台形", adj: [0.25], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const o = Math.min(a * ss(w, hh), w / 2); return poly([[o, 0], [w - o, 0], [w, hh], [0, hh]]); },
      handles: [handleX((w, hh, [a]) => [Math.min(a * ss(w, hh), w / 2), 0], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const o = Math.min(a * ss(w, hh), w / 2) * 0.66; return [o, 0, w - o, hh]; },
    },
    diamond: { label: "ひし形", d: (w, hh) => poly([[w / 2, 0], [w, hh / 2], [w / 2, hh], [0, hh / 2]]), text: (w, hh) => [w / 4, hh / 4, (3 * w) / 4, (3 * hh) / 4] },
    pentagon: { label: "五角形", d: (w, hh) => poly(regular(5, w, hh)), text: (w, hh) => [w * 0.19, hh * 0.32, w * 0.81, hh * 0.92] },
    hexagon: {
      label: "六角形", adj: [0.25], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const o = Math.min(a * ss(w, hh), w / 2); return poly([[o, 0], [w - o, 0], [w, hh / 2], [w - o, hh], [o, hh], [0, hh / 2]]); },
      handles: [handleX((w, hh, [a]) => [Math.min(a * ss(w, hh), w / 2), 0], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const o = Math.min(a * ss(w, hh), w / 2) * 0.6; return [o, hh * 0.1, w - o, hh * 0.9]; },
    },
    heptagon: { label: "七角形", d: (w, hh) => poly(regular(7, w, hh)), text: (w, hh) => [w * 0.15, hh * 0.2, w * 0.85, hh * 0.85] },
    octagon: {
      label: "八角形", adj: [0.2929], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const c = a * ss(w, hh); return poly([[c, 0], [w - c, 0], [w, c], [w, hh - c], [w - c, hh], [c, hh], [0, hh - c], [0, c]]); },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), 0], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const c = (a * ss(w, hh)) / 2; return [c, c, w - c, hh - c]; },
    },
    decagon: { label: "十角形", d: (w, hh) => poly(regular(10, w, hh)), text: (w, hh) => [w * 0.1, hh * 0.1, w * 0.9, hh * 0.9] },
    dodecagon: { label: "十二角形", d: (w, hh) => poly(regular(12, w, hh)), text: (w, hh) => [w * 0.1, hh * 0.1, w * 0.9, hh * 0.9] },
    pie: {
      label: "部分円", adj: [0, 270], lim: [[-360, 720], [-360, 720]],
      d: (w, hh, [a0, a1]) => { const [cx, cy] = [w / 2, hh / 2]; return `M${P(cx, cy)} L${P(...on(cx, cy, cx, cy, a0))} ${arcTo(cx, cy, cx, cy, a0, cw(a0, a1))} Z`; },
      handles: [
        handleX((w, hh, [a0]) => on(w / 2, hh / 2, w / 2, hh / 2, a0), (x, y, w, hh, [, a1]) => [deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2))), a1]),
        handleX((w, hh, [, a1]) => on(w / 2, hh / 2, w / 2, hh / 2, a1), (x, y, w, hh, [a0]) => [a0, deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2)))]),
      ],
    },
    chord: {
      label: "弦", adj: [45, 270], lim: [[-360, 720], [-360, 720]],
      d: (w, hh, [a0, a1]) => { const [cx, cy] = [w / 2, hh / 2]; return `M${P(...on(cx, cy, cx, cy, a0))} ${arcTo(cx, cy, cx, cy, a0, cw(a0, a1))} Z`; },
      handles: [
        handleX((w, hh, [a0]) => on(w / 2, hh / 2, w / 2, hh / 2, a0), (x, y, w, hh, [, a1]) => [deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2))), a1]),
        handleX((w, hh, [, a1]) => on(w / 2, hh / 2, w / 2, hh / 2, a1), (x, y, w, hh, [a0]) => [a0, deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2)))]),
      ],
    },
    teardrop: {
      label: "涙形", adj: [1], lim: [[0, 2]],
      d: (w, hh, [a]) => {
        const [cx, cy] = [w / 2, hh / 2];
        const tip = [cx + (cx * Math.max(a, Math.SQRT1_2)), cy - (cy * Math.max(a, Math.SQRT1_2))];
        return `M${P(0, cy)} ${arcTo(cx, cy, cx, cy, 180, 270)} L${P(...tip)} L${P(w, cy)} ${arcTo(cx, cy, cx, cy, 0, 90)} ${arcTo(cx, cy, cx, cy, 90, 180)} Z`;
      },
      handles: [handleX((w, hh, [a]) => [w / 2 + (w / 2) * Math.max(a, Math.SQRT1_2), hh / 2 - (hh / 2) * Math.max(a, Math.SQRT1_2)], (x, y, w) => [(x - w / 2) / (w / 2)])],
      text: (w, hh) => [w * 0.146, hh * 0.146, w * 0.854, hh * 0.854],
    },
    frame: {
      label: "フレーム", adj: [0.125], lim: [[0, 0.5]], rule: "evenodd",
      d: (w, hh, [a]) => { const t = a * ss(w, hh); return `${poly([[0, 0], [w, 0], [w, hh], [0, hh]])} ${poly([[t, t], [t, hh - t], [w - t, hh - t], [w - t, t]])}`; },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), hh / 2], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const t = a * ss(w, hh); return [t, t, w - t, hh - t]; },
    },
    halfFrame: {
      label: "フレーム（半分）", adj: [0.3333], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const t = a * ss(w, hh); return poly([[0, 0], [w, 0], [w - t, t], [t, t], [t, hh - t], [0, hh]]); },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), hh / 2], (x, y, w, hh) => [x / ss(w, hh)])],
    },
    corner: {
      label: "L 字", adj: [0.5], lim: [[0, 1]],
      d: (w, hh, [a]) => { const t = Math.min(a * ss(w, hh), w, hh); return poly([[0, 0], [t, 0], [t, hh - t], [w, hh - t], [w, hh], [0, hh]]); },
      handles: [handleX((w, hh, [a]) => [Math.min(a * ss(w, hh), w, hh), 0], (x, y, w, hh) => [x / ss(w, hh)])],
    },
    diagStripe: {
      label: "斜め縞", adj: [0.5], lim: [[0, 1]],
      d: (w, hh, [a]) => poly([[0, a * hh], [a * w, 0], [w, 0], [0, hh]]),
      handles: [handleX((w, hh, [a]) => [a * w, 0], (x, y, w) => [x / w])],
    },
    plus: {
      label: "十字形", adj: [0.25], lim: [[0, 0.5]],
      d: (w, hh, [a]) => {
        const c = a * ss(w, hh);
        return poly([[c, 0], [w - c, 0], [w - c, c], [w, c], [w, hh - c], [w - c, hh - c], [w - c, hh], [c, hh], [c, hh - c], [0, hh - c], [0, c], [c, c]]);
      },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), 0], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const c = a * ss(w, hh); return [0, c, w, hh - c]; },
    },
    plaque: {
      label: "ブローチ", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => {
        const c = a * ss(w, hh);
        return `M${P(c, 0)} L${P(w - c, 0)} A${r2(c)} ${r2(c)} 0 0 0 ${P(w, c)} L${P(w, hh - c)} A${r2(c)} ${r2(c)} 0 0 0 ${P(w - c, hh)} L${P(c, hh)} A${r2(c)} ${r2(c)} 0 0 0 ${P(0, hh - c)} L${P(0, c)} A${r2(c)} ${r2(c)} 0 0 0 ${P(c, 0)} Z`;
      },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), 0], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const c = a * ss(w, hh) * 0.72; return [c, c, w - c, hh - c]; },
    },
    can: {
      label: "円柱", adj: [0.25], lim: [[0, 1]],
      d: (w, hh, [a]) => { const ry = Math.min((a * ss(w, hh)) / 2, hh / 2); return `M${P(0, ry)} ${arcTo(w / 2, ry, w / 2, ry, 180, 360)} L${P(w, hh - ry)} ${arcTo(w / 2, hh - ry, w / 2, ry, 0, 180)} Z`; },
      extras: (w, hh, [a]) => { const ry = Math.min((a * ss(w, hh)) / 2, hh / 2); return [{ d: ellipse(w / 2, ry, w / 2, ry), tone: "light" }]; },
      handles: [handleX((w, hh, [a]) => [w / 2, Math.min(a * ss(w, hh), hh)], (x, y, w, hh) => [y / ss(w, hh)])],
      text: (w, hh, [a]) => { const ry = Math.min((a * ss(w, hh)) / 2, hh / 2); return [0, ry * 2, w, hh - ry]; },
    },
    cube: {
      label: "直方体", adj: [0.25], lim: [[0, 1]],
      d: (w, hh, [a]) => { const d = Math.min(a * ss(w, hh), w, hh); return poly([[0, d], [d, 0], [w, 0], [w, hh - d], [w - d, hh], [0, hh]]); },
      extras: (w, hh, [a]) => {
        const d = Math.min(a * ss(w, hh), w, hh);
        return [{ d: poly([[0, d], [d, 0], [w, 0], [w - d, d]]), tone: "light" }, { d: poly([[w - d, d], [w, 0], [w, hh - d], [w - d, hh]]), tone: "dark" }];
      },
      handles: [handleX((w, hh, [a]) => [Math.min(a * ss(w, hh), w), 0], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const d = Math.min(a * ss(w, hh), w, hh); return [0, d, w - d, hh]; },
    },
    bevel: {
      label: "額縁", adj: [0.125], lim: [[0, 0.5]],
      d: (w, hh) => poly([[0, 0], [w, 0], [w, hh], [0, hh]]),
      extras: (w, hh, [a]) => {
        const b = a * ss(w, hh);
        return [
          { d: poly([[0, 0], [w, 0], [w - b, b], [b, b]]), tone: "light" },
          { d: poly([[0, 0], [b, b], [b, hh - b], [0, hh]]), tone: "light" },
          { d: poly([[w, 0], [w, hh], [w - b, hh - b], [w - b, b]]), tone: "dark" },
          { d: poly([[0, hh], [b, hh - b], [w - b, hh - b], [w, hh]]), tone: "dark" },
        ];
      },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), hh / 2], (x, y, w, hh) => [x / ss(w, hh)])],
      text: (w, hh, [a]) => { const b = a * ss(w, hh); return [b, b, w - b, hh - b]; },
    },
    donut: {
      label: "ドーナツ", adj: [0.25], lim: [[0, 0.5]], rule: "evenodd",
      d: (w, hh, [a]) => { const t = a * ss(w, hh); return `${ellipse(w / 2, hh / 2, w / 2, hh / 2)} ${ellipse(w / 2, hh / 2, Math.max(0, w / 2 - t), Math.max(0, hh / 2 - t))}`; },
      handles: [handleX((w, hh, [a]) => [a * ss(w, hh), hh / 2], (x, y, w, hh) => [x / ss(w, hh)])],
    },
    noSmoking: {
      label: "禁止", adj: [0.1875], lim: [[0.02, 0.45]], rule: "evenodd",
      d: (w, hh, [a]) => noSmoking(w, hh, a),
      handles: [handleX((w, hh, [a]) => [a * w, hh / 2], (x, y, w) => [x / w])],
    },
    blockArc: {
      label: "アーチ", adj: [180, 0, 0.25], lim: [[-360, 720], [-360, 720], [0, 0.5]],
      d: (w, hh, [a0, a1, t]) => {
        const [cx, cy] = [w / 2, hh / 2];
        const th = t * ss(w, hh);
        const [rx2, ry2] = [Math.max(1, cx - th), Math.max(1, cy - th)];
        const end = cw(a0, a1);
        return `M${P(...on(cx, cy, cx, cy, a0))} ${arcTo(cx, cy, cx, cy, a0, end)} L${P(...on(cx, cy, rx2, ry2, end))} ${arcTo(cx, cy, rx2, ry2, end, a0)} Z`;
      },
      handles: [
        handleX((w, hh, [a0]) => on(w / 2, hh / 2, w / 2, hh / 2, a0), (x, y, w, hh, [, a1, t]) => [deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2))), a1, t]),
        handleX((w, hh, [a0, a1, t]) => { const th = t * ss(w, hh); return on(w / 2, hh / 2, w / 2 - th, hh / 2 - th, a1); }, (x, y, w, hh, [a0, , t]) => [a0, deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2))), t]),
      ],
    },
    foldedCorner: {
      label: "メモ", adj: [0.1667], lim: [[0, 0.5]],
      d: (w, hh, [a]) => { const c = a * ss(w, hh); return poly([[0, 0], [w, 0], [w, hh - c], [w - c, hh], [0, hh]]); },
      extras: (w, hh, [a]) => { const c = a * ss(w, hh); return [{ d: poly([[w - c, hh], [w - 0.8 * c, hh - 0.8 * c], [w, hh - c]]), tone: "dark" }]; },
      handles: [handleX((w, hh, [a]) => [w - a * ss(w, hh), hh], (x, y, w, hh) => [(w - x) / ss(w, hh)])],
    },
    smileyFace: {
      label: "スマイル", adj: [0.0465], lim: [[-0.0465, 0.0465]],
      d: (w, hh) => ellipse(w / 2, hh / 2, w / 2, hh / 2),
      extras: (w, hh, [a]) => [
        { d: `${ellipse(w * 0.35, hh * 0.38, w * 0.06, hh * 0.06)} ${ellipse(w * 0.65, hh * 0.38, w * 0.06, hh * 0.06)}`, tone: "dark" },
        { d: `M${P(w * 0.3, hh * 0.68)} Q${P(w * 0.5, hh * (0.68 + a * 6))} ${P(w * 0.7, hh * 0.68)}`, tone: "line" },
      ],
      handles: [handleX((w, hh, [a]) => [w / 2, hh * (0.68 + a * 3)], (x, y, w, hh) => [(y / hh - 0.68) / 3])],
    },
    heart: {
      label: "ハート",
      d: (w, hh) => `M${P(0.5 * w, 0.22 * hh)} C${P(0.5 * w, 0.05 * hh)} ${P(0.28 * w, -0.02 * hh)} ${P(0.13 * w, 0.08 * hh)} C${P(-0.04 * w, 0.2 * hh)} ${P(0, 0.45 * hh)} ${P(0.12 * w, 0.58 * hh)} L${P(0.5 * w, hh)} L${P(0.88 * w, 0.58 * hh)} C${P(w, 0.45 * hh)} ${P(1.04 * w, 0.2 * hh)} ${P(0.87 * w, 0.08 * hh)} C${P(0.72 * w, -0.02 * hh)} ${P(0.5 * w, 0.05 * hh)} ${P(0.5 * w, 0.22 * hh)} Z`,
      text: (w, hh) => [w * 0.2, hh * 0.18, w * 0.8, hh * 0.68],
    },
    lightningBolt: { label: "稲妻", d: (w, hh) => poly(scale21600(LIGHTNING, w, hh)), text: (w, hh) => [w * 0.33, hh * 0.35, w * 0.6, hh * 0.6] },
    sun: {
      label: "太陽", adj: [0.25], lim: [[0.12, 0.47]],
      d: (w, hh, [a]) => {
        const [cx, cy] = [w / 2, hh / 2];
        const rc = 0.5 - a;
        const rays = Array.from({ length: 8 }, (_, k) => {
          const ang = k * 45;
          const base = rc + 0.06;
          return poly([on(cx, cy, w * base, hh * base, ang - 9), on(cx, cy, w / 2, hh / 2, ang), on(cx, cy, w * base, hh * base, ang + 9)]);
        });
        return `${ellipse(cx, cy, w * rc, hh * rc)} ${rays.join(" ")}`;
      },
      handles: [handleX((w, hh, [a]) => [w / 2 - w * (0.5 - a), hh / 2], (x, y, w) => [0.5 - (w / 2 - x) / w])],
      text: (w, hh, [a]) => { const r = (0.5 - a) * 0.7; return [w * (0.5 - r), hh * (0.5 - r), w * (0.5 + r), hh * (0.5 + r)]; },
    },
    moon: {
      label: "月", adj: [0.5], lim: [[0.05, 0.95]],
      d: (w, hh, [a]) => `M${P(w, 0)} A${r2(w)} ${r2(hh / 2)} 0 0 0 ${P(w, hh)} A${r2(w * (1 - a))} ${r2(hh / 2)} 0 0 1 ${P(w, 0)} Z`,
      handles: [handleX((w, hh, [a]) => [w * a, hh / 2], (x, y, w) => [x / w])],
      text: (w, hh, [a]) => [w * 0.1, hh * 0.3, w * a * 0.9, hh * 0.7],
    },
    cloud: { label: "雲", d: (w, hh) => cloudPath(w, hh), text: (w, hh) => [w * 0.18, hh * 0.22, w * 0.82, hh * 0.78] },
    arc: {
      label: "円弧", adj: [270, 0], lim: [[-360, 720], [-360, 720]], open: true,
      d: (w, hh, [a0, a1]) => `M${P(...on(w / 2, hh / 2, w / 2, hh / 2, a0))} ${arcTo(w / 2, hh / 2, w / 2, hh / 2, a0, cw(a0, a1))}`,
      handles: [
        handleX((w, hh, [a0]) => on(w / 2, hh / 2, w / 2, hh / 2, a0), (x, y, w, hh, [, a1]) => [deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2))), a1]),
        handleX((w, hh, [, a1]) => on(w / 2, hh / 2, w / 2, hh / 2, a1), (x, y, w, hh, [a0]) => [a0, deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2)))]),
      ],
    },
    bracketPair: {
      label: "大かっこ", adj: [0.1667], lim: [[0, 0.5]], open: true,
      d: (w, hh, [a]) => { const r = a * ss(w, hh); return `M${P(r, 0)} A${r2(r)} ${r2(r)} 0 0 0 ${P(0, r)} L${P(0, hh - r)} A${r2(r)} ${r2(r)} 0 0 0 ${P(r, hh)} M${P(w - r, 0)} A${r2(r)} ${r2(r)} 0 0 1 ${P(w, r)} L${P(w, hh - r)} A${r2(r)} ${r2(r)} 0 0 1 ${P(w - r, hh)}`; },
      handles: [handleX((w, hh, [a]) => [0, a * ss(w, hh)], (x, y, w, hh) => [y / ss(w, hh)])],
      text: (w, hh, [a]) => { const r = a * ss(w, hh) * 0.3; return [r, r, w - r, hh - r]; },
    },
    bracePair: {
      label: "中かっこ", adj: [0.0833], lim: [[0, 0.25]], open: true,
      d: (w, hh, [a]) => {
        const r = a * ss(w, hh);
        const m = hh / 2;
        const left = `M${P(2 * r, 0)} Q${P(r, 0)} ${P(r, r)} L${P(r, m - r)} Q${P(r, m)} ${P(0, m)} Q${P(r, m)} ${P(r, m + r)} L${P(r, hh - r)} Q${P(r, hh)} ${P(2 * r, hh)}`;
        const right = `M${P(w - 2 * r, 0)} Q${P(w - r, 0)} ${P(w - r, r)} L${P(w - r, m - r)} Q${P(w - r, m)} ${P(w, m)} Q${P(w - r, m)} ${P(w - r, m + r)} L${P(w - r, hh - r)} Q${P(w - r, hh)} ${P(w - 2 * r, hh)}`;
        return `${left} ${right}`;
      },
      handles: [handleX((w, hh, [a]) => [0, a * ss(w, hh)], (x, y, w, hh) => [y / ss(w, hh)])],
      text: (w, hh, [a]) => { const r = a * ss(w, hh) * 2; return [r, 0, w - r, hh]; },
    },
    leftBracket: {
      label: "左大かっこ", adj: [0.0833], lim: [[0, 0.5]], open: true,
      d: (w, hh, [a]) => { const r = Math.min(a * hh, hh / 2); return `M${P(w, 0)} A${r2(w)} ${r2(r)} 0 0 0 ${P(0, r)} L${P(0, hh - r)} A${r2(w)} ${r2(r)} 0 0 0 ${P(w, hh)}`; },
      handles: [handleX((w, hh, [a]) => [0, Math.min(a * hh, hh / 2)], (x, y, w, hh) => [y / hh])],
    },
    rightBracket: {
      label: "右大かっこ", adj: [0.0833], lim: [[0, 0.5]], open: true,
      d: (w, hh, [a]) => { const r = Math.min(a * hh, hh / 2); return `M${P(0, 0)} A${r2(w)} ${r2(r)} 0 0 1 ${P(w, r)} L${P(w, hh - r)} A${r2(w)} ${r2(r)} 0 0 1 ${P(0, hh)}`; },
      handles: [handleX((w, hh, [a]) => [w, Math.min(a * hh, hh / 2)], (x, y, w, hh) => [y / hh])],
    },
    leftBrace: {
      label: "左中かっこ", adj: [0.0833], lim: [[0, 0.25]], open: true,
      d: (w, hh, [a]) => {
        const r = a * hh;
        const m = hh / 2;
        return `M${P(w, 0)} Q${P(w / 2, 0)} ${P(w / 2, r)} L${P(w / 2, m - r)} Q${P(w / 2, m)} ${P(0, m)} Q${P(w / 2, m)} ${P(w / 2, m + r)} L${P(w / 2, hh - r)} Q${P(w / 2, hh)} ${P(w, hh)}`;
      },
      handles: [handleX((w, hh, [a]) => [w / 2, a * hh], (x, y, w, hh) => [y / hh])],
    },
    rightBrace: {
      label: "右中かっこ", adj: [0.0833], lim: [[0, 0.25]], open: true,
      d: (w, hh, [a]) => {
        const r = a * hh;
        const m = hh / 2;
        return `M${P(0, 0)} Q${P(w / 2, 0)} ${P(w / 2, r)} L${P(w / 2, m - r)} Q${P(w / 2, m)} ${P(w, m)} Q${P(w / 2, m)} ${P(w / 2, m + r)} L${P(w / 2, hh - r)} Q${P(w / 2, hh)} ${P(0, hh)}`;
      },
      handles: [handleX((w, hh, [a]) => [w / 2, a * hh], (x, y, w, hh) => [y / hh])],
    },

    // ---- ブロック矢印
    rightArrow: arrowShape("矢印: 右", "right"),
    leftArrow: arrowShape("矢印: 左", "left"),
    upArrow: arrowShape("矢印: 上", "up"),
    downArrow: arrowShape("矢印: 下", "down"),
    leftRightArrow: {
      label: "矢印: 左右", adj: [0.5, 0.5], lim: [[0.05, 1], [0, 1]],
      d: (w, hh, [a, b]) => poly(doubleArrow(w, hh, a * hh, b * ss(w, hh))),
      handles: [handleX((w, hh, [, b]) => [Math.min(b * ss(w, hh), w / 2), 0], (x, y, w, hh, [a]) => [a, x / ss(w, hh)]), handleX((w, hh, [a]) => [w / 2, hh / 2 - (a * hh) / 2], (x, y, w, hh, [, b]) => [(hh - 2 * y) / hh, b])],
      text: (w, hh, [a, b]) => [Math.min(b * ss(w, hh), w / 2) * 0.5, hh / 2 - (a * hh) / 2, w - Math.min(b * ss(w, hh), w / 2) * 0.5, hh / 2 + (a * hh) / 2],
    },
    upDownArrow: {
      label: "矢印: 上下", adj: [0.5, 0.5], lim: [[0.05, 1], [0, 1]],
      d: (w, hh, [a, b]) => poly(swap(doubleArrow(hh, w, a * w, b * ss(w, hh)))),
      handles: [handleX((w, hh, [, b]) => [0, Math.min(b * ss(w, hh), hh / 2)], (x, y, w, hh, [a]) => [a, y / ss(w, hh)]), handleX((w, hh, [a]) => [w / 2 - (a * w) / 2, hh / 2], (x, y, w, hh, [, b]) => [(w - 2 * x) / w, b])],
      text: (w, hh, [a, b]) => [w / 2 - (a * w) / 2, Math.min(b * ss(w, hh), hh / 2) * 0.5, w / 2 + (a * w) / 2, hh - Math.min(b * ss(w, hh), hh / 2) * 0.5],
    },
    quadArrow: {
      label: "矢印: 四方向", adj: [0.225, 0.225, 0.225], lim: [[0.02, 0.5], [0.02, 0.5], [0.02, 0.5]],
      d: (w, hh, [a, b, c]) => {
        const m = ss(w, hh);
        const [cx, cy] = [w / 2, hh / 2];
        const t = (a * m) / 2;
        const hw = Math.max(t, b * m);
        const hl = Math.min(c * m, w / 2 - t, hh / 2 - t);
        return poly([[0, cy], [hl, cy - hw], [hl, cy - t], [cx - t, cy - t], [cx - t, hl], [cx - hw, hl], [cx, 0], [cx + hw, hl], [cx + t, hl], [cx + t, cy - t], [w - hl, cy - t], [w - hl, cy - hw], [w, cy],
          [w - hl, cy + hw], [w - hl, cy + t], [cx + t, cy + t], [cx + t, hh - hl], [cx + hw, hh - hl], [cx, hh], [cx - hw, hh - hl], [cx - t, hh - hl], [cx - t, cy + t], [hl, cy + t], [hl, cy + hw]]);
      },
      handles: [handleX((w, hh, [, , c]) => [Math.min(c * ss(w, hh), w / 2), hh / 2], (x, y, w, hh, [a, b]) => [a, b, x / ss(w, hh)])],
    },
    bentArrow: {
      label: "矢印: 折線", adj: [0.25, 0.25, 0.25], lim: [[0.02, 0.5], [0.02, 0.5], [0.02, 0.5]],
      d: (w, hh, [a, b, c]) => {
        const m = ss(w, hh);
        const t = a * m;
        const hh2 = Math.max(t / 2, b * m);
        const hl = Math.min(c * m, w - t);
        const yTop = hh2 - t / 2;
        const yBot = hh2 + t / 2;
        const r = Math.min(t, hh - yTop, w - hl);
        return `M${P(0, hh)} L${P(0, yTop + r)} A${r2(r)} ${r2(r)} 0 0 1 ${P(r, yTop)} L${P(w - hl, yTop)} L${P(w - hl, 0)} L${P(w, hh2)} L${P(w - hl, 2 * hh2)} L${P(w - hl, yBot)} L${P(t, yBot)} L${P(t, hh)} Z`;
      },
      handles: [handleX((w, hh, [a, b, c]) => [w - Math.min(c * ss(w, hh), w), 0], (x, y, w, hh, [a, b]) => [a, b, (w - x) / ss(w, hh)])],
    },
    uturnArrow: {
      label: "矢印: U ターン", adj: [0.25, 0.25, 0.25], lim: [[0.02, 0.45], [0.02, 0.45], [0.02, 0.6]],
      d: (w, hh, [a, b, c]) => {
        const m = ss(w, hh);
        const t = a * m;
        const hw = Math.max(t / 2 + 1, b * m);
        const hl = Math.min(c * m, hh * 0.6);
        const xc = w - hw;
        const rxo = Math.max(t, (xc + t / 2) / 2);
        const ryo = Math.max(t + 1, Math.min(rxo, (hh - hl) * 0.7));
        const right = 2 * rxo;
        return `M${P(0, hh)} L${P(0, ryo)} A${r2(rxo)} ${r2(ryo)} 0 0 1 ${P(right, ryo)} L${P(right, hh - hl)} L${P(xc + hw, hh - hl)} L${P(xc, hh)} L${P(xc - hw, hh - hl)} L${P(right - t, hh - hl)} L${P(right - t, ryo)} A${r2(Math.max(0, rxo - t))} ${r2(Math.max(0, ryo - t))} 0 0 0 ${P(t, ryo)} L${P(t, hh)} Z`;
      },
      handles: [handleX((w, hh, [a, b, c]) => [w, hh - Math.min(c * ss(w, hh), hh * 0.6)], (x, y, w, hh, [a, b]) => [a, b, (hh - y) / ss(w, hh)])],
    },
    bentUpArrow: {
      label: "矢印: 上向き折線", adj: [0.25, 0.25, 0.25], lim: [[0.02, 0.5], [0.02, 0.5], [0.02, 0.6]],
      d: (w, hh, [a, b, c]) => {
        const m = ss(w, hh);
        const t = a * m;
        const hw = Math.max(t / 2, b * m);
        const hl = Math.min(c * m, hh - t);
        const xc = w - hw;
        return poly([[0, hh - t], [xc - t / 2, hh - t], [xc - t / 2, hl], [w - 2 * hw, hl], [xc, 0], [w, hl], [xc + t / 2, hl], [xc + t / 2, hh], [0, hh]]);
      },
      handles: [handleX((w, hh, [a, b, c]) => [w, Math.min(c * ss(w, hh), hh)], (x, y, w, hh, [a, b]) => [a, b, y / ss(w, hh)])],
    },
    circularArrow: {
      label: "矢印: 環状", adj: [180, 330, 0.125], lim: [[-360, 720], [-360, 720], [0.03, 0.3]],
      d: (w, hh, [a0, a1, t]) => {
        const [cx, cy] = [w / 2, hh / 2];
        const th = t * ss(w, hh);
        const hw = th * 1.1;
        const ro = [cx - hw + th / 2, cy - hw + th / 2];
        const ri = [Math.max(1, ro[0] - th), Math.max(1, ro[1] - th)];
        const rm = [(ro[0] + ri[0]) / 2, (ro[1] + ri[1]) / 2];
        const end = cw(a0, a1);
        const tipAngle = end + deg((hw * 1.4) / Math.max(1, (rm[0] + rm[1]) / 2));
        return `M${P(...on(cx, cy, ro[0], ro[1], a0))} ${arcTo(cx, cy, ro[0], ro[1], a0, end)} L${P(...on(cx, cy, rm[0] + hw, rm[1] + hw, end))} L${P(...on(cx, cy, rm[0], rm[1], tipAngle))} L${P(...on(cx, cy, rm[0] - hw, rm[1] - hw, end))} L${P(...on(cx, cy, ri[0], ri[1], end))} ${arcTo(cx, cy, ri[0], ri[1], end, a0)} Z`;
      },
      handles: [
        handleX((w, hh, [a0, , t]) => on(w / 2, hh / 2, w / 2 - t * ss(w, hh) * 0.6, hh / 2 - t * ss(w, hh) * 0.6, a0), (x, y, w, hh, [, a1, t]) => [deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2))), a1, t]),
        handleX((w, hh, [, a1, t]) => on(w / 2, hh / 2, w / 2 - t * ss(w, hh) * 0.6, hh / 2 - t * ss(w, hh) * 0.6, a1), (x, y, w, hh, [a0, , t]) => [a0, deg(Math.atan2((y - hh / 2) / (hh / 2), (x - w / 2) / (w / 2))), t]),
      ],
    },
    stripedRightArrow: {
      label: "矢印: ストライプ", adj: [0.5, 0.5], lim: [[0.05, 1], [0, 2]],
      d: (w, hh, [a, b]) => {
        const m = ss(w, hh);
        const t = a * hh;
        const x0 = (5 / 32) * m;
        const arrow = rightArrow(w - x0, hh, t, Math.min(b * m, w - x0)).map(([x, y]) => [x + x0, y]);
        const y1 = hh / 2 - t / 2;
        const y2 = hh / 2 + t / 2;
        return `${poly(arrow)} ${poly([[0, y1], [m / 32, y1], [m / 32, y2], [0, y2]])} ${poly([[m / 16, y1], [m / 8, y1], [m / 8, y2], [m / 16, y2]])}`;
      },
      handles: [handleX((w, hh, [a, b]) => [w - Math.min(b * ss(w, hh), w), 0], (x, y, w, hh, [a]) => [a, (w - x) / ss(w, hh)])],
    },
    notchedRightArrow: {
      label: "矢印: V 字型", adj: [0.5, 0.5], lim: [[0.05, 1], [0, 2]],
      d: (w, hh, [a, b]) => {
        const t = a * hh;
        const hl = Math.min(b * ss(w, hh), w);
        const pts = rightArrow(w, hh, t, hl);
        return poly([...pts, [Math.min(t * 0.5, w - hl), hh / 2]]);
      },
      handles: [handleX((w, hh, [a, b]) => [w - Math.min(b * ss(w, hh), w), 0], (x, y, w, hh, [a]) => [a, (w - x) / ss(w, hh)])],
    },
    homePlate: {
      label: "矢印: 五方向", adj: [0.5], lim: [[0, 2]],
      d: (w, hh, [a]) => { const d = Math.min(a * ss(w, hh), w); return poly([[0, 0], [w - d, 0], [w, hh / 2], [w - d, hh], [0, hh]]); },
      handles: [handleX((w, hh, [a]) => [w - Math.min(a * ss(w, hh), w), 0], (x, y, w, hh) => [(w - x) / ss(w, hh)])],
      text: (w, hh, [a]) => [0, 0, w - Math.min(a * ss(w, hh), w) / 2, hh],
    },
    chevron: {
      label: "矢印: 山形", adj: [0.5], lim: [[0, 2]],
      d: (w, hh, [a]) => { const d = Math.min(a * ss(w, hh), w / 2); return poly([[0, 0], [w - d, 0], [w, hh / 2], [w - d, hh], [0, hh], [d, hh / 2]]); },
      handles: [handleX((w, hh, [a]) => [w - Math.min(a * ss(w, hh), w / 2), 0], (x, y, w, hh) => [(w - x) / ss(w, hh)])],
      text: (w, hh, [a]) => { const d = Math.min(a * ss(w, hh), w / 2); return [d, 0, w - d, hh]; },
    },
    rightArrowCallout: {
      label: "吹き出し: 右矢印", adj: [0.64, 0.25, 0.35, 0.25], lim: [[0.1, 0.95], [0.02, 0.5], [0.02, 0.5], [0.02, 1]],
      d: (w, hh, [bwr, tr, hr, hlr]) => {
        const bw = bwr * w;
        const cy = hh / 2;
        const t = Math.min(tr * hh, hr * hh) * 1;
        const hhh = Math.max(t / 2, hr * hh);
        const hl = Math.min(hlr * ss(w, hh), w - bw);
        return poly([[0, 0], [bw, 0], [bw, cy - t / 2], [w - hl, cy - t / 2], [w - hl, cy - hhh], [w, cy], [w - hl, cy + hhh], [w - hl, cy + t / 2], [bw, cy + t / 2], [bw, hh], [0, hh]]);
      },
      handles: [handleX((w, hh, [bwr]) => [bwr * w, 0], (x, y, w, hh, [, tr, hr, hlr]) => [x / w, tr, hr, hlr])],
      text: (w, hh, [bwr]) => [0, 0, bwr * w, hh],
    },
    downArrowCallout: {
      label: "吹き出し: 下矢印", adj: [0.64, 0.25, 0.35, 0.25], lim: [[0.1, 0.95], [0.02, 0.5], [0.02, 0.5], [0.02, 1]],
      d: (w, hh, [bhr, tr, hr, hlr]) => {
        const bh = bhr * hh;
        const cx = w / 2;
        const t = Math.min(tr * w, hr * w);
        const hww = Math.max(t / 2, hr * w);
        const hl = Math.min(hlr * ss(w, hh), hh - bh);
        return poly([[0, 0], [w, 0], [w, bh], [cx + t / 2, bh], [cx + t / 2, hh - hl], [cx + hww, hh - hl], [cx, hh], [cx - hww, hh - hl], [cx - t / 2, hh - hl], [cx - t / 2, bh], [0, bh]]);
      },
      handles: [handleX((w, hh, [bhr]) => [0, bhr * hh], (x, y, w, hh, [, tr, hr, hlr]) => [y / hh, tr, hr, hlr])],
      text: (w, hh, [bhr]) => [0, 0, w, bhr * hh],
    },

    // ---- 数式図形
    mathPlus: {
      label: "加算記号", adj: [0.235], lim: [[0.02, 0.6]],
      d: (w, hh, [a]) => {
        const m = ss(w, hh);
        const t = (a * m) / 2;
        const g = m * 0.07;
        const [cx, cy] = [w / 2, hh / 2];
        return poly([[g, cy - t], [cx - t, cy - t], [cx - t, g], [cx + t, g], [cx + t, cy - t], [w - g, cy - t], [w - g, cy + t], [cx + t, cy + t], [cx + t, hh - g], [cx - t, hh - g], [cx - t, cy + t], [g, cy + t]]);
      },
    },
    mathMinus: {
      label: "減算記号", adj: [0.235], lim: [[0.02, 0.6]],
      d: (w, hh, [a]) => { const t = (a * ss(w, hh)) / 2; const g = w * 0.07; return poly([[g, hh / 2 - t], [w - g, hh / 2 - t], [w - g, hh / 2 + t], [g, hh / 2 + t]]); },
    },
    mathMultiply: {
      label: "乗算記号", adj: [0.235], lim: [[0.02, 0.5]],
      d: (w, hh, [a]) => {
        const t = a / 2;
        const L = 0.5 * Math.SQRT2 - 0.1;
        const pts = [];
        for (const th of [45, 135, 225, 315]) {
          const end = on(0.5, 0.5, L, L, th);
          const n = [-Math.sin(rad(th)), Math.cos(rad(th))];
          pts.push([end[0] - t * n[0], end[1] - t * n[1]], [end[0] + t * n[0], end[1] + t * n[1]], on(0.5, 0.5, t * Math.SQRT2, t * Math.SQRT2, th + 45));
        }
        return poly(pts.map(([x, y]) => [x * w, y * hh]));
      },
    },
    mathDivide: {
      label: "除算記号", adj: [0.235], lim: [[0.02, 0.4]],
      d: (w, hh, [a]) => {
        const m = ss(w, hh);
        const t = (a * m) / 2;
        const g = w * 0.07;
        const r = t * 1.15;
        return `${poly([[g, hh / 2 - t], [w - g, hh / 2 - t], [w - g, hh / 2 + t], [g, hh / 2 + t]])} ${ellipse(w / 2, hh / 2 - t - r * 2.2, r, r)} ${ellipse(w / 2, hh / 2 + t + r * 2.2, r, r)}`;
      },
    },
    mathEqual: {
      label: "等号", adj: [0.235], lim: [[0.02, 0.4]],
      d: (w, hh, [a]) => {
        const t = (a * ss(w, hh)) / 2;
        const gap = t * 0.9;
        const g = w * 0.07;
        return `${poly([[g, hh / 2 - gap - t * 2], [w - g, hh / 2 - gap - t * 2], [w - g, hh / 2 - gap], [g, hh / 2 - gap]])} ${poly([[g, hh / 2 + gap], [w - g, hh / 2 + gap], [w - g, hh / 2 + gap + t * 2], [g, hh / 2 + gap + t * 2]])}`;
      },
    },
    mathNotEqual: {
      label: "不等号", adj: [0.235], lim: [[0.02, 0.4]],
      d: (w, hh, [a]) => {
        const t = (a * ss(w, hh)) / 2;
        const gap = t * 0.9;
        const g = w * 0.07;
        const s1 = w * 0.06;
        return `${poly([[g, hh / 2 - gap - t * 2], [w - g, hh / 2 - gap - t * 2], [w - g, hh / 2 - gap], [g, hh / 2 - gap]])} ${poly([[g, hh / 2 + gap], [w - g, hh / 2 + gap], [w - g, hh / 2 + gap + t * 2], [g, hh / 2 + gap + t * 2]])} ${poly([[w * 0.62 - s1, hh * 0.08], [w * 0.62 + s1, hh * 0.08], [w * 0.38 + s1, hh * 0.92], [w * 0.38 - s1, hh * 0.92]])}`;
      },
    },

    // ---- フローチャート
    flowProcess: { label: "フローチャート: 処理", d: (w, hh) => poly([[0, 0], [w, 0], [w, hh], [0, hh]]) },
    flowAlternate: { label: "フローチャート: 代替処理", d: (w, hh) => rrect(0, 0, w, hh, Array(4).fill(ss(w, hh) / 6)), text: (w, hh) => { const i = (ss(w, hh) / 6) * 0.29; return [i, i, w - i, hh - i]; } },
    flowDecision: { label: "フローチャート: 判断", d: (w, hh) => poly([[w / 2, 0], [w, hh / 2], [w / 2, hh], [0, hh / 2]]), text: (w, hh) => [w / 4, hh / 4, (3 * w) / 4, (3 * hh) / 4] },
    flowData: { label: "フローチャート: データ", d: (w, hh) => poly([[w / 5, 0], [w, 0], [(4 * w) / 5, hh], [0, hh]]), text: (w, hh) => [w / 5, 0, (4 * w) / 5, hh] },
    flowPredefined: {
      label: "フローチャート: 定義済み処理", d: (w, hh) => poly([[0, 0], [w, 0], [w, hh], [0, hh]]),
      extras: (w, hh) => [{ d: `M${P(w / 8, 0)} L${P(w / 8, hh)} M${P((7 * w) / 8, 0)} L${P((7 * w) / 8, hh)}`, tone: "line" }], text: (w, hh) => [w / 8, 0, (7 * w) / 8, hh],
    },
    flowInternalStorage: {
      label: "フローチャート: 内部記憶", d: (w, hh) => poly([[0, 0], [w, 0], [w, hh], [0, hh]]),
      extras: (w, hh) => [{ d: `M${P(w / 8, 0)} L${P(w / 8, hh)} M${P(0, hh / 8)} L${P(w, hh / 8)}`, tone: "line" }], text: (w, hh) => [w / 8, hh / 8, w, hh],
    },
    flowDocument: {
      label: "フローチャート: 書類",
      d: (w, hh) => `M0 0 L${P(w, 0)} L${P(w, 0.8 * hh)} C${P(0.83 * w, 0.67 * hh)} ${P(0.66 * w, 0.7 * hh)} ${P(0.5 * w, 0.8 * hh)} C${P(0.33 * w, 0.92 * hh)} ${P(0.16 * w, 0.95 * hh)} ${P(0, 0.85 * hh)} Z`,
      text: (w, hh) => [0, 0, w, 0.75 * hh],
    },
    flowMultidocument: {
      label: "フローチャート: 複数書類",
      d: (w, hh) => {
        const doc = (x, y, ww, hhh) => `M${P(x, y)} L${P(x + ww, y)} L${P(x + ww, y + 0.8 * hhh)} C${P(x + 0.83 * ww, y + 0.67 * hhh)} ${P(x + 0.66 * ww, y + 0.7 * hhh)} ${P(x + 0.5 * ww, y + 0.8 * hhh)} C${P(x + 0.33 * ww, y + 0.92 * hhh)} ${P(x + 0.16 * ww, y + 0.95 * hhh)} ${P(x, y + 0.85 * hhh)} Z`;
        return [doc(0.15 * w, 0, 0.85 * w, 0.8 * hh), doc(0.075 * w, 0.1 * hh, 0.85 * w, 0.8 * hh), doc(0, 0.2 * hh, 0.85 * w, 0.8 * hh)];
      },
      text: (w, hh) => [0, 0.2 * hh, 0.85 * w, 0.8 * hh],
    },
    flowTerminator: { label: "フローチャート: 端子", d: (w, hh) => rrect(0, 0, w, hh, Array(4).fill(Math.min(w, hh) / 2)), text: (w, hh) => [Math.min(w, hh) * 0.3, 0, w - Math.min(w, hh) * 0.3, hh] },
    flowPreparation: { label: "フローチャート: 準備", d: (w, hh) => poly([[w / 5, 0], [(4 * w) / 5, 0], [w, hh / 2], [(4 * w) / 5, hh], [w / 5, hh], [0, hh / 2]]), text: (w, hh) => [w / 5, 0, (4 * w) / 5, hh] },
    flowManualInput: { label: "フローチャート: 手操作入力", d: (w, hh) => poly([[0, hh / 5], [w, 0], [w, hh], [0, hh]]), text: (w, hh) => [0, hh / 5, w, hh] },
    flowManualOperation: { label: "フローチャート: 手作業", d: (w, hh) => poly([[0, 0], [w, 0], [(4 * w) / 5, hh], [w / 5, hh]]), text: (w, hh) => [w / 5, 0, (4 * w) / 5, hh] },
    flowConnector: { label: "フローチャート: 結合子", d: (w, hh) => ellipse(w / 2, hh / 2, w / 2, hh / 2), text: (w, hh) => [w * 0.146, hh * 0.146, w * 0.854, hh * 0.854] },
    flowOffpage: { label: "フローチャート: 他ページ結合子", d: (w, hh) => poly([[0, 0], [w, 0], [w, 0.8 * hh], [w / 2, hh], [0, 0.8 * hh]]), text: (w, hh) => [0, 0, w, 0.8 * hh] },
    flowCard: { label: "フローチャート: カード", d: (w, hh) => poly([[w / 5, 0], [w, 0], [w, hh], [0, hh], [0, hh / 5]]), text: (w, hh) => [0, hh / 5, w, hh] },
    flowPunchedTape: { label: "フローチャート: せん孔テープ", d: (w, hh) => wavePath(w, hh, 0.1, 1), text: (w, hh) => [0, hh * 0.2, w, hh * 0.8] },
    flowSummingJunction: {
      label: "フローチャート: 和接合", d: (w, hh) => ellipse(w / 2, hh / 2, w / 2, hh / 2),
      extras: (w, hh) => { const [a, b] = [on(w / 2, hh / 2, w / 2, hh / 2, 225), on(w / 2, hh / 2, w / 2, hh / 2, 45)]; const [c, d] = [on(w / 2, hh / 2, w / 2, hh / 2, 315), on(w / 2, hh / 2, w / 2, hh / 2, 135)]; return [{ d: `M${P(...a)} L${P(...b)} M${P(...c)} L${P(...d)}`, tone: "line" }]; },
    },
    flowOr: { label: "フローチャート: 論理和", d: (w, hh) => ellipse(w / 2, hh / 2, w / 2, hh / 2), extras: (w, hh) => [{ d: `M${P(w / 2, 0)} L${P(w / 2, hh)} M${P(0, hh / 2)} L${P(w, hh / 2)}`, tone: "line" }] },
    flowCollate: { label: "フローチャート: 照合", d: (w, hh) => poly([[0, 0], [w, 0], [0, hh], [w, hh]]) },
    flowSort: { label: "フローチャート: 分類", d: (w, hh) => poly([[w / 2, 0], [w, hh / 2], [w / 2, hh], [0, hh / 2]]), extras: (w, hh) => [{ d: `M${P(0, hh / 2)} L${P(w, hh / 2)}`, tone: "line" }] },
    flowExtract: { label: "フローチャート: 抜き出し", d: (w, hh) => poly([[w / 2, 0], [w, hh], [0, hh]]), text: (w, hh) => [w / 4, hh / 2, (3 * w) / 4, hh] },
    flowMerge: { label: "フローチャート: 組合せ", d: (w, hh) => poly([[0, 0], [w, 0], [w / 2, hh]]), text: (w, hh) => [w / 4, 0, (3 * w) / 4, hh / 2] },
    flowStoredData: {
      label: "フローチャート: 記憶データ",
      d: (w, hh) => `M${P(w / 6, 0)} L${P(w, 0)} A${r2(w / 6)} ${r2(hh / 2)} 0 0 0 ${P(w, hh)} L${P(w / 6, hh)} A${r2(w / 6)} ${r2(hh / 2)} 0 0 1 ${P(w / 6, 0)} Z`,
      text: (w, hh) => [w / 6, 0, (5 * w) / 6, hh],
    },
    flowDelay: { label: "フローチャート: 論理積ゲート", d: (w, hh) => `M0 0 L${P(w / 2, 0)} A${r2(w / 2)} ${r2(hh / 2)} 0 0 1 ${P(w / 2, hh)} L${P(0, hh)} Z`, text: (w, hh) => [0, hh / 7, w * 0.85, hh * 0.86] },
    flowMagneticDisk: {
      label: "フローチャート: 磁気ディスク",
      d: (w, hh) => { const ry = hh / 6; return `M${P(0, ry)} ${arcTo(w / 2, ry, w / 2, ry, 180, 360)} L${P(w, hh - ry)} ${arcTo(w / 2, hh - ry, w / 2, ry, 0, 180)} Z`; },
      extras: (w, hh) => [{ d: `M${P(0, hh / 6)} ${arcTo(w / 2, hh / 6, w / 2, hh / 6, 180, 0)}`, tone: "line" }],
      text: (w, hh) => [0, hh / 3, w, (5 * hh) / 6],
    },
    flowDirectAccess: {
      label: "フローチャート: 直接アクセス記憶",
      d: (w, hh) => { const rx = w / 6; return `M${P(rx, 0)} L${P(w - rx, 0)} ${arcTo(w - rx, hh / 2, rx, hh / 2, -90, 90)} L${P(rx, hh)} ${arcTo(rx, hh / 2, rx, hh / 2, 90, 270)} Z`; },
      extras: (w, hh) => [{ d: `M${P(w - w / 6, 0)} ${arcTo(w - w / 6, hh / 2, w / 6, hh / 2, -90, -270)}`, tone: "line" }],
      text: (w, hh) => [w / 6, 0, (2 * w) / 3, hh],
    },
    flowDisplay: { label: "フローチャート: 表示", d: (w, hh) => `M${P(w / 6, 0)} L${P((5 * w) / 6, 0)} A${r2(w / 6)} ${r2(hh / 2)} 0 0 1 ${P((5 * w) / 6, hh)} L${P(w / 6, hh)} L${P(0, hh / 2)} Z`, text: (w, hh) => [w / 6, 0, (5 * w) / 6, hh] },

    // ---- 星とリボン
    explosion1: { label: "爆発: 8 pt", d: (w, hh) => poly(scale21600(EXPLOSION1, w, hh)), text: (w, hh) => [w * 0.21, hh * 0.29, w * 0.78, hh * 0.7] },
    explosion2: { label: "爆発: 14 pt", d: (w, hh) => poly(scale21600(EXPLOSION2, w, hh)), text: (w, hh) => [w * 0.22, hh * 0.31, w * 0.75, hh * 0.71] },
    star4: starShape("星: 4 pt", 4, 0.25),
    star5: starShape("星: 5 pt", 5, 0.382),
    star6: starShape("星: 6 pt", 6, 0.577),
    star7: starShape("星: 7 pt", 7, 0.69),
    star8: starShape("星: 8 pt", 8, 0.75),
    star10: starShape("星: 10 pt", 10, 0.85),
    star12: starShape("星: 12 pt", 12, 0.75),
    star16: starShape("星: 16 pt", 16, 0.75),
    star24: starShape("星: 24 pt", 24, 0.75),
    star32: starShape("星: 32 pt", 32, 0.75),
    wave: {
      label: "波線", adj: [0.125], lim: [[0, 0.2]],
      d: (w, hh, [a]) => wavePath(w, hh, a, 1),
      handles: [handleX((w, hh, [a]) => [0, a * hh], (x, y, w, hh) => [y / hh])],
      text: (w, hh, [a]) => [0, a * hh * 2, w, hh - a * hh * 2],
    },
    doubleWave: {
      label: "小波", adj: [0.0625], lim: [[0, 0.12]],
      d: (w, hh, [a]) => wavePath(w, hh, a, 2),
      handles: [handleX((w, hh, [a]) => [0, a * hh], (x, y, w, hh) => [y / hh])],
      text: (w, hh, [a]) => [0, a * hh * 2, w, hh - a * hh * 2],
    },

    // ---- 吹き出し
    wedgeRectCallout: { label: "吹き出し: 四角形", adj: [-0.2083, 0.625], lim: [[-3, 3], [-3, 3]], d: (w, hh, adj) => wedgeRect(w, hh, adj), handles: [tailHandle] },
    wedgeRoundRectCallout: { label: "吹き出し: 角を丸めた四角形", adj: [-0.2083, 0.625], lim: [[-3, 3], [-3, 3]], d: (w, hh, adj) => wedgeRect(w, hh, adj, ss(w, hh) / 6), handles: [tailHandle], text: (w, hh) => { const i = (ss(w, hh) / 6) * 0.29; return [i, i, w - i, hh - i]; } },
    wedgeEllipseCallout: {
      label: "吹き出し: 円形", adj: [-0.2083, 0.625], lim: [[-3, 3], [-3, 3]],
      d: (w, hh, adj) => {
        const [cx, cy] = [w / 2, hh / 2];
        const [tx, ty] = wedgeTip(w, hh, adj);
        if (((tx - cx) / cx) ** 2 + ((ty - cy) / cy) ** 2 <= 1) return ellipse(cx, cy, cx, cy);
        const th = deg(Math.atan2((ty - cy) / cy, (tx - cx) / cx));
        return `M${P(...on(cx, cy, cx, cy, th + 12))} ${arcTo(cx, cy, cx, cy, th + 12, th + 348)} L${P(tx, ty)} Z`;
      },
      handles: [tailHandle], text: (w, hh) => [w * 0.146, hh * 0.146, w * 0.854, hh * 0.854],
    },
    cloudCallout: {
      label: "思考の吹き出し: 雲形", adj: [-0.2083, 0.625], lim: [[-3, 3], [-3, 3]],
      d: (w, hh, adj) => {
        const [tx, ty] = wedgeTip(w, hh, adj);
        const [cx, cy] = [w / 2, hh / 2];
        // Thought bubbles grow from the tip towards the edge of the cloud.
        const th = Math.atan2((ty - cy) / cy, (tx - cx) / cx);
        const edge = [cx + Math.cos(th) * cx * 0.92, cy + Math.sin(th) * cy * 0.92];
        const outside = Math.hypot((tx - cx) / cx, (ty - cy) / cy) > 1;
        const bubbles = !outside ? [] : [0, 0.4, 0.75].map((k, i) => {
          const r = ss(w, hh) * (0.035 + i * 0.03);
          return ellipse(tx + (edge[0] - tx) * k, ty + (edge[1] - ty) * k, r, r);
        });
        return `${cloudPath(w, hh)} ${bubbles.join(" ")}`;
      },
      handles: [tailHandle], text: (w, hh) => [w * 0.18, hh * 0.22, w * 0.82, hh * 0.78],
    },
    lineCallout: {
      label: "吹き出し: 線", adj: [-0.38, 1.1], lim: [[-3, 3], [-3, 3]],
      d: (w, hh) => poly([[0, 0], [w, 0], [w, hh], [0, hh]]),
      extras: (w, hh, adj) => { const [tx, ty] = wedgeTip(w, hh, adj); const start = [clamp(tx, 0, w) === tx ? tx : tx < 0 ? 0 : w, ty > hh ? hh : ty < 0 ? 0 : hh / 2]; return [{ d: `M${P(...start)} L${P(tx, ty)}`, tone: "line" }]; },
      handles: [tailHandle],
    },
    ...ACTION_BUTTONS,
  };
  for (const [key, shape] of Object.entries(SHAPES)) shape.key = key;

  const SHAPE_GROUPS = [
    ["四角形", ["rect", "roundRect", "snip1Rect", "snip2SameRect", "snip2DiagRect", "round1Rect", "round2SameRect", "round2DiagRect"]],
    ["基本図形", ["ellipse", "triangle", "rtTriangle", "parallelogram", "trapezoid", "diamond", "pentagon", "hexagon", "heptagon", "octagon", "decagon", "dodecagon", "pie", "chord", "teardrop", "frame", "halfFrame", "corner", "diagStripe", "plus", "plaque", "can", "cube", "bevel", "donut", "noSmoking", "blockArc", "foldedCorner", "smileyFace", "heart", "lightningBolt", "sun", "moon", "cloud", "arc", "bracketPair", "bracePair", "leftBracket", "rightBracket", "leftBrace", "rightBrace"]],
    ["ブロック矢印", ["rightArrow", "leftArrow", "upArrow", "downArrow", "leftRightArrow", "upDownArrow", "quadArrow", "bentArrow", "uturnArrow", "bentUpArrow", "circularArrow", "stripedRightArrow", "notchedRightArrow", "homePlate", "chevron", "rightArrowCallout", "downArrowCallout"]],
    ["数式図形", ["mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual", "mathNotEqual"]],
    ["フローチャート", ["flowProcess", "flowAlternate", "flowDecision", "flowData", "flowPredefined", "flowInternalStorage", "flowDocument", "flowMultidocument", "flowTerminator", "flowPreparation", "flowManualInput", "flowManualOperation", "flowConnector", "flowOffpage", "flowCard", "flowPunchedTape", "flowSummingJunction", "flowOr", "flowCollate", "flowSort", "flowExtract", "flowMerge", "flowStoredData", "flowDelay", "flowMagneticDisk", "flowDirectAccess", "flowDisplay"]],
    ["星とリボン", ["explosion1", "explosion2", "star4", "star5", "star6", "star7", "star8", "star10", "star12", "star16", "star24", "star32", "wave", "doubleWave"]],
    ["吹き出し", ["wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout", "cloudCallout", "lineCallout"]],
    ["動作設定ボタン", Object.keys(ACTION_BUTTONS)],
  ];

  /** A shape's adjustments, each within its range (defaults for what is missing). */
  function adjOf(shape, adj) {
    const defaults = shape.adj || [];
    return defaults.map((value, i) => {
      const given = Array.isArray(adj) ? Number(adj[i]) : NaN;
      const [lo, hi] = shape.lim?.[i] ?? [-Infinity, Infinity];
      return Number.isFinite(given) ? clamp(given, lo, hi) : value;
    });
  }

  /** The outline of a shape in a w × h box: { paths, extras, text: [l, t, r, b], open, rule }. */
  function geometry(key, w, hh, adj, path = null) {
    if (key === "custom" && Array.isArray(path?.pts) && path.pts.length > 1) return customGeometry(path, Math.max(1, w), Math.max(1, hh));
    const shape = SHAPES[key] || SHAPES.rect;
    const a = adjOf(shape, adj);
    const ww = Math.max(1, w);
    const hhh = Math.max(1, hh);
    const d = shape.d(ww, hhh, a);
    return {
      key: shape.key, paths: Array.isArray(d) ? d : [d], extras: shape.extras ? shape.extras(ww, hhh, a) : [],
      text: shape.text ? shape.text(ww, hhh, a) : [0, 0, ww, hhh], open: Boolean(shape.open), rule: shape.rule || "nonzero", adj: a,
    };
  }

  /**
   * A shape drawn by hand (曲線・フリーフォーム・フリーハンド): its points are fractions of its box, joined by
   * straight lines or a smooth curve through them (Catmull-Rom as cubic Béziers); an open one has no fill.
   */
  function freeformD(path, w, hh) {
    const pts = path.pts.map(([x, y]) => [x * w, y * hh]);
    const n = pts.length;
    const f = (v) => r2(v);
    if (!path.curve || n < 3) return `M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(" L")}${path.closed ? " Z" : ""}`;
    const at = (i) => (path.closed ? pts[(i + n) % n] : pts[Math.max(0, Math.min(n - 1, i))]);
    let d = `M${f(pts[0][0])} ${f(pts[0][1])}`;
    const last = path.closed ? n : n - 1;
    for (let i = 0; i < last; i += 1) {
      const [p0, p1, p2, p3] = [at(i - 1), at(i), at(i + 1), at(i + 2)];
      d += ` C${f(p1[0] + (p2[0] - p0[0]) / 6)} ${f(p1[1] + (p2[1] - p0[1]) / 6)} ${f(p2[0] - (p3[0] - p1[0]) / 6)} ${f(p2[1] - (p3[1] - p1[1]) / 6)} ${f(p2[0])} ${f(p2[1])}`;
    }
    return path.closed ? `${d} Z` : d;
  }
  /** A hand-drawn shape; merged shapes (図形の結合) add more closed rings (holes and separate parts), drawn even-odd. */
  function customGeometry(path, w, hh) {
    const parts = Array.isArray(path.parts) ? path.parts.filter((ring) => Array.isArray(ring) && ring.length > 2) : [];
    const d = [freeformD(path, w, hh), ...parts.map((pts) => freeformD({ pts, closed: true }, w, hh))].join(" ");
    return { key: "custom", paths: [d], extras: [], text: [0, 0, w, hh], open: !path.closed && !parts.length, rule: parts.length ? "evenodd" : "nonzero", adj: [] };
  }

  // ---------------------------------------------------------------- rich text (what people type into shapes and text boxes)
  // Text is kept as a small subset of HTML. Everything is rebuilt from scratch out of an inert <template>, so only
  // known tags and a handful of style properties survive (no scripts, links only to web pages and mail).

  const INLINE = { B: "b", STRONG: "b", I: "i", EM: "i", U: "u", S: "s", STRIKE: "s", DEL: "s", SUB: "sub", SUP: "sup", SPAN: "span", FONT: "span", A: "a" };
  const BLOCKS = new Set(["P", "DIV", "H1", "H2", "H3", "H4", "H5", "H6", "BLOCKQUOTE", "LI"]);
  // 下線の種類 besides the plain line: 二重線・太線・点線・破線・波線.
  const U_LINES = new Set(["double", "thick", "dotted", "dashed", "wavy"]);
  const ALIGN = new Set(["left", "center", "right", "justify"]);
  // 均等割り付け (distributed): justified, the last line too — kept on a box, a cell or a paragraph.
  const ALIGN_ALL = new Set([...ALIGN, "distributed"]);
  const alignCss = (a) => (a === "distributed" ? "justify" : a || null);
  const alignLastCss = (a) => (a === "distributed" ? "justify" : null);

  function hexColor(value) {
    const text = String(value || "").trim().toLowerCase();
    if (/^#[0-9a-f]{6}$/.test(text)) return text;
    if (/^#[0-9a-f]{3}$/.test(text)) return `#${text.slice(1).split("").map((c) => c + c).join("")}`;
    const m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+))?\s*\)$/.exec(text);
    if (m) {
      if (m[4] !== undefined && Number(m[4]) === 0) return null;
      return `#${[m[1], m[2], m[3]].map((v) => clamp(Number(v), 0, 255).toString(16).padStart(2, "0")).join("")}`;
    }
    const named = { black: "#000000", white: "#ffffff", red: "#ff0000", navy: "#000080", gray: "#808080", grey: "#808080" }[text];
    return named || null;
  }
  const FONT_TAG_SIZES = { 1: 20, 2: 26, 3: 32, 4: 36, 5: 48, 6: 64, 7: 96 };
  function pxSize(value) {
    // The editor shows sizes as calc(48px * var(--os, 1)) (the box's shrink factor): the 48px is the size.
    const text = String(value || "").trim().replace(/^calc\(\s*([\d.]+(?:px|pt))\s*\*[^)]*\)\)?$/, "$1");
    const m = /^([\d.]+)(px|pt)?$/.exec(text);
    if (!m) return null;
    const n = Number(m[1]) * (m[2] === "pt" ? PX_PER_PT : 1);
    return Number.isFinite(n) && n > 0 ? clamp(Math.round(n * 10) / 10, 8, 400) : null;
  }

  /** The styles we keep on an inline span. */
  function inlineStyle(el) {
    const out = {};
    const style = el.style;
    const color = hexColor(style?.color || (el.tagName === "FONT" ? el.getAttribute("color") : ""));
    if (color) out.color = color;
    const bg = hexColor(style?.backgroundColor);
    if (bg && bg !== "#ffffff") out["background-color"] = bg;
    const size = pxSize(style?.fontSize) ?? (el.tagName === "FONT" ? FONT_TAG_SIZES[el.getAttribute("size")] ?? null : null);
    if (size) out["font-size"] = `${size}px`;
    const face = String(style?.fontFamily || "").replace(/^['\"]|['\"]$/g, "");
    if (/^[\p{L}\p{N}\s._+\-]{1,100}$/u.test(face)) out["font-family"] = `"${face}"`;
    const weight = String(style?.fontWeight || "");
    if (weight === "bold" || Number(weight) >= 600) out["font-weight"] = "700";
    else if (weight === "normal" || (Number(weight) > 0 && Number(weight) < 600)) out["font-weight"] = "400";
    if (style?.fontStyle === "italic") out["font-style"] = "italic";
    else if (style?.fontStyle === "normal") out["font-style"] = "normal";
    const deco = `${style?.textDecoration || ""} ${style?.textDecorationLine || ""}`;
    const lines = ["underline", "line-through"].filter((word) => deco.includes(word));
    if (lines.length) out["text-decoration"] = lines.join(" ");
    else if (/\bnone\b/.test(deco)) out["text-decoration"] = "none";
    return out;
  }
  const styleText = (style) => Object.entries(style).map(([k, v]) => `${k}: ${v}`).join("; ");

  const FIELDS = new Set(["slideno", "total", "date"]);
  function cleanInline(src, doc, depth = 0) {
    const out = [];
    for (const node of [...src.childNodes]) {
      if (node.nodeType === 3) { if (node.data) out.push(doc.createTextNode(node.data.replace(/ /g, " "))); continue; }
      if (node.nodeType !== 1 || depth > 12) continue;
      const tag = node.tagName.toUpperCase();
      if (tag === "BR") { out.push(doc.createElement("br")); continue; }
      if (BLOCKS.has(tag) || tag === "UL" || tag === "OL") {
        // A block inside a line: keep its words on a new line.
        if (out.length) out.push(doc.createElement("br"));
        out.push(...cleanInline(node, doc, depth + 1));
        continue;
      }
      if (["SCRIPT", "STYLE", "TEMPLATE", "IFRAME", "OBJECT", "EMBED", "SVG", "MATH", "IMG", "VIDEO", "AUDIO", "CANVAS", "INPUT", "TEXTAREA", "SELECT", "BUTTON", "NOSCRIPT", "TITLE", "META", "LINK"].includes(tag)) continue;
      const mapped = INLINE[tag];
      const children = cleanInline(node, doc, depth + 1);
      if (!mapped) { out.push(...children); continue; }
      if (!children.length) continue;
      if (mapped === "a") {
        const href = String(node.getAttribute("href") || "").trim();
        if (!/^(https?:\/\/|mailto:)/i.test(href)) { out.push(...children); continue; }
        const a = doc.createElement("a");
        a.setAttribute("href", href.slice(0, 2000));
        a.setAttribute("target", "_blank");
        a.setAttribute("rel", "noopener noreferrer");
        a.append(...children);
        out.push(a);
        continue;
      }
      if (mapped === "span") {
        // Fields filled in when the slide is drawn (挿入 → スライド番号・日付と時刻の自動更新).
        const field = node.getAttribute("data-field");
        if (FIELDS.has(field)) {
          const span = doc.createElement("span");
          span.setAttribute("data-field", field);
          span.textContent = (node.textContent || "#").slice(0, 24);
          out.push(span);
          continue;
        }
        const style = inlineStyle(node);
        if (!Object.keys(style).length) { out.push(...children); continue; }
        const span = doc.createElement("span");
        span.setAttribute("style", styleText(style));
        span.append(...children);
        out.push(span);
        continue;
      }
      const el = doc.createElement(mapped);
      // 下線の種類・二重取り消し線 (PowerPoint's フォント dialog).
      const line = String(node.getAttribute?.("data-line") || "");
      if ((mapped === "u" && U_LINES.has(line)) || (mapped === "s" && line === "double")) el.setAttribute("data-line", line);
      el.append(...children);
      out.push(el);
    }
    return out;
  }

  /** A length in px from a style ("24px", "0"), within lo..hi (else null). */
  const pxLen = (value, lo, hi) => {
    const m = /^(-?\d+(?:\.\d+)?)(px)?$/.exec(String(value ?? "").trim());
    const n = m ? Math.round(Number(m[1]) * 100) / 100 : NaN;
    return Number.isFinite(n) && n !== 0 && n >= lo && n <= hi ? n : null;
  };
  // ---------------------------------------------------------------- タブ設定 (tab stops of a paragraph)
  const TAB_ALIGNS = { l: "left", c: "center", r: "right" };
  const DEFAULT_TAB = 144;
  /** "l300,r900" → [{ type: "l", x: 300 }, { type: "r", x: 900 }] (sorted, one to a place, at most 8, within the slide). */
  function parseTabs(text) {
    const seen = new Set();
    const stops = [];
    for (const part of String(text ?? "").split(",")) {
      const m = /^\s*([lcr])(\d{1,4}(?:\.\d{1,2})?)\s*$/.exec(part);
      if (!m) continue;
      const x = Math.round(Number(m[2]) * 100) / 100;
      if (!(x >= 1 && x <= 3840) || seen.has(x)) continue;
      seen.add(x);
      stops.push({ type: m[1], x });
    }
    return stops.sort((a, b) => a.x - b.x).slice(0, 8);
  }
  const tabsText = (stops) => (stops || []).map((s) => `${s.type}${s.x}`).join(",");
  /**
   * Where the words between tab characters sit. Every word but a last left-aligned one gets a box: [{ w, align }] (w in px, null for
   * "as long as it is"). A left stop starts the next word at it; a right (centre) stop ends (centres) it there; the room between two
   * stops that do not meet is shared half and half. Past the last stop, tabs fall on every inch (144 px) as in PowerPoint.
   * `start` is where the first word begins (the stops are measured from the same edge).
   */
  function tabBoxes(count, stops, start = 0) {
    const all = (stops || []).slice();
    let at = all.length ? all[all.length - 1].x : 0;
    while (all.length < count - 1) { at = (Math.floor(at / DEFAULT_TAB) + 1) * DEFAULT_TAB; all.push({ type: "l", x: at }); }
    const out = [];
    let cursor = start;
    for (let k = 0; k < count; k += 1) {
      const before = k ? all[k - 1] : null;
      const after = k < count - 1 ? all[k] : null;
      let end;
      if (before?.type === "r") end = before.x;
      else if (before?.type === "c") end = before.x + Math.max(0, before.x - cursor);
      else if (!after) end = null;
      else if (after.type === "l") end = after.x;
      else end = cursor + (after.x - cursor) / 2;
      out.push({ w: end === null ? null : r2(Math.max(0, end - cursor)), align: before ? TAB_ALIGNS[before.type] : "left" });
      if (end !== null) cursor = Math.max(cursor, end);
    }
    return out;
  }
  /** The nodes of a paragraph cut at its tab characters: an array of arrays of nodes (an inline tag across a tab is repeated on each side). */
  function splitTabs(parent, doc) {
    const segs = [[]];
    for (const node of [...parent.childNodes]) {
      if (node.nodeType === 3) {
        node.data.split("\t").forEach((piece, i) => { if (i) segs.push([]); if (piece) segs[segs.length - 1].push(doc.createTextNode(piece)); });
      } else if (node.nodeType === 1 && node.textContent.includes("\t")) {
        splitTabs(node, doc).forEach((nodes, i) => {
          if (i) segs.push([]);
          if (nodes.length) { const copy = node.cloneNode(false); copy.append(...nodes); segs[segs.length - 1].push(copy); }
        });
      } else segs[segs.length - 1].push(node);
    }
    return segs;
  }
  /** タブ設定 on drawing: the paragraphs of a text box that hold tab characters become boxes at their stops (nothing is measured). */
  function layoutTabs(tx) {
    const doc = root.document;
    for (const p of [...tx.querySelectorAll("p")]) {
      if (!p.textContent.includes("\t")) continue;
      const segs = splitTabs(p, doc);
      // Stops are measured from the text area's left edge; the first line starts at the paragraph's margin plus its first-line indent.
      const start = (pxLen(p.style?.marginLeft, 0, 1200) || 0) + (pxLen(p.style?.textIndent, -1200, 1200) || 0);
      const boxes = tabBoxes(segs.length, parseTabs(p.getAttribute("data-tabs")), start);
      p.replaceChildren(...segs.map((nodes, i) => {
        const span = doc.createElement("span");
        span.className = "hs-tabbox";
        if (boxes[i].w !== null) span.setAttribute("style", `width: ${boxes[i].w}px; text-align: ${boxes[i].align}`);
        else span.classList.add("is-end");
        span.append(...nodes);
        return span;
      }));
    }
  }
  function blockAttrs(src, el) {
    const align = String(src.style?.textAlign || src.getAttribute?.("align") || "").toLowerCase();
    const last = String(src.style?.textAlignLast || src.style?.getPropertyValue?.("text-align-last") || "").toLowerCase();
    const styles = [];
    if (align === "justify" && last === "justify") styles.push("text-align: justify", "text-align-last: justify");
    else if (ALIGN.has(align)) styles.push(`text-align: ${align}`);
    // 段落 → インデント: before the text, and the first line indented or hanging (a negative indent).
    const left = pxLen(src.style?.marginLeft, 0, 1200);
    const first = pxLen(src.style?.textIndent, -1200, 1200);
    if (left) styles.push(`margin-left: ${left}px`);
    if (first) styles.push(`text-indent: ${first}px`);
    if (styles.length) el.setAttribute("style", styles.join("; "));
    const indent = Number(src.getAttribute?.("data-indent"));
    if (Number.isInteger(indent) && indent > 0) el.setAttribute("data-indent", String(Math.min(indent, 4)));
    // 段落 → タブ設定: where the tab stops are (data-tabs="l300,r900": left / center / right at that many px from the text's left).
    const tabs = tabsText(parseTabs(src.getAttribute?.("data-tabs")));
    if (tabs && el.tagName === "P") el.setAttribute("data-tabs", tabs);
    // A paragraph's own bullet or number ("■", "①", "1."), as PowerPoint draws it (a hanging mark).
    const bullet = String(src.getAttribute?.("data-bullet") || "");
    if (bullet.trim() && /^[^<>"&\s]{1,4}$/u.test(bullet) && el.tagName !== "LI") el.setAttribute("data-bullet", bullet);
  }

  // 箇条書き・段落番号の種類 (PowerPoint's galleries): a list's own marker, kept on the ul / ol as data-style.
  // A list's marks may take the theme's first accent or its muted text colour (the text's own is the default) and
  // another size (%).
  const LIST_MARKS = { accent: "var(--accent)", muted: "var(--muted)" };
  const LIST_MARK_SIZES = ["75", "125", "150"];
  const LIST_STYLES = {
    ul: { disc: "●", circle: "○", square: "■", diamond: "◆", arrow: "➢", check: "✓", star: "★", dash: "–" },
    ol: { decimal: ["1.", "2.", "3."], paren: ["1)", "2)", "3)"], circled: ["①", "②", "③"], "lower-alpha": ["a.", "b.", "c."], "upper-alpha": ["A.", "B.", "C."], "lower-roman": ["i.", "ii.", "iii."], "upper-roman": ["I.", "II.", "III."], kanji: ["一、", "二、", "三、"] },
  };
  function cleanList(src, doc, depth) {
    const tag = src.tagName.toUpperCase() === "OL" ? "ol" : "ul";
    const list = doc.createElement(tag);
    const style = String(src.getAttribute?.("data-style") || "");
    if (LIST_STYLES[tag][style] && style !== (tag === "ul" ? "disc" : "decimal")) list.setAttribute("data-style", style);
    // 箇条書きと段落番号: the first number, and the marks' colour (the accent or the muted colour) and size.
    const start = Math.round(Number(src.getAttribute?.("start")));
    if (tag === "ol" && Number.isInteger(start) && start > 1 && start < 10000) list.setAttribute("start", String(start));
    const mark = String(src.getAttribute?.("data-mark") || "");
    if (LIST_MARKS[mark]) list.setAttribute("data-mark", mark);
    const msize = String(src.getAttribute?.("data-msize") || "");
    if (LIST_MARK_SIZES.includes(msize)) list.setAttribute("data-msize", msize);
    for (const child of [...src.childNodes]) {
      if (child.nodeType === 1 && ["UL", "OL"].includes(child.tagName.toUpperCase()) && depth < 4) {
        const nested = cleanList(child, doc, depth + 1);
        if (nested.childNodes.length) (list.lastElementChild || list).append(nested);
        continue;
      }
      if (child.nodeType === 1 && child.tagName.toUpperCase() === "LI") {
        const li = doc.createElement("li");
        blockAttrs(child, li);
        const inner = doc.createElement("div");
        for (const sub of [...child.childNodes]) {
          if (sub.nodeType === 1 && ["UL", "OL"].includes(sub.tagName.toUpperCase()) && depth < 4) {
            const nested = cleanList(sub, doc, depth + 1);
            if (nested.childNodes.length) li.append(...cleanInline(inner, doc), nested);
            inner.replaceChildren();
          } else inner.append(sub.cloneNode(true));
        }
        const rest = cleanInline(inner, doc);
        if (rest.length) {
          const first = li.firstChild;
          if (first && first.nodeType === 1 && ["UL", "OL"].includes(first.tagName)) li.append(...rest);
          else li.prepend(...rest);
        }
        if (li.childNodes.length) list.append(li);
        continue;
      }
      const loose = cleanInline({ childNodes: [child] }, doc);
      if (loose.length) { const li = doc.createElement("li"); li.append(...loose); list.append(li); }
    }
    return list;
  }

  /** Any HTML → the rich text we keep (paragraphs, lists, inline marks), as a fragment of new, safe nodes. */
  function richFragment(html) {
    const doc = root.document;
    const frag = doc.createDocumentFragment();
    const tpl = doc.createElement("template");
    tpl.innerHTML = String(html ?? "").slice(0, 1_000_000);
    const source = tpl.content || tpl;
    let para = null;
    const flush = () => { if (para) { frag.append(para); para = null; } };
    const visit = (node, depth = 0) => {
      if (node.nodeType === 1) {
        const tag = node.tagName.toUpperCase();
        if (tag === "UL" || tag === "OL") { flush(); const list = cleanList(node, doc, 0); if (list.childNodes.length) frag.append(list); return; }
        if (BLOCKS.has(tag) && tag !== "LI") {
          // A block that only holds other blocks (Chrome's <div><div>…) is unwrapped.
          if ([...node.childNodes].some((child) => child.nodeType === 1 && (BLOCKS.has(child.tagName.toUpperCase()) || ["UL", "OL"].includes(child.tagName.toUpperCase()))) && depth < 6) {
            flush();
            for (const child of [...node.childNodes]) visit(child, depth + 1);
            flush();
            return;
          }
          flush();
          const p = doc.createElement("p");
          blockAttrs(node, p);
          const kids = cleanInline(node, doc);
          if (/^H[1-6]$/.test(tag) && kids.length) { const b = doc.createElement("b"); b.append(...kids); p.append(b); } else p.append(...kids);
          frag.append(p);
          return;
        }
      }
      // Loose inline content gathers into a paragraph.
      const kids = cleanInline({ childNodes: [node] }, doc);
      if (!kids.length) return;
      if (!para) para = doc.createElement("p");
      for (const kid of kids) {
        if (kid.nodeName === "BR" && para.childNodes.length && para.lastChild?.nodeName !== "BR") { /* a line break inside the paragraph */ }
        para.append(kid);
      }
    };
    for (const node of [...source.childNodes]) visit(node);
    flush();
    // A trailing <br> in a paragraph is how editors keep an empty line open; drop the extra one.
    for (const p of frag.querySelectorAll ? frag.querySelectorAll("p") : []) {
      if (p.lastChild?.nodeName === "BR" && p.childNodes.length > 1) p.lastChild.remove();
    }
    return frag;
  }

  const sanitized = new Map();
  /** The safe, canonical form of rich text (cached). */
  function sanitizeRich(html) {
    const key = String(html ?? "");
    if (sanitized.has(key)) return sanitized.get(key);
    const box = root.document.createElement("div");
    box.append(richFragment(key));
    const out = box.innerHTML;
    if (sanitized.size > 600) sanitized.clear();
    sanitized.set(key, out);
    return out;
  }

  const escapeHtml = (text) => String(text ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
  /** Plain text (lines) → rich text paragraphs. */
  function textToRich(text) {
    return String(text ?? "").split(/\r?\n/).map((line) => `<p>${escapeHtml(line) || "<br>"}</p>`).join("");
  }
  /** Rich text → plain text, a line per paragraph or list item. */
  function richToText(html) {
    const box = root.document.createElement("div");
    box.append(richFragment(html));
    const lines = [];
    const isList = (node) => node.nodeType === 1 && (node.nodeName === "UL" || node.nodeName === "OL");
    const textOf = (el) => {
      let out = "";
      for (const node of el.childNodes) out += node.nodeType === 3 ? node.data : node.nodeName === "BR" ? "\n" : node.nodeType === 1 && !isList(node) ? textOf(node) : "";
      return out;
    };
    const walk = (node) => {
      for (const child of node.childNodes) {
        if (child.nodeType === 1 && (child.nodeName === "P" || child.nodeName === "LI")) {
          lines.push(textOf(child));
          for (const sub of child.childNodes) if (isList(sub)) walk(sub);
        } else if (child.nodeType === 1) walk(child);
      }
    };
    walk(box);
    return lines.join("\n");
  }

  // ---------------------------------------------------------------- ノートの書式 (the speaker notes' own rich text)
  /**
   * The notes as rich text for the notes pane, the presenter view and the notes pages: words, bold / italic / underline /
   * strike, bullets and numbers, links and line breaks. Colours, sizes, alignment and indents are left out — the notes keep
   * each screen's own look (the presenter view has its own colours).
   */
  function noteHtml(html) {
    const box = root.document.createElement("div");
    box.append(richFragment(String(html ?? "").slice(0, 60_000)));
    const doc = root.document;
    for (const el of [...box.querySelectorAll("*")]) {
      if (el.nodeName === "SPAN") {
        // Words made bold, italic, underlined or struck with a style (an editor's way) become the tags that mean it.
        const st = el.getAttribute("style") || "";
        let inner = [...el.childNodes];
        const wrap = (tag) => { const t = doc.createElement(tag); t.append(...inner); inner = [t]; };
        if (/font-weight:\s*(bold|[6-9]00)/i.test(st)) wrap("b");
        if (/font-style:\s*italic/i.test(st)) wrap("i");
        if (/text-decoration[^;]*underline/i.test(st)) wrap("u");
        if (/text-decoration[^;]*line-through/i.test(st)) wrap("s");
        el.replaceWith(...inner);
        continue;
      }
      for (const attr of [...el.attributes]) {
        const keep = (el.nodeName === "A" && ["href", "target", "rel"].includes(attr.name)) || (el.nodeName === "OL" && attr.name === "start");
        if (!keep) el.removeAttribute(attr.name);
      }
    }
    return box.innerHTML;
  }
  /** Notes (rich) → the plain words the studio keeps in `notes`: a line per paragraph, "・" before a bullet, "1. " before a number. */
  function noteToText(html) {
    const box = root.document.createElement("div");
    box.append(richFragment(html));
    const lines = [];
    const isList = (node) => node.nodeType === 1 && (node.nodeName === "UL" || node.nodeName === "OL");
    const textOf = (el) => {
      let out = "";
      for (const node of el.childNodes) out += node.nodeType === 3 ? node.data : node.nodeName === "BR" ? "\n" : node.nodeType === 1 && !isList(node) ? textOf(node) : "";
      return out;
    };
    const walk = (node, depth = 0) => {
      for (const child of node.childNodes) {
        if (child.nodeType !== 1) continue;
        if (child.nodeName === "P") lines.push(textOf(child).replace(/\n$/, ""));
        else if (isList(child)) {
          let n = Math.max(1, Math.round(Number(child.getAttribute("start")) || 1));
          for (const li of child.childNodes) {
            if (li.nodeType !== 1 || li.nodeName !== "LI") continue;
            lines.push(`${"  ".repeat(depth)}${child.nodeName === "OL" ? `${n++}. ` : "・"}${textOf(li).replace(/\n$/, "")}`);
            for (const sub of li.childNodes) if (isList(sub)) walk({ childNodes: [sub] }, depth + 1);
          }
        } else walk(child, depth);
      }
    };
    walk(box);
    return lines.join("\n");
  }
  const sameNotes = (a, b) => String(a ?? "").replace(/[ \t]+(\n|$)/g, "$1").replace(/\s+$/g, "") === String(b ?? "").replace(/[ \t]+(\n|$)/g, "$1").replace(/\s+$/g, "");
  /** Whether the formatting is worth keeping: anything beyond plain paragraphs (a bold word, a list, a link). */
  const noteFormatted = (html) => /<(b|i|u|s|a|ul|ol|sup|sub)[ >]/.test(html);
  /**
   * What a slide's notes are, for every place that shows them: `rich` (HTML — the slide's `notesRich` when it still says what
   * `notes` says, else the plain lines as paragraphs), `text` (the plain words) and whether the formatting is in use.
   */
  function noteOf(slide) {
    const text = String(slide?.notes ?? "");
    const rich = typeof slide?.notesRich === "string" ? noteHtml(slide.notesRich) : "";
    if (rich && noteFormatted(rich) && sameNotes(noteToText(rich), text)) return { rich, text, formatted: true };
    return { rich: text.trim() ? textToRich(text) : "", text, formatted: false };
  }
  /** Notes typed in the pane: the formatted words and the plain words they make (the formatting kept only when it matters). */
  function noteFields(html) {
    const rich = noteHtml(html);
    const text = noteToText(rich);
    return { notes: text, notesRich: noteFormatted(rich) && text.trim() ? rich.slice(0, 120_000) : undefined };
  }

  // ---------------------------------------------------------------- objects: kinds, defaults, normalization

  const KINDS = ["shape", "text", "image", "line", "icon", "video", "audio", "lottie", "table", "chart", "smartart", "ink", "zoom", "camera", "equation", "model", "poll"];
  const KIND_LABELS = { shape: "図形", text: "テキスト ボックス", image: "図", line: "直線", icon: "アイコン", video: "ビデオ", audio: "オーディオ", lottie: "アニメーション", table: "表", chart: "グラフ", smartart: "SmartArt", ink: "インク", zoom: "ズーム", camera: "カメオ", equation: "数式", model: "3D モデル", poll: "アンケート" };
  // Media playback (PowerPoint's 再生 tab): how loud, and where the sound sits on the slides.
  const VOLUMES = { 0: "ミュート", 0.33: "小", 0.66: "中", 1: "大" };
  // 表のスタイル: PowerPoint's built-in styles in the theme's colours (the exporter writes the same style).
  const TABLE_STYLES = { medium2: "中間スタイル 2 - アクセント 1", light1: "淡色スタイル 1 - アクセント 1", light2: "淡色スタイル 2 - アクセント 1", medium2a2: "中間スタイル 2 - アクセント 2", grid: "表のグリッド", lines: "横線だけ", plain: "スタイルなし" };
  // The short names under the gallery's little pictures (the full name is the tooltip).
  const TABLE_STYLE_SHORT = { medium2: "中間 2", light1: "淡色 1", light2: "淡色 2", medium2a2: "中間 2（2）", grid: "グリッド", lines: "横線", plain: "なし" };
  const CHART_KINDS = { bar: "縦棒", "clustered-bar": "集合縦棒", "stacked-bar": "積み上げ縦棒", "100-stacked-bar": "100%積み上げ縦棒", line: "折れ線", "multi-line": "折れ線（複数）", area: "面", pie: "円", donut: "ドーナツ", scatter: "散布図", radar: "レーダー", waterfall: "ウォーターフォール", funnel: "じょうご", combo: "複合（棒と折れ線）",
    hbar: "集合横棒", "stacked-hbar": "積み上げ横棒", "stacked-area": "積み上げ面", bubble: "バブル", histogram: "ヒストグラム", boxplot: "箱ひげ図", treemap: "ツリーマップ", sunburst: "サンバースト", pareto: "パレート図" };
  const DASHES = {
    solid: ["実線", null], roundDot: ["丸点線", [0, 2]], squareDot: ["角点線", [1, 1]], dash: ["破線", [4, 3]], dashDot: ["一点鎖線", [4, 3, 1, 3]],
    longDash: ["長破線", [8, 3]], longDashDot: ["長鎖線", [8, 3, 1, 3]], longDashDotDot: ["長二点鎖線", [8, 3, 1, 3, 1, 3]],
  };
  // 線の書式設定: how a line ends (線端), how its corners join (結合点) and a compound line (複合線: the stroke cut
  // down its middle by one see-through band, or two).
  const LINE_CAPS = { flat: "フラット", round: "丸", square: "四角" };
  const LINE_JOINS = { round: "丸", bevel: "面取り", miter: "とがった" };
  const COMPOUNDS = { dbl: "二重線", tri: "三重線" };
  const SVG_CAPS = { flat: "butt", round: "round", square: "square" };
  const ARROWHEADS = { none: "なし", triangle: "矢印", arrow: "開いた矢印", stealth: "鋭い矢印", diamond: "ひし形", oval: "円" };
  const ROUTES = { straight: "直線", elbow: "カギ線", curve: "曲線" };
  const AUTOFIT = { none: "自動調整なし", shrink: "はみ出す場合だけ自動調整する", grow: "テキストに合わせて図形のサイズを調整する" };
  const FITS = { fill: "図形に合わせる", cover: "トリミングして埋める", contain: "全体を入れる" };
  // What only HTML can do with an object (インタラクション): how it answers the mouse, what a click on it shows,
  // and the motion it keeps while the slide is on screen. They work while presenting and in the exported file.
  const IX_HOVERS = { lift: "浮き上がる", zoom: "拡大する", glow: "光る", tilt: "3Dで傾く", focus: "ほかを薄くする" };
  const IX_LOOPS = { float: "ふわふわ浮く", pulse: "鼓動する", sway: "ゆらゆら揺れる", spin: "回り続ける", shine: "光が走る", bounce: "弾む" };
  const IX_CLICKS = { popup: "詳細を開く", zoom: "拡大して見せる", flip: "裏返す", reveal: "ほかの部品を表示・非表示", spot: "スポットライトを当てる" };

  const DEFAULTS = {
    shape: { shape: "rect", fill: "#4472c4", stroke: "none", strokeW: 2, fs: 36, color: "#ffffff", align: "center", valign: "middle", pad: [7, 14, 7, 14], autofit: "none", lh: 1.35, wrap: true },
    text: { shape: "rect", fill: "none", stroke: "none", strokeW: 2, fs: 40, color: "#000000", align: "left", valign: "top", pad: [7, 14, 7, 14], autofit: "grow", lh: 1.45, wrap: true },
    image: { fit: "fill", stroke: "none", strokeW: 2 },
    line: { stroke: "#4472c4", strokeW: 4, dash: "solid", head: "none", tail: "none", route: "straight", headSize: 2, tailSize: 2 },
    icon: { color: "#4472c4", strokeW: 1.75 },
    video: { fit: "cover", autoplay: false, loop: false, muted: false },
    audio: { autoplay: false, loop: false, color: "#4472c4" },
    smartart: {},
    ink: {},
    zoom: { back: true, stroke: "#4472c4", strokeW: 3 },
    equation: { fs: 56, color: "#000000", align: "center" },
    camera: { mask: "ellipse", stroke: "none", strokeW: 3 },
    lottie: { fit: "contain", autoplay: true, loop: true },
    table: { style: "medium2", header: true, banded: true, firstCol: false, lastRow: false, fs: 28, color: "#000000", lh: 1.35 },
    chart: {},
  };
  /**
   * PowerPoint's defaults in a theme: a shape is filled with アクセント 1 and its words are 背景 1; a text box, an equation
   * and a table write in テキスト 1; lines, icons and zoom frames are アクセント 1.
   */
  function themeDefaults(themeId) {
    const c = themeColors(themeId);
    DEFAULTS.shape.fill = c.accent1;
    DEFAULTS.shape.color = c.bg1;
    DEFAULTS.text.color = c.tx1;
    DEFAULTS.equation.color = c.tx1;
    DEFAULTS.table.color = c.tx1;
    DEFAULTS.line.stroke = c.accent1;
    DEFAULTS.icon.color = c.accent1;
    DEFAULTS.audio.color = c.accent1;
    DEFAULTS.zoom.stroke = c.accent1;
  }
  useTheme(E.DEFAULT_THEME);
  /** An object's settings, with its kind's defaults for whatever it does not set. */
  const withDefaults = (o) => ({ ...DEFAULTS[o.kind], ...o });

  const num = (value, lo, hi, fallback = null) => {
    const n = Number(value);
    return Number.isFinite(n) ? clamp(Math.round(n * 100) / 100, lo, hi) : fallback;
  };
  const colorOrNone = (value) => (value === "none" ? "none" : hexColor(value));
  let idSeed = 0;
  const newId = () => `o${Date.now().toString(36).slice(-5)}${(idSeed++).toString(36)}${Math.random().toString(36).slice(2, 5)}`;

  /**
   * Playback of a video or a sound (PowerPoint's 再生 tab): trimmed start and end, fades, volume, back to the start
   * when done; a sound may play on across slides (a number of slides, 999 = to the end), hide its icon during the
   * show, and take a colour for its icon; a video may play full screen or show only while playing.
   */
  function normalizePlayback(raw, kind) {
    const o = {};
    const start = num(raw.trimStart, 0, 86400);
    if (start) o.trimStart = start;
    const end = num(raw.trimEnd, 0, 86400);
    if (end && end > (start || 0) + 0.1) o.trimEnd = end;
    for (const key of ["fadeIn", "fadeOut"]) { const v = num(raw[key], 0, 60); if (v) o[key] = v; }
    const volume = num(raw.volume, 0, 1);
    if (volume != null && volume !== 1) o.volume = volume;
    if (raw.rewind === true) o.rewind = true;
    if (kind === "audio") {
      const across = raw.across === "all" ? 999 : Math.round(num(raw.across, 0, 999, 0));
      if (across > 1) o.across = across;
      if (raw.hideIcon === true) o.hideIcon = true;
      const color = hexColor(raw.color);
      if (color) o.color = color;
    } else {
      if (raw.fullscreen === true) o.fullscreen = true;
      if (raw.hideIdle === true) o.hideIdle = true;
    }
    const marks = normalizeBookmarks(raw.bookmarks);
    if (marks.length) o.bookmarks = marks;
    return o;
  }
  /**
   * ブックマーク (再生 → ブックマークの追加): named moments in a video or a sound, in time order. The show jumps to
   * one from the bar, and an animation can start as the media passes one (トリガー → ブックマーク時, "id@mark").
   */
  function normalizeBookmarks(list) {
    const seen = new Set();
    const out = [];
    for (const b of (Array.isArray(list) ? list : []).slice(0, 64)) {
      if (!b || typeof b !== "object" || b.t == null || b.t === "") continue;
      const t = num(b.t, 0, 86400);
      if (t == null) continue;
      let id = typeof b.id === "string" && /^[A-Za-z0-9_-]{1,16}$/.test(b.id) ? b.id : `b${Math.round(t * 100).toString(36)}`;
      for (let n = 2; seen.has(id); n += 1) id = `${id.slice(0, 12)}-${n}`;
      seen.add(id);
      out.push({ id, t, name: typeof b.name === "string" ? b.name.replace(/\s+/g, " ").trim().slice(0, 40) : "" });
    }
    return out.sort((a, b) => a.t - b.t).slice(0, 32).map((b, i) => ({ ...b, name: b.name || `ブックマーク ${i + 1}` }));
  }

  // Ink (描画): strokes written with a pen or a highlighter, each a list of points as fractions of the object's box.
  function normalizeStrokes(list) {
    return (Array.isArray(list) ? list : []).slice(0, 400).map((st) => {
      const pts = (Array.isArray(st?.pts) ? st.pts : []).filter((p) => Array.isArray(p) && Number.isFinite(Number(p[0])) && Number.isFinite(Number(p[1]))).slice(0, 4000)
        .map(([x, y]) => [Math.round(clamp(Number(x), -0.5, 1.5) * 10000) / 10000, Math.round(clamp(Number(y), -0.5, 1.5) * 10000) / 10000]);
      if (!pts.length) return null;
      const color = hexColor(st.color);
      return { pts, color: color ? color.toLowerCase() : "#000000", width: num(st.width, 1, 80, 6), ...(st.highlighter === true ? { highlighter: true } : {}) };
    }).filter(Boolean);
  }
  /** A smooth path through a stroke's points (midpoint quadratics), in the box's pixels. */
  function inkPath(pts, w, hh) {
    const p = pts.map(([x, y]) => [x * w, y * hh]);
    if (p.length === 1) return `M${r2(p[0][0])} ${r2(p[0][1])} l0.01 0`;
    let d = `M${r2(p[0][0])} ${r2(p[0][1])}`;
    for (let i = 1; i < p.length - 1; i += 1) d += ` Q${r2(p[i][0])} ${r2(p[i][1])} ${r2((p[i][0] + p[i + 1][0]) / 2)} ${r2((p[i][1] + p[i + 1][1]) / 2)}`;
    return `${d} L${r2(p[p.length - 1][0])} ${r2(p[p.length - 1][1])}`;
  }
  function inkBody(o, rotEl) {
    const svg = s("svg", { class: "hs-obj-geom hs-ink", width: r2(o.w), height: r2(o.h), viewBox: `0 0 ${r2(Math.max(1, o.w))} ${r2(Math.max(1, o.h))}`, overflow: "visible", "aria-hidden": "true" });
    for (const st of o.strokes || []) {
      svg.append(s("path", { d: inkPath(st.pts, o.w, o.h), fill: "none", stroke: st.color, "stroke-width": st.width, "stroke-linecap": st.highlighter ? "square" : "round", "stroke-linejoin": "round", "stroke-opacity": st.highlighter ? 0.9 : null, class: st.highlighter ? "hs-ink-hl" : null }));
    }
    rotEl.append(svg);
  }

  /** Slide numbers, the number of slides and today's date in text (fields), as the slide is drawn. */
  function fillFields(el, ctx) {
    const fields = el.querySelectorAll?.("span[data-field]");
    if (!fields?.length) return;
    const story = E.storyMap(ctx.deck?.slides || []);
    const now = new Date();
    for (const f of fields) {
      const kind = f.getAttribute("data-field");
      // スライド番号の開始番号 (deck.firstNumber) shifts the number the field shows.
      const first = Number.isInteger(ctx.deck?.firstNumber) && ctx.deck.firstNumber >= 0 && ctx.deck.firstNumber <= 9999 ? ctx.deck.firstNumber : 1;
      f.textContent = kind === "slideno" ? String((story.no[ctx.index] ?? (ctx.index ?? 0) + 1) + first - 1)
        : kind === "total" ? String(story.order.length || (ctx.deck?.slides || []).length || 1)
          : `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日`;
    }
  }

  // ---------------------------------------------------------------- 数式 (TeX → MathML)

  const TEX_SYMBOLS = {
    alpha: "α", beta: "β", gamma: "γ", delta: "δ", epsilon: "ε", zeta: "ζ", eta: "η", theta: "θ", iota: "ι", kappa: "κ", lambda: "λ", mu: "μ", nu: "ν", xi: "ξ", pi: "π", rho: "ρ", sigma: "σ", tau: "τ", upsilon: "υ", phi: "φ", chi: "χ", psi: "ψ", omega: "ω",
    Gamma: "Γ", Delta: "Δ", Theta: "Θ", Lambda: "Λ", Xi: "Ξ", Pi: "Π", Sigma: "Σ", Phi: "Φ", Psi: "Ψ", Omega: "Ω",
    infty: "∞", partial: "∂", nabla: "∇", degree: "°", ldots: "…", cdots: "⋯", prime: "′",
  };
  const TEX_OPS = {
    times: "×", cdot: "·", div: "÷", pm: "±", mp: "∓", le: "≤", leq: "≤", ge: "≥", geq: "≥", ne: "≠", neq: "≠", approx: "≈", equiv: "≡", sim: "∼", propto: "∝",
    to: "→", rightarrow: "→", leftarrow: "←", Rightarrow: "⇒", Leftarrow: "⇐", leftrightarrow: "↔", Leftrightarrow: "⇔", in: "∈", notin: "∉", subset: "⊂", supset: "⊃", cup: "∪", cap: "∩", forall: "∀", exists: "∃", therefore: "∴", because: "∵",
  };
  const TEX_BIG = { sum: "∑", prod: "∏", int: "∫", oint: "∮", lim: "lim", max: "max", min: "min", log: "log", ln: "ln", sin: "sin", cos: "cos", tan: "tan", exp: "exp" };
  const MATH_NS = "http://www.w3.org/1998/Math/MathML";
  /** A small TeX subset as MathML elements (built node by node, never from markup). */
  function texToMathML(tex, doc = root.document) {
    const m = (tag, kids = [], text = null) => { const el = doc.createElementNS(MATH_NS, tag); if (text != null) el.textContent = text; for (const k of kids) if (k) el.append(k); return el; };
    let i = 0;
    const src = String(tex);
    const peek = () => src[i];
    const skip = () => { while (i < src.length && /\s/.test(src[i])) i += 1; };
    function group() {
      skip();
      if (peek() === "{") { i += 1; const kids = sequence("}"); i += 1; return kids.length === 1 ? kids[0] : m("mrow", kids); }
      return atom();
    }
    function command() {
      i += 1;
      let name = "";
      while (i < src.length && /[A-Za-z]/.test(src[i])) { name += src[i]; i += 1; }
      if (!name) { const ch = src[i] || ""; i += 1; return ch === "\\" ? m("mspace", [], null) : m("mo", [], ch); }
      if (name === "frac") return m("mfrac", [group(), group()]);
      if (name === "sqrt") {
        skip();
        if (peek() === "[") { i += 1; let idx = ""; while (i < src.length && src[i] !== "]") { idx += src[i]; i += 1; } i += 1; return m("mroot", [group(), texRow(idx)]); }
        return m("msqrt", [group()]);
      }
      if (name === "text" || name === "mathrm") { skip(); let t = ""; if (peek() === "{") { i += 1; let depth = 1; while (i < src.length && depth) { if (src[i] === "{") depth += 1; if (src[i] === "}") depth -= 1; if (depth) t += src[i]; i += 1; } } return m(name === "text" ? "mtext" : "mi", [], t); }
      if (name === "left" || name === "right") { skip(); const ch = src[i] === "\\" ? (i += 2, src[i - 1]) : src[i++]; return m("mo", [], ch === "." ? "" : ch); }
      if (name === "quad" || name === "qquad") return m("mspace", [], null);
      if (TEX_SYMBOLS[name]) return m("mi", [], TEX_SYMBOLS[name]);
      if (TEX_OPS[name]) return m("mo", [], TEX_OPS[name]);
      if (TEX_BIG[name]) { const el = m(TEX_BIG[name].length > 1 ? "mi" : "mo", [], TEX_BIG[name]); if (TEX_BIG[name].length === 1) el.setAttribute("largeop", "true"); return el; }
      return m("mi", [], name);
    }
    function atom() {
      skip();
      const ch = peek();
      if (ch === undefined) return null;
      if (ch === "\\") return command();
      if (ch === "{") return group();
      if (/[0-9.]/.test(ch)) { let n = ""; while (i < src.length && /[0-9.]/.test(src[i])) { n += src[i]; i += 1; } return m("mn", [], n); }
      i += 1;
      if (/[A-Za-z]/.test(ch)) return m("mi", [], ch);
      if (/[-+=<>*/!|(),:;\[\]]/.test(ch)) return m("mo", [], ch === "-" ? "−" : ch === "*" ? "×" : ch);
      return m("mi", [], ch);
    }
    function sequence(until) {
      const out = [];
      while (i < src.length && src[i] !== until) {
        skip();
        if (i >= src.length || src[i] === until) break;
        let base = atom();
        if (!base) break;
        // Subscripts and superscripts (in either order); big operators take them as limits.
        let sub = null;
        let sup = null;
        for (let k = 0; k < 2; k += 1) {
          skip();
          if (src[i] === "_" && !sub) { i += 1; sub = group(); }
          else if (src[i] === "^" && !sup) { i += 1; sup = group(); }
        }
        const big = base.getAttribute?.("largeop") === "true";
        if (sub && sup) base = m(big ? "munderover" : "msubsup", [base, sub, sup]);
        else if (sub) base = m(big ? "munder" : "msub", [base, sub]);
        else if (sup) base = m(big ? "mover" : "msup", [base, sup]);
        out.push(base);
      }
      return out;
    }
    const texRow = (t) => texToMathML(t, doc).firstChild || m("mrow");
    const math = m("math", [m("mrow", sequence(undefined))]);
    math.setAttribute("display", "block");
    return math;
  }
  function equationBody(o, rotEl) {
    const box = h("div", { class: "hs-equation", style: { "font-size": `${o.fs}px`, color: o.color, "justify-content": o.align === "left" ? "flex-start" : o.align === "right" ? "flex-end" : "center" } });
    box.append(texToMathML(o.tex));
    rotEl.append(box);
  }

  // スライド ズーム: the target slide drawn small inside the object (not again inside itself).
  let zoomDepth = 0;
  function zoomBody(o, rotEl, ctx) {
    const slides = ctx.deck?.slides || [];
    const at = slides.findIndex((slide) => slide.sid === o.target);
    const box = h("div", { class: "hs-zoom", style: { "--zb": o.stroke === "none" ? "transparent" : o.stroke, "--zw": `${o.strokeW ?? 3}px` } });
    if (at < 0 || zoomDepth > 0) {
      box.append(h("div", { class: "hs-zoom-missing" }, at < 0 ? "リンク先のスライドがありません" : `スライド ${at + 1}`));
    } else {
      zoomDepth += 1;
      try {
        const el = E.render(slides[at], { deck: ctx.deck, index: at, mode: "thumb", assetBase: ctx.assetBase, assetMap: ctx.assetMap, mediaUrls: ctx.mediaUrls });
        box.append(E.mount(el, { contain: true, className: "hs-zoom-thumb" }));
      } finally { zoomDepth -= 1; }
    }
    rotEl.append(box);
  }
  /** カメオ: a placeholder while editing; while presenting, the camera (the player starts it). */
  /**
   * 3D モデル: its still picture (or a placeholder); models.js draws it live on the stage and in a slide show. The
   * box carries what the live drawing needs: where the file is and the view.
   */
  function modelBody(o, rotEl, ctx) {
    const url = o.src.startsWith("builtin:") ? o.src : E.resolveSrc(o.src, ctx);
    const box = h("div", { class: "hs-model", "data-src": url || "", "data-view": JSON.stringify(o.view || {}) });
    if (o.poster) box.append(h("img", { class: "hs-model-poster", src: o.poster, alt: o.decorative ? "" : o.alt || "", draggable: "false" }));
    else box.append(h("div", { class: "hs-model-placeholder" }, E.icon("box", "hs-model-icon") || "", h("span", {}, "3D モデル")));
    rotEl.append(box);
  }

  /**
   * アンケート: the question, then each choice with a bar for its share of the answers (counts come in while
   * presenting: motion.js fills them). In a show a choice can be clicked to answer.
   */
  function pollBody(o, rotEl, ctx) {
    const fs = o.fs || 30;
    const box = h("div", { class: "hs-poll", "data-poll": o.id, style: { "--poll-fs": `${fs}px`, "--poll-bar": o.fill || "#c9d6ee" } },
      h("div", { class: "hs-poll-q" }, o.question),
      h("ol", { class: "hs-poll-opts" }, o.options.map((label, k) => h("li", { class: "hs-poll-opt", "data-option": String(k), role: ctx.live ? "button" : null, tabindex: ctx.live ? "0" : null },
        h("span", { class: "hs-poll-label" }, label),
        h("span", { class: "hs-poll-bar" }, h("i", { style: { width: "0%" } })),
        h("span", { class: "hs-poll-count" }, "0")))),
      h("div", { class: "hs-poll-total" }, ctx.live ? "回答をクリック・タップ" : "発表中に回答を集めます"));
    rotEl.append(box);
  }

  function cameraBody(o, rotEl, ctx) {
    const radius = o.mask === "ellipse" ? "50%" : o.mask === "roundRect" ? "12%" : "0";
    const box = h("div", { class: "hs-camera", style: { "border-radius": radius, "--cb": o.stroke === "none" ? "transparent" : o.stroke, "--cw": `${o.strokeW ?? 3}px` } });
    if (ctx.live) box.append(h("video", { class: "hs-camera-feed", autoplay: true, muted: true, playsinline: true }));
    else box.append(h("div", { class: "hs-camera-placeholder" }, h("span", {}, "カメラ")));
    rotEl.append(box);
  }

  /** One object, as kept in a deck: unknown fields go, numbers are clamped, text is made safe. Null if unusable. */
  function normalizeObject(raw) {
    if (!raw || typeof raw !== "object" || !KINDS.includes(raw.kind)) return null;
    const o = { id: /^[A-Za-z0-9_-]{1,32}$/.test(raw.id ?? "") ? raw.id : newId(), kind: raw.kind };
    if (typeof raw.name === "string" && raw.name.trim()) o.name = raw.name.trim().slice(0, 60);
    if (o.kind === "line") {
      for (const key of ["x1", "y1", "x2", "y2"]) o[key] = num(raw[key], -W * 2, W * 3, key.startsWith("x") ? (key === "x1" ? 400 : 800) : 540);
      if (ROUTES[raw.route] && raw.route !== "straight") o.route = raw.route;
      for (const end of ["from", "to"]) {
        const ref = raw[end];
        if (ref && typeof ref === "object" && /^[A-Za-z0-9_-]{1,32}$/.test(ref.id ?? "") && Number.isInteger(ref.site) && ref.site >= 0 && ref.site < 8) o[end] = { id: ref.id, site: ref.site };
      }
      if (ARROWHEADS[raw.head] && raw.head !== "none") o.head = raw.head;
      if (ARROWHEADS[raw.tail] && raw.tail !== "none") o.tail = raw.tail;
      for (const key of ["headSize", "tailSize"]) { const v = num(raw[key], 1, 3); if (v != null && v !== 2) o[key] = Math.round(v); }
    } else {
      o.x = num(raw.x, -W * 2, W * 3, 100);
      o.y = num(raw.y, -H * 2, H * 3, 100);
      o.w = num(raw.w, 1, W * 4, 300);
      o.h = num(raw.h, 1, H * 4, 200);
      const rot = num(raw.rot, -3600, 3600, 0);
      if (rot) { const r = mod(rot + 180, 360) - 180; if (r) o.rot = Math.round(r * 100) / 100; }
      if (raw.flipH === true) o.flipH = true;
      if (raw.flipV === true) o.flipV = true;
    }
    const opacity = num(raw.opacity, 0, 1);
    if (opacity != null && opacity < 1) o.opacity = opacity;
    // 3-D 回転 (図形の効果): the object turned about its horizontal (x) and vertical (y) axis, in degrees, seen from a camera whose
    // field of view is `p` degrees (45 unless said; 0 is a flat, isometric view). Shapes, text boxes and pictures.
    if (raw.rot3d && typeof raw.rot3d === "object" && ["shape", "text", "image"].includes(o.kind)) {
      const x = Math.round(num(raw.rot3d.x, -89, 89, 0) * 10) / 10;
      const y = Math.round(num(raw.rot3d.y, -89, 89, 0) * 10) / 10;
      const p = Math.round(num(raw.rot3d.p, 0, 120, 45));
      if (Math.abs(x) >= 0.5 || Math.abs(y) >= 0.5) o.rot3d = { x, y, ...(p !== 45 ? { p } : {}) };
    }
    if (raw.shadow && typeof raw.shadow === "object") {
      const color = hexColor(raw.shadow.color);
      if (color) o.shadow = { dx: num(raw.shadow.dx, -500, 500, 0), dy: num(raw.shadow.dy, -500, 500, 0),
        blur: num(raw.shadow.blur, 0, 500, 0), color, opacity: num(raw.shadow.opacity, 0, 1, 1) };
    }
    // 図形の効果 besides 影: 反射 (reflect: the share of the object mirrored below it, how strong, and the gap), 光彩
    // (glow: a halo of a colour) and ぼかし (soft: the edges fade out over so many px). Shapes, text boxes and pictures.
    if (["shape", "text", "image"].includes(o.kind)) {
      if (raw.reflect && typeof raw.reflect === "object") {
        const size = Number(raw.reflect.size);
        if (Number.isFinite(size) && size >= 0.05) o.reflect = { size: r2(Math.min(1, size)), opacity: r2(num(raw.reflect.opacity, 0, 1, 0.4)), gap: Math.round(num(raw.reflect.gap, 0, 100, 0)) };
      }
      if (raw.glow && typeof raw.glow === "object") {
        const color = hexColor(raw.glow.color);
        if (color) o.glow = { r: Math.round(num(raw.glow.r, 1, 200, 10)), color: color.toLowerCase(), opacity: r2(num(raw.glow.opacity, 0.05, 1, 0.6)) };
      }
      const soft = num(raw.soft, 0, 200);
      if (soft != null && soft >= 1) o.soft = Math.round(soft);
    }
    for (const key of ["locked", "hidden"]) if (raw[key] === true) o[key] = true;
    // 代替テキスト (what a screen reader says for the object) or 装飾用 (a screen reader skips it).
    if (raw.decorative === true) o.decorative = true;
    else if (typeof raw.alt === "string" && raw.alt.trim()) o.alt = raw.alt.trim().slice(0, 500);
    if (typeof raw.group === "string" && /^[A-Za-z0-9_-]{1,32}$/.test(raw.group)) o.group = raw.group;
    // The layout item an object came from (図形に変換): its "詳しく" card and deep-dive page open from it.
    if (typeof raw.item === "string" && /^[A-Za-z]{1,20}(\[\d{1,2}\])?$/.test(raw.item)) o.item = raw.item;
    // Lines, outlines and fills
    const stroke = colorOrNone(raw.stroke);
    if (stroke) o.stroke = stroke;
    const strokeW = num(raw.strokeW, 0.5, 200);
    if (strokeW != null) o.strokeW = strokeW;
    if (DASHES[raw.dash] && raw.dash !== "solid") o.dash = raw.dash;
    if (LINE_CAPS[raw.cap] && raw.cap !== "flat") o.cap = raw.cap;
    if (LINE_JOINS[raw.join]) o.join = raw.join;
    if (COMPOUNDS[raw.cmpd]) o.cmpd = raw.cmpd;
    if (["shape", "text"].includes(o.kind)) {
      o.shape = SHAPES[raw.shape] ? raw.shape : "rect";
      // A shape drawn by hand keeps its points (fractions of its box).
      if (raw.shape === "custom" && o.kind === "shape") {
        const pts = (Array.isArray(raw.path?.pts) ? raw.path.pts : []).filter((p) => Array.isArray(p) && Number.isFinite(Number(p[0])) && Number.isFinite(Number(p[1]))).slice(0, 5000).map(([x, y]) => [Math.round(clamp(Number(x), -1, 2) * 10000) / 10000, Math.round(clamp(Number(y), -1, 2) * 10000) / 10000]);
        if (pts.length > 1) { o.shape = "custom"; o.path = { pts, ...(raw.path.closed ? { closed: true } : {}), ...(raw.path.curve ? { curve: true } : {}) }; }
        // More rings of a merged shape (図形の結合): closed, straight-edged, fractions of the box like the points.
        const ring = (list) => (Array.isArray(list) ? list : []).filter((p) => Array.isArray(p) && Number.isFinite(Number(p[0])) && Number.isFinite(Number(p[1]))).slice(0, 5000).map(([x, y]) => [Math.round(clamp(Number(x), -1, 2) * 10000) / 10000, Math.round(clamp(Number(y), -1, 2) * 10000) / 10000]);
        const parts = (Array.isArray(raw.path?.parts) ? raw.path.parts : []).slice(0, 400).map(ring).filter((r) => r.length > 2);
        if (o.path && parts.length) { o.path.parts = parts; o.path.closed = true; delete o.path.curve; }
      }
      if (Array.isArray(raw.adj) && SHAPES[o.shape]?.adj) o.adj = adjOf(SHAPES[o.shape], raw.adj).map((v) => Math.round(v * 10000) / 10000);
      const fill = colorOrNone(raw.fill);
      if (fill) o.fill = fill;
      if (raw.gradient && typeof raw.gradient === "object" && Array.isArray(raw.gradient.stops)) {
        const stops = raw.gradient.stops.slice(0, 100).map((stop) => ({ at: num(stop?.at, 0, 1), color: hexColor(stop?.color), opacity: num(stop?.opacity, 0, 1, 1) }))
          .filter((stop) => stop.at != null && stop.color);
        if (stops.length) o.gradient = { angle: num(raw.gradient.angle, -360, 360, 0), stops };
      }
      // パターンの塗りつぶし: a kind (PATTERN_KINDS) and its two colours.
      if (raw.pattern && typeof raw.pattern === "object" && PATTERN_KINDS[raw.pattern.kind]) {
        const fg = hexColor(raw.pattern.fg);
        const bg = hexColor(raw.pattern.bg);
        if (fg && bg) o.pattern = { kind: raw.pattern.kind, fg: fg.toLowerCase(), bg: bg.toLowerCase() };
      }
      // 図形の塗りつぶし → 図: a picture fills the shape (cut to its outline), stretched to cover it or tiled.
      if (typeof raw.fillImg === "string" && /^(data:image\/(png|jpeg|webp|gif)[;,]|idb:|asset:|https:\/\/|blob:)/i.test(raw.fillImg.trim()) && raw.fillImg.length <= 30_000_000) {
        o.fillImg = raw.fillImg.trim();
        if (raw.fillTile === true) o.fillTile = true;
      }
      const fillOpacity = num(raw.fillOpacity, 0, 1);
      if (fillOpacity != null && fillOpacity < 1) o.fillOpacity = fillOpacity;
    }
    // Text
    if (["shape", "text"].includes(o.kind)) {
      if (typeof raw.text === "string" && raw.text.trim()) {
        const text = sanitizeRich(raw.text);
        if (text.replace(/<[^>]+>/g, "").trim() || /<br>/.test(text)) o.text = text;
      }
      const fs = num(raw.fs, 8, 400);
      if (fs != null) o.fs = fs;
      const color = hexColor(raw.color);
      if (color) o.color = color;
      for (const key of ["bold", "italic", "underline", "strike", "vertical"]) if (typeof raw[key] === "boolean") o[key] = raw[key];
      // 文字列の方向: すべてのテキストを90度・270度回転 (the letters on their sides).
      if (raw.textRot === 90 || raw.textRot === 270) o.textRot = raw.textRot;
      if (raw.lockRatio === true) o.lockRatio = true;
      if (typeof raw.wrap === "boolean") o.wrap = raw.wrap;
      if (U_LINES.has(raw.uline)) o.uline = raw.uline;
      if (raw.sline === "double") o.sline = "double";
      if (ALIGN_ALL.has(raw.align)) o.align = raw.align;
      if (["top", "middle", "bottom"].includes(raw.valign)) o.valign = raw.valign;
      if (FONTS[raw.font] && raw.font !== "body") o.font = raw.font;
      if (typeof raw.fontFace === "string" && /^[\p{L}\p{N}\s._+\-]{1,100}$/u.test(raw.fontFace)) o.fontFace = raw.fontFace;
      const lh = num(raw.lh, 0.8, 4);
      if (lh != null) o.lh = lh;
      const ls = num(raw.ls, -0.2, 1);
      if (ls) o.ls = ls;
      const psp = num(raw.psp, 0, 4);
      if (psp) o.psp = psp;
      // フォント → 文字飾り: すべて大文字 / 小型英大文字 (Latin letters only change).
      if (raw.caps === "all" || raw.caps === "small") o.caps = raw.caps;
      // 文字の効果 (図形の書式 → 文字の効果): 影 (tshadow: how far, how blurred, colour, strength) and 光彩 (tglow: a halo) on the letters themselves.
      if (raw.tshadow && typeof raw.tshadow === "object") {
        const color = hexColor(raw.tshadow.color);
        const shadow = color ? { dx: r2(num(raw.tshadow.dx, -100, 100, 0)), dy: r2(num(raw.tshadow.dy, -100, 100, 0)), blur: r2(num(raw.tshadow.blur, 0, 100, 0)), color: color.toLowerCase(), opacity: r2(num(raw.tshadow.opacity, 0, 1, 1)) } : null;
        if (shadow && (shadow.dx || shadow.dy || shadow.blur) && shadow.opacity > 0) o.tshadow = shadow;
      }
      if (raw.tglow && typeof raw.tglow === "object") {
        const color = hexColor(raw.tglow.color);
        if (color) o.tglow = { r: Math.round(num(raw.tglow.r, 1, 100, 8)), color: color.toLowerCase(), opacity: r2(num(raw.tglow.opacity, 0.05, 1, 0.6)) };
      }
      // 文字の輪郭 (toutline): a line of a colour and a width (px) round the letters, drawn behind their fill.
      if (raw.toutline && typeof raw.toutline === "object") {
        const color = hexColor(raw.toutline.color);
        const w = Number(raw.toutline.w);
        if (color && Number.isFinite(w) && w >= 0.25) o.toutline = { w: r2(Math.min(24, w)), color: color.toLowerCase() };
      }
      if (Array.isArray(raw.pad) && raw.pad.length === 4) o.pad = raw.pad.map((v) => num(v, 0, 400, 0));
      if (AUTOFIT[raw.autofit]) o.autofit = raw.autofit;
      // 段組み: the words run in two to four columns.
      const cols = Math.round(num(raw.cols, 1, 4, 1));
      if (cols > 1) { o.cols = cols; const gap = num(raw.colGap, 0, 200); if (gap != null) o.colGap = gap; }
    }
    // Pictures, videos, animations
    if (["image", "video", "audio", "lottie"].includes(o.kind)) {
      const src = typeof raw.src === "string" ? raw.src.trim() : "";
      if (!/^(data:(image|video|audio|application\/json)|idb:|asset:|https?:\/\/|blob:)/i.test(src) || src.length > 80_000_000) return null;
      o.src = src;
      if (typeof raw.fileName === "string" && raw.fileName.trim()) o.fileName = raw.fileName.trim().slice(0, 200);
      if (FITS[raw.fit]) o.fit = raw.fit;
      if (o.kind === "image") {
        const crop = raw.crop;
        if (crop && typeof crop === "object") {
          const c = { l: num(crop.l, 0, 0.95, 0), t: num(crop.t, 0, 0.95, 0), r: num(crop.r, 0, 0.95, 0), b: num(crop.b, 0, 0.95, 0) };
          if (c.l + c.r < 0.98 && c.t + c.b < 0.98 && (c.l || c.t || c.r || c.b)) o.crop = c;
        }
        if (SHAPES[raw.mask] && raw.mask !== "rect") o.mask = raw.mask;
        if (Array.isArray(raw.adj) && SHAPES[o.mask]?.adj) o.adj = adjOf(SHAPES[o.mask], raw.adj).map((v) => Math.round(v * 10000) / 10000);
        for (const key of ["bright", "contrast", "sat"]) { const v = num(raw[key], -1, 1); if (v) o[key] = v; }
        if (raw.lockRatio === false) o.lockRatio = false;
        if (raw.gray === true) o.gray = true;
      } else {
        for (const key of ["autoplay", "loop", "muted"]) if (typeof raw[key] === "boolean") o[key] = raw[key];
      }
      if (o.kind === "video" || o.kind === "audio") Object.assign(o, normalizePlayback(raw, o.kind));
      // ビデオの表紙画像 (PowerPoint's poster frame): what shows before the video plays.
      if (o.kind === "video" && typeof raw.poster === "string" && /^data:image\/(png|jpeg|webp);base64,/.test(raw.poster) && raw.poster.length <= 3_000_000) o.poster = raw.poster;
      // スライド ショーの記録: the slide's narration (played with the slide unless 「ナレーションを付けない」).
      if (o.kind === "audio" && raw.narration === true) o.narration = true;
      // 再生 →「キャプションの挿入」: subtitles for a video (WebVTT text, shown with the video).
      if (o.kind === "video" && typeof raw.captions === "string" && /^\uFEFF?WEBVTT/.test(raw.captions.trim()) && raw.captions.length <= 300_000) {
        o.captions = raw.captions.replace(/^\uFEFF/, "");
        if (typeof raw.captionLang === "string" && /^[a-z]{2,3}(-[A-Za-z]{2,4})?$/.test(raw.captionLang)) o.captionLang = raw.captionLang;
      }
    }
    if (o.kind === "zoom") {
      // スライド ズーム: a live picture of another slide; clicking it in the show goes there (and back).
      if (typeof raw.target !== "string" || !/^[A-Za-z0-9_-]{1,32}$/.test(raw.target)) return null;
      o.target = raw.target;
      if (raw.back === false) o.back = false;
      const stroke = colorOrNone(raw.stroke);
      if (stroke) o.stroke = stroke;
      const strokeW = num(raw.strokeW, 0, 40);
      if (strokeW != null) o.strokeW = strokeW;
    }
    if (o.kind === "equation") {
      // 数式: kept as TeX (\frac{a}{b}, x^2, \sqrt{x}, \sum_{i=1}^{n}…), drawn as MathML.
      const tex = typeof raw.tex === "string" ? raw.tex.replace(/[\u0000-\u001f]/g, " ").slice(0, 2000) : "";
      if (!tex.trim()) return null;
      o.tex = tex;
      const fs = num(raw.fs, 12, 300);
      if (fs != null) o.fs = fs;
      const color = hexColor(raw.color);
      if (color) o.color = color;
      if (["left", "center", "right"].includes(raw.align)) o.align = raw.align;
    }
    if (o.kind === "model") {
      // 3D モデル: a GLB / glTF file (or a built-in model), its view, and a still picture of it (models.js).
      const src = typeof raw.src === "string" ? raw.src.trim() : "";
      if (!/^(data:(model\/gltf-binary|model\/gltf\+json|application\/octet-stream|application\/json)[;,]|idb:|https?:\/\/|blob:|builtin:[a-z]{2,16}$)/i.test(src) || src.length > 120_000_000) return null;
      o.src = src;
      if (typeof raw.fileName === "string" && raw.fileName.trim()) o.fileName = raw.fileName.trim().slice(0, 200);
      if (typeof raw.poster === "string" && /^data:image\/(png|jpeg|webp);base64,/.test(raw.poster) && raw.poster.length <= 3_000_000) o.poster = raw.poster;
      const v = raw.view && typeof raw.view === "object" ? raw.view : {};
      const view = { yaw: num(v.yaw, -360, 360, 0), pitch: num(v.pitch, -89, 89, 0), roll: num(v.roll, -180, 180, 0), zoom: num(v.zoom, 0.3, 4, 1), panX: num(v.panX, -2, 2, 0), panY: num(v.panY, -2, 2, 0) };
      const kept = Object.fromEntries(Object.entries(view).filter(([k, val]) => (k === "zoom" ? val !== 1 : val !== 0)).map(([k, val]) => [k, Math.round(val * 1000) / 1000]));
      if (Object.keys(kept).length) o.view = kept;
    }
    if (o.kind === "poll") {
      // アンケート (PowerPoint Live's polls): a question and its choices; the answers come in while presenting.
      const plainLine = (t, max) => String(t ?? "").replace(/<[^>]*>/g, "").replace(/[\u0000-\u001f]/g, " ").trim().slice(0, max);
      o.question = plainLine(raw.question, 200) || "質問";
      const options = (Array.isArray(raw.options) ? raw.options : []).map((t) => plainLine(t, 60)).filter(Boolean).slice(0, 8);
      if (options.length < 2) return null;
      o.options = options;
      const fill = hexColor(raw.fill);
      if (fill) o.fill = fill;
      const fs = num(raw.fs, 12, 120);
      if (fs != null) o.fs = fs;
    }
    if (o.kind === "camera") {
      // カメオ: the presenter's camera, live on the slide (a circle or another shape).
      if (["rect", "roundRect", "ellipse"].includes(raw.mask)) o.mask = raw.mask;
      const stroke = colorOrNone(raw.stroke);
      if (stroke) o.stroke = stroke;
      const strokeW = num(raw.strokeW, 0, 40);
      if (strokeW != null) o.strokeW = strokeW;
    }
    if (o.kind === "ink") {
      const strokes = normalizeStrokes(raw.strokes);
      if (!strokes.length) return null;
      o.strokes = strokes;
    }
    if (o.kind === "smartart") {
      o.smartart = normalizeSmartart(raw.smartart);
      const fs = num(raw.fs, 8, 200);
      if (fs != null) o.fs = fs;
    }
    if (o.kind === "table") {
      const table = normalizeTable(raw);
      if (!table) return null;
      Object.assign(o, table);
      // The words' size, colour, line spacing and font for the whole table (cells may set their own colour).
      const fs = num(raw.fs, 8, 400);
      if (fs != null) o.fs = fs;
      const color = hexColor(raw.color);
      if (color) o.color = color;
      const lh = num(raw.lh, 0.8, 4);
      if (lh != null) o.lh = lh;
      if (FONTS[raw.font] && raw.font !== "body") o.font = raw.font;
    }
    if (o.kind === "chart") {
      const chartData = normalizeChart(raw.chart);
      if (!chartData) return null;
      o.chart = chartData;
    }
    if (o.kind === "icon") {
      if (!E.icons[raw.icon]) return null;
      o.icon = raw.icon;
      const color = hexColor(raw.color);
      if (color) o.color = color;
    }
    // Clicking the object while presenting
    const action = raw.action;
    if (action && typeof action === "object") {
      const jump = jumpAction(action);
      if (jump) o.action = jump;
      else if (action.type === "popup") {
        // A card (or, with a breakdown or a source, a panel from the right) with words of its own.
        const pop = { type: "popup" };
        if (typeof action.title === "string" && action.title.trim()) pop.title = action.title.trim().slice(0, 80);
        const words = sanitizeRich(typeof action.text === "string" ? action.text : "");
        if (words && richToText(words).trim()) pop.text = words;
        const rows = (Array.isArray(action.rows) ? action.rows : []).filter((r) => r && String(r.label ?? "").trim()).slice(0, 8).map((r) => ({ label: String(r.label).trim().slice(0, 40), value: String(r.value ?? "").trim().slice(0, 40) }));
        if (rows.length) pop.rows = rows;
        if (typeof action.source === "string" && action.source.trim()) pop.source = action.source.trim().slice(0, 120);
        if (pop.title || pop.text || pop.rows) o.action = pop;
      } else if (action.type === "zoom" || action.type === "spot") o.action = { type: action.type };
      else if (action.type === "flip") {
        // The back of the card: its words and its colour (the front's colour when none is set).
        const back = sanitizeRich(typeof action.back === "string" ? action.back : "");
        o.action = { type: "flip", back: back && richToText(back).trim() ? back : "<p>（裏の文字）</p>" };
        const fill = hexColor(action.fill);
        if (fill) o.action.fill = fill;
      } else if (action.type === "reveal") {
        // Objects that wait hidden until this one is clicked (clicking again hides them); "only" hides what the
        // slide's other reveal buttons showed, so a row of them works as tabs.
        const targets = [...new Set((Array.isArray(action.targets) ? action.targets : []).filter((id) => typeof id === "string" && /^[A-Za-z0-9_-]{1,40}$/.test(id) && id !== o.id))].slice(0, 40);
        if (targets.length) o.action = { type: "reveal", targets, ...(action.only === true ? { only: true } : {}) };
      }
    }
    // マウスの通過 (PowerPoint's Mouse Over action): moving the pointer onto the object while presenting jumps to a
    // slide (next, previous, first, last, one of them), ends the show or plays a sound — not a link (a browser opens
    // a page only for a click).
    const over = raw.overAction && typeof raw.overAction === "object" ? jumpAction(raw.overAction) : null;
    if (over && over.type !== "url") o.overAction = over;
    if (IX_HOVERS[raw.hover]) o.hover = raw.hover;
    if (IX_LOOPS[raw.loop]) o.loop = raw.loop;
    if (typeof raw.tip === "string" && raw.tip.trim()) o.tip = raw.tip.trim().slice(0, 200);
    return o;
  }

  /** A click or mouse-over action that moves the show (or opens a link, or only plays a sound), with its sound:
   *  PowerPoint's 動作設定 (ハイパーリンク・サウンドの再生). null for anything else. */
  function jumpAction(action) {
    const sound = typeof action.sound === "string" && Object.hasOwn(E.TRANSITION_SOUNDS || {}, action.sound) ? { sound: action.sound } : {};
    if (["next", "prev", "first", "last", "end"].includes(action.type)) return { type: action.type, ...sound };
    if (action.type === "slide" && /^[A-Za-z0-9_-]{1,32}$/.test(action.to ?? "")) return { type: "slide", to: action.to, ...sound };
    if (action.type === "url" && /^(https?:\/\/|mailto:)/i.test(String(action.href || "").trim())) return { type: "url", href: String(action.href).trim().slice(0, 2000), ...sound };
    if (action.type === "sound" && sound.sound) return { type: "sound", ...sound };
    return null;
  }

  /** Column widths or row heights: n positive shares that add up to 1 (equal when unusable). */
  function shares(list, n) {
    const values = Array.from({ length: n }, (_, i) => Number(Array.isArray(list) ? list[i] : NaN));
    if (!values.every((v) => Number.isFinite(v) && v > 0)) return Array.from({ length: n }, () => Math.round((1 / n) * 100000) / 100000);
    const sum = values.reduce((a, b) => a + b, 0);
    return values.map((v) => Math.round((v / sum) * 100000) / 100000);
  }
  /** A table's cells (rich text, fill, colour, bold, alignment, merged spans), columns and rows, style options. */
  const TABLE_FILTER_OPS = new Set(["eq", "ne", "contains", "gt", "ge", "lt", "le", "nonblank", "blank"]);
  function normalizeTable(raw) {
    // Excel スプレッドシート: a table with `sheet` lets cells hold formulas (`f`), whose words (`text`) the editor's sheet.mjs works out.
    const sheet = raw.sheet === true;
    const rows = (Array.isArray(raw.cells) ? raw.cells : []).filter(Array.isArray).slice(0, 500);
    if (!rows.length) return null;
    const nCols = clamp(Math.max(...rows.map((row) => row.length), 1), 1, 100);
    const cells = rows.map((row) => Array.from({ length: nCols }, (_, c) => {
      const src = row[c] && typeof row[c] === "object" ? row[c] : {};
      const cell = {};
      if (typeof src.text === "string" && src.text.trim()) { const text = sanitizeRich(src.text); if (text.replace(/<[^>]+>/g, "").trim()) cell.text = text; }
      const fill = hexColor(src.fill);
      if (fill) cell.fill = fill;
      const color = hexColor(src.color);
      if (color) cell.color = color;
      const fs = num(src.fs, 8, 400);
      if (fs != null) cell.fs = fs;
      const lh = num(src.lh, 0.8, 4);
      if (lh != null) cell.lh = lh;
      if (typeof src.fontFace === "string" && /^[\p{L}\p{N}\s._+\-]{1,100}$/u.test(src.fontFace)) cell.fontFace = src.fontFace;
      if (Array.isArray(src.pad) && src.pad.length === 4) cell.pad = src.pad.map((v) => num(v, 0, 400, 0));
      for (const key of ["bold", "italic", "underline", "strike"]) if (src[key] === true) cell[key] = true;
      if (ALIGN_ALL.has(src.align)) cell.align = src.align;
      if (["top", "middle", "bottom"].includes(src.valign)) cell.valign = src.valign;
      if (src.vertical === true) cell.vertical = true;
      if (sheet && typeof src.f === "string" && src.f.trim().startsWith("=") && src.f.trim().length <= 500) cell.f = src.f.trim();
      // How a formula's number is shown: decimals, thousands commas, a percentage, a ¥ or $ in front.
      const dec = Math.round(Number(src.dec));
      if (src.dec != null && Number.isFinite(dec) && dec >= 0 && dec <= 6) cell.dec = dec;
      if (src.sep === true) cell.sep = true;
      if (src.pct === true) cell.pct = true;
      if (src.cur === "¥" || src.cur === "$") cell.cur = src.cur;
      const rs = Math.round(Number(src.rs) || 1);
      const cs = Math.round(Number(src.cs) || 1);
      if (rs > 1) cell.rs = rs;
      if (cs > 1) cell.cs = cs;
      // Borders a cell draws itself (a table brought over from PowerPoint): { c: colour, w: width } or "none".
      for (const side of ["bt", "br", "bb", "bl"]) {
        const b = src[side];
        if (b === "none") cell[side] = "none";
        else if (b && typeof b === "object" && hexColor(b.c)) cell[side] = { c: hexColor(b.c), w: num(b.w, 0.5, 24, 1) };
      }
      return cell;
    }));
    // Merged cells: a span stays inside the table, the cells it covers are marked, and spans never overlap.
    const covered = cells.map((row) => row.map(() => false));
    cells.forEach((row, r) => row.forEach((cell, c) => {
      if (covered[r][c]) { for (const key of Object.keys(cell)) delete cell[key]; cell.merged = true; return; }
      const rs = clamp(cell.rs || 1, 1, cells.length - r);
      const cs = clamp(cell.cs || 1, 1, nCols - c);
      let fits = true;
      for (let i = r; i < r + rs; i += 1) for (let j = c; j < c + cs; j += 1) if ((i !== r || j !== c) && covered[i][j]) fits = false;
      if (!fits) { delete cell.rs; delete cell.cs; return; }
      if (rs > 1) cell.rs = rs; else delete cell.rs;
      if (cs > 1) cell.cs = cs; else delete cell.cs;
      for (let i = r; i < r + rs; i += 1) for (let j = c; j < c + cs; j += 1) if (i !== r || j !== c) covered[i][j] = true;
    }));
    const out = { cells, cols: shares(raw.cols, nCols), rows: shares(raw.rows, cells.length) };
    if (TABLE_STYLES[raw.style] && raw.style !== "medium2") out.style = raw.style;
    for (const key of ["header", "banded", "firstCol", "lastRow"]) if (typeof raw[key] === "boolean") out[key] = raw[key];
    if (sheet) {
      out.sheet = true;
      // フィルター: one column's condition; the rows it hides (`hide`) are worked out again whenever the sheet is normalised
      // (the editor's hook), and kept as they were saved where there is none (an exported show).
      const flt = raw.filter;
      if (flt && typeof flt === "object" && Number.isInteger(flt.col) && flt.col >= 0 && flt.col < nCols && TABLE_FILTER_OPS.has(flt.op)) {
        out.filter = { col: flt.col, op: flt.op };
        if (typeof flt.value === "string" && flt.value.trim()) out.filter.value = flt.value.trim().slice(0, 100);
        const hide = [...new Set((Array.isArray(raw.hide) ? raw.hide : []).filter((r) => Number.isInteger(r) && r >= 0 && r < cells.length))].sort((a, b) => a - b);
        if (hide.length && hide.length < cells.length) out.hide = hide;
      }
      E.sheetCalc?.(out);
    }
    return out;
  }
  // As many categories and series as a PowerPoint chart brought over may carry.
  const CHART_MAX_LABELS = 500;
  const CHART_MAX_SERIES = 100;
  // The chart kinds that take 近似曲線, 棒の間隔・重なり, 項目を逆順, 折れ線の書式 and 目盛間隔.
  const TREND_CHART_KINDS = new Set(["bar", "line", "multi-line", "scatter"]);
  const BAR_GAP_CHARTS = new Set(["bar", "combo", "pareto", "clustered-bar", "stacked-bar", "100-stacked-bar", "hbar", "stacked-hbar"]);
  const REVERSE_CHARTS = new Set(["bar", "combo", "clustered-bar", "stacked-bar", "100-stacked-bar", "line", "multi-line", "area", "stacked-area", "hbar", "stacked-hbar", "pie", "donut", "funnel", "radar"]);
  const LINE_CHART_KINDS = new Set(["line", "multi-line"]);
  const STEP_CHARTS = new Set(["line", "multi-line", "area", "scatter"]);
  const MARKER_SHAPES = ["none", "circle", "square", "diamond", "triangle"];
  const LABEL_ALL_CHARTS = new Set(["line", "multi-line", "area", "clustered-bar"]);
  const LABEL_POS_CHARTS = new Set(["line", "multi-line", "area"]);
  const SLICE_LABEL_PARTS = ["category", "value", "percent"];
  /** A chart's kind, category labels and series (numbers), title and unit. */
  function normalizeChart(raw) {
    if (!raw || typeof raw !== "object") return null;
    const labels = (Array.isArray(raw.labels) ? raw.labels : []).slice(0, CHART_MAX_LABELS).map((label) => String(label ?? "").trim().slice(0, 500));
    // 箱ひげ図 and ヒストグラム read samples: a group may be shorter than the others, its empty cells stay empty (null).
    const gappy = raw.type === "boxplot" || raw.type === "histogram";
    const series = (Array.isArray(raw.series) ? raw.series : []).filter((s) => s && typeof s === "object").slice(0, CHART_MAX_SERIES).map((s, i) => ({
      name: String(s.name ?? "").trim().slice(0, 500) || `系列${i + 1}`,
      values: labels.map((_, j) => {
        const r = Array.isArray(s.values) ? s.values[j] : null;
        if (gappy && (r == null || r === "")) return null;
        const v = Number(r);
        return Number.isFinite(v) ? Math.round(clamp(v, -1e12, 1e12) * 10000) / 10000 : gappy ? null : 0;
      }),
    }));
    if (!labels.length || !series.length) return null;
    const out = { type: CHART_KINDS[raw.type] ? raw.type : "bar", labels, series };
    for (const key of ["title", "unit"]) if (typeof raw[key] === "string" && raw[key].trim()) out[key] = raw[key].trim().slice(0, key === "unit" ? 10 : 80);
    // 色の変更: one of the colour sets (カラフル・モノクロ; none is the theme's own order).
    if (CHART_COLORS[raw.colors]) out.colors = raw.colors;
    // グラフ要素を追加 (and クイック レイアウト): data labels off, the legend's place (or none), gridlines off.
    const opts = {};
    if (raw.opts?.labels === false) opts.labels = false;
    if (["bottom", "right", "none"].includes(raw.opts?.legend)) opts.legend = raw.opts.legend;
    if (raw.opts?.grid === false) opts.grid = false;
    // 軸ラベル and 近似曲線 (グラフ要素).
    for (const key of ["axisX", "axisY"]) if (typeof raw.opts?.[key] === "string" && raw.opts[key].trim()) opts[key] = raw.opts[key].trim().slice(0, 40);
    // 近似曲線: a kind (線形・指数・対数・多項式・累乗・移動平均), its equation and R², a moving average's span.
    const trend = raw.opts?.trend === true ? "linear" : raw.opts?.trend;
    if (TREND_CHART_KINDS.has(out.type) && E.TREND_KINDS[trend]) {
      opts.trend = trend;
      if (raw.opts.trendEq === true && trend !== "movavg") opts.trendEq = true;
      if (raw.opts.trendR2 === true && trend !== "movavg") opts.trendR2 = true;
      const period = Math.round(Number(raw.opts.trendPeriod));
      if (trend === "movavg" && period >= 2 && period <= 12 && period !== 2) opts.trendPeriod = period;
    }
    // 棒の間隔の幅・系列の重なり (% of a bar), 項目を逆順, 折れ線の滑らかさとマーカー, 円の角度・ドーナツの穴・切り出し, 目盛間隔.
    const whole = (v, lo, hi) => { const n = v == null || v === "" ? NaN : Math.round(Number(v)); return Number.isFinite(n) && n >= lo && n <= hi ? n : null; };
    if (BAR_GAP_CHARTS.has(out.type)) {
      const gap = whole(raw.opts?.gap, 0, 500);
      const overlap = whole(raw.opts?.overlap, -100, 100);
      if (gap != null) opts.gap = gap;
      if (overlap != null && (out.type === "clustered-bar" || out.type === "hbar") && series.length > 1) opts.overlap = overlap;
    }
    if (raw.opts?.reverse === true && REVERSE_CHARTS.has(out.type)) opts.reverse = true;
    if (LINE_CHART_KINDS.has(out.type)) {
      if (raw.opts?.smooth === true) opts.smooth = true;
      if (MARKER_SHAPES.includes(raw.opts?.marker) && raw.opts.marker !== "circle") opts.marker = raw.opts.marker;
    }
    if (out.type === "pie" || out.type === "donut") {
      const angle = whole(raw.opts?.angle, 1, 359);
      const explode = whole(raw.opts?.explode, 1, 40);
      const hole = whole(raw.opts?.hole, 10, 90);
      if (angle != null) opts.angle = angle;
      if (explode != null) opts.explode = explode;
      if (hole != null && out.type === "donut") opts.hole = hole;
    }
    // データ ラベル: every point labelled (すべての点), where a line's labels sit, and a pie's labels outside its slices.
    if (raw.opts?.labelAll === true && LABEL_ALL_CHARTS.has(out.type)) opts.labelAll = true;
    if (["below", "right"].includes(raw.opts?.labelPos) && LABEL_POS_CHARTS.has(out.type)) opts.labelPos = raw.opts.labelPos;
    if ((out.type === "pie" || out.type === "donut") && Array.isArray(raw.opts?.sliceLabels)) {
      const parts = SLICE_LABEL_PARTS.filter((part) => raw.opts.sliceLabels.includes(part));
      if (parts.length) opts.sliceLabels = parts;
    }
    if (STEP_CHARTS.has(out.type)) {
      const step = Number(raw.opts?.axisStep);
      if (Number.isFinite(step) && step > 0 && step < 1e12) opts.axisStep = step;
    }
    // 数値の書式 (グラフ要素): decimals, a display unit (千・万・百万・億) and signs before and after the number.
    const nf = raw.opts?.numFmt;
    if (nf && typeof nf === "object") {
      const numFmt = {};
      const decimals = Number(nf.decimals);
      if (nf.decimals != null && nf.decimals !== "" && Number.isInteger(decimals) && decimals >= 0 && decimals <= 4) numFmt.decimals = decimals;
      if ([1000, 10000, 1000000, 100000000].includes(Number(nf.scale))) numFmt.scale = Number(nf.scale);
      for (const [key, max] of [["prefix", 4], ["suffix", 6]]) if (typeof nf[key] === "string" && nf[key].trim()) numFmt[key] = nf[key].replace(/[\u0000-\u001f<>]/g, "").trim().slice(0, max);
      if (Object.keys(numFmt).length) opts.numFmt = numFmt;
    }
    // 誤差範囲 (グラフ要素): a fixed amount or a share of each value, on the bars and lines of a chart with axes.
    const eb = raw.opts?.errorBars;
    if (eb && typeof eb === "object" && ["fixed", "percent"].includes(eb.type) && ["bar", "combo", "clustered-bar", "line", "multi-line"].includes(out.type)) {
      const amount = Number(eb.amount);
      if (Number.isFinite(amount) && amount > 0 && amount <= (eb.type === "percent" ? 100 : 1e12)) opts.errorBars = { type: eb.type, amount };
    }
    // 第2軸: a combination chart's line on a scale of its own (drawn down the right side).
    if (raw.opts?.axis2 === true && out.type === "combo" && series.length >= 2) opts.axis2 = true;
    // 軸の書式: the value axis's own minimum and maximum (a maximum above the minimum).
    const axisMin = raw.opts?.axisMin != null && raw.opts.axisMin !== "" ? Number(raw.opts.axisMin) : NaN;
    const axisMax = raw.opts?.axisMax != null && raw.opts.axisMax !== "" ? Number(raw.opts.axisMax) : NaN;
    if (Number.isFinite(axisMin) && Math.abs(axisMin) < 1e12) opts.axisMin = axisMin;
    if (Number.isFinite(axisMax) && Math.abs(axisMax) < 1e12 && axisMax > (opts.axisMin ?? -Infinity)) opts.axisMax = axisMax;
    // ウォーターフォール: the bars that show the running total (PowerPoint's 合計として設定), besides labels like 合計.
    if (Array.isArray(raw.opts?.totals)) {
      const totals = [...new Set(raw.opts.totals.map(Number).filter((i) => Number.isInteger(i) && i >= 0 && i < labels.length))].sort((a, b) => a - b).slice(0, 50);
      if (totals.length) opts.totals = totals;
    }
    // グラフ フィルター: series and categories kept in the data but not drawn (at least one of each stays shown).
    for (const [key, count] of [["hideSeries", series.length], ["hideLabels", labels.length]]) {
      if (!Array.isArray(raw.opts?.[key])) continue;
      const list = [...new Set(raw.opts[key].map(Number).filter((i) => Number.isInteger(i) && i >= 0 && i < count))].sort((a, b) => a - b);
      if (list.length && list.length < count) opts[key] = list;
    }
    // データ テーブル: the values in a table under the chart.
    if (raw.opts?.table === true) opts.table = true;
    if (Object.keys(opts).length) out.opts = opts;
    const style = normalizeChartStyle(raw.style, labels.length);
    if (style) out.style = style;
    return out;
  }

  // ---------------------------------------------------------------- charts as PowerPoint draws them
  // A chart brought over from PowerPoint keeps its formatting in `chart.style`: the colour of each series and
  // point, gap width and overlap, line widths and markers, data labels (number format, position, font, custom
  // text), axes (labels, lines, scale, gridlines, reversed order), the legend and the plot area. Such a chart is
  // drawn the way PowerPoint draws it rather than with the layouts' look. Sizes are slide pixels.

  const LABEL_POS = new Set(["outEnd", "inEnd", "ctr", "inBase", "t", "b", "l", "r", "bestFit"]);
  const MARKERS = new Set(["none", "circle", "square", "diamond", "triangle", "dash", "dot", "x", "plus", "star"]);
  const SERIES_KINDS = new Set(["bar", "line", "area"]);

  function normalizeChartStyle(raw, categoryCount = CHART_MAX_LABELS) {
    if (!raw || typeof raw !== "object") return null;
    const font = (f) => {
      if (!f || typeof f !== "object") return null;
      const out = {};
      const size = num(f.size, 6, 200);
      if (size != null) out.size = size;
      if (typeof f.bold === "boolean") out.bold = f.bold;
      if (f.italic === true) out.italic = true;
      const color = hexColor(f.color);
      if (color) out.color = color;
      return Object.keys(out).length ? out : null;
    };
    const lineColor = (v) => (v === "none" ? "none" : hexColor(v));
    const fmt = (v) => (typeof v === "string" && v.trim() ? v.slice(0, 60) : null);
    const label = (l) => {
      if (!l || typeof l !== "object") return null;
      const out = {};
      for (const key of ["val", "pct", "cat", "ser"]) if (l[key] === true) out[key] = true;
      if (l.show === false) out.show = false;
      if (LABEL_POS.has(l.pos)) out.pos = l.pos;
      if (fmt(l.format)) out.format = fmt(l.format);
      const f = font(l.font);
      if (f) out.font = f;
      if (Array.isArray(l.runs)) {
        const runs = l.runs.filter((r) => r && typeof r.t === "string").slice(0, 8).map((r) => ({ t: r.t.slice(0, 80), ...(font(r.font) ? { font: font(r.font) } : {}) }));
        if (runs.length) out.runs = runs;
      }
      return Object.keys(out).length ? out : null;
    };
    const st = {};
    if (CHART_KINDS[raw.type]) st.type = raw.type;
    if (raw.dir === "bar" || raw.dir === "col") st.dir = raw.dir;
    if (raw.stack === "stacked" || raw.stack === "percent") st.stack = raw.stack;
    const gap = num(raw.gap, 0, 500);
    if (gap != null) st.gap = gap;
    const overlap = num(raw.overlap, -100, 100);
    if (overlap != null) st.overlap = overlap;
    const hole = num(raw.hole, 0, 90);
    if (hole != null) st.hole = hole;
    const angle = num(raw.angle, 0, 360);
    if (angle != null) st.angle = angle;
    for (const key of ["font", "title"]) { const f = font(raw[key]); if (f) st[key] = f; }
    if (raw.legend && typeof raw.legend === "object" && ["b", "t", "r", "l", "tr"].includes(raw.legend.pos)) {
      st.legend = { pos: raw.legend.pos, ...(font(raw.legend.font) ? { font: font(raw.legend.font) } : {}) };
    }
    for (const key of ["cat", "val"]) {
      const a = raw[key];
      if (!a || typeof a !== "object") continue;
      const out = {};
      if (a.hide === true) out.hide = true;
      if (a.reverse === true) out.reverse = true;
      if (a.edge === true) out.edge = true;
      const f = font(a.font);
      if (f) out.font = f;
      const line = lineColor(a.line);
      if (line) out.line = line;
      const grid = lineColor(a.grid);
      if (grid) out.grid = grid;
      for (const k of ["min", "max", "step"]) { const v = num(a[k], -1e12, 1e12); if (v != null) out[k] = v; }
      if (out.step != null && out.step <= 0) delete out.step;
      if (fmt(a.format)) out.format = fmt(a.format);
      if (Object.keys(out).length) st[key] = out;
    }
    if (raw.plot && typeof raw.plot === "object") {
      const p = { x: num(raw.plot.x, 0, 1), y: num(raw.plot.y, 0, 1), w: num(raw.plot.w, 0.05, 1), h: num(raw.plot.h, 0.05, 1) };
      if (Object.values(p).every((v) => v != null)) st.plot = { ...p, ...(raw.plot.outer === true ? { outer: true } : {}) };
    }
    if (Array.isArray(raw.series)) {
      st.series = raw.series.slice(0, CHART_MAX_SERIES).map((sr) => {
        if (!sr || typeof sr !== "object") return {};
        const out = {};
        if (SERIES_KINDS.has(sr.kind)) out.kind = sr.kind;
        const color = hexColor(sr.color);
        if (color) out.color = color;
        const width = num(sr.width, 0, 60);
        if (width != null) out.width = width;
        if (DASHES[sr.dash] && sr.dash !== "solid") out.dash = sr.dash;
        if (sr.smooth === true) out.smooth = true;
        if (sr.marker && typeof sr.marker === "object") {
          const m = { s: MARKERS.has(sr.marker.s) ? sr.marker.s : "none" };
          const z = num(sr.marker.z, 2, 80);
          if (z != null) m.z = z;
          const mc = hexColor(sr.marker.color);
          if (mc) m.color = mc;
          if (Array.isArray(sr.marker.at)) m.at = sr.marker.at.map(Number).filter((i) => Number.isInteger(i) && i >= 0 && i < categoryCount).slice(0, categoryCount);
          out.marker = m;
        }
        if (Array.isArray(sr.points)) {
          const pts = sr.points.filter((p) => p && Number.isInteger(p.i) && p.i >= 0 && p.i < categoryCount && hexColor(p.color)).slice(0, categoryCount).map((p) => ({ i: p.i, color: hexColor(p.color) }));
          if (pts.length) out.points = pts;
        }
        const l = label(sr.label);
        if (l) out.label = l;
        if (Array.isArray(sr.pointLabels)) {
          const pls = sr.pointLabels.filter((p) => p && Number.isInteger(p.i) && p.i >= 0 && p.i < categoryCount).slice(0, categoryCount).map((p) => ({ i: p.i, ...(label(p) || {}) }));
          if (pls.length) out.pointLabels = pls;
        }
        return out;
      });
    }
    return Object.keys(st).length ? st : null;
  }

  /** A number in an Excel number format: 0, 0.0, #,##0, 0%, 0.0%, 0"%", "¥"#,##0, [Red] and ; sections, General. */
  function numFormat(value, code) {
    const n = Number(value);
    if (!Number.isFinite(n)) return String(value ?? "");
    const general = (v) => {
      const a = Math.abs(v);
      if (a !== 0 && (a >= 1e11 || a < 1e-9)) return v.toExponential(4).replace(/\.?0+e/, "E");
      return String(Math.round(v * 1e9) / 1e9);
    };
    if (!code || /^general$/i.test(code.trim())) return general(n);
    const sections = [];
    let cur = "";
    let quoted = false;
    for (const ch of code) {
      if (ch === '"') quoted = !quoted;
      if (ch === ";" && !quoted) { sections.push(cur); cur = ""; } else cur += ch;
    }
    sections.push(cur);
    let section = sections[0];
    let neg = n < 0;
    if (n < 0 && sections.length > 1) { section = sections[1]; neg = false; } else if (n === 0 && sections.length > 2) section = sections[2];
    if (/^general$/i.test(section.trim())) return (neg ? "-" : "") + general(Math.abs(n));
    // Split into literal text and the one numeric placeholder block.
    const parts = [];
    let block = null;
    let percent = 0;
    for (let i = 0; i < section.length; i += 1) {
      const ch = section[i];
      if (ch === '"') { const end = section.indexOf('"', i + 1); parts.push({ t: section.slice(i + 1, end < 0 ? undefined : end) }); i = end < 0 ? section.length : end; continue; }
      if (ch === "\\" && i + 1 < section.length) { parts.push({ t: section[i + 1] }); i += 1; continue; }
      if (ch === "[") { const end = section.indexOf("]", i); i = end < 0 ? section.length : end; continue; }
      if (ch === "_" || ch === "*") { i += 1; if (ch === "_") parts.push({ t: " " }); continue; }
      if ("0#?,.".includes(ch) && (block === null || parts.at(-1) === block)) {
        if (block === null) { block = { num: "" }; parts.push(block); }
        block.num += ch;
        continue;
      }
      if (ch === "%") percent += 1;
      if (/[eE]/.test(ch) && block) return (neg ? "-" : "") + general(Math.abs(n));
      parts.push({ t: ch });
    }
    let v = Math.abs(n) * 100 ** percent;
    if (!block) return (neg ? "-" : "") + parts.map((p) => p.t ?? "").join("");
    const [intPart, decPart = ""] = block.num.split(".");
    const scale = (/,+$/.exec(intPart) || [""])[0].length;
    v /= 1000 ** scale;
    const minDec = (decPart.match(/0/g) || []).length;
    const maxDec = (decPart.match(/[0#?]/g) || []).length;
    const grouping = intPart.replace(/,+$/, "").includes(",");
    const minInt = Math.max(1, (intPart.match(/0/g) || []).length);
    let text = v.toLocaleString("en-US", { minimumIntegerDigits: Math.min(21, minInt), minimumFractionDigits: minDec, maximumFractionDigits: maxDec, useGrouping: grouping });
    if (!/0/.test(intPart) && text.startsWith("0") && v < 1 && maxDec) text = text.slice(1);
    block.t = text;
    return (neg && Math.round(v * 10 ** maxDec) !== 0 ? "-" : "") + parts.map((p) => p.t ?? "").join("");
  }

  /** About how wide a text is drawn (CJK full width, Latin about half). */
  function textWidth(text, size) {
    let w = 0;
    for (const ch of String(text)) {
      const c = ch.codePointAt(0);
      w += c < 0x2e80 ? (/[ilIjt.,:;'|!()[\]\s]/.test(ch) ? 0.32 : /[mwMW%@]/.test(ch) ? 0.86 : /[A-Z0-9#$&]/.test(ch) ? 0.62 : 0.54) : c >= 0xff61 && c <= 0xff9f ? 0.5 : 1;
    }
    return w * size;
  }

  /** A label broken into lines no wider than `width` (at most `max` lines). */
  function wrapText(text, size, width, max = 3) {
    const out = [];
    let line = "";
    for (const ch of String(text)) {
      if (line && textWidth(line + ch, size) > width) {
        out.push(line);
        line = ch.trim() ? ch : "";
        if (out.length === max - 1) { line += [...String(text)].slice([...out.join("")].length + 1).join(""); break; }
      } else line += ch;
    }
    if (line) out.push(line);
    if (out.length === max && textWidth(out.at(-1), size) > width) {
      let last = out.at(-1);
      while (last.length > 1 && textWidth(`${last}…`, size) > width) last = last.slice(0, -1);
      out[out.length - 1] = `${last}…`;
    }
    return out;
  }

  /** The value axis PowerPoint picks: 5% headroom, then the smallest round step (1, 2, 5 × 10ⁿ) that gives at most
   *  ten divisions (fewer on a short axis); zero stays in unless the values sit far from it. */
  function valueScale(lo, hi, opt = {}) {
    const fixedMin = opt.min != null;
    const fixedMax = opt.max != null;
    let min = fixedMin ? opt.min : lo >= 0 ? (hi > 0 && (hi - lo) / hi < 1 / 6 ? lo : 0) : lo;
    let max = fixedMax ? opt.max : hi <= 0 && lo < 0 ? 0 : hi;
    if (!fixedMax && max > 0) max += (max - Math.min(min, max)) * 0.05;
    if (!fixedMin && min < 0) min -= (max - min) * 0.05;
    if (max <= min) max = min + 1;
    const most = Math.max(2, Math.min(10, Math.floor(opt.divisions || 10)));
    let step = opt.step;
    if (!step) {
      const range = max - min;
      const exp = 10 ** Math.floor(Math.log10(range / most));
      step = [1, 2, 5, 10, 20].map((m) => m * exp).find((v) => range / v <= most + 1e-9) || 10 * exp;
    }
    if (!fixedMin) min = Math.floor(min / step + 1e-9) * step;
    if (!fixedMax) max = Math.max(min + step, Math.ceil(max / step - 1e-9) * step);
    return { min, max, step };
  }

  /** The chart a PowerPoint deck carried, drawn in SVG with its own formatting (chart.style). */
  function officeChart(c, W0, H0) {
    const st = c.style || {};
    const W = Math.max(60, W0);
    const H = Math.max(40, H0);
    const svg = s("svg", { class: "hs-ochart", viewBox: `0 0 ${r2(W)} ${r2(H)}`, preserveAspectRatio: "none", role: "img", "aria-label": c.title || "グラフ" });
    const base = { size: 24, color: "#595959", ...(st.font || {}) };
    const fontOf = (...fs) => Object.assign({}, base, ...fs.filter(Boolean));
    const textEl = (x, y, str, f, anchor = "start", extra = {}) => s("text", { x: r2(x), y: r2(y), "text-anchor": anchor, fill: f.color, style: { "font-size": `${f.size}px`, "font-weight": f.bold ? 700 : 400, "font-style": f.italic ? "italic" : null }, ...extra }, str);
    const sameType = !st.type || st.type === c.type;
    const sts = (i) => (sameType ? st.series?.[i] : { color: st.series?.[i]?.color }) || {};
    const palette = ["var(--c1)", "var(--c2)", "var(--c3)", "var(--c4)", "var(--c5)", "var(--c6)"];
    const colorOf = (i) => sts(i).color || palette[i % palette.length];
    const pointColor = (si, i) => sts(si).points?.find((p) => p.i === i)?.color || colorOf(si);
    const isPie = c.type === "donut";
    const baseKind = /line/.test(c.type) ? "line" : "bar";
    const kindOf = (i) => sts(i).kind || (c.type === "combo" ? (i === 0 ? "bar" : "line") : baseKind);
    const stack = sameType && st.stack ? st.stack : c.type === "100-stacked-bar" ? "percent" : c.type === "stacked-bar" ? "stacked" : null;
    const horizontal = sameType && st.dir === "bar";
    const n = c.labels.length;
    const pad = 10;
    let box = { x: pad, y: pad, w: W - pad * 2, h: H - pad * 2 };
    // Title
    if (c.title) {
      const f = fontOf({ size: Math.round(base.size * 1.2), bold: true }, st.title);
      svg.append(textEl(W / 2, box.y + f.size, c.title, f, "middle", { class: "hs-ochart-title" }));
      box = { ...box, y: box.y + f.size * 1.5, h: box.h - f.size * 1.5 };
    }
    // Legend
    const legendItems = isPie ? c.labels.map((label, i) => ({ label, color: pointColor(0, i), kind: "bar" })) : c.series.map((sr, i) => ({ label: sr.name, color: colorOf(i), kind: kindOf(i), i }));
    if (st.legend && legendItems.length) {
      const f = fontOf(st.legend.font);
      const sw = f.size * 0.75;
      const itemW = (it) => sw + f.size * 0.4 + textWidth(it.label, f.size) + f.size * 1.1;
      const g = s("g", { class: "hs-ochart-legend" });
      // While presenting, a series' key shows or hides it (motion.js); a pie's keys only name its slices.
      const keys = !isPie && legendItems.length > 1;
      const drawItem = (it, x, y) => {
        const key = keys ? s("g", { class: "hs-ochart-key hs-control", "data-series": String(it.i) }) : g;
        if (keys) g.append(key);
        if (keys) key.append(s("rect", { x: r2(x - 6), y: r2(y - f.size * 0.75), width: r2(itemW(it) - f.size * 0.6), height: r2(f.size * 1.5), fill: "transparent" }));
        if (it.kind === "line") {
          key.append(s("line", { x1: r2(x), x2: r2(x + sw * 1.6), y1: r2(y), y2: r2(y), stroke: it.color, "stroke-width": Math.max(2, sts(it.i).width || 3) }));
          x += sw * 0.6;
        } else key.append(s("rect", { x: r2(x), y: r2(y - sw / 2), width: r2(sw), height: r2(sw), fill: it.color }));
        key.append(textEl(x + sw + f.size * 0.4, y + f.size * 0.36, it.label, f));
      };
      const pos = st.legend.pos;
      if (pos === "b" || pos === "t") {
        const rows = [[]];
        let rowW = 0;
        for (const it of legendItems) {
          const w = itemW(it);
          if (rowW + w > box.w && rows.at(-1).length) { rows.push([]); rowW = 0; }
          rows.at(-1).push(it);
          rowW += w;
        }
        const lh = f.size * 1.5;
        const y0 = pos === "b" ? box.y + box.h - lh * rows.length + lh / 2 : box.y + lh / 2;
        rows.forEach((row, ri) => {
          const total = row.reduce((sum, it) => sum + itemW(it), 0) - f.size * 1.1;
          let x = box.x + (box.w - total) / 2;
          for (const it of row) { drawItem(it, x, y0 + ri * lh); x += itemW(it); }
        });
        box = pos === "b" ? { ...box, h: box.h - lh * rows.length - 4 } : { ...box, y: box.y + lh * rows.length + 4, h: box.h - lh * rows.length - 4 };
      } else {
        const lw = Math.min(box.w * 0.4, Math.max(...legendItems.map(itemW)));
        const lh = f.size * 1.45;
        const total = lh * legendItems.length;
        const x0 = pos === "l" ? box.x : box.x + box.w - lw;
        const y0 = pos === "tr" ? box.y + lh / 2 : box.y + (box.h - total) / 2 + lh / 2;
        legendItems.forEach((it, i) => drawItem(it, x0, y0 + i * lh));
        box = pos === "l" ? { ...box, x: box.x + lw + 8, w: box.w - lw - 8 } : { ...box, w: box.w - lw - 8 };
      }
      svg.append(g);
    }
    const manual = st.plot ? { x: st.plot.x * W, y: st.plot.y * H, w: st.plot.w * W, h: st.plot.h * H } : null;
    if (isPie) {
      pieChart(svg, c, manual || box, { st, sts, pointColor, fontOf, textEl });
      return svg;
    }
    // Values and the value axis
    const valuesOf = (si) => c.series[si].values;
    const barIdx = c.series.map((_, i) => i).filter((i) => kindOf(i) === "bar");
    const lineIdx = c.series.map((_, i) => i).filter((i) => kindOf(i) !== "bar");
    const totals = c.labels.map((_, i) => {
      let pos = 0, neg = 0;
      for (const si of barIdx) { const v = valuesOf(si)[i] || 0; if (v >= 0) pos += v; else neg += v; }
      return { pos, neg };
    });
    const shown = (si, i) => {
      const v = valuesOf(si)[i] || 0;
      if (stack === "percent" && kindOf(si) === "bar") { const t = totals[i].pos - totals[i].neg; return t ? (v / t) * 100 : 0; }
      return v;
    };
    let lo = 0, hi = 0;
    const consider = (v) => { lo = Math.min(lo, v); hi = Math.max(hi, v); };
    let first = true;
    const take = (v) => { if (first) { lo = v; hi = v; first = false; } else consider(v); };
    if (stack && barIdx.length) for (const t of totals) { if (stack === "percent") { take(t.neg ? -100 : 0); take(t.pos ? 100 : 0); } else { take(t.pos); take(t.neg); } }
    else for (const si of barIdx) valuesOf(si).forEach((v) => take(v));
    for (const si of lineIdx) valuesOf(si).forEach((v) => take(v));
    if (first) { lo = 0; hi = 1; }
    if (barIdx.length) { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
    const va = st.val || {};
    const ca = st.cat || {};
    const valAxisLen = horizontal ? (manual || box).w * 0.7 : (manual || box).h * 0.8;
    const percentAxis = stack === "percent" && !lineIdx.length;
    const scale = valueScale(lo, hi, { min: va.min ?? (percentAxis && lo >= 0 ? 0 : undefined), max: va.max ?? (percentAxis ? (hi > 0 ? 100 : 0) : undefined), step: va.step, divisions: valAxisLen / (fontOf(va.font).size * 1.6) });
    // A 100% stacked chart's axis runs from 0 to 1 in Excel (shown as 0%–100%); ours from 0 to 100.
    const tickText = (v) => (percentAxis ? numFormat(v / 100, va.format || "0%") : numFormat(v, va.format));
    const ticks = [];
    for (let v = scale.min, k = 0; v <= scale.max + scale.step * 1e-6 && k < 60; v += scale.step, k += 1) ticks.push(Math.round(v / scale.step) * scale.step);
    const valFont = fontOf(va.font);
    const catFont = fontOf(ca.font);
    // The plot area: as the chart placed it, or what is left after the axis labels.
    let plot;
    let labelLeft = (manual || box).x;
    if (manual && !st.plot.outer) plot = manual;
    else {
      const area = manual || box;
      labelLeft = area.x;
      const valLabelW = va.hide ? 0 : Math.max(...ticks.map((v) => textWidth(tickText(v), valFont.size))) + 10;
      if (horizontal) {
        const catW = ca.hide ? 0 : Math.min(area.w * 0.5, Math.max(...c.labels.map((l) => textWidth(l, catFont.size)), 0) + 12);
        const valH = va.hide ? 0 : valFont.size * 1.5;
        plot = { x: area.x + catW, y: area.y + (va.hide ? 0 : 0), w: area.w - catW - (va.hide ? 0 : valFont.size), h: area.h - valH };
      } else {
        const slotW = (area.w - valLabelW) / Math.max(1, n);
        const lines = ca.hide ? 0 : Math.max(1, ...c.labels.map((l) => wrapText(l, catFont.size, slotW * 0.92).length));
        plot = { x: area.x + valLabelW, y: area.y + (va.hide ? 4 : valFont.size * 0.6), w: area.w - valLabelW - 6, h: area.h - lines * catFont.size * 1.25 - 8 - (va.hide ? 4 : valFont.size * 0.6) };
      }
    }
    plot.w = Math.max(10, plot.w);
    plot.h = Math.max(10, plot.h);
    // Positions along the value axis and the category axis
    const valLen = horizontal ? plot.w : plot.h;
    const vPos = (v) => {
      const t = (clamp(v, scale.min, scale.max) - scale.min) / (scale.max - scale.min || 1);
      return horizontal ? plot.x + t * valLen : plot.y + plot.h - t * valLen;
    };
    const edge = sameType && ca.edge && !barIdx.length;
    const slot = (horizontal ? plot.h : plot.w) / Math.max(1, edge ? n - 1 || 1 : n);
    // Category i's centre along its axis; PowerPoint draws a bar chart's first category at the bottom.
    const order = (i) => (horizontal ? (ca.reverse ? i : n - 1 - i) : ca.reverse ? n - 1 - i : i);
    const cPos = (i) => (horizontal ? plot.y : plot.x) + (edge ? order(i) * slot : (order(i) + 0.5) * slot);
    const zero = vPos(clamp(0, scale.min, scale.max));
    // Gridlines and axis lines
    const grid = s("g", { class: "hs-ochart-grid" });
    if (va.grid && va.grid !== "none") {
      for (const v of ticks) {
        const p = vPos(v);
        grid.append(horizontal ? s("line", { x1: r2(p), x2: r2(p), y1: r2(plot.y), y2: r2(plot.y + plot.h), stroke: va.grid, "stroke-width": 1.5 }) : s("line", { x1: r2(plot.x), x2: r2(plot.x + plot.w), y1: r2(p), y2: r2(p), stroke: va.grid, "stroke-width": 1.5 }));
      }
    }
    svg.append(grid);
    if (!va.hide) {
      const g = s("g", { class: "hs-ochart-vaxis" });
      const valAtTop = horizontal && ca.reverse;
      for (const v of ticks) {
        const p = vPos(v);
        if (horizontal) g.append(textEl(p, valAtTop ? plot.y - valFont.size * 0.5 : plot.y + plot.h + valFont.size * 1.15, tickText(v), valFont, "middle"));
        else g.append(textEl(plot.x - 8, p + valFont.size * 0.35, tickText(v), valFont, "end"));
      }
      if (va.line && va.line !== "none") g.append(horizontal ? s("line", { x1: r2(plot.x), x2: r2(plot.x + plot.w), y1: r2(valAtTop ? plot.y : plot.y + plot.h), y2: r2(valAtTop ? plot.y : plot.y + plot.h), stroke: va.line, "stroke-width": 1.5 }) : s("line", { x1: r2(plot.x), x2: r2(plot.x), y1: r2(plot.y), y2: r2(plot.y + plot.h), stroke: va.line, "stroke-width": 1.5 }));
      svg.append(g);
    }
    // Bars
    const gap = (sameType ? st.gap : null) ?? 150;
    const overlap = stack ? 100 : (sameType ? st.overlap : null) ?? 0;
    const k = stack ? 1 : Math.max(1, barIdx.length);
    const groupW = slot / (1 + gap / 100);
    const bw = groupW / (k - ((k - 1) * overlap) / 100);
    const labelsLayer = s("g", { class: "hs-ochart-labels" });
    const labelFor = (si, i, value, where) => {
      const ss = sts(si);
      const own = ss.pointLabels?.find((p) => p.i === i);
      const l = { ...(ss.label || {}), ...(own || {}) };
      if (own && own.show === false) return;
      if (!own && (!ss.label || ss.label.show === false)) return;
      if (!l.runs && !l.val && !l.cat && !l.ser && !l.pct) return;
      const f = fontOf(ss.label?.font, own?.font);
      const parts = l.runs ? null : [l.ser ? c.series[si].name : "", l.cat ? c.labels[i] : "", l.val ? numFormat(value, l.format) : ""].filter(Boolean);
      const text = l.runs ? l.runs.map((r) => r.t).join("") : parts.join(", ");
      const width = textWidth(text, Math.max(f.size, ...(l.runs || []).map((r) => r.font?.size || 0)));
      const { x, y, anchor } = where(l.pos, width, f);
      const el = textEl(x, y, l.runs ? "" : text, f, anchor, { "data-s": String(si) });
      if (l.runs) for (const r of l.runs) { const rf = fontOf(r.font); el.append(s("tspan", { fill: rf.color, style: { "font-size": `${rf.size}px`, "font-weight": rf.bold ? 700 : 400 } }, r.t)); }
      labelsLayer.append(el);
    };
    const bars = s("g", { class: "hs-ochart-bars" });
    const stackBase = c.labels.map(() => ({ pos: 0, neg: 0 }));
    barIdx.forEach((si, bi) => {
      c.labels.forEach((label, i) => {
        const raw = valuesOf(si)[i] || 0;
        const v = shown(si, i);
        let from = 0, to = v;
        if (stack) {
          const b = stackBase[i];
          if (v >= 0) { from = b.pos; to = b.pos + v; b.pos = to; } else { from = b.neg; to = b.neg + v; b.neg = to; }
        }
        const p0 = vPos(from), p1 = vPos(to);
        const off = (slot - groupW) / 2 + (stack ? 0 : bi * bw * (1 - overlap / 100));
        const start = (horizontal ? plot.y : plot.x) + order(i) * slot + off;
        const color = pointColor(si, i);
        const rect = horizontal
          ? { x: Math.min(p0, p1), y: start, width: Math.abs(p1 - p0), height: bw }
          : { x: start, y: Math.min(p0, p1), width: bw, height: Math.abs(p1 - p0) };
        bars.append(s("rect", { class: `hs-obar hs-mark${horizontal ? " h" : ""}`, x: r2(rect.x), y: r2(rect.y), width: r2(Math.max(0, rect.width)), height: r2(Math.max(0, rect.height)), fill: color, "data-paint": color.startsWith("#") ? color : null, "data-s": String(si), "data-neg": to < from ? "1" : null, "data-tip": `${label}・${c.series[si].name}：${numFormat(raw, sts(si).label?.format)}`, style: { "--i": String(i) } }));
        labelFor(si, i, raw, (pos, width, f) => {
          const where = pos || (stack ? "ctr" : "outEnd");
          const up = to >= from;
          const mid = start + bw / 2;
          if (horizontal) {
            const end = vPos(to), begin = vPos(from);
            const dir = up ? 1 : -1;
            const x = where === "inEnd" ? end - dir * 6 : where === "ctr" ? (end + begin) / 2 : where === "inBase" ? begin + dir * 6 : end + dir * 6;
            const anchor = where === "ctr" ? "middle" : (where === "inEnd") === up ? "end" : "start";
            return { x, y: mid + f.size * 0.35, anchor };
          }
          const end = vPos(to), begin = vPos(from);
          const dir = up ? -1 : 1;
          const y = where === "inEnd" ? end - dir * (f.size * 1.0) : where === "ctr" ? (end + begin) / 2 + f.size * 0.35 : where === "inBase" ? begin + dir * 6 + (up ? 0 : f.size) : end + dir * 6 + (up ? 0 : f.size * 0.8);
          return { x: mid, y, anchor: "middle" };
        });
      });
    });
    svg.append(bars);
    // Lines and areas
    const lines = s("g", { class: "hs-ochart-lines" });
    for (const si of lineIdx) {
      const ss = sts(si);
      const color = colorOf(si);
      const pts = c.labels.map((_, i) => [cPos(i), vPos(valuesOf(si)[i] || 0)]).map(([a, b]) => (horizontal ? [b, a] : [a, b]));
      if (!pts.length) continue;
      const d = ss.smooth ? smoothPath(pts) : `M${pts.map((p) => P(p[0], p[1])).join(" L")}`;
      if (kindOf(si) === "area") {
        const zx = horizontal ? zero : null;
        const close = horizontal ? ` L${P(zx, pts.at(-1)[1])} L${P(zx, pts[0][1])} Z` : ` L${P(pts.at(-1)[0], zero)} L${P(pts[0][0], zero)} Z`;
        lines.append(s("path", { class: "hs-oarea", d: d + close, fill: color, "data-paint": color.startsWith("#") ? color : null, "data-s": String(si) }));
      } else {
        const width = ss.width ?? 4.5;
        const dash = DASHES[ss.dash]?.[1];
        lines.append(s("path", { class: "hs-oline", d, fill: "none", stroke: color, "stroke-width": r2(width), "stroke-linejoin": "round", "stroke-linecap": "round", "stroke-dasharray": dash ? dash.map((v) => r2(v * width)).join(" ") : null, "data-paint": color.startsWith("#") ? color : null, "data-s": String(si) }));
      }
      const m = ss.marker || (c.type === "combo" || kindOf(si) === "area" ? { s: "none" } : { s: "circle", z: 10 });
      pts.forEach(([x, y], i) => {
        const on = m.s !== "none" || m.at?.includes(i);
        const tip = `${c.labels[i]}・${c.series[si].name}：${numFormat(valuesOf(si)[i] || 0, ss.label?.format)}`;
        if (on) lines.append(marker(m.s === "none" ? "circle" : m.s, x, y, m.z || 10, m.color || color, tip, si));
        // Every point answers the mouse while presenting, marked or not (an invisible target).
        else lines.append(s("circle", { class: "hs-mark hs-ohit", cx: r2(x), cy: r2(y), r: 16, "data-s": String(si), "data-tip": tip }));
        labelFor(si, i, valuesOf(si)[i] || 0, (pos, width, f) => {
          const where = pos || "r";
          const r = (m.z || 10) / 2 + 6;
          if (where === "l") return { x: x - r, y: y + f.size * 0.35, anchor: "end" };
          if (where === "t") return { x, y: y - r, anchor: "middle" };
          if (where === "b") return { x, y: y + r + f.size * 0.8, anchor: "middle" };
          if (where === "ctr") return { x, y: y + f.size * 0.35, anchor: "middle" };
          return { x: x + r, y: y + f.size * 0.35, anchor: "start" };
        });
      });
    }
    svg.append(lines);
    // The category axis
    if (!ca.hide) {
      const g = s("g", { class: "hs-ochart-caxis" });
      const line = ca.line || "#d9d9d9";
      if (line !== "none") g.append(horizontal ? s("line", { x1: r2(zero), x2: r2(zero), y1: r2(plot.y), y2: r2(plot.y + plot.h), stroke: line, "stroke-width": 1.5 }) : s("line", { x1: r2(plot.x), x2: r2(plot.x + plot.w), y1: r2(zero), y2: r2(zero), stroke: line, "stroke-width": 1.5 }));
      c.labels.forEach((label, i) => {
        const p = cPos(i);
        if (horizontal) {
          const rows = wrapText(label, catFont.size, Math.max(40, plot.x - labelLeft - 8), 2);
          rows.forEach((row, ri) => g.append(textEl(plot.x - 8, p + catFont.size * 0.35 + (ri - (rows.length - 1) / 2) * catFont.size * 1.2, row, catFont, "end")));
        } else {
          const rows = wrapText(label, catFont.size, slot * 0.92, 3);
          rows.forEach((row, ri) => g.append(textEl(p, plot.y + plot.h + catFont.size * 1.15 + ri * catFont.size * 1.2, row, catFont, "middle")));
        }
      });
      svg.append(g);
    } else if (ca.line && ca.line !== "none") {
      svg.append(horizontal ? s("line", { x1: r2(zero), x2: r2(zero), y1: r2(plot.y), y2: r2(plot.y + plot.h), stroke: ca.line, "stroke-width": 1.5 }) : s("line", { x1: r2(plot.x), x2: r2(plot.x + plot.w), y1: r2(zero), y2: r2(zero), stroke: ca.line, "stroke-width": 1.5 }));
    }
    svg.append(labelsLayer);
    return svg;
  }

  /** A smooth line through the points (Catmull-Rom as cubic Béziers, like Office's smoothed lines). */
  function smoothPath(pts) {
    if (pts.length < 3) return `M${pts.map((p) => P(p[0], p[1])).join(" L")}`;
    let d = `M${P(pts[0][0], pts[0][1])}`;
    for (let i = 0; i < pts.length - 1; i += 1) {
      const p0 = pts[i - 1] || pts[i], p1 = pts[i], p2 = pts[i + 1], p3 = pts[i + 2] || p2;
      d += ` C${P(p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6)} ${P(p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6)} ${P(p2[0], p2[1])}`;
    }
    return d;
  }

  function marker(kind, x, y, size, color, tip, si = 0) {
    const r = size / 2;
    const attrs = { class: "hs-omark hs-mark", fill: color, "data-tip": tip, "data-s": String(si) };
    if (kind === "square") return s("rect", { ...attrs, x: r2(x - r), y: r2(y - r), width: r2(size), height: r2(size) });
    if (kind === "diamond") return s("path", { ...attrs, d: `M${P(x, y - r)} L${P(x + r, y)} L${P(x, y + r)} L${P(x - r, y)} Z` });
    if (kind === "triangle") return s("path", { ...attrs, d: `M${P(x, y - r)} L${P(x + r, y + r)} L${P(x - r, y + r)} Z` });
    if (kind === "dash") return s("rect", { ...attrs, x: r2(x - r), y: r2(y - size / 8), width: r2(size), height: r2(size / 4) });
    if (kind === "x" || kind === "plus" || kind === "star") {
      const d = kind === "plus" ? `M${P(x - r, y)} L${P(x + r, y)} M${P(x, y - r)} L${P(x, y + r)}` : `M${P(x - r, y - r)} L${P(x + r, y + r)} M${P(x + r, y - r)} L${P(x - r, y + r)}${kind === "star" ? ` M${P(x, y - r)} L${P(x, y + r)}` : ""}`;
      return s("path", { class: "hs-omark hs-mark", d, stroke: color, "stroke-width": Math.max(1.5, size / 6), fill: "none", "data-tip": tip, "data-s": String(si) });
    }
    return s("circle", { ...attrs, cx: r2(x), cy: r2(y), r: r2(kind === "dot" ? r / 2 : r) });
  }

  function pieChart(svg, c, area, { st, sts, pointColor, fontOf, textEl }) {
    const values = c.series[0].values.map((v) => Math.max(0, v || 0));
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const cx = area.x + area.w / 2;
    const cy = area.y + area.h / 2;
    const R = Math.max(4, Math.min(area.w, area.h) / 2 - 4);
    const hole = (st.type && st.type !== c.type ? 50 : st.hole ?? 50) / 100;
    let a = ((st.angle || 0) - 90) * (Math.PI / 180);
    const g = s("g", { class: "hs-ochart-pie", "data-c": `${r2(cx)},${r2(cy)}` });
    const labels = s("g", { class: "hs-ochart-labels" });
    const ss = sts(0);
    values.forEach((v, i) => {
      const sweep = (v / total) * Math.PI * 2;
      const a1 = a + sweep;
      const color = pointColor(0, i);
      const large = sweep > Math.PI ? 1 : 0;
      const pt = (r, ang) => P(cx + r * Math.cos(ang), cy + r * Math.sin(ang));
      let d;
      if (sweep >= Math.PI * 2 - 1e-6) d = hole ? `M${pt(R, 0)} A${r2(R)} ${r2(R)} 0 1 1 ${pt(R, Math.PI)} A${r2(R)} ${r2(R)} 0 1 1 ${pt(R, 0)} M${pt(R * hole, 0)} A${r2(R * hole)} ${r2(R * hole)} 0 1 0 ${pt(R * hole, Math.PI)} A${r2(R * hole)} ${r2(R * hole)} 0 1 0 ${pt(R * hole, 0)} Z` : `M${pt(R, 0)} A${r2(R)} ${r2(R)} 0 1 1 ${pt(R, Math.PI)} A${r2(R)} ${r2(R)} 0 1 1 ${pt(R, 0)} Z`;
      else if (hole) d = `M${pt(R, a)} A${r2(R)} ${r2(R)} 0 ${large} 1 ${pt(R, a1)} L${pt(R * hole, a1)} A${r2(R * hole)} ${r2(R * hole)} 0 ${large} 0 ${pt(R * hole, a)} Z`;
      else d = `M${P(cx, cy)} L${pt(R, a)} A${r2(R)} ${r2(R)} 0 ${large} 1 ${pt(R, a1)} Z`;
      if (v > 0) g.append(s("path", { class: "hs-oslice hs-mark", d, fill: color, "fill-rule": "evenodd", stroke: "#ffffff", "stroke-width": 2, "data-paint": color.startsWith("#") ? color : null, "data-tip": `${c.labels[i]}：${numFormat(c.series[0].values[i], ss.label?.format)}（${Math.round((v / total) * 100)}%）`, style: { "--i": String(i) } }));
      const own = ss.pointLabels?.find((p) => p.i === i);
      const l = { ...(ss.label || {}), ...(own || {}) };
      const show = own ? own.show !== false : ss.label && ss.label.show !== false;
      if (show && v > 0 && (l.runs || l.val || l.pct || l.cat || l.ser)) {
        const f = fontOf(ss.label?.font, own?.font);
        const mid = a + sweep / 2;
        const outside = l.pos === "outEnd";
        const r = outside ? R + f.size * 0.9 : hole ? R * (1 + hole) / 2 : l.pos === "inEnd" ? R * 0.78 : R * 0.62;
        const text = l.runs ? l.runs.map((x) => x.t).join("") : [l.ser ? c.series[0].name : "", l.cat ? c.labels[i] : "", l.val ? numFormat(c.series[0].values[i], l.format) : "", l.pct ? `${Math.round((v / total) * 100)}%` : ""].filter(Boolean).join(l.cat && (l.val || l.pct) ? "\n" : ", ");
        const x = cx + r * Math.cos(mid);
        const y = cy + r * Math.sin(mid);
        const anchor = outside ? (Math.cos(mid) > 0.2 ? "start" : Math.cos(mid) < -0.2 ? "end" : "middle") : "middle";
        const rows = text.split("\n");
        rows.forEach((row, ri) => labels.append(textEl(x, y + f.size * 0.35 + (ri - (rows.length - 1) / 2) * f.size * 1.2, row, f, anchor)));
      }
      a = a1;
    });
    svg.append(g, labels);
  }
  /** The engine's chart spec (as the layouts' charts) for a chart object. */
  /** The chart as drawn: without the series and categories グラフ フィルター hides (a waterfall's totals follow). */
  function visibleChart(c) {
    const hs = new Set(c.opts?.hideSeries || []);
    const hl = new Set(c.opts?.hideLabels || []);
    const flip = c.opts?.reverse === true && REVERSE_CHARTS.has(c.type);
    if (!hs.size && !hl.size) return flip ? { ...c, labels: [...c.labels].reverse(), series: c.series.map((s) => ({ ...s, values: [...s.values].reverse() })) } : c;
    const keep = c.labels.map((_, i) => !hl.has(i));
    const at = [];
    let n = 0;
    keep.forEach((k, i) => { if (k) at[i] = n++; });
    const { hideSeries: _s, hideLabels: _l, ...opts } = c.opts;
    if (opts.totals) { opts.totals = opts.totals.filter((i) => keep[i]).map((i) => at[i]); if (!opts.totals.length) delete opts.totals; }
    const out = { ...c, labels: c.labels.filter((_, i) => keep[i]), series: c.series.filter((_, i) => !hs.has(i)).map((s) => ({ ...s, values: s.values.filter((_, i) => keep[i]) })) };
    if (flip) { out.labels.reverse(); for (const s of out.series) s.values.reverse(); }
    if (Object.keys(opts).length) out.opts = opts; else delete out.opts;
    return out;
  }
  // データ テーブル: the charts with categories along an axis.
  const TABLE_CHARTS = new Set(["bar", "clustered-bar", "stacked-bar", "100-stacked-bar", "line", "multi-line", "area", "stacked-area", "combo", "waterfall"]);
  function dataTable(c) {
    const shown = (v) => (v == null || v === "" ? "" : c.opts?.numFmt ? E.formatNumber(v, c.opts.numFmt, true) : Number(v).toLocaleString("ja-JP"));
    return h("table", { class: "hs-chart-table" },
      h("thead", {}, h("tr", {}, h("th", {}), c.labels.map((label) => h("th", {}, label)))),
      h("tbody", {}, c.series.map((series) => h("tr", {}, h("th", {}, series.name), series.values.map((v) => h("td", {}, shown(v)))))));
  }

  function chartSpec(chart) {
    const c = visibleChart(chart);
    const data = { ...(c.title ? { title: c.title } : {}), ...(c.unit ? { unit: c.unit } : {}) };
    const first = c.series[0]?.values || [];
    if (c.type === "stacked-bar" || c.type === "100-stacked-bar") {
      data.barData = c.labels.map((label, i) => ({ label, values: c.series.map((s) => s.values[i] ?? 0) }));
      data.legendLabels = c.series.map((s) => s.name);
    } else if (["multi-line", "clustered-bar", "area", "scatter", "radar", "hbar", "stacked-hbar", "stacked-area", "bubble", "histogram", "boxplot", "treemap", "sunburst"].includes(c.type)) {
      data.xAxisLabels = c.labels;
      data.series = c.series.map((s) => ({ label: s.name, values: s.values }));
    } else if (c.type === "pareto") {
      // パレート図: the first series from the largest, a line of the running share (%) on a scale of its own that ends at 100.
      const order = c.labels.map((_, i) => i).sort((a, b) => (first[b] ?? 0) - (first[a] ?? 0) || a - b);
      const total = order.reduce((sum, i) => sum + Math.max(0, first[i] ?? 0), 0) || 1;
      let run = 0;
      data.items = order.map((i) => { run += Math.max(0, first[i] ?? 0); return { label: c.labels[i], barValue: first[i] ?? 0, value: Math.round((run / total) * 1000) / 10 }; });
      data.legendLabels = [c.series[0]?.name, "累積比率"].filter(Boolean);
    } else if (c.type === "combo") {
      const line = c.series[1]?.values || first;
      data.items = c.labels.map((label, i) => ({ label, barValue: first[i] ?? 0, value: line[i] ?? 0 }));
      data.legendLabels = [c.series[0]?.name, c.series[1]?.name || c.series[0]?.name].filter(Boolean);
    } else data.items = c.labels.map((label, i) => ({ label, value: first[i] ?? 0 }));
    if (c.type === "pareto") return { chartType: "combo", data, opts: { ...(c.opts || {}), axis2: true, axis2Pct: true } };
    return { chartType: c.type, data, ...(c.opts ? { opts: c.opts } : {}) };
  }

  /** A slide's objects: the usable ones, ids made unique. */
  function normalizeObjects(list) {
    const seen = new Set();
    const out = [];
    for (const raw of Array.isArray(list) ? list : []) {
      const o = normalizeObject(raw);
      if (!o) continue;
      if (seen.has(o.id)) o.id = newId();
      seen.add(o.id);
      out.push(o);
      if (out.length >= 5000) break;
    }
    // Connectors only stay attached to objects that exist; reveal buttons only show objects that exist.
    for (const o of out) for (const end of ["from", "to"]) if (o[end] && !seen.has(o[end].id)) delete o[end];
    for (const o of out) {
      if (o.action?.type !== "reveal") continue;
      o.action.targets = o.action.targets.filter((id) => seen.has(id));
      if (!o.action.targets.length) delete o.action;
    }
    return out;
  }

  // ---------------------------------------------------------------- where an object is

  /** Corners of an object's box after rotation (slide pixels). */
  function corners(o) {
    if (o.kind === "line") { const [p1, p2] = lineEnds(o); return [p1, p2]; }
    const cx = o.x + o.w / 2;
    const cy = o.y + o.h / 2;
    const a = rad(o.rot || 0);
    const [c, sn] = [Math.cos(a), Math.sin(a)];
    return [[-1, -1], [1, -1], [1, 1], [-1, 1]].map(([kx, ky]) => {
      const dx = (kx * o.w) / 2;
      const dy = (ky * o.h) / 2;
      return [cx + dx * c - dy * sn, cy + dx * sn + dy * c];
    });
  }
  /** The upright box around an object (what guides and alignment use). */
  function bounds(o, all = null) {
    if (o.kind === "line") {
      const [[x1, y1], [x2, y2]] = lineEnds(o, all);
      return { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) };
    }
    const pts = corners(o);
    const xs = pts.map((p) => p[0]);
    const ys = pts.map((p) => p[1]);
    return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
  }
  /** Where a connector may attach on an object: the middles of its sides (top, right, bottom, left), rotated. */
  function sites(o) {
    const cx = o.x + o.w / 2;
    const cy = o.y + o.h / 2;
    const a = rad(o.rot || 0);
    const [c, sn] = [Math.cos(a), Math.sin(a)];
    return [[0, -1], [1, 0], [0, 1], [-1, 0]].map(([kx, ky]) => {
      const dx = (kx * o.w) / 2;
      const dy = (ky * o.h) / 2;
      return [cx + dx * c - dy * sn, cy + dx * sn + dy * c];
    });
  }
  /** A line's two ends, following the objects it is attached to. */
  function lineEnds(o, all = null) {
    let p1 = [o.x1, o.y1];
    let p2 = [o.x2, o.y2];
    if (all) {
      const at = (ref) => { const target = ref && all.find((other) => other.id === ref.id && other.kind !== "line"); return target ? sites(target)[ref.site % 4] : null; };
      p1 = at(o.from) || p1;
      p2 = at(o.to) || p2;
    }
    return [p1, p2];
  }
  /** The side a connector leaves an attached object from (unit vector), for elbows and curves. */
  function siteDir(ref, all) {
    const target = ref && all?.find((other) => other.id === ref.id);
    if (!target) return null;
    const base = [[0, -1], [1, 0], [0, 1], [-1, 0]][ref.site % 4];
    const a = rad(target.rot || 0);
    return [base[0] * Math.cos(a) - base[1] * Math.sin(a), base[0] * Math.sin(a) + base[1] * Math.cos(a)];
  }

  /** The path a line takes between its ends, and the direction it arrives at each end. */
  function linePath(o, all) {
    const [p1, p2] = lineEnds(o, all);
    const route = o.route || "straight";
    if (route === "elbow") {
      const d1 = siteDir(o.from, all);
      const d2 = siteDir(o.to, all);
      const horizontal = d1 ? Math.abs(d1[0]) > 0.5 : Math.abs(p2[0] - p1[0]) >= Math.abs(p2[1] - p1[1]);
      const pts = horizontal
        ? [p1, [(p1[0] + p2[0]) / 2, p1[1]], [(p1[0] + p2[0]) / 2, p2[1]], p2]
        : [p1, [p1[0], (p1[1] + p2[1]) / 2], [p2[0], (p1[1] + p2[1]) / 2], p2];
      if (d2 && horizontal !== Math.abs(d2[0]) > 0.5) pts.splice(2, 1, horizontal ? [p2[0], p1[1]] : [p1[0], p2[1]]);
      const clean = pts.filter((p, i) => i === 0 || Math.hypot(p[0] - pts[i - 1][0], p[1] - pts[i - 1][1]) > 0.01);
      return { pts: clean, d: `M${clean.map((p) => P(...p)).join(" L")}`, start: clean[1] || p2, end: clean[clean.length - 2] || p1 };
    }
    if (route === "curve") {
      const d1 = siteDir(o.from, all);
      const d2 = siteDir(o.to, all);
      const len = Math.hypot(p2[0] - p1[0], p2[1] - p1[1]) * 0.45;
      const horizontal = Math.abs(p2[0] - p1[0]) >= Math.abs(p2[1] - p1[1]);
      const c1 = d1 ? [p1[0] + d1[0] * len, p1[1] + d1[1] * len] : horizontal ? [(p1[0] + p2[0]) / 2, p1[1]] : [p1[0], (p1[1] + p2[1]) / 2];
      const c2 = d2 ? [p2[0] + d2[0] * len, p2[1] + d2[1] * len] : horizontal ? [(p1[0] + p2[0]) / 2, p2[1]] : [p2[0], (p1[1] + p2[1]) / 2];
      return { pts: [p1, p2], d: `M${P(...p1)} C${P(...c1)} ${P(...c2)} ${P(...p2)}`, start: c1, end: c2, ctrl: [c1, c2] };
    }
    return { pts: [p1, p2], d: `M${P(...p1)} L${P(...p2)}`, start: p2, end: p1 };
  }

  /** An arrowhead at `tip`, arriving from `from`. Returns { d, fill, back } (back: how far to pull the line in). */
  function arrowhead(kind, tip, from, sw, size) {
    const len = Math.max(10, sw * (size === 1 ? 2.2 : size === 3 ? 4.4 : 3.2));
    const half = len * 0.55;
    const ang = Math.atan2(tip[1] - from[1], tip[0] - from[0]);
    const [c, sn] = [Math.cos(ang), Math.sin(ang)];
    const at = (along, across) => [tip[0] - along * c - across * sn, tip[1] - along * sn + across * c];
    if (kind === "triangle") return { d: poly([tip, at(len, half), at(len, -half)]), fill: true, back: len * 0.9 };
    if (kind === "stealth") return { d: poly([tip, at(len, half), at(len * 0.65, 0), at(len, -half)]), fill: true, back: len * 0.6 };
    if (kind === "arrow") return { d: `M${P(...at(len, half))} L${P(...tip)} L${P(...at(len, -half))}`, fill: false, back: 0 };
    if (kind === "diamond") return { d: poly([tip, at(len / 2, half * 0.8), at(len, 0), at(len / 2, -half * 0.8)]), fill: true, back: len / 2 };
    if (kind === "oval") { const cx = at(len / 2, 0); return { d: ellipse(cx[0], cx[1], len / 2, len / 2), fill: true, back: len / 2 }; }
    return null;
  }

  // ---------------------------------------------------------------- rendering

  /** 複合線 (二重線・三重線): a mask that cuts see-through bands down the middle of a stroke `sw` wide (two bands
   *  of a third for a double line; two fifths and a fifth between them for a triple one). Returns its `url(#…)`. */
  function compoundMask(svg, cmpd, sw, paths, cap, join) {
    const id = `cm${newId().slice(1)}`;
    const mask = s("mask", { id, maskUnits: "userSpaceOnUse", x: -4000, y: -4000, width: 12000, height: 12000 });
    mask.append(s("rect", { x: -4000, y: -4000, width: 12000, height: 12000, fill: "#ffffff" }));
    const bands = cmpd === "tri" ? [[(sw * 3) / 5, "#000000"], [sw / 5, "#ffffff"]] : [[sw / 3, "#000000"]];
    for (const [width, color] of bands) for (const d of paths) mask.append(s("path", { d, fill: "none", stroke: color, "stroke-width": r2(width), "stroke-linecap": cap || "butt", "stroke-linejoin": join || "miter", "stroke-miterlimit": 8 }));
    svg.append(s("defs", {}, mask));
    return `url(#${id})`;
  }
  const dashArray = (dash, sw) => {
    const pattern = DASHES[dash]?.[1];
    return pattern ? pattern.map((v) => r2(Math.max(v * sw, v === 0 ? 0.01 : 1))).join(" ") : null;
  };

  /** An object's rich text, with every point size following the box's shrink factor (--os). */
  function textNode(o, scale) {
    const tx = h("div", { class: "hs-obj-tx" });
    const frag = richFragment(o.text || "");
    tx.append(frag);
    if (!o.vertical && !o.textRot && (o.text || "").includes("\t")) layoutTabs(tx);
    for (const el of tx.querySelectorAll("[style*='font-size']")) {
      const size = pxSize(el.style.fontSize);
      if (size) el.style.fontSize = `calc(${size}px * var(--os, 1))`;
    }
    const style = {
      "font-size": `calc(${o.fs}px * var(--os, 1))`, color: o.color, "line-height": String(o.lh), "text-align": alignCss(o.align), "text-align-last": alignLastCss(o.align),
      "font-weight": o.bold ? "700" : null, "font-style": o.italic ? "italic" : null,
      "text-decoration": [o.underline ? "underline" : "", o.strike ? "line-through" : ""].filter(Boolean).join(" ") || null,
      "text-decoration-style": (o.underline && o.uline && o.uline !== "thick" ? o.uline : null) || (o.strike && o.sline === "double" && !o.underline ? "double" : null),
      "text-decoration-thickness": o.underline && o.uline === "thick" ? ".12em" : null,
      "letter-spacing": o.ls ? `${o.ls}em` : null, "--psp": o.psp ? `${o.psp}em` : null,
      "text-transform": o.caps === "all" ? "uppercase" : null, "font-variant": o.caps === "small" ? "small-caps" : null,
      "text-shadow": textShadowOf(o),
      "-webkit-text-stroke": o.toutline ? `${o.toutline.w}px ${o.toutline.color}` : null, "paint-order": o.toutline ? "stroke fill" : null,
      "font-family": o.fontFace ? `"${o.fontFace}", sans-serif` : o.font ? FONTS[o.font][1] : null, "white-space": o.wrap === false ? "pre" : null,
    };
    for (const [k, v] of Object.entries(style)) if (v != null) tx.style.setProperty(k, String(v));
    if (scale && scale < 1) tx.style.setProperty("--os", String(scale));
    return tx;
  }

  const rgbaOf = (hex, a) => `rgba(${[1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16)).join(",")},${a})`;
  /** 文字の効果: the letters' 影 and 光彩 as a CSS text-shadow (the halo is drawn twice, wide then close, so it reads on a light fill). */
  function textShadowOf(o) {
    const layers = [];
    if (o.tshadow) layers.push(`${o.tshadow.dx}px ${o.tshadow.dy}px ${o.tshadow.blur}px ${rgbaOf(o.tshadow.color, o.tshadow.opacity)}`);
    if (o.tglow) layers.push(`0 0 ${o.tglow.r}px ${rgbaOf(o.tglow.color, o.tglow.opacity)}`, `0 0 ${r2(o.tglow.r / 3)}px ${rgbaOf(o.tglow.color, o.tglow.opacity)}`);
    return layers.join(", ") || null;
  }
  /** 図形の効果: 影 and 光彩 are drop-shadow filters, 反射 a mirror below (the browser's box-reflect), on the object's own element. */
  function applyEffects(o, rotEl) {
    const filters = [];
    if (o.shadow) filters.push(`drop-shadow(${o.shadow.dx}px ${o.shadow.dy}px ${o.shadow.blur}px ${rgbaOf(o.shadow.color, o.shadow.opacity)})`);
    // The halo is drawn twice (wide, then close) so it reads on a light fill.
    if (o.glow) filters.push(`drop-shadow(0 0 ${o.glow.r}px ${rgbaOf(o.glow.color, o.glow.opacity)})`, `drop-shadow(0 0 ${r2(o.glow.r / 3)}px ${rgbaOf(o.glow.color, o.glow.opacity)})`);
    if (filters.length) rotEl.style.filter = filters.join(" ");
    if (o.reflect) rotEl.style.setProperty("-webkit-box-reflect", `below ${o.reflect.gap}px linear-gradient(to bottom, transparent ${Math.round((1 - o.reflect.size) * 100)}%, rgba(0,0,0,${o.reflect.opacity}))`);
  }
  /** ぼかし: the edges of the shape (or picture) fade out — a mask fading on all four sides, `r` px wide. */
  function softEdges(el, r) {
    const fade = (dir) => `linear-gradient(to ${dir}, transparent, #000 ${r}px, #000 calc(100% - ${r}px), transparent)`;
    const mask = `${fade("right")}, ${fade("bottom")}`;
    for (const [key, value] of [["-webkit-mask-image", mask], ["mask-image", mask], ["-webkit-mask-composite", "source-in"], ["mask-composite", "intersect"]]) el.style.setProperty(key, value);
  }

  // ---------------------------------------------------------------- パターンの塗りつぶし
  // PowerPoint's pattern fills are tiny 8 × 8 bitmaps repeated over the shape. These are drawn the same way, a cell being 1.5 px on
  // the slide, in two colours; a few of PowerPoint's own are drawn by the nearest of these.
  const BAYER = [[0, 32, 8, 40, 2, 34, 10, 42], [48, 16, 56, 24, 50, 18, 58, 26], [12, 44, 4, 36, 14, 46, 6, 38], [60, 28, 52, 20, 62, 30, 54, 22], [3, 35, 11, 43, 1, 33, 9, 41], [51, 19, 59, 27, 49, 17, 57, 25], [15, 47, 7, 39, 13, 45, 5, 37], [63, 31, 55, 23, 61, 29, 53, 21]];
  const m8 = (n) => ((n % 8) + 8) % 8;
  const m4 = (n) => ((n % 4) + 4) % 4;
  const PATTERN_KINDS = {};
  for (const p of [5, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90]) PATTERN_KINDS[`pct${p}`] = [`${p}%`, (x, y) => BAYER[y][x] < Math.round(p * 0.64)];
  Object.assign(PATTERN_KINDS, {
    ltHorz: ["薄い横線", (x, y) => y === 0], horz: ["横線", (x, y) => y % 4 === 0], narHorz: ["細かい横線", (x, y) => y % 2 === 0], dkHorz: ["濃い横線", (x, y) => y % 4 < 2],
    ltVert: ["薄い縦線", (x) => x === 0], vert: ["縦線", (x) => x % 4 === 0], narVert: ["細かい縦線", (x) => x % 2 === 0], dkVert: ["濃い縦線", (x) => x % 4 < 2],
    ltDnDiag: ["薄い右下がり斜線", (x, y) => m8(x - y) === 0], dnDiag: ["右下がり斜線", (x, y) => m4(x - y) === 0], dkDnDiag: ["濃い右下がり斜線", (x, y) => m4(x - y) < 2], wdDnDiag: ["太い右下がり斜線", (x, y) => m8(x - y) < 3],
    ltUpDiag: ["薄い右上がり斜線", (x, y) => m8(x + y) === 0], upDiag: ["右上がり斜線", (x, y) => m4(x + y) === 0], dkUpDiag: ["濃い右上がり斜線", (x, y) => m4(x + y) < 2], wdUpDiag: ["太い右上がり斜線", (x, y) => m8(x + y) < 3],
    smGrid: ["小さい格子", (x, y) => x % 4 === 0 || y % 4 === 0], lgGrid: ["大きい格子", (x, y) => x === 0 || y === 0],
    smCheck: ["小さい市松模様", (x, y) => ((x >> 1) + (y >> 1)) % 2 === 0], lgCheck: ["大きい市松模様", (x, y) => ((x >> 2) + (y >> 2)) % 2 === 0],
    diagCross: ["斜め格子", (x, y) => m4(x - y) === 0 || m4(x + y) === 0], openDmnd: ["ひし形格子", (x, y) => m8(x - y) === 0 || m8(x + y) === 0],
  });
  /** An SVG <pattern> of a pattern fill: the background colour and the kind's cells in the other (an unknown kind is 50%). */
  function patternNode(id, { kind, fg, bg }, cell = 1.5) {
    const rule = (PATTERN_KINDS[kind] || PATTERN_KINDS.pct50)[1];
    const size = 8 * cell;
    let d = "";
    for (let y = 0; y < 8; y += 1) {
      for (let x = 0; x < 8; x += 1) {
        if (!rule(x, y)) continue;
        let run = 1;
        while (x + run < 8 && rule(x + run, y)) run += 1;
        d += `M${r2(x * cell)} ${r2(y * cell)}h${r2(run * cell)}v${r2(cell)}h${r2(-run * cell)}z`;
        x += run - 1;
      }
    }
    return s("pattern", { id, patternUnits: "userSpaceOnUse", width: size, height: size }, s("rect", { width: size, height: size, fill: bg }), s("path", { d, fill: fg, "shape-rendering": "crispEdges" }));
  }

  function shapeBody(o, rotEl, scale, ctx = {}) {
    const g = geometry(o.shape, o.w, o.h, o.adj, o.path);
    const sw = o.stroke !== "none" ? o.strokeW : 0;
    const fill = g.open ? "none" : o.fill;
    const svg = s("svg", { class: "hs-obj-geom", width: r2(o.w), height: r2(o.h), viewBox: `0 0 ${r2(Math.max(1, o.w))} ${r2(Math.max(1, o.h))}`, overflow: "visible", "aria-hidden": "true" });
    applyEffects(o, rotEl);
    if (o.soft) softEdges(svg, o.soft);
    let paint = fill;
    if (!g.open && o.gradient?.stops?.length) {
      const id = newId();
      const angle = rad(o.gradient.angle || 0);
      const dx = Math.cos(angle) * 50, dy = Math.sin(angle) * 50;
      const grad = s("linearGradient", { id, x1: `${50 - dx}%`, y1: `${50 - dy}%`, x2: `${50 + dx}%`, y2: `${50 + dy}%` });
      for (const stop of o.gradient.stops) grad.append(s("stop", { offset: `${stop.at * 100}%`, "stop-color": stop.color, "stop-opacity": stop.opacity }));
      svg.append(s("defs", {}, grad));
      paint = `url(#${id})`;
    }
    // パターンの塗りつぶし: two colours in a repeating cell pattern (over a gradient).
    if (!g.open && o.pattern) {
      const id = newId();
      svg.append(s("defs", {}, patternNode(id, o.pattern)));
      paint = `url(#${id})`;
    }
    // 図で塗りつぶし: the picture covers the shape's box (cut to its outline), or repeats as tiles.
    const picture = !g.open && o.fillImg ? E.resolveSrc(o.fillImg, ctx) : "";
    if (picture) {
      const id = newId();
      const [pw, ph] = o.fillTile ? [160, 160] : [r2(Math.max(1, o.w)), r2(Math.max(1, o.h))];
      svg.append(s("defs", {}, s("pattern", { id, patternUnits: "userSpaceOnUse", x: 0, y: 0, width: pw, height: ph },
        s("image", { href: picture, x: 0, y: 0, width: pw, height: ph, preserveAspectRatio: "xMidYMid slice" }))));
      paint = `url(#${id})`;
    }
    const cap = o.cap ? SVG_CAPS[o.cap] : o.dash === "roundDot" ? "round" : null;
    const join = o.join || "miter";
    // 複合線: the fill is drawn alone and the outline over it, cut by a mask.
    const compound = sw && o.cmpd ? compoundMask(svg, o.cmpd, sw, g.paths, cap, join) : null;
    for (const d of g.paths) {
      svg.append(s("path", {
        d, fill: paint === "none" ? "none" : paint, "fill-opacity": fill !== "none" && o.fillOpacity != null ? o.fillOpacity : null, "fill-rule": g.rule,
        stroke: sw && !compound ? o.stroke : "none", "stroke-width": sw && !compound ? sw : null, "stroke-dasharray": sw && !compound ? dashArray(o.dash, sw) : null,
        "stroke-linecap": cap, "stroke-linejoin": join, "stroke-miterlimit": 8,
      }));
      if (compound) svg.append(s("path", { d, fill: "none", stroke: o.stroke, "stroke-width": sw, "stroke-dasharray": dashArray(o.dash, sw), "stroke-linecap": cap, "stroke-linejoin": join, "stroke-miterlimit": 8, mask: compound }));
    }
    for (const extra of g.extras) {
      if (extra.tone === "line") svg.append(s("path", { d: extra.d, fill: "none", stroke: sw ? o.stroke : "rgba(0,0,0,.45)", "stroke-width": sw || 1.5 }));
      else svg.append(s("path", { d: extra.d, fill: fill === "none" ? "none" : extra.tone === "dark" ? "#000000" : "#ffffff", "fill-opacity": extra.tone === "dark" ? 0.12 : 0.35, stroke: sw ? o.stroke : "none", "stroke-width": sw || null }));
    }
    rotEl.append(svg);
    {
      // Every shape has a text box (empty until something is typed in it).
      const [l, t, r, b] = g.text;
      const [pt, pr, pb, pl] = o.pad;
      const box = h("div", { class: ["hs-obj-text", o.vertical ? "is-vertical" : "", o.textRot ? `is-rot${o.textRot}` : ""], "data-valign": o.valign, style: { left: `${r2(l + pl)}px`, top: `${r2(t + pt)}px`, width: `${r2(Math.max(0, r - l - pl - pr))}px`, height: `${r2(Math.max(0, b - t - pt - pb))}px` } });
      const tx = textNode(o, scale);
      if (o.cols > 1) { tx.style.columnCount = String(o.cols); tx.style.columnGap = `${o.colGap ?? 28}px`; tx.style.width = "100%"; }
      box.append(tx);
      rotEl.append(box);
    }
  }

  /** The colours a table style is drawn in (the theme's accent, its 80 % and 60 % lighter tints, 背景 1, テキスト 1). */
  function tableStyleVars(style, themeId = PALETTE.sheet?.themeId) {
    const c = themeColors(themeId);
    const a = style === "medium2a2" ? c.accent2 : c.accent1;
    return { "--tb-a": a, "--tb-t20": lumColor(a, 20000, 80000), "--tb-t40": lumColor(a, 40000, 60000), "--tb-lt": c.bg1, "--tb-dk": c.tx1, "--tb-mid": lumColor(c.tx1, 50000, 50000) };
  }
  /** A table: a PowerPoint table style in the theme's colours, each cell's rich text; merged cells span; rows grow with their text. */
  function tableBody(o, rotEl) {
    const classes = ["hs-otable", `ts-${o.style || "medium2"}`, o.header !== false ? "has-header" : "", o.banded !== false ? "is-banded" : "", o.firstCol ? "has-first-col" : "", o.lastRow ? "has-last-row" : ""].filter(Boolean);
    const table = h("table", { class: classes, style: { "font-size": `calc(${o.fs}px * var(--os, 1))`, color: o.color, "line-height": String(o.lh), "font-family": o.font ? FONTS[o.font][1] : null, ...tableStyleVars(o.style || "medium2") } });
    table.append(h("colgroup", {}, o.cols.map((f) => h("col", { style: { width: `${(f * 100).toFixed(3)}%` } }))));
    const body = h("tbody");
    // フィルター: the rows it hides are not drawn; the rows left share the table's height as theirs did.
    const hidden = new Set(o.hide || []);
    const shown = o.rows.reduce((sum, share, r) => sum + (hidden.has(r) ? 0 : share), 0) || 1;
    o.cells.forEach((row, r) => {
      if (hidden.has(r)) return;
      const tr = h("tr", { "data-r": String(r), style: { height: `${r2((o.rows[r] / (hidden.size ? shown : 1)) * o.h)}px` } });
      row.forEach((cell, c) => {
        if (cell.merged) return;
        const edge = (side) => (cell[side] === "none" ? "none" : cell[side] ? `${cell[side].w}px solid ${cell[side].c}` : null);
        const td = h("td", { "data-r": String(r), "data-c": String(c), rowspan: cell.rs > 1 ? String(cell.rs) : null, colspan: cell.cs > 1 ? String(cell.cs) : null, "data-fill": cell.fill || null,
          style: { background: cell.fill || null, "vertical-align": cell.valign || null, "border-top": edge("bt"), "border-right": edge("br"), "border-bottom": edge("bb"), "border-left": edge("bl"),
            padding: cell.pad ? cell.pad.map((v) => `${v}px`).join(" ") : null } });
        const tx = h("div", { class: ["hs-cell-tx", cell.vertical ? "is-vertical" : ""], style: { "text-align": alignCss(cell.align), "text-align-last": alignLastCss(cell.align), color: cell.color || null, "font-weight": cell.bold ? "700" : null, "font-style": cell.italic ? "italic" : null,
          "font-size": cell.fs ? `calc(${cell.fs}px * var(--os, 1))` : null, "line-height": cell.lh ?? null,
          "font-family": cell.fontFace ? `"${cell.fontFace}", sans-serif` : null,
          "text-decoration": [cell.underline ? "underline" : "", cell.strike ? "line-through" : ""].filter(Boolean).join(" ") || null } });
        tx.append(richFragment(cell.text || ""));
        td.append(tx);
        tr.append(td);
      });
      body.append(tr);
    });
    table.append(body);
    rotEl.append(h("div", { class: "hs-obj-tablebox" }, table));
  }

  /**
   * グラフのデザイン → 色の変更 (PowerPoint's colour sets): カラフル starts the series at another accent; モノクロ uses one
   * accent in darker and lighter tones. The theme's own order (accent 1 to 6) is the default and is not stored.
   * chartColorVars() makes the CSS variables (--c1…--c6 for series, --k1…--k6 for stacked parts) in a theme.
   */
  const CHART_COLORS = {
    colorful2: { label: "カラフル（アクセント 2 から）", start: 2 },
    colorful3: { label: "カラフル（アクセント 3 から）", start: 3 },
    colorful4: { label: "カラフル（アクセント 4 から）", start: 4 },
    mono1: { label: "モノクロ（アクセント 1）", mono: 1 },
    mono2: { label: "モノクロ（アクセント 2）", mono: 2 },
    mono3: { label: "モノクロ（アクセント 3）", mono: 3 },
    mono4: { label: "モノクロ（アクセント 4）", mono: 4 },
    mono5: { label: "モノクロ（アクセント 5）", mono: 5 },
    mono6: { label: "モノクロ（アクセント 6）", mono: 6 },
    gray: { label: "モノクロ（グレー）", mono: "tx1" },
  };
  /** The colours of a chart's colour set in a theme ({} for the theme's own). */
  function chartColorVars(key, themeId = PALETTE.sheet?.themeId) {
    const def = CHART_COLORS[key];
    if (!def) return {};
    const c = themeColors(themeId);
    let series;
    if (def.mono) {
      // One colour and its tones from the colour sheet (so each is a theme colour too), lighter and darker in turn.
      const base = c[def.mono === "tx1" ? "tx1" : `accent${def.mono}`].toLowerCase();
      const tones = lumVariants(base).map(([mod, off]) => lumColor(base, mod, off));
      series = [base, ...[3, 2, 1, 4, 0].map((i) => tones[i])];
    } else {
      series = [0, 1, 2, 3, 4, 5].map((i) => c[`accent${((def.start - 1 + i) % 6) + 1}`]);
    }
    const vars = { "--accent": series[0], "--c-muted": lumColor(series[0], 20000, 80000) };
    series.forEach((v, i) => { vars[`--c${i + 1}`] = v; vars[`--k${i + 1}`] = v; });
    return vars;
  }
  /**
   * グラフのデザイン → クイック レイアウト: a set of chart elements at once (data labels, the legend's place, gridlines,
   * axis titles). The title, the unit, the axis bounds, a trendline and the waterfall's totals stay as they are.
   */
  const CHART_LAYOUTS = {
    layout1: { label: "レイアウト 1", desc: "データ ラベル・凡例（上）・目盛線", labels: true, legend: "top", grid: true },
    layout2: { label: "レイアウト 2", desc: "データ ラベル・凡例（右）", labels: true, legend: "right", grid: false },
    layout3: { label: "レイアウト 3", desc: "凡例（下）・目盛線（データ ラベルなし）", labels: false, legend: "bottom", grid: true },
    layout4: { label: "レイアウト 4", desc: "データ ラベルだけ（凡例・目盛線なし）", labels: true, legend: "none", grid: false },
    layout5: { label: "レイアウト 5", desc: "軸ラベル・凡例（上）・目盛線", labels: false, legend: "top", grid: true, axes: true },
    layout6: { label: "レイアウト 6", desc: "軸ラベル・データ ラベル・凡例（下）", labels: true, legend: "bottom", grid: false, axes: true },
    layout7: { label: "レイアウト 7", desc: "要素なし（グラフだけ）", labels: false, legend: "none", grid: false },
  };
  /** The quick layout a chart matches now (or ""). */
  function chartLayoutOf(chart) {
    const o = chart?.opts || {};
    const axes = Boolean(o.axisX || o.axisY);
    return Object.keys(CHART_LAYOUTS).find((k) => {
      const l = CHART_LAYOUTS[k];
      return (o.labels !== false) === l.labels && (o.legend || "top") === l.legend && (o.grid !== false) === l.grid && axes === Boolean(l.axes);
    }) || "";
  }
  /** The chart with a quick layout's elements (axis titles kept, or named after the unit when the layout adds them). */
  function applyChartLayout(chart, key) {
    const l = CHART_LAYOUTS[key];
    if (!chart || !l) return chart;
    const { labels: _l, legend: _g, grid: _d, axisX, axisY, ...rest } = chart.opts || {};
    const opts = { ...rest };
    if (!l.labels) opts.labels = false;
    if (l.legend !== "top") opts.legend = l.legend;
    if (!l.grid) opts.grid = false;
    if (l.axes) { opts.axisX = axisX || "項目"; opts.axisY = axisY || chart.unit || "値"; }
    const out = { ...chart };
    if (Object.keys(opts).length) out.opts = opts; else delete out.opts;
    return out;
  }
  /**
   * グラフのデザイン → グラフ スタイル: a look for the whole chart at once — one of the colour sets with the chart
   * elements (data labels, the legend's place, gridlines) and the bars' gap or a line's curve and markers. The title, the unit,
   * axis titles and bounds, a trendline, the filter and the data table stay as they are; what a chart's kind has no use for
   * (a gap on a line, a curve on a bar) is left out of the choice.
   */
  const CHART_STYLES = {
    style1: { label: "スタイル 1", desc: "テーマの配色・データ ラベル・凡例（上）・目盛線（標準）", colors: "", labels: true, legend: "top", grid: true },
    style2: { label: "スタイル 2", desc: "モノクロ・データ ラベルだけ（すっきり）", colors: "mono1", labels: true, legend: "none", grid: false, gap: 60 },
    style3: { label: "スタイル 3", desc: "モノクロ・太い棒・データ ラベル・目盛線", colors: "mono1", labels: true, legend: "top", grid: true, gap: 20 },
    style4: { label: "スタイル 4", desc: "グレー・細い棒・凡例（下）・目盛線", colors: "gray", labels: false, legend: "bottom", grid: true, gap: 200 },
    style5: { label: "スタイル 5", desc: "グレー・データ ラベル・凡例（右）・目盛線", colors: "gray", labels: true, legend: "right", grid: true, gap: 80 },
    style6: { label: "スタイル 6", desc: "カラフル（アクセント 2 から）・データ ラベルだけ", colors: "colorful2", labels: true, legend: "none", grid: false, gap: 50 },
    style7: { label: "スタイル 7", desc: "カラフル（アクセント 2 から）・凡例（下）・目盛線", colors: "colorful2", labels: false, legend: "bottom", grid: true, gap: 120 },
    style8: { label: "スタイル 8", desc: "モノクロ・なめらかな線・ひし形のマーカー・凡例（下）", colors: "mono1", labels: false, legend: "bottom", grid: true, gap: 100, smooth: true, marker: "diamond" },
  };
  /** The chart style a chart matches now (or ""): only what its kind can show is compared. */
  function chartStyleOf(chart) {
    if (!chart || chart.style) return "";
    const o = chart.opts || {};
    const bars = BAR_GAP_CHARTS.has(chart.type);
    const lines = LINE_CHART_KINDS.has(chart.type);
    return Object.keys(CHART_STYLES).find((k) => {
      const s = CHART_STYLES[k];
      if ((chart.colors || "") !== s.colors || (o.labels !== false) !== s.labels || (o.legend || "top") !== s.legend || (o.grid !== false) !== s.grid) return false;
      if (bars && (o.gap ?? null) !== (s.gap ?? null)) return false;
      if (lines && (Boolean(o.smooth) !== Boolean(s.smooth) || (o.marker || "circle") !== (s.marker || "circle"))) return false;
      return true;
    }) || "";
  }
  /** The chart in a style (what the style sets is replaced; everything else of the chart stays). */
  function applyChartStyle(chart, key) {
    const s = CHART_STYLES[key];
    if (!chart || !s || chart.style) return chart;
    const { labels: _l, legend: _g, grid: _d, gap: _p, smooth: _s, marker: _m, ...rest } = chart.opts || {};
    const opts = { ...rest };
    if (!s.labels) opts.labels = false;
    if (s.legend !== "top") opts.legend = s.legend;
    if (!s.grid) opts.grid = false;
    if (BAR_GAP_CHARTS.has(chart.type) && s.gap != null) opts.gap = s.gap;
    if (LINE_CHART_KINDS.has(chart.type)) {
      if (s.smooth) opts.smooth = true;
      if (s.marker && s.marker !== "circle") opts.marker = s.marker;
    }
    const out = { ...chart };
    if (s.colors) out.colors = s.colors; else delete out.colors;
    if (Object.keys(opts).length) out.opts = opts; else delete out.opts;
    return out;
  }
  /** A chart drawn by the engine's own charts (the same look as the layouts' charts), with an optional title. */
  // Kinds only the studio's own charts draw: an imported chart turned into one of them keeps its colours' order only.
  const STUDIO_ONLY_CHARTS = new Set(["hbar", "stacked-hbar", "stacked-area", "bubble", "histogram", "boxplot", "treemap", "sunburst", "pareto"]);
  function chartBody(o, rotEl) {
    if (o.chart.style && !STUDIO_ONLY_CHARTS.has(o.chart.type)) {
      rotEl.append(h("div", { class: "hs-obj-chart is-office" }, officeChart(o.chart, o.w, o.h)));
      return;
    }
    const box = h("div", { class: "hs-obj-chart", "data-colors": o.chart.colors || null });
    for (const [name, value] of Object.entries(chartColorVars(o.chart.colors))) box.style.setProperty(name, value);
    if (o.chart.title) box.append(h("div", { class: "hs-obj-chart-title" }, o.chart.title));
    const titleH = o.chart.title ? 52 : 0;
    const shown = visibleChart(o.chart);
    const table = o.chart.opts?.table && TABLE_CHARTS.has(o.chart.type) ? dataTable(shown) : null;
    const tableH = table ? Math.round(Math.min(o.h * 0.45, (shown.series.length + 1) * 32 + 4)) : 0;
    box.append(E.chart(chartSpec(o.chart), { w: Math.max(240, Math.round(o.w)), h: Math.max(140, Math.round(o.h - titleH - tableH)), key: `obj:${o.id}` }));
    if (table) { table.style.height = `${tableH}px`; box.append(table); }
    rotEl.append(box);
  }

  /** The trim, fades, volume and the rest on a media element, for the player (mediaPlay) to follow. */
  function playbackData(el, o) {
    const set = (key, value) => { if (value != null && value !== false) el.setAttribute(`data-${key}`, value === true ? "" : String(value)); };
    set("trim-start", o.trimStart);
    set("trim-end", o.trimEnd);
    set("fade-in", o.fadeIn);
    set("fade-out", o.fadeOut);
    set("volume", o.volume);
    set("rewind", o.rewind);
    set("across", o.across);
    set("fullscreen", o.fullscreen);
    if (o.bookmarks?.length) el.setAttribute("data-bookmarks", JSON.stringify(o.bookmarks.map(({ id, t }) => ({ id, t }))));
  }
  /** The bookmarks of a media element (from data-bookmarks): [{ id, t }]. */
  function mediaMarks(el) {
    try { const list = JSON.parse(el?.getAttribute("data-bookmarks") || "[]"); return Array.isArray(list) ? list.filter((b) => b && Number.isFinite(b.t)) : []; } catch { return []; }
  }
  /** The bookmark marks on the show's bar (a sound) or over the picture (a video), placed along the trimmed part. */
  function bookmarkMarks(o) {
    return (o.bookmarks || []).map((b) => h("b", { class: "hs-mark", "data-t": String(b.t), "data-mark": b.id, title: `${b.name}（${Math.floor(b.t / 60)}:${String(Math.floor(b.t % 60)).padStart(2, "0")}）`, role: "button", "aria-label": `ブックマーク：${b.name}` }));
  }
  function placeMarks(el) {
    const host = el.closest(".hs-obj");
    if (!host) return;
    const [start, end] = mediaSpan(el);
    for (const mark of host.querySelectorAll(".hs-mark[data-t]")) {
      const t = Number(mark.dataset.t);
      const inside = Number.isFinite(end) && t >= start && t <= end;
      mark.hidden = !inside;
      if (inside) mark.style.left = `${(((t - start) / Math.max(0.01, end - start)) * 100).toFixed(3)}%`;
    }
  }
  /** Bookmark marks placed once the length is known (and again if it changes). */
  function watchMarks(el) {
    if (!el.hasAttribute("data-bookmarks")) return;
    const place = () => placeMarks(el);
    el.addEventListener("loadedmetadata", place);
    el.addEventListener("durationchange", place);
    if (el.readyState >= 1) (globalThis.requestAnimationFrame || setTimeout)(place);
  }

  const SPEAKER = '<path d="M5 9.5h3.6L13.5 5v14L8.6 14.5H5z" fill="currentColor" stroke="none"/><path d="M16.2 9a4.2 4.2 0 0 1 0 6M18.8 6.5a7.8 7.8 0 0 1 0 11" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>';
  /**
   * A sound on the slide (挿入 → オーディオ): a speaker icon; while presenting, the sound itself and a small bar to
   * play and pause it, as PowerPoint shows when the pointer is over the icon.
   */
  function audioBody(o, rotEl, ctx) {
    const box = h("div", { class: "hs-audio", "data-hide-icon": ctx.live && o.hideIcon ? "" : null });
    const icon = s("svg", { class: "hs-audio-icon", viewBox: "0 0 24 24", "aria-hidden": "true" });
    icon.innerHTML = SPEAKER;
    icon.style.color = o.color || DEFAULTS.audio.color;
    box.append(icon);
    const url = ctx.live ? E.resolveSrc(o.src, ctx) : "";
    if (ctx.live && url) {
      const audio = h("audio", { src: url, preload: "auto", loop: o.loop && !o.trimEnd ? true : null, "data-autoplay": o.autoplay ? "" : null, "data-loop": o.loop ? "" : null, "data-narration": o.narration ? "" : null });
      playbackData(audio, o);
      box.append(audio, h("div", { class: "hs-audio-bar" },
        h("button", { type: "button", class: "hs-audio-play", "aria-label": "再生・一時停止", title: "再生・一時停止" }),
        h("span", { class: "hs-audio-track" }, h("i"), ...bookmarkMarks(o)),
        h("span", { class: "hs-audio-time" }, "0:00")));
      watchMarks(audio);
    } else if (ctx.mode !== "thumb" && o.fileName) box.title = o.fileName;
    rotEl.append(box);
  }

  // ---------------------------------------------------------------- SmartArt (挿入 → SmartArt)
  // A diagram drawn from a list of items with levels (PowerPoint's テキスト ウィンドウ): the layout places boxes,
  // arrows and lines in the theme's colours (saScheme). The parts are ordinary shapes and lines,
  // so 図形に変換 hands them over as objects, and an animation can bring the items in one by one.

  const SMARTART_LAYOUTS = {
    blocks: { label: "基本ブロック リスト", group: "リスト" },
    vlist: { label: "縦方向箇条書きリスト", group: "リスト" },
    process: { label: "基本ステップ", group: "手順" },
    chevron: { label: "矢印型ステップ", group: "手順" },
    vprocess: { label: "縦方向ステップ", group: "手順" },
    steps: { label: "上向きステップ", group: "手順" },
    timeline: { label: "基本タイムライン", group: "手順" },
    cycle: { label: "基本の循環", group: "循環" },
    radial: { label: "基本の放射", group: "循環" },
    hierarchy: { label: "組織図", group: "階層構造" },
    venn: { label: "基本ベン図", group: "集合関係" },
    target: { label: "ターゲット", group: "集合関係" },
    matrix: { label: "基本マトリックス", group: "マトリックス" },
    pyramid: { label: "基本ピラミッド", group: "ピラミッド" },
  };
  const SMARTART_GROUPS = ["リスト", "手順", "循環", "階層構造", "集合関係", "マトリックス", "ピラミッド"];
  // SmartArt のデザイン → 色の変更 (PowerPoint's colour sets), made from the theme's colours when the SmartArt is drawn.
  const SMARTART_COLORS = { accent1: "アクセント 1", colorful: "カラフル - アクセント 2 から 4", light: "アクセント 1 - 淡色", gray: "グレー", outline: "アクセント 1 - 線のみ" };
  const SMARTART_STYLES = { round: "角丸", square: "四角", soft: "丸みを強く" };
  /** A SmartArt colour set in the current theme: fills of the main and sub shapes, the arrows, the lines and the words. */
  function saScheme(key) {
    const c = themeColors(PALETTE.sheet?.themeId);
    const tint = (hex, pct) => lumColor(hex, (100 - pct) * 1000, pct * 1000);
    const base = { ink: c.bg1, subInk: c.tx1, tx: c.tx1, bg: c.bg1 };
    switch (key) {
      case "colorful": return { ...base, main: [c.accent2, c.accent3, c.accent4], sub: [tint(c.accent2, 80), tint(c.accent3, 80), tint(c.accent4, 80)], arrow: tint(c.accent2, 40), line: c.accent2 };
      case "light": return { ...base, ink: c.tx1, main: [tint(c.accent1, 60)], sub: [tint(c.accent1, 80)], arrow: tint(c.accent1, 40), line: c.accent1 };
      case "gray": return { ...base, ink: c.tx1, main: [lumColor(c.bg1, 85000, 0)], sub: [lumColor(c.bg1, 95000, 0)], arrow: lumColor(c.bg1, 75000, 0), line: c.tx2 };
      case "outline": return { ...base, ink: c.tx1, main: [c.bg1], sub: [c.bg1], arrow: tint(c.accent1, 40), line: c.accent1, stroke: c.accent1, subStroke: tint(c.accent1, 40) };
      default: return { ...base, main: [c.accent1], sub: [tint(c.accent1, 80)], arrow: tint(c.accent1, 40), line: c.accent1 };
    }
  }
  const SA_SAMPLE = {
    hierarchy: [["社長", 0], ["営業本部", 1], ["第1営業部", 2], ["第2営業部", 2], ["管理本部", 1], ["総務部", 2]],
    radial: [["お客様", 0], ["品ぞろえ", 0], ["価格", 0], ["接客", 0], ["売場", 0]],
    venn: [["品質", 0], ["価格", 0], ["利便性", 0]],
    matrix: [["強み×機会", 0], ["強み×脅威", 0], ["弱み×機会", 0], ["弱み×脅威", 0]],
  };
  /** Items to start a new SmartArt with (PowerPoint shows [テキスト]; the studio gives words to replace). */
  function smartartSample(layout) {
    const list = SA_SAMPLE[layout] || [["計画", 0], ["目標と期限を決める", 1], ["実行", 0], ["店舗で試す", 1], ["評価", 0], ["数字で振り返る", 1]];
    return list.map(([text, level]) => ({ text, level }));
  }

  function normalizeSmartart(raw) {
    const sa = raw && typeof raw === "object" ? raw : {};
    const items = (Array.isArray(sa.items) ? sa.items : []).slice(0, 60)
      .map((it) => ({ text: String(it?.text ?? "").replace(/[\u0000-\u0008\u000b-\u001f\u007f]/g, "").replace(/\r?\n/g, " ").slice(0, 300), level: Math.round(clamp(Number(it?.level) || 0, 0, 4)) }));
    if (!items.length) items.push({ text: "", level: 0 });
    items[0].level = 0;
    for (let i = 1; i < items.length; i += 1) items[i].level = Math.min(items[i].level, items[i - 1].level + 1);
    const out = { layout: SMARTART_LAYOUTS[sa.layout] ? sa.layout : "blocks", items };
    if (SMARTART_COLORS[sa.color] && sa.color !== "accent1") out.color = sa.color;
    if (SMARTART_STYLES[sa.style] && sa.style !== "round") out.style = sa.style;
    const fsScale = num(sa.fsScale, 0.5, 2);
    if (fsScale && fsScale !== 1) out.fsScale = fsScale;
    if (sa.rtl === true) out.rtl = true;
    return out;
  }

  /** Top-level items, each with the items under it (a deeper level belongs to the item above it). */
  function saTree(items) {
    const tops = [];
    items.forEach((it, index) => {
      if (it.level === 0 || !tops.length) tops.push({ ...it, index, kids: [] });
      else tops[tops.length - 1].kids.push({ ...it, index });
    });
    return tops;
  }
  const emWidth = (text) => [...String(text)].reduce((sum, ch) => sum + (/[\u0000-ÿ｡-ﾟ]/.test(ch) ? 0.56 : 1), 0);
  /** The biggest size (px) at which the lines fit the box, wrapping as they would (a rough measure, no layout). */
  function fitFs(lines, w, hh, { max = 44, min = 12, lh = 1.3, padX = 24, padY = 14 } = {}) {
    const innerW = Math.max(8, w - padX);
    const innerH = Math.max(8, hh - padY);
    const list = lines.filter((line) => line !== undefined).map((line) => String(line || " "));
    for (let fs = max; fs > min; fs -= 1) {
      const rows = list.reduce((n, line) => n + Math.max(1, Math.ceil((emWidth(line) * fs) / innerW - 0.02)), 0);
      if (rows * fs * lh <= innerH && list.every((line) => Math.min(emWidth(line), 4) * fs <= innerW)) return fs;
    }
    return min;
  }

  /**
   * The parts of a SmartArt in its own box (0…w, 0…h): shapes ({ kind: "shape", … }) and arrows/lines
   * ({ kind: "line", … }), each with `step` (the top-level item it belongs to, -1 for what is always there) and
   * `item` (the item whose words it shows).
   */
  function smartartParts(o) {
    const sa = o.smartart || normalizeSmartart(null);
    const w = Math.max(40, o.w);
    const hh = Math.max(40, o.h);
    const scheme = saScheme(sa.color);
    const tops = saTree(sa.items);
    const n = Math.max(1, tops.length);
    const k = sa.fsScale || 1;
    const corner = sa.style === "square" ? "rect" : "roundRect";
    const cornerAdj = sa.style === "soft" ? [0.35] : undefined;
    const mainFill = (i) => scheme.main[i % scheme.main.length];
    const subFill = (i) => scheme.sub[i % scheme.sub.length];
    const outline = scheme.stroke ? { stroke: scheme.stroke, strokeW: 3 } : { stroke: "none" };
    const subOutline = scheme.subStroke ? { stroke: scheme.subStroke, strokeW: 2 } : { stroke: "none" };
    const parts = [];
    const node = (x, y, nw, nh, text, extra = {}) => parts.push({ kind: "shape", shape: corner, adj: corner === "roundRect" ? cornerAdj : undefined, x, y, w: nw, h: nh, fill: mainFill(extra.step ?? 0), ...outline, text, color: scheme.ink, align: "center", valign: "middle", pad: [8, 12, 8, 12], lh: 1.3, autofit: "none", role: "node", ...extra });
    const arrow = (x1, y1, x2, y2, extra = {}) => parts.push({ kind: "line", x1, y1, x2, y2, stroke: scheme.line, strokeW: 4, tail: "triangle", tailSize: 2, role: "connector", ...extra });
    const rich = (text) => (text ? textToRich(text) : "");
    const list = (kids, centered = false) => (!kids.length ? "" : centered ? kids.map((kid) => `<p>${escapeHtml(kid.text) || "<br>"}</p>`).join("")
      : `<ul>${kids.map((kid) => `<li>${escapeHtml(kid.text) || "<br>"}</li>`).join("")}</ul>`);
    const uniform = (boxes, opts) => Math.round(Math.min(...boxes.map(([lines, bw, bh]) => fitFs(lines, bw, bh, opts))) * k);
    const rtl = (x, bw) => (sa.rtl ? w - x - bw : x);
    const layout = sa.layout;

    if (layout === "blocks") {
      // As many columns as make the boxes closest to 16:10.
      let cols = 1;
      let best = Infinity;
      for (let c = 1; c <= n; c += 1) {
        const rows = Math.ceil(n / c);
        const ratio = (w / c) / (hh / rows);
        const score = Math.abs(Math.log(ratio / 1.6));
        if (score < best) { best = score; cols = c; }
      }
      const rows = Math.ceil(n / cols);
      const gap = Math.min(w, hh) * 0.04;
      const bw = (w - gap * (cols - 1)) / cols;
      const bh = (hh - gap * (rows - 1)) / rows;
      const fs = uniform(tops.map((t) => [[t.text, ...t.kids.map((kid) => kid.text)], bw, bh]));
      // The last row, when it is not full, sits in the middle (as PowerPoint's 基本ブロック リスト).
      const lastRow = rows - 1;
      const inLast = n - lastRow * cols;
      const shift = inLast < cols ? ((cols - inLast) * (bw + gap)) / 2 : 0;
      tops.forEach((t, i) => node(rtl((i % cols) * (bw + gap) + (Math.floor(i / cols) === lastRow ? shift : 0), bw), Math.floor(i / cols) * (bh + gap), bw, bh, rich(t.text) + list(t.kids), { step: i, item: t.index, fs, align: t.kids.length ? "left" : "center", fill: mainFill(i) }));
    } else if (layout === "vlist") {
      const gap = hh * 0.03;
      const rowH = (hh - gap * (n - 1)) / n;
      const headW = w * 0.3;
      const fs = uniform(tops.map((t) => [[t.text], headW, rowH]));
      const subFs = uniform(tops.map((t) => [t.kids.map((kid) => kid.text).concat(t.kids.length ? [] : [""]), w - headW, rowH]), { max: Math.max(12, fs - 4) });
      tops.forEach((t, i) => {
        const y = i * (rowH + gap);
        parts.push({ kind: "shape", shape: "rect", x: rtl(headW * 0.5, w - headW * 0.5), y, w: w - headW * 0.5, h: rowH, fill: subFill(i), ...subOutline, text: list(t.kids), color: scheme.subInk, align: "left", valign: "middle", pad: [8, 16, 8, headW * 0.5 + 20], lh: 1.3, autofit: "none", fs: subFs, role: "sub", step: i, item: t.kids[0]?.index ?? t.index });
        node(rtl(0, headW), y, headW, rowH, rich(t.text), { step: i, item: t.index, fs, fill: mainFill(i) });
      });
    } else if (layout === "process" || layout === "chevron") {
      const chevron = layout === "chevron";
      const gap = chevron ? -w * 0.012 : w * 0.06;
      const bw = (w - gap * (n - 1)) / n;
      const hasKids = tops.some((t) => t.kids.length);
      const bh = chevron ? Math.min(hh * (hasKids ? 0.36 : 0.7), bw * 0.55) : Math.min(hh * (hasKids ? 0.5 : 0.8), bw * 0.8);
      const y = hasKids ? hh * 0.05 : (hh - bh) / 2;
      const fs = uniform(tops.map((t) => [[t.text], chevron ? bw * 0.7 : bw, bh]));
      const subFs = uniform(tops.map((t) => [t.kids.map((kid) => kid.text).concat([""]), bw, hh - bh - y - hh * 0.06]), { max: Math.max(12, fs - 4) });
      tops.forEach((t, i) => {
        const x = rtl(i * (bw + gap), bw);
        if (chevron) parts.push({ kind: "shape", shape: i === 0 && !sa.rtl ? "homePlate" : "chevron", flipH: sa.rtl || undefined, adj: [0.32], x, y, w: bw, h: bh, fill: mainFill(i), ...outline, text: rich(t.text), color: scheme.ink, align: "center", valign: "middle", pad: [6, bh * 0.36, 6, bh * 0.36], lh: 1.25, autofit: "none", fs, role: "node", step: i, item: t.index });
        else node(x, y, bw, bh, rich(t.text), { step: i, item: t.index, fs, fill: mainFill(i) });
        if (t.kids.length) parts.push({ kind: "shape", shape: "rect", x, y: y + bh + hh * 0.04, w: bw, h: hh - bh - y - hh * 0.06, fill: "none", stroke: "none", text: list(t.kids), color: scheme.tx, align: "left", valign: "top", pad: [6, 10, 6, 10], lh: 1.35, autofit: "none", fs: subFs, role: "sub", step: i, item: t.kids[0].index });
        if (!chevron && i > 0) {
          const ax = sa.rtl ? x + bw + gap * 0.82 : x - gap * 0.82;
          const bx = sa.rtl ? x + bw + gap * 0.18 : x - gap * 0.18;
          parts.push({ kind: "shape", shape: sa.rtl ? "leftArrow" : "rightArrow", x: Math.min(ax, bx), y: y + bh / 2 - Math.min(bh * 0.18, gap * 0.4), w: Math.abs(bx - ax), h: Math.min(bh * 0.36, gap * 0.8), fill: scheme.arrow, stroke: "none", text: "", color: scheme.tx, align: "center", valign: "middle", pad: [0, 0, 0, 0], lh: 1.2, autofit: "none", role: "connector", step: i });
        }
      });
    } else if (layout === "vprocess") {
      const gap = hh * 0.08;
      const bh = (hh - gap * (n - 1)) / n;
      const hasKids = tops.some((t) => t.kids.length);
      const bw = hasKids ? w * 0.36 : Math.min(w, Math.max(w * 0.5, bh * 4));
      const x0 = hasKids ? 0 : (w - bw) / 2;
      const fs = uniform(tops.map((t) => [[t.text], bw, bh]));
      const subFs = uniform(tops.map((t) => [t.kids.map((kid) => kid.text).concat([""]), w - bw - w * 0.04, bh]), { max: Math.max(12, fs - 4) });
      tops.forEach((t, i) => {
        const y = i * (bh + gap);
        node(rtl(x0, bw), y, bw, bh, rich(t.text), { step: i, item: t.index, fs, fill: mainFill(i) });
        if (t.kids.length) parts.push({ kind: "shape", shape: "rect", x: rtl(bw + w * 0.04, w - bw - w * 0.04), y, w: w - bw - w * 0.04, h: bh, fill: subFill(i), ...subOutline, text: list(t.kids), color: scheme.subInk, align: "left", valign: "middle", pad: [6, 14, 6, 14], lh: 1.3, autofit: "none", fs: subFs, role: "sub", step: i, item: t.kids[0].index });
        if (i > 0) parts.push({ kind: "shape", shape: "downArrow", x: rtl(x0 + bw / 2 - gap * 0.45, gap * 0.9), y: y - gap * 0.85, w: gap * 0.9, h: gap * 0.7, fill: scheme.arrow, stroke: "none", text: "", color: scheme.tx, align: "center", valign: "middle", pad: [0, 0, 0, 0], lh: 1.2, autofit: "none", role: "connector", step: i });
      });
    } else if (layout === "steps") {
      const gap = w * 0.015;
      const bw = (w - gap * (n - 1)) / n;
      const fs = uniform(tops.map((t) => [[t.text, ...t.kids.map((kid) => kid.text)], bw, hh / Math.max(2, n) * 1.4]));
      // The boxes rise from left to right; an arrow runs just above their top-left corners.
      const heightOf = (i) => hh * (0.3 + (0.58 * (i + 1)) / n);
      tops.forEach((t, i) => {
        const bh = heightOf(i);
        node(rtl(i * (bw + gap), bw), hh - bh, bw, bh, rich(t.text) + list(t.kids), { step: i, item: t.index, fs, valign: "top", align: t.kids.length ? "left" : "center", fill: mainFill(i), shape: "rect" });
      });
      const lift = hh * 0.05;
      const slope = (heightOf(1) - heightOf(0)) / (bw + gap);
      const y0 = hh - heightOf(0) - lift;
      // It stops at the top of the box (keeping its slope, so it never cuts into the tall boxes).
      const reach = Math.min(w, slope > 0 ? (y0 - 6) / slope : w);
      parts.push({ kind: "line", x1: rtl(0, 0), y1: y0, x2: rtl(reach, 0), y2: y0 - slope * reach, stroke: scheme.line, strokeW: 4, tail: "triangle", tailSize: 3, role: "deco", step: -1 });
    } else if (layout === "timeline") {
      const y = hh / 2;
      parts.push({ kind: "line", x1: 0, y1: y, x2: w, y2: y, stroke: scheme.line, strokeW: 6, tail: "triangle", tailSize: 3, role: "deco", step: -1 });
      const slot = w / n;
      const bh = hh * 0.36;
      const fs = uniform(tops.map((t) => [[t.text, ...t.kids.map((kid) => kid.text)], slot * 0.92, bh]));
      tops.forEach((t, i) => {
        const cx = rtl(slot * i + slot / 2, 0);
        const up = i % 2 === 0;
        const dot = Math.min(hh * 0.07, slot * 0.18);
        parts.push({ kind: "shape", shape: "ellipse", x: cx - dot / 2, y: y - dot / 2, w: dot, h: dot, fill: scheme.line, stroke: "none", text: "", color: scheme.tx, align: "center", valign: "middle", pad: [0, 0, 0, 0], lh: 1.2, autofit: "none", role: "connector", step: i });
        node(cx - slot * 0.46, up ? y - dot - bh - hh * 0.04 : y + dot + hh * 0.04, slot * 0.92, bh, rich(t.text) + list(t.kids, true), { step: i, item: t.index, fs, fill: mainFill(i), valign: up ? "bottom" : "top" });
      });
    } else if (layout === "cycle" || layout === "radial") {
      const radial = layout === "radial";
      const ring = radial ? tops.slice(1) : tops;
      const m = Math.max(1, ring.length);
      const cx = w / 2;
      const cy = hh / 2;
      const size = Math.min(w, hh);
      // As big as the circle allows without the items touching (and, radial, without touching the centre).
      const sinA = Math.sin(Math.PI / Math.max(2, m));
      const nodeR = Math.min(size * (radial ? 0.155 : 0.2), (0.4 * sinA * size) / (1 + 0.8 * sinA));
      const R = size / 2 - nodeR - 4;
      // A circle's words sit in the square inside it (about 1.4 × the radius wide).
      const linesOf = (t) => [t.text, ...t.kids.map((kid) => kid.text)];
      const fs = uniform([...ring.map((t) => [linesOf(t), nodeR * 1.42, nodeR * 1.42]), ...(radial && tops[0] ? [[linesOf(tops[0]), size * 0.24, size * 0.24]] : [])], { padX: 16, padY: 12 });
      if (radial && tops[0]) parts.push({ kind: "shape", shape: "ellipse", x: cx - size * 0.17, y: cy - size * 0.17, w: size * 0.34, h: size * 0.34, fill: mainFill(0), ...outline, text: rich(tops[0].text) + list(tops[0].kids, true), color: scheme.ink, align: "center", valign: "middle", pad: [8, 12, 8, 12], lh: 1.25, autofit: "none", fs: Math.round(fs * 1.15), role: "node", step: 0, item: tops[0].index });
      const at = (i) => { const a = -Math.PI / 2 + (sa.rtl ? -1 : 1) * ((2 * Math.PI * i) / m); return [cx + R * Math.cos(a), cy + R * Math.sin(a), a]; };
      ring.forEach((t, i) => {
        const [x, y] = at(i);
        const step = radial ? i + 1 : i;
        parts.push({ kind: "shape", shape: "ellipse", x: x - nodeR, y: y - nodeR, w: nodeR * 2, h: nodeR * 2, fill: mainFill(step), ...outline, text: rich(t.text) + list(t.kids, true), color: scheme.ink, align: "center", valign: "middle", pad: [6, 8, 6, 8], lh: 1.25, autofit: "none", fs, role: "node", step, item: t.index });
        if (radial) {
          const len = Math.hypot(x - cx, y - cy) || 1;
          const ux = (x - cx) / len;
          const uy = (y - cy) / len;
          parts.push({ kind: "line", x1: cx + ux * size * 0.17, y1: cy + uy * size * 0.17, x2: x - ux * nodeR, y2: y - uy * nodeR, stroke: scheme.line, strokeW: 4, role: "connector", step });
        } else if (m > 1) {
          // An arrow along the circle from the item before to this one.
          const [px, py, pa] = at((i - 1 + m) % m);
          const [, , a] = at(i);
          const da = (Math.asin(Math.min(1, (nodeR * 1.15) / R)));
          const dir = sa.rtl ? -1 : 1;
          const a0 = pa + dir * da;
          let a1 = a - dir * da;
          if (!sa.rtl && a1 < a0) a1 += 2 * Math.PI;
          if (sa.rtl && a1 > a0) a1 -= 2 * Math.PI;
          const am = (a0 + a1) / 2;
          parts.push({ kind: "line", x1: cx + R * Math.cos(a0), y1: cy + R * Math.sin(a0), x2: cx + R * Math.cos(a1), y2: cy + R * Math.sin(a1), via: [cx + R * 1.0 * Math.cos(am) * 1.0, cy + R * Math.sin(am)], stroke: scheme.arrow === "#d9d9d9" ? "#808080" : scheme.line, strokeW: 4, tail: "triangle", tailSize: 2, role: "connector", step, curve: true });
          void px; void py;
        }
      });
    } else if (layout === "hierarchy") {
      // A tree by level: each item hangs under the nearest item above it with a smaller level.
      const nodes = sa.items.map((it, index) => ({ ...it, index, kids: [], parent: null }));
      const stack = [];
      const roots = [];
      for (const nd of nodes) {
        while (stack.length && stack[stack.length - 1].level >= nd.level) stack.pop();
        if (stack.length) { nd.parent = stack[stack.length - 1]; nd.parent.kids.push(nd); } else roots.push(nd);
        stack.push(nd);
      }
      const depth = Math.max(...nodes.map((nd) => nd.level)) + 1;
      const leaves = (nd) => (nd.kids.length ? nd.kids.reduce((sum, kid) => sum + leaves(kid), 0) : 1);
      const total = roots.reduce((sum, r) => sum + leaves(r), 0);
      const colW = w / total;
      const rowGap = hh * 0.12;
      const rowH = (hh - rowGap * (depth - 1)) / depth;
      const bw = Math.min(colW * 0.86, w * 0.3);
      const fs = uniform(nodes.map((nd) => [[nd.text], bw, rowH]));
      let col = 0;
      const place = (nd) => {
        if (!nd.kids.length) { nd.cx = colW * (col + 0.5); col += 1; } else { nd.kids.forEach(place); nd.cx = (nd.kids[0].cx + nd.kids[nd.kids.length - 1].cx) / 2; }
        nd.cy = nd.level * (rowH + rowGap);
      };
      roots.forEach(place);
      // Steps follow the outline order, so "one by one" comes in from the top as the text pane lists them.
      nodes.forEach((nd, i) => {
        const x = rtl(nd.cx - bw / 2, bw);
        node(x, nd.cy, bw, rowH, rich(nd.text), { step: i, item: nd.index, fs, fill: nd.level === 0 ? mainFill(0) : nd.level === 1 ? (scheme.main[1] || subFill(0)) : subFill(i) });
        if (nd.parent) {
          const px = rtl(nd.parent.cx, 0);
          const cxs = rtl(nd.cx, 0);
          const midY = nd.cy - rowGap / 2;
          parts.push({ kind: "line", x1: px, y1: nd.parent.cy + rowH, x2: cxs, y2: nd.cy, elbow: midY, stroke: scheme.line, strokeW: 3, role: "connector", step: i });
        }
      });
    } else if (layout === "venn") {
      const m = Math.min(n, 6);
      const size = Math.min(w, hh);
      const r = m === 1 ? size * 0.45 : size * (m <= 3 ? 0.3 : 0.26);
      const R = m === 1 ? 0 : m === 2 ? r * 0.62 : r * 0.72;
      const fs = uniform(tops.slice(0, m).map((t) => [[t.text, ...t.kids.map((kid) => kid.text)], r * 1.1, r * 0.8]));
      // The circles' centres, then the whole set moved to the middle of the box.
      const angle = (i) => (m === 2 ? Math.PI * i + Math.PI : -Math.PI / 2 + (2 * Math.PI * i) / m);
      const ys = tops.slice(0, m).map((_, i) => R * Math.sin(angle(i)) * (m === 2 ? 0 : 1));
      const dy = -(Math.min(...ys) + Math.max(...ys)) / 2;
      tops.slice(0, m).forEach((t, i) => {
        const a = angle(i);
        const cx = w / 2 + R * Math.cos(a);
        const cy = hh / 2 + ys[i] + dy;
        // The words sit in the part of the circle away from the others.
        const tx = cx + Math.cos(a) * r * 0.32;
        const ty = cy + Math.sin(a) * r * 0.32 * (m === 2 ? 0 : 1);
        parts.push({ kind: "shape", shape: "ellipse", x: cx - r, y: cy - r, w: r * 2, h: r * 2, fill: mainFill(i), fillOpacity: scheme.stroke ? undefined : 0.75, ...(scheme.stroke ? { fill: "none", stroke: scheme.stroke, strokeW: 3 } : { stroke: "none" }), text: "", color: scheme.ink, align: "center", valign: "middle", pad: [0, 0, 0, 0], lh: 1.25, autofit: "none", role: "node", step: i, item: t.index });
        parts.push({ kind: "shape", shape: "rect", x: tx - r * 0.55, y: ty - r * 0.4, w: r * 1.1, h: r * 0.8, fill: "none", stroke: "none", text: rich(t.text) + list(t.kids, true), color: scheme.tx, align: "center", valign: "middle", pad: [0, 0, 0, 0], lh: 1.25, autofit: "none", fs, role: "label", step: i, item: t.index });
      });
    } else if (layout === "target") {
      const size = Math.min(hh, w * 0.55);
      const m = n;
      const cx = sa.rtl ? w - size / 2 : size / 2;
      const cy = hh / 2;
      const fs = uniform(tops.map((t) => [[t.text, ...t.kids.map((kid) => kid.text)], w - size * 1.05, hh / m]));
      tops.forEach((t, i) => {
        const r = (size / 2) * (1 - i / m);
        parts.push({ kind: "shape", shape: "ellipse", x: cx - r, y: cy - r, w: r * 2, h: r * 2, fill: mainFill(i), ...(scheme.stroke ? { fill: scheme.bg, stroke: scheme.stroke, strokeW: 3 } : { stroke: scheme.bg, strokeW: 3 }), text: "", color: scheme.ink, align: "center", valign: "middle", pad: [0, 0, 0, 0], lh: 1.25, autofit: "none", role: "node", step: i, item: t.index });
        // Labels in rows beside the rings; each line starts in the middle of its own ring, pointing at its label.
        const labelH = hh / m;
        const labelY = i * labelH;
        const rMid = (r + (size / 2) * (1 - (i + 1) / m)) / 2;
        const lx1 = sa.rtl ? w - size * 1.08 : size * 1.08;
        const ly1 = labelY + labelH / 2;
        const level = ly1 - cy;
        const flat = Math.abs(level) < rMid;
        const toward = Math.atan2(level, lx1 - cx);
        const lx0 = flat ? cx + (sa.rtl ? -1 : 1) * Math.sqrt(rMid * rMid - level * level) : cx + rMid * Math.cos(toward);
        const ly0 = flat ? ly1 : cy + rMid * Math.sin(toward);
        parts.push({ kind: "line", x1: lx0, y1: ly0, x2: lx1, y2: ly1, stroke: scheme.line, strokeW: 2, head: "oval", headSize: 1, role: "connector", step: i });
        parts.push({ kind: "shape", shape: "rect", x: sa.rtl ? 0 : size * 1.1, y: labelY, w: w - size * 1.1, h: labelH, fill: "none", stroke: "none", text: rich(t.text) + list(t.kids), color: scheme.tx, align: sa.rtl ? "right" : "left", valign: "middle", pad: [0, 8, 0, 8], lh: 1.25, autofit: "none", fs, role: "label", step: i, item: t.index });
      });
    } else if (layout === "matrix") {
      const gap = Math.min(w, hh) * 0.03;
      const bw = (w - gap) / 2;
      const bh = (hh - gap) / 2;
      const four = tops.slice(0, 4);
      const fs = uniform(four.map((t) => [[t.text, ...t.kids.map((kid) => kid.text)], bw, bh]));
      four.forEach((t, i) => node(rtl((i % 2) * (bw + gap), bw), Math.floor(i / 2) * (bh + gap), bw, bh, rich(t.text) + list(t.kids), { step: i, item: t.index, fs, fill: mainFill(i), align: t.kids.length ? "left" : "center" }));
    } else if (layout === "pyramid") {
      const bandH = hh / n;
      // The words fit where the band is wide enough: the lower half of the top triangle, the middle of the others.
      const roomOf = (i) => (i === 0 ? (w / n) * 0.6 : ((w * i) / n + (w * (i + 1)) / n) / 2 * 0.8);
      const fitBand = (t, i) => [[t.text, ...t.kids.map((kid) => kid.text)], roomOf(i), i === 0 ? bandH * 0.55 : bandH];
      // The bands share one size; the narrow top may need a smaller one of its own.
      const lower = tops.length > 1 ? uniform(tops.slice(1).map((t, i) => fitBand(t, i + 1))) : Infinity;
      const fsOf = (t, i) => Math.min(lower, Math.round(fitFs(...fitBand(t, i)) * k));
      tops.forEach((t, i) => {
        const y = i * bandH;
        const topW = (w * i) / n;
        const botW = (w * (i + 1)) / n;
        // A band: a triangle at the top, trapezoids below (the top width as the trapezoid's adjustment).
        const shape = i === 0 ? "triangle" : "trapezoid";
        const adj = i === 0 ? [0.5] : [Math.min(0.5, ((botW - topW) / 2) / bandH)];
        parts.push({ kind: "shape", shape, adj, x: (w - botW) / 2, y, w: botW, h: bandH - 3, fill: mainFill(i), ...(scheme.stroke ? outline : { stroke: scheme.bg, strokeW: 3 }), text: rich(t.text) + list(t.kids, true), color: scheme.ink, align: "center", valign: i === 0 ? "bottom" : "middle", pad: [6, 10, 6, 10], lh: 1.25, autofit: "none", fs: fsOf(t, i), role: "node", step: i, item: t.index });
      });
    }
    // Words on a fill take whichever of 背景 1 and テキスト 1 reads better on it (white on a light accent would not).
    const readable = (fill) => (contrastRatio(scheme.bg, fill) >= contrastRatio(scheme.tx, fill) ? scheme.bg : scheme.tx);
    return parts.map((part) => (part.kind === "shape" ? { ...part, fs: part.fs || Math.round(24 * k), ...(part.text && part.fill && part.fill !== "none" && !part.fillOpacity ? { color: readable(part.fill) } : {}) } : part));
  }

  /** The polyline a connector of a SmartArt follows (straight, an elbow, or along a circle). */
  function saLinePoints(part) {
    if (part.elbow != null) return [[part.x1, part.y1], [part.x1, part.elbow], [part.x2, part.elbow], [part.x2, part.y2]];
    if (part.curve && part.via) {
      // A quadratic through `via`, flattened.
      const [vx, vy] = part.via;
      const cx = 2 * vx - (part.x1 + part.x2) / 2;
      const cy = 2 * vy - (part.y1 + part.y2) / 2;
      const pts = [];
      for (let i = 0; i <= 16; i += 1) { const t = i / 16; const u = 1 - t; pts.push([u * u * part.x1 + 2 * u * t * cx + t * t * part.x2, u * u * part.y1 + 2 * u * t * cy + t * t * part.y2]); }
      return pts;
    }
    return [[part.x1, part.y1], [part.x2, part.y2]];
  }

  /** A SmartArt drawn: one box per top-level item (its shapes and the connector into it), for "one by one". */
  function smartartBody(o, rotEl, scale) {
    const parts = smartartParts(o);
    const root = h("div", { class: "hs-smartart", "data-layout": o.smartart?.layout || "blocks" });
    const steps = new Map();
    for (const part of parts) { if (!steps.has(part.step)) steps.set(part.step, []); steps.get(part.step).push(part); }
    const order = [...steps.keys()].sort((a, b) => a - b);
    for (const step of order) {
      const group = steps.get(step);
      const boxes = group.flatMap((part) => (part.kind === "line" ? saLinePoints(part) : [[part.x, part.y], [part.x + part.w, part.y + part.h]]));
      const minX = Math.min(...boxes.map((p) => p[0])) - 8;
      const minY = Math.min(...boxes.map((p) => p[1])) - 8;
      const maxX = Math.max(...boxes.map((p) => p[0])) + 8;
      const maxY = Math.max(...boxes.map((p) => p[1])) + 8;
      const wrap = h("div", { class: ["hs-sa-step", step < 0 ? "hs-sa-fixed" : ""], "data-step": String(step), style: { left: `${r2(minX)}px`, top: `${r2(minY)}px`, width: `${r2(maxX - minX)}px`, height: `${r2(maxY - minY)}px` } });
      const lines = group.filter((part) => part.kind === "line");
      if (lines.length) {
        const svg = s("svg", { class: "hs-sa-lines", width: r2(maxX - minX), height: r2(maxY - minY), viewBox: `${r2(minX)} ${r2(minY)} ${r2(maxX - minX)} ${r2(maxY - minY)}`, overflow: "visible", "aria-hidden": "true" });
        for (const part of lines) {
          const pts = saLinePoints(part);
          let draw = pts;
          const heads = [];
          if (part.tail) {
            const head = arrowhead(part.tail, pts[pts.length - 1], pts[pts.length - 2], part.strokeW, part.tailSize ?? 2);
            if (head) { heads.push(head); const [tip, from] = [pts[pts.length - 1], pts[pts.length - 2]]; const len = Math.hypot(tip[0] - from[0], tip[1] - from[1]) || 1; const back = Math.min(head.back, len * 0.9); draw = [...pts.slice(0, -1), [tip[0] - ((tip[0] - from[0]) / len) * back, tip[1] - ((tip[1] - from[1]) / len) * back]]; }
          }
          if (part.head) { const head = arrowhead(part.head, pts[0], pts[1], part.strokeW, part.headSize ?? 2); if (head) heads.push(head); }
          svg.append(s("path", { d: `M${draw.map(([x, y]) => `${r2(x)} ${r2(y)}`).join(" L")}`, fill: "none", stroke: part.stroke, "stroke-width": part.strokeW, "stroke-linejoin": "round", "stroke-linecap": "round" }));
          for (const head of heads) svg.append(s("path", { d: head.d, fill: head.fill ? part.stroke : "none", stroke: part.stroke, "stroke-width": head.fill ? Math.max(1, part.strokeW * 0.4) : part.strokeW }));
        }
        wrap.append(svg);
      }
      for (const part of group.filter((p) => p.kind === "shape")) {
        const shape = withDefaults({ ...part, kind: "shape", pad: part.pad || [8, 12, 8, 12] });
        const el = h("div", { class: ["hs-sa-node", `hs-sa-${part.role || "node"}`], "data-item": part.item != null ? String(part.item) : null, style: { left: `${r2(part.x - minX)}px`, top: `${r2(part.y - minY)}px`, width: `${r2(part.w)}px`, height: `${r2(part.h)}px` } });
        const body = h("div", { class: "hs-obj-rot" });
        if (part.flipH) body.style.transform = "scale(-1, 1)";
        shapeBody(shape, body, scale);
        if (part.flipH) { const text = body.querySelector(".hs-obj-text"); if (text) text.style.transform = "scale(-1, 1)"; }
        el.append(body);
        wrap.append(el);
      }
      root.append(wrap);
    }
    rotEl.append(root);
  }

  /** 図形に変換: the SmartArt's parts as objects on the slide (grouped), in slide coordinates. */
  function smartartObjects(o) {
    const group = newId();
    const a = rad(o.rot || 0);
    const [cos, sin] = [Math.cos(a), Math.sin(a)];
    const cx = o.x + o.w / 2;
    const cy = o.y + o.h / 2;
    const place = (x, y) => { const dx = o.x + x - cx; const dy = o.y + y - cy; return [cx + dx * cos - dy * sin, cy + dx * sin + dy * cos]; };
    return smartartParts(o).map((part) => {
      if (part.kind === "line") {
        const pts = saLinePoints(part).map(([x, y]) => place(x, y));
        const [p1, p2] = [pts[0], pts[pts.length - 1]];
        const line = { id: newId(), kind: "line", x1: r2(p1[0]), y1: r2(p1[1]), x2: r2(p2[0]), y2: r2(p2[1]), stroke: part.stroke, strokeW: part.strokeW, group, ...(part.tail ? { tail: part.tail, tailSize: part.tailSize ?? 2 } : {}), ...(part.head ? { head: part.head, headSize: part.headSize ?? 2 } : {}) };
        if (part.elbow != null) line.route = "elbow";
        if (part.curve) line.route = "curve";
        return line;
      }
      const [nx, ny] = place(part.x + part.w / 2, part.y + part.h / 2);
      const { role, step, item, via, elbow, curve, ...rest } = part;
      void role; void step; void item; void via; void elbow; void curve;
      return { ...rest, id: newId(), kind: "shape", x: r2(nx - part.w / 2), y: r2(ny - part.h / 2), w: r2(part.w), h: r2(part.h), rot: o.rot || undefined, group };
    }).map((x) => normalizeObject(x)).filter(Boolean);
  }

  // ---------------------------------------------------------------- playing media (video and sound)

  const mediaNum = (el, key) => { const v = Number(el.getAttribute(`data-${key}`)); return el.hasAttribute(`data-${key}`) && Number.isFinite(v) ? v : null; };
  /** Where a media element's trimmed part begins and ends (seconds). */
  function mediaSpan(el) {
    const start = mediaNum(el, "trim-start") || 0;
    const dur = Number.isFinite(el.duration) ? el.duration : Infinity;
    const end = Math.min(mediaNum(el, "trim-end") || dur, dur);
    return [start, end > start ? end : dur];
  }
  /** Volume at this moment: the set volume, faded in after the start and out before the end. */
  function mediaVolume(el) {
    const base = mediaNum(el, "volume") ?? 1;
    const [start, end] = mediaSpan(el);
    const t = el.currentTime;
    let k = 1;
    const fin = mediaNum(el, "fade-in");
    const fout = mediaNum(el, "fade-out");
    if (fin) k = Math.min(k, Math.max(0, (t - start) / fin));
    if (fout && Number.isFinite(end)) k = Math.min(k, Math.max(0, (end - t) / fout));
    return Math.max(0, Math.min(1, base * k));
  }
  /** Watches a media element once: the trimmed end, looping inside the trim, fades, rewinding, the bar. */
  function mediaWatch(el) {
    if (el.hsWatch) return;
    el.hsWatch = true;
    let raf = 0;
    const box = el.closest(".hs-audio");
    const tick = () => {
      raf = 0;
      const [start, end] = mediaSpan(el);
      if (Number.isFinite(end) && el.currentTime >= end - 0.03) {
        if (el.hasAttribute("data-loop")) el.currentTime = start;
        else { el.pause(); if (el.hasAttribute("data-rewind")) el.currentTime = start; }
      }
      el.volume = mediaVolume(el);
      if (box) {
        const span = Number.isFinite(end) ? end - start : 0;
        const at = Math.max(0, el.currentTime - start);
        const bar = box.querySelector(".hs-audio-track i");
        if (bar) bar.style.width = `${span ? Math.min(100, (at / span) * 100) : 0}%`;
        const time = box.querySelector(".hs-audio-time");
        if (time) time.textContent = `${Math.floor(at / 60)}:${String(Math.floor(at % 60)).padStart(2, "0")}`;
      }
      if (!el.paused) raf = requestAnimationFrame(tick);
    };
    el.addEventListener("play", () => { box?.classList.add("playing"); el.closest(".hs-hide-idle")?.classList.add("playing"); if (!raf) raf = requestAnimationFrame(tick); });
    el.addEventListener("pause", () => { box?.classList.remove("playing"); el.closest(".hs-hide-idle")?.classList.remove("playing"); tick(); });
    el.addEventListener("ended", () => { if (el.hasAttribute("data-rewind")) el.currentTime = mediaSpan(el)[0]; });
    el.addEventListener("seeked", tick);
  }
  /** Play a video or a sound from where it was (or the trimmed start), with its fades and volume. */
  function mediaPlay(el, { fromStart = false } = {}) {
    if (!el) return Promise.resolve();
    mediaWatch(el);
    const begin = () => {
      const [start, end] = mediaSpan(el);
      if (fromStart || el.currentTime < start || el.currentTime >= end - 0.05) el.currentTime = start;
      el.volume = mediaVolume(el);
    };
    if (el.readyState >= 1) begin(); else el.addEventListener("loadedmetadata", begin, { once: true });
    if (el.tagName === "VIDEO" && el.hasAttribute("data-fullscreen") && !document.fullscreenElement) el.requestFullscreen?.().catch(() => {});
    return el.play().catch(() => {
      // A browser that refuses sound without a click plays a video silently instead.
      if (el.tagName === "VIDEO") { el.muted = true; return el.play().catch(() => {}); }
      return null;
    });
  }
  function mediaPause(el, { stop = false } = {}) {
    if (!el) return;
    try { el.pause(); } catch { /* detached */ }
    if (stop || el.hasAttribute("data-rewind")) { try { el.currentTime = mediaSpan(el)[0]; } catch { /* not loaded */ } }
  }
  const mediaToggle = (el) => (el.paused ? mediaPlay(el) : mediaPause(el));

  function imageBody(o, rotEl, ctx) {
    const url = E.resolveSrc(o.src, ctx);
    const frame = h("div", { class: "hs-obj-img" });
    if (o.soft) softEdges(frame, o.soft);
    if (o.mask) {
      const g = geometry(o.mask, o.w, o.h, o.adj);
      frame.style.clipPath = `path(${g.rule === "evenodd" ? "evenodd, " : ""}'${g.paths.join(" ")}')`;
    }
    if (url) {
      const img = h("img", { src: url, alt: o.decorative ? "" : o.alt || "", draggable: "false", decoding: "async" });
      if (o.crop) {
        const { l, t, r, b } = o.crop;
        Object.assign(img.style, { position: "absolute", width: `${(100 / (1 - l - r)).toFixed(4)}%`, height: `${(100 / (1 - t - b)).toFixed(4)}%`, left: `${((-l / (1 - l - r)) * 100).toFixed(4)}%`, top: `${((-t / (1 - t - b)) * 100).toFixed(4)}%`, "max-width": "none" });
      } else img.style.objectFit = o.fit === "cover" ? "cover" : o.fit === "contain" ? "contain" : "fill";
      const filters = [o.bright ? `brightness(${r2(1 + o.bright)})` : "", o.contrast ? `contrast(${r2(1 + o.contrast)})` : "", o.sat ? `saturate(${r2(1 + o.sat)})` : "", o.gray ? "grayscale(1)" : ""].filter(Boolean).join(" ");
      if (filters) img.style.filter = filters;
      frame.append(img);
    } else frame.append(h("span", { class: "hs-obj-missing" }, E.icon("image"), "画像がありません"));
    rotEl.append(frame);
    if (o.stroke !== "none") {
      const g = geometry(o.mask || "rect", o.w, o.h, o.adj);
      const svg = s("svg", { class: "hs-obj-geom", width: r2(o.w), height: r2(o.h), viewBox: `0 0 ${r2(Math.max(1, o.w))} ${r2(Math.max(1, o.h))}`, overflow: "visible", "aria-hidden": "true" });
      const cap = o.cap ? SVG_CAPS[o.cap] : null;
      const compound = o.cmpd ? compoundMask(svg, o.cmpd, o.strokeW, g.paths, cap, o.join || "miter") : null;
      for (const d of g.paths) svg.append(s("path", { d, fill: "none", stroke: o.stroke, "stroke-width": o.strokeW, "stroke-dasharray": dashArray(o.dash, o.strokeW), "stroke-linecap": cap, "stroke-linejoin": o.join || null, ...(compound ? { mask: compound } : {}) }));
      rotEl.append(svg);
    }
  }

  function lineNode(o, all) {
    const route = linePath(o, all);
    const sw = o.stroke === "none" ? 0 : o.strokeW;
    const xs = route.pts.map((p) => p[0]);
    const ys = route.pts.map((p) => p[1]);
    const [minX, minY] = [Math.min(...xs), Math.min(...ys)];
    const [w, hh] = [Math.max(1, Math.max(...xs) - minX), Math.max(1, Math.max(...ys) - minY)];
    // A line says nothing to a screen reader unless it has alternative text.
    const el = h("div", { class: "hs-obj hs-obj-line", "data-el": o.id, "data-kind": "line", "data-stroke": sw ? o.stroke : "none", "data-bbox": [minX, minY, w, hh].map(r2).join(","), "aria-hidden": o.decorative || !o.alt ? "true" : null, role: !o.decorative && o.alt ? "img" : null, "aria-label": !o.decorative && o.alt ? o.alt : null, style: { left: `${r2(minX)}px`, top: `${r2(minY)}px`, width: `${r2(w)}px`, height: `${r2(hh)}px` } });
    const svg = s("svg", { class: "hs-obj-geom", width: r2(w), height: r2(hh), viewBox: `${r2(minX)} ${r2(minY)} ${r2(w)} ${r2(hh)}`, overflow: "visible", "aria-hidden": "true" });
    const [p1, p2] = [route.pts[0], route.pts[route.pts.length - 1]];
    // Pull the line in under filled heads so a thick line does not poke through their tips.
    const heads = [];
    const trim = (tip, toward, kind, size) => {
      const head = kind ? arrowhead(kind, tip, toward, sw || 2, size) : null;
      if (!head) return tip;
      heads.push(head);
      const len = Math.hypot(tip[0] - toward[0], tip[1] - toward[1]) || 1;
      const back = Math.min(head.back, len * 0.9);
      return [tip[0] - ((tip[0] - toward[0]) / len) * back, tip[1] - ((tip[1] - toward[1]) / len) * back];
    };
    let d = route.d;
    if (sw && (o.head || o.tail)) {
      const a = trim(p1, route.start, o.head, o.headSize ?? 2);
      const b = trim(p2, route.end, o.tail, o.tailSize ?? 2);
      if (route.ctrl) d = `M${P(...a)} C${P(...route.ctrl[0])} ${P(...route.ctrl[1])} ${P(...b)}`;
      else { const pts = [a, ...route.pts.slice(1, -1), b]; d = `M${pts.map((p) => P(...p)).join(" L")}`; }
    }
    if (sw) {
      const cap = o.cap ? SVG_CAPS[o.cap] : o.dash === "roundDot" ? "round" : "butt";
      const join = o.join || "round";
      const compound = o.cmpd ? compoundMask(svg, o.cmpd, sw, [d], cap, join) : null;
      svg.append(s("path", { d, fill: "none", stroke: o.stroke, "stroke-width": sw, "stroke-dasharray": dashArray(o.dash, sw), "stroke-linecap": cap, "stroke-linejoin": join, ...(compound ? { mask: compound } : {}) }));
      for (const head of heads) svg.append(s("path", { d: head.d, fill: head.fill ? o.stroke : "none", stroke: o.stroke, "stroke-width": head.fill ? Math.max(1, sw * 0.4) : sw, "stroke-linejoin": "miter", "stroke-linecap": "round" }));
    }
    const fx = h("div", { class: "hs-obj-fx" });
    const move = h("div", { class: "hs-obj-move" });
    fx.append(svg);
    move.append(fx);
    el.append(move);
    if (o.opacity != null) el.style.opacity = String(o.opacity);
    return el;
  }

  /** 3-D 回転: the CSS that turns an object about its x and y axes, with the camera's field of view as the perspective. */
  function rot3dCss(o) {
    if (!o.rot3d) return "";
    const { x, y, p = 45 } = o.rot3d;
    const persp = p > 0 ? `perspective(${r2(Math.max(o.w, o.h) / 2 / Math.tan((p * Math.PI) / 360))}px) ` : "";
    return `${persp}rotateX(${x}deg) rotateY(${y}deg)`;
  }

  /** One object as HTML: a positioned box → a layer for motion paths → a layer for effects → a rotated body. */
  function objectNode(raw, ctx, all, scale) {
    const o = withDefaults(raw);
    if (o.kind === "line") return lineNode(o, all);
    const el = h("div", {
      class: ["hs-obj", `hs-obj-${o.kind}`], "data-el": o.id, "data-kind": o.kind,
      "data-fill": ["shape", "text"].includes(o.kind) && o.fill !== "none" && !(o.shape === "custom" ? !o.path?.closed : SHAPES[o.shape]?.open) ? o.fill : null,
      "data-stroke": ["shape", "text", "image"].includes(o.kind) && o.stroke !== "none" ? o.stroke : null,
      "data-shadow": o.shadow && ["shape", "text", "image"].includes(o.kind) ? "1" : null,
      "data-tshadow": o.tshadow && ["shape", "text"].includes(o.kind) ? "1" : null,
      "data-toutline": o.toutline && ["shape", "text"].includes(o.kind) ? o.toutline.color : null,
      "data-pattern": o.pattern && ["shape", "text"].includes(o.kind) ? `${o.pattern.fg},${o.pattern.bg}` : null,
      "data-autofit": o.autofit && o.autofit !== "none" ? o.autofit : null, "data-item": o.item || null,
      "data-bbox": Object.values(bounds(o)).map(r2).join(","),
      // Screen readers: a decorative object is skipped; a picture-like object says its alternative text.
      "aria-hidden": o.decorative ? "true" : null,
      role: !o.decorative && o.alt && !["shape", "text", "table", "image"].includes(o.kind) ? "img" : null,
      "aria-label": !o.decorative && o.alt && o.kind !== "image" ? o.alt : null,
      style: { left: `${r2(o.x)}px`, top: `${r2(o.y)}px`, width: `${r2(o.w)}px`, height: `${r2(o.h)}px` },
    });
    const move = h("div", { class: "hs-obj-move" });
    const fx = h("div", { class: "hs-obj-fx" });
    const rot = h("div", { class: "hs-obj-rot" });
    if (o.kind === "image") applyEffects(o, rot);
    const transform = [rot3dCss(o), o.rot ? `rotate(${o.rot}deg)` : "", o.flipH || o.flipV ? `scale(${o.flipH ? -1 : 1}, ${o.flipV ? -1 : 1})` : ""].filter(Boolean).join(" ");
    if (transform) rot.style.transform = transform;
    if (o.kind === "shape" || o.kind === "text") {
      shapeBody(o, rot, scale, ctx);
      // A flipped shape keeps its words readable: the text is flipped back.
      const text = rot.querySelector(".hs-obj-text");
      if (text && (o.flipH || o.flipV)) text.style.transform = `scale(${o.flipH ? -1 : 1}, ${o.flipV ? -1 : 1})`;
    } else if (o.kind === "image") imageBody(o, rot, ctx);
    else if (o.kind === "table") tableBody(o, rot);
    else if (o.kind === "smartart") smartartBody(o, rot, scale);
    else if (o.kind === "ink") inkBody(o, rot);
    else if (o.kind === "zoom") zoomBody(o, rot, ctx);
    else if (o.kind === "camera") cameraBody(o, rot, ctx);
    else if (o.kind === "model") modelBody(o, rot, ctx);
    else if (o.kind === "poll") pollBody(o, rot, ctx);
    else if (o.kind === "equation") equationBody(o, rot);
    else if (o.kind === "chart") chartBody(o, rot);
    else if (o.kind === "icon") {
      const icon = E.icon(o.icon, "hs-obj-icon");
      if (icon) { icon.style.color = o.color; icon.setAttribute("stroke-width", String(o.strokeW)); rot.append(icon); }
    } else if (o.kind === "video" || o.kind === "lottie") {
      const yt = E.youtubeId(o.src);
      const desc = { kind: yt ? "youtube" : o.kind, src: o.src, yt, motion: "none", fit: o.fit === "contain" ? "contain" : "cover", autoplay: o.autoplay !== false, loop: o.loop !== false, muted: o.muted !== false, name: o.fileName || "" };
      if (o.kind === "lottie") desc.fit = o.fit === "cover" ? "cover" : "contain";
      const media = E.mediaEl(desc, ctx, "hs-obj-media");
      const video = media.querySelector("video");
      if (video) playbackData(video, o);
      if (video && ctx.live && o.bookmarks?.length) { media.append(h("div", { class: "hs-media-marks" }, bookmarkMarks(o))); watchMarks(video); }
      if (video && o.poster) video.setAttribute("poster", o.poster);
      if (!video && o.poster && !yt) media.prepend(h("img", { class: "hs-video-poster", src: o.poster, alt: "", draggable: "false" }));
      if (video && o.captions) {
        video.setAttribute("crossorigin", "anonymous");
        video.append(h("track", { kind: "captions", src: `data:text/vtt;charset=utf-8,${encodeURIComponent(o.captions)}`, srclang: o.captionLang || "ja", label: "字幕", default: true }));
      }
      if (video && ctx.live && o.hideIdle) media.classList.add("hs-hide-idle");
      rot.append(media);
    } else if (o.kind === "audio") audioBody(o, rot, ctx);
    if (o.action?.type === "flip") {
      // A card that turns over on a click: the front and the back share one 3D turn (the back is its mirror).
      const flip = h("div", { class: "hs-obj-flip" });
      const back = h("div", { class: "hs-obj-rot hs-obj-back" });
      back.style.transform = ["rotateY(180deg)", transform].filter(Boolean).join(" ");
      const backFill = o.action.fill || (o.fill && o.fill !== "none" ? o.fill : themeColors(PALETTE.sheet?.themeId).bg2);
      shapeBody(withDefaults({ ...o, kind: "shape", shape: o.kind === "shape" ? o.shape : "roundRect", text: o.action.back, fill: backFill, flipH: false, flipV: false, autofit: "none", valign: "middle", align: o.kind === "shape" || o.kind === "text" ? o.align : "center" }), back, scale);
      flip.append(rot, back);
      fx.append(flip);
    } else fx.append(rot);
    move.append(fx);
    el.append(move);
    if (o.opacity != null) el.style.opacity = String(o.opacity);
    fillFields(el, ctx);
    if (ctx.live) {
      if (o.action) { el.dataset.action = o.action.type; el.title = o.action.type === "url" ? o.action.href : ""; }
      if (o.overAction) el.dataset.over = o.overAction.type;
      if (o.action?.type === "popup") el.dataset.detail = `obj:${o.id}`;
      if (o.hover) el.dataset.hover = o.hover;
      if (o.loop) el.dataset.loop = o.loop;
      if (o.tip) el.dataset.tip = o.tip;
      if (o.group) el.dataset.group = o.group;
      // A looping group turns about the group's centre, not each member's own.
      if ((o.loop || o.hover) && ctx.groupBoxes?.[o.group]) {
        const g = ctx.groupBoxes[o.group];
        el.style.setProperty("--ix-ox", `${r2(g.x + g.w / 2 - o.x)}px`);
        el.style.setProperty("--ix-oy", `${r2(g.y + g.h / 2 - o.y)}px`);
      }
    }
    return el;
  }

  // 背景の書式設定: a slide's own background — a colour, a picture (with its transparency, filling the slide or tiled),
  // or both. It lies under the layout and its objects.
  function normalizeBackground(raw) {
    if (!raw || typeof raw !== "object") return null;
    const out = {};
    const color = hexColor(raw.color);
    if (color) out.color = color.toLowerCase();
    const src = typeof raw.image === "string" ? raw.image : "";
    if (src && /^(data:image\/(png|jpeg|webp|gif|svg\+xml)[;,]|idb:|asset:|https?:\/\/|blob:)/i.test(src) && src.length <= 80_000_000) {
      out.image = src;
      const t = Number(raw.transparency);
      if (Number.isFinite(t) && t > 0) out.transparency = Math.min(1, Math.round(t * 100) / 100);
      if (raw.tile === true) out.tile = true;
    }
    return Object.keys(out).length ? out : null;
  }
  function backgroundLayer(slide, ctx) {
    const bg = normalizeBackground(slide?.background);
    if (!bg) return null;
    const layer = h("div", { class: "hs-bg", "aria-hidden": "true", style: { background: bg.color || null } });
    const url = bg.image ? E.resolveSrc(bg.image, ctx) : "";
    if (url) {
      layer.append(h("div", { class: ["hs-bg-img", bg.tile ? "is-tiled" : ""], style: { "background-image": `url("${url.replace(/["\\\n\r]/g, (c) => encodeURIComponent(c))}")`, opacity: bg.transparency ? String(r2(1 - bg.transparency)) : null } }));
    }
    return layer;
  }

  /**
   * スライド マスター: the objects every slide shows (deck.masterObjects), under the layout and the slide's own
   * objects. A slide can leave them out (slide.hideMaster).
   * They are drawn, not interacted with: a screen reader skips them and a click goes through.
   */
  function masterLayer(deck, slide, ctx) {
    const list = Array.isArray(deck?.masterObjects) ? deck.masterObjects : [];
    if (!list.length || slide?.hideMaster === true) return null;
    const layer = h("div", { class: "hs-master-layer", "aria-hidden": "true" });
    for (const o of list) {
      if (!o || o.hidden || !KINDS.includes(o.kind)) continue;
      try {
        const node = objectNode(o, { ...ctx, live: false }, list, undefined);
        if (node) layer.append(node);
      } catch { /* a broken object never takes the slide down */ }
    }
    return layer.childNodes.length ? layer : null;
  }

  /** The slide's objects as one layer above the layout. */
  function objectLayer(slide, ctx, opts = {}) {
    const list = Array.isArray(slide?.elements) ? slide.elements : [];
    if (!list.length) return null;
    const layer = h("div", { class: "hs-objects" });
    const scales = opts.fit?.objs || {};
    // While presenting, what a reveal button shows waits hidden until it is clicked.
    const waiting = new Set(ctx.live ? list.flatMap((o) => (o?.action?.type === "reveal" && !o.hidden ? o.action.targets : [])) : []);
    const groupBoxes = {};
    if (ctx.live) {
      for (const o of list) {
        if (!o?.group || !(o.loop || o.hover) || groupBoxes[o.group]) continue;
        const boxes = list.filter((m) => m?.group === o.group && m.kind !== "line").map((m) => bounds(m));
        const x = Math.min(...boxes.map((b) => b.x));
        const y = Math.min(...boxes.map((b) => b.y));
        groupBoxes[o.group] = { x, y, w: Math.max(...boxes.map((b) => b.x + b.w)) - x, h: Math.max(...boxes.map((b) => b.y + b.h)) - y };
      }
    }
    const live = { ...ctx, groupBoxes };
    // 読み取り順序 (what a screen reader reads first) may differ from the stacking order: the nodes then follow the
    // reading order and each keeps its place in the stack with a z-index.
    const reading = Array.isArray(slide.readingOrder) && slide.readingOrder.length ? readingOrderOf(slide) : null;
    const nodes = [];
    list.forEach((o, stack) => {
      if (!o || o.hidden || !KINDS.includes(o.kind)) return;
      try {
        const node = objectNode(o, live, list, scales[o.id]);
        if (node && waiting.has(o.id)) node.classList.add("hs-ix-wait");
        if (node && reading) node.style.zIndex = String(stack + 1);
        if (node) nodes.push([o.id, node]);
      } catch { /* a broken object never takes the slide down */ }
    });
    if (reading) { const at = new Map(reading.map((id, i) => [id, i])); nodes.sort((a, b) => (at.get(a[0]) ?? 1e6) - (at.get(b[0]) ?? 1e6)); }
    for (const [, node] of nodes) layer.append(node);
    return layer;
  }

  /**
   * The slide's objects in reading order: the ids of slide.readingOrder that still exist, then the others in the
   * stacking order (the order of slide.elements, which is also the reading order when none is set).
   */
  function readingOrderOf(slide) {
    const ids = (Array.isArray(slide?.elements) ? slide.elements : []).map((o) => o?.id).filter(Boolean);
    const known = new Set(ids);
    const first = [...new Set((Array.isArray(slide?.readingOrder) ? slide.readingOrder : []).filter((id) => known.has(id)))];
    const picked = new Set(first);
    return [...first, ...ids.filter((id) => !picked.has(id))];
  }

  /** The cards a slide's objects open on a click (「詳細を開く」), as the player's details: [{ target: "obj:<id>", … }]. */
  function objectDetails(slide) {
    return (Array.isArray(slide?.elements) ? slide.elements : []).filter((o) => o?.action?.type === "popup" && !o.hidden)
      .map((o) => ({ target: `obj:${o.id}`, title: o.action.title || "", text: o.action.text || "", html: true, rows: o.action.rows || [], source: o.action.source || "" }));
  }

  /** Shrink the text of objects set to "shrink on overflow" until it fits (the slide must be in the document). */
  function fitObjects(slideEl) {
    const out = {};
    for (const el of slideEl.querySelectorAll('.hs-obj[data-autofit="shrink"]')) {
      const box = el.querySelector(".hs-obj-text");
      const tx = box?.querySelector(".hs-obj-tx");
      if (!box || !tx) continue;
      const over = () => tx.scrollHeight > box.clientHeight + 1 || tx.scrollWidth > box.clientWidth + 1;
      let k = 1;
      tx.style.setProperty("--os", "1");
      while (over() && k > 0.3) {
        k = Math.round((k - 0.04) * 100) / 100;
        tx.style.setProperty("--os", String(k));
      }
      if (k < 1) out[el.dataset.el] = k;
    }
    return Object.keys(out).length ? out : null;
  }

  /** What an object says, as plain text (search, the AI's summary, the selection pane). */
  function objectText(o) {
    if (o?.kind === "table") return o.cells.map((row) => row.filter((cell) => !cell.merged).map((cell) => (cell.text ? richToText(cell.text).trim() : "")).join("\t")).join("\n").trim();
    if (o?.kind === "chart") return [o.chart.title, ...o.chart.labels].filter(Boolean).join(" ");
    return o && (o.kind === "shape" || o.kind === "text") && o.text ? richToText(o.text).trim() : "";
  }

  /** The default name PowerPoint would give ("正方形/長方形 3", "テキスト ボックス 2"). */
  function objectName(o, index = 0) {
    if (o.name) return o.name;
    const base = o.kind === "shape" ? (o.shape === "custom" ? "フリーフォーム" : (SHAPES[o.shape]?.label || "図形").replace(/^.*: /, "")) : o.kind === "line" ? (o.head || o.tail ? "矢印" : o.route === "elbow" ? "カギ線コネクタ" : o.route === "curve" ? "曲線コネクタ" : "直線") : KIND_LABELS[o.kind] || "オブジェクト";
    return `${base} ${index + 1}`;
  }

  // ---------------------------------------------------------------- 変形 (Morph)

  // What can be the same thing on two slides: not what plays (videos, sounds, animations, 3D models, cameras, polls) or points elsewhere.
  const MORPH_KINDS = new Set(["shape", "text", "image", "line", "icon", "table", "chart", "smartart", "ink", "equation"]);
  const morphable = (o) => Boolean(o && !o.hidden && MORPH_KINDS.has(o.kind) && o.id);
  /** A name that starts with "!!" says "this is the same as the one with this name on the next slide" (as in PowerPoint). */
  const isMorphName = (name) => typeof name === "string" && name.startsWith("!!") && name.length > 2;
  /** What makes two parts the same when they were not named: the same words, the same picture, the same icon. */
  function morphLook(o) {
    if (o.kind === "image" && o.src) return `image:${o.src}`;
    if (o.kind === "icon" && o.icon) return `icon:${o.icon}`;
    if ((o.kind === "shape" || o.kind === "text") && o.text) {
      const words = richToText(o.text).replace(/\s+/g, " ").trim();
      if (words) return `${o.kind}:${words}`;
    }
    return "";
  }
  /**
   * 画面切り替え「変形」: the objects of the slide being left that are the same as objects of the slide coming in, as
   * [from, to] pairs. In this order: the same "!!" name; the same id and kind (a duplicated slide keeps its ids, so the
   * parts moved on the copy are the same parts); the same words, picture or icon where that is only one on each side.
   */
  function morphPairs(fromList, toList) {
    const from = (Array.isArray(fromList) ? fromList : []).filter(morphable);
    const to = (Array.isArray(toList) ? toList : []).filter(morphable);
    const pairs = [];
    const usedFrom = new Set();
    const usedTo = new Set();
    const take = (a, b) => { pairs.push([a, b]); usedFrom.add(a); usedTo.add(b); };
    for (const b of to) {
      if (!isMorphName(b.name)) continue;
      const a = from.find((m) => !usedFrom.has(m) && m.name === b.name);
      if (a) take(a, b);
    }
    for (const b of to) {
      if (usedTo.has(b)) continue;
      const a = from.find((m) => !usedFrom.has(m) && m.id === b.id && m.kind === b.kind && !isMorphName(m.name));
      if (a) take(a, b);
    }
    const looks = (list, used) => { const map = new Map(); for (const o of list) { if (used.has(o)) continue; const look = morphLook(o); if (look) map.set(look, [...(map.get(look) || []), o]); } return map; };
    const fromLooks = looks(from, usedFrom);
    const toLooks = looks(to, usedTo);
    for (const [look, bs] of toLooks) {
      const as = fromLooks.get(look);
      if (as?.length === 1 && bs.length === 1) take(as[0], bs[0]);
    }
    return pairs;
  }

  Object.assign(E, {
    PX_PER_PT, PX_PER_CM, PALETTE, useTheme, retheme, paletteSheet, themeColors, themeColorRef, lumColor, lumVariants, STANDARD_COLORS, HIGHLIGHTS, FONTS, SHAPES, SHAPE_GROUPS, LIST_STYLES, OBJECT_KINDS: KINDS, KIND_LABELS, DASHES, ARROWHEADS, ROUTES, AUTOFIT, FITS, OBJECT_DEFAULTS: DEFAULTS, IX_HOVERS, IX_LOOPS, IX_CLICKS,
    LINE_CAPS, LINE_JOINS, COMPOUNDS, LIST_MARKS, LIST_MARK_SIZES, U_LINES, CHART_COLORS, chartColorVars, tableStyleVars, CHART_LAYOUTS, chartLayoutOf, applyChartLayout, CHART_STYLES, chartStyleOf, applyChartStyle, visibleChart, TREND_CHART_KINDS, BAR_GAP_CHARTS, REVERSE_CHARTS, LINE_CHART_KINDS, STEP_CHARTS, MARKER_SHAPES, LABEL_ALL_CHARTS, LABEL_POS_CHARTS, SLICE_LABEL_PARTS, TABLE_CHARTS, TABLE_STYLES, TABLE_STYLE_SHORT, CHART_KINDS, CHART_MAX_LABELS, CHART_MAX_SERIES, chartSpec, officeChart, numFormat, freeformD, objectDetails, richNodes: (html) => richFragment(html),
    geometry, adjOf, sanitizeRich, richFragment, textToRich, richToText, PATTERN_KINDS, patternNode, parseTabs, tabsText, tabBoxes, noteHtml, noteToText, noteOf, noteFields, noteFormatted, hexColor, normalizeObject, normalizeObjects, withDefaults, newObjectId: newId,
    corners, bounds, sites, lineEnds, linePath, objectLayer, masterLayer, morphPairs, MORPH_KINDS, normalizeBackground, backgroundLayer, readingOrderOf, objectNode, fitObjects, objectText, objectName,
    VOLUMES, normalizePlayback, normalizeBookmarks, mediaMarks, placeMarks, mediaPlay, mediaPause, mediaToggle, mediaSpan,
    SMARTART_LAYOUTS, SMARTART_GROUPS, SMARTART_COLORS, SMARTART_STYLES, normalizeSmartart, smartartParts, smartartObjects, smartartSample,
    normalizeStrokes, inkPath, texToMathML,
  });
})(typeof window !== "undefined" ? window : globalThis);
