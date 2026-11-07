// 描画 (PowerPoint's Draw tab): write on a slide with a pen or a highlighter (mouse, touch or a stylus), erase
// strokes, lasso strokes to move them, turn ink into shapes (a line, a rectangle, an ellipse, a triangle…), and
// play the ink back as it was written. Ink is an object ({ kind: "ink", strokes: [{ pts, color, width,
// highlighter }] }, points as fractions of its box) drawn by objects.js; strokes written in a row join one object.

const PENS = [
  { id: "black", label: "ペン（黒）", color: "#000000", width: 6 },
  { id: "blue", label: "ペン（青）", color: "#0070c0", width: 6 },
  { id: "gray", label: "ペン（グレー）", color: "#808080", width: 6 },
  { id: "red", label: "ペン（赤・強調）", color: "#c00000", width: 6 },
  { id: "hlYellow", label: "蛍光ペン（黄）", color: "#ffff00", width: 28, highlighter: true },
  { id: "hlGreen", label: "蛍光ペン（明るい緑）", color: "#00ff00", width: 28, highlighter: true },
];
const WIDTHS = [[3, "極細"], [6, "細"], [10, "中"], [16, "太"], [28, "極太"]];
const JOIN_MS = 2500;

/** Fewer points where the line is straight (Ramer–Douglas–Peucker). */
export function simplify(pts, tolerance) {
  if (pts.length < 3) return pts;
  const keep = new Array(pts.length).fill(false);
  keep[0] = keep[pts.length - 1] = true;
  const stack = [[0, pts.length - 1]];
  const dist = (p, a, b) => {
    const [dx, dy] = [b[0] - a[0], b[1] - a[1]];
    const len = dx * dx + dy * dy;
    const t = len ? Math.max(0, Math.min(1, ((p[0] - a[0]) * dx + (p[1] - a[1]) * dy) / len)) : 0;
    return Math.hypot(p[0] - (a[0] + t * dx), p[1] - (a[1] + t * dy));
  };
  while (stack.length) {
    const [a, b] = stack.pop();
    let far = -1;
    let best = tolerance;
    for (let i = a + 1; i < b; i += 1) { const d = dist(pts[i], pts[a], pts[b]); if (d > best) { best = d; far = i; } }
    if (far >= 0) { keep[far] = true; stack.push([a, far], [far, b]); }
  }
  return pts.filter((_, i) => keep[i]);
}

/**
 * インクを図形に変換: what a stroke (slide pixels) looks like — a line, a triangle, a rectangle, an ellipse, another
 * polygon — or null when it is handwriting.
 */
export function recognize(pts) {
  if (pts.length < 2) return null;
  const xs = pts.map((p) => p[0]);
  const ys = pts.map((p) => p[1]);
  const box = { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
  const diag = Math.hypot(box.w, box.h);
  if (diag < 24) return null;
  const length = pts.reduce((sum, p, i) => (i ? sum + Math.hypot(p[0] - pts[i - 1][0], p[1] - pts[i - 1][1]) : 0), 0);
  const [a, b] = [pts[0], pts[pts.length - 1]];
  const ends = Math.hypot(b[0] - a[0], b[1] - a[1]);
  // Nearly straight from one end to the other: a line.
  if (length < ends * 1.12) return { kind: "line", x1: a[0], y1: a[1], x2: b[0], y2: b[1] };
  // A shape must close on itself.
  if (ends > Math.max(40, diag * 0.28) || length < diag * 1.8) return null;
  const cx = xs.reduce((s, v) => s + v, 0) / xs.length;
  const cy = ys.reduce((s, v) => s + v, 0) / ys.length;
  const ring = [...pts];
  const corners = simplify([...ring, ring[0]], diag * 0.09).slice(0, -1);
  // Round: the distance from the middle (scaled to the box) hardly changes.
  const radial = pts.map(([x, y]) => Math.hypot((x - cx) / Math.max(1, box.w / 2), (y - cy) / Math.max(1, box.h / 2)));
  const mean = radial.reduce((s, v) => s + v, 0) / radial.length;
  const spread = Math.sqrt(radial.reduce((s, v) => s + (v - mean) ** 2, 0) / radial.length) / (mean || 1);
  if (corners.length >= 5 && spread < 0.16) return { kind: "shape", shape: "ellipse", ...box };
  if (corners.length === 3) return { kind: "shape", shape: "custom", ...box, path: { pts: corners.map(([x, y]) => [(x - box.x) / Math.max(1, box.w), (y - box.y) / Math.max(1, box.h)]), closed: true } };
  if (corners.length === 4) {
    // Right angles and sides along the axes: a rectangle; otherwise the four-sided shape as drawn.
    const angle = (p, q, r) => { const v1 = [p[0] - q[0], p[1] - q[1]]; const v2 = [r[0] - q[0], r[1] - q[1]]; return Math.abs((Math.atan2(v1[0] * v2[1] - v1[1] * v2[0], v1[0] * v2[0] + v1[1] * v2[1]) * 180) / Math.PI); };
    const square = corners.every((q, i) => Math.abs(angle(corners[(i + 3) % 4], q, corners[(i + 1) % 4]) - 90) < 22);
    const level = corners.every((q, i) => { const r = corners[(i + 1) % 4]; const deg = (Math.atan2(r[1] - q[1], r[0] - q[0]) * 180) / Math.PI; return Math.min(...[0, 90, 180, -90, -180].map((t) => Math.abs(deg - t))) < 15; });
    if (square && level) return { kind: "shape", shape: "rect", ...box };
  }
  if (corners.length >= 4 && corners.length <= 8 && spread >= 0.16) return { kind: "shape", shape: "custom", ...box, path: { pts: corners.map(([x, y]) => [(x - box.x) / Math.max(1, box.w), (y - box.y) / Math.max(1, box.h)]), closed: true } };
  if (spread < 0.25) return { kind: "shape", shape: "ellipse", ...box };
  return null;
}

/** Strokes (slide pixels) as one ink object: its box and the points as fractions of it. */
export function inkObject(strokes, id) {
  const all = strokes.flatMap((st) => st.pts);
  const pad = Math.max(...strokes.map((st) => st.width)) / 2 + 1;
  const xs = all.map((p) => p[0]);
  const ys = all.map((p) => p[1]);
  const x = Math.min(...xs) - pad;
  const y = Math.min(...ys) - pad;
  const w = Math.max(8, Math.max(...xs) - Math.min(...xs) + pad * 2);
  const hh = Math.max(8, Math.max(...ys) - Math.min(...ys) + pad * 2);
  return {
    id, kind: "ink", x: Math.round(x * 100) / 100, y: Math.round(y * 100) / 100, w: Math.round(w * 100) / 100, h: Math.round(hh * 100) / 100,
    strokes: strokes.map((st) => ({ ...st, pts: st.pts.map(([px, py]) => [Math.round(((px - x) / w) * 10000) / 10000, Math.round(((py - y) / hh) * 10000) / 10000]) })),
  };
}
/** An ink object's strokes in slide pixels (its box may have been moved or resized). */
export function strokesOf(o) {
  return (o.strokes || []).map((st) => ({ ...st, pts: st.pts.map(([fx, fy]) => [o.x + fx * o.w, o.y + fy * o.h]) }));
}

export function createInk(editor, app, kit) {
  const { E, h } = app;
  const { btn, drop, group, col, menu, updater } = kit;
  let tool = null; // "pen" | "eraser" | "lasso"
  let pen = PENS[0];
  let width = null; // a width chosen for the pen (null: the pen's own)
  let toShapes = false;
  let layer = null;
  let lastInk = null; // { id, at } — strokes in a row join the same object

  function setTool(next) {
    tool = tool === next && next !== "pen" ? null : next;
    unmount();
    if (tool) mount();
    kit.refreshRibbon?.();
  }
  function choosePen(p) { pen = p; width = null; tool = "pen"; unmount(); mount(); kit.refreshRibbon?.(); }
  const stop = () => { tool = null; unmount(); kit.refreshRibbon?.(); };

  function unmount() {
    layer?.remove();
    layer = null;
    document.removeEventListener("keydown", onKey, true);
  }
  function onKey(event) {
    if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); stop(); }
  }
  /** A layer over the slide that catches the pen (pointer events: mouse, touch, stylus with pressure). */
  function mount() {
    const wrap = editor.state.wrap;
    if (!wrap || !app.onSlide()) { tool = null; return; }
    editor.clear();
    layer = h("div", { class: ["ed-pathdraw", "ed-ink-layer", `tool-${tool}`], title: tool === "pen" ? `${pen.label}：ドラッグで書く（Escで終わる）` : tool === "eraser" ? "消しゴム：消したい線をなぞる" : "投げ縄：囲んだ線を選ぶ" });
    const svg = E.s("svg", { viewBox: `0 0 ${E.W} ${E.H}`, preserveAspectRatio: "none" });
    layer.append(svg);
    wrap.append(layer);
    document.addEventListener("keydown", onKey, true);
    layer.addEventListener("pointerdown", (event) => {
      if (event.button !== 0 && event.pointerType === "mouse") return;
      // The canvas below must not start a selection (it would redraw the stage under the pen).
      event.preventDefault();
      event.stopPropagation();
      layer.setPointerCapture(event.pointerId);
      const pts = [editor.toSlide(event)];
      const w = width ?? pen.width;
      const live = E.s("path", { fill: "none", stroke: tool === "pen" ? pen.color : tool === "eraser" ? "#c00000" : "#2b6be0", "stroke-width": tool === "pen" ? w : 2, "stroke-linecap": "round", "stroke-linejoin": "round", "stroke-opacity": tool === "pen" && pen.highlighter ? 0.9 : 1, "stroke-dasharray": tool === "lasso" ? "8 6" : null });
      svg.append(live);
      const move = (ev) => {
        for (const e of ev.getCoalescedEvents?.() || [ev]) {
          const p = editor.toSlide(e);
          const last = pts[pts.length - 1];
          if (Math.hypot(p[0] - last[0], p[1] - last[1]) >= 1.5) pts.push(p);
        }
        live.setAttribute("d", E.inkPath(pts.map(([x, y]) => [x / E.W, y / E.H]), E.W, E.H));
        if (tool === "eraser") eraseAt(pts[pts.length - 1], { preview: true });
      };
      const up = () => {
        layer?.removeEventListener("pointermove", move);
        layer?.removeEventListener("pointerup", up);
        layer?.removeEventListener("pointercancel", up);
        live.remove();
        if (tool === "pen") finishStroke(pts, w);
        else if (tool === "eraser") eraseAlong(pts);
        else if (tool === "lasso") lassoSelect(pts);
      };
      layer.addEventListener("pointermove", move);
      layer.addEventListener("pointerup", up);
      layer.addEventListener("pointercancel", up);
    });
  }

  function finishStroke(raw, w) {
    const pts = simplify(raw, 0.6);
    const list = editor.objects();
    if (toShapes && !pen.highlighter) {
      const shape = recognize(pts);
      if (shape) {
        const o = shape.kind === "line"
          ? { id: E.newObjectId(), kind: "line", x1: shape.x1, y1: shape.y1, x2: shape.x2, y2: shape.y2, stroke: pen.color, strokeW: Math.max(2, w) }
          : { id: E.newObjectId(), kind: "shape", ...shape, fill: "none", stroke: pen.color, strokeW: Math.max(2, w) };
        editor.commit([...list, o], { select: editor.selection });
        lastInk = null;
        remount();
        return;
      }
    }
    const stroke = { pts, color: pen.color, width: w, ...(pen.highlighter ? { highlighter: true } : {}) };
    // Strokes in a row (within a few seconds, near each other) belong to one ink object, as PowerPoint groups them.
    const prev = lastInk && Date.now() - lastInk.at < JOIN_MS ? list.find((o) => o.id === lastInk.id && o.kind === "ink") : null;
    let next;
    let id;
    if (prev && near(prev, pts)) {
      id = prev.id;
      next = list.map((o) => (o.id === id ? { ...o, ...inkObject([...strokesOf(o), stroke], id) } : o));
    } else {
      id = E.newObjectId();
      next = [...list, inkObject([stroke], id)];
    }
    editor.commit(next, { select: [] });
    lastInk = { id, at: Date.now() };
    remount();
  }
  const near = (o, pts) => pts.some(([x, y]) => x > o.x - 240 && x < o.x + o.w + 240 && y > o.y - 240 && y < o.y + o.h + 240);
  // The stage redraws after a change: the pen layer goes back on top of the new slide.
  // At once, not on the next frame: until the layer is back, a stroke would land on the slide as a selection.
  function remount() { if (!tool) return; unmount(); mount(); }

  // ---------------------------------------------------------------- 消しゴム・投げ縄

  const R = 14;
  function hitStroke(st, p, r = R) {
    return st.pts.some((q, i) => {
      if (Math.hypot(q[0] - p[0], q[1] - p[1]) <= r + st.width / 2) return true;
      const n = st.pts[i + 1];
      if (!n) return false;
      const [dx, dy] = [n[0] - q[0], n[1] - q[1]];
      const len = dx * dx + dy * dy;
      const t = len ? Math.max(0, Math.min(1, ((p[0] - q[0]) * dx + (p[1] - q[1]) * dy) / len)) : 0;
      return Math.hypot(p[0] - (q[0] + t * dx), p[1] - (q[1] + t * dy)) <= r + st.width / 2;
    });
  }
  function eraseAt(p, { preview = false } = {}) {
    if (!preview || !layer) return;
    for (const o of editor.objects()) {
      if (o.kind !== "ink") continue;
      const node = document.querySelector(`.slide-wrap .hs-obj[data-el="${CSS.escape(o.id)}"]`);
      strokesOf(o).forEach((st, i) => { if (hitStroke(st, p)) node?.querySelectorAll(".hs-ink path")[i]?.setAttribute("opacity", "0.2"); });
    }
  }
  /** Every stroke the eraser passed over goes (an ink object with nothing left goes too): one undo step. */
  function eraseAlong(pts) {
    let changed = false;
    const next = [];
    for (const o of editor.objects()) {
      if (o.kind !== "ink" || o.locked) { next.push(o); continue; }
      const strokes = strokesOf(o);
      const kept = strokes.filter((st) => !pts.some((p) => hitStroke(st, p)));
      if (kept.length === strokes.length) { next.push(o); continue; }
      changed = true;
      if (kept.length) next.push({ ...o, ...inkObject(kept, o.id) });
    }
    if (changed) editor.commit(next, { select: [] });
    else editor.draw();
    remount();
  }
  const inside = (p, poly) => {
    let hit = false;
    for (let i = 0, j = poly.length - 1; i < poly.length; j = i, i += 1) {
      const [xi, yi] = poly[i];
      const [xj, yj] = poly[j];
      if ((yi > p[1]) !== (yj > p[1]) && p[0] < ((xj - xi) * (p[1] - yi)) / (yj - yi || 1e-9) + xi) hit = !hit;
    }
    return hit;
  };
  /** 投げ縄選択: the strokes inside the loop become one ink object of their own, selected (to move, copy or delete). */
  function lassoSelect(poly) {
    if (poly.length < 3) return;
    const taken = [];
    const next = [];
    for (const o of editor.objects()) {
      if (o.kind !== "ink" || o.locked) { next.push(o); continue; }
      const strokes = strokesOf(o);
      const caught = strokes.filter((st) => st.pts.filter((p) => inside(p, poly)).length >= st.pts.length * 0.6);
      if (!caught.length) { next.push(o); continue; }
      taken.push(...caught);
      const rest = strokes.filter((st) => !caught.includes(st));
      if (rest.length) next.push({ ...o, ...inkObject(rest, o.id) });
    }
    if (!taken.length) { app.toast("囲んだ中にインクがありません"); remount(); return; }
    const picked = inkObject(taken, E.newObjectId());
    tool = null;
    unmount();
    editor.commit([...next, picked], { select: [picked.id] });
    kit.refreshRibbon?.();
    app.toast("囲んだインクを選びました（ドラッグで移動・Deleteで削除）");
  }

  // ---------------------------------------------------------------- インクを図形に変換・描画で再生

  /** The selected ink, or all ink on the slide, recognised stroke by stroke (what is handwriting stays ink). */
  function convertSelected() {
    const list = editor.objects();
    const sel = new Set(editor.selection);
    const targets = list.filter((o) => o.kind === "ink" && (!sel.size || sel.has(o.id)));
    if (!targets.length) { app.toast("インクがありません"); return; }
    let made = 0;
    const next = [];
    for (const o of list) {
      if (!targets.includes(o)) { next.push(o); continue; }
      const keep = [];
      for (const st of strokesOf(o)) {
        const shape = st.highlighter ? null : recognize(st.pts);
        if (!shape) { keep.push(st); continue; }
        made += 1;
        next.push(shape.kind === "line"
          ? { id: E.newObjectId(), kind: "line", x1: shape.x1, y1: shape.y1, x2: shape.x2, y2: shape.y2, stroke: st.color, strokeW: Math.max(2, st.width) }
          : { id: E.newObjectId(), kind: "shape", ...shape, fill: "none", stroke: st.color, strokeW: Math.max(2, st.width) });
      }
      if (keep.length) next.push({ ...o, ...inkObject(keep, o.id) });
    }
    if (!made) { app.toast("図形にできる線がありませんでした（手書きの文字はインクのままです）"); return; }
    editor.commit(next, { select: [] });
    app.toast(`${made}本の線を図形にしました`);
  }
  /** 描画で再生: the ink on this slide is written again, stroke by stroke (the same "線を描く" effect the show uses). */
  function replay() {
    const ink = editor.objects().filter((o) => o.kind === "ink" && !o.hidden);
    if (!ink.length) { app.toast("このスライドにインクがありません"); return; }
    app.previewEffect?.(ink.map((o) => ({ id: `rp${o.id}`, el: o.id, cls: "in", fx: "draw", start: "with", dur: Math.min(8000, 900 + strokesOf(o).length * 450), delay: 0 })));
  }
  /** Add 描画で再生 to the slide's animations (the ink is written in the show, click by click). */
  function addReplayAnimation() {
    const ink = editor.objects().filter((o) => o.kind === "ink" && !o.hidden);
    if (!ink.length) { app.toast("このスライドにインクがありません"); return; }
    const timeline = [...(app.slide()?.timeline || []).filter((e) => !(e.fx === "draw" && ink.some((o) => o.id === e.el)))];
    ink.forEach((o, i) => timeline.push({ id: `a${Math.random().toString(36).slice(2, 9)}`, el: o.id, cls: "in", fx: "draw", start: i ? "after" : "click", dur: Math.min(8000, 900 + strokesOf(o).length * 450), delay: 0 }));
    app.setTimeline(timeline);
    app.toast("発表でインクが書かれていくように、アニメーション「線を描く」を付けました");
  }

  // ---------------------------------------------------------------- the 描画 tab

  function drawTab() {
    const penBtn = (p) => {
      const b = h("button", { type: "button", class: ["ink-pen", p.highlighter ? "hl" : ""], title: p.label, "aria-label": p.label, onclick: () => choosePen(p) }, h("i", { style: { background: p.color } }));
      updater(() => b.classList.toggle("on", tool === "pen" && pen.id === p.id));
      return b;
    };
    return [
      group("描画ツール",
        btn("selectAll", "選択", "書くのをやめて図形を選ぶ（Esc）", () => stop(), { big: true, pressed: () => !tool }),
        btn("lasso", "投げ縄|選択", "インクを囲んで選ぶ", () => setTool("lasso"), { big: true, pressed: () => tool === "lasso" }),
        btn("eraser", "消しゴム", "なぞった線を消す", () => setTool("eraser"), { big: true, pressed: () => tool === "eraser" }),
        h("div", { class: "ink-pens" }, PENS.map(penBtn)),
        drop("weight", "太さ", "ペンの太さ", () => menu(WIDTHS.map(([v, label]) => ({ label: `${label}（${v}px）`, on: (width ?? pen.width) === v, run: () => { width = v; if (tool !== "pen") choosePen(pen); kit.refreshRibbon?.(); } }))), {})),
      group("変換",
        btn("inkShape", "インクを|図形に変換", "オンにすると書いた線が図形（直線・四角形・楕円・三角形…）になります。選んだインクは今すぐ変換", () => { if (editor.selectedObjects().some((o) => o.kind === "ink")) convertSelected(); else { toShapes = !toShapes; kit.refreshRibbon?.(); app.toast(toShapes ? "インクを図形に変換：オン（書いた線が図形になります）" : "インクを図形に変換：オフ"); } }, { big: true, pressed: () => toShapes })),
      group("再生",
        btn("play", "描画で|再生", "このスライドのインクを書いた順に再生して見せる", () => replay(), { big: true }),
        btn("motionPath", "発表で|再生", "発表でインクが書かれていくアニメーションを付ける", () => addReplayAnimation(), { big: true })),
    ];
  }

  // The stage may be redrawn (another slide, zoom): keep the pen layer on the slide that is shown.
  editor.subscribe(() => { if (tool && layer && !layer.isConnected) remount(); });

  return { drawTab, stop, get tool() { return tool; }, recognize, convertSelected, replay };
}
