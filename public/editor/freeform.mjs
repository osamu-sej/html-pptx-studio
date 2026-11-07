// Shapes drawn by hand, as PowerPoint's 曲線・フリーフォーム: 図形・フリーフォーム: フリーハンド, and 頂点の編集.
// 曲線: click to put points, the line runs smoothly through them; フリーフォーム: click for corners, press and drag
// to draw freely; フリーハンド: press, draw, let go. Ending near the first point closes the shape (it gets a fill);
// a double-click or Enter ends it open (a line). A drawn shape keeps its points as fractions of its box.

import * as ops from "./ops.mjs";

const CLOSE_PX = 14;

export function createFreeform(editor, app) {
  const { E, h } = app;
  let drawing = null;
  let editing = null;

  // ---------------------------------------------------------------- drawing

  function start(mode) {
    stop();
    const wrap = editor.state.wrap;
    if (!wrap) return;
    editor.clear();
    const pts = [];
    let hover = null;
    let pressed = false;
    let freehand = mode === "scribble";
    const layer = h("div", { class: "ed-pathdraw ed-freeform", title: mode === "scribble" ? "押したままなぞって描き、離すと終わります" : "クリックで点を打つ・最初の点で閉じる・ダブルクリックかEnterで終わる・Escでやめる" });
    const svg = E.s("svg", { viewBox: `0 0 ${E.W} ${E.H}`, preserveAspectRatio: "none" });
    layer.append(svg);
    wrap.append(layer);
    const near = (p, q) => Math.hypot(p[0] - q[0], p[1] - q[1]) * editor.scale() < CLOSE_PX;
    const paint = () => {
      const all = hover && !pressed ? [...pts, hover] : pts;
      if (!all.length) { svg.replaceChildren(); return; }
      const path = { pts: all, curve: mode === "curve" };
      svg.replaceChildren(E.s("path", { d: E.freeformD(path, 1, 1), fill: "none", stroke: "#1f3864", "stroke-width": 4, "vector-effect": "non-scaling-stroke" }),
        ...(pts.length > 2 && hover && near(hover, pts[0]) ? [E.s("circle", { cx: pts[0][0], cy: pts[0][1], r: 12, fill: "none", stroke: "#16a34a", "stroke-width": 4 })] : []));
    };
    const finish = (closed = false) => {
      const points = mode === "scribble" || freehand ? simplify(pts) : pts;
      stop();
      if (points.length < 2) { app.toast("2つ以上の点で描いてください"); return; }
      make(points, { closed, curve: mode === "curve" || mode === "scribble" || freehand });
    };
    layer.addEventListener("pointerdown", (ev) => {
      ev.preventDefault();
      ev.stopPropagation();
      const p = editor.toSlide(ev);
      layer.setPointerCapture(ev.pointerId);
      pressed = true;
      if (pts.length > 2 && near(p, pts[0]) && mode !== "scribble") { finish(true); return; }
      const last = pts[pts.length - 1];
      if (!last || Math.hypot(p[0] - last[0], p[1] - last[1]) > 3) pts.push(p);
      paint();
    });
    layer.addEventListener("pointermove", (ev) => {
      const p = editor.toSlide(ev);
      if (pressed && mode !== "curve") {
        const last = pts[pts.length - 1];
        if (last && Math.hypot(p[0] - last[0], p[1] - last[1]) * editor.scale() > 5) { pts.push(p); if (mode === "polygon") freehand = true; }
      } else hover = p;
      paint();
    });
    layer.addEventListener("pointerup", () => {
      pressed = false;
      if (mode === "scribble" && pts.length > 1) finish(pts.length > 6 && near(pts[pts.length - 1], pts[0]));
    });
    layer.addEventListener("dblclick", (ev) => { ev.preventDefault(); ev.stopPropagation(); finish(false); });
    const key = (ev) => {
      if (ev.key === "Escape") { ev.preventDefault(); ev.stopPropagation(); stop(); app.toast("描くのをやめました"); }
      else if (ev.key === "Enter") { ev.preventDefault(); ev.stopPropagation(); finish(false); }
    };
    document.addEventListener("keydown", key, true);
    drawing = { layer, key };
    paint();
    app.toast(mode === "curve" ? "曲線：クリックで点を打ち、ダブルクリックで終わります（最初の点に戻ると閉じた図形に）" : mode === "polygon" ? "フリーフォーム：クリックで角、ドラッグで自由に。最初の点に戻ると閉じた図形に（ダブルクリックで線のまま終わる）" : "フリーハンド：押したままなぞって描きます");
  }
  function stop() {
    if (!drawing) return;
    drawing.layer.remove();
    document.removeEventListener("keydown", drawing.key, true);
    drawing = null;
  }

  /** The drawn points (slide pixels) as a shape: its box, and the points as fractions of it. */
  function make(points, { closed, curve }) {
    const xs = points.map((p) => p[0]);
    const ys = points.map((p) => p[1]);
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    const w = Math.max(4, Math.max(...xs) - x);
    const hh = Math.max(4, Math.max(...ys) - y);
    const path = { pts: points.map(([px, py]) => [round4((px - x) / w), round4((py - y) / hh)]), ...(closed ? { closed: true } : {}), ...(curve ? { curve: true } : {}) };
    const o = ops.makeObject("shape", { x, y, w, h: hh }, { shape: "custom", path, ...(closed ? { fill: E.OBJECT_DEFAULTS.shape.fill } : { fill: "none", stroke: E.OBJECT_DEFAULTS.line.stroke, strokeW: 4 }) });
    editor.commit([...editor.objects(), o], { select: [o.id] });
  }
  const round4 = (v) => Math.round(v * 10000) / 10000;
  function simplify(pts, tolerance = 4) {
    if (pts.length < 3) return pts;
    const keep = new Array(pts.length).fill(false);
    keep[0] = keep[pts.length - 1] = true;
    const stack = [[0, pts.length - 1]];
    while (stack.length) {
      const [a, b] = stack.pop();
      let far = -1;
      let dist = 0;
      for (let i = a + 1; i < b; i += 1) { const d = ops.distToSegment(pts[i], pts[a], pts[b]); if (d > dist) { dist = d; far = i; } }
      if (far >= 0 && dist > tolerance) { keep[far] = true; stack.push([a, far], [far, b]); }
    }
    return pts.filter((_, i) => keep[i]);
  }

  // ---------------------------------------------------------------- 頂点の編集

  /** Show a drawn shape's points to drag (click one, Delete removes it; double-click a side adds one). */
  function editPoints(id) {
    const o = editor.objects().find((x) => x.id === id);
    if (o?.kind !== "shape" || o.shape !== "custom") return;
    editing = { id, pick: null };
    document.addEventListener("keydown", onEditKey, true);
    document.addEventListener("pointerdown", onEditOutside, true);
    editor.state.wrap?.addEventListener("dblclick", onAddPoint, true);
    editor.draw();
    app.toast("頂点の編集：点をドラッグで動かす・点を選んでDeleteで消す・辺をダブルクリックで点を足す（Escで終わる）");
  }
  function endEdit() {
    if (!editing) return;
    editing = null;
    document.removeEventListener("keydown", onEditKey, true);
    document.removeEventListener("pointerdown", onEditOutside, true);
    editor.state.wrap?.removeEventListener("dblclick", onAddPoint, true);
    editor.draw();
  }
  // A merged shape (図形の結合) has more rings than its first: every ring's points can be edited (picks are [ring, i]).
  function onEditKey(event) {
    if (event.key === "Escape" || event.key === "Enter") { event.preventDefault(); event.stopPropagation(); endEdit(); return; }
    if ((event.key === "Delete" || event.key === "Backspace") && editing?.pick) {
      event.preventDefault();
      event.stopPropagation();
      const o = editor.objects().find((x) => x.id === editing.id);
      const [r, i] = editing.pick;
      const rings = absolute(o);
      rings[r] = rings[r].filter((_, j) => j !== i);
      const min = o.path.closed || r > 0 ? 3 : 2;
      if (rings[r].length < min) {
        // A part of a merged shape goes when too few points are left; the first ring always stays.
        if (r === 0 || rings.length < 2) { app.toast("これ以上は消せません"); return; }
        rings.splice(r, 1);
      }
      editing.pick = null;
      save(o, rings);
    }
  }
  function onEditOutside(event) {
    if (event.target.closest?.(".ed-vertex, .ribbon, .rb-pop")) return;
    const o = editor.objects().find((x) => x.id === editing?.id);
    const k = editor.scale();
    const [px, py] = editor.toSlide(event);
    if (o && px >= o.x - 8 / k && px <= o.x + o.w + 8 / k && py >= o.y - 8 / k && py <= o.y + o.h + 8 / k) return;
    endEdit();
  }
  function onAddPoint(event) {
    const o = editor.objects().find((x) => x.id === editing?.id);
    if (!o) return;
    event.preventDefault();
    event.stopPropagation();
    const p = editor.toSlide(event);
    const rings = absolute(o);
    let best = [0, 0];
    let dist = Infinity;
    rings.forEach((abs, r) => {
      const closed = o.path.closed || r > 0;
      const segs = abs.length - (closed ? 0 : 1);
      for (let i = 0; i < segs; i += 1) { const d = ops.distToSegment(p, abs[i], abs[(i + 1) % abs.length]); if (d < dist) { dist = d; best = [r, i]; } }
    });
    rings[best[0]].splice(best[1] + 1, 0, p);
    save(o, rings);
  }
  const absolute = (o) => [o.path.pts, ...(o.path.parts || [])].map((ring) => ring.map(([fx, fy]) => [o.x + fx * o.w, o.y + fy * o.h]));
  function save(o, rings) {
    const all = rings.flat();
    const xs = all.map((p) => p[0]);
    const ys = all.map((p) => p[1]);
    const x = Math.min(...xs);
    const y = Math.min(...ys);
    const w = Math.max(4, Math.max(...xs) - x);
    const hh = Math.max(4, Math.max(...ys) - y);
    const frac = (ring) => ring.map(([px, py]) => [round4((px - x) / w), round4((py - y) / hh)]);
    const [first, ...parts] = rings.map(frac);
    const path = { ...o.path, pts: first };
    if (parts.length) path.parts = parts; else delete path.parts;
    editor.commit(editor.objects().map((item) => (item.id === o.id ? { ...item, x, y, w, h: hh, rot: undefined, path } : item)), { select: [o.id] });
  }

  function overlay(list, k) {
    if (!editing) return [];
    const o = list.find((x) => x.id === editing.id);
    if (!o?.path) return [];
    const rings = editing.preview || absolute(o);
    // While a point moves, the shape's new outline follows it.
    const ghost = editing.preview ? [(() => {
      const svg = E.s("svg", { class: "ed-anim-path", width: E.W * k, height: E.H * k, viewBox: `0 0 ${E.W} ${E.H}`, "aria-hidden": "true" });
      for (const [r, pts] of rings.entries()) svg.append(E.s("path", { d: E.freeformD({ ...o.path, pts, ...(r > 0 ? { closed: true, curve: false } : {}) }, 1, 1), fill: "none", stroke: "#2b6be0", "stroke-width": 2 / k, "stroke-dasharray": `${6 / k} ${4 / k}` }));
      return svg;
    })()] : [];
    const on = (r, i) => editing.pick && editing.pick[0] === r && editing.pick[1] === i;
    return [...ghost, ...rings.flatMap((pts, r) => pts.map(([x, y], i) => h("span", { class: ["ed-handle", "ed-vertex", on(r, i) ? "on" : ""], "data-handle": `vtx:${r}:${i}`, title: "ドラッグで動かす・選んでDeleteで消す", style: { left: `${x * k}px`, top: `${y * k}px` } })))];
  }
  editor.overlay(overlay);
  editor.handle("vtx:", (event, handle) => {
    if (!editing) return;
    const [, rs, is] = handle.split(":");
    const r = Number(rs);
    const i = Number(is);
    const o = editor.objects().find((x) => x.id === editing.id);
    if (!o) return;
    editing.pick = [r, i];
    const start = absolute(o);
    const [sx, sy] = editor.toSlide(event);
    let moved = false;
    const move = (ev) => {
      const [x, y] = editor.toSlide(ev);
      moved = true;
      editing.preview = start.map((ring, rr) => (rr === r ? ring.map((p, j) => (j === i ? [p[0] + x - sx, p[1] + y - sy] : p)) : ring));
      editor.draw();
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      const rings = editing?.preview;
      if (editing) editing.preview = null;
      if (moved && rings) save(o, rings); else editor.draw();
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  });

  return { start, stop, editPoints, endEdit, get editing() { return Boolean(editing); } };
}
