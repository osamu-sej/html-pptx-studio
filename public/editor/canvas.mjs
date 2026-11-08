// The slide canvas, PowerPoint style: click to select (Shift/Ctrl to add), drag to move (smart guides, Shift for a
// straight line, Ctrl to copy, Alt to place freely), handles to resize (Shift keeps the shape, Ctrl from the centre)
// and rotate (Shift: 15° steps), yellow handles to adjust a shape, a box drawn on the background to select several,
// double-click to type in a shape, groups, the clipboard (objects, pictures, text) and the keyboard.
// Objects are slide.elements (see public/engine/objects.js); every change is one undo step.

import * as ops from "./ops.mjs";
import { autoCorrect, autoCorrectOptions } from "./autocorrect.mjs";
import { cellPatch, richOf } from "./sheet.mjs";

const SNAP_SCREEN = 7; // px on screen
const NUDGE = { plain: 5, fine: 1, big: 25 };
const FONT_SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96, 120, 150, 200];
// グリッドとガイド → 間隔 (cm).
export const GRID_STEPS = [0.1, 0.2, 0.25, 0.5, 1, 2];
const STYLE_KEYS = ["fill", "gradient", "pattern", "fillOpacity", "stroke", "strokeW", "dash", "fs", "color", "bold", "italic", "underline", "uline", "strike", "sline", "caps", "tshadow", "tglow", "toutline", "align", "valign", "font", "lh", "ls", "psp", "pad", "autofit", "opacity", "head", "tail", "headSize", "tailSize", "route", "vertical"];

const stored = (key, fallback) => { try { const v = localStorage.getItem(`hpptx-editor-${key}`); return v == null ? fallback : JSON.parse(v); } catch { return fallback; } };
const store = (key, value) => { try { localStorage.setItem(`hpptx-editor-${key}`, JSON.stringify(value)); } catch { /* private window */ } };

export function createCanvas(app) {
  const { E, h } = app;
  const ed = {
    sel: [], slideIndex: -1, entered: null, typing: null, tool: null, drag: null,
    wrap: null, slideEl: null, layer: null, painter: null, clipboard: null, pasteCount: 0, suppressClick: false,
    zoom: stored("zoom", "fit"), grid: stored("grid", false), snapGrid: stored("snapGrid", false), smart: stored("smart", true), guides: stored("guides", true), gridStep: GRID_STEPS.includes(stored("gridStep", 0.25)) ? stored("gridStep", 0.25) : 0.25, alignTo: stored("alignTo", "auto") === "slide" ? "slide" : "auto",
    menu: null, range: null, picker: null, cellRange: null,
  };
  const subs = new Set();
  const emit = () => { for (const fn of subs) { try { fn(); } catch (error) { console.error(error); } } };

  // ---------------------------------------------------------------- data

  const slide = () => app.slide();
  const objects = () => (Array.isArray(slide()?.elements) ? slide().elements : []);
  const byId = (id, list = objects()) => list.find((o) => o.id === id);
  const selected = (list = objects()) => list.filter((o) => ed.sel.includes(o.id));
  const visible = (list = objects()) => list.filter((o) => !o.hidden);
  const ends = (list) => (o) => E.lineEnds(o, list);

  // `timeline` replaces the slide's animations in the same undo step (copies of animated objects bring theirs).
  // While a dialog applies several settings (フォント・段落), they make one undo step together.
  let batching = null;
  function commit(next, { select = null, undo = true, timeline = undefined } = {}) {
    if (select) ed.sel = [...select];
    const step = undo && !batching?.done;
    if (batching && undo) batching.done = true;
    app.setObjects(next.map((o) => E.normalizeObject(o) || o).filter(Boolean), { undo: step, timeline });
    emit();
  }
  function batch(fn) {
    const outer = batching;
    batching = outer || { done: false };
    try { fn(); } finally { batching = outer; }
  }
  const timeline = () => (Array.isArray(slide()?.timeline) ? slide().timeline : []);
  function change(fn, opts) {
    const list = objects();
    const next = fn(list);
    if (next === list) return;
    commit(next, opts);
  }

  // ---------------------------------------------------------------- geometry on screen

  // The slide fills the width of its wrapper (engine mount): measure the wrapper, which is right even before
  // the slide's own scaling has been applied.
  const scale = () => { const r = ed.wrap?.getBoundingClientRect(); return r && r.width ? r.width / E.W : 1; };
  function toSlide(event) {
    const r = ed.wrap.getBoundingClientRect();
    const k = r.width / E.W || 1;
    return [(event.clientX - r.left) / k, (event.clientY - r.top) / k];
  }
  const ctx = () => ({ ...app.renderOptions(), mode: "edit", live: false, deck: app.deck() });

  // ---------------------------------------------------------------- attaching to the rendered slide

  // The overlay follows the slide whenever its size changes (window, ribbon, side panels, zoom), and in
  // 「全体表示」 the slide follows the room the stage has.
  const watch = typeof ResizeObserver === "function" ? {
    slide: new ResizeObserver(() => { if (!ed.drag && ed.layer?.isConnected) { ed.wrap.querySelector(".ed-grid")?.style.setProperty("--grid", `${gridPx()}px`); draw(); } }),
    room: new ResizeObserver(() => { if (ed.wrap?.isConnected) applyZoom(); }),
    wrap: null,
    body: null,
  } : null;
  function observe(wrap) {
    if (!watch) return;
    const body = wrap.closest(".stage-body");
    if (watch.wrap !== wrap) { if (watch.wrap) watch.slide.unobserve(watch.wrap); watch.slide.observe(wrap); watch.wrap = wrap; }
    if (body && watch.body !== body) { if (watch.body) watch.room.unobserve(watch.body); watch.room.observe(body); watch.body = body; }
  }

  /** Called by the studio each time the slide on the stage is drawn. */
  function attach(wrap, slideEl, index) {
    if (index !== ed.slideIndex) {
      ed.sel = [];
      ed.entered = null;
      ed.typing = null;
      ed.slideIndex = index;
    }
    ed.wrap = wrap;
    ed.slideEl = slideEl;
    const ids = new Set(objects().map((o) => o.id));
    ed.sel = ed.sel.filter((id) => ids.has(id));
    wrap.classList.add("ed-wrap");
    wrap.classList.toggle("ed-drawing", Boolean(ed.tool));
    ed.layer = h("div", { class: "ed-layer", "aria-hidden": "true" });
    wrap.append(ed.layer);
    if (ed.grid) wrap.append(gridLayer());
    prepareHits();
    wrap.addEventListener("pointerdown", onPointerDown);
    wrap.addEventListener("dblclick", onDoubleClick);
    wrap.addEventListener("contextmenu", onContextMenu);
    applyZoom();
    observe(wrap);
    syncGrow();
    draw();
    emit();
  }

  /** Lines get a wide invisible stroke to click on; text boxes and pictures take clicks anywhere on them. */
  function prepareHits(scope = ed.slideEl) {
    for (const svg of scope?.querySelectorAll?.(".hs-obj-line .hs-obj-geom") ?? []) {
      const path = svg.querySelector("path");
      if (!path || svg.querySelector(".ed-hit")) continue;
      const hit = document.createElementNS("http://www.w3.org/2000/svg", "path");
      hit.setAttribute("d", path.getAttribute("d"));
      hit.setAttribute("class", "ed-hit");
      hit.setAttribute("fill", "none");
      hit.setAttribute("stroke", "transparent");
      hit.setAttribute("stroke-width", String(Math.max(22, 14 / scale())));
      svg.prepend(hit);
    }
  }

  // グリッドとガイド: lines every gridStep cm (doubled until they are at least 10 px apart on the screen).
  function gridPx() {
    const k = scale();
    let step = ops.PX_PER_CM * ed.gridStep;
    while (step * k < 10) step *= 2;
    return step * k;
  }
  function gridLayer() {
    return h("div", { class: "ed-grid", style: { "--grid": `${gridPx()}px` } });
  }

  // ---------------------------------------------------------------- zoom

  // Zoom as PowerPoint's: a percentage of the slide's real size (13.33 inches at 96 dpi = 1280 px wide), or "fit"
  // (現在のウィンドウ サイズに合わせる): the whole slide in the room the stage has, as large as it fits.
  const REAL_W = 1280;
  const ZOOM_MIN = 10;
  const ZOOM_MAX = 400;
  // An older studio stored the zoom as a share of the 1920 px slide (0.5 … 2).
  ed.zoom = ed.zoom === "fit" ? "fit" : Number(ed.zoom) > 0 ? Math.round(Number(ed.zoom) <= 4 ? Number(ed.zoom) * (E.W / REAL_W) * 100 : Number(ed.zoom)) : "fit";
  // The room inside the stage's padding (app.css .stage-body; more on the top and left when the rulers show).
  function fitWidth(body) {
    const st = getComputedStyle(body);
    const padX = parseFloat(st.paddingLeft) + parseFloat(st.paddingRight);
    const padY = parseFloat(st.paddingTop) + parseFloat(st.paddingBottom);
    return Math.max(160, Math.min(body.clientWidth - padX, (body.clientHeight - padY) * (E.W / E.H)));
  }
  function applyZoom() {
    const single = ed.wrap?.closest(".stage-single");
    const body = ed.wrap?.closest(".stage-body");
    if (!single || !body) return;
    const width = ed.zoom === "fit" ? fitWidth(body) : REAL_W * ed.zoom / 100;
    single.style.width = `${Math.round(width)}px`;
    single.style.maxWidth = "none";
    app.zoomChanged?.();
  }
  /** "fit" or a percentage (10–400). The slide rescales in place: no redraw of the stage. */
  function setZoom(zoom) {
    ed.zoom = zoom === "fit" ? "fit" : Math.round(Math.max(ZOOM_MIN, Math.min(ZOOM_MAX, Number(zoom) || 100)));
    store("zoom", ed.zoom);
    if (ed.wrap?.isConnected) { applyZoom(); requestAnimationFrame(() => draw()); } else app.zoomChanged?.();
  }
  /** The percentage shown now (in "fit", what the fitted slide comes to). */
  function zoomPercent() {
    if (ed.zoom !== "fit") return ed.zoom;
    const w = ed.wrap?.getBoundingClientRect().width;
    return w ? Math.round((w / REAL_W) * 100) : 100;
  }
  /** One step in or out (the status bar's − / ＋, ⌘＋ホイール), from what is shown now. */
  function zoomStep(direction) {
    const now = zoomPercent();
    const next = direction > 0 ? Math.min(ZOOM_MAX, Math.round(now * 1.1 + 1)) : Math.max(ZOOM_MIN, Math.round(now / 1.1 - 1));
    setZoom(next);
  }

  // ---------------------------------------------------------------- the overlay (selection, handles, guides)

  function draw(preview = null, extra = null) {
    const layer = ed.layer;
    if (!layer || !ed.slideEl) return;
    const list = preview || objects();
    const k = scale();
    layer.replaceChildren();
    const chosen = list.filter((o) => ed.sel.includes(o.id) && !o.hidden);
    const asGroup = chosen.length > 1 && chosen.every((o) => o.group && o.group === chosen[0].group) && ed.entered !== chosen[0].group;
    const single = chosen.length === 1 ? chosen[0] : null;
    for (const o of chosen) layer.append(outline(o, list, k, { quiet: !single, typing: ed.typing?.id === o.id }));
    if (single && !single.locked && ed.typing?.id !== single.id) layer.append(...handlesFor(single, list, k));
    else if (chosen.length > 1 && !chosen.some((o) => o.locked)) layer.append(boxHandles(ops.bounds(chosen, ends(list)), k, { rotate: asGroup, group: asGroup }));
    // Animation order badges and other marks drawn by others (the animation pane).
    for (const fn of overlays) layer.append(...(fn(list, k) || []));
    if (extra) layer.append(...extra);
    // The deck's guides lie under everything else drawn here.
    if (ed.guides) layer.prepend(...guideMarks(k));
  }

  // ---------------------------------------------------------------- guides (PowerPoint's ガイド)
  // Dashed lines across the slide that objects snap to: dragged to move (out of the slide to delete, with Ctrl to
  // copy), added from the 表示 tab or a right-click. They belong to the deck (deck.guides: slide px).
  const GUIDE_MAX = 20;
  function customGuides() {
    const g = app.deck()?.guides;
    return { x: Array.isArray(g?.x) ? g.x : [], y: Array.isArray(g?.y) ? g.y : [] };
  }
  /** What moving objects snap to: the slide's own guides (content area, middle) and the deck's. */
  function allGuides() {
    const g = customGuides();
    return { x: [...ops.SLIDE_GUIDES.x, ...g.x], y: [...ops.SLIDE_GUIDES.y, ...g.y] };
  }
  function guideMarks(k) {
    const g = customGuides();
    const title = "ガイド：ドラッグで移動（スライドの外へ出すと削除・Ctrlを押しながらでコピー）";
    return [
      ...g.x.map((v, i) => h("div", { class: "ed-cguide v", "data-axis": "x", "data-i": String(i), style: { left: `${v * k}px` }, title })),
      ...g.y.map((v, i) => h("div", { class: "ed-cguide h", "data-axis": "y", "data-i": String(i), style: { top: `${v * k}px` }, title })),
    ];
  }
  function setGuides(next) {
    app.setGuides({ x: next.x.slice(0, GUIDE_MAX), y: next.y.slice(0, GUIDE_MAX) });
    draw();
  }
  /** A new guide (vertical "x" or horizontal "y"), in the middle of the slide unless told where. */
  function addGuide(axis, at = axis === "x" ? E.W / 2 : E.H / 2) {
    const g = customGuides();
    if (g[axis].length >= GUIDE_MAX) { app.toast(`ガイドは縦・横それぞれ${GUIDE_MAX}本までです`); return; }
    if (!ed.guides) setView("guides", true);
    setGuides({ ...g, [axis]: [...g[axis], Math.round(at * 2) / 2] });
    app.toast(axis === "x" ? "垂直ガイドを追加しました（ドラッグで移動）" : "水平ガイドを追加しました（ドラッグで移動）");
  }
  function clearGuides() { setGuides({ x: [], y: [] }); }
  function dragGuide(event, axis, i) {
    event.preventDefault();
    event.stopPropagation();
    const g = customGuides();
    const copy = event.ctrlKey || event.metaKey;
    const el = event.target.closest(".ed-cguide");
    const size = axis === "x" ? E.W : E.H;
    const label = h("div", { class: "ed-cguide-label" });
    ed.layer.append(label);
    let at = g[axis][i];
    const move = (e) => {
      const p = toSlide(e);
      const k = scale();
      at = Math.round((axis === "x" ? p[0] : p[1]) * 2) / 2;
      const out = at < 0 || at > size;
      el.style[axis === "x" ? "left" : "top"] = `${at * k}px`;
      el.classList.toggle("out", out);
      // As PowerPoint shows it: the distance from the middle of the slide, in cm.
      const cm = Math.abs((at - size / 2) / ops.PX_PER_CM).toFixed(2);
      label.textContent = out ? "削除" : `${cm} cm`;
      label.style.left = `${p[0] * k + 10}px`;
      label.style.top = `${p[1] * k - 26}px`;
    };
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
      label.remove();
      const next = { x: [...g.x], y: [...g.y] };
      if (at < 0 || at > size) { if (!copy) next[axis].splice(i, 1); } else if (copy) next[axis].push(at); else next[axis][i] = at;
      if (JSON.stringify(next) !== JSON.stringify(g)) setGuides(next); else draw();
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }
  const overlays = new Set();
  const handles = new Map();
  // "Click an object" mode (the animation painter, a trigger): the next object clicked is handed over.
  function pick(fn, hint) {
    ed.picker = fn ? { fn, hint } : null;
    wrapClass();
    if (fn && hint) app.toast(hint);
    emit();
  }
  function wrapClass() { ed.wrap?.classList.toggle("ed-picking", Boolean(ed.picker)); }

  function outline(o, list, k, { quiet = false, typing = false } = {}) {
    if (o.kind === "line") {
      const [[x1, y1], [x2, y2]] = E.lineEnds(o, list);
      const len = Math.hypot(x2 - x1, y2 - y1) * k;
      const ang = Math.atan2(y2 - y1, x2 - x1);
      return h("div", { class: ["ed-line-outline", quiet ? "quiet" : ""], style: { left: `${x1 * k}px`, top: `${y1 * k}px`, width: `${len}px`, transform: `rotate(${ang}rad)` } });
    }
    return h("div", {
      class: ["ed-outline", quiet ? "quiet" : "", typing ? "typing" : "", o.locked ? "locked" : ""],
      style: { left: `${o.x * k}px`, top: `${o.y * k}px`, width: `${o.w * k}px`, height: `${o.h * k}px`, transform: o.rot ? `rotate(${o.rot}deg)` : null },
    }, o.locked ? h("span", { class: "ed-lock", title: "ロック中（右クリックで解除）" }, "🔒") : null);
  }

  /** Handles for one object: eight to resize, one to rotate, yellow ones to adjust its shape. */
  function handlesFor(o, list, k) {
    if (o.kind === "line") {
      const [p1, p2] = E.lineEnds(o, list);
      return [p1, p2].map(([x, y], i) => h("span", { class: ["ed-handle", "ed-end", (i ? o.to : o.from) ? "attached" : ""], "data-handle": `end${i + 1}`, style: { left: `${x * k}px`, top: `${y * k}px` }, title: "ドラッグで端を動かす（図形の点に近づけるとつながります）" }));
    }
    const box = boxHandles({ x: o.x, y: o.y, w: o.w, h: o.h }, k, { rotate: true, rot: o.rot || 0 });
    const shape = ["shape", "text"].includes(o.kind) ? E.SHAPES[o.shape] : o.kind === "image" && o.mask ? E.SHAPES[o.mask] : null;
    if (shape?.handles?.length && o.w * k > 24 && o.h * k > 24) {
      const adj = E.adjOf(shape, o.adj);
      shape.handles.forEach((handle, i) => {
        let [x, y] = handle.pos(o.w, o.h, adj);
        if (o.flipH) x = o.w - x;
        if (o.flipV) y = o.h - y;
        box.firstChild.append(h("span", { class: "ed-handle ed-adjust", "data-handle": `adj${i}`, style: { left: `${x * k}px`, top: `${y * k}px` }, title: "ドラッグで形を調整" }));
      });
    }
    return [box];
  }

  function boxHandles(b, k, { rotate = false, rot = 0, group = false } = {}) {
    const frame = h("div", { class: ["ed-frame", group ? "group" : ""], style: { left: `${b.x * k}px`, top: `${b.y * k}px`, width: `${b.w * k}px`, height: `${b.h * k}px`, transform: rot ? `rotate(${rot}deg)` : null } });
    const small = b.w * k < 36 || b.h * k < 36;
    for (const name of ops.HANDLE_NAMES) {
      if (small && name.length === 1) continue;
      frame.append(h("span", { class: `ed-handle ed-size ed-${name}`, "data-handle": name }));
    }
    if (rotate) frame.append(h("span", { class: "ed-stem" }), h("span", { class: "ed-handle ed-rotate", "data-handle": "rotate", title: "ドラッグで回転（Shiftで15°ずつ）" }, "⟳"));
    return h("div", { class: "ed-box" }, frame);
  }

  // ---------------------------------------------------------------- drawing objects again while they move

  function paint(list, ids) {
    if (!ed.slideEl) return;
    let layer = ed.slideEl.querySelector(".hs-objects");
    if (!layer) {
      layer = h("div", { class: "hs-objects" });
      const overlay = ed.slideEl.querySelector(".hs-overlay");
      if (overlay) overlay.before(layer); else ed.slideEl.append(layer);
    }
    const touched = new Set(ids);
    // Connectors follow the objects they are attached to.
    for (const o of list) if (o.kind === "line" && ((o.from && touched.has(o.from.id)) || (o.to && touched.has(o.to.id)))) touched.add(o.id);
    const c = ctx();
    for (const id of touched) {
      const o = byId(id, list);
      const old = layer.querySelector(`[data-el="${CSS.escape(id)}"]`);
      if (!o || o.hidden) { old?.remove(); continue; }
      const node = E.objectNode(o, c, list);
      if (old) old.replaceWith(node);
      else {
        // Keep the drawing order: insert before the next object that is already drawn.
        const index = list.findIndex((item) => item.id === id);
        const after = list.slice(index + 1).map((item) => layer.querySelector(`[data-el="${CSS.escape(item.id)}"]`)).find(Boolean);
        if (after) after.before(node); else layer.append(node);
      }
    }
    prepareHits(layer);
  }

  // ---------------------------------------------------------------- pointer: select, move, resize, rotate, adjust, draw

  function targetsFor(list, moving) {
    const set = new Set(moving);
    return visible(list).filter((o) => !set.has(o.id)).map((o) => ops.bounds([o], ends(list)));
  }

  /** The object under the pointer (by position: after a drag the events go to the stage, not the object). */
  function objectAt(event, list) {
    for (const el of document.elementsFromPoint(event.clientX, event.clientY)) {
      const node = el.closest?.(".hs-obj");
      if (node && ed.slideEl?.contains(node)) return byId(node.dataset.el, list) || null;
      if (el === ed.wrap) break;
    }
    return null;
  }

  function hitObject(event, list) {
    const found = objectAt(event, list);
    if (found) return found;
    // A selected object can be dragged from anywhere inside its box.
    const p = toSlide(event);
    return [...selected(list)].reverse().find((o) => !o.hidden && ops.hits(o, p, 4 / scale())) || null;
  }

  function onPointerDown(event) {
    if (event.button === 2) return;
    if (event.button !== 0 || !ed.slideEl || app.busy()) return;
    if (event.target.closest(".inline-tools, .ph-handle, .hs-placed, .ed-menu, .motion-banner, .ed-comment-pin")) return;
    closeMenu();
    const guide = event.target.closest?.(".ed-cguide");
    if (guide) { dragGuide(event, guide.dataset.axis, Number(guide.dataset.i)); return; }
    const list = objects();
    const p = toSlide(event);
    const handle = event.target.closest?.(".ed-handle")?.dataset.handle;
    // Typing in a table: another cell takes the caret, Shift+click (or a drag) selects a block of cells.
    if (ed.typing?.cell && event.target.closest?.("td[data-r]") && ed.typing.node.contains(event.target)) {
      const td = event.target.closest("td[data-r]");
      const here = ops.tableOrigin(byId(ed.typing.id), Number(td.dataset.r), Number(td.dataset.c));
      const [r0, c0] = ed.typing.cell;
      if (event.shiftKey) { event.preventDefault(); setCellRange([r0, c0, here[0], here[1]]); return; }
      if (here[0] !== r0 || here[1] !== c0) {
        event.preventDefault();
        const id = ed.typing.id;
        stopTyping(true);
        startTyping(id, { cell: here, at: [event.clientX, event.clientY] });
        return;
      }
      setCellRange(null);
      dragCells(event, [r0, c0]);
      return;
    }
    // Typing: clicks inside the text keep editing.
    if (ed.typing && event.target.closest(".ed-typing-tx")) return;
    if (ed.typing) stopTyping(true);
    if (ed.tool) { start(event, { type: "draw", tool: ed.tool, from: p, list }); return; }
    // Handles drawn by others on the overlay (a motion path's points) drag themselves.
    const external = handle && [...handles.entries()].find(([prefix]) => handle.startsWith(prefix));
    if (external) { event.preventDefault(); external[1](event, handle, p); return; }
    if (handle) { start(event, { type: handle === "rotate" ? "rotate" : handle.startsWith("adj") ? "adjust" : handle.startsWith("end") ? "end" : "resize", handle, from: p, list }); return; }
    const hit = hitObject(event, list);
    if (hit) {
      const additive = event.shiftKey || event.ctrlKey || event.metaKey;
      // Clicking outside the group being edited leaves it; a grouped object selects its whole group.
      if (ed.entered && hit.group !== ed.entered) ed.entered = null;
      const ids = hit.group && ed.entered !== hit.group ? ops.withGroups(list, [hit.id]) : [hit.id];
      if (ed.painter) { pasteFormat([hit.id]); event.preventDefault(); return; }
      if (ed.picker) { const fn = ed.picker.fn; ed.picker = null; wrapClass(); fn(hit.id, ids); event.preventDefault(); emit(); return; }
      if (additive) {
        const on = ids.every((id) => ed.sel.includes(id));
        ed.sel = on ? ed.sel.filter((id) => !ids.includes(id)) : [...new Set([...ed.sel, ...ids])];
        draw();
        emit();
        event.preventDefault();
        return;
      }
      if (!ids.every((id) => ed.sel.includes(id))) { ed.sel = ids; draw(); emit(); }
      app.focusStage();
      start(event, { type: "move", from: p, list });
      return;
    }
    // The background: a box selects what it encloses; a plain click clears the selection.
    if (!(event.shiftKey || event.ctrlKey || event.metaKey) && ed.sel.length) { ed.sel = []; ed.entered = null; draw(); emit(); }
    start(event, { type: "marquee", from: p, list, quiet: true });
  }

  function start(event, drag) {
    ed.drag = { ...drag, pointer: event.pointerId, moved: false, client: [event.clientX, event.clientY], startSel: [...ed.sel] };
    // A press on the background may be a plain click on the slide (a title to edit, a drill badge): the
    // pointer is only captured once it really drags, so that click still reaches what is under it.
    if (!drag.quiet) { event.preventDefault(); capture(event); }
    // The window follows the drag, not the stage: the studio may draw the slide again while the button is down
    // (a new wrap replaces the old one, as on a slow machine right after an object was placed), and the drag
    // must still end where the button comes up.
    window.addEventListener("pointermove", onPointerMove);
    window.addEventListener("pointerup", onPointerUp);
    window.addEventListener("pointercancel", onPointerUp);
  }

  function capture(event) { try { ed.wrap.setPointerCapture(event.pointerId); } catch { /* gone */ } }

  function onPointerMove(event) {
    const drag = ed.drag;
    if (!drag) return;
    if (!drag.moved && Math.hypot(event.clientX - drag.client[0], event.clientY - drag.client[1]) < 4) return;
    if (!drag.moved && drag.quiet) capture(event);
    drag.moved = true;
    const p = toSlide(event);
    const k = scale();
    const tolerance = SNAP_SCREEN / k;
    const snapping = ed.smart && !event.altKey;
    const result = { list: drag.list, extra: [] };
    if (drag.type === "move") moveDrag(drag, p, event, tolerance, snapping, result);
    else if (drag.type === "resize") resizeDrag(drag, p, event, tolerance, snapping, result);
    else if (drag.type === "rotate") rotateDrag(drag, p, event, result);
    else if (drag.type === "adjust") adjustDrag(drag, p, result);
    else if (drag.type === "end") endDrag(drag, p, event, tolerance, result);
    else if (drag.type === "marquee") marqueeDrag(drag, p, result);
    else if (drag.type === "draw") drawDrag(drag, p, event, result);
    drag.result = result.list;
    if (result.list !== drag.list) paint(result.list, drag.touched ?? ed.sel);
    draw(result.list, [...result.extra, ...guideLines(result.lines || [], k)]);
  }

  function onPointerUp(event) {
    const drag = ed.drag;
    ed.drag = null;
    window.removeEventListener("pointermove", onPointerMove);
    window.removeEventListener("pointerup", onPointerUp);
    window.removeEventListener("pointercancel", onPointerUp);
    if (!drag) return;
    if (drag.moved) ed.suppressClick = true;
    if (drag.type === "draw") { finishDraw(drag, toSlide(event), event); return; }
    if (drag.type === "marquee") {
      if (drag.moved && drag.box) {
        const inside = visible(drag.list).filter((o) => { const b = ops.bounds([o], ends(drag.list)); return b.x >= drag.box.x && b.y >= drag.box.y && b.x + b.w <= drag.box.x + drag.box.w && b.y + b.h <= drag.box.y + drag.box.h; });
        const ids = ops.withGroups(drag.list, inside.map((o) => o.id));
        ed.sel = event.shiftKey || event.ctrlKey || event.metaKey ? [...new Set([...drag.startSel, ...ids])] : ids;
        emit();
      }
      draw();
      return;
    }
    if (!drag.moved || !drag.result || drag.result === drag.list) { draw(); return; }
    const next = drag.result;
    // Ctrl+drag leaves the originals and drops copies.
    if (drag.type === "move" && (event.ctrlKey || event.metaKey) && drag.copy) {
      const { list, ids } = ops.duplicate(drag.list, ed.sel, 0);
      const moved = ops.moveBy(list, ids, drag.dx, drag.dy);
      commit(moved, { select: ids });
      return;
    }
    commit(next);
  }

  // -- move
  function moveDrag(drag, p, event, tolerance, snapping, result) {
    const ids = ed.sel.filter((id) => !byId(id, drag.list)?.locked);
    if (!ids.length) return;
    let dx = p[0] - drag.from[0];
    let dy = p[1] - drag.from[1];
    if (event.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
    const moving = drag.list.filter((o) => ids.includes(o.id));
    const b0 = ops.bounds(moving, ends(drag.list));
    let lines = [];
    if (snapping || ed.snapGrid) {
      const snap = ops.snapBox({ x: b0.x + dx, y: b0.y + dy, w: b0.w, h: b0.h }, snapping ? targetsFor(drag.list, ids) : [], { tolerance, guides: snapping && ed.guides ? allGuides() : null, grid: ed.snapGrid ? ops.PX_PER_CM * ed.gridStep : 0 });
      if (!(event.shiftKey && dx === 0)) dx += snap.dx;
      if (!(event.shiftKey && dy === 0)) dy += snap.dy;
      lines = snap.lines;
    }
    drag.dx = dx;
    drag.dy = dy;
    drag.copy = event.ctrlKey || event.metaKey;
    ed.wrap.classList.toggle("ed-copying", drag.copy);
    result.list = ops.moveBy(drag.list, ids, dx, dy);
    result.lines = lines;
    drag.touched = ids;
  }

  // -- resize
  function resizeDrag(drag, p, event, tolerance, snapping, result) {
    const chosen = drag.list.filter((o) => ed.sel.includes(o.id));
    if (!chosen.length) return;
    const fromCenter = event.ctrlKey || event.metaKey;
    if (chosen.length === 1) {
      const o = chosen[0];
      // Pictures (and objects set to keep their proportions) keep them from the corners; Shift does it for any.
      const keep = event.shiftKey || (drag.handle.length === 2 && (o.lockRatio || ["image", "icon", "video", "audio", "lottie"].includes(o.kind)));
      let next = ops.resizeBox(o, drag.handle, p, { keepRatio: keep, fromCenter });
      let lines = [];
      // Upright boxes snap their moving edges to other objects and the guides.
      if (snapping && !o.rot) {
        const targets = targetsFor(drag.list, [o.id]);
        const xs = [...targets.flatMap((t) => [t.x, t.x + t.w / 2, t.x + t.w]), 0, E.W / 2, E.W, ...(ed.guides ? allGuides().x : [])];
        const ys = [...targets.flatMap((t) => [t.y, t.y + t.h / 2, t.y + t.h]), 0, E.H / 2, E.H, ...(ed.guides ? allGuides().y : [])];
        const [hx, hy] = { nw: [-1, -1], n: [0, -1], ne: [1, -1], e: [1, 0], se: [1, 1], s: [0, 1], sw: [-1, 1], w: [-1, 0] }[drag.handle];
        if (hx && !keep) {
          const edge = hx > 0 ? next.x + next.w : next.x;
          const snapX = ops.snapValue(edge, xs, tolerance);
          if (snapX != null) { if (hx > 0) next.w = Math.max(4, snapX - next.x); else { next.w = Math.max(4, next.x + next.w - snapX); next.x = snapX; } lines.push({ axis: "x", at: snapX, from: Math.min(next.y, 0), to: Math.max(next.y + next.h, E.H), kind: "object" }); }
        }
        if (hy && !keep) {
          const edge = hy > 0 ? next.y + next.h : next.y;
          const snapY = ops.snapValue(edge, ys, tolerance);
          if (snapY != null) { if (hy > 0) next.h = Math.max(4, snapY - next.y); else { next.h = Math.max(4, next.y + next.h - snapY); next.y = snapY; } lines.push({ axis: "y", at: snapY, from: 0, to: E.W, kind: "object" }); }
        }
      }
      result.list = ops.update(drag.list, [o.id], next);
      result.lines = lines;
      result.extra.push(sizeTip(next, p));
      drag.touched = [o.id];
      return;
    }
    // Several objects: the box around them scales, and everything in it with it.
    const b0 = ops.bounds(chosen, ends(drag.list));
    const nb = ops.resizeBox({ ...b0 }, drag.handle, p, { keepRatio: event.shiftKey, fromCenter });
    const sx = nb.w / (b0.w || 1);
    const sy = nb.h / (b0.h || 1);
    const map = (x, y) => [nb.x + (nb.flipH ? b0.x + b0.w - x : x - b0.x) * sx, nb.y + (nb.flipV ? b0.y + b0.h - y : y - b0.y) * sy];
    result.list = drag.list.map((o) => {
      if (!ed.sel.includes(o.id)) return o;
      if (o.kind === "line") { const [x1, y1] = map(o.x1, o.y1); const [x2, y2] = map(o.x2, o.y2); return { ...o, x1, y1, x2, y2 }; }
      const turned = Math.abs(((o.rot || 0) % 180)) > 45 && Math.abs(((o.rot || 0) % 180)) < 135;
      const w = o.w * (turned ? sy : sx);
      const hh = o.h * (turned ? sx : sy);
      const [cx, cy] = map(o.x + o.w / 2, o.y + o.h / 2);
      return { ...o, x: cx - w / 2, y: cy - hh / 2, w, h: hh, ...(o.kind !== "text" && o.fs && Math.abs(sx - sy) < 0.001 && event.shiftKey ? { fs: Math.max(8, o.fs * sx) } : {}) };
    });
    drag.touched = ed.sel;
  }

  function sizeTip(b, p) {
    const k = scale();
    return h("div", { class: "ed-tip", style: { left: `${p[0] * k + 14}px`, top: `${p[1] * k + 14}px` } }, `幅 ${ops.toCm(b.w)} cm × 高さ ${ops.toCm(b.h)} cm`);
  }

  // -- rotate
  function rotateDrag(drag, p, event, result) {
    const chosen = drag.list.filter((o) => ed.sel.includes(o.id));
    if (!chosen.length) return;
    if (chosen.length === 1 && chosen[0].kind !== "line") {
      const rot = ops.angleTo(chosen[0], p, event.shiftKey);
      result.list = ops.update(drag.list, [chosen[0].id], { rot: rot || undefined });
      result.extra.push(h("div", { class: "ed-tip", style: { left: `${p[0] * scale() + 14}px`, top: `${p[1] * scale() + 14}px` } }, `${Math.round(rot)}°`));
    } else {
      const b = ops.bounds(chosen, ends(drag.list));
      const center = { x: b.x, y: b.y, w: b.w, h: b.h };
      const a0 = drag.a0 ?? (drag.a0 = ops.angleTo(center, drag.from, false));
      let delta = ops.angleTo(center, p, false) - a0;
      if (event.shiftKey) delta = Math.round(delta / 15) * 15;
      // A group turns as one; separate objects each turn about their own centre.
      const grouped = chosen.every((o) => o.group && o.group === chosen[0].group);
      result.list = grouped ? ops.rotateBy(drag.list, ed.sel, delta) : chosen.reduce((list, o) => ops.rotateBy(list, [o.id], delta), drag.list);
    }
    drag.touched = ed.sel;
  }

  // -- adjust (yellow handles)
  function adjustDrag(drag, p, result) {
    const o = byId(ed.sel[0], drag.list);
    if (!o) return;
    const shape = E.SHAPES[o.kind === "image" ? o.mask : o.shape];
    const handle = shape?.handles?.[Number(drag.handle.slice(3))];
    if (!handle) return;
    // Pointer → the shape's own frame (unrotated, unflipped).
    const [lx, ly] = ops.rotatePoint(p, [o.x + o.w / 2, o.y + o.h / 2], -(o.rot || 0));
    let x = lx - o.x;
    let y = ly - o.y;
    if (o.flipH) x = o.w - x;
    if (o.flipV) y = o.h - y;
    const adj = E.adjOf(shape, handle.set(x, y, o.w, o.h, E.adjOf(shape, o.adj)));
    result.list = ops.update(drag.list, [o.id], { adj: adj.map((v) => Math.round(v * 10000) / 10000) });
    drag.touched = [o.id];
  }

  // -- line ends: drag; near a shape's connection point the line attaches to it
  function endDrag(drag, p, event, tolerance, result) {
    const o = byId(ed.sel[0], drag.list);
    if (!o || o.kind !== "line") return;
    const which = drag.handle === "end1" ? 1 : 2;
    const [p1, p2] = E.lineEnds(o, drag.list);
    const other = which === 1 ? p2 : p1;
    let [x, y] = p;
    if (event.shiftKey) {
      const ang = Math.round(Math.atan2(y - other[1], x - other[0]) / (Math.PI / 4)) * (Math.PI / 4);
      const len = Math.hypot(x - other[0], y - other[1]);
      [x, y] = [other[0] + Math.cos(ang) * len, other[1] + Math.sin(ang) * len];
    }
    let ref = null;
    const sitesShown = [];
    for (const t of visible(drag.list)) {
      if (t.kind === "line" || t.id === o.id) continue;
      if (!ops.hits(t, [x, y], 40 / scale())) continue;
      E.sites(t).forEach((s, site) => {
        sitesShown.push(s);
        if (Math.hypot(s[0] - x, s[1] - y) <= Math.max(tolerance * 1.6, 12) && (!ref || Math.hypot(s[0] - x, s[1] - y) < ref.d)) ref = { id: t.id, site, d: Math.hypot(s[0] - x, s[1] - y), at: s };
      });
    }
    const k = scale();
    for (const s of sitesShown) result.extra.push(h("span", { class: ["ed-site", ref && ref.at === s ? "on" : ""], style: { left: `${s[0] * k}px`, top: `${s[1] * k}px` } }));
    const key = which === 1 ? "from" : "to";
    const patch = which === 1 ? { x1: ref ? ref.at[0] : x, y1: ref ? ref.at[1] : y } : { x2: ref ? ref.at[0] : x, y2: ref ? ref.at[1] : y };
    patch[key] = ref ? { id: ref.id, site: ref.site } : undefined;
    // The other end keeps where it is drawn.
    if (which === 1) Object.assign(patch, { x2: p2[0], y2: p2[1] }); else Object.assign(patch, { x1: p1[0], y1: p1[1] });
    result.list = ops.update(drag.list, [o.id], patch);
    drag.touched = [o.id];
  }

  // -- the selection box
  function marqueeDrag(drag, p, result) {
    const box = { x: Math.min(drag.from[0], p[0]), y: Math.min(drag.from[1], p[1]), w: Math.abs(p[0] - drag.from[0]), h: Math.abs(p[1] - drag.from[1]) };
    drag.box = box;
    const k = scale();
    result.extra.push(h("div", { class: "ed-marquee", style: { left: `${box.x * k}px`, top: `${box.y * k}px`, width: `${box.w * k}px`, height: `${box.h * k}px` } }));
  }

  // -- drawing a new object
  function drawBox(drag, p, event) {
    let [x2, y2] = p;
    const [x1, y1] = drag.from;
    if (drag.tool.kind === "line") {
      if (event.shiftKey) {
        const ang = Math.round(Math.atan2(y2 - y1, x2 - x1) / (Math.PI / 4)) * (Math.PI / 4);
        const len = Math.hypot(x2 - x1, y2 - y1);
        [x2, y2] = [x1 + Math.cos(ang) * len, y1 + Math.sin(ang) * len];
      }
      return { x1, y1, x2, y2 };
    }
    let w = Math.abs(x2 - x1);
    let hh = Math.abs(y2 - y1);
    if (event.shiftKey) { const m = Math.max(w, hh); w = m; hh = m; }
    const x = event.ctrlKey || event.metaKey ? x1 - w : x2 < x1 ? x1 - w : x1;
    const y = event.ctrlKey || event.metaKey ? y1 - hh : y2 < y1 ? y1 - hh : y1;
    return event.ctrlKey || event.metaKey ? { x, y, w: w * 2, h: hh * 2 } : { x, y, w, h: hh };
  }
  function drawDrag(drag, p, event, result) {
    const k = scale();
    const box = drawBox(drag, p, event);
    drag.box = box;
    if (drag.tool.kind === "line") {
      const len = Math.hypot(box.x2 - box.x1, box.y2 - box.y1) * k;
      result.extra.push(h("div", { class: "ed-draw-line", style: { left: `${box.x1 * k}px`, top: `${box.y1 * k}px`, width: `${len}px`, transform: `rotate(${Math.atan2(box.y2 - box.y1, box.x2 - box.x1)}rad)` } }));
    } else {
      result.extra.push(h("div", { class: "ed-draw-box", style: { left: `${box.x * k}px`, top: `${box.y * k}px`, width: `${box.w * k}px`, height: `${box.h * k}px` } }), sizeTip(box, p));
    }
  }
  function finishDraw(drag, p, event) {
    const tool = drag.tool;
    let box = drag.moved ? drag.box : null;
    // A click without dragging places the object at its usual size where it was clicked.
    if (!box || (tool.kind !== "line" && (box.w < 6 || box.h < 6)) || (tool.kind === "line" && Math.hypot(box.x2 - box.x1, box.y2 - box.y1) < 6)) {
      const [x, y] = drag.from;
      if (tool.kind === "line") box = { x1: x, y1: y, x2: x + 300, y2: y };
      else if (tool.kind === "text") box = { x, y, w: tool.vertical ? 90 : 520, h: tool.vertical ? 420 : 70 };
      else { const [w, hh] = ops.defaultSize(tool.shape); box = { x: x - w / 2, y: y - hh / 2, w, h: hh }; }
    }
    // 描画モードのロック keeps the tool for the next one (Esc ends it).
    if (!tool.lock) setTool(null);
    const o = newObjectFor(tool, box);
    commit([...objects(), o], { select: [o.id] });
    if (o.kind === "text") requestAnimationFrame(() => startTyping(o.id, { caret: "end" }));
  }

  function newObjectFor(tool, box) {
    // 既定の図形・テキスト ボックス・線に設定 (the deck's own): the look a new one starts with.
    const own = app.objectDefaults?.() || {};
    if (tool.kind === "line") {
      const extra = { line: {}, arrow: { tail: "triangle" }, double: { head: "triangle", tail: "triangle" }, elbow: { route: "elbow", tail: "triangle" }, curve: { route: "curve", tail: "triangle" } }[tool.variant || "line"] || {};
      return ops.makeObject("line", box, { ...(own.line || {}), ...extra });
    }
    if (tool.kind === "text") return ops.makeObject("text", box, { ...(own.text || {}), text: "<p><br></p>", ...(tool.vertical ? { vertical: true, autofit: "none" } : {}) });
    // 動作設定ボタン come with their click (次へ・前へ・最初へ…) set, as in PowerPoint.
    const preset = E.SHAPES[tool.shape]?.action;
    return ops.makeObject("shape", box, { ...(own.shape || {}), shape: tool.shape || "rect", ...(preset ? { action: { ...preset } } : {}) });
  }

  function setTool(tool) {
    if (ed.tool?.lock && !tool) app.toast("描画モードのロックを解除しました");
    ed.tool = tool;
    ed.wrap?.classList.toggle("ed-drawing", Boolean(tool));
    emit();
  }

  // ---------------------------------------------------------------- double click, typing

  function onDoubleClick(event) {
    if (!ed.slideEl || ed.tool) return;
    const list = objects();
    const hit = objectAt(event, list);
    if (!hit) return;
    event.preventDefault();
    // In a group, the first double-click selects the one object.
    if (hit.group && ed.entered !== hit.group) {
      ed.entered = hit.group;
      ed.sel = [hit.id];
      draw();
      emit();
      if (!["shape", "text"].includes(hit.kind)) return;
    }
    if (["shape", "text"].includes(hit.kind) && !hit.locked) startTyping(hit.id, { at: [event.clientX, event.clientY] });
    else if (hit.kind === "table" && !hit.locked) {
      const td = cellAt(event, hit.id);
      startTyping(hit.id, { cell: td ? [Number(td.dataset.r), Number(td.dataset.c)] : [0, 0], at: [event.clientX, event.clientY] });
    } else if (hit.kind === "image") app.showTab?.("picture");
    else if (hit.kind === "chart") app.editChart?.(hit.id);
    else if (hit.kind === "equation" && !hit.locked) app.editEquation?.(hit.id);
    else if (hit.kind === "poll" && !hit.locked) app.editPoll?.(hit.id);
    else if (hit.kind === "zoom") app.showTab?.("zoomTool");
    else if (hit.kind === "ink") app.showTab?.("draw");
    else if (hit.kind === "smartart" && !hit.locked) {
      // The item under the pointer opens in the SmartArt's テキスト ウィンドウ.
      const node = document.elementsFromPoint(event.clientX, event.clientY).map((el) => el.closest?.("[data-item]")).find((el) => el?.closest?.(`[data-el="${hit.id}"]`));
      const item = Number(node?.dataset.item);
      app.editSmartart?.(hit.id, Number.isFinite(item) ? item : null);
    }
  }

  /** A block of cells picked while typing in a table (rows r0…r1, columns c0…c1), shown tinted on the stage. */
  function setCellRange(range) {
    ed.cellRange = range && (range[0] !== range[2] || range[1] !== range[3]) ? range : null;
    paintCellRange();
    emit();
  }
  function paintCellRange() {
    const node = ed.typing?.node || (ed.sel.length === 1 ? ed.slideEl?.querySelector(`.hs-obj[data-el="${CSS.escape(ed.sel[0])}"]`) : null);
    for (const td of ed.slideEl?.querySelectorAll("td.ed-cell-sel") || []) td.classList.remove("ed-cell-sel");
    const r = ed.cellRange;
    if (!r || !node) return;
    const [ra, rb, ca, cb] = [Math.min(r[0], r[2]), Math.max(r[0], r[2]), Math.min(r[1], r[3]), Math.max(r[1], r[3])];
    for (const td of node.querySelectorAll("td[data-r]")) {
      const i = Number(td.dataset.r);
      const j = Number(td.dataset.c);
      if (i >= ra && i <= rb && j >= ca && j <= cb) td.classList.add("ed-cell-sel");
    }
  }
  function dragCells(event, from) {
    const id = ed.typing.id;
    const move = (ev) => {
      if (!ed.typing || ed.typing.id !== id) return;
      const td = cellAt(ev, id);
      if (!td) return;
      const here = [Number(td.dataset.r), Number(td.dataset.c)];
      if (here[0] === from[0] && here[1] === from[1]) { if (ed.cellRange) setCellRange(null); return; }
      window.getSelection()?.removeAllRanges();
      setCellRange([from[0], from[1], here[0], here[1]]);
    };
    const up = () => { window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", up); };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  }
  /** Where table commands act: the picked block, the cell being typed in, or (table selected) every cell. */
  function tableTarget() {
    const id = ed.typing?.cell ? ed.typing.id : ed.sel.length === 1 ? ed.sel[0] : null;
    const o = id ? byId(id) : null;
    if (!o || o.kind !== "table") return null;
    if (ed.cellRange) { const [r0, c0, r1, c1] = ed.cellRange; return { o, r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1), cells: true }; }
    if (ed.typing?.cell) { const [r, c] = ed.typing.cell; const cell = o.cells[r][c]; return { o, r0: r, c0: c, r1: r + (cell.rs || 1) - 1, c1: c + (cell.cs || 1) - 1, cells: true }; }
    return { o, r0: 0, c0: 0, r1: o.cells.length - 1, c1: o.cells[0].length - 1, cells: false };
  }
  /** Change the table that table commands act on (typing stops first; the cell picked stays picked). */
  function changeTable(fn, { keepTyping = true, keepRange = true } = {}) {
    const target = tableTarget();
    if (!target) return null;
    const cell = ed.typing?.cell;
    const range = ed.cellRange;
    if (ed.typing) stopTyping(true);
    const o = byId(target.o.id);
    const next = fn(o, target);
    if (!next) return null;
    commit(objects().map((item) => (item.id === o.id ? next : item)), { select: [o.id] });
    if (keepTyping && cell && next.cells?.[cell[0]]?.[cell[1]] && !next.cells[cell[0]][cell[1]].merged) { startTyping(o.id, { cell, caret: "end" }); if (range && keepRange) setCellRange(range); }
    return next;
  }
  /** A table whose rows no longer hold their words grows to fit (measured on the stage). */
  function tableFit(o, node) {
    const trs = [...(node?.querySelectorAll("tr") || [])];
    if (!trs.length || trs.length !== o.cells.length) return {};
    const k = scale();
    const heights = trs.map((tr) => tr.getBoundingClientRect().height / k);
    const total = heights.reduce((a, b) => a + b, 0);
    if (total <= o.h + 1) return {};
    return { h: Math.round(total * 100) / 100, rows: heights.map((v) => Math.round((v / total) * 100000) / 100000) };
  }

  /** The table cell under the pointer (a <td> of that table on the stage). */
  function cellAt(event, id) {
    const node = ed.slideEl?.querySelector(`.hs-obj[data-el="${CSS.escape(id)}"]`);
    return document.elementsFromPoint(event.clientX, event.clientY).map((el) => el.closest?.("td[data-r]")).find((td) => td && node?.contains(td)) || null;
  }

  function startTyping(id, { caret = "point", at = null, replaceWith = null, cell = null } = {}) {
    const o = byId(id);
    if (!o || !["shape", "text", "table"].includes(o.kind) || !ed.slideEl) return false;
    const node = ed.slideEl.querySelector(`.hs-obj[data-el="${CSS.escape(id)}"]`);
    let tx = node?.querySelector(".hs-obj-tx");
    let where = null;
    if (o.kind === "table") {
      // A table types into one cell at a time (the cell that covers a merged block).
      where = ops.tableOrigin(o, ...(cell || [0, 0]));
      tx = node?.querySelector(`td[data-r="${where[0]}"][data-c="${where[1]}"] .hs-cell-tx`);
      ed.cellRange = null;
      // Excel スプレッドシート: a formula cell shows its formula while it is edited (what it works out to shows otherwise).
      const own = o.sheet ? o.cells[where[0]]?.[where[1]] : null;
      if (own?.f && tx) tx.innerHTML = richOf(own.f);
    }
    if (!tx) return false;
    // 段落 → タブ設定: the words are edited as they were typed (tab characters), not as the boxes they are drawn in.
    if (o.kind !== "table" && /\t/.test(o.text || "")) tx.replaceChildren(E.richFragment(o.text));
    tx.classList.toggle("ed-has-tab", /\t/.test(tx.textContent));
    ed.sel = [id];
    ed.typing = { id, before: tx.innerHTML, node, tx, cell: where };
    if (!tx.innerHTML.trim()) tx.innerHTML = "<p><br></p>";
    if (replaceWith != null) tx.innerHTML = `<p>${replaceWith.replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[c]) || "<br>"}</p>`;
    tx.setAttribute("contenteditable", "true");
    // 校閲 → スペル チェック: the browser's own dictionary underlines while typing.
    tx.spellcheck = Boolean(app.spellcheck?.());
    tx.classList.add("ed-typing-tx");
    node.classList.add("ed-typing");
    try { document.execCommand("defaultParagraphSeparator", false, "p"); document.execCommand("styleWithCSS", false, true); } catch { /* old browser */ }
    tx.focus({ preventScroll: true });
    const selection = window.getSelection();
    let range = null;
    if (at && document.caretRangeFromPoint) range = document.caretRangeFromPoint(at[0], at[1]);
    if (!range || !tx.contains(range.startContainer)) {
      range = document.createRange();
      range.selectNodeContents(tx);
      if (caret !== "all") range.collapse(false);
    }
    selection.removeAllRanges();
    selection.addRange(range);
    tx.addEventListener("input", onTypingInput);
    tx.addEventListener("keydown", onTypingKey);
    tx.addEventListener("paste", onTypingPaste);
    tx.addEventListener("focusout", onTypingBlur);
    document.addEventListener("selectionchange", rememberRange);
    draw();
    emit();
    return true;
  }

  function rememberRange() {
    const selection = window.getSelection();
    if (ed.typing && selection.rangeCount && ed.typing.tx.contains(selection.anchorNode)) ed.range = selection.getRangeAt(0).cloneRange();
  }
  /** Put the caret back where it was (after a ribbon control took the focus). */
  function restoreRange() {
    if (!ed.typing) return false;
    ed.typing.tx.focus({ preventScroll: true });
    if (ed.range) { const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(ed.range); }
    return true;
  }

  /**
   * A list made while typing in a paragraph lands inside it (<p><ol>…</ol></p>, which HTML does not allow and the
   * sanitizer splits into empty paragraphs): the list is lifted out, the caret staying at the end of its item.
   */
  function unnestLists() {
    const t = ed.typing;
    if (!t) return;
    const selection = window.getSelection();
    const range = selection.rangeCount ? selection.getRangeAt(0) : null;
    const at = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement);
    const item = at?.closest?.("li");
    let moved = false;
    for (const list of t.tx.querySelectorAll("p > ul, p > ol, div > ul, div > ol")) {
      const holder = list.parentElement;
      if (holder === t.tx || !t.tx.contains(holder)) continue;
      const others = [...holder.childNodes].filter((n) => n !== list && n.nodeName !== "BR" && !(n.nodeType === 3 && !n.textContent.trim()));
      if (others.length) holder.after(list); else holder.replaceWith(list);
      moved = true;
    }
    if (moved && item?.isConnected) {
      const caret = document.createRange();
      // An empty item (just its line break): before the break, or the first letters typed would go on a new line.
      if (item.textContent) { caret.selectNodeContents(item); caret.collapse(false); } else caret.setStart(item, 0);
      selection.removeAllRanges();
      selection.addRange(caret);
    }
  }
  /** 下線の種類・二重取り消し線 on the words chosen while typing: underlined (struck) first, then the line's kind. */
  function markLine(tag, value) {
    const t = ed.typing;
    const selection = window.getSelection();
    if (!t || !selection.rangeCount || selection.isCollapsed) return;
    const command = tag === "u" ? "underline" : "strikeThrough";
    if (!document.queryCommandState(command)) document.execCommand(command);
    const range = selection.getRangeAt(0);
    for (const el of t.tx.querySelectorAll(tag)) {
      if (!range.intersectsNode(el)) continue;
      if (value && value !== "single") el.setAttribute("data-line", value); else el.removeAttribute("data-line");
    }
  }
  /** The paragraphs the caret or the selection is in (while typing). */
  function typedBlocks() {
    const t = ed.typing;
    const selection = window.getSelection();
    if (!t || !selection.rangeCount) return [];
    const range = selection.getRangeAt(0);
    const blocks = [...t.tx.querySelectorAll("p, div, li")].filter((b) => range.intersectsNode(b) && !b.querySelector("p, div, li"));
    if (!blocks.length) {
      const node = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
      const block = node?.closest?.("p, div, li");
      if (block && block !== t.tx && t.tx.contains(block)) blocks.push(block);
    }
    return blocks;
  }
  function onTypingInput(event) {
    if (event?.inputType === "insertText" && event.data && !event.isComposing) runAutoCorrect(event.data);
    growWhileTyping();
  }
  /**
   * オートコレクト: what was just typed may finish "(c)" (→ ©) or a list marker at a paragraph's start ("・ ", "1. ").
   * The correction is typed like any other text, so ⌘Z brings back what was typed.
   */
  function runAutoCorrect(typed) {
    const t = ed.typing;
    const selection = window.getSelection();
    if (!t || !selection.rangeCount || !selection.isCollapsed) return;
    const range = selection.getRangeAt(0);
    const at = range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement;
    if (!at || !t.tx.contains(at)) return;
    const block = at.closest("p, div, li");
    const root = block && block !== t.tx && t.tx.contains(block) ? block : t.tx;
    const head = document.createRange();
    head.setStart(root, 0);
    head.setEnd(range.startContainer, range.startOffset);
    // A formula in a sheet is left as typed ("<--" or "(c)" in it is not meant as an arrow or a ©).
    if (t.cell && /^\s*[=＝]/.test(head.toString()) && byId(t.id)?.sheet) return;
    const fix = autoCorrect(head.toString(), typed, { inList: Boolean(at.closest("li")), options: autoCorrectOptions() });
    if (!fix) return;
    const back = (n) => { for (let i = 0; i < n; i += 1) selection.modify("extend", "backward", "character"); };
    if (fix.replace) {
      back(fix.replace);
      document.execCommand("insertText", false, fix.with);
      return;
    }
    back(fix.remove);
    document.execCommand("delete");
    document.execCommand(fix.list === "bullet" ? "insertUnorderedList" : "insertOrderedList");
    unnestLists();
    const node = selection.anchorNode;
    const list = (node?.nodeType === 1 ? node : node?.parentElement)?.closest?.("ul, ol");
    if (list && t.tx.contains(list) && fix.style !== (fix.list === "bullet" ? "disc" : "decimal")) list.setAttribute("data-style", fix.style);
  }
  /** Whether Tab indents here: in a list, or with nothing of the paragraph before the caret (and nothing selected). */
  function tabIndents() {
    const sel = window.getSelection();
    if (!sel?.rangeCount || !ed.typing?.tx.contains(sel.anchorNode)) return true;
    const range = sel.getRangeAt(0);
    const node = range.startContainer.nodeType === 3 ? range.startContainer.parentElement : range.startContainer;
    if (node?.closest("li")) return true;
    if (!range.collapsed) return false;
    const block = node?.closest("p, div") || ed.typing.tx;
    const before = document.createRange();
    before.selectNodeContents(block);
    before.setEnd(range.startContainer, range.startOffset);
    return before.toString() === "";
  }
  function onTypingKey(event) {
    // F4 (繰り返し) and ⇧F7 (類義語) work while typing too: the box's own handler stops the key from reaching the page's.
    if ((event.key === "F4" || event.key === "F7") && keydown(event)) { event.stopPropagation(); return; }
    if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); stopTyping(true); return; }
    if (event.key === "Tab" && ed.typing?.cell) {
      // Tab: the next cell (a new row after the last one), Shift+Tab: the one before.
      event.preventDefault();
      event.stopPropagation();
      const { id, cell: [r, c] } = ed.typing;
      stopTyping(true);
      const o = byId(id);
      const order = [];
      o.cells.forEach((row, i) => row.forEach((x, j) => { if (!x.merged) order.push([i, j]); }));
      const at = order.findIndex(([i, j]) => i === r && j === c);
      let to = order[at + (event.shiftKey ? -1 : 1)];
      if (!to && !event.shiftKey) {
        commit(objects().map((item) => (item.id === id ? ops.tableInsertRow(item, item.cells.length) : item)));
        to = [byId(id).cells.length - 1, 0];
      }
      if (to) startTyping(id, { cell: to, caret: "all" });
      return;
    }
    // Excel スプレッドシート: Enter ends the cell and goes to the one below (⇧Enter above); ⌥Enter is a new line.
    if (event.key === "Enter" && ed.typing?.cell && !event.altKey && !event.isComposing && byId(ed.typing.id)?.sheet) {
      event.preventDefault();
      event.stopPropagation();
      const { id, cell: [r, c] } = ed.typing;
      stopTyping(true);
      const next = Math.max(0, Math.min(byId(id).cells.length - 1, r + (event.shiftKey ? -1 : 1)));
      startTyping(id, { cell: [next, c], caret: "all" });
      return;
    }
    if (event.key === "Tab") {
      event.preventDefault();
      // At the start of a paragraph (or in a list) Tab indents; anywhere else it types a tab, which falls on the paragraph's tab stops.
      if (event.shiftKey || tabIndents()) document.execCommand(event.shiftKey ? "outdent" : "indent");
      else { document.execCommand("insertText", false, "\t"); ed.typing?.tx.classList.add("ed-has-tab"); }
    }
    event.stopPropagation();
  }
  function onTypingPaste(event) {
    // Text only, as the box's own style (rich text from other apps would bring their colours and fonts).
    event.preventDefault();
    const text = event.clipboardData?.getData("text/plain") ?? "";
    // Cells from Excel (or a table of a web page) in a table cell fill the cells from there on, as PowerPoint does.
    const grid = ed.typing?.cell ? ops.clipboardGrid({ text, html: event.clipboardData?.getData("text/html") || "" }) : null;
    if (grid) {
      const [r, c] = ed.typing.cell;
      changeTable((o) => ops.tableFill(o, r, c, grid));
      app.toast(`セルに貼り付けました（${grid.rows.length}行×${grid.rows[0].length}列。足りない行・列は増やしました）`);
      return;
    }
    document.execCommand("insertText", false, text);
  }
  function onTypingBlur(event) {
    // A ribbon or pane control that keeps the caret (data-keeps-text) does not end typing.
    if (event.relatedTarget?.closest?.("[data-keeps-text], .ed-menu")) return;
    setTimeout(() => { if (ed.typing && !ed.typing.tx.contains(document.activeElement) && !document.activeElement?.closest?.("[data-keeps-text]")) stopTyping(true); }, 0);
  }

  /** The box height a "grow" object needs for its text. */
  function neededSize(o, tx) {
    const g = E.geometry(o.shape || "rect", o.w, o.h, o.adj, o.path);
    const pad = o.pad || E.OBJECT_DEFAULTS[o.kind]?.pad || [7, 14, 7, 14];
    if (o.vertical) {
      const insetX = o.w - (g.text[2] - g.text[0]);
      return { w: Math.max(30, Math.ceil(tx.scrollWidth + pad[1] + pad[3] + insetX)) };
    }
    const insetY = o.h - (g.text[3] - g.text[1]);
    const out = { h: Math.max(20, Math.ceil(tx.scrollHeight + pad[0] + pad[2] + insetY)) };
    if (o.wrap === false) out.w = Math.max(30, Math.ceil(tx.scrollWidth + pad[1] + pad[3] + (o.w - (g.text[2] - g.text[0]))));
    return out;
  }
  function growWhileTyping() {
    const t = ed.typing;
    if (!t) return;
    const o = { ...E.withDefaults(byId(t.id)) };
    if (o.autofit !== "grow") return;
    const size = neededSize(o, t.tx);
    const box = t.node.querySelector(".hs-obj-text");
    if (size.h != null && Math.abs(size.h - o.h) > 0.5) {
      t.node.style.height = `${size.h}px`;
      if (box) box.style.height = `${Math.max(0, parseFloat(box.style.height) + size.h - o.h)}px`;
      t.size = size;
    }
    if (size.w != null && Math.abs(size.w - o.w) > 0.5) { t.node.style.width = `${size.w}px`; if (box) box.style.width = `${Math.max(0, parseFloat(box.style.width) + size.w - o.w)}px`; t.size = { ...(t.size || {}), w: size.w }; }
    const live = { ...byId(t.id), ...(t.size || {}) };
    draw(objects().map((item) => (item.id === t.id ? live : item)));
  }

  function stopTyping(save = true) {
    const t = ed.typing;
    if (!t) return;
    ed.typing = null;
    ed.range = null;
    ed.cellRange = null;
    document.removeEventListener("selectionchange", rememberRange);
    t.tx.removeEventListener("input", onTypingInput);
    t.tx.removeEventListener("keydown", onTypingKey);
    t.tx.removeEventListener("paste", onTypingPaste);
    t.tx.removeEventListener("focusout", onTypingBlur);
    t.tx.removeAttribute("contenteditable");
    t.tx.classList.remove("ed-typing-tx");
    t.node.classList.remove("ed-typing");
    const html = E.sanitizeRich(t.tx.innerHTML);
    const o = byId(t.id);
    if (!o) return;
    const empty = !html.replace(/<[^>]+>/g, "").trim();
    if (t.cell) {
      // A cell keeps its words; a table that no longer fits its text grows (rows keep their proportions).
      // In a sheet, "=…" typed in a cell is a formula (and typing words over a formula takes it away).
      const own = o.sheet ? o.cells[t.cell[0]]?.[t.cell[1]] : null;
      const words = own ? E.richToText(html) : "";
      const formula = own && (own.f || /^\s*=/.test(words));
      const patch = formula ? (cell) => cellPatch(cell, words) : { text: empty ? undefined : html };
      let next = save && html !== E.sanitizeRich(t.before) ? ops.tableCells(o, t.cell[0], t.cell[1], t.cell[0], t.cell[1], patch) : o;
      next = { ...next, ...tableFit(next, t.node) };
      if (next !== o && JSON.stringify(next) !== JSON.stringify(o)) commit(objects().map((item) => (item.id === o.id ? next : item)));
      else { app.rerender(); emit(); }
      return;
    }
    // An empty text box disappears when you leave it, as in PowerPoint.
    if (save && empty && o.kind === "text") {
      commit(objects().filter((item) => item.id !== o.id), { select: [] });
      return;
    }
    const before = E.sanitizeRich(t.before);
    const patch = {};
    if (save && html !== before) patch.text = empty ? undefined : html;
    const grown = E.withDefaults(o).autofit === "grow" ? neededSize({ ...E.withDefaults(o) }, t.tx) : {};
    if (grown.h != null && Math.abs(grown.h - o.h) > 0.5) patch.h = grown.h;
    if (grown.w != null && Math.abs(grown.w - o.w) > 0.5) patch.w = grown.w;
    if (Object.keys(patch).length) commit(ops.update(objects(), [o.id], patch));
    else { app.rerender(); emit(); }
  }

  /** Boxes set to grow with their text keep their height in step (after a font change, a new size…). */
  function syncGrow() {
    if (ed.typing || !ed.slideEl) return;
    const fixes = [];
    for (const o of objects()) {
      if (o.kind === "table") {
        const fit = tableFit(o, ed.slideEl.querySelector(`.hs-obj[data-el="${CSS.escape(o.id)}"]`));
        if (fit.h != null) fixes.push([o.id, fit]);
        continue;
      }
      const full = E.withDefaults(o);
      if (full.autofit !== "grow" || !["shape", "text"].includes(o.kind)) continue;
      const tx = ed.slideEl.querySelector(`.hs-obj[data-el="${CSS.escape(o.id)}"] .hs-obj-tx`);
      if (!tx) continue;
      const size = neededSize(full, tx);
      const patch = {};
      if (size.h != null && Math.abs(size.h - o.h) > 1) patch.h = size.h;
      if (size.w != null && Math.abs(size.w - o.w) > 1) patch.w = size.w;
      if (Object.keys(patch).length) fixes.push([o.id, patch]);
    }
    if (!fixes.length) return;
    let list = objects();
    for (const [id, patch] of fixes) list = ops.update(list, [id], patch);
    // Not an undo step of its own: it follows from the change just made.
    queueMicrotask(() => commit(list, { undo: false }));
  }

  // ---------------------------------------------------------------- keyboard

  function keydown(event) {
    // ⇧⌘V pastes the words only (the paste event that follows reads this).
    if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "v") ed.plainPaste = event.shiftKey;
    if (!app.canEdit()) return false;
    const meta = event.ctrlKey || event.metaKey;
    const key = event.key;
    // ⇧F7: 類義語 (also while typing: the words selected are looked up, and the pick takes their place).
    if (key === "F7" && event.shiftKey && !meta && !event.altKey) { event.preventDefault(); app.openSynonyms?.(); return true; }
    // F4: 繰り返し (also while typing: the format goes to the words selected now).
    if (key === "F4" && !meta && !event.altKey && !event.shiftKey) { event.preventDefault(); repeat(); return true; }
    if (ed.typing) return false;
    const active = document.activeElement;
    if (active && (/^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName) || active.isContentEditable) && !active.closest?.(".ed-wrap")) return false;
    // A pane with keys of its own (the thumbnails: Delete deletes the slide there).
    if (active?.closest?.("[data-own-keys]")) return false;
    if (key === "Escape") {
      if (ed.menu) { closeMenu(); return true; }
      if (ed.tool) { setTool(null); return true; }
      if (ed.painter) { ed.painter = null; app.toast("書式のコピーを終えました"); emit(); return true; }
      if (ed.picker) { pick(null); return true; }
      if (ed.entered) { const g = ed.entered; ed.entered = null; ed.sel = objects().filter((o) => o.group === g).map((o) => o.id); draw(); emit(); return true; }
      if (ed.sel.length) { ed.sel = []; draw(); emit(); return true; }
      return false;
    }
    if (meta && key.toLowerCase() === "a" && objects().length) { event.preventDefault(); ed.sel = visible().map((o) => o.id); draw(); emit(); return true; }
    if (!ed.sel.length) return false;
    const list = objects();
    if (key === "Delete" || key === "Backspace") { event.preventDefault(); removeSelection(); return true; }
    if (/^Arrow/.test(key)) {
      event.preventDefault();
      const step = meta || event.altKey ? NUDGE.fine : event.shiftKey ? NUDGE.big : NUDGE.plain;
      const [dx, dy] = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[key];
      const ids = ed.sel.filter((id) => !byId(id, list)?.locked);
      if (ids.length) commit(ops.moveBy(list, ids, dx, dy));
      return true;
    }
    if (key === "Tab") {
      event.preventDefault();
      const order = visible(list);
      if (!order.length) return true;
      const at = order.findIndex((o) => o.id === ed.sel[ed.sel.length - 1]);
      const next = order[(at + (event.shiftKey ? -1 : 1) + order.length) % order.length];
      ed.sel = [next.id];
      ed.entered = next.group || null;
      draw();
      emit();
      return true;
    }
    if ((key === "Enter" || key === "F2") && ed.sel.length === 1) {
      const o = byId(ed.sel[0], list);
      if (o && ["shape", "text"].includes(o.kind)) { event.preventDefault(); startTyping(o.id, { caret: "all" }); return true; }
    }
    if (meta && key.toLowerCase() === "d") { event.preventDefault(); duplicateSelection(); return true; }
    if (meta && key.toLowerCase() === "g") { event.preventDefault(); if (event.shiftKey) ungroupSelection(); else groupSelection(); return true; }
    if (meta && event.shiftKey && (key === "C" || key === "c")) { event.preventDefault(); copyFormat(); return true; }
    if (meta && event.shiftKey && (key === "V" || key === "v")) { event.preventDefault(); pasteFormat(ed.sel); return true; }
    if (meta && event.shiftKey && (key === "}" || key === "]")) { event.preventDefault(); order("front"); return true; }
    if (meta && event.shiftKey && (key === "{" || key === "[")) { event.preventDefault(); order("back"); return true; }
    if (meta && !event.shiftKey && key === "]") { event.preventDefault(); order("forward"); return true; }
    if (meta && !event.shiftKey && key === "[") { event.preventDefault(); order("backward"); return true; }
    if (meta && ["b", "i", "u"].includes(key.toLowerCase()) && !event.shiftKey) { event.preventDefault(); textFormat({ b: "bold", i: "italic", u: "underline" }[key.toLowerCase()]); return true; }
    if (meta && event.shiftKey && (key === ">" || key === ".")) { event.preventDefault(); textFormat("grow"); return true; }
    if (meta && event.shiftKey && (key === "<" || key === ",")) { event.preventDefault(); textFormat("shrink"); return true; }
    if (meta && ["l", "e", "r", "j"].includes(key.toLowerCase()) && !event.shiftKey) { event.preventDefault(); textFormat("align", { l: "left", e: "center", r: "right", j: "justify" }[key.toLowerCase()]); return true; }
    if (meta && event.shiftKey && key.toLowerCase() === "j") { event.preventDefault(); textFormat("align", "distributed"); return true; }
    // Typing a letter on a selected shape replaces its text, as in PowerPoint.
    if (!meta && !event.altKey && key.length === 1 && ed.sel.length === 1) {
      const o = byId(ed.sel[0], list);
      if (o && ["shape", "text"].includes(o.kind) && !o.locked) {
        event.preventDefault();
        startTyping(o.id, { replaceWith: key, caret: "end" });
        growWhileTyping();
        return true;
      }
    }
    return false;
  }

  // ---------------------------------------------------------------- clipboard

  const CLIP_TYPE = "application/x-hpptx-objects";
  // A copy carries the objects' animations ({ objects, anims }); an older copy is a plain list of objects.
  const payloadObjects = (payload) => (Array.isArray(payload) ? payload : Array.isArray(payload?.objects) ? payload.objects : []);
  function copyPayload() {
    const list = objects();
    const chosen = selected(list);
    if (!chosen.length) return null;
    // Connectors keep their drawn ends when copied without the objects they were attached to.
    const ids = new Set(chosen.map((o) => o.id));
    const groups = new Set(chosen.map((o) => o.group).filter(Boolean));
    const anims = timeline().filter((e) => ids.has(e.el) || (e.el?.startsWith("grp:") && groups.has(e.el.slice(4)))).map((e) => JSON.parse(JSON.stringify(e)));
    return { objects: copyObjects(list, chosen, ids), anims };
  }
  function copyObjects(list, chosen, ids) {
    return chosen.map((o) => {
      if (o.kind !== "line") return JSON.parse(JSON.stringify(o));
      const [[x1, y1], [x2, y2]] = E.lineEnds(o, list);
      const copy = { ...JSON.parse(JSON.stringify(o)), x1, y1, x2, y2 };
      if (copy.from && !ids.has(copy.from.id)) delete copy.from;
      if (copy.to && !ids.has(copy.to.id)) delete copy.to;
      return copy;
    });
  }
  /** The copied objects' animations, pointed at the copies (appended after the slide's own). */
  function copiedAnims(anims, idMap, groupMap) {
    const out = [];
    for (const e of anims || []) {
      const el = e.el?.startsWith("grp:") ? (groupMap.has(e.el.slice(4)) ? `grp:${groupMap.get(e.el.slice(4))}` : null) : idMap.get(e.el);
      if (!el) continue;
      const copy = { ...JSON.parse(JSON.stringify(e)), id: ops.newId().replace(/^o/, "a"), el };
      if (copy.trigger) {
        const [tid, mark] = String(copy.trigger).split("@");
        if (idMap.has(tid)) copy.trigger = mark ? `${idMap.get(tid)}@${mark}` : idMap.get(tid); else delete copy.trigger;
      }
      out.push(copy);
    }
    return out;
  }
  function onCopy(event, cut = false) {
    if (!app.canEdit() || ed.typing || !ed.sel.length) return false;
    const active = document.activeElement;
    if (active && (/^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName) || active.isContentEditable)) return false;
    const payload = copyPayload();
    if (!payload) return false;
    ed.clipboard = payload;
    ed.pasteCount = 0;
    try { localStorage.setItem("hpptx-editor-clipboard", JSON.stringify(payload)); } catch { /* too big */ }
    const text = payload.objects.map((o) => E.objectText(o)).filter(Boolean).join("\n");
    event.clipboardData?.setData(CLIP_TYPE, JSON.stringify(payload));
    event.clipboardData?.setData("text/plain", text || " ");
    event.preventDefault();
    if (cut) removeSelection();
    app.toast(cut ? "切り取りました" : `${payload.objects.length}個のオブジェクトをコピーしました`);
    return true;
  }
  async function onPaste(event) {
    if (!app.canEdit() || ed.typing) return false;
    const active = document.activeElement;
    if (active && (/^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName) || active.isContentEditable)) return false;
    const data = event.clipboardData;
    const files = [...(data?.files || [])].filter((file) => /^(image|video)\//.test(file.type));
    const own = data?.getData(CLIP_TYPE);
    // ⇧⌘V: テキストのみ保持 (the words of copied objects as a plain text box).
    const plain = ed.plainPaste;
    ed.plainPaste = false;
    if (own) { event.preventDefault(); pasteObjects(JSON.parse(own), { mode: plain ? "text" : "keep" }); return true; }
    if (files.length) { event.preventDefault(); await app.insertFiles(files); return true; }
    const text = data?.getData("text/plain");
    // Cells from Excel (tab-separated) or a table of a web page: a table, as PowerPoint pastes them.
    const grid = plain ? null : ops.clipboardGrid({ text: text || "", html: data?.getData("text/html") || "" });
    if (grid) {
      event.preventDefault();
      const o = { ...ops.tableFromGrid(grid), id: ops.newId() };
      commit([...objects(), o], { select: [o.id] });
      app.showTab?.("tableDesign");
      app.toast(`表として貼り付けました（${grid.rows.length}行×${grid.rows[0].length}列）`);
      return true;
    }
    if (text && text.trim()) {
      event.preventDefault();
      const lines = text.replace(/\r/g, "").split("\n").slice(0, 60);
      const o = ops.makeObject("text", { x: 200, y: 260, w: 900, h: 70 }, { text: E.textToRich(lines.join("\n")) });
      commit([...objects(), o], { select: [o.id] });
      app.toast("テキストボックスとして貼り付けました");
      return true;
    }
    return false;
  }
  /**
   * 貼り付けオプション: "keep" (元の書式を保持), "theme" (貼り付け先のスタイルを使用: the colours moved to the deck's
   * palette) or "text" (テキストのみ保持: the words only, in a new text box).
   */
  function pasteObjects(payload, { at = null, mode = "keep" } = {}) {
    let pasted = payloadObjects(payload);
    if (!pasted.length) return;
    if (mode === "picture") { pasteAsPicture(pasted, at); return; }
    if (mode === "text") {
      const words = pasted.map((o) => E.objectText(o)).filter(Boolean).join("\n").trim();
      if (!words) { app.toast("貼り付ける文字がありません"); return; }
      const o = ops.makeObject("text", { x: 200, y: 260, w: 900, h: 70 }, { text: E.textToRich(words.split("\n").slice(0, 60).join("\n")) });
      commit([...objects(), o], { select: [o.id] });
      app.toast("テキストのみ貼り付けました");
      return;
    }
    if (mode === "theme") pasted = pasted.map((o) => ops.snapToPalette(o, E.PALETTE));
    ed.pasteCount += 1;
    let offset = ed.pasteCount * 20;
    const list = objects();
    // Pasted on another slide, the objects keep their place; pasted again, each copy steps down a little.
    if (!list.some((o) => pasted.some((p) => p.id === o.id))) offset = (ed.pasteCount - 1) * 20;
    const { list: next, ids, idMap, groupMap } = ops.paste(list, pasted, offset);
    let out = next;
    if (at) {
      const b = ops.bounds(out.filter((o) => ids.includes(o.id)), ends(out));
      out = ops.moveBy(out, ids, at[0] - (b.x + b.w / 2), at[1] - (b.y + b.h / 2));
    }
    const anims = copiedAnims(payload?.anims, idMap, groupMap);
    commit(out, { select: ids, ...(anims.length ? { timeline: [...timeline(), ...anims] } : {}) });
  }
  /** 図として貼り付け: the copied objects drawn into one PNG picture, put where they were (or at `at`). */
  async function pasteAsPicture(pasted, at) {
    let made;
    try { made = await app.objectsPicture?.(pasted); } catch (error) { app.toast(`図として貼り付けられませんでした（${error.message || error}）`); return; }
    if (!made) return;
    const src = await app.storeBlob(made.blob, "図.png");
    const words = pasted.map((o) => E.objectText(o)).filter(Boolean).join(" ").replace(/\s+/g, " ").trim().slice(0, 200);
    ed.pasteCount += 1;
    const list = objects();
    const step = list.some((o) => pasted.some((x) => x.id === o.id)) ? ed.pasteCount * 20 : (ed.pasteCount - 1) * 20;
    const box = at ? { x: at[0] - made.w / 2, y: at[1] - made.h / 2, w: made.w, h: made.h } : { x: made.x + step, y: made.y + step, w: made.w, h: made.h };
    const o = ops.makeObject("image", box, { src, fileName: "図.png", ...(words ? { alt: words } : {}) });
    commit([...list, o], { select: [o.id] });
    app.toast("図として貼り付けました（文字や形はもう直せません。元の部品はコピーしたまま）");
  }
  function pasteFromMemory(mode = "keep") {
    let payload = ed.clipboard;
    if (!payload) { try { payload = JSON.parse(localStorage.getItem("hpptx-editor-clipboard") || "null"); } catch { payload = null; } }
    if (payload) pasteObjects(payload, { mode });
    else app.toast("貼り付けるオブジェクトがありません（⌘V で画像や文字も貼り付けられます）");
  }

  // ---------------------------------------------------------------- commands (ribbon, menu, keys)

  function removeSelection() {
    const list = objects();
    const ids = ed.sel.filter((id) => !byId(id, list)?.locked);
    if (!ids.length) return;
    commit(ops.remove(list, ids, ends(list)), { select: [] });
  }
  // 複製 (⌘D) repeats a move: a copy moved after it was made sets where the next copy goes (as PowerPoint does).
  let lastDup = null; // { to: ids of the copies, from: where the first source was, made: where its copy was put }
  const posOf = (o) => (o?.kind === "line" ? [o.x1, o.y1] : [o?.x ?? 0, o?.y ?? 0]);
  function duplicateSelection() {
    const source = objects();
    const chosen = new Set(ed.sel);
    const groups = new Set(source.filter((o) => chosen.has(o.id)).map((o) => o.group).filter(Boolean));
    const first = source.find((o) => chosen.has(o.id));
    let offset = [24, 24];
    if (lastDup && lastDup.to.length === ed.sel.length && lastDup.to.every((id) => chosen.has(id)) && first) {
      const now = posOf(first);
      if (Math.hypot(now[0] - lastDup.made[0], now[1] - lastDup.made[1]) > 0.5) offset = [now[0] - lastDup.from[0], now[1] - lastDup.from[1]];
      else offset = lastDup.offset;
    }
    const made = ops.duplicate(source, ed.sel, 0);
    const list = ops.moveBy(made.list, made.ids, offset[0], offset[1]);
    const { ids, idMap, groupMap } = made;
    if (first) { const [x, y] = posOf(first); lastDup = { to: ids, from: [x, y], made: [x + offset[0], y + offset[1]], offset }; }
    const anims = copiedAnims(timeline().filter((e) => chosen.has(e.el) || (e.el?.startsWith("grp:") && groups.has(e.el.slice(4)))), idMap, groupMap);
    commit(list, { select: ids, ...(anims.length ? { timeline: [...timeline(), ...anims] } : {}) });
  }
  /** 図のリセット: trimming, adjustments, shape, outline and shadow back to none; 図とサイズのリセット also gives each
   *  picture its own size again (made smaller only to fit in the slide, its middle where it was) and takes off its turn
   *  and flips. Every picture in `ids` at once (one Undo). */
  const PICTURE_LOOK = { crop: undefined, mask: undefined, adj: undefined, bright: undefined, contrast: undefined, sat: undefined, gray: undefined, opacity: undefined, stroke: undefined, strokeW: undefined, shadow: undefined };
  function resetPicture(ids, withSize = false) {
    const list = (Array.isArray(ids) ? ids : [ids]).filter((id) => byId(id)?.kind === "image" && !byId(id).locked);
    if (!list.length) return;
    commit(ops.update(objects(), list, (o) => {
      const patch = { ...PICTURE_LOOK };
      const img = withSize ? ed.wrap?.querySelector(`.hs-obj[data-el="${CSS.escape(o.id)}"] .hs-obj-img img`) : null;
      if (img?.naturalWidth) {
        const k = Math.min(1, (ops.W * 0.8) / img.naturalWidth, (ops.H * 0.8) / img.naturalHeight);
        const [w, h] = [Math.round(img.naturalWidth * k), Math.round(img.naturalHeight * k)];
        Object.assign(patch, { w, h, x: Math.round(o.x + o.w / 2 - w / 2), y: Math.round(o.y + o.h / 2 - h / 2), rot: undefined, flipH: undefined, flipV: undefined, fit: undefined });
      }
      return patch;
    }));
  }
  function groupSelection() {
    if (ed.sel.length < 2) return app.toast("グループ化するには2つ以上選んでください");
    const list = ops.group(objects(), ed.sel);
    ed.entered = null;
    commit(list);
  }
  function ungroupSelection() {
    if (!selected().some((o) => o.group)) return app.toast("グループが選ばれていません");
    ed.entered = null;
    const list = objects();
    // 再グループ化 remembers each group taken apart (its members' ids).
    const groups = [...new Set(selected().filter((o) => o.group).map((o) => o.group))];
    ed.regroup = [...(ed.regroup || []), ...groups.map((g) => list.filter((o) => o.group === g).map((o) => o.id))].slice(-20);
    commit(ops.ungroup(list, ed.sel));
  }
  /** 再グループ化: the group last taken apart (the one the chosen objects belong to) put together again. */
  function regroupIds() {
    const list = objects();
    const alive = (ids) => ids.filter((id) => list.some((o) => o.id === id && !o.group));
    const mine = (ed.regroup || []).map(alive).filter((ids) => ids.length >= 2);
    return (ed.sel.length ? mine.filter((ids) => ids.some((id) => ed.sel.includes(id))) : mine).at(-1) || null;
  }
  function regroupSelection() {
    const ids = regroupIds();
    if (!ids) return app.toast("再グループ化できるグループがありません（グループ解除したオブジェクトを選んでください）");
    ed.entered = null;
    ed.regroup = (ed.regroup || []).filter((g) => !g.some((id) => ids.includes(id)));
    commit(ops.group(objects(), ids), { select: ids });
  }
  function order(how) { commit(ops.reorder(objects(), ops.withGroups(objects(), ed.sel), how)); }
  // 配置 → スライドに合わせて配置 / 選択したオブジェクトを揃える (one object always aligns to the slide).
  function alignSelection(how, toSlide = false) { const list = objects(); commit(ops.align(list, ed.sel, how, { toSlide: toSlide || ed.alignTo === "slide" || ed.sel.length === 1, ends: ends(list) })); }
  function distributeSelection(axis) {
    if (ed.sel.length < 3) return app.toast("等間隔に並べるには3つ以上選んでください");
    const list = objects();
    commit(ops.distribute(list, ed.sel, axis, { ends: ends(list) }));
  }
  function rotateSelection(deg) { commit(ops.rotateBy(objects(), ed.sel, deg)); }
  function flipSelection(axis) { commit(ops.flip(objects(), ed.sel, axis)); }
  function setLocked(on) { commit(ops.update(objects(), ed.sel, { locked: on ? true : undefined })); }
  function setHidden(ids, on) { commit(ops.update(objects(), ids, { hidden: on ? true : undefined }), { select: on ? ed.sel.filter((id) => !ids.includes(id)) : ed.sel }); }
  function rename(id, name) { commit(ops.update(objects(), [id], { name: name.trim() || undefined })); }
  function moveInOrder(id, index) { commit(ops.moveTo(objects(), id, index)); }

  /** Change settings of the selected objects (fill, line, size…). A function gets each object. */
  function apply(patch, { ids, undo = true } = {}) {
    const target = ids ?? ed.sel;
    if (!target.length) return;
    commit(ops.update(objects(), target, patch), { undo });
    // 繰り返し (F4): the last change made to the selection is what comes next.
    if (ids === undefined && undo && !repeating) lastAction = { run: () => apply(patch) };
  }
  // 繰り返し: do the last format again to what is selected now (PowerPoint's F4, and ⌘Y when there is nothing to redo).
  let lastAction = null;
  let repeating = false;
  function repeat() {
    if (!lastAction) { app.toast("繰り返す操作がありません"); return false; }
    if (!ed.sel.length && !ed.typing) { app.toast("繰り返す部品（か文字）を選んでください"); return false; }
    repeating = true;
    try { lastAction.run(); } finally { repeating = false; }
    return true;
  }

  // -- text: inside the text being typed (the selected words) or the whole of each selected object
  function textFormat(kind, value) {
    if (!repeating) lastAction = { run: () => textFormat(kind, value) };
    const t = ed.typing;
    // A table: a block of cells, the cell's own settings (vertical alignment), or the whole table when it is
    // selected without typing.
    if (t?.cell && (ed.cellRange || kind === "valign")) { tableFormat(kind, value); return; }
    if (!t && tableTarget()) { tableFormat(kind, value); return; }
    if (t) {
      restoreRange();
      const cmd = { bold: "bold", italic: "italic", underline: "underline", strike: "strikeThrough", sup: "superscript", sub: "subscript", clear: "removeFormat", bullet: "insertUnorderedList", number: "insertOrderedList", indent: "indent", outdent: "outdent" }[kind];
      if (cmd) { document.execCommand(cmd); if (kind === "bullet" || kind === "number") unnestLists(); }
      else if (kind === "color") document.execCommand("foreColor", false, value);
      else if (kind === "highlight") document.execCommand("hiliteColor", false, value || "transparent");
      else if (kind === "link") document.execCommand("createLink", false, value);
      else if (kind === "unlink") document.execCommand("unlink");
      else if (kind === "uline" || kind === "sline") markLine(kind === "uline" ? "u" : "s", value);
      else if (kind === "align") {
        document.execCommand({ left: "justifyLeft", center: "justifyCenter", right: "justifyRight", justify: "justifyFull", distributed: "justifyFull" }[value]);
        // 均等割り付け: justified to the last line too, on the paragraphs being typed in.
        for (const block of typedBlocks()) block.style.textAlignLast = value === "distributed" ? "justify" : "";
      }
      else if (kind === "size" || kind === "grow" || kind === "shrink") sizeSelectedWords(kind, value);
      // 段落 → インデント on the paragraphs being typed in.
      else if (kind === "paraTabs") { const text = E.tabsText(E.parseTabs(E.tabsText(value || []))); for (const block of typedBlocks()) { if (text) block.setAttribute("data-tabs", text); else block.removeAttribute("data-tabs"); } }
      else if (kind === "paraIndent") for (const block of typedBlocks()) { block.style.marginLeft = value?.left ? `${value.left}px` : ""; block.style.textIndent = value?.first ? `${value.first}px` : ""; }
      else if (["valign", "lh", "vertical", "direction", "font", "autofit", "pad", "wrap", "psp", "ls", "caps", "tfx", "listStyle", "listProps"].includes(kind)) { stopTyping(true); applyText(kind, value); return; }
      growWhileTyping();
      emit();
      return;
    }
    applyText(kind, value);
  }
  function tableFormat(kind, value) {
    changeTable((o, t) => {
      const block = [];
      for (let i = t.r0; i <= t.r1; i += 1) for (let j = t.c0; j <= t.c1; j += 1) if (!o.cells[i][j].merged) block.push(o.cells[i][j]);
      const set = (patch) => ops.tableCells(o, t.r0, t.c0, t.r1, t.c1, patch);
      // Settings on the whole table clear the same setting written into its words.
      const clearWords = (what) => (cell) => (cell.text ? { text: ops.clearInline(E, cell.text, what) || undefined } : {});
      switch (kind) {
        case "bold": case "italic": case "underline": case "strike": {
          const on = !block.every((cell) => cell[kind]);
          return set((cell) => ({ [kind]: on || undefined, ...clearWords(kind)(cell) }));
        }
        case "color": return set((cell) => ({ color: value, ...clearWords("color")(cell) }));
        case "align": return set((cell) => ({ align: value === "left" ? undefined : value, ...clearWords("align")(cell) }));
        case "valign": return set({ valign: value === "middle" ? undefined : value });
        case "size": return { ...o, fs: ops.fromPt(value), cells: o.cells.map((row) => row.map((cell) => (cell.text ? { ...cell, text: ops.clearInline(E, cell.text, "size") || undefined } : cell))) };
        case "grow": case "shrink": return { ...o, fs: ops.fromPt(stepSize((o.fs || 28) / ops.PX_PER_PT, kind === "grow" ? 1 : -1)) };
        case "font": return { ...o, font: value === "body" ? undefined : value };
        case "clear": return set((cell) => ({ bold: undefined, italic: undefined, underline: undefined, strike: undefined, color: undefined, text: cell.text ? ["bold", "italic", "underline", "strike", "color", "size", "highlight"].reduce((html, what) => ops.clearInline(E, html, what), cell.text) || undefined : undefined }));
        default: return null;
      }
    }, { keepTyping: Boolean(ed.typing) });
  }
  function sizeSelectedWords(kind, value) {
    const o = E.withDefaults(byId(ed.typing.id));
    const current = (() => { const node = window.getSelection()?.anchorNode; const el = node?.nodeType === 3 ? node.parentElement : node; return el ? parseFloat(getComputedStyle(el).fontSize) || o.fs : o.fs; })();
    const pt = kind === "size" ? Number(value) : stepSize(current / ops.PX_PER_PT, kind === "grow" ? 1 : -1);
    // Chrome writes <font size="7"> for a size; it becomes the size asked for.
    document.execCommand("styleWithCSS", false, false);
    document.execCommand("fontSize", false, "7");
    document.execCommand("styleWithCSS", false, true);
    for (const font of ed.typing.tx.querySelectorAll('font[size="7"]')) {
      const span = document.createElement("span");
      span.style.fontSize = `${ops.fromPt(pt)}px`;
      span.append(...font.childNodes);
      font.replaceWith(span);
      const range = document.createRange();
      range.selectNodeContents(span);
      const selection = window.getSelection();
      selection.removeAllRanges();
      selection.addRange(range);
    }
  }
  const stepSize = (pt, dir) => (dir > 0 ? FONT_SIZES.find((s) => s > pt + 0.01) ?? pt + 8 : [...FONT_SIZES].reverse().find((s) => s < pt - 0.01) ?? Math.max(6, pt - 2));

  function applyText(kind, value) {
    const list = objects();
    const ids = ed.sel.filter((id) => ["shape", "text"].includes(byId(id, list)?.kind));
    if (!ids.length) return;
    const all = ids.map((id) => E.withDefaults(byId(id, list)));
    const toggle = (key) => !all.every((o) => o[key]);
    commit(ops.update(list, ids, (raw) => {
      const o = E.withDefaults(raw);
      const text = raw.text;
      switch (kind) {
        case "bold": case "italic": case "underline": case "strike": {
          const on = toggle(kind);
          const patch = { [kind]: on || undefined, text: ops.clearInline(E, text, kind) || undefined };
          // Off: the line's kind goes with it.
          if (!on && kind === "underline") patch.uline = undefined;
          if (!on && kind === "strike") patch.sline = undefined;
          return patch;
        }
        // 下線の種類・二重取り消し線: the whole box underlined (struck) with that line.
        case "uline": return { underline: true, uline: value === "single" ? undefined : value, text: ops.clearInline(E, text, "underline") || undefined };
        case "sline": return { strike: true, sline: value === "double" ? "double" : undefined, text: ops.clearInline(E, text, "strike") || undefined };
        case "color": return { color: value, text: ops.clearInline(E, text, "color") || undefined };
        case "size": return { fs: ops.fromPt(value), text: ops.clearInline(E, text, "size") || undefined };
        case "grow": case "shrink": return { fs: ops.fromPt(stepSize(o.fs / ops.PX_PER_PT, kind === "grow" ? 1 : -1)), text: ops.clearInline(E, text, "size") || undefined };
        case "align": return { align: value, text: ops.clearInline(E, text, "align") || undefined };
        case "valign": return { valign: value };
        case "bullet": case "number": { const type = kind === "bullet" ? "bullet" : "number"; return { text: ops.setList(E, text, ops.listOf(text) === type ? null : type) }; }
        // 箇条書き・段落番号の種類: the list with that marker (the text made a list of that kind if it was not).
        case "listStyle": { const [type, style] = value; return { text: ops.setList(E, text, type, style) }; }
        case "listProps": return ops.listPropsOf(text) || /<(ul|ol)[ >]/.test(text) ? { text: ops.setListProps(E, text, value) } : null;
        case "lh": return { lh: Number(value) };
        case "psp": return { psp: Number(value) || undefined };
        case "ls": return { ls: Number(value) || undefined };
        case "caps": return { caps: value === "all" || value === "small" ? value : undefined };
        // 文字の効果: the letters' 影 and 光彩 (a patch of the two, each or neither).
        case "tfx": return { tshadow: value?.tshadow || undefined, tglow: value?.tglow || undefined, toutline: value?.toutline || undefined };
        case "paraIndent": return { text: ops.setIndent(E, text, value || {}) || undefined };
        case "paraTabs": return { text: ops.setTabs(E, text, value) || undefined };
        case "vertical": return { vertical: !o.vertical || undefined };
        // 文字列の方向: "h" 横書き, "v" 縦書き, 90 / 270 the letters turned on their sides.
        case "direction": return value === "v" ? { vertical: true, textRot: undefined } : value === 90 || value === 270 ? { vertical: undefined, textRot: value } : { vertical: undefined, textRot: undefined };
        case "font": return { font: value === "body" ? undefined : value };
        case "autofit": return { autofit: value };
        case "wrap": return { wrap: value ? undefined : false };
        case "pad": return { pad: value };
        case "clear": return { bold: undefined, italic: undefined, underline: undefined, strike: undefined, text: ["bold", "italic", "underline", "strike", "color", "size", "highlight"].reduce((acc, k) => ops.clearInline(E, acc, k), text) || undefined };
        case "highlight": {
          const cleared = ops.clearInline(E, text, "highlight");
          if (!value || value === "none") return { text: cleared || undefined };
          const box = document.createElement("div");
          box.append(E.richFragment(cleared || ""));
          for (const p of box.querySelectorAll("p, li")) { const span = document.createElement("span"); span.style.backgroundColor = value; span.append(...[...p.childNodes].filter((n) => !(n.nodeName === "UL" || n.nodeName === "OL"))); p.prepend(span); }
          return { text: E.sanitizeRich(box.innerHTML) };
        }
        default: return null;
      }
    }));
  }

  /** What the text controls should show for the selection (or the words being typed). */
  function textState() {
    const list = objects();
    const table = tableTarget();
    if (table && (!ed.typing || ed.cellRange)) {
      const o = E.withDefaults(table.o);
      const block = [];
      for (let i = table.r0; i <= table.r1; i += 1) for (let j = table.c0; j <= table.c1; j += 1) if (!o.cells[i][j].merged) block.push(o.cells[i][j]);
      const first = block[0] || {};
      const all = (key) => block.length > 0 && block.every((cell) => cell[key]);
      return { fs: ops.toPt(o.fs), color: first.color || o.color, bold: all("bold"), italic: all("italic"), underline: all("underline"), strike: all("strike"), align: first.align || "left", valign: first.valign || "middle", vertical: false, lh: o.lh, font: o.font || "body", autofit: "none", list: null, pad: null, wrap: true, table: true };
    }
    const chosen = selected(list).filter((o) => ["shape", "text"].includes(o.kind) || (o.kind === "table" && ed.typing?.cell)).map((o) => E.withDefaults(o));
    if (!chosen.length) return null;
    const o = chosen[0];
    const out = { tshadow: o.tshadow, tglow: o.tglow, toutline: o.toutline, fs: ops.toPt(o.fs), color: o.color, bold: chosen.every((x) => x.bold), italic: chosen.every((x) => x.italic), underline: chosen.every((x) => x.underline), strike: chosen.every((x) => x.strike), align: o.align, valign: o.valign, vertical: Boolean(o.vertical), lh: o.lh, font: o.font || "body", autofit: o.autofit, textRot: o.textRot || null, list: ops.listOf(o.text), listStyle: ops.listStyleOf(o.text), listProps: ops.listPropsOf(o.text), uline: o.underline ? o.uline || "single" : null, sline: o.strike ? o.sline || "single" : null, pad: o.pad, wrap: o.wrap !== false, caps: o.caps || null, ls: o.ls || 0, psp: o.psp || 0, indent: ops.indentOf(o.text), tabs: ops.tabsOf(E, o.text) };
    if (ed.typing) {
      try {
        out.bold = document.queryCommandState("bold");
        out.italic = document.queryCommandState("italic");
        out.underline = document.queryCommandState("underline");
        out.strike = document.queryCommandState("strikeThrough");
        const node = window.getSelection()?.anchorNode;
        const el = node?.nodeType === 3 ? node.parentElement : node;
        if (el && ed.typing.tx.contains(el)) out.fs = Math.round((parseFloat(getComputedStyle(el).fontSize) / ops.PX_PER_PT) * 10) / 10;
      } catch { /* not focused */ }
    }
    return out;
  }

  // -- format painter (書式のコピー/貼り付け)
  /** 書式のコピー: once (the next object clicked takes it), or kept on (a double-click) until Esc. */
  function copyFormat({ sticky = false } = {}) {
    const o = selected()[0];
    if (!o) return;
    const full = E.withDefaults(o);
    ed.painter = Object.fromEntries(STYLE_KEYS.filter((key) => full[key] !== undefined).map((key) => [key, JSON.parse(JSON.stringify(full[key]))]));
    ed.painter.kind = o.kind;
    ed.painter.sticky = sticky;
    app.toast(sticky ? "書式をコピーしました。クリックした図形に続けて貼り付けます（Escで終わる）" : "書式をコピーしました。貼り付ける図形をクリックしてください（Escでやめる）");
    emit();
  }
  function pasteFormat(ids) {
    if (!ed.painter || !ids.length) return;
    const style = ed.painter;
    commit(ops.update(objects(), ids, (o) => {
      const keys = o.kind === "line" ? ["stroke", "strokeW", "dash", "head", "tail", "headSize", "tailSize", "opacity"] : o.kind === "image" ? ["stroke", "strokeW", "dash", "opacity"] : STYLE_KEYS.filter((key) => !["head", "tail", "headSize", "tailSize", "route"].includes(key));
      const patch = Object.fromEntries(keys.filter((key) => style[key] !== undefined && (style.kind !== "line" || ["stroke", "strokeW", "dash", "opacity"].includes(key) || o.kind === "line")).map((key) => [key, style[key]]));
      // A plain fill painted over a gradient takes the gradient off; a gradient brings its first colour as the fill.
      if (keys.includes("gradient") && style.kind !== "line" && style.fill !== undefined && style.gradient === undefined) patch.gradient = undefined;
      if (keys.includes("pattern") && style.kind !== "line" && style.fill !== undefined && style.pattern === undefined) patch.pattern = undefined;
      return patch;
    }), { select: ids });
    if (!style.sticky) ed.painter = null;
    emit();
  }

  // ---------------------------------------------------------------- inserting (ribbon, drops)

  /** Add ready-made objects (pictures, icons…), centred on `at` or the slide, and select them. */
  function insert(objs, { at = null, select = true } = {}) {
    const list = objects();
    const placed = objs.map((o, i) => {
      if (o.kind === "line") return o;
      const cx = (at?.[0] ?? E.W / 2) + i * 24;
      const cy = (at?.[1] ?? E.H / 2 + 40) + i * 24;
      return { ...o, x: Math.round(cx - o.w / 2), y: Math.round(cy - o.h / 2) };
    });
    commit([...list, ...placed], { select: select ? placed.map((o) => o.id) : ed.sel });
    return placed;
  }

  // ---------------------------------------------------------------- right-click menu

  function onContextMenu(event) {
    if (!ed.slideEl) return;
    const list = objects();
    const hit = hitObject(event, list);
    event.preventDefault();
    if (hit && !ed.sel.includes(hit.id)) { ed.sel = hit.group && ed.entered !== hit.group ? ops.withGroups(list, [hit.id]) : [hit.id]; draw(); emit(); }
    if (!hit && !ed.sel.length) ed.sel = [];
    const p = toSlide(event);
    openMenu(event.clientX, event.clientY, menuItems(hit, p));
  }

  function menuItems(hit, p) {
    const any = ed.sel.length > 0;
    const one = ed.sel.length === 1 ? byId(ed.sel[0]) : null;
    const locked = selected().some((o) => o.locked);
    return [
      any && { label: "切り取り", keys: "⌘X", run: () => { const payload = copyPayload(); ed.clipboard = payload; ed.pasteCount = 0; try { localStorage.setItem("hpptx-editor-clipboard", JSON.stringify(payload)); } catch { /* full */ } removeSelection(); } },
      any && { label: "コピー", keys: "⌘C", run: () => { ed.clipboard = copyPayload(); ed.pasteCount = 0; try { localStorage.setItem("hpptx-editor-clipboard", JSON.stringify(ed.clipboard)); } catch { /* full */ } app.toast("コピーしました"); } },
      { label: "貼り付け", keys: "⌘V", run: () => { if (ed.clipboard || localStorage.getItem("hpptx-editor-clipboard")) { const payload = ed.clipboard || JSON.parse(localStorage.getItem("hpptx-editor-clipboard")); pasteObjects(payload, { at: hit ? null : p }); } else app.toast("貼り付けるオブジェクトがありません"); } },
      { label: "図として貼り付け", run: () => { const payload = ed.clipboard || (() => { try { return JSON.parse(localStorage.getItem("hpptx-editor-clipboard") || "null"); } catch { return null; } })(); if (payload) pasteObjects(payload, { at: hit ? null : p, mode: "picture" }); else app.toast("貼り付けるオブジェクトがありません"); } },
      any && { label: "複製", keys: "⌘D", run: duplicateSelection },
      any && !locked && { label: "削除", keys: "Delete", run: removeSelection },
      "-",
      one && ["shape", "text"].includes(one.kind) && !one.locked && { label: "テキストの編集", keys: "Enter", run: () => startTyping(one.id, { caret: "end" }) },
      any && { label: "書式のコピー", keys: "⇧⌘C", run: copyFormat },
      any && ed.painter && { label: "書式の貼り付け", keys: "⇧⌘V", run: () => pasteFormat(ed.sel) },
      "-",
      any && { label: "最前面へ移動", keys: "⇧⌘]", run: () => order("front") },
      any && { label: "前面へ移動", keys: "⌘]", run: () => order("forward") },
      any && { label: "背面へ移動", keys: "⌘[", run: () => order("backward") },
      any && { label: "最背面へ移動", keys: "⇧⌘[", run: () => order("back") },
      "-",
      ed.sel.length > 1 && { label: "グループ化", keys: "⌘G", run: groupSelection },
      regroupIds() && { label: "再グループ化", run: regroupSelection },
      selected().some((o) => o.group) && { label: "グループ解除", keys: "⇧⌘G", run: ungroupSelection },
      any && { label: locked ? "ロックを解除" : "ロック（動かないようにする）", run: () => setLocked(!locked) },
      any && { label: "図形の書式設定…", run: () => app.openPanel("format") },
      one && { label: "代替テキストを編集…", run: () => app.editAlt?.(one.id) },
      one && { label: "リンク・動作の設定…", run: () => app.openPanel("format", "action") },
      // A link on the object: open it, or take it away.
      one?.action?.type === "url" && { label: "リンクを開く", run: () => window.open(one.action.href, "_blank", "noopener") },
      one?.action?.type === "url" && !one.locked && { label: "リンクの削除", run: () => commit(ops.update(objects(), [one.id], () => ({ action: undefined }))) },
      any && { label: "図として保存…", run: () => app.saveAsPicture?.(ed.sel.slice()) },
      // 既定の図形・テキスト ボックス・線に設定: new ones in this deck start with this one's look.
      one?.kind === "shape" && { label: "既定の図形に設定", run: () => app.setObjectDefault?.("shape", one) },
      one?.kind === "text" && { label: "既定のテキスト ボックスに設定", run: () => app.setObjectDefault?.("text", one) },
      one?.kind === "line" && { label: "既定の線に設定", run: () => app.setObjectDefault?.("line", one) },
      any && { label: "アニメーション…", run: () => { app.showTab("animation", { open: true }); app.openAnimationPane(); } },
      one?.kind === "shape" && one.shape === "custom" && !one.locked && { label: "頂点の編集", run: () => app.editPoints?.(one.id) },
      one?.kind === "image" && !one.locked && { label: "トリミング", run: () => app.startCrop?.(one.id) },
      one?.kind === "image" && !one.locked && { label: "図とサイズのリセット", run: () => resetPicture([one.id], true) },
      one?.kind === "chart" && { label: "データの編集…", run: () => app.editChart?.(one.id) },
      one?.kind === "smartart" && !one.locked && { label: "テキスト ウィンドウ", run: () => app.editSmartart?.(one.id, null) },
      one?.kind === "smartart" && !one.locked && { label: "図形に変換", run: () => app.smartartToShapes?.(one.id) },
      one?.kind === "equation" && !one.locked && { label: "数式の編集…", run: () => app.editEquation?.(one.id) },
      one?.kind === "poll" && !one.locked && { label: "アンケートの編集…", run: () => app.editPoll?.(one.id) },
      one?.kind === "zoom" && { label: "ズーム先へ移動", run: () => { const i = app.slideOfSid?.(one.target) ?? -1; if (i >= 0) app.select?.(i); } },
      one?.kind === "ink" && { label: "描画タブ（ペン・消しゴム・インクを図形に変換）", run: () => app.showTab?.("draw") },
      !any && { label: "すべて選択", keys: "⌘A", run: () => { ed.sel = visible().map((o) => o.id); draw(); emit(); } },
      !any && { label: "図形を描く…", run: () => app.showTab?.("insert") },
      !any && slide()?.background?.image && { label: "背景の保存…", run: () => app.saveBackground?.() },
      !any && app.canConvert?.() && "-",
      !any && app.canConvert?.() && { label: "図形に変換（レイアウトを部品に分ける）", run: () => app.convertSlide() },
      !any && "-",
      !any && { label: "垂直ガイドを追加", run: () => addGuide("x", p[0]) },
      !any && { label: "水平ガイドを追加", run: () => addGuide("y", p[1]) },
      !any && (customGuides().x.length || customGuides().y.length) && { label: "ガイドをすべて削除", run: clearGuides },
    ].filter(Boolean);
  }

  function openMenu(x, y, items) {
    closeMenu();
    const menu = h("div", { class: "ed-menu", role: "menu" });
    let lastSep = true;
    for (const item of items) {
      if (item === "-") { if (!lastSep) menu.append(h("div", { class: "ed-menu-sep" })); lastSep = true; continue; }
      lastSep = false;
      menu.append(h("button", { type: "button", role: "menuitem", onmousedown: (e) => e.preventDefault(), onclick: () => { closeMenu(); item.run(); } }, h("span", {}, item.label), item.keys ? h("kbd", {}, item.keys) : null));
    }
    document.body.append(menu);
    const r = menu.getBoundingClientRect();
    menu.style.left = `${Math.min(x, innerWidth - r.width - 8)}px`;
    menu.style.top = `${Math.min(y, innerHeight - r.height - 8)}px`;
    ed.menu = menu;
    setTimeout(() => document.addEventListener("pointerdown", closeOnOutside, true), 0);
  }
  function closeOnOutside(event) { if (!ed.menu?.contains(event.target)) closeMenu(); }
  function closeMenu() {
    ed.menu?.remove();
    ed.menu = null;
    document.removeEventListener("pointerdown", closeOnOutside, true);
  }

  // ---------------------------------------------------------------- guides drawn while moving

  function guideLines(lines, k) {
    return lines.map((line) => {
      if (line.axis === "x") return h("div", { class: `ed-guide v ${line.kind}`, style: { left: `${line.at * k}px`, top: `${line.from * k}px`, height: `${(line.to - line.from) * k}px` } });
      if (line.axis === "y") return h("div", { class: `ed-guide h ${line.kind}`, style: { top: `${line.at * k}px`, left: `${line.from * k}px`, width: `${(line.to - line.from) * k}px` } });
      if (line.axis === "gap") return h("div", { class: "ed-gap h", style: { top: `${line.at * k}px`, left: `${line.from * k}px`, width: `${(line.to - line.from) * k}px` } });
      return h("div", { class: "ed-gap v", style: { left: `${line.at * k}px`, top: `${line.from * k}px`, height: `${(line.to - line.from) * k}px` } });
    });
  }

  // ---------------------------------------------------------------- the stage's click (layout text) after a drag

  function consumeClick() {
    const was = ed.suppressClick;
    ed.suppressClick = false;
    return was;
  }

  // ---------------------------------------------------------------- view settings

  function setView(key, value) {
    ed[key] = value;
    store(key, value);
    app.rerender();
    emit();
  }

  return {
    get state() { return ed; },
    get selection() { return [...ed.sel]; },
    get typing() { return Boolean(ed.typing); },
    selectedObjects: () => selected(),
    batch, resetPicture, regroupSelection, canRegroup: () => Boolean(regroupIds()),
    objects,
    attach, draw, emit,
    subscribe(fn) { subs.add(fn); return () => subs.delete(fn); },
    overlay(fn) { overlays.add(fn); return () => overlays.delete(fn); },
    handle(prefix, fn) { handles.set(prefix, fn); return () => handles.delete(prefix); },
    pick, get picking() { return Boolean(ed.picker); },
    tableTarget, changeTable, setCellRange, get cellRange() { return ed.cellRange; },
    toSlide, scale, timeline,
    select(ids) { ed.sel = [...ids]; ed.entered = null; draw(); emit(); },
    clear() { if (ed.typing) stopTyping(true); ed.sel = []; ed.entered = null; ed.tool = null; draw(); emit(); },
    setTool, get tool() { return ed.tool; },
    startTyping, stopTyping, restoreRange,
    keydown, onCopy, onPaste, consumeClick,
    insert, pasteObjects, pasteFromMemory, commit, apply, textFormat, textState, repeat, canRepeat: () => Boolean(lastAction),
    // 類義語: the words selected while typing, the text round them, and a pick that takes their place.
    selectedWords: () => (ed.typing && ed.range ? ed.range.toString() : ""),
    typingText: () => ed.typing?.tx.textContent || "",
    // Excel スプレッドシート: the cell being typed in (the formula bar shows and sets its words).
    typingCell: () => (ed.typing?.cell ? { id: ed.typing.id, cell: [...ed.typing.cell] } : null),
    setTypingText(text, { focus = false } = {}) {
      const t = ed.typing;
      if (!t?.cell) return false;
      t.tx.innerHTML = richOf(String(text ?? ""));
      if (focus) {
        t.tx.focus({ preventScroll: true });
        const range = document.createRange();
        range.selectNodeContents(t.tx);
        range.collapse(false);
        const selection = window.getSelection();
        selection.removeAllRanges();
        selection.addRange(range);
      }
      return true;
    },
    replaceWords(text) { if (!restoreRange()) return false; document.execCommand("insertText", false, text); return true; },
    removeSelection, duplicateSelection, groupSelection, ungroupSelection, order, alignSelection, distributeSelection, rotateSelection, flipSelection,
    setLocked, setHidden, rename, moveInOrder, copyFormat, pasteFormat,
    get painter() { return ed.painter; },
    endPainter: () => { if (ed.painter) { ed.painter = null; app.toast("書式のコピーを終えました"); emit(); } },
    setZoom, zoomPercent, zoomStep, get zoom() { return ed.zoom; }, setView, applyZoom, openMenu, addGuide, clearGuides, customGuides,
    FONT_SIZES, closeMenu,
  };
}
