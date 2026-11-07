// The animation editor, PowerPoint style: the ribbon's アニメーション and 画面切り替え tabs, the アニメーション
// ウィンドウ (side pane: order, timing bars, every option), the order marks and motion paths on the stage (a
// path's points can be dragged), and drawing a path of one's own. Animations are slide.timeline; the engine
// (engine/animate.js) plays them.

import * as ops from "./ops.mjs";
import { ico } from "./icons.mjs";

const CLASS_ICON = { in: "star", em: "star", out: "star", path: "motionPath", media: "media" };
const START_ICON = { click: "mouse", with: null, after: "clock" };
const GALLERY_QUICK = [["in", "fade"], ["in", "flyIn"], ["in", "wipe"], ["in", "zoom"], ["in", "floatIn"], ["in", "appear"]];
const PATH_COLORS = { start: "#16a34a", end: "#dc2626" };

export function createAnimations(editor, app, kit) {
  const { E, h } = app;
  const { btn, drop, group, col, row, menu, openPop, closePop, updater, tabNow } = kit;
  const pane = document.getElementById("animPane");
  let picked = [];
  let painter = null;
  let drawing = null;
  let dragPath = null;
  let auto = stored("autopreview", true);

  function stored(key, fallback) { try { const v = localStorage.getItem(`hpptx-editor-${key}`); return v == null ? fallback : JSON.parse(v); } catch { return fallback; } }
  function store(key, value) { try { localStorage.setItem(`hpptx-editor-${key}`, JSON.stringify(value)); } catch { /* optional */ } }

  // ---------------------------------------------------------------- reading

  const slide = () => app.slide();
  const timeline = () => (Array.isArray(slide()?.timeline) ? slide().timeline : []);
  const objects = () => editor.objects();
  const byId = (id) => objects().find((o) => o.id === id);
  const onSlide = () => Boolean(slide()) && app.state().view === "single";
  const plan = () => E.timelinePlan(slide() || {});
  const newId = () => `a${ops.newId().slice(1)}`;
  const label = (e) => (e.cls === "path" ? E.ANIM_PATHS[e.fx]?.label : E.animLabel(e.cls, e.fx)) || e.fx;

  /** The name of what an animation moves (selection-pane names, a group's, or a part of the layout). */
  function targetName(el) {
    if (el.startsWith("@")) return el === "@title" ? "タイトル" : el === "@takeaway" ? "キーメッセージ" : `本文の項目 ${Number(el.slice(2)) + 1}`;
    if (el.startsWith("grp:")) {
      const members = objects().filter((o) => o.group === el.slice(4));
      return `グループ（${members.length}）${members[0] ? `：${E.objectName(members[0], objects().indexOf(members[0]))}…` : ""}`;
    }
    const o = byId(el);
    return o ? E.objectName(o, objects().indexOf(o)) : "（見つからない）";
  }
  /** A trigger's name: the object clicked, or a bookmark of a video or a sound ("id@mark"). */
  function triggerName(key) {
    const [id, mark] = String(key).split("@");
    if (!mark) return `クリック：${targetName(id)}`;
    return `ブックマーク時：${targetName(id)}「${byId(id)?.bookmarks?.find((b) => b.id === mark)?.name || mark}」`;
  }
  /** Every bookmark on the slide as a trigger choice: [key, label]. */
  const bookmarkTriggers = () => objects().filter((o) => (o.kind === "video" || o.kind === "audio") && o.bookmarks?.length)
    .flatMap((o) => o.bookmarks.map((b) => [`${o.id}@${b.id}`, `${targetName(o.id)}「${b.name}」`]));

  /** What the selection animates: a whole group as one ("grp:…"), otherwise each object. */
  function selectionTargets() {
    const sel = editor.selection;
    const list = objects();
    const out = [];
    const seen = new Set();
    for (const id of sel) {
      const o = byId(id);
      if (!o) continue;
      if (o.group && editor.state.entered !== o.group) {
        const members = list.filter((m) => m.group === o.group);
        if (members.every((m) => sel.includes(m.id))) { if (!seen.has(o.group)) { seen.add(o.group); out.push(`grp:${o.group}`); } continue; }
      }
      out.push(o.id);
    }
    return out;
  }
  const targetsOf = (e) => (e.el.startsWith("grp:") ? objects().filter((o) => o.group === e.el.slice(4)).map((o) => o.id) : e.el.startsWith("@") ? [] : [e.el]);

  /** Picks are of this slide's animations only (a pick left on another slide is forgotten). */
  function prunePicked() {
    const list = timeline();
    if (picked.some((id) => !list.some((e) => e.id === id))) picked = picked.filter((id) => list.some((e) => e.id === id));
  }
  /** The animations the ribbon's options change: those picked in the pane, else the selection's. */
  function active() {
    const list = timeline();
    const chosen = list.filter((e) => picked.includes(e.id));
    if (chosen.length) return chosen;
    const targets = selectionTargets();
    return list.filter((e) => targets.includes(e.el));
  }

  // ---------------------------------------------------------------- changing

  function commit(next, { pick = null, preview = null } = {}) {
    if (pick) picked = pick;
    app.setTimeline(next);
    picked = picked.filter((id) => timeline().some((e) => e.id === id));
    refresh();
    // After the stage has redrawn with the change (a redraw would end the preview).
    if (preview && auto) setTimeout(() => previewEntries(preview), 80);
  }
  function refresh() { kit.refreshRibbon(); renderPane(); editor.draw(); }
  const update = (ids, patch) => timeline().map((e) => (ids.includes(e.id) ? clean({ ...e, ...(typeof patch === "function" ? patch(e) : patch) }) : e));
  const clean = (e) => Object.fromEntries(Object.entries(e).filter(([, v]) => v !== undefined && v !== null));

  function boxOf(el) {
    if (el.startsWith("grp:")) { const m = objects().filter((o) => o.group === el.slice(4)); return ops.bounds(m, (o) => E.lineEnds(o, objects())); }
    const o = byId(el);
    return o ? ops.bounds([o], (x) => E.lineEnds(x, objects())) : { x: 0, y: 0, w: 200, h: 200 };
  }
  function makeEntry(el, cls, fx, start) {
    const e = { id: newId(), el, cls, fx, start, dur: E.animDefaultDur(cls, fx), delay: 0 };
    if (cls === "path") e.path = E.pathPreset(fx, boxOf(el));
    return e;
  }
  const mediaOk = (el) => ["video", "audio", "lottie"].includes(byId(el)?.kind);

  /** Add an effect to everything selected (the first on a click, the others with it), as アニメーションの追加. */
  function add(cls, fx) {
    const targets = selectionTargets().filter((el) => cls !== "media" || mediaOk(el));
    if (!targets.length) { app.toast(cls === "media" ? "動画かアニメーションを選んでください" : "アニメーションを付ける図形・文字・画像を選んでください（レイアウトの部品はアニメーション ウィンドウから）"); return; }
    if (cls === "path" && fx === "custom") { startDrawing(targets[0]); return; }
    const entries = targets.map((el, i) => makeEntry(el, cls, fx, i ? "with" : "click"));
    commit([...timeline(), ...entries], { pick: entries.map((e) => e.id), preview: entries });
  }

  /** The gallery: the picked animations (or the selection's) become this effect; what has none gets it. */
  function apply(cls, fx) {
    prunePicked();
    if (cls === "path" && fx === "custom") { const t = selectionTargets()[0] || active()[0]?.el; if (t) startDrawing(t, active().find((e) => e.el === t && e.cls === "path")?.id); else app.toast("パスを描く図形を選んでください"); return; }
    const chosen = active();
    const targets = selectionTargets();
    if (!chosen.length && !targets.length) { app.toast("アニメーションを付ける図形・文字・画像を選んでください"); return; }
    const changed = [];
    let list = timeline();
    if (picked.length) {
      list = list.map((e) => { if (!picked.includes(e.id)) return e; const n = replaceFx(e, cls, fx); changed.push(n); return n; });
    } else {
      // An object's animations are replaced by the one chosen (the first keeps its place and start).
      const done = new Set();
      list = list.flatMap((e) => {
        if (!targets.includes(e.el)) return [e];
        if (done.has(e.el)) return [];
        done.add(e.el);
        const n = replaceFx(e, cls, fx);
        changed.push(n);
        return [n];
      });
      const fresh = targets.filter((el) => !done.has(el) && (cls !== "media" || mediaOk(el))).map((el, i) => makeEntry(el, cls, fx, i || done.size ? "with" : "click"));
      if (fresh.length) { fresh[0].start = done.size ? "with" : "click"; list = [...list, ...fresh]; changed.push(...fresh); }
    }
    commit(list, { pick: changed.map((e) => e.id), preview: changed });
  }
  function replaceFx(e, cls, fx) {
    const { dir: _d, amount: _a, color: _c, path: _p, by: _b, ease: _e, autoReverse: _r, ...keep } = e;
    const n = { ...keep, cls, fx, dur: E.animDefaultDur(cls, fx) };
    if (cls === "path") n.path = E.pathPreset(fx, boxOf(e.el));
    return n;
  }
  function removeEntries(ids) {
    if (!ids.length) return;
    commit(timeline().filter((e) => !ids.includes(e.id)), { pick: [] });
  }
  function none() {
    const ids = active().map((e) => e.id);
    if (!ids.length) { app.toast("外すアニメーションがありません"); return; }
    removeEntries(ids);
  }
  function setOption(patch) {
    const ids = active().map((e) => e.id);
    if (!ids.length) { app.toast("アニメーションを選んでください"); return; }
    commit(update(ids, patch), { preview: patch.dir || patch.amount || patch.by || patch.color ? timeline().filter((e) => ids.includes(e.id)) : null });
  }
  function move(step) {
    const ids = active().map((e) => e.id);
    if (!ids.length) return;
    commit(ops.moveAnimations(timeline(), ids, step), { pick: ids });
  }

  // ---------------------------------------------------------------- preview (the stage plays the slide)

  /** Play where these animations are: their click step on the stage, then back to editing. */
  function previewEntries(entries) {
    const p = plan();
    const first = entries.find((e) => !e.trigger);
    if (!first) return;
    const g = p.main.findIndex((grp) => grp.items.some((i) => i.e.id === first.id));
    if (g < 0) return;
    const lsteps = Number(editor.state.slideEl?.dataset.lsteps || 0);
    app.previewMotion({ from: g === 0 ? 0 : lsteps + g, auto: false, brief: true });
    const total = p.main[g]?.total || 600;
    setTimeout(() => { if (app.previewing()) app.stopPreview(); }, Math.min(8000, total + 900));
  }
  function playFromPicked() {
    const e = timeline().find((x) => picked.includes(x.id));
    if (!e) { app.previewMotion(); return; }
    const p = plan();
    const g = p.main.findIndex((grp) => grp.items.some((i) => i.e.id === e.id));
    const lsteps = Number(editor.state.slideEl?.dataset.lsteps || 0);
    app.previewMotion({ from: g <= 0 ? 0 : lsteps + g });
  }

  // ---------------------------------------------------------------- the animation painter and triggers

  function copyAnimations() {
    const src = active();
    if (!src.length) { app.toast("アニメーションの付いた図形を1つ選んでください"); return; }
    painter = src.map((e) => JSON.parse(JSON.stringify(e)));
    editor.pick((id, ids) => {
      const target = selectionTargetFor(id, ids);
      const others = timeline().filter((e) => e.el !== target);
      const copies = painter.map((e, i) => ({ ...JSON.parse(JSON.stringify(e)), id: newId(), el: target, ...(i === 0 ? { start: e.start } : {}) }));
      for (const c of copies) if (c.cls === "path") c.path = JSON.parse(JSON.stringify(c.path));
      painter = null;
      commit([...others, ...copies], { pick: copies.map((c) => c.id), preview: copies });
      app.toast("アニメーションを貼り付けました");
    }, "アニメーションを貼り付ける図形をクリックしてください（Escでやめる）");
  }
  const selectionTargetFor = (id, ids) => (ids.length > 1 && byId(id)?.group ? `grp:${byId(id).group}` : id);
  function pickTrigger() {
    if (!active().length) { app.toast("トリガーを付けるアニメーションを選んでください"); return; }
    const ids = active().map((e) => e.id);
    editor.pick((id) => {
      if (ids.some((x) => timeline().find((e) => e.id === x)?.el === id)) { app.toast("アニメーションする図形そのものはトリガーにできません"); return; }
      commit(update(ids, { trigger: id }), { pick: ids });
      app.toast(`「${targetName(id)}」をクリックすると再生します`);
    }, "クリックすると再生させる図形（トリガー）をクリックしてください（Escでやめる）");
  }

  // ---------------------------------------------------------------- galleries

  function fxButton(cls, fx, run, { on = false } = {}) {
    const name = cls === "path" ? E.ANIM_PATHS[fx]?.label : E.animLabel(cls, fx);
    const html = cls !== "path" && E.animIsHtml(cls, fx);
    return h("button", { type: "button", class: ["an-fx", `an-${cls}`, html ? "an-html" : "", on ? "on" : ""], title: html ? `${name}（HTMLならでは：PowerPointにはない動き）` : name, "data-fx": `${cls}:${fx}`, onmousedown: (event) => event.preventDefault(), onclick: () => run(cls, fx) },
      cls === "path" ? pathThumb(fx) : ico(CLASS_ICON[cls], 22), h("small", {}, name));
  }
  function pathThumb(fx) {
    const path = E.pathPreset(fx, { w: 100, h: 100 });
    const pts = E.pathPoints(path, 40);
    const xs = pts.map((p) => p[0]);
    const ys = pts.map((p) => p[1]);
    const minX = Math.min(...xs); const minY = Math.min(...ys);
    const span = Math.max(1, Math.max(...xs) - minX, Math.max(...ys) - minY);
    const k = 18 / span;
    const svg = E.s("svg", { viewBox: "-3 -3 26 26", width: 24, height: 24, class: "an-path-thumb", "aria-hidden": "true" });
    svg.append(E.s("path", { d: pts.map(([x, y], i) => `${i ? "L" : "M"}${((x - minX) * k).toFixed(1)} ${((y - minY) * k).toFixed(1)}`).join(" "), fill: "none", stroke: "currentColor", "stroke-width": 1.6, "stroke-dasharray": "2.5 2" }));
    svg.append(E.s("circle", { cx: ((pts[0][0] - minX) * k).toFixed(1), cy: ((pts[0][1] - minY) * k).toFixed(1), r: 2.2, fill: PATH_COLORS.start }));
    return svg;
  }
  // What only HTML can do (typewriter, decode, a pen drawing, figures counting, a chart growing…) comes first.
  // 3D モデルの動き (到着・ターンテーブル・スイング・ジャンプしてターン・退出) are offered only for 3D models.
  const isModelFx = (cls, fx) => Boolean((cls === "em" ? E.ANIM_EM : E.ANIM_IN)[fx]?.model);
  const htmlKeys = (cls) => Object.keys(cls === "em" ? E.ANIM_EM : E.ANIM_IN).filter((fx) => E.animIsHtml(cls, fx) && !isModelFx(cls, fx) && !(cls === "out" && E.ANIM_IN[fx].noExit));
  const pptKeys = (cls) => Object.keys(cls === "em" ? E.ANIM_EM : E.ANIM_IN).filter((fx) => !E.animIsHtml(cls, fx) && !isModelFx(cls, fx));
  const modelKeys = (cls) => Object.keys(cls === "em" ? E.ANIM_EM : E.ANIM_IN).filter((fx) => isModelFx(cls, fx));
  function htmlBlock(go) {
    return h("div", { class: "an-html-block" },
      h("div", { class: "rb-gallery-head an-head-html" }, h("b", { class: "an-html-badge" }, "HTML"), "HTMLならでは（PowerPointにはない動き）"),
      h("div", { class: "an-grid" }, [...htmlKeys("in").map((fx) => fxButton("in", fx, go)), ...htmlKeys("em").map((fx) => fxButton("em", fx, go)), ...htmlKeys("out").map((fx) => fxButton("out", fx, go))]));
  }
  /** Every effect, by kind (the gallery's ▾ and アニメーションの追加). */
  function gallery(run, { title = "", htmlOnly = false } = {}) {
    return (close) => {
      const go = (cls, fx) => { close(); run(cls, fx); };
      const hasMedia = selectionTargets().some(mediaOk);
      const models = selectionTargets().length > 0 && selectionTargets().every((id) => byId(id)?.kind === "model");
      const section = (cls, keys) => [h("div", { class: `rb-gallery-head an-head-${cls}` }, E.ANIM_CLASSES[cls]), h("div", { class: "an-grid" }, keys.map((fx) => fxButton(cls, fx, go)))];
      if (htmlOnly) return h("div", { class: "rb-gallery an-gallery" }, title ? h("div", { class: "rb-menu-head" }, title) : null, htmlBlock(go));
      return h("div", { class: "rb-gallery an-gallery" },
        title ? h("div", { class: "rb-menu-head" }, title) : null,
        h("button", { type: "button", class: "an-none", onclick: () => { close(); none(); } }, "なし（アニメーションを外す）"),
        models ? h("div", { class: "an-model-block" }, h("div", { class: "rb-gallery-head an-head-model" }, "3D"),
          h("div", { class: "an-grid" }, [...modelKeys("in").map((fx) => fxButton("in", fx, go)), ...modelKeys("em").map((fx) => fxButton("em", fx, go)), ...modelKeys("out").map((fx) => fxButton("out", fx, go))])) : null,
        htmlBlock(go),
        section("in", pptKeys("in")),
        section("em", pptKeys("em")),
        section("out", pptKeys("out")),
        section("path", Object.keys(E.ANIM_PATHS)),
        hasMedia ? section("media", Object.keys(E.ANIM_MEDIA)) : null);
    };
  }

  /** 効果のオプション: directions, how much, colour, words or letters, and a path's shape. */
  function optionsMenu() {
    return (close) => {
      const list = active();
      const e = list[0];
      if (!e) return h("div", { class: "rb-note an-note" }, "アニメーションを選ぶと、方向などを選べます。");
      const same = list.every((x) => x.cls === e.cls && x.fx === e.fx);
      const items = [];
      const dirs = same ? E.animDirs(e.cls, e.fx) : null;
      if (dirs) { items.push({ head: "方向" }); for (const [k, l] of dirs) items.push({ label: l, on: (e.dir || defaultOf(e, "dir")) === k, run: () => setOption({ dir: k }) }); }
      const def = defOf(e);
      if (same && def?.amounts) { items.push("-", { head: e.fx === "spin" ? "回転の量" : e.fx === "transparency" ? "透明度" : "大きさ" }); for (const [k, l] of def.amounts) items.push({ label: l, on: (e.amount ?? def.amount) === k, run: () => setOption({ amount: k }) }); }
      if (same && def?.color) {
        items.push("-", { head: "色" });
        const palette = e.fx === "fontColor" ? E.PALETTE.text : e.fx === "lineColor" ? E.PALETTE.line.filter(([c]) => c !== "#ffffff") : E.PALETTE.fill.filter(([c]) => c !== "#ffffff");
        for (const [c, l] of palette) items.push({ label: l, icon: null, swatch: c, on: (e.color || def.color) === c, run: () => setOption({ color: c }) });
      }
      if (["in", "out", "em"].includes(e.cls) && list.every((x) => textual(x.el)) && !defOf(e)?.html) {
        items.push("-", { head: "テキストの動作" });
        for (const [k, l] of Object.entries(E.ANIM_BY)) items.push({ label: l, on: (e.by || "all") === k, run: () => setOption({ by: k === "all" ? undefined : k }) });
      }
      if (["in", "out", "em"].includes(e.cls) && list.every((x) => byId(x.el)?.kind === "smartart") && !defOf(e)?.html && !defOf(e)?.custom) {
        items.push("-", { head: "グラフィックの動作（SmartArt）" });
        for (const [k, l] of Object.entries(E.ANIM_BY_SMARTART)) items.push({ label: l, on: (e.by || "all") === k, run: () => setOption({ by: k === "all" ? undefined : k }) });
      }
      if (e.cls === "path") {
        items.push("-", { head: "パス" });
        items.push({ label: "パスの反転（逆にたどる）", icon: "reset", run: () => setOption({ path: reversePath(e.path) }) });
        items.push({ label: e.path?.closed ? "パスを開く" : "パスを閉じる（元の位置に戻る）", run: () => setOption({ path: { ...e.path, closed: e.path?.closed ? undefined : true } }) });
        items.push({ label: e.path?.curve ? "直線でつなぐ" : "曲線でつなぐ", run: () => setOption({ path: { ...e.path, curve: e.path?.curve ? undefined : true } }) });
        items.push({ label: "パスを描き直す…", icon: "pen", run: () => startDrawing(e.el, e.id) });
        items.push({ head: "スライド上の点をドラッグして形を変えられます（Shift＋終点で全体を回転・拡大）" });
      }
      if (!items.length) return h("div", { class: "rb-note an-note" }, "この効果に選べるオプションはありません。");
      return menuWithSwatches(items)(close);
    };
  }
  const defOf = (e) => (e.cls === "out" ? E.ANIM_IN[e.fx] : e.cls === "in" ? E.ANIM_IN[e.fx] : e.cls === "em" ? E.ANIM_EM[e.fx] : null);
  const defaultOf = (e, key) => defOf(e)?.[key];
  const textual = (el) => (el.startsWith("@") ? true : ["shape", "text"].includes(byId(el)?.kind));
  function reversePath(path) {
    const pts = [...path.pts].reverse();
    const [x0, y0] = pts[0];
    return { ...path, pts: pts.map(([x, y]) => [x - x0, y - y0]) };
  }
  // menu() of ui.mjs, with a colour chip on items that name a colour.
  function menuWithSwatches(items) {
    return (close) => {
      const el = menu(items.map((item) => (item && item.swatch ? { ...item, icon: null } : item)))(close);
      const buttons = [...el.querySelectorAll("button")];
      let bi = 0;
      for (const item of items.filter((x) => x && x !== "-" && !x.head)) {
        const b = buttons[bi++];
        if (item.swatch && b) b.querySelector(".rb-noicon")?.replaceWith(h("i", { class: "an-swatch", style: { background: item.swatch } }));
      }
      return el;
    };
  }

  // ---------------------------------------------------------------- the ribbon tabs

  function seconds(get, set, { title, min = 0, max = 60, stepS = 0.25, enabled = null } = {}) {
    const input = h("input", { type: "number", class: "rb-num an-sec", min, max, step: stepS, title, "aria-label": title,
      onchange: (event) => { const v = parseFloat(event.target.value); if (Number.isFinite(v)) set(Math.max(min, Math.min(max, v))); },
      onkeydown: (event) => { if (event.key === "Enter") event.target.blur(); } });
    updater(() => { const v = get(); input.disabled = enabled ? !enabled() : v == null; if (document.activeElement !== input) input.value = v == null ? "" : String(Math.round(v * 100) / 100); });
    return input;
  }
  function select(options, get, set, title) {
    const el = h("select", { class: "rb-font an-select", title, "aria-label": title, onchange: (event) => set(event.target.value) }, options.map(([v, l]) => h("option", { value: v }, l)));
    updater(() => { const v = get(); el.disabled = v == null; if (v != null) el.value = v; });
    return el;
  }
  const field = (labelText, control) => h("label", { class: "rb-field an-field" }, h("span", {}, labelText), control);
  const firstActive = () => active()[0] || null;

  function animationTab() {
    const quick = h("div", { class: "an-quick" }, GALLERY_QUICK.map(([cls, fx]) => fxButton(cls, fx, apply)));
    updater(() => {
      const e = firstActive();
      for (const b of quick.querySelectorAll(".an-fx")) b.classList.toggle("on", Boolean(e) && b.dataset.fx === `${e.cls}:${e.fx}`);
    });
    return [
      group("プレビュー",
        btn("previewPlay", "プレビュー", "このスライドのアニメーションを再生", () => app.previewMotion(), { big: true }),
        col(btn("check", "自動", "効果を選んだらすぐ再生して見せる（自動プレビュー）", () => { auto = !auto; store("autopreview", auto); kit.refreshRibbon(); }, { pressed: () => auto }))),
      group("アニメーション",
        h("div", { class: "an-quick-wrap" }, quick,
          drop("chevron", "", "すべての効果（開始・強調・終了・軌跡）", () => gallery(apply), {})),
        drop("effectOpts", "効果の|オプション", "方向・量・色・テキストの動作・パス", () => optionsMenu(), { big: true, enabled: () => active().length > 0 })),
      group("HTMLならでは",
        drop("magic", "HTMLの|効果", "PowerPointにはない動き：タイプライター・デコード・マスク・ぼかし・線を描く・カウントアップ・グラフが伸びる・光が走る・波紋・マーカー・スポットライト", () => gallery(apply, { title: "HTMLならではの効果", htmlOnly: true }), { big: true })),
      group("詳細設定",
        drop("starPlus", "アニメーション|の追加", "選んだ図形にもう1つアニメーションを付ける", () => gallery(add, { title: "アニメーションの追加" }), { big: true, enabled: () => selectionTargets().length > 0 }),
        col(btn("animPane", "ウィンドウ", "アニメーション ウィンドウ：順番とタイミングの一覧", () => app.openAnimationPane()),
          drop("bolt", "トリガー", "開始のタイミング：ほかの図形をクリックしたときに再生", () => triggerMenu(), { enabled: () => active().length > 0 }),
          btn("painter", "コピー/貼り付け", "アニメーションのコピー/貼り付け：このアニメーションをクリックした図形にも付ける", copyAnimations, { enabled: () => active().length > 0, pressed: () => Boolean(painter) }))),
      group("タイミング",
        col(field("開始", select(Object.entries(E.ANIM_STARTS), () => firstActive()?.start ?? null, (v) => setOption({ start: v }), "開始")),
          field("継続時間", seconds(() => (firstActive() ? firstActive().dur / 1000 : null), (v) => setOption({ dur: Math.round(v * 1000) || 1 }), { title: "継続時間（秒）", min: 0.01, max: 60 })),
          field("遅延", seconds(() => (firstActive() ? firstActive().delay / 1000 : null), (v) => setOption({ delay: Math.round(v * 1000) }), { title: "遅延（秒）" }))),
        col(btn("up", "前へ", "順番を前にする：このアニメーションを1つ前へ", () => move(-1), { enabled: () => active().length > 0 }),
          btn("down", "後へ", "順番を後にする：このアニメーションを1つ後へ", () => move(1), { enabled: () => active().length > 0 }),
          btn("trash", "削除", "選んだアニメーションを外す", none, { enabled: () => active().length > 0 }))),
    ];
  }
  function triggerMenu() {
    const e = firstActive();
    const own = new Set(active().map((x) => x.el));
    const items = [{ label: "スライドのクリック順（トリガーなし）", on: !e?.trigger, run: () => setOption({ trigger: undefined }) }, "-", { head: "次の図形をクリックしたとき" }];
    for (const o of objects()) if (!own.has(o.id) && !o.hidden) items.push({ label: E.objectName(o, objects().indexOf(o)), on: e?.trigger === o.id, run: () => setOption({ trigger: o.id }) });
    // ブックマーク時: as a video or a sound passes one of its bookmarks (再生 → ブックマークの追加).
    const marks = bookmarkTriggers();
    items.push("-", { head: "ブックマーク時" });
    if (marks.length) for (const [key, text] of marks) items.push({ label: text, on: e?.trigger === key, run: () => setOption({ trigger: key }) });
    else items.push({ label: "（ビデオ・オーディオにブックマークがありません）", disabled: true });
    items.push("-", { label: "スライド上でクリックして選ぶ…", icon: "bolt", run: pickTrigger });
    return menu(items);
  }

  function transitionTab() {
    const current = () => slide()?.transition || "auto";
    const set = (key) => {
      const type = key === "auto" ? app.deck()?.transition || "fade" : key;
      const keep = (E.TRANSITION_OPTIONS[type] || []).some(([k]) => k === slide()?.transitionDir);
      app.setSlideFields({ transition: key === "auto" ? undefined : key, ...(keep ? {} : { transitionDir: undefined }) });
      // 変形: say what will travel from the slide before (the same parts: a duplicated slide's, or the same "!!" name).
      if (key === "morph") {
        const before = app.deck()?.slides[app.index() - 1];
        const n = before ? E.morphPairs(before.elements, slide()?.elements).length : 0;
        app.toast(n ? `前のスライドと同じ部品 ${n} 個がつながって動きます` : "前のスライドと同じ部品がありません（スライドを複製して部品を動かすか、同じ「!!」で始まる名前を付けると、つながって動きます）");
      }
      if (auto && key !== "auto") app.previewTransition();
    };
    const kinds = [["auto", "資料の設定"], ...Object.entries(E.TRANSITIONS).map(([k, l]) => [k, l.replace(/（.*$/, "")])];
    const items = kinds.map(([key, name]) => {
      const b = h("button", { type: "button", class: "an-tr", "data-tr": key, title: key === "auto" ? `資料全体の切り替え（${E.TRANSITIONS[app.deck()?.transition] || "フェード"}）に従う` : E.TRANSITIONS[key], onclick: () => set(key) }, h("span", { class: `an-tr-icon an-tr-${key}` }), h("small", {}, name));
      updater(() => b.classList.toggle("on", current() === key));
      return b;
    });
    // 効果のオプション: the directions the slide's transition can come from (its first is the default).
    const typeNow = () => { const s = slide(); return s?.transition || app.deck()?.transition || "fade"; };
    const optionMenu = () => {
      const list = E.TRANSITION_OPTIONS[typeNow()] || [];
      const now = slide()?.transitionDir;
      return menu([{ head: "方向" }, ...list.map(([key, label], k) => ({ label, on: now ? now === key : k === 0, run: () => { app.setSlideFields({ transitionDir: k === 0 ? undefined : key }); if (auto) app.previewTransition(); } }))]);
    };
    // サウンド: a short sound as the slide arrives (played once here to hear it).
    const soundMenu = () => menu([
      { label: "［サウンドなし］", on: !slide()?.transitionSound, run: () => app.setSlideFields({ transitionSound: undefined }) },
      "-",
      ...Object.entries(E.TRANSITION_SOUNDS).map(([key, label]) => ({ label, on: slide()?.transitionSound === key, run: () => { app.setSlideFields({ transitionSound: key }); E.playSound?.(key); } })),
    ]);
    const soundName = h("small", { class: "an-sound-name" });
    updater(() => { soundName.textContent = slide()?.transitionSound ? E.TRANSITION_SOUNDS[slide().transitionSound] : "［サウンドなし］"; });
    const advanceOn = h("input", { type: "checkbox", onchange: (event) => app.setSlideFields({ advance: event.target.checked ? 5 : undefined }) });
    updater(() => { advanceOn.checked = slide()?.advance != null; });
    return [
      group("プレビュー", btn("previewPlay", "プレビュー", "このスライドへの切り替えを見る", () => app.previewTransition(), { big: true })),
      group("画面切り替え", h("div", { class: "an-tr-grid" }, items),
        drop("effectOpts", "効果の|オプション", "切り替えの方向（押し上げ・ワイプ・スライドなど）", optionMenu, { big: true, enabled: () => Boolean(slide() && E.TRANSITION_OPTIONS[typeNow()]) })),
      group("タイミング",
        col(drop("audio", "サウンド", "スライドが現れるときに鳴らす音", soundMenu, { enabled: () => Boolean(slide()) }),
          soundName),
        col(field("期間", seconds(() => { const s = slide(); const type = s?.transition || app.deck()?.transition || "fade"; return s ? (s.transitionDur ?? E.transitionMs(type)) / 1000 : null; }, (v) => app.setSlideFields({ transitionDur: Math.round(v * 1000) }), { title: "切り替えにかける時間（秒）", min: 0.1, max: 10, stepS: 0.05, enabled: () => Boolean(slide()) && (slide().transition || app.deck()?.transition || "fade") !== "none" })),
          btn("applyAll", "すべてに適用", "このスライドの切り替え・サウンド・期間・自動の切り替えを全スライドに", () => { const s = slide(); app.setSlideFields({ transition: s.transition, transitionDir: s.transitionDir, transitionDur: s.transitionDur, transitionSound: s.transitionSound, advance: s.advance }, { all: true }); app.toast("すべてのスライドに同じ切り替えを設定しました"); })),
        col(h("label", { class: "rb-field an-field an-check" }, h("input", { type: "checkbox", checked: true, disabled: true }), h("span", {}, "クリック時")),
          h("label", { class: "rb-field an-field an-check" }, advanceOn, h("span", {}, "自動的に切り替え")),
          field("秒後", seconds(() => slide()?.advance ?? null, (v) => app.setSlideFields({ advance: v }), { title: "自動的に切り替えるまでの秒数", min: 0, max: 600, stepS: 0.5, enabled: () => Boolean(slide()) })))),
    ];
  }

  // ---------------------------------------------------------------- the animation pane

  function renderPane() {
    if (!pane || pane.hidden) return;
    const keepScroll = pane.querySelector(".fp-body")?.scrollTop ?? 0;
    const head = h("div", { class: "fp-head an-pane-head" }, h("b", {}, "アニメーション ウィンドウ"));
    const body = h("div", { class: "fp-body" });
    // The whole deck at once (at the end of the pane, below what moves on this slide).
    const deckBox = h("div", { class: "an-sec" }, h("div", { class: "an-sec-head" }, ico("slide", 14), h("b", {}, "資料全体")),
      h("div", { class: "an-reset-actions" },
        h("p", { class: "hint" }, "全ページの動きとクリック動作をゼロにできます。文字・図形・画像とスライドは残り、⌘Zで復元できます。"),
        h("button", { type: "button", class: "btn btn-sm", onclick: () => app.resetAllActions() }, "全ページの動き・操作を削除"),
        h("button", { type: "button", class: "btn btn-sm", onclick: () => app.prepareActionPlanPrompt() }, "✦ AIにHTML演出を相談")));
    if (!onSlide()) {
      body.append(h("p", { class: "hint" }, "1枚表示（標準）で、図形・文字・画像やレイアウトの部品にアニメーションを付けられます。"), deckBox);
      pane.replaceChildren(head, body);
      return;
    }
    // Everything that moves on this slide, in the order it happens: on its own as it arrives, the animations
    // (clicks, with, after, triggers), then what answers the mouse (HTML) — each editable right here.
    body.append(autoSection(), h("div", { class: "an-sec-head an-sec-gap" }, ico("star", 14), h("b", {}, "アニメーション"), h("small", {}, "クリックの順・タイミング")));
    const p = plan();
    head.append(h("span", { class: "an-pane-btns" },
      h("button", { type: "button", class: "btn btn-sm", title: "このスライドのアニメーションを最初から再生", onclick: () => app.previewMotion() }, "▶ すべて再生"),
      h("button", { type: "button", class: "btn btn-sm", disabled: !picked.length || null, title: "選んだアニメーションから再生", onclick: playFromPicked }, "▶ 選択から")));
    const scale = Math.max(2000, ...p.main.map((g) => g.total), ...[...p.triggers.values()].flatMap((seq) => seq.map((g) => g.total)));
    const listEl = h("ol", { class: "an-list" });
    if (!p.list.length) {
      body.append(h("div", { class: "an-empty" },
        h("p", {}, "このスライドにはまだアニメーションがありません。"),
        h("p", { class: "hint" }, "図形・文字・画像を選び、リボンの「アニメーション」から効果を選びます。タイトルや本文の項目には下の「レイアウトの部品」から付けられます。")));
    }
    p.main.forEach((grp, g) => grp.items.forEach((item, i) => listEl.append(rowEl(item, { number: i === 0 && g > 0 ? String(g) : i === 0 && g === 0 ? "0" : "", scale, seq: null }))));
    body.append(listEl);
    for (const [trigger, seq] of p.triggers) {
      body.append(h("div", { class: "an-trigger-head" }, ico("bolt", 14), `トリガー：${triggerName(trigger)}`));
      const tl = h("ol", { class: "an-list" });
      seq.forEach((grp) => grp.items.forEach((item, i) => tl.append(rowEl(item, { number: i === 0 ? "⚡" : "", scale, seq: trigger }))));
      body.append(tl);
    }
    if (p.list.length) body.append(ruler(scale));
    body.append(h("div", { class: "an-actions" },
      h("button", { type: "button", class: "btn btn-sm", disabled: !active().length || null, onclick: () => move(-1), title: "順番を前にする" }, "▲"),
      h("button", { type: "button", class: "btn btn-sm", disabled: !active().length || null, onclick: () => move(1), title: "順番を後にする" }, "▼"),
      h("button", { type: "button", class: "btn btn-sm", disabled: !picked.length || null, onclick: () => removeEntries([...picked]) }, "削除"),
      h("span", { class: "spacer" }),
      h("button", { type: "button", class: "btn btn-sm", onclick: (event) => openPop(event.currentTarget, gallery(add, { title: "アニメーションの追加" })) }, "＋ 追加"),
      h("button", { type: "button", class: "btn btn-sm", onclick: (event) => openPop(event.currentTarget, layoutMenu()) }, "＋ レイアウトの部品")));
    const chosen = timeline().filter((e) => picked.includes(e.id));
    if (chosen.length) body.append(details(chosen));
    body.append(interactionSection(), deckBox);
    pane.replaceChildren(head, body);
    body.scrollTop = keepScroll;
  }

  /** What moves on its own as the slide arrives (the layout's and the deck's motion): one row each, changed here. */
  function autoSection() {
    const info = app.autoMotion?.();
    const sec = h("div", { class: "an-sec an-auto" }, h("div", { class: "an-sec-head" }, ico("magic", 14), h("b", {}, "スライドの自動の動き"), h("small", {}, "おまかせ＝資料・レイアウトの設定")));
    if (!info) return sec;
    for (const item of info.rows) {
      sec.append(h("label", { class: "an-auto-row" }, h("span", {}, item.label),
        h("select", { "data-auto": item.key, "aria-label": item.label, onchange: (event) => app.setAutoMotion(item.key, event.target.value) },
          item.options.map(([value, text]) => h("option", { value, selected: value === item.value || null }, text)))));
    }
    if (info.source) sec.append(h("p", { class: "hint" }, "PowerPointから見た目どおりに取り込んだページは、元の見た目を守るため、レイアウトの自動の動き（登場・背景・文字）を付けません。部品の動きは下のアニメーションとインタラクションで付けます。"));
    return sec;
  }

  const ACTION_WORDS = { next: "次のスライド", prev: "前のスライド", first: "最初のスライド", last: "最後のスライド", end: "スライドショーの終了", slide: "スライドへ移動", url: "Webページを開く", sound: "サウンドの再生" };
  /** The objects that answer the mouse, a click or keep moving (HTML only): each setting a chip, × takes it off. */
  function interactionSection() {
    const list = objects().filter((o) => !o.hidden && (o.hover || o.tip || o.loop || o.action || o.overAction));
    const sec = h("div", { class: "an-sec an-ix" }, h("div", { class: "an-sec-head" }, ico("hover", 14), h("b", {}, "インタラクション（HTML）"), h("small", {}, "マウス・クリック・ずっと動く")));
    if (!list.length) {
      sec.append(h("p", { class: "hint" }, "マウスを乗せたとき・クリックしたとき・ずっと動く動きは、まだありません。"),
        h("div", { class: "an-sec-btns" },
          h("button", { type: "button", class: "btn btn-sm", title: "このスライドの部品に、種類に合うHTMLの動きを付ける（⌘Zで戻せます）", onclick: () => app.enhance("slide") }, "✦ おまかせで付ける"),
          h("button", { type: "button", class: "btn btn-sm", onclick: () => app.showTab("interact", { open: true }) }, "インタラクション タブ")));
      return sec;
    }
    const ul = h("ul", { class: "an-ix-list" });
    for (const o of list) {
      const chips = [];
      if (o.hover) chips.push(["hover", `マウス：${E.IX_HOVERS[o.hover]}`, { hover: undefined }]);
      if (o.tip) chips.push(["tip", `説明「${o.tip.length > 10 ? `${o.tip.slice(0, 10)}…` : o.tip}」`, { tip: undefined }]);
      if (o.action) chips.push(["click", `クリック：${E.IX_CLICKS[o.action.type] || ACTION_WORDS[o.action.type] || o.action.type}`, { action: undefined }]);
      if (o.overAction) chips.push(["over", `マウスの通過：${ACTION_WORDS[o.overAction.type] || o.overAction.type}`, { overAction: undefined }]);
      if (o.loop) chips.push(["loop", `ずっと：${E.IX_LOOPS[o.loop]}`, { loop: undefined }]);
      ul.append(h("li", { class: ["an-ix-row", editor.selection.includes(o.id) ? "on" : ""], "data-id": o.id },
        h("button", { type: "button", class: "an-ix-name", title: "選んで、リボンの「インタラクション」で変える", onclick: () => { editor.select([o.id]); app.showTab("interact"); } }, E.objectName(o, objects().indexOf(o))),
        h("span", { class: "an-ix-chips" }, chips.map(([kind, text, patch]) => h("span", { class: `an-ix-chip an-ix-${kind}` }, h("span", {}, text),
          h("button", { type: "button", title: "外す", "aria-label": `${text}を外す`, onclick: () => editor.apply(patch, { ids: [o.id] }) }, "×"))))));
    }
    sec.append(ul);
    return sec;
  }

  function rowEl(item, { number, scale, seq }) {
    const e = item.e;
    const on = picked.includes(e.id);
    const li = h("li", { class: ["an-row", on ? "on" : ""], draggable: "true", "data-id": e.id, title: `${targetName(e.el)}：${E.ANIM_CLASSES[e.cls]}「${label(e)}」・${E.ANIM_STARTS[e.start]}・${(e.dur / 1000).toFixed(2)}秒${e.delay ? `・遅延${(e.delay / 1000).toFixed(2)}秒` : ""}`,
      onclick: (event) => { if (event.target.closest("button")) return; pickRow(e, event); },
      ondragstart: (event) => { event.dataTransfer.setData("text/x-hpptx-anim", e.id); event.dataTransfer.effectAllowed = "move"; },
      ondragover: (event) => { if ([...event.dataTransfer.types].includes("text/x-hpptx-anim")) { event.preventDefault(); li.classList.add("drop"); } },
      ondragleave: () => li.classList.remove("drop"),
      ondrop: (event) => { event.preventDefault(); li.classList.remove("drop"); dropOn(event.dataTransfer.getData("text/x-hpptx-anim"), e, seq); } },
    h("span", { class: "an-num" }, number),
    h("span", { class: "an-start" }, START_ICON[e.start] ? ico(START_ICON[e.start], 13) : ""),
    h("span", { class: `an-kind an-${e.cls}` }, ico(CLASS_ICON[e.cls], 15)),
    h("span", { class: "an-name" }, h("b", {}, targetName(e.el)), h("small", {}, label(e))),
    h("span", { class: "an-track" }, h("i", { class: `an-bar an-${e.cls}`, style: { left: `${(item.begin / scale) * 100}%`, width: `${Math.max(1.5, ((item.end - item.begin) / scale) * 100)}%` } })),
    h("button", { type: "button", class: "an-more", title: "メニュー", "aria-label": "メニュー", onclick: (event) => { pickRow(e, null, true); openPop(event.currentTarget, rowMenu(e)); } }, "▾"));
    return li;
  }
  function pickRow(e, event, keep = false) {
    if (event?.shiftKey && picked.length) {
      const order = timeline().map((x) => x.id);
      const a = order.indexOf(picked[picked.length - 1]);
      const b = order.indexOf(e.id);
      picked = [...new Set([...picked, ...order.slice(Math.min(a, b), Math.max(a, b) + 1)])];
    } else if (event && (event.metaKey || event.ctrlKey)) picked = picked.includes(e.id) ? picked.filter((x) => x !== e.id) : [...picked, e.id];
    else if (!keep || !picked.includes(e.id)) picked = [e.id];
    // The object it moves is selected on the stage too.
    const ids = [...new Set(timeline().filter((x) => picked.includes(x.id)).flatMap(targetsOf))];
    editor.select(ids);
    refresh();
  }
  function rowMenu(e) {
    return menu([
      ...Object.entries(E.ANIM_STARTS).map(([k, l]) => ({ label: l, icon: START_ICON[k], on: e.start === k, run: () => setOption({ start: k }) })),
      "-",
      { label: "効果のオプション…", icon: "effectOpts", run: () => requestAnimationFrame(() => openPop(pane.querySelector(".an-details") || pane, optionsMenu())) },
      { label: "ここから再生", icon: "previewPlay", run: playFromPicked },
      { label: "順番を前にする", icon: "up", run: () => move(-1) },
      { label: "順番を後にする", icon: "down", run: () => move(1) },
      "-",
      { label: "削除", icon: "trash", run: () => removeEntries([e.id]) },
    ]);
  }
  /** Dragging a row onto another moves it there (into that trigger's sequence, or out of one). */
  function dropOn(id, target, seq) {
    if (!id || id === target.id) return;
    const list = timeline();
    const moving = list.find((e) => e.id === id);
    if (!moving) return;
    const rest = list.filter((e) => e.id !== id);
    const at = rest.findIndex((e) => e.id === target.id);
    const placed = { ...moving };
    if (seq) placed.trigger = seq; else delete placed.trigger;
    rest.splice(at < 0 ? rest.length : at, 0, placed);
    commit(rest, { pick: [id] });
  }
  function layoutMenu() {
    const parts = editor.state.slideEl ? E.layoutTargets(editor.state.slideEl) : [];
    if (!parts.length) return () => h("div", { class: "rb-note an-note" }, "このスライドのレイアウトには、アニメーションを付けられる部品がありません。");
    return menu([
      { head: "フェードで登場させる（あとで効果を変えられます）" },
      ...parts.map((part) => ({ label: part.label, run: () => {
        const e = makeEntry(part.el, "in", "fade", "click");
        commit([...timeline(), e], { pick: [e.id], preview: [e] });
        if (part.el.startsWith("@g")) app.toast("本文の項目にアニメーションを付けると、レイアウトの「中身の出し方」の代わりに、アニメーションの順番で出ます");
      } })),
    ]);
  }
  function ruler(scale) {
    const marks = [];
    const step = scale > 8000 ? 2000 : 1000;
    for (let t = 0; t <= scale; t += step) marks.push(h("span", { style: { left: `${(t / scale) * 100}%` } }, `${t / 1000}`));
    return h("div", { class: "an-ruler" }, h("span", { class: "an-ruler-label" }, "秒"), h("div", { class: "an-ruler-track" }, marks));
  }

  /** The picked animations' settings, all in one place (PowerPoint's 効果 and タイミング dialogs). */
  function details(chosen) {
    const e = chosen[0];
    const same = chosen.every((x) => x.cls === e.cls && x.fx === e.fx);
    const sec = (key, title, opts) => h("input", { type: "number", min: opts?.min ?? 0, max: opts?.max ?? 60, step: 0.05, value: (e[key] / 1000).toFixed(2), "aria-label": title,
      onchange: (event) => { const v = parseFloat(event.target.value); if (Number.isFinite(v)) setOption({ [key]: Math.max(key === "dur" ? 1 : 0, Math.round(v * 1000)) }); } });
    const choice = (options, value, set, title) => h("select", { "aria-label": title, onchange: (event) => set(event.target.value) }, options.map(([v, l]) => h("option", { value: v, selected: String(v) === String(value) || null }, l)));
    const line = (labelText, ...controls) => h("div", { class: "fp-line" }, h("span", { class: "fp-label" }, labelText), h("div", { class: "fp-ctrl" }, ...controls.filter(Boolean)));
    const effects = [...["in", "em", "out"].flatMap((cls) => Object.keys(cls === "em" ? E.ANIM_EM : E.ANIM_IN).map((fx) => [`${cls}:${fx}`, `${E.ANIM_CLASSES[cls]}：${E.animLabel(cls, fx)}`])),
      ...Object.keys(E.ANIM_PATHS).map((fx) => [`path:${fx}`, `軌跡：${E.ANIM_PATHS[fx].label}`]), ...(mediaOk(e.el) ? Object.keys(E.ANIM_MEDIA).map((fx) => [`media:${fx}`, `メディア：${E.ANIM_MEDIA[fx].label}`]) : [])];
    const dirs = same ? E.animDirs(e.cls, e.fx) : null;
    const def = defOf(e);
    const triggers = [["", "スライドのクリック順"], ...objects().filter((o) => !chosen.some((x) => targetsOf(x).includes(o.id))).map((o) => [o.id, `クリック：${E.objectName(o, objects().indexOf(o))}`]),
      ...bookmarkTriggers().map(([key, text]) => [key, `ブックマーク時：${text}`])];
    const box = h("div", { class: "fp-sec an-details" },
      h("div", { class: "an-details-head" }, h("b", {}, chosen.length > 1 ? `${chosen.length}個のアニメーション` : `${targetName(e.el)}`), h("small", {}, chosen.length > 1 ? "" : `${E.ANIM_CLASSES[e.cls]}「${label(e)}」`)),
      line("効果", same ? choice(effects, `${e.cls}:${e.fx}`, (v) => { const [cls, fx] = v.split(":"); if (cls === "path" && fx === "custom") { startDrawing(e.el, e.id); return; } apply(cls, fx); }, "効果") : h("span", { class: "hint" }, "（いろいろ）")),
      dirs ? line("方向", choice(dirs, e.dir || def?.dir, (v) => setOption({ dir: v }), "方向")) : null,
      same && def?.amounts ? line(e.fx === "spin" ? "回転" : e.fx === "transparency" ? "透明度" : "大きさ", choice(def.amounts, e.amount ?? def.amount, (v) => setOption({ amount: Number(v) }), "量")) : null,
      same && def?.color ? line("色", h("span", { class: "fp-swatches" }, (e.fx === "fontColor" ? E.PALETTE.text : e.fx === "lineColor" ? E.PALETTE.line.filter(([c]) => c !== "#ffffff") : E.PALETTE.fill.filter(([c]) => c !== "#ffffff")).map(([c, l]) => h("button", { type: "button", class: ["fp-sw", (e.color || def.color) === c ? "on" : ""], title: l, "aria-label": l, style: { "--c": c }, onclick: () => setOption({ color: c }) })))) : null,
      ["in", "out", "em"].includes(e.cls) && chosen.every((x) => textual(x.el)) ? line("テキスト", choice(Object.entries(E.ANIM_BY), e.by || "all", (v) => setOption({ by: v === "all" ? undefined : v }), "テキストの動作")) : null,
      ["in", "out", "em"].includes(e.cls) && chosen.every((x) => byId(x.el)?.kind === "smartart") ? line("SmartArt", choice(Object.entries(E.ANIM_BY_SMARTART), e.by || "all", (v) => setOption({ by: v === "all" ? undefined : v }), "グラフィックの動作")) : null,
      line("開始", choice(Object.entries(E.ANIM_STARTS), e.start, (v) => setOption({ start: v }), "開始")),
      line("継続時間", sec("dur", "継続時間（秒）", { min: 0.01 }), h("small", {}, "秒"), choice([["", "速さ…"], ...E.ANIM_SPEEDS.map(([ms, l]) => [ms, l])], "", (v) => v && setOption({ dur: Number(v) }), "速さ")),
      line("遅延", sec("delay", "遅延（秒）"), h("small", {}, "秒")),
      e.cls !== "media" ? line("繰り返し", choice(Object.entries(E.ANIM_REPEATS), e.repeat ?? 1, (v) => setOption({ repeat: v === "1" ? undefined : v === "click" || v === "slide" ? v : Number(v) }), "繰り返し")) : null,
      e.cls !== "media" ? line("", h("label", { class: "fp-check" }, h("input", { type: "checkbox", checked: e.rewind || null, onchange: (event) => setOption({ rewind: event.target.checked || undefined }) }), "再生が終了したら巻き戻す")) : null,
      e.cls !== "media" && e.cls !== "in" && e.cls !== "out" ? line("", h("label", { class: "fp-check" }, h("input", { type: "checkbox", checked: e.autoReverse || null, onchange: (event) => setOption({ autoReverse: event.target.checked || undefined }) }), "自動的に元に戻す（往復）")) : null,
      e.cls !== "media" ? line("滑らかさ", choice(Object.entries(E.ANIM_EASES), e.ease || "auto", (v) => setOption({ ease: v === "auto" ? undefined : v }), "滑らかさ")) : null,
      // 効果のオプション: アニメーションの後 (faded, hidden) and a sound as it starts.
      e.cls !== "media" && e.cls !== "out" ? line("アニメーションの後", choice(Object.entries(E.ANIM_AFTERS), e.after || "none", (v) => setOption({ after: v === "none" ? undefined : v }), "アニメーションの後")) : null,
      line("サウンド", choice([["", "［サウンドなし］"], ...Object.entries(E.TRANSITION_SOUNDS || {}).filter(([k]) => k !== "stop")], e.sound || "", (v) => { setOption({ sound: v || undefined }); if (v) E.playSound?.(v); }, "サウンド")),
      line("トリガー", choice(triggers, e.trigger || "", (v) => setOption({ trigger: v || undefined }), "トリガー")),
      e.cls === "path" ? line("パス",
        h("button", { type: "button", class: "btn btn-sm", onclick: () => setOption({ path: reversePath(e.path) }) }, "反転"),
        h("button", { type: "button", class: "btn btn-sm", onclick: () => setOption({ path: { ...e.path, closed: e.path.closed ? undefined : true } }) }, e.path.closed ? "開く" : "閉じる"),
        h("button", { type: "button", class: "btn btn-sm", onclick: () => setOption({ path: { ...e.path, curve: e.path.curve ? undefined : true } }) }, e.path.curve ? "直線" : "曲線"),
        h("button", { type: "button", class: "btn btn-sm", onclick: () => startDrawing(e.el, e.id) }, "描き直す")) : null,
      e.cls === "path" ? h("p", { class: "hint" }, "スライド上の点（緑＝始点・赤＝終点）をドラッグして形を変えられます。Shiftを押しながら終点を動かすと、パス全体が回転・拡大します。") : null);
    return box;
  }

  // ---------------------------------------------------------------- marks on the stage

  const showing = () => onSlide() && (app.panel() === "anim" || tabNow() === "animation");
  /** Order numbers by each animated thing and the motion paths of what is picked (drawn on the editor's overlay). */
  function overlay(list, k) {
    if (!showing() || !editor.state.slideEl) return [];
    const out = [];
    const p = plan();
    const marks = new Map();
    const note = (el, text, on) => { if (!marks.has(el)) marks.set(el, []); marks.get(el).push({ text, on }); };
    p.main.forEach((grp, g) => grp.items.forEach((item) => note(item.e.el, String(g), picked.includes(item.e.id))));
    for (const seq of p.triggers.values()) for (const grp of seq) for (const item of grp.items) note(item.e.el, "⚡", picked.includes(item.e.id));
    for (const [el, items] of marks) {
      const b = markBox(el, list);
      if (!b) continue;
      const texts = [...new Set(items.map((i) => i.text))];
      out.push(h("div", { class: ["ed-anim-badge", items.some((i) => i.on) ? "on" : ""], style: { left: `${b.x * k - 4}px`, top: `${b.y * k - 4}px` } }, texts.slice(0, 4).join(",") + (texts.length > 4 ? "…" : "")));
    }
    // Motion paths: the picked ones, or those of the selected objects.
    const sel = new Set(editor.selection);
    const paths = p.list.filter((e) => e.cls === "path" && (picked.includes(e.id) || targetsOf(e).some((id) => sel.has(id))));
    for (const e of paths) out.push(...pathMarks(e, list, k, picked.length === 1 && picked[0] === e.id));
    return out;
  }
  function markBox(el, list) {
    if (el.startsWith("@")) {
      const slideEl = editor.state.slideEl;
      const node = el === "@title" ? slideEl.querySelector('[data-field="title"]') : el === "@takeaway" ? slideEl.querySelector('[data-field="takeaway"]') : slideEl.querySelector(`[data-g="${Number(el.slice(2))}"]`);
      if (!node) return null;
      const base = slideEl.getBoundingClientRect();
      const r = node.getBoundingClientRect();
      const k = base.width / E.W || 1;
      return { x: (r.left - base.left) / k, y: (r.top - base.top) / k, w: r.width / k, h: r.height / k };
    }
    const members = el.startsWith("grp:") ? list.filter((o) => o.group === el.slice(4) && !o.hidden) : list.filter((o) => o.id === el && !o.hidden);
    return members.length ? ops.bounds(members, (o) => E.lineEnds(o, list)) : null;
  }
  function pathMarks(e, list, k, editable) {
    const b = markBox(e.el, list);
    if (!b) return [];
    const path = dragPath?.id === e.id ? { ...e.path, pts: dragPath.pts } : e.path;
    const cx = b.x + b.w / 2;
    const cy = b.y + b.h / 2;
    const svg = E.s("svg", { class: "ed-anim-path", width: E.W * k, height: E.H * k, viewBox: `0 0 ${E.W} ${E.H}`, "aria-hidden": "true" });
    svg.append(E.s("path", { d: E.pathD(path), transform: `translate(${cx} ${cy})`, fill: "none", stroke: "#1f3864", "stroke-width": 2.5 / k, "stroke-dasharray": `${8 / k} ${6 / k}`, "vector-effect": "non-scaling-stroke" }));
    const [ex, ey] = E.pathEnd(path);
    // A ghost of the object where the path ends.
    svg.append(E.s("rect", { x: cx + ex - b.w / 2, y: cy + ey - b.h / 2, width: b.w, height: b.h, fill: "none", stroke: "#1f3864", "stroke-opacity": 0.35, "stroke-width": 1.5 / k, "stroke-dasharray": `${4 / k} ${4 / k}`, "vector-effect": "non-scaling-stroke" }));
    const out = [svg];
    const dot = (x, y, color, handle, title) => h("span", { class: ["ed-handle", "ed-anim-pt", handle ? "" : "fixed"], "data-handle": handle || null, title, style: { left: `${(cx + x) * k}px`, top: `${(cy + y) * k}px`, "--c": color } });
    out.push(dot(0, 0, PATH_COLORS.start, null, "始点（図形の中心）"));
    if (editable) path.pts.forEach(([x, y], i) => { if (i > 0) out.push(dot(x, y, i === path.pts.length - 1 && !path.closed ? PATH_COLORS.end : "#1f3864", `anim:${e.id}:${i}`, i === path.pts.length - 1 ? "終点（Shiftで全体を回転・拡大）" : "点をドラッグ")); });
    else if (!path.closed) out.push(dot(ex, ey, PATH_COLORS.end, null, "終点"));
    return out;
  }

  // Dragging a point of the picked path.
  editor.handle("anim:", (event, handle) => {
    const [, id, index] = handle.split(":");
    const e = timeline().find((x) => x.id === id);
    if (!e?.path) return;
    const i = Number(index);
    const b = markBox(e.el, objects());
    const center = [b.x + b.w / 2, b.y + b.h / 2];
    const start = e.path.pts.map((p) => [...p]);
    dragPath = { id, pts: start.map((p) => [...p]) };
    const onMove = (ev) => {
      const [px, py] = editor.toSlide(ev);
      const local = [px - center[0], py - center[1]];
      if (ev.shiftKey && i === start.length - 1) {
        // Turn and stretch the whole path about its start so its end follows the pointer.
        const [ox, oy] = start[i];
        const a = Math.atan2(local[1], local[0]) - Math.atan2(oy, ox);
        const s = Math.hypot(...local) / (Math.hypot(ox, oy) || 1);
        dragPath.pts = start.map(([x, y]) => [Math.round((x * Math.cos(a) - y * Math.sin(a)) * s), Math.round((x * Math.sin(a) + y * Math.cos(a)) * s)]);
      } else dragPath.pts = start.map((p, j) => (j === i ? [Math.round(local[0]), Math.round(local[1])] : p));
      editor.draw();
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      const pts = dragPath.pts;
      dragPath = null;
      if (JSON.stringify(pts) !== JSON.stringify(start)) commit(update([id], (x) => ({ path: { ...x.path, pts } })), { pick: [id] });
      else editor.draw();
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  });

  // ---------------------------------------------------------------- drawing a path of one's own

  /**
   * ユーザー設定パス: click to put points (a press-and-drag draws freehand), double-click or Enter to finish,
   * Esc to stop. The path starts at the object's centre. `replace` redraws an existing path animation.
   */
  function startDrawing(el, replace = null) {
    const wrap = editor.state.wrap;
    if (!wrap) return;
    stopDrawing();
    const b = markBox(el, objects());
    if (!b) return;
    const center = [b.x + b.w / 2, b.y + b.h / 2];
    const pts = [[0, 0]];
    let hover = null;
    let freehand = false;
    const layer = h("div", { class: "ed-pathdraw", title: "クリックで点を打つ・ドラッグで自由に描く・ダブルクリックかEnterで終わる・Escでやめる" });
    const svg = E.s("svg", { viewBox: `0 0 ${E.W} ${E.H}`, preserveAspectRatio: "none" });
    layer.append(svg);
    wrap.append(layer);
    const local = (ev) => { const [x, y] = editor.toSlide(ev); return [Math.round(x - center[0]), Math.round(y - center[1])]; };
    const paint = () => {
      const all = hover ? [...pts, hover] : pts;
      svg.replaceChildren(E.s("path", { d: all.map(([x, y], i) => `${i ? "L" : "M"}${x + center[0]} ${y + center[1]}`).join(" "), fill: "none", stroke: "#1f3864", "stroke-width": 4, "stroke-dasharray": "10 7" }),
        E.s("circle", { cx: center[0], cy: center[1], r: 9, fill: PATH_COLORS.start }));
    };
    const finish = () => {
      if (pts.length < 2) { app.toast("2つ以上の点を打ってください"); return; }
      const simple = simplify(pts);
      stopDrawing();
      const path = { pts: simple, ...(freehand ? { curve: true } : {}) };
      if (replace) commit(update([replace], { cls: "path", fx: "custom", path }), { pick: [replace], preview: timeline().filter((x) => x.id === replace) });
      else { const e = { ...makeEntry(el, "path", "custom", "click"), path }; commit([...timeline(), e], { pick: [e.id], preview: [e] }); }
    };
    let pressed = false;
    layer.addEventListener("pointerdown", (ev) => { ev.preventDefault(); ev.stopPropagation(); pressed = true; layer.setPointerCapture(ev.pointerId); const p = local(ev); const last = pts[pts.length - 1]; if (Math.hypot(p[0] - last[0], p[1] - last[1]) > 4) pts.push(p); paint(); });
    layer.addEventListener("pointermove", (ev) => {
      const p = local(ev);
      if (pressed) { const last = pts[pts.length - 1]; if (Math.hypot(p[0] - last[0], p[1] - last[1]) > 14) { pts.push(p); freehand = true; } } else hover = p;
      paint();
    });
    layer.addEventListener("pointerup", () => { pressed = false; });
    layer.addEventListener("dblclick", (ev) => { ev.preventDefault(); ev.stopPropagation(); finish(); });
    const key = (ev) => { if (ev.key === "Escape") { ev.preventDefault(); ev.stopPropagation(); stopDrawing(); app.toast("パスを描くのをやめました"); } else if (ev.key === "Enter") { ev.preventDefault(); ev.stopPropagation(); finish(); } };
    document.addEventListener("keydown", key, true);
    drawing = { layer, key };
    paint();
    app.toast("クリックで点を打ち、ダブルクリックかEnterで終わります（ドラッグで自由に描けます・Escでやめる）");
  }
  function stopDrawing() {
    if (!drawing) return;
    drawing.layer.remove();
    document.removeEventListener("keydown", drawing.key, true);
    drawing = null;
  }
  /** Fewer points for a freehand line (Ramer–Douglas–Peucker, 6 px). */
  function simplify(pts, tolerance = 6) {
    if (pts.length < 3) return pts;
    const keep = new Array(pts.length).fill(false);
    keep[0] = keep[pts.length - 1] = true;
    const stack = [[0, pts.length - 1]];
    while (stack.length) {
      const [a, b] = stack.pop();
      let far = -1;
      let dist = 0;
      for (let i = a + 1; i < b; i += 1) {
        const d = ops.distToSegment(pts[i], pts[a], pts[b]);
        if (d > dist) { dist = d; far = i; }
      }
      if (far >= 0 && dist > tolerance) { keep[far] = true; stack.push([a, far], [far, b]); }
    }
    return pts.filter((_, i) => keep[i]);
  }

  editor.overlay(overlay);
  editor.subscribe(() => {
    // Picking an object on the stage picks its animations in the pane.
    prunePicked();
    const targets = new Set(selectionTargets());
    const sel = new Set(editor.selection);
    // A part of the layout (@title…) has no object to select: picking its row empties the selection, and it stays picked.
    if (picked.length && !timeline().filter((e) => picked.includes(e.id)).every((e) => targets.has(e.el) || targetsOf(e).some((id) => sel.has(id)) || (e.el.startsWith("@") && !sel.size))) picked = [];
    renderPane();
  });

  return { animationTab, transitionTab, renderPane, stopDrawing, get picked() { return [...picked]; } };
}
