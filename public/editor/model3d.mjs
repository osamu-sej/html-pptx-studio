// 挿入 → 3D モデル and the 3D モデル tab (PowerPoint's 3D Models): a GLB / glTF file from this device or a built-in
// model (ストック 3D モデル); its view — the 3D モデル ビュー gallery, or turning it on the slide by the handle in its
// middle — パンとズーム (drag inside to move it in its box, the magnifier or the wheel to zoom), 3D モデルのリセット and
// 3D モデルの変更. Each change draws its poster again (the still picture thumbnails and printing use); on the stage
// and in a slide show the model is drawn live (engine/models.js).

import { ico } from "./icons.mjs";

const clampView = (v) => ({
  yaw: Math.max(-360, Math.min(360, v.yaw || 0)), pitch: Math.max(-89, Math.min(89, v.pitch || 0)), roll: v.roll || 0,
  zoom: Math.max(0.3, Math.min(4, v.zoom ?? 1)), panX: Math.max(-2, Math.min(2, v.panX || 0)), panY: Math.max(-2, Math.min(2, v.panY || 0)),
});
const viewOf = (o) => clampView({ ...(o?.view || {}) });
const tidy = (v) => Object.fromEntries(Object.entries(clampView(v)).filter(([k, x]) => (k === "zoom" ? Math.abs(x - 1) > 0.001 : Math.abs(x) > 0.001)).map(([k, x]) => [k, Math.round(x * 100) / 100]));

export function createModels(editor, app, kit) {
  const { h, E } = app;
  const { btn, drop, group, col, menu } = kit;
  const models = () => editor.selectedObjects().filter((o) => o.kind === "model");
  const one = () => (models().length === 1 ? models()[0] : null);
  let panMode = false;
  const urlOf = (src) => (String(src).startsWith("builtin:") ? src : app.imageUrl(src));
  const liveBox = (id) => document.querySelector(`#stageBody .hs-obj[data-el="${id}"] .hs-model`);

  /** The still picture of a model at a view (its box's shape), or undefined when it cannot be drawn. */
  async function posterOf(o, view = viewOf(o)) {
    try {
      const width = 640;
      return await E.modelPoster(urlOf(o.src), tidy(view), { width, height: Math.max(80, Math.min(1600, Math.round((width * o.h) / Math.max(1, o.w)))) });
    } catch (error) {
      app.toast(`3D モデルを表示できません（${error.message || error}）`);
      return undefined;
    }
  }
  /** Set a model's view (one undo step) with its new poster. */
  async function setView(o, view) {
    const next = tidy(view);
    const poster = await posterOf(o, next);
    editor.apply((x) => (x.id === o.id ? { view: Object.keys(next).length ? next : undefined, ...(poster ? { poster } : {}) } : null), { ids: [o.id] });
  }

  // ---------------------------------------------------------------- inserting

  async function insertModel(src, { fileName = null, name = null } = {}) {
    const side = 560;
    const o = { id: E.newObjectId(), kind: "model", x: 0, y: 0, w: side, h: side, src, view: { yaw: -30, pitch: 18 }, ...(fileName ? { fileName } : {}), ...(name ? { name } : {}) };
    app.toast("3D モデルを読み込んでいます…");
    const poster = await posterOf(o);
    if (!poster) return null;
    const [placed] = editor.insert([{ ...o, poster }]);
    app.toast("3D モデルを入れました（中央のハンドルで回転、3D モデル タブでビュー・パンとズーム）");
    return placed;
  }
  async function fromDevice({ replace = null } = {}) {
    const [file] = await app.pickFiles(".glb,.gltf,model/gltf-binary,model/gltf+json", false);
    if (!file) return;
    if (file.size > 100_000_000) { app.toast("3D モデルが大きすぎます（100MBまで）"); return; }
    if (/\.gltf$/i.test(file.name)) {
      // A glTF must carry its buffers and pictures inside (data:); otherwise use the .glb export of the model.
      const text = await file.text();
      if (/"uri"\s*:\s*"(?!data:)/.test(text)) { app.toast("ほかのファイルを参照する glTF は入れられません（GLB で書き出してください）"); return; }
    }
    const type = /\.gltf$/i.test(file.name) ? "model/gltf+json" : "model/gltf-binary";
    const src = await app.storeBlob(new Blob([await file.arrayBuffer()], { type }), file.name);
    if (replace) {
      const poster = await posterOf({ ...replace, src });
      if (!poster) return;
      editor.apply((x) => (x.id === replace.id ? { src, fileName: file.name, poster } : null), { ids: [replace.id] });
      app.toast("3D モデルを変更しました（⌘Zで戻せます）");
      return;
    }
    await insertModel(src, { fileName: file.name });
  }
  /** 挿入 → 3D モデル ▾ */
  function insertMenu() {
    return menu([
      { label: "このデバイス…", icon: "model3d", run: () => fromDevice() },
      "-", { head: "ストック 3D モデル" },
      ...Object.entries(E.MODEL_BUILTINS).map(([key, label]) => ({ label, icon: "model3d", run: () => insertModel(`builtin:${key}`, { name: label }) })),
    ]);
  }

  // ---------------------------------------------------------------- the 3D モデル tab

  function viewGallery() {
    return (close) => {
      const o = one();
      return h("div", { class: "rb-gallery m3d-views" }, h("div", { class: "rb-menu-head" }, "3D モデル ビュー"),
        h("div", { class: "m3d-view-grid" }, Object.entries(E.MODEL_VIEWS).map(([key, [label, yaw, pitch]]) => h("button", {
          type: "button", class: "m3d-view", "data-view": key, title: label,
          onclick: () => { close(); for (const m of models()) setView(m, { ...viewOf(m), yaw, pitch }); },
        }, h("span", { class: "m3d-cube", style: { transform: `rotateX(${-pitch}deg) rotateY(${-yaw}deg)` } }, ...["f", "b", "l", "r", "t", "d"].map((f) => h("i", { class: `m3d-face m3d-${f}` }))), h("small", {}, label)))),
        o ? null : h("p", { class: "hint" }, "3D モデルを1つ選ぶと、ビューを選べます。"));
    };
  }
  function tab() {
    const any = () => models().length > 0;
    return [
      group("調整",
        btn("change", "3D モデル|の変更", "別の GLB / glTF ファイルに差し替える（サイズ・ビュー・アニメーションはそのまま）", () => { const o = one(); if (o) fromDevice({ replace: o }); }, { big: true, enabled: () => Boolean(one()) }),
        drop("reset", "3D モデル|のリセット", "ビュー（回転・パンとズーム）を最初に戻す", () => menu([
          { label: "3D モデルのリセット", icon: "reset", run: () => { for (const m of models()) setView(m, { yaw: -30, pitch: 18 }); } },
          { label: "3D モデルとサイズのリセット", icon: "reset", run: async () => {
            for (const m of models()) {
              const side = Math.round(Math.max(m.w, m.h));
              const next = { ...m, w: side, h: side };
              const poster = await posterOf(next, { yaw: -30, pitch: 18 });
              editor.apply((x) => (x.id === m.id ? { w: side, h: side, view: { yaw: -30, pitch: 18 }, ...(poster ? { poster } : {}) } : null), { ids: [m.id] });
            }
          } },
        ]), { big: true, enabled: any })),
      group("3D モデル ビュー", drop("model3d", "ビュー", "正面・背面・左・右・上・下・斜めから見る", viewGallery(), { big: true, enabled: any })),
      group("サイズ",
        btn("zoomIn", "パンと|ズーム", "オン：モデルをドラッグして枠の中で動かし、右の虫めがね（またはホイール）で拡大・縮小", () => { panMode = !panMode; editor.draw(); kit.refreshRibbon?.(); }, { big: true, pressed: () => panMode, enabled: () => Boolean(one()) })),
      group("回転",
        col(numberField("横の回転", "yaw", -360, 360), numberField("縦の回転", "pitch", -89, 89))),
    ];
  }
  function numberField(label, key, min, max) {
    const input = h("input", { type: "number", class: "rb-num", min, max, step: 5, "aria-label": `${label}（度）`,
      onchange: () => { const o = one(); if (o) setView(o, { ...viewOf(o), [key]: Number(input.value) || 0 }); } });
    kit.updater(() => { const o = one(); input.disabled = !o; if (o && document.activeElement !== input) input.value = String(Math.round(viewOf(o)[key])); });
    return h("label", { class: "rb-field" }, h("span", {}, label), input, h("span", {}, "°"));
  }

  // ---------------------------------------------------------------- on the slide: turn, pan, zoom

  editor.overlay((list, k) => {
    const o = list.find((x) => x.kind === "model" && editor.selection.length === 1 && editor.selection[0] === x.id && !x.locked && !x.hidden);
    if (!o) return [];
    const cx = (o.x + o.w / 2) * k;
    const cy = (o.y + o.h / 2) * k;
    const out = [];
    if (panMode) {
      out.push(h("span", { class: "ed-handle m3d-pan", "data-handle": `m3d:pan:${o.id}`, title: "ドラッグで枠の中を動かす（ホイールで拡大・縮小）", style: { left: `${o.x * k}px`, top: `${o.y * k}px`, width: `${o.w * k}px`, height: `${o.h * k}px` },
        onwheel: (event) => { event.preventDefault(); zoomBy(o, event.deltaY < 0 ? 1.1 : 1 / 1.1); } }));
      out.push(h("span", { class: "ed-handle m3d-zoom", "data-handle": `m3d:zoom:${o.id}`, title: "上下にドラッグで拡大・縮小", style: { left: `${(o.x + o.w) * k + 22}px`, top: `${cy}px` }}, ico("zoomIn", 14)));
    }
    // In パンとズーム, a drag inside moves the model: the turning handle steps aside.
    if (!panMode) out.push(h("span", { class: "ed-handle m3d-rotate", "data-handle": `m3d:rot:${o.id}`, title: "ドラッグで3Dに回転（Shiftで15°ずつ）", style: { left: `${cx}px`, top: `${cy}px` }}, ico("rotate3d", 16)));
    return out;
  });

  let zoomTimer = 0;
  function zoomBy(o, factor) {
    const box = liveBox(o.id);
    const view = JSON.parse(box?.dataset.view || JSON.stringify(viewOf(o)));
    view.zoom = Math.max(0.3, Math.min(4, (view.zoom ?? 1) * factor));
    if (box) box.dataset.view = JSON.stringify(view);
    clearTimeout(zoomTimer);
    zoomTimer = setTimeout(() => setView(o, view), 350);
  }

  editor.handle("m3d:", (event, handle) => {
    const [, what, id] = handle.split(":");
    const o = editor.objects().find((x) => x.id === id);
    if (!o) return;
    event.preventDefault();
    event.stopPropagation();
    const start = viewOf(o);
    const box = liveBox(id);
    const [sx, sy] = [event.clientX, event.clientY];
    const k = editor.scale();
    let view = start;
    const onMove = (ev) => {
      const dx = ev.clientX - sx;
      const dy = ev.clientY - sy;
      if (what === "rot") {
        let yaw = start.yaw + dx * 0.6;
        let pitch = start.pitch + dy * 0.4;
        if (ev.shiftKey) { yaw = Math.round(yaw / 15) * 15; pitch = Math.round(pitch / 15) * 15; }
        view = clampView({ ...start, yaw: ((yaw + 540) % 360) - 180, pitch });
      } else if (what === "pan") view = clampView({ ...start, panX: start.panX - (dx / (o.w * k)) * 2, panY: start.panY + (dy / (o.h * k)) * 2 });
      else if (what === "zoom") view = clampView({ ...start, zoom: start.zoom * Math.exp(-dy / 160) });
      if (box) box.dataset.view = JSON.stringify(view);
    };
    const onUp = () => {
      window.removeEventListener("pointermove", onMove);
      window.removeEventListener("pointerup", onUp);
      if (JSON.stringify(view) !== JSON.stringify(start)) setView(o, view);
    };
    window.addEventListener("pointermove", onMove);
    window.addEventListener("pointerup", onUp);
  });

  return { insertMenu, tab, isModel: () => models().length > 0, insertModel, fromDevice, setView, get panMode() { return panMode; } };
}
