// Cropping a picture on the slide, PowerPoint style (図の形式 → トリミング): the whole picture shows faintly,
// the part kept is framed with black crop marks; drag the marks to crop, drag the picture to choose what shows,
// Enter or a click outside keeps it, Esc puts it back. Rotated pictures are cropped in the 書式 pane.

import { ico } from "./icons.mjs";

const MARKS = ["nw", "n", "ne", "e", "se", "s", "sw", "w"];

// トリミング ▾ (PowerPoint's): 縦横比 crops to a shape of picture, 塗りつぶし fills the frame (cropping what is over),
// 枠に合わせる shows the whole picture inside the frame. `aspect` is the picture's own width / height.
const fullOf = (o) => { const c = o.crop || { l: 0, t: 0, r: 0, b: 0 }; return { fw: o.w / (1 - c.l - c.r), fh: o.h / (1 - c.t - c.b) }; };
const round4 = (v) => Math.round(v * 10000) / 10000;
const tidy = (c) => (c.l || c.t || c.r || c.b ? { l: round4(c.l), t: round4(c.t), r: round4(c.r), b: round4(c.b) } : undefined);

/** 縦横比: the largest middle part of the picture of that shape; the frame takes its size (the picture's scale kept). */
export function cropToAspect(o, ratio, aspect = null) {
  const { fw } = fullOf(o);
  const a = aspect || fw / fullOf(o).fh;
  const fh = fw / a;
  const [rw, rh] = a > ratio ? [fh * ratio, fh] : [fw, fw / ratio];
  const cx = o.x + o.w / 2;
  const cy = o.y + o.h / 2;
  const crop = tidy({ l: (fw - rw) / 2 / fw, r: (fw - rw) / 2 / fw, t: (fh - rh) / 2 / fh, b: (fh - rh) / 2 / fh });
  return { x: cx - rw / 2, y: cy - rh / 2, w: rw, h: rh, crop };
}
/** 塗りつぶし: the frame stays; the picture fills it, its overhang cropped evenly. */
export function cropFill(o, aspect) {
  const frame = o.w / o.h;
  if (aspect > frame) { const keep = frame / aspect; return { crop: tidy({ l: (1 - keep) / 2, r: (1 - keep) / 2, t: 0, b: 0 }) }; }
  const keep = aspect / frame;
  return { crop: tidy({ l: 0, r: 0, t: (1 - keep) / 2, b: (1 - keep) / 2 }) };
}
/** 枠に合わせる: the whole picture, as large as fits in the frame (its middle where the frame's was). */
export function cropFit(o, aspect) {
  const frame = o.w / o.h;
  const [w, h] = aspect > frame ? [o.w, o.w / aspect] : [o.h * aspect, o.h];
  return { x: o.x + (o.w - w) / 2, y: o.y + (o.h - h) / 2, w, h, crop: undefined };
}

export function createCrop(editor, app) {
  const { E, h } = app;
  let crop = null;

  const byId = (id) => editor.objects().find((o) => o.id === id);
  const node = (id) => editor.state.slideEl?.querySelector(`.hs-obj[data-el="${CSS.escape(id)}"]`);

  function start(id) {
    const o = byId(id);
    if (!o || o.kind !== "image") { app.toast("トリミングする画像を1つ選んでください"); return; }
    if (o.rot) { app.toast("回転した画像は、書式パネルの「トリミング（%）」で切り取れます"); app.openPanel("format", "picture"); return; }
    if (o.locked) return;
    const c = o.crop || { l: 0, t: 0, r: 0, b: 0 };
    const fw = o.w / (1 - c.l - c.r);
    const fh = o.h / (1 - c.t - c.b);
    const src = node(id)?.querySelector("img")?.src;
    if (!src) return;
    crop = { id, src, full: { x: o.x - c.l * fw, y: o.y - c.t * fh, w: fw, h: fh }, box: { x: o.x, y: o.y, w: o.w, h: o.h } };
    node(id)?.classList.add("ed-cropping");
    editor.state.wrap?.classList.add("ed-crop-mode");
    document.addEventListener("keydown", onKey, true);
    document.addEventListener("pointerdown", onOutside, true);
    editor.draw();
    app.toast("黒い印をドラッグして切り取り、画像をドラッグして位置を決めます（Enterで確定・Escで戻す）");
    editor.emit();
  }
  function finish(keep = true) {
    if (!crop) return;
    const { id, full, box } = crop;
    crop = null;
    document.removeEventListener("keydown", onKey, true);
    document.removeEventListener("pointerdown", onOutside, true);
    node(id)?.classList.remove("ed-cropping");
    editor.state.wrap?.classList.remove("ed-crop-mode");
    if (keep) {
      const c = { l: (box.x - full.x) / full.w, t: (box.y - full.y) / full.h, r: (full.x + full.w - box.x - box.w) / full.w, b: (full.y + full.h - box.y - box.h) / full.h };
      const r4 = (v) => Math.max(0, Math.round(v * 10000) / 10000);
      const next = { l: r4(c.l), t: r4(c.t), r: r4(c.r), b: r4(c.b) };
      const none = !next.l && !next.t && !next.r && !next.b;
      editor.commit(editor.objects().map((o) => (o.id === id ? { ...o, x: box.x, y: box.y, w: box.w, h: box.h, crop: none ? undefined : next } : o)), { select: [id] });
    } else editor.draw();
    editor.emit();
  }
  function onKey(event) {
    if (event.key === "Enter") { event.preventDefault(); event.stopPropagation(); finish(true); }
    else if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); finish(false); }
  }
  function onOutside(event) {
    if (event.target.closest?.(".ed-crop, .ribbon, .rb-pop")) return;
    finish(true);
  }

  function overlay(list, k) {
    if (!crop) return [];
    const { full, box, src } = crop;
    const ghost = h("img", { class: "ed-crop-ghost", src, alt: "", draggable: "false", style: { left: `${full.x * k}px`, top: `${full.y * k}px`, width: `${full.w * k}px`, height: `${full.h * k}px` } });
    const kept = h("div", { class: "ed-handle ed-crop ed-crop-box", "data-handle": "crop:pan", title: "ドラッグで見える位置を決める", style: { left: `${box.x * k}px`, top: `${box.y * k}px`, width: `${box.w * k}px`, height: `${box.h * k}px` } },
      h("img", { src, alt: "", draggable: "false", style: { left: `${(full.x - box.x) * k}px`, top: `${(full.y - box.y) * k}px`, width: `${full.w * k}px`, height: `${full.h * k}px` } }));
    for (const m of MARKS) kept.append(h("span", { class: `ed-handle ed-crop ed-crop-mark ed-crop-${m}`, "data-handle": `crop:${m}`, title: "ドラッグで切り取る" }));
    return [ghost, kept];
  }
  editor.overlay(overlay);
  editor.handle("crop:", (event, handle) => {
    if (!crop) return;
    const kind = handle.slice(5);
    const [sx, sy] = editor.toSlide(event);
    const box0 = { ...crop.box };
    const full0 = { ...crop.full };
    const min = 12;
    const move = (ev) => {
      if (!crop) return;
      const [x, y] = editor.toSlide(ev);
      const dx = x - sx;
      const dy = y - sy;
      if (kind === "pan") {
        // The picture slides under the frame, never leaving a gap inside it.
        const fx = Math.min(box0.x, Math.max(box0.x + box0.w - full0.w, full0.x + dx));
        const fy = Math.min(box0.y, Math.max(box0.y + box0.h - full0.h, full0.y + dy));
        crop.full = { ...full0, x: fx, y: fy };
      } else {
        let { x: bx, y: by, w: bw, h: bh } = box0;
        const right = box0.x + box0.w;
        const bottom = box0.y + box0.h;
        if (kind.includes("w")) { bx = Math.max(full0.x, Math.min(right - min, box0.x + dx)); bw = right - bx; }
        if (kind.includes("e")) bw = Math.max(min, Math.min(full0.x + full0.w - box0.x, box0.w + dx));
        if (kind.includes("n")) { by = Math.max(full0.y, Math.min(bottom - min, box0.y + dy)); bh = bottom - by; }
        if (kind.includes("s")) bh = Math.max(min, Math.min(full0.y + full0.h - box0.y, box0.h + dy));
        crop.box = { x: bx, y: by, w: bw, h: bh };
      }
      editor.draw();
    };
    const up = () => { window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", up); };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  });

  return { start, finish, get active() { return Boolean(crop); }, icon: () => ico("crop") };
}
