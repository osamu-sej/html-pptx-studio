// 配布資料マスター・ノート マスター (表示 → マスター表示): the page that handouts and notes pages are printed on — its
// orientation, where 配布資料 slides sit, where the slide and the notes sit on a notes page, the header, date, footer and
// page number (on or off, their words, where they stand), and the words and pictures every page carries.
// The functions are pure; the dialog (createPrintMasters) edits the stored form, `deck.printMasters = { handout, notes }`,
// which keeps only what differs from the studio's own page (the one print.mjs has always drawn).

import * as ops from "./ops.mjs";

export const MARGIN = 48;
export const PAGES = { portrait: { w: 794, h: 1123 }, landscape: { w: 1123, h: 794 } };
export const pageOf = (orientation) => (orientation === "landscape" ? PAGES.landscape : PAGES.portrait);
export const KINDS = { handout: "配布資料マスター", notes: "ノート マスター" };
export const HF_KEYS = ["header", "date", "footer", "pageNo"];
export const HF_NAMES = { header: "ヘッダー", date: "日付", footer: "フッター", pageNo: "ページ番号" };
export const HANDOUT_LAYOUTS = { h1: 1, h2: 2, h3: 3, h4: 4, h6: 6, h9: 9 };
export const ALIGN_NAMES = { left: "左", center: "中央", right: "右" };
const HF_ON = { header: true, date: true, footer: false, pageNo: true };
export const MAX_OBJECTS = 30;
export const FS = { min: 8, max: 24, normal: 11 };

const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
const whole = (v, lo, hi) => { const n = Math.round(Number(v)); return Number.isFinite(n) ? clamp(n, lo, hi) : null; };

/** Where a placeholder stands until it is moved: the corners of the page, as the printouts have always had them. */
export function placeholderBox(key, orientation, fs = FS.normal) {
  const page = pageOf(orientation);
  const w = Math.round((page.w - MARGIN * 2) / 2);
  const left = key === "header" || key === "footer";
  const top = key === "header" || key === "date";
  return { x: left ? MARGIN : page.w - MARGIN - w, y: top ? 22 : page.h - 20 - Math.round(fs * 1.4), w, align: left ? "left" : "right" };
}
/** Where the slide and the notes stand on a notes page until they are moved. */
export function notesBoxes(orientation) {
  const page = pageOf(orientation);
  if (orientation === "landscape") return { slide: { x: MARGIN, y: 70, w: 520 }, notes: { x: 600, y: 70, w: page.w - 600 - MARGIN, h: page.h - 70 - 60, fs: 14 } };
  return { slide: { x: 78, y: 70, w: page.w - MARGIN * 2 - 60 }, notes: { x: 78, y: 470, w: page.w - 156, h: page.h - 470 - 60, fs: 14 } };
}

/**
 * The master as it is kept: only what is not the default, every number checked. `objects` is the list of words and pictures
 * (the studio passes the engine's own check for objects; they are kept in page pixels). Null when it is all default.
 */
export function normalizeMaster(value, kind, { objects = (list) => list } = {}) {
  if (!value || typeof value !== "object") return null;
  const out = {};
  const orientation = value.orientation === "landscape" ? "landscape" : "portrait";
  if (orientation === "landscape") out.orientation = "landscape";
  const page = pageOf(orientation);
  if (kind === "handout" && HANDOUT_LAYOUTS[value.perPage]) out.perPage = value.perPage;
  const fs = whole(value.fs, FS.min, FS.max);
  if (fs != null && fs !== FS.normal) out.fs = fs;
  for (const key of HF_KEYS) {
    const own = value[key];
    if (!own || typeof own !== "object") continue;
    const kept = {};
    if (typeof own.on === "boolean" && own.on !== HF_ON[key]) kept.on = own.on;
    if (typeof own.text === "string" && own.text.trim() && key !== "pageNo") kept.text = own.text.slice(0, 200);
    const base = placeholderBox(key, orientation, out.fs ?? FS.normal);
    const x = whole(own.x, 0, page.w - 40); const y = whole(own.y, 0, page.h - 10); const w = whole(own.w, 40, page.w);
    if (x != null && x !== base.x) kept.x = x;
    if (y != null && y !== base.y) kept.y = y;
    if (w != null && w !== base.w) kept.w = w;
    if (["left", "center", "right"].includes(own.align) && own.align !== base.align) kept.align = own.align;
    if (Object.keys(kept).length) out[key] = kept;
  }
  if (kind === "notes") {
    const base = notesBoxes(orientation);
    const slide = {};
    for (const [k, [lo, hi]] of Object.entries({ x: [0, page.w - 80], y: [0, page.h - 60], w: [120, page.w] })) { const v = whole(value.slide?.[k], lo, hi); if (v != null && v !== base.slide[k]) slide[k] = v; }
    if (Object.keys(slide).length) out.slide = slide;
    const notes = {};
    for (const [k, [lo, hi]] of Object.entries({ x: [0, page.w - 80], y: [0, page.h - 60], w: [120, page.w], h: [40, page.h], fs: [8, 32] })) { const v = whole(value.notes?.[k], lo, hi); if (v != null && v !== base.notes[k]) notes[k] = v; }
    if (Object.keys(notes).length) out.notes = notes;
  }
  const list = objects(Array.isArray(value.objects) ? value.objects.slice(0, MAX_OBJECTS) : []);
  if (list?.length) out.objects = list.slice(0, MAX_OBJECTS);
  return Object.keys(out).length ? out : null;
}
/** `deck.printMasters`, each master checked (null when there is none to keep). */
export function normalizeMasters(value, options) {
  const out = {};
  for (const kind of Object.keys(KINDS)) { const m = normalizeMaster(value?.[kind], kind, options); if (m) out[kind] = m; }
  return Object.keys(out).length ? out : null;
}

/** The master with every default filled in: the page, the placeholders and (notes) the two boxes, all in page pixels. */
export function resolveMaster(master, kind) {
  const m = master && typeof master === "object" ? master : {};
  const orientation = m.orientation === "landscape" ? "landscape" : "portrait";
  const fs = m.fs ?? FS.normal;
  const ph = {};
  for (const key of HF_KEYS) {
    const own = m[key] || {};
    const base = placeholderBox(key, orientation, fs);
    ph[key] = { on: own.on ?? HF_ON[key], text: own.text ?? "", x: own.x ?? base.x, y: own.y ?? base.y, w: own.w ?? base.w, align: own.align ?? base.align };
  }
  const out = { kind, orientation, page: pageOf(orientation), fs, ph, objects: Array.isArray(m.objects) ? m.objects : [], perPage: m.perPage || null };
  if (kind === "notes") {
    const base = notesBoxes(orientation);
    out.slide = { ...base.slide, ...(m.slide || {}) };
    out.notes = { ...base.notes, ...(m.notes || {}) };
  }
  return out;
}

/** The print dialog's header and footer fields as the master has them (the dialog may change them for one printing). */
export function printDefaults(resolved) {
  const { ph } = resolved;
  return { header: ph.header.on, headerText: ph.header.text, date: ph.date.on, pageNo: ph.pageNo.on, footer: ph.footer.on, footerText: ph.footer.text };
}

/** Every picture and object the masters hold, as pseudo-slides (so the media of a deck is found and carried there too). */
export function masterSlides(printMasters) {
  return Object.keys(KINDS).filter((kind) => printMasters?.[kind]?.objects?.length).map((kind) => ({ elements: printMasters[kind].objects }));
}

// ---------------------------------------------------------------- the dialog

const SCALE = { portrait: 0.42, landscape: 0.5 };
const clone = (v) => JSON.parse(JSON.stringify(v ?? {}));
const TEXT_COLORS = [["#000000", "黒"], ["#404040", "濃いグレー"], ["#808080", "グレー"]];

/**
 * 表示 → 配布資料マスター・ノート マスター: the page drawn small with everything on it that can be moved (drag it; the corner
 * handle sizes it), and the settings beside it. OK keeps the master in the deck (one undo step); the printing dialog starts from it.
 */
export function createPrintMasters(app) {
  const { h, E } = app;

  async function open(kind) {
    const deck = app.deck();
    if (!deck || !KINDS[kind]) return;
    const { handoutBoxes } = await import("./print.mjs");
    let work = clone(app.printMasters()[kind]);
    let selected = null;
    const slideCache = new Map();
    const dialog = h("dialog", { class: "pm-dialog", "aria-label": KINDS[kind] });
    const preview = h("div", { class: "pm-preview" });
    const sheetBox = h("div", { class: "pm-sheet" });
    const page = h("div", { class: "pm-page" });
    sheetBox.append(page);
    preview.append(sheetBox);
    const side = h("div", { class: "pm-form" });
    const resolved = () => resolveMaster(work, kind);
    const k = () => SCALE[resolved().orientation];

    /** A slide of the deck drawn `w` px wide (kept, so a drag does not draw it again). */
    function thumb(i, w) {
      const index = Math.max(0, Math.min(deck.slides.length - 1, i));
      if (!slideCache.has(index)) slideCache.set(index, app.renderSlide(index));
      const el = slideCache.get(index);
      el.style.transform = `scale(${w / E.W})`;
      el.style.transformOrigin = "0 0";
      return h("div", { class: "pm-thumb", style: { width: `${w}px`, height: `${Math.round((w * 9) / 16)}px` } }, el);
    }
    const today = () => new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long", day: "numeric" });

    // ---- where each movable thing is, and how a drag changes it
    function rectOf(id) {
      const r = resolved();
      if (HF_KEYS.includes(id)) { const p = r.ph[id]; return { x: p.x, y: p.y, w: p.w, h: Math.round(r.fs * 1.4) }; }
      if (id === "slide") return { ...r.slide, h: Math.round((r.slide.w * 9) / 16) };
      if (id === "notes") return { ...r.notes };
      const o = r.objects.find((x) => `o:${x.id}` === id);
      return o ? { x: o.x, y: o.y, w: o.w, h: o.h } : null;
    }
    function setRect(id, rect) {
      const r = resolved();
      const x = Math.round(clamp(rect.x, 0, r.page.w - 20));
      const y = Math.round(clamp(rect.y, 0, r.page.h - 10));
      const w = Math.round(clamp(rect.w, 24, r.page.w));
      if (HF_KEYS.includes(id)) work[id] = { ...(work[id] || {}), x, y, w };
      else if (id === "slide") work.slide = { x, y, w: Math.max(120, w) };
      else if (id === "notes") work.notes = { ...(work.notes || {}), x, y, w: Math.max(120, w), h: Math.round(clamp(rect.h, 40, r.page.h)) };
      else work.objects = (work.objects || []).map((o) => (`o:${o.id}` === id ? { ...o, x, y, w, h: Math.round(clamp(rect.h, 8, r.page.h)) } : o));
    }
    const nameOf = (id) => (HF_KEYS.includes(id) ? HF_NAMES[id] : id === "slide" ? "スライド" : id === "notes" ? "ノート" : "図・文字");

    // ---- the page
    function drawPreview() {
      const r = resolved();
      const scale = k();
      sheetBox.style.width = `${Math.round(r.page.w * scale)}px`;
      sheetBox.style.height = `${Math.round(r.page.h * scale)}px`;
      Object.assign(page.style, { width: `${r.page.w}px`, height: `${r.page.h}px`, transform: `scale(${scale})`, transformOrigin: "0 0" });
      page.replaceChildren();
      if (r.objects.length) {
        const layer = E.objectLayer({ elements: r.objects }, app.renderCtx?.() || {});
        if (layer) page.append(h("div", { class: "hs-slide pr-objects", "data-theme": E.PALETTE.sheet?.themeId || E.DEFAULT_THEME, style: { width: `${r.page.w}px`, height: `${r.page.h}px` } }, layer));
      }
      const frame = (id, box, inner, { resize = true, cls = "" } = {}) => {
        const el = h("div", { class: ["pm-box", cls, selected === id ? "sel" : ""], "data-pm": id, title: `${nameOf(id)}（ドラッグで動かす）`, style: { left: `${box.x}px`, top: `${box.y}px`, width: `${box.w}px`, height: `${box.h}px` } }, inner);
        if (resize && selected === id) el.append(h("i", { class: "pm-grip", "aria-label": "大きさを変える" }));
        page.append(el);
        return el;
      };
      if (kind === "handout") {
        const layout = r.perPage || "h6";
        const boxes = handoutBoxes(layout, HANDOUT_LAYOUTS[layout], { order: "rows", page: r.page });
        boxes.forEach((b, i) => {
          page.append(h("div", { class: "pm-slotbox", style: { left: `${b.x}px`, top: `${b.y}px`, width: `${b.w}px`, height: `${b.h}px` } }, thumb(i, b.w)));
          if (b.lines) page.append(h("div", { class: "pr-lines", style: { left: `${b.lines.x}px`, top: `${b.y}px`, width: `${b.lines.w}px`, height: `${b.h}px` } }));
        });
      } else {
        const s = rectOf("slide");
        frame("slide", s, thumb(app.index(), s.w));
        const n = r.notes;
        frame("notes", rectOf("notes"), h("div", { class: "pm-notes", style: { fontSize: `${n.fs}px` } }, [...Array(14)].map((_, i) => h("i", { style: { width: `${i % 4 === 3 ? 55 : 96}%` } }))), { cls: "pm-notesbox" });
      }
      for (const o of r.objects) frame(`o:${o.id}`, rectOf(`o:${o.id}`), null, { cls: "pm-obj" });
      for (const key of HF_KEYS) {
        const p = r.ph[key];
        const text = key === "header" ? p.text || deck.title || "ヘッダー" : key === "date" ? p.text || today() : key === "footer" ? p.text || "フッター" : "1";
        frame(key, rectOf(key), h("span", { style: { fontSize: `${r.fs}px`, textAlign: p.align } }, text), { resize: key !== "pageNo" && true, cls: ["pm-ph", p.on ? "" : "off"].join(" ") });
      }
    }

    // ---- dragging
    preview.addEventListener("pointerdown", (event) => {
      const target = event.target.closest?.("[data-pm]");
      if (!target) { if (selected) { selected = null; drawPreview(); drawForm(); } return; }
      event.preventDefault();
      const id = target.dataset.pm;
      const resizing = Boolean(event.target.closest?.(".pm-grip"));
      const select = selected !== id;
      selected = id;
      const from = rectOf(id);
      const [sx, sy] = [event.clientX, event.clientY];
      const scale = k();
      let moved = false;
      const move = (ev) => {
        const dx = (ev.clientX - sx) / scale;
        const dy = (ev.clientY - sy) / scale;
        if (!moved && Math.hypot(dx, dy) < 2) return;
        moved = true;
        setRect(id, resizing ? { ...from, w: from.w + dx, h: from.h + dy } : { ...from, x: from.x + dx, y: from.y + dy });
        drawPreview();
      };
      const up = () => { window.removeEventListener("pointermove", move); window.removeEventListener("pointerup", up); if (select || moved) { drawPreview(); drawForm(); } };
      window.addEventListener("pointermove", move);
      window.addEventListener("pointerup", up);
      if (select) { drawPreview(); drawForm(); }
    });

    // ---- the settings
    const sel = (value, entries, onset) => h("select", { onchange: (e) => onset(e.target.value) }, entries.map(([v, label]) => h("option", { value: v, selected: v === value || null }, label)));
    const numberField = (value, { min, max }, onset) => h("input", { type: "number", min, max, value, class: "pm-num", onchange: (e) => { const n = Math.round(Number(e.target.value)); if (Number.isFinite(n)) onset(clamp(n, min, max)); } });
    function selectedObject() { return selected?.startsWith("o:") ? resolved().objects.find((o) => `o:${o.id}` === selected) : null; }
    function changeObject(patch) { work.objects = (work.objects || []).map((o) => (`o:${o.id}` === selected ? E.normalizeObject({ ...o, ...patch }) || o : o)); drawPreview(); }
    async function addImage() {
      const [file] = (await app.pickFiles("image/png,image/jpeg,image/webp,image/svg+xml,image/gif", false)) || [];
      if (!file) return;
      try {
        const stored = await app.storePicture(file);
        const w = Math.min(220, stored.w || 220);
        const o = ops.makeObject("image", { x: MARGIN, y: 90, w, h: Math.round((w * (stored.h || 1)) / (stored.w || 1)) }, { src: stored.src, alt: file.name.replace(/\.[^.]+$/, ""), fit: "contain" });
        addObject(o);
      } catch { app.toast("画像を読み込めませんでした"); }
    }
    function addObject(o) {
      const made = E.normalizeObject(o);
      if (!made) return;
      if ((work.objects || []).length >= MAX_OBJECTS) { app.toast(`置けるのは ${MAX_OBJECTS} 個までです`); return; }
      work.objects = [...(work.objects || []), made];
      selected = `o:${made.id}`;
      drawPreview();
      drawForm();
    }
    function drawForm() {
      const r = resolved();
      const obj = selectedObject();
      const orient = sel(r.orientation, [["portrait", "縦（A4）"], ["landscape", "横（A4）"]], (v) => { if (v === "landscape") work.orientation = "landscape"; else delete work.orientation; drawPreview(); });
      const rows = HF_KEYS.map((key) => {
        const p = r.ph[key];
        return h("div", { class: "pm-ph-row" },
          h("label", { class: "sh-choice" }, h("input", { type: "checkbox", checked: p.on || null, "data-hf": key, onchange: (e) => { work[key] = { ...(work[key] || {}), on: e.target.checked }; drawPreview(); } }), h("span", {}, HF_NAMES[key])),
          key === "pageNo" ? h("span", { class: "hint" }, "ページの番号") : h("input", { type: "text", class: "pm-text", "aria-label": `${HF_NAMES[key]}の文字`, placeholder: key === "date" ? "（空欄＝印刷する日）" : key === "header" ? "（空欄＝資料のタイトル）" : "", value: p.text, maxlength: 200, oninput: (e) => { work[key] = { ...(work[key] || {}), text: e.target.value }; drawPreview(); } }),
          sel(p.align, Object.entries(ALIGN_NAMES), (v) => { work[key] = { ...(work[key] || {}), align: v }; drawPreview(); }));
      });
      const objectList = h("div", { class: "pm-objs" }, r.objects.map((o) => h("div", { class: ["pm-obj-row", selected === `o:${o.id}` ? "on" : ""], "data-obj": o.id },
        h("button", { type: "button", class: "btn btn-sm", onclick: () => { selected = `o:${o.id}`; drawPreview(); drawForm(); } }, o.kind === "image" ? `画像：${o.alt || ""}` : `文字：${E.richToText(o.text || "").slice(0, 16)}`),
        h("button", { type: "button", class: "btn btn-ghost btn-sm", "aria-label": "削除", onclick: () => { work.objects = (work.objects || []).filter((x) => x.id !== o.id); if (selected === `o:${o.id}`) selected = null; drawPreview(); drawForm(); } }, "✕"))));
      const detail = obj?.kind === "text"
        ? h("div", { class: "pm-detail" },
          h("label", {}, "文字", h("textarea", { rows: 2, class: "pm-obj-text", maxlength: 400, oninput: (e) => changeObject({ text: E.textToRich(e.target.value) }) }, E.richToText(obj.text || ""))),
          h("div", { class: "pm-line" }, h("label", {}, "大きさ", numberField(obj.fs || 14, { min: 8, max: 72 }, (v) => changeObject({ fs: v }))),
            h("label", {}, "色", sel(obj.color || "#000000", TEXT_COLORS, (v) => changeObject({ color: v }))),
            h("label", {}, "位置", sel(obj.align || "left", Object.entries(ALIGN_NAMES), (v) => changeObject({ align: v })))))
        : obj ? h("p", { class: "hint" }, "図はドラッグで動かし、右下の角で大きさを変えます。") : null;
      side.replaceChildren(...[
        h("label", {}, "用紙の向き", orient),
        kind === "handout" ? h("label", {}, "1ページのスライド数（初めの配置）", sel(r.perPage || "h6", Object.entries(HANDOUT_LAYOUTS).map(([key, n]) => [key, `${n}スライド`]), (v) => { work.perPage = v; drawPreview(); })) : null,
        h("label", {}, "ヘッダー・フッターの文字の大きさ", numberField(r.fs, { min: FS.min, max: FS.max }, (v) => { work.fs = v; drawPreview(); })),
        kind === "notes" ? h("label", {}, "ノートの文字の大きさ", numberField(r.notes.fs, { min: 8, max: 32 }, (v) => { work.notes = { ...(work.notes || {}), fs: v }; drawPreview(); })) : null,
        h("fieldset", { class: "pr-hf" }, h("legend", {}, "ヘッダー・日付・フッター・ページ番号（印刷では印刷画面で1回ごとに変えられます）"), ...rows),
        h("fieldset", { class: "pr-hf" }, h("legend", {}, "全ページに出す図・文字"),
          h("div", { class: "pm-add" }, h("button", { type: "button", class: "btn btn-sm pm-add-text", onclick: () => addObject(ops.makeObject("text", { x: MARGIN, y: 80, w: 300, h: 34 }, { text: "<p>文字</p>", fs: 14 })) }, "＋ 文字"), h("button", { type: "button", class: "btn btn-sm pm-add-image", onclick: addImage }, "＋ 図…")),
          objectList, detail),
        h("p", { class: "hint" }, "ページの上のものはドラッグで動かせます（選ぶと右下に大きさを変える角が出ます）。")].filter(Boolean));
    }

    const close = (save) => {
      if (save) app.setPrintMaster(kind, work);
      dialog.close();
      dialog.remove();
    };
    dialog.append(
      h("div", { class: "dialog-head" }, h("h3", {}, KINDS[kind]), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => close(false) }, "✕")),
      h("div", { class: "dialog-body pm-body" }, side, preview),
      h("div", { class: "dialog-foot" },
        h("button", { class: "btn btn-ghost pm-reset", type: "button", onclick: () => { work = {}; selected = null; drawPreview(); drawForm(); } }, "初期の状態に戻す"),
        h("button", { class: "btn", type: "button", onclick: () => close(false) }, "キャンセル"),
        h("button", { class: "btn btn-primary pm-ok", type: "button", onclick: () => close(true) }, "OK")));
    dialog.addEventListener("cancel", (event) => { event.preventDefault(); close(false); });
    document.body.append(dialog);
    dialog.showModal();
    drawForm();
    drawPreview();
  }
  return { open };
}
