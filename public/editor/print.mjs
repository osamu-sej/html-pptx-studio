// 印刷 (PowerPoint's File → Print): which slides (all, this one, a range like "1,3,5-8", with or without hidden
// slides), the layout (full-page slides, notes pages, the outline, handouts with 1/2/3/4/6/9 slides a page, in rows
// or columns), colour or grayscale, frames, and the header and footer of notes and handouts — with a preview of the
// pages. Printing (or saving as PDF) goes through the browser's print dialog. The page itself — its orientation, the places
// of the header, footer, date, page number, slide and notes, and the words and pictures on every page — is the deck's
// 配布資料マスター・ノート マスター (printmaster.mjs).

import { pageOf, printDefaults, resolveMaster } from "./printmaster.mjs";
import { createNotesEditor } from "./notesedit.mjs";

export const LAYOUTS = {
  full: "フル ページ サイズのスライド",
  notes: "ノート",
  outline: "アウトライン",
  h1: "配布資料（1スライド）",
  h2: "配布資料（2スライド）",
  h3: "配布資料（3スライド・メモ欄付き）",
  h4: "配布資料（4スライド）",
  h6: "配布資料（6スライド）",
  h9: "配布資料（9スライド）",
};
// A4 portrait at 96 dpi, in CSS pixels.
export const PAGE = { w: 794, h: 1123, margin: 48 };
const GRID = { h1: [1, 1], h2: [1, 2], h3: [1, 3], h4: [2, 2], h6: [2, 3], h9: [3, 3] };

/** "1,3,5-8" → [0, 2, 4, 5, 6, 7] (places in the deck); null when the text cannot be read. */
export function parseRange(text, total) {
  const out = [];
  for (const part of String(text || "").replace(/[，、]/g, ",").replace(/[－―ー〜~]/g, "-").split(",").map((p) => p.trim()).filter(Boolean)) {
    const m = part.match(/^(\d+)(?:\s*-\s*(\d+))?$/);
    if (!m) return null;
    const a = Number(m[1]);
    const b = m[2] ? Number(m[2]) : a;
    if (a < 1 || b < a || b > total) return null;
    for (let i = a; i <= b; i += 1) if (!out.includes(i - 1)) out.push(i - 1);
  }
  return out.length ? out : null;
}

/** Which slides are printed: [index] in deck order. */
export function printedSlides(deck, { range = "all", current = 0, custom = "", hidden = false } = {}) {
  const total = deck?.slides?.length || 0;
  let list = range === "current" ? [current] : range === "custom" ? parseRange(custom, total) || [] : deck.slides.map((_, i) => i);
  if (!hidden && range !== "current") list = list.filter((i) => !deck.slides[i]?.hidden);
  return list;
}

/** Slides on pages for a layout: [[index, …], …] (full and notes pages hold one; outline pages are made later). */
export function paginate(list, layout) {
  const per = layout in GRID ? GRID[layout][0] * GRID[layout][1] : 1;
  const pages = [];
  for (let i = 0; i < list.length; i += per) pages.push(list.slice(i, i + per));
  return pages;
}

/** Where each slide goes on a handout page (CSS px from the page's corner) — in rows, or down the columns. */
export function handoutBoxes(layout, count, { order = "rows", page = PAGE } = {}) {
  // On a landscape page the grid lies down too: 6 slides are 3 across and 2 down.
  const [cols, rows] = page.w > page.h && layout !== "h1" && layout !== "h3" ? [...GRID[layout]].reverse() : GRID[layout];
  const top = PAGE.margin + 36;
  const bottom = page.h - PAGE.margin - 30;
  const left = PAGE.margin;
  const width = page.w - PAGE.margin * 2;
  const gap = 24;
  if (layout === "h3") {
    const w = Math.round(width * 0.46);
    const hh = Math.round((w * 9) / 16);
    const step = (bottom - top - hh) / 2;
    return Array.from({ length: count }, (_, k) => ({ x: left, y: Math.round(top + step * k), w, h: hh, lines: { x: left + w + gap, w: width - w - gap } }));
  }
  const cellW = (width - gap * (cols - 1)) / cols;
  const cellH = (bottom - top - gap * (rows - 1)) / rows;
  const w = Math.floor(Math.min(cellW, (cellH * 16) / 9));
  const hh = Math.floor((w * 9) / 16);
  return Array.from({ length: count }, (_, k) => {
    const c = order === "columns" ? Math.floor(k / rows) : k % cols;
    const r = order === "columns" ? k % rows : Math.floor(k / cols);
    return { x: Math.round(left + c * (cellW + gap) + (cellW - w) / 2), y: Math.round(top + r * (cellH + gap) + (cellH - hh) / 2), w, h: hh };
  });
}

/** A slide's comments as printed lines (PowerPoint's 「コメントを印刷する」): each comment, then its replies. */
export function commentLines(slide) {
  const day = (at) => { const d = new Date(at || ""); return Number.isNaN(d.getTime()) ? "" : d.toLocaleString("ja-JP", { year: "numeric", month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" }); };
  const line = (x, level, done) => ({ level, by: String(x?.by || "").trim() || "（名前なし）", at: day(x?.at), text: String(x?.text || ""), done: Boolean(done) });
  const out = [];
  for (const c of Array.isArray(slide?.comments) ? slide.comments : []) {
    if (!c || !String(c.text || "").trim()) continue;
    out.push(line(c, 0, c.done));
    for (const r of Array.isArray(c.replies) ? c.replies : []) if (r && String(r.text || "").trim()) out.push(line(r, 1, false));
  }
  return out;
}

/** Comment lines in page-sized runs: about `rows` rows a page, a row holding `perRow` characters (each comment takes
 *  its name line, its text rows and a gap; a comment longer than a page gets a page of its own). */
export function commentChunks(lines, { perRow = 52, rows = 44 } = {}) {
  const out = [];
  let run = [];
  let used = 0;
  for (const l of lines) {
    const need = 2 + String(l.text).split(/\n/).reduce((n, t) => n + Math.max(1, Math.ceil(t.length / perRow)), 0);
    if (run.length && used + need > rows) { out.push(run); run = []; used = 0; }
    run.push(l);
    used += need;
  }
  if (run.length) out.push(run);
  return out;
}

export function createPrinter(app) {
  const { h, E } = app;
  const strip = (t) => E.strip(String(t ?? ""));
  const SETTINGS_KEY = "hpptx-print";
  const load = () => { try { return JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}"); } catch { return {}; } };
  const save = (o) => { try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(o)); } catch { /* private window */ } };

  /** A slide drawn at `w` CSS px wide (the engine draws it at 1920 × 1080 and it is scaled down). */
  function thumb(i, w, frame) {
    const el = app.renderSlide(i);
    const k = w / E.W;
    el.style.transform = `scale(${k})`;
    el.style.transformOrigin = "0 0";
    return h("div", { class: ["pr-slide", frame ? "framed" : ""], style: { width: `${w}px`, height: `${Math.round(w * 9 / 16)}px` } }, el);
  }
  /** The master a layout is printed on: notes pages use the ノート マスター, handouts and the outline the 配布資料マスター. */
  const masterFor = (layout) => { const kind = layout === "notes" ? "notes" : "handout"; return resolveMaster(app.printMasters?.()?.[kind], kind); };
  /** A page of the master's size (the master's pictures and words lie under everything else on it). */
  function sheet(master, classes) {
    const page = h("div", { class: classes, style: { width: `${master.page.w}px`, height: `${master.page.h}px` } });
    if (master.objects.length) {
      const layer = E.objectLayer({ elements: master.objects }, app.renderCtx?.() || {});
      if (layer) page.append(h("div", { class: "hs-slide pr-objects", "data-theme": E.PALETTE.sheet?.themeId || E.DEFAULT_THEME, style: { width: `${master.page.w}px`, height: `${master.page.h}px` } }, layer));
    }
    return page;
  }
  /** ヘッダー・日付・フッター・ページ番号, each where the master puts it. */
  function chrome(page, n, opts, master) {
    const deck = app.deck();
    const today = new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long", day: "numeric" });
    const place = (key, text) => { const p = master.ph[key]; page.append(h("div", { class: `pr-ph pr-ph-${key}`, style: { left: `${p.x}px`, top: `${p.y}px`, width: `${p.w}px`, textAlign: p.align, fontSize: `${master.fs}px` } }, text)); };
    if (opts.header) place("header", opts.headerText || deck.title);
    if (opts.date) place("date", master.ph.date.text || today);
    if (opts.footer) place("footer", opts.footerText || "");
    if (opts.pageNo) place("pageNo", String(n));
  }
  /** The outline of a slide: its title, then its words (key message, points, text of its parts). */
  function outlineOf(slide) {
    const lines = [];
    for (const key of ["subtitle", "subhead", "takeaway", "message", "text"]) if (strip(slide[key])) lines.push(strip(slide[key]));
    for (const key of ["points", "steps", "leftItems", "rightItems"]) for (const item of Array.isArray(slide[key]) ? slide[key] : []) if (strip(typeof item === "string" ? item : item?.title || item?.label)) lines.push(strip(typeof item === "string" ? item : item.title || item.label));
    for (const item of Array.isArray(slide.items) ? slide.items : []) { const t = strip(typeof item === "string" ? item : [item?.title || item?.label, item?.desc].filter(Boolean).join("：")); if (t) lines.push(t); }
    for (const o of slide.elements || []) { if (o.hidden) continue; const t = strip(E.objectText ? E.objectText(o) : o.text); if (t && t !== strip(slide.title)) lines.push(t.slice(0, 200)); }
    return [...new Set(lines)].slice(0, 30);
  }

  /** The pages of one slide's comments (after the slide, or at the end): who wrote each, when, and the replies. */
  function commentsPages(i, full, first, opts) {
    const lines = commentLines(app.deck().slides[i]);
    const chunks = commentChunks(lines, full ? { perRow: 50, rows: 15 } : { perRow: 52, rows: 44 });
    return chunks.map((chunk, k) => {
      const page = h("div", { class: [full ? "pr-full" : "pr-page", "pr-comments"] },
        h("h2", {}, `スライド ${i + 1} のコメント${chunks.length > 1 ? `（${k + 1}/${chunks.length}）` : ""}`),
        h("ul", {}, chunk.map((l) => h("li", { class: l.level ? "pr-cm-reply" : "" },
          h("div", { class: "pr-cm-meta" }, h("b", {}, l.by), l.at ? h("span", {}, l.at) : null, l.done ? h("span", {}, "解決済み") : null),
          h("div", { class: "pr-cm-text" }, l.text.split(/\n/).map((t) => h("p", {}, t || "\u00a0")))))));
      if (!full) chrome(page, first + k, opts, masterFor(opts.layout));
      return page;
    });
  }

  /** The pages for the settings (each an element of PAGE size, or a slide-size page for full slides). */
  function pages(opts) {
    const deck = app.deck();
    const list = printedSlides(deck, { range: opts.range, current: app.index(), custom: opts.custom, hidden: opts.hidden });
    const out = [];
    // Comments: after their slide (full pages, notes pages), or at the end (outline, handouts).
    const commentsAfter = (i) => { if (opts.comments) out.push(...commentsPages(i, opts.layout === "full", out.length + 1, opts)); };
    if (opts.layout === "full") {
      for (const i of list) {
        const el = app.renderSlide(i);
        out.push(h("div", { class: ["pr-full", opts.frame ? "framed" : ""] }, el));
        commentsAfter(i);
      }
      return out;
    }
    const master = masterFor(opts.layout);
    if (opts.layout === "outline") {
      let page = null;
      let used = 0;
      const start = () => { page = sheet(master, "pr-page pr-outline"); out.push(page); chrome(page, out.length, opts, master); used = 0; };
      start();
      for (const i of list) {
        const slide = deck.slides[i];
        const lines = outlineOf(slide);
        const need = 1 + lines.length;
        if (used && used + need > 34) start();
        page.append(h("div", { class: "pr-ol" }, h("div", { class: "pr-ol-title" }, h("b", {}, `${i + 1}`), strip(slide.title) || "（タイトルなし）"), lines.length ? h("ul", {}, lines.map((t) => h("li", {}, t))) : null));
        used += need;
      }
      for (const i of list) commentsAfter(i);
      return out;
    }
    if (opts.layout === "notes") {
      for (const i of list) {
        const page = sheet(master, "pr-page pr-notes");
        chrome(page, out.length + 1, opts, master);
        const box = thumb(i, master.slide.w, opts.frame);
        box.classList.add("pr-notes-slide");
        Object.assign(box.style, { left: `${master.slide.x}px`, top: `${master.slide.y}px` });
        const n = master.notes;
        page.append(box, h("div", { class: "pr-notes-text", style: { left: `${n.x}px`, top: `${n.y}px`, width: `${n.w}px`, height: `${n.h}px`, right: "auto", bottom: "auto", fontSize: `${n.fs}px` } }, noteNodes(deck.slides[i])));
        out.push(page);
        commentsAfter(i);
      }
      return out;
    }
    for (const group of paginate(list, opts.layout)) {
      const page = sheet(master, `pr-page pr-handout ${opts.layout}`);
      chrome(page, out.length + 1, opts, master);
      handoutBoxes(opts.layout, group.length, { order: opts.order, page: master.page }).forEach((b, k) => {
        const box = thumb(group[k], b.w, opts.frame);
        Object.assign(box.style, { position: "absolute", left: `${b.x}px`, top: `${b.y}px` });
        page.append(box);
        if (b.lines) page.append(h("div", { class: "pr-lines", style: { left: `${b.lines.x}px`, top: `${b.y}px`, width: `${b.lines.w}px`, height: `${b.h}px` } }));
      });
      out.push(page);
    }
    for (const i of list) commentsAfter(i);
    return out;
  }

  /** Into the print root, then the browser's print dialog (A4 portrait, or slide-size pages). */
  async function print(opts) {
    await app.beforePrint?.();
    const root = document.getElementById("printRoot");
    const list = pages(opts);
    if (!list.length) { app.toast("印刷するスライドがありません（範囲を確かめてください）"); return; }
    root.replaceChildren(...list);
    root.className = ["print-root", opts.layout === "full" ? "pr-mode-full" : "pr-mode-page", opts.color === "gray" ? "pr-gray" : "", opts.color === "bw" ? "pr-bw" : ""].filter(Boolean).join(" ");
    const landscape = opts.layout !== "full" && masterFor(opts.layout).orientation === "landscape";
    const style = h("style", { id: "printPageStyle" }, opts.layout === "full" ? "@media print { @page { size: 1920px 1080px; margin: 0; } }" : `@media print { @page { size: A4 ${landscape ? "landscape" : "portrait"}; margin: 0; } }`);
    document.getElementById("printPageStyle")?.remove();
    document.head.append(style);
    await Promise.all([...root.querySelectorAll("img")].map((img) => img.decode?.().catch(() => {})));
    await E.mountLottie?.(root, { play: false, frame: 0.5 });
    await document.fonts?.ready;
    const cleanup = () => { E.stopLottie?.(root); root.replaceChildren(); root.className = "print-root"; style.remove(); window.removeEventListener("afterprint", cleanup); };
    window.addEventListener("afterprint", cleanup);
    app.toast(opts.layout === "full" ? "印刷画面で「PDFとして保存」を選び、余白を「なし」にしてください" : `印刷画面で用紙をA4・${landscape ? "横" : "縦"}、余白を「なし」にしてください（PDFとして保存もできます）`);
    setTimeout(() => window.print(), 80);
  }

  function openDialog() {
    const deck = app.deck();
    if (!deck) return;
    const saved = load();
    const opts = { range: "all", custom: "", hidden: false, layout: "full", order: "rows", frame: false, color: "color", header: true, date: true, pageNo: true, footer: false, headerText: "", footerText: "", comments: false, ...saved, ...(saved.range === "custom" ? {} : { custom: "" }) };
    // 配布資料マスター・ノート マスター: the header, date, footer and page number start as the master has them (and a handout master's
    // 1ページあたりのスライド数 is the layout when none was used before); they can be changed for one printing here.
    if (!saved.layout && masterFor("h6").perPage) opts.layout = masterFor("h6").perPage;
    Object.assign(opts, printDefaults(masterFor(opts.layout)));
    let masterKind = opts.layout === "notes" ? "notes" : "handout";
    const preview = h("div", { class: "pr-preview", "aria-label": "印刷プレビュー" });
    const count = h("span", { class: "hint pr-count" });
    const sel = (name, entries, value) => h("select", { name }, entries.map(([v, label]) => h("option", { value: v, selected: v === value || null }, label)));
    const form = h("form", { class: "pr-form", onsubmit: (e) => e.preventDefault(), oninput: () => refresh(), onchange: (event) => { if (event.target?.name === "layout") adoptMaster(event.target.value); refresh(); } },
      h("label", {}, "印刷範囲", sel("range", [["all", "すべてのスライドを印刷"], ["current", "現在のスライドを印刷"], ["custom", "ユーザー設定の範囲"]], opts.range)),
      h("label", {}, "スライド指定", h("input", { name: "custom", type: "text", value: opts.custom, placeholder: "例：1,3,5-8" })),
      h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name: "hidden", checked: opts.hidden || null }), h("span", {}, "非表示スライドを印刷する")),
      h("label", {}, "印刷レイアウト", sel("layout", Object.entries(LAYOUTS), opts.layout)),
      h("label", {}, "順序（配布資料）", sel("order", [["rows", "横（左から右へ）"], ["columns", "縦（上から下へ）"]], opts.order)),
      h("label", {}, "色", sel("color", [["color", "カラー"], ["gray", "グレースケール"], ["bw", "単純白黒"]], opts.color)),
      h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name: "frame", checked: opts.frame || null }), h("span", {}, "スライドに枠を付ける")),
      h("label", { class: "sh-choice", title: "コメントのあるスライドの後（配布資料・アウトラインでは最後）に、コメントと返信のページを入れます" }, h("input", { type: "checkbox", name: "comments", checked: opts.comments || null }), h("span", {}, "コメントを印刷する")),
      h("fieldset", { class: "pr-hf" }, h("legend", {}, "ヘッダーとフッター（ノート・配布資料・アウトライン）"),
        h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name: "header", checked: opts.header || null }), h("span", {}, "ヘッダー"), h("input", { type: "text", name: "headerText", value: opts.headerText, placeholder: deck.title })),
        h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name: "date", checked: opts.date || null }), h("span", {}, "日付")),
        h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name: "pageNo", checked: opts.pageNo || null }), h("span", {}, "ページ番号")),
        h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name: "footer", checked: opts.footer || null }), h("span", {}, "フッター"), h("input", { type: "text", name: "footerText", value: opts.footerText, placeholder: "例：社内限り" }))));
    /** Choosing another kind of page (notes ⇄ handouts) takes the header, date, footer and page number of that master. */
    function adoptMaster(layout) {
      const kind = layout === "notes" ? "notes" : "handout";
      if (kind === masterKind || layout === "full") return;
      masterKind = kind;
      const d = printDefaults(masterFor(layout));
      const f = (name) => form.elements.namedItem(name);
      for (const key of ["header", "date", "pageNo", "footer"]) f(key).checked = d[key];
      f("headerText").value = d.headerText;
      f("footerText").value = d.footerText;
    }
    // Read from the fields themselves: a field turned off for this layout (the header and footer of full pages) keeps
    // its setting for the next layout instead of reading as off.
    const read = () => {
      const f = (name) => form.elements.namedItem(name);
      return { range: f("range").value, custom: String(f("custom").value || ""), hidden: f("hidden").checked, layout: f("layout").value, order: f("order").value, color: f("color").value, frame: f("frame").checked, comments: f("comments").checked, header: f("header").checked, date: f("date").checked, pageNo: f("pageNo").checked, footer: f("footer").checked, headerText: String(f("headerText").value || ""), footerText: String(f("footerText").value || "") };
    };
    let timer = null;
    function refresh() {
      clearTimeout(timer);
      timer = setTimeout(() => {
        const o = read();
        form.querySelector('[name="custom"]').disabled = o.range !== "custom";
        form.querySelector('[name="order"]').disabled = !["h4", "h6", "h9"].includes(o.layout);
        form.querySelector(".pr-hf").disabled = o.layout === "full";
        if (o.range === "custom" && !parseRange(o.custom, deck.slides.length)) { preview.replaceChildren(h("p", { class: "hint" }, `スライド番号を「1,3,5-8」の形で（1〜${deck.slides.length}）`)); count.textContent = ""; return; }
        const list = pages(o);
        count.textContent = `${printedSlides(deck, { ...o, current: app.index() }).length}枚のスライド・${list.length}ページ`;
        const pg = o.layout === "full" ? { w: E.W, h: E.H } : masterFor(o.layout).page;
        const k = 300 / pg.w;
        preview.className = ["pr-preview", o.color === "gray" ? "pr-gray" : "", o.color === "bw" ? "pr-bw" : ""].filter(Boolean).join(" ");
        preview.replaceChildren(...list.slice(0, 12).map((p) => {
          const w = pg.w;
          const hh = pg.h;
          p.style.transform = `scale(${k})`;
          p.style.transformOrigin = "0 0";
          return h("div", { class: "pr-sheet", style: { width: `${Math.round(w * k)}px`, height: `${Math.round(hh * k)}px` } }, p);
        }), list.length > 12 ? h("p", { class: "hint" }, `ほか ${list.length - 12} ページ`) : "");
      }, 120);
    }
    const dialog = h("dialog", { class: "print-dialog", "aria-label": "印刷" },
      h("div", { class: "dialog-head" }, h("h3", {}, "印刷"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body pr-body" }, form, h("div", { class: "pr-side" }, count, preview)),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary pr-go", onclick: () => {
          const o = read();
          if (o.range === "custom" && !parseRange(o.custom, deck.slides.length)) { app.toast("スライド指定を確かめてください（例：1,3,5-8）"); return; }
          save(o);
          dialog.close();
          print(o);
        } }, "印刷（PDFに保存）")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => { clearTimeout(timer); dialog.remove(); });
    dialog.showModal();
    refresh();
  }

  /** 表示 → ノート (the notes page view): the slide above its speaker notes, page by page; the notes can be typed in. */
  /** A slide's notes for a page: the formatted words (bold, lists, links) as they were typed, or the plain lines. */
  function noteNodes(slide) {
    const note = E.noteOf(slide);
    if (note.formatted) {
      const box = h("div", { class: "pr-notes-rich" });
      box.append(E.richFragment(note.rich));
      return [box];
    }
    return String(slide.notes || "").split(/\n/).map((line) => h("p", {}, line || "\u00a0"));
  }
  function notesView() {
    const deck = app.deck();
    if (!deck) return;
    let i = app.index();
    const pageBox = h("div", { class: "nv-page" });
    const count = h("span", { class: "nv-count" });
    const editorBox = createNotesEditor(app, { label: "ノート" });
    editorBox.el.classList.add("nv-notes");
    const show = () => {
      app.select(i);
      const slide = app.deck().slides[i];
      const box = thumb(i, 640, true);
      pageBox.replaceChildren(box);
      editorBox.show(slide, { force: true });
      count.textContent = `${i + 1} / ${app.deck().slides.length}`;
    };
    const go = (d) => { const n = Math.max(0, Math.min(app.deck().slides.length - 1, i + d)); if (n !== i) { i = n; show(); } };
    const dialog = h("dialog", { class: "notes-view", "aria-label": "ノート表示" },
      h("div", { class: "dialog-head" }, h("h3", {}, "ノート表示"), h("div", { class: "nv-nav" },
        h("button", { type: "button", class: "btn btn-sm", "aria-label": "前のスライド", onclick: () => go(-1) }, "◀"), count,
        h("button", { type: "button", class: "btn btn-sm", "aria-label": "次のスライド", onclick: () => go(1) }, "▶")),
      h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body nv-body" }, h("div", { class: "nv-sheet" }, pageBox, editorBox.el)));
    dialog.addEventListener("keydown", (e) => { if (editorBox.el.contains(e.target)) return; if (e.key === "ArrowRight" || e.key === "PageDown") go(1); if (e.key === "ArrowLeft" || e.key === "PageUp") go(-1); });
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
    show();
    editorBox.focus();
  }

  return { openDialog, print, pages, notesView };
}
