// The editor's controls, PowerPoint style: the ribbon over the slide (ホーム・挿入・図形の書式・図の形式・表示),
// the 書式 pane on the right (fill, line, text, size and position, picture, link) with the selection pane, and the
// galleries they open (shapes, the theme's colours, icons, photos).

import * as ops from "./ops.mjs";
import { GRID_STEPS } from "./canvas.mjs";
import { ico } from "./icons.mjs";
import { createAnimations } from "./anim.mjs";
import { createInteractions } from "./interact.mjs";
import { createTableUi } from "./tables.mjs";
import { createCrop, cropFill, cropFit, cropToAspect } from "./crop.mjs";
import { createFreeform } from "./freeform.mjs";
import { createMedia } from "./media.mjs";
import { MERGE_HINTS, MERGE_MODES, canMerge, clippingLib, mergeObjects } from "./merge.mjs";
import { createSmartArt } from "./smartart.mjs";
import { createInk } from "./ink.mjs";
import { createExtras } from "./extras.mjs";
import { ART_EFFECTS, RECOLORS, createPictureTools } from "./picture.mjs";
import { createModels } from "./model3d.mjs";
import { createPolls } from "./poll.mjs";
import { createQuick } from "./quick.mjs";
import { pickColor } from "./eyedropper.mjs";
import { createReadAloud } from "./readaloud.mjs";

const SIZES_PT = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 66, 72, 80, 88, 96];
const LINE_WIDTHS = [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6, 8, 12];
// 文字の間隔 (PowerPoint's: tight to very loose), in em.
const CHAR_SPACING = [[-0.1, "より狭く"], [-0.05, "狭く"], [0, "標準"], [0.05, "広く"], [0.12, "より広く"]];
// 挿入 → 記号と特殊文字: what Japanese business slides use most (no < > &, which are markup).
const SYMBOLS = ["①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "※", "・", "→", "←", "↑", "↓", "⇒", "⇔", "○", "●", "◎", "△", "▲", "▼", "□", "■", "◇", "◆", "★", "☆", "✓", "✕", "±", "×", "÷", "≒", "≠", "≦", "≧", "∞", "℃", "％", "‰", "㎡", "㎏", "㎞", "￥", "€", "＄", "〒", "©", "®", "™", "…", "〜", "「", "」", "『", "』", "【", "】", "〔", "〕", "♪", "☎", "✉"];
const LINE_HEIGHTS = [[1, "1.0"], [1.15, "1.15"], [1.35, "1.35"], [1.5, "1.5"], [1.75, "1.75"], [2, "2.0"], [2.5, "2.5"], [3, "3.0"]];
// 図形のスタイル (PowerPoint's theme styles): for テキスト 1 and each accent, 枠線のみ, 塗りつぶし and a light fill, in
// the current theme's colours (they follow the theme when it changes).
function quickStyles(E) {
  const c = E.themeColors(E.PALETTE.sheet?.themeId);
  const slots = [["テキスト 1", c.tx1], ...[1, 2, 3, 4, 5, 6].map((i) => [`アクセント ${i}`, c[`accent${i}`]])];
  return [
    ...slots.map(([name, color]) => [`枠線のみ - ${name}`, { fill: c.bg1, stroke: color, strokeW: 3, color: c.tx1 }]),
    ...slots.map(([name, color]) => [`塗りつぶし - ${name}`, { fill: color, stroke: "none", color: c.bg1 }]),
    ...slots.map(([name, color]) => [`パステル - ${name}`, { fill: E.lumColor(color, 20000, 80000), stroke: "none", color: c.tx1 }]),
  ];
}
const ACTIONS = [["none", "なし"], ["next", "次のスライド"], ["prev", "前のスライド"], ["first", "最初のスライド"], ["last", "最後のスライド"], ["slide", "スライドを指定…"], ["url", "URL（Webページ・メール）…"], ["end", "スライドショーの終了"]];
const MASKS = ["rect", "roundRect", "ellipse", "triangle", "diamond", "pentagon", "hexagon", "octagon", "star5", "star8", "heart", "cloud", "teardrop", "wedgeRoundRectCallout", "flowTerminator", "snip2DiagRect", "round2DiagRect", "chevron", "homePlate"];

export function createEditorUi(editor, app) {
  const { E, h } = app;
  const ribbon = document.getElementById("ribbon");
  const pane = document.getElementById("formatPane");
  let tab = "home";
  // Ribbon display options, as PowerPoint's: "full" (always shown), "tabs" (only the tabs; a tab opens the ribbon
  // over the slide until a command or a click elsewhere), "auto" (hidden behind a thin bar at the top).
  const RIBBON_MODES = { full: "常にリボンを表示する", tabs: "タブのみを表示する", auto: "自動的に非表示にする（全画面表示モード）" };
  let ribbonMode = (() => { try { const m = localStorage.getItem("hpptx-ribbon-mode"); return RIBBON_MODES[m] ? m : "full"; } catch { return "full"; } })();
  let floating = false;
  let grayView = false; // 表示 → グレースケール (only how the slide is shown here)
  let hideInk = false; // 校閲 → インクの非表示 (only while editing)
  let built = "";
  let updaters = [];
  let pop = null;
  let quick = null; // 検索 (⌥Q) and the クイック アクセス ツール バー (quick.mjs)

  // ---------------------------------------------------------------- small building blocks

  const selected = () => editor.selectedObjects();
  const kinds = () => new Set(selected().map((o) => o.kind));
  const hasText = () => selected().some((o) => o.kind === "shape" || o.kind === "text");
  const hasShape = () => selected().some((o) => ["shape", "text", "line", "icon"].includes(o.kind));
  const hasImage = () => kinds().has("image");
  const one = () => (selected().length === 1 ? selected()[0] : null);
  const any = () => selected().length > 0;
  const onSlide = () => Boolean(app.slide()) && app.state().view === "single";

  // A big button's label breaks only where "|" says (新しい|スライド), as PowerPoint's two-line labels.
  const plain = (label) => (label || "").replaceAll("|", "");
  const caption = (label) => (label ? h("span", {}, ...label.split("|").flatMap((part, i) => (i ? [h("wbr"), part] : [part]))) : null);

  /** A ribbon button: icon + label, pressed/enabled follow the selection. */
  function btn(icon, label, title, run, { pressed = null, enabled = null, big = false, keep = false } = {}) {
    const el = h("button", { type: "button", class: ["rb-btn", big ? "big" : ""], title: title || plain(label), "aria-label": title || plain(label), "data-keeps-text": keep ? "" : null,
      onmousedown: (event) => { if (keep) event.preventDefault(); },
      // A command from the ribbon floating over the slide closes it afterwards (PowerPoint's タブのみを表示).
      onclick: (event) => { event.preventDefault(); const fromFloat = floating && ribbon.contains(el); run(event); if (fromFloat && floating && !pop) closeFloating(); } }, ico(icon, big ? 26 : 18), caption(label));
    updaters.push(() => {
      if (enabled) el.disabled = !enabled();
      if (pressed) el.setAttribute("aria-pressed", String(Boolean(pressed())));
    });
    return el;
  }
  /** A button that opens a menu or gallery under it. */
  function drop(icon, label, title, build, { enabled = null, big = false, keep = false, swatch = null } = {}) {
    const el = h("button", { type: "button", class: ["rb-btn", "rb-drop", big ? "big" : ""], title: title || plain(label), "aria-haspopup": "true", "data-keeps-text": keep ? "" : null,
      onmousedown: (event) => { if (keep) event.preventDefault(); },
      onclick: (event) => { event.preventDefault(); openPop(el, build); } },
    ico(icon, big ? 26 : 18), caption(label), swatch ? h("i", { class: "rb-swatch" }) : null, h("b", { class: "rb-caret" }, "▾"));
    updaters.push(() => {
      if (enabled) el.disabled = !enabled();
      if (swatch) el.querySelector(".rb-swatch").style.background = swatch() || "transparent";
    });
    return el;
  }
  /** The ▾ half of a split button (箇条書き ▾): the menu beside a button that acts at once. */
  function caret(title, build, { enabled = null } = {}) {
    const el = h("button", { type: "button", class: "rb-btn rb-drop rb-caret-only", title, "aria-label": title, "aria-haspopup": "true", "data-keeps-text": "",
      onmousedown: (event) => event.preventDefault(), onclick: (event) => { event.preventDefault(); openPop(el, build); } }, h("b", { class: "rb-caret" }, "▾"));
    updaters.push(() => { if (enabled) el.disabled = !enabled(); });
    return el;
  }
  const group = (label, ...items) => h("div", { class: "rb-group" }, h("div", { class: "rb-items" }, ...items.flat().filter(Boolean)), h("div", { class: "rb-label" }, label));
  const col = (...items) => h("div", { class: "rb-col" }, ...items.flat().filter(Boolean));
  const row = (...items) => h("div", { class: "rb-row" }, ...items.flat().filter(Boolean));

  // ---------------------------------------------------------------- popovers (menus and galleries)

  function openPop(anchor, build) {
    // Measure first: the anchor may sit in the popover being closed (a button of a folded ribbon group).
    const r = anchor.getBoundingClientRect();
    const nested = pop?.el.contains(anchor);
    const fromRibbon = floating && (ribbon.contains(anchor) || (nested && pop?.fromRibbon));
    closePop(true);
    // A builder may return the content, or a function that builds it with `close`.
    let content = build(closePop);
    while (typeof content === "function") content = content(closePop);
    if (!content) return;
    const el = h("div", { class: "rb-pop", role: "dialog", "data-keeps-text": "" }, content);
    document.body.append(el);
    if (nested) anchor = el;
    const pr = el.getBoundingClientRect();
    el.style.left = `${Math.max(8, Math.min(r.left, innerWidth - pr.width - 8))}px`;
    el.style.top = `${Math.min(r.bottom + 4, innerHeight - pr.height - 8)}px`;
    pop = { el, anchor, fromRibbon };
    setTimeout(() => document.addEventListener("pointerdown", outside, true), 0);
    document.addEventListener("keydown", popKey, true);
  }
  // A click elsewhere closes the menu (and the floating ribbon it came from, unless the click is on the ribbon).
  function outside(event) { if (pop && !pop.el.contains(event.target) && !pop.anchor.contains(event.target)) closePop(ribbon.contains(event.target)); }
  function popKey(event) { if (event.key === "Escape" && pop) { event.stopPropagation(); closePop(true); } }
  /** Closes the open menu or gallery; one opened from the floating ribbon closes the ribbon too (a command was chosen). */
  function closePop(keepRibbon = false) {
    const fromRibbon = pop?.fromRibbon;
    pop?.el.remove();
    pop = null;
    document.removeEventListener("pointerdown", outside, true);
    document.removeEventListener("keydown", popKey, true);
    if (fromRibbon && !keepRibbon && floating) { floating = false; renderRibbon(true); }
  }
  const menu = (items) => (close) => h("div", { class: "rb-menu" }, items.filter(Boolean).map((item) => (item === "-" ? h("div", { class: "rb-sep" })
    : item.head ? h("div", { class: "rb-menu-head" }, item.head)
      : h("button", { type: "button", class: item.on ? "on" : "", title: item.title || null, disabled: item.disabled || null, onmousedown: (e) => e.preventDefault(), onclick: () => { item.hover?.(false); close(); item.run(); },
        onmouseenter: item.hover ? () => item.hover(true) : null, onmouseleave: item.hover ? () => item.hover(false) : null }, item.icon ? ico(item.icon) : h("span", { class: "rb-noicon" }), h("span", {}, item.label), item.keys ? h("kbd", {}, item.keys) : null))));

  // トリミング ▾: 縦横比, 塗りつぶし, 枠に合わせる (the picture's own shape from the stage's <img>).
  const CROP_RATIOS = [["1:1", 1], ["2:3", 2 / 3], ["3:2", 3 / 2], ["3:4", 3 / 4], ["4:3", 4 / 3], ["3:5", 3 / 5], ["5:3", 5 / 3], ["4:5", 4 / 5], ["5:4", 5 / 4], ["9:16", 9 / 16], ["16:9", 16 / 9], ["16:10", 16 / 10]];
  function pictureAspect(o) {
    const img = document.querySelector(`#stageBody .hs-obj[data-el="${CSS.escape(o.id)}"] img`);
    if (img?.naturalWidth && img.naturalHeight) return img.naturalWidth / img.naturalHeight;
    const c = o.crop || { l: 0, t: 0, r: 0, b: 0 };
    return (o.w / (1 - c.l - c.r)) / (o.h / (1 - c.t - c.b));
  }
  function cropMenu() {
    const o = one();
    if (o?.kind !== "image") return null;
    const apply = (patch, label) => { editor.apply(patch, { ids: [o.id] }); app.toast(`${label}（⌘Zで戻せます）`); };
    return menu([
      { head: "縦横比" },
      ...CROP_RATIOS.map(([label, r]) => ({ label, run: () => apply(cropToAspect(o, r, pictureAspect(o)), `縦横比 ${label} にトリミングしました`) })),
      "-",
      { label: "塗りつぶし（枠いっぱいに）", icon: "zoomFit", run: () => apply(cropFill(o, pictureAspect(o)), "枠いっぱいに合わせました") },
      { label: "枠に合わせる（全体を表示）", icon: "zoomFit", run: () => apply(cropFit(o, pictureAspect(o)), "画像の全体が枠に入るようにしました") },
    ]);
  }

  // 箇条書き・段落番号の種類: PowerPoint's galleries (なし first); the marker stays on the list (data-style).
  const LIST_NAMES = { disc: "黒丸", circle: "白丸", square: "四角", diamond: "ひし形", arrow: "矢印", check: "チェック マーク", star: "星", dash: "ダッシュ", decimal: "1. 2. 3.", paren: "1) 2) 3)", circled: "① ② ③", "lower-alpha": "a. b. c.", "upper-alpha": "A. B. C.", "lower-roman": "i. ii. iii.", "upper-roman": "I. II. III.", kanji: "一、二、三、" };
  function listGallery(type) {
    const tag = type === "number" ? "ol" : "ul";
    return (close) => {
      const st = editor.textState();
      const current = st?.list === type ? st.listStyle : st?.list ? "" : "none";
      const marks = (style) => (tag === "ol" ? E.LIST_STYLES.ol[style] : [0, 1, 2].map(() => E.LIST_STYLES.ul[style]));
      const cell = (style, body, label) => h("button", { type: "button", class: ["rb-list-sw", current === (style || "none") ? "on" : ""], title: label, "aria-label": label, "data-list-style": style || "none",
        onmousedown: (event) => event.preventDefault(), onclick: () => { close(); editor.textFormat("listStyle", [style ? type : null, style]); } }, body);
      return h("div", { class: "rb-gallery rb-lists" }, h("div", { class: "rb-gallery-head" }, type === "number" ? "段落番号" : "箇条書き"),
        h("div", { class: "rb-list-grid" },
          cell(null, h("span", { class: "rb-list-none" }, "なし"), "なし"),
          Object.keys(E.LIST_STYLES[tag]).map((style) => cell(style, h("span", { class: "rb-list-prev" }, marks(style).map((m) => h("span", {}, h("b", {}, m), h("i")))), LIST_NAMES[style] || style))),
        h("button", { type: "button", class: "rb-menu-foot rb-list-more", disabled: !st?.list, title: st?.list ? "" : "先に箇条書き・段落番号にしてください", onmousedown: (event) => event.preventDefault(), onclick: () => { close(); listDialog(type); } }, "箇条書きと段落番号…"));
    };
  }
  /** 箇条書きと段落番号: the marks' size and colour (the text's, the accent or the muted colour) and a numbered list's first number. */
  function listDialog(type) {
    const now = editor.textState()?.listProps || { start: 1, mark: "", msize: "100" };
    const size = h("select", { class: "ld-size", "aria-label": "サイズ" }, ...[["75", "75%"], ["100", "100%（文字と同じ）"], ["125", "125%"], ["150", "150%"]].map(([v, l]) => h("option", { value: v, selected: now.msize === v }, l)));
    const color = h("select", { class: "ld-color", "aria-label": "色" }, ...[["", "文字と同じ"], ["accent", "アクセント"], ["muted", "薄い文字の色"]].map(([v, l]) => h("option", { value: v, selected: now.mark === v }, l)));
    const start = h("input", { type: "number", min: "1", max: "9999", value: String(now.start || 1), class: "ld-start sh-num", "aria-label": "開始" });
    const numbered = now.type === "number" || type === "number";
    const dialog = h("dialog", { class: "list-dialog", "aria-label": "箇条書きと段落番号" },
      h("div", { class: "dialog-head" }, h("h3", {}, "箇条書きと段落番号"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body sh-form-col" },
        h("label", { class: "sh-inline" }, "サイズ（文字に対して） ", size),
        h("label", { class: "sh-inline" }, "色 ", color),
        numbered ? h("label", { class: "sh-inline" }, "開始 ", start) : null,
        h("p", { class: "hint" }, "色はテーマの色です（テーマを変えると一緒に変わります）。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary ld-ok", onclick: () => {
          const props = { mark: color.value, msize: size.value };
          if (numbered) props.start = Math.max(1, Math.min(9999, Math.round(Number(start.value) || 1)));
          dialog.close();
          editor.textFormat("listProps", props);
        } }, "OK")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
  }

  /**
   * The colour menu, as PowerPoint's: テーマの色 (the theme's ten colours over their five lighter and darker tones),
   * 標準の色, then その他の色 and スポイト. A list that is not the theme's (the highlighter's) is a row of swatches.
   */
  function colors(list, current, pick, { none = null, custom = true, note = "", extra = null } = {}) {
    return (close) => {
      const wrap = h("div", { class: "rb-colors" });
      const now = String(current).toLowerCase();
      const sw = (color, label) => h("button", {
        type: "button", class: ["rb-sw", now === color ? "on" : ""], title: `${label}（${color.toUpperCase()}）`, "aria-label": label, style: { "--c": color },
        onmousedown: (e) => e.preventDefault(), onclick: () => { close(); pick(color); },
      });
      if (none) wrap.append(h("button", { type: "button", class: ["rb-none", current === "none" ? "on" : ""], onmousedown: (e) => e.preventDefault(), onclick: () => { close(); pick("none"); } }, none));
      const sheet = E.PALETTE.sheet;
      if (sheet && [E.PALETTE.text, E.PALETTE.fill, E.PALETTE.line].includes(list)) {
        wrap.append(h("div", { class: "rb-colors-head" }, "テーマの色"),
          h("div", { class: "rb-sheet" }, sheet.columns.map((col) => h("div", { class: "rb-sheet-col" }, sw(col.hex, col.label), h("div", { class: "rb-sheet-tones" }, col.variants.map((v) => sw(v.hex, v.label)))))),
          h("div", { class: "rb-colors-head" }, "標準の色"),
          h("div", { class: "rb-swatches rb-standard" }, sheet.standard.map(([color, label]) => sw(color, label))));
      } else {
        wrap.append(h("div", { class: "rb-swatches" }, list.map(([color, label]) => sw(color, label))));
      }
      if (custom) {
        const input = h("input", { type: "color", value: /^#[0-9a-f]{6}$/i.test(current) ? current : "#4472c4", onchange: (event) => { close(); pick(event.target.value); } });
        wrap.append(h("label", { class: "rb-custom" }, input, h("span", {}, "その他の色…"), h("small", {}, "テーマの色はテーマを変えると一緒に変わります")));
      }
      // スポイト: a colour from the slide (or anywhere on the screen); where only the list's colours may be used, the nearest.
      wrap.append(h("button", { type: "button", class: "rb-eyedrop", onmousedown: (e) => e.preventDefault(), onclick: async () => {
        close();
        const got = await pickColor({ root: document.getElementById("stageBody") || document.body, toast: app.toast });
        if (!got) return;
        const color = custom ? got : ops.nearestColor(got, list.map(([c]) => c.toLowerCase()));
        if (color !== got) app.toast(`スポイト：${got.toUpperCase()} にいちばん近い色（${color.toUpperCase()}）にしました`);
        pick(color);
      } }, ico("eyedrop", 16), h("span", {}, "スポイト")));
      // More ways to fill (図・並べて表示), each closing the menu first.
      for (const item of (typeof extra === "function" ? extra() : extra) || []) {
        wrap.append(h("button", { type: "button", class: ["rb-extra", item.on ? "on" : ""], disabled: item.disabled || null, "data-fill": item.key || null, onmousedown: (e) => e.preventDefault(), onclick: () => { close(); item.run(); } }, item.icon ? ico(item.icon, 16) : null, h("span", {}, item.label)));
      }
      if (note) wrap.append(h("p", { class: "rb-note" }, note));
      return wrap;
    };
  }

  /** The shape gallery, by PowerPoint's categories (and the lines). */
  /** 書式のコピー/貼り付け: a click copies for one object; a double-click keeps it on for many (Esc ends it). */
  function painterBtn() {
    const b = btn("painter", "書式", "書式のコピー/貼り付け（⇧⌘C → クリック。ダブルクリックで続けて貼り付け）", () => {
      if (editor.painter?.sticky) { editor.endPainter(); return; }
      if (editor.painter) editor.pasteFormat(editor.selection); else editor.copyFormat();
    }, { enabled: () => any() || Boolean(editor.painter), pressed: () => Boolean(editor.painter) });
    b.addEventListener("dblclick", (event) => { event.preventDefault(); editor.copyFormat({ sticky: true }); });
    return b;
  }
  // 最近使用した図形 (at the top of the shapes gallery, as in PowerPoint; kept in this browser).
  const RECENT_SHAPES = "hpptx-recent-shapes";
  const recentShapes = () => { try { return JSON.parse(localStorage.getItem(RECENT_SHAPES) || "[]").filter((k) => E.SHAPES[k]).slice(0, 12); } catch { return []; } };
  const rememberShape = (key) => { try { localStorage.setItem(RECENT_SHAPES, JSON.stringify([key, ...recentShapes().filter((k) => k !== key)].slice(0, 12))); } catch { /* private window */ } };
  // The shape gallery. Where it draws new shapes (with the lines), a right-click on a shape or a line is PowerPoint's
  // 描画モードのロック: the same one drawn again and again until Esc.
  function shapeGallery(onPick, { lines = true } = {}) {
    const lockable = lines;
    const pickShape = (close, key, lock = false) => { close(); rememberShape(key); onPick({ kind: "shape", shape: key, ...(lock ? { lock: true } : {}) }); };
    const lockOn = (close, pick) => (lockable ? (event) => { event.preventDefault(); pick(true); } : null);
    return (close) => {
      const box = h("div", { class: "rb-gallery shapes" });
      if (lockable) box.append(h("div", { class: "rb-gallery-hint" }, "右クリックで描画モードのロック（同じ図形を続けて描く・Escで終わる）"));
      const recent = recentShapes();
      if (recent.length) {
        box.append(h("div", { class: "rb-gallery-head" }, "最近使用した図形"));
        box.append(h("div", { class: "rb-gallery-grid rb-recent-shapes" }, recent.map((key) => h("button", {
          type: "button", title: E.SHAPES[key].label, "aria-label": E.SHAPES[key].label, "data-shape": key, onclick: () => pickShape(close, key), oncontextmenu: lockOn(close, (lock) => pickShape(close, key, lock)),
        }, shapeThumb(key)))));
      }
      if (lines) {
        box.append(h("div", { class: "rb-gallery-head" }, "線"));
        box.append(h("div", { class: "rb-gallery-grid" }, [["line", "直線"], ["arrow", "矢印"], ["double", "両方向矢印"], ["elbow", "カギ線コネクタ"], ["curve", "曲線コネクタ"]].map(([variant, label]) => h("button", {
          type: "button", title: label, "aria-label": label, onclick: () => { close(); onPick({ kind: "line", variant }); }, oncontextmenu: lockOn(close, () => { close(); onPick({ kind: "line", variant, lock: true }); }),
        }, lineThumb(variant))), [["curve", "曲線"], ["polygon", "フリーフォーム: 図形"], ["scribble", "フリーフォーム: フリーハンド"]].map(([mode, label]) => h("button", {
          type: "button", title: label, "aria-label": label, "data-freeform": mode, onclick: () => { close(); onPick({ kind: "freeform", mode }); },
        }, freeformThumb(mode)))));
      }
      for (const [name, keys] of E.SHAPE_GROUPS) {
        box.append(h("div", { class: "rb-gallery-head" }, name));
        box.append(h("div", { class: "rb-gallery-grid" }, keys.map((key) => h("button", {
          type: "button", title: E.SHAPES[key].label, "aria-label": E.SHAPES[key].label, "data-shape": key, onclick: () => pickShape(close, key), oncontextmenu: lockOn(close, (lock) => pickShape(close, key, lock)),
        }, shapeThumb(key)))));
      }
      return box;
    };
  }
  function shapeThumb(key, w = 26, hh = 20) {
    const g = E.geometry(key, w, hh, E.SHAPES[key].adj);
    const svg = E.s("svg", { viewBox: `-2 -2 ${w + 4} ${hh + 4}`, width: w + 4, height: hh + 4, "aria-hidden": "true" });
    for (const d of g.paths) svg.append(E.s("path", { d, fill: g.open ? "none" : "#dce4f2", stroke: "#1f3864", "stroke-width": 1, "fill-rule": g.rule }));
    for (const extra of g.extras) svg.append(E.s("path", { d: extra.d, fill: extra.tone === "line" ? "none" : extra.tone === "dark" ? "rgba(0,0,0,.15)" : "rgba(255,255,255,.5)", stroke: "#1f3864", "stroke-width": 0.8 }));
    return svg;
  }
  function freeformThumb(mode) {
    const svg = E.s("svg", { viewBox: "0 0 30 24", width: 30, height: 24, "aria-hidden": "true" });
    const d = mode === "curve" ? "M3 18 C8 4 14 4 16 12 S24 22 27 6" : mode === "polygon" ? "M4 19 L9 5 L17 11 L25 4 L22 20 Z" : "M3 16 c3-6 5 4 8-2 s4-8 6 0 s5 6 7-4 s2-4 3 0";
    svg.append(E.s("path", { d, fill: mode === "polygon" ? "#dce4f2" : "none", stroke: "#1f3864", "stroke-width": 1.5, "stroke-linejoin": "round", "stroke-linecap": "round" }));
    return svg;
  }
  function lineThumb(variant) {
    const svg = E.s("svg", { viewBox: "0 0 30 24", width: 30, height: 24, "aria-hidden": "true" });
    const d = variant === "elbow" ? "M4 20 H15 V4 H26" : variant === "curve" ? "M4 20 C15 20 15 4 26 4" : "M4 20 L26 4";
    svg.append(E.s("path", { d, fill: "none", stroke: "#1f3864", "stroke-width": 1.6 }));
    if (variant !== "line") svg.append(E.s("path", { d: variant === "elbow" ? "M21 1 L27 4 L21 7 Z" : variant === "curve" ? "M21 1 L27 4 L21 7 Z" : "M20 3 L26 4 L23 9 Z", fill: "#1f3864" }));
    if (variant === "double") svg.append(E.s("path", { d: "M4 20 L10 21 L7 15 Z", fill: "#1f3864" }));
    return svg;
  }

  function iconGallery(onPick) {
    return (close) => {
      const search = h("input", { type: "search", placeholder: "アイコンを探す（例：時計、店舗）", class: "rb-search" });
      const grid = h("div", { class: "rb-gallery-grid icons" });
      const fill = () => {
        const q = search.value.trim();
        grid.replaceChildren(...Object.entries(E.icons).filter(([key, entry]) => !q || key.includes(q) || String(entry.label || "").includes(q)).map(([key, entry]) => {
          const icon = E.icon(key);
          return h("button", { type: "button", title: entry.label || key, "aria-label": entry.label || key, onclick: () => { close(); onPick(key); } }, icon, h("small", {}, entry.label || key));
        }));
      };
      search.addEventListener("input", fill);
      fill();
      setTimeout(() => search.focus(), 0);
      return h("div", { class: "rb-gallery icons" }, search, grid);
    };
  }

  function photoGallery(onPick) {
    return (close) => h("div", { class: "rb-gallery photos" }, h("div", { class: "rb-gallery-head" }, "内蔵の写真（社内で使える素材）"),
      h("div", { class: "rb-gallery-grid photos" }, Object.entries(E.PHOTOS).map(([key, [file, label]]) => h("button", { type: "button", title: label, onclick: () => { close(); onPick(key); } },
        h("img", { src: `/assets/${file}`, alt: "", loading: "lazy" }), h("small", {}, label)))));
  }

  // ---------------------------------------------------------------- commands shared by the ribbon and the pane

  const fillOf = () => { const o = selected().find((x) => ["shape", "text"].includes(x.kind)); return o ? E.withDefaults(o).fill : null; };
  const strokeOf = () => { const o = selected()[0]; return o ? E.withDefaults(o).stroke : null; };
  const textColorOf = () => editor.textState()?.color ?? "#1a1a1a";
  // A colour replaces a picture fill (図で塗りつぶし), as in PowerPoint.
  const setFill = (color) => editor.apply((o) => (["shape", "text"].includes(o.kind) ? { fill: color, fillImg: undefined, fillTile: undefined, gradient: undefined, pattern: undefined } : null));
  /** 塗りつぶし → グラデーション: a gradient of the theme's colours (kept with its first colour as the fill). */
  const setGradient = (setKey, directionKey) => editor.apply((o) => {
    if (!["shape", "text"].includes(o.kind)) return null;
    const g = ops.makeGradient(setKey, directionKey, E.themeColors(E.PALETTE.sheet?.themeId));
    return { gradient: g, fill: g.stops[0].color, fillImg: undefined, fillTile: undefined, pattern: undefined };
  });
  const gradientOf = () => editor.selectedObjects().find((o) => o.gradient)?.gradient || null;
  const patternOf = () => editor.selectedObjects().find((o) => o.pattern)?.pattern || null;
  /** パターンの塗りつぶし: the kinds drawn small (pick one), then the two colours (the theme's). */
  function patternDialog() {
    const now = patternOf() || { kind: "ltUpDiag", fg: "#b7c3da", bg: "#ffffff" };
    const colors = [...new Map([...E.PALETTE.fill, ...E.PALETTE.line].map(([c, label]) => [c, label])).entries()];
    const fg = select("前景", colors, now.fg);
    const bg = select("背景", E.PALETTE.fill.map(([c, label]) => [c, label]), now.bg);
    let kind = now.kind;
    const grid = h("div", { class: "pat-grid", role: "listbox", "aria-label": "パターン" });
    const tiles = new Map();
    const draw = () => {
      for (const [key, [label]] of Object.entries(E.PATTERN_KINDS)) {
        const id = `pt-${key}`;
        const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
        svg.setAttribute("viewBox", "0 0 48 36");
        svg.setAttribute("width", "48");
        svg.setAttribute("height", "36");
        const defs = document.createElementNS("http://www.w3.org/2000/svg", "defs");
        defs.append(E.patternNode(id, { kind: key, fg: fg.value, bg: bg.value }));
        const rect = document.createElementNS("http://www.w3.org/2000/svg", "rect");
        rect.setAttribute("width", "48");
        rect.setAttribute("height", "36");
        rect.setAttribute("fill", `url(#${id})`);
        svg.append(defs, rect);
        let tile = tiles.get(key);
        if (!tile) {
          tile = h("button", { type: "button", class: "pat-tile", role: "option", "data-pattern": key, title: label, onclick: () => { kind = key; mark(); } });
          tiles.set(key, tile);
          grid.append(tile);
        }
        tile.replaceChildren(svg);
      }
      mark();
    };
    const mark = () => { for (const [key, tile] of tiles) { tile.classList.toggle("on", key === kind); tile.setAttribute("aria-selected", String(key === kind)); } };
    fg.addEventListener("change", draw);
    bg.addEventListener("change", draw);
    draw();
    const dialog = settingsDialog("pattern-dialog", "パターンの塗りつぶし", h("div", {}, grid, labeled("前景", fg), labeled("背景", bg), h("p", { class: "hint" }, "PowerPoint のパターンに近い見た目です。")), () => {
      editor.apply((o) => (["shape", "text"].includes(o.kind) ? { pattern: { kind, fg: fg.value, bg: bg.value }, fill: bg.value, fillImg: undefined, fillTile: undefined, gradient: undefined } : null));
    });
  }
  function gradientDialog() {
    const current = ops.gradientChoice(gradientOf(), E.themeColors(E.PALETTE.sheet?.themeId)) || { set: ops.GRADIENT_SETS[0][0], direction: "down" };
    const sets = select("色の組み合わせ", ops.GRADIENT_SETS.map(([key, label]) => [key, label]), current.set);
    const dirs = select("方向", ops.GRADIENT_DIRECTIONS.map(([key, label]) => [key, label]), current.direction);
    const bar = h("div", { class: "gr-preview", "aria-hidden": "true" });
    const paint = () => { const g = ops.makeGradient(sets.value, dirs.value, E.themeColors(E.PALETTE.sheet?.themeId)); bar.style.background = `linear-gradient(${g.angle + 90}deg, ${g.stops[0].color}, ${g.stops[1].color})`; };
    sets.addEventListener("change", paint);
    dirs.addEventListener("change", paint);
    paint();
    settingsDialog("gradient-dialog", "グラデーションの塗りつぶし", h("div", {}, labeled("色の組み合わせ", sets), labeled("方向", dirs), bar, h("p", { class: "hint" }, "テーマの色の組み合わせです（文字が読める明るさ）。")), () => setGradient(sets.value, dirs.value));
  }
  // ---------------------------------------------------------------- 図形の効果 (影・反射・光彩・ぼかし)
  const EFFECT_KINDS = ["shape", "text", "image"];
  const hasEffectTarget = () => selected().some((o) => EFFECT_KINDS.includes(o.kind));
  const setEffect = (patch) => editor.apply((o) => (EFFECT_KINDS.includes(o.kind) ? patch : null));
  const effectOf = (key) => selected().find((o) => EFFECT_KINDS.includes(o.kind))?.[key];
  function effectsMenu() {
    const [shadow, reflect, glow, soft, rot3d] = ["shadow", "reflect", "glow", "soft", "rot3d"].map(effectOf);
    const direction = ops.shadowDirection(shadow);
    return menu([
      { head: "影" },
      ...ops.SHADOW_DIRECTIONS.map(([key, label]) => ({ label: `影：${label}`, on: direction === key, run: () => setEffect({ shadow: ops.shadowPreset(key) }) })),
      { label: "影なし", disabled: !shadow, run: () => setEffect({ shadow: undefined }) },
      "-", { head: "反射" },
      ...ops.REFLECTIONS.map(([key, label]) => ({ label: `反射：${label}`, on: Boolean(reflect) && Math.abs(reflect.size - ops.reflectionPreset(key).size) < 0.01, run: () => setEffect({ reflect: ops.reflectionPreset(key) }) })),
      { label: "反射なし", disabled: !reflect, run: () => setEffect({ reflect: undefined }) },
      "-", { head: "光彩" },
      ...ops.GLOW_SIZES.map(([r, label]) => ({ label: `光彩：${label}`, on: glow?.r === r, run: () => setEffect({ glow: { r, color: glow?.color || ops.GLOW_COLORS[2][0], opacity: glow?.opacity ?? 0.6 } }) })),
      { label: "光彩なし", disabled: !glow, run: () => setEffect({ glow: undefined }) },
      "-", { head: "ぼかし（縁をぼかす）" },
      ...ops.SOFT_SIZES.map(([r, label]) => ({ label: `ぼかし：${label}`, on: soft === r, run: () => setEffect({ soft: r }) })),
      { label: "ぼかしなし", disabled: !soft, run: () => setEffect({ soft: undefined }) },
      "-", { head: "3-D 回転" },
      ...ops.ROT3D_PRESETS.map(([key, label]) => ({ label: `3-D：${label}`, on: ops.rot3dKey(rot3d) === key, run: () => setEffect({ rot3d: ops.rot3dPreset(key) }) })),
      { label: "3-D 回転なし", disabled: !rot3d, run: () => setEffect({ rot3d: undefined }) },
      "-", { label: "効果のオプション…", icon: "effects", run: () => effectsDialog() },
      { label: "すべての効果をなくす", icon: "trash", disabled: !(shadow || reflect || glow || soft || rot3d), run: () => setEffect({ shadow: undefined, reflect: undefined, glow: undefined, soft: undefined, rot3d: undefined }) },
    ]);
  }
  /** 効果のオプション: each effect's numbers (影＝色・透明度・ぼかし・距離・角度、反射＝大きさ・透明度・間隔、光彩＝大きさ・色・透明度、ぼかし＝大きさ). */
  function effectsDialog() {
    const polar = ops.shadowPolar(effectOf("shadow"));
    const reflect = effectOf("reflect");
    const glow = effectOf("glow");
    const soft = effectOf("soft");
    const rot3d = effectOf("rot3d");
    const input = (name, value, { min = 0, max = 100, step = 1 } = {}) => h("input", { type: "number", name, "aria-label": name, value: String(value), min: String(min), max: String(max), step: String(step) });
    const on = (name, checked) => h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name, checked: checked || null }), h("span", {}, "付ける"));
    const sColor = select("shadowColor", ops.SHADOW_COLORS, polar.color);
    const gColor = select("glowColor", ops.GLOW_COLORS, glow?.color || ops.GLOW_COLORS[2][0]);
    const body = h("div", { class: "fx-body" },
      fieldset("影", on("shadowOn", Boolean(effectOf("shadow"))), labeled("色", sColor), labeled("透明度（%）", input("shadowTransparency", Math.round((1 - polar.opacity) * 100))), labeled("ぼかし（pt）", input("shadowBlur", polar.blur, { max: 100, step: 0.5 })), labeled("距離（pt）", input("shadowDistance", polar.distance, { max: 100, step: 0.5 })), labeled("角度（度）", input("shadowAngle", polar.angle, { min: -180, max: 180 }))),
      fieldset("反射", on("reflectOn", Boolean(reflect)), labeled("大きさ（%）", input("reflectSize", Math.round((reflect?.size ?? 0.5) * 100), { min: 5 })), labeled("透明度（%）", input("reflectTransparency", Math.round((1 - (reflect?.opacity ?? 0.4)) * 100))), labeled("間隔（pt）", input("reflectGap", (reflect?.gap ?? 2) / 2, { max: 50, step: 0.5 }))),
      fieldset("光彩", on("glowOn", Boolean(glow)), labeled("大きさ（pt）", input("glowSize", (glow?.r ?? 16) / 2, { min: 0.5, max: 100, step: 0.5 })), labeled("色", gColor), labeled("透明度（%）", input("glowTransparency", Math.round((1 - (glow?.opacity ?? 0.6)) * 100)))),
      fieldset("ぼかし", on("softOn", Boolean(soft)), labeled("大きさ（pt）", input("softSize", (soft ?? 10) / 2, { min: 0.5, max: 100, step: 0.5 }))),
      fieldset("3-D 回転", on("rot3dOn", Boolean(rot3d)), labeled("X 回転（°）", input("rotX", rot3d?.x ?? 0, { min: -89, max: 89, step: 1 })), labeled("Y 回転（°）", input("rotY", rot3d?.y ?? 0, { min: -89, max: 89, step: 1 })), labeled("遠近（°。0 は平行）", input("rotP", rot3d?.p ?? 45, { min: 0, max: 120, step: 1 })),
        h("p", { class: "hint" }, "X が正だと上が、Y が正だと右が奥に下がります。")));
    settingsDialog("fx-dialog", "図形の効果のオプション", body, () => {
      const dlg = document.querySelector(".fx-dialog");
      const f = (name) => dlg.querySelector(`[name="${name}"]`);
      const n = (name, fallback = 0) => { const v = Number(f(name).value); return Number.isFinite(v) ? v : fallback; };
      const patch = {
        shadow: f("shadowOn").checked ? ops.shadowFromPolar({ distance: n("shadowDistance"), angle: n("shadowAngle"), blur: n("shadowBlur"), opacity: 1 - n("shadowTransparency") / 100, color: f("shadowColor").value }) : undefined,
        reflect: f("reflectOn").checked ? { size: Math.min(1, Math.max(0.05, n("reflectSize", 50) / 100)), opacity: 1 - n("reflectTransparency", 60) / 100, gap: Math.round(n("reflectGap") * 2) } : undefined,
        glow: f("glowOn").checked ? { r: Math.max(1, Math.round(n("glowSize", 8) * 2)), color: f("glowColor").value, opacity: 1 - n("glowTransparency", 40) / 100 } : undefined,
        soft: f("softOn").checked ? Math.max(1, Math.round(n("softSize", 5) * 2)) : undefined,
        rot3d: f("rot3dOn").checked && (Math.abs(n("rotX")) >= 0.5 || Math.abs(n("rotY")) >= 0.5) ? { x: Math.max(-89, Math.min(89, n("rotX"))), y: Math.max(-89, Math.min(89, n("rotY"))), p: Math.max(0, Math.min(120, Math.round(n("rotP", 45)))) } : undefined,
      };
      setEffect(patch);
    });
  }
  // ---------------------------------------------------------------- 文字の効果 (文字の影・光彩)
  const textFxOf = (key) => editor.textState()?.[key];
  const setTextFx = (patch) => editor.textFormat("tfx", { tshadow: textFxOf("tshadow"), tglow: textFxOf("tglow"), toutline: textFxOf("toutline"), ...patch });
  function textEffectsMenu() {
    const [shadow, glow, outline] = [textFxOf("tshadow"), textFxOf("tglow"), textFxOf("toutline")];
    const direction = ops.shadowDirection(shadow);
    return menu([
      { head: "文字の影" },
      ...ops.SHADOW_DIRECTIONS.map(([key, label]) => ({ label: `影：${label}`, on: direction === key, run: () => setTextFx({ tshadow: ops.textShadowPreset(key) }) })),
      { label: "影なし", disabled: !shadow, run: () => setTextFx({ tshadow: undefined }) },
      "-", { head: "光彩" },
      ...ops.TEXT_GLOW_SIZES.map(([r, label]) => ({ label: `光彩：${label}`, on: glow?.r === r, run: () => setTextFx({ tglow: { r, color: glow?.color || ops.GLOW_COLORS[2][0], opacity: glow?.opacity ?? 0.7 } }) })),
      { label: "光彩なし", disabled: !glow, run: () => setTextFx({ tglow: undefined }) },
      "-", { head: "文字の輪郭" },
      ...ops.TEXT_OUTLINE_COLORS.map(([color, label]) => ({ label: `輪郭：${label}`, on: outline?.color === color, run: () => setTextFx({ toutline: ops.textOutline(color, outline) }) })),
      ...ops.TEXT_OUTLINE_WIDTHS.map(([w, label]) => ({ label: `輪郭の太さ：${label}`, on: outline?.w === w, run: () => setTextFx({ toutline: { w, color: outline?.color || ops.TEXT_OUTLINE_COLORS[1][0] } }) })),
      { label: "輪郭なし", disabled: !outline, run: () => setTextFx({ toutline: undefined }) },
      "-", { label: "文字の効果のオプション…", icon: "effects", run: () => textEffectsDialog() },
      { label: "すべての文字の効果をなくす", icon: "trash", disabled: !(shadow || glow || outline), run: () => setTextFx({ tshadow: undefined, tglow: undefined, toutline: undefined }) },
    ]);
  }
  /** 文字の効果のオプション: 影＝色・透明度・ぼかし・距離・角度、光彩＝大きさ・色・透明度 (pt, like the shape's). */
  function textEffectsDialog() {
    const polar = ops.shadowPolar(textFxOf("tshadow") || { dx: 3, dy: 3, blur: 4, color: "#000000", opacity: 0.4 });
    const glow = textFxOf("tglow");
    const outline = textFxOf("toutline");
    const input = (name, value, { min = 0, max = 100, step = 1 } = {}) => h("input", { type: "number", name, "aria-label": name, value: String(value), min: String(min), max: String(max), step: String(step) });
    const on = (name, checked) => h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name, checked: checked || null }), h("span", {}, "付ける"));
    const sColor = select("shadowColor", ops.SHADOW_COLORS, polar.color);
    const gColor = select("glowColor", ops.GLOW_COLORS, glow?.color || ops.GLOW_COLORS[2][0]);
    const oColor = select("outlineColor", ops.TEXT_OUTLINE_COLORS, outline?.color || ops.TEXT_OUTLINE_COLORS[1][0]);
    const body = h("div", { class: "fx-body" },
      fieldset("文字の影", on("shadowOn", Boolean(textFxOf("tshadow"))), labeled("色", sColor), labeled("透明度（%）", input("shadowTransparency", Math.round((1 - polar.opacity) * 100))), labeled("ぼかし（pt）", input("shadowBlur", polar.blur, { max: 50, step: 0.5 })), labeled("距離（pt）", input("shadowDistance", polar.distance, { max: 50, step: 0.5 })), labeled("角度（°）", input("shadowAngle", polar.angle, { min: -180, max: 180 }))),
      fieldset("文字の光彩", on("glowOn", Boolean(glow)), labeled("大きさ（pt）", input("glowSize", (glow?.r ?? 10) / 2, { min: 0.5, max: 50, step: 0.5 })), labeled("色", gColor), labeled("透明度（%）", input("glowTransparency", Math.round((1 - (glow?.opacity ?? 0.7)) * 100)))),
      fieldset("文字の輪郭", on("outlineOn", Boolean(outline)), labeled("太さ（pt）", input("outlineWidth", (outline?.w ?? 2) / 2, { min: 0.25, max: 12, step: 0.25 })), labeled("色", oColor)));
    settingsDialog("tfx-dialog", "文字の効果のオプション", body, () => {
      const dlg = document.querySelector(".tfx-dialog");
      const f = (name) => dlg.querySelector(`[name="${name}"]`);
      const n = (name, fallback = 0) => { const v = Number(f(name).value); return Number.isFinite(v) ? v : fallback; };
      setTextFx({
        tshadow: f("shadowOn").checked ? ops.shadowFromPolar({ distance: n("shadowDistance"), angle: n("shadowAngle"), blur: n("shadowBlur"), opacity: 1 - n("shadowTransparency") / 100, color: f("shadowColor").value }) : undefined,
        tglow: f("glowOn").checked ? { r: Math.max(1, Math.round(n("glowSize", 5) * 2)), color: f("glowColor").value, opacity: 1 - n("glowTransparency", 30) / 100 } : undefined,
        toutline: f("outlineOn").checked ? { w: Math.max(0.5, Math.round(n("outlineWidth", 1) * 4) / 2), color: f("outlineColor").value } : undefined,
      });
    });
  }
  /** 図形の塗りつぶし → 図: a picture from this device fills the selected shapes (cut to their outline). */
  async function pictureFill() {
    const [file] = await app.pickFiles("image/png,image/jpeg,image/webp,image/gif");
    if (!file) return;
    try {
      const { src } = await app.storePicture(file);
      editor.apply((o) => (["shape", "text"].includes(o.kind) ? { fillImg: src } : null));
      app.toast("図で塗りつぶしました（図形の形に切り抜き。「並べて表示」で繰り返しにできます）");
    } catch (error) { app.toast(`画像を読み込めませんでした（${error.message}）`); }
  }
  const pictureFilled = () => editor.selectedObjects().some((o) => o.fillImg);
  const fillExtras = () => [
    { key: "picture", label: "図…", icon: "image", run: () => pictureFill() },
    { key: "tile", label: "図を並べて表示", icon: "grid", on: editor.selectedObjects().some((o) => o.fillTile), disabled: !pictureFilled(), run: () => { const on = !editor.selectedObjects().some((o) => o.fillTile); editor.apply((o) => (o.fillImg ? { fillTile: on || undefined } : null)); } },
    { key: "gradient", label: "グラデーション…", icon: "gradient", run: () => gradientDialog() },
    { key: "nogradient", label: "グラデーションを外す", icon: "trash", disabled: !gradientOf(), run: () => editor.apply((o) => (o.gradient ? { gradient: undefined } : null)) },
    { key: "pattern", label: "パターン…", icon: "grid", run: () => patternDialog() },
    { key: "nopattern", label: "パターンを外す", icon: "trash", disabled: !patternOf(), run: () => editor.apply((o) => (o.pattern ? { pattern: undefined } : null)) },
    { key: "nopicture", label: "図を外す", icon: "trash", disabled: !pictureFilled(), run: () => editor.apply((o) => (o.fillImg ? { fillImg: undefined, fillTile: undefined } : null)) },
  ];
  const setStroke = (color) => editor.apply((o) => (["shape", "text", "image", "line"].includes(o.kind) ? { stroke: color, ...(color !== "none" && o.kind !== "line" && E.withDefaults(o).stroke === "none" ? { strokeW: E.withDefaults(o).strokeW || 2 } : {}) } : o.kind === "icon" ? { color } : null));
  const setStrokeWidth = (pt) => editor.apply((o) => (["shape", "text", "image", "line"].includes(o.kind) ? { strokeW: ops.fromPt(pt), ...(E.withDefaults(o).stroke === "none" && o.kind !== "line" ? { stroke: "#1f3864" } : {}) } : o.kind === "icon" ? { strokeW: Math.max(0.5, pt) } : null));
  /** 線の書式設定 (複合線・線端・結合点): a change to the outline of whatever has one. */
  const setLineLook = (patch) => editor.apply((o) => (["shape", "text", "image", "line"].includes(o.kind) ? patch : null));
  const setDash = (dash) => editor.apply((o) => (["shape", "text", "image", "line"].includes(o.kind) ? { dash: dash === "solid" ? undefined : dash } : null));

  async function insertImages() {
    const files = await app.pickFiles("image/*", true);
    if (files.length) { await app.insertFiles(files); showTab("picture"); }
  }
  async function insertVideoFile() {
    const files = await app.pickFiles("video/mp4,video/webm,video/quicktime,.json,application/json", false);
    if (files.length) await app.insertFiles(files);
  }
  async function insertFromUrl(kind) {
    const url = await app.ask(kind === "image" ? "URLから画像を入れる" : "URLで動画（YouTube）・アニメーションを入れる", kind === "image" ? "画像のURL（https://…）" : "YouTubeのURL、動画（mp4）やLottie（.json）のURL", "");
    if (url) await app.insertUrl(url.trim(), kind);
  }
  /** 挿入 → QR コード: the words (a web address, usually) as a QR code picture; an address is also a link, clicked while presenting. */
  async function insertQr() {
    const answer = await app.ask("QR コード", "QR コードにする文字（Webページのアドレスなど。300文字まで）", "https://");
    const value = String(answer ?? "").trim();
    if (!value || value === "https://") return;
    if ([...value].length > 300) { app.toast("QR コードにできるのは300文字までです"); return; }
    try {
      const { qrSvg } = await import("./online.mjs");
      const svg = await qrSvg(value);
      const link = /^https?:\/\/\S+$/i.test(value);
      const [o] = editor.insert([ops.makeObject("image", { x: 0, y: 0, w: 360, h: 360 }, { src: `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`, alt: `QR コード：${value.slice(0, 80)}`, ...(link ? { action: { type: "url", href: value } } : {}) })]);
      if (o) showTab("picture");
    } catch { app.toast("QR コードを作れませんでした"); }
  }
  function insertIcon(key) {
    const [o] = editor.insert([ops.makeObject("icon", { x: 0, y: 0, w: 144, h: 144 }, { icon: key, color: "#1f3864" })]);
    if (o) showTab("shape");
  }
  function insertPhoto(key) {
    const [file] = E.PHOTOS[key];
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(960 / img.naturalWidth, 640 / img.naturalHeight, 1);
      editor.insert([ops.makeObject("image", { x: 0, y: 0, w: img.naturalWidth * scale, h: img.naturalHeight * scale }, { src: `asset:${key}`, alt: E.PHOTOS[key][1] })]);
      showTab("picture");
    };
    img.onerror = () => editor.insert([ops.makeObject("image", { x: 0, y: 0, w: 960, h: 540 }, { src: `asset:${key}` })]);
    img.src = `/assets/${file}`;
  }
  /** 挿入 → 記号と特殊文字・日付と時刻: at the caret while typing, else a new text box in the middle of the slide. */
  function insertWords(text) {
    if (editor.typing && editor.restoreRange()) { document.execCommand("insertText", false, text); return; }
    const width = Math.min(1600, Math.max(160, [...text].length * 46 + 48));
    editor.insert([ops.makeObject("text", { x: 0, y: 0, w: width, h: 96 }, { text: `<p>${text}</p>`, fs: 40 })]);
  }
  function symbolGallery() {
    return (close) => h("div", { class: "rb-gallery symbols" }, h("div", { class: "rb-gallery-head" }, "記号と特殊文字"),
      h("div", { class: "rb-symbols" }, SYMBOLS.map((ch) => h("button", { type: "button", title: ch, "aria-label": ch, onmousedown: (e) => e.preventDefault(), onclick: () => { close(); insertWords(ch); } }, ch))));
  }
  function dateMenu() {
    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth() + 1;
    const d = now.getDate();
    const wd = "日月火水木金土"[now.getDay()];
    const era = y >= 2019 ? `令和${y - 2018 === 1 ? "元" : y - 2018}年${m}月${d}日` : `${y}年${m}月${d}日`;
    const pad = (n) => String(n).padStart(2, "0");
    return menu([{ head: "日付と時刻（今日）" }, ...[`${y}年${m}月${d}日`, `${y}年${m}月${d}日（${wd}）`, `${y}/${pad(m)}/${pad(d)}`, `${y}年${m}月`, era, `${m}月${d}日（${wd}）`, `${pad(now.getHours())}:${pad(now.getMinutes())}`].map((text) => ({ label: text, run: () => insertWords(text) }))]);
  }
  function pickTool(tool) {
    if (tool.kind === "freeform") { freeform.start(tool.mode); return; }
    editor.setTool(tool);
    if (tool.lock) { app.toast("描画モードをロックしました。同じ図形を続けて描けます（Escで終わる）"); return; }
    app.toast(tool.kind === "line" ? "スライド上をドラッグして線を引いてください（Shiftで水平・垂直・45°）" : tool.kind === "text" ? "クリックまたはドラッグでテキストボックスを置き、そのまま入力できます" : "スライド上をドラッグして描いてください（クリックで標準の大きさ。Shiftで縦横同じ比率）");
  }

  async function editLink() {
    // While typing, the words picked become a link (挿入 → リンク), as in PowerPoint.
    if (editor.typing) {
      const selection = window.getSelection();
      const picked = selection && !selection.isCollapsed && selection.toString().trim();
      if (picked) {
        const range = selection.getRangeAt(0).cloneRange();
        const url = await app.ask("リンクの挿入", `「${picked.slice(0, 30)}」から開くWebページ（https://…）かメール（mailto:…）。空にするとリンクを外します`, "https://");
        if (url == null) return;
        editor.restoreRange();
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        if (!url.trim() || url.trim() === "https://") { editor.textFormat("unlink"); return; }
        if (!/^(https?:\/\/|mailto:)/i.test(url.trim())) { app.toast("URLは https:// か mailto: で始めてください"); return; }
        editor.textFormat("link", url.trim());
        app.toast("リンクを付けました（発表中にクリックすると開きます）");
        return;
      }
    }
    const o = one();
    if (!o) return app.toast("リンクを付けるオブジェクトを1つ選んでください");
    openPop(document.querySelector('[data-rb="link"]') || ribbon, (close) => linkForm(o, close));
  }

  // PowerPoint's 動作設定: マウスのクリック and マウスの通過, each a jump (or a link, for a click) and a sound. An HTML
  // click (詳細を開く・裏返す…, set on the インタラクション tab) stays as it is unless another click is chosen.
  function linkForm(o, close = null) {
    const slideOptions = (a) => app.deck().slides.map((slide, i) => h("option", { value: String(i), selected: a?.type === "slide" && app.slideOfSid(a.to) === i ? true : null }, `${i + 1}. ${E.strip(slide.title || slide.message || E.TYPE_LABELS[slide.type] || "")}`.slice(0, 40)));
    const soundSelect = (a, label) => h("select", { "aria-label": label, class: "rb-link-sound" }, h("option", { value: "" }, "サウンドなし"), Object.entries(E.TRANSITION_SOUNDS).map(([v, l]) => h("option", { value: v, selected: a?.sound === v || null }, `サウンド：${l}`)));
    // One way of acting (click or mouse over): what it does, where to, and its sound.
    function part(a, { label, choices, keep, keepNames = false }) {
      const current = keep ? "keep" : a?.type === "sound" ? "none" : a?.type || "none";
      const kind = h("select", { "aria-label": label }, keep ? h("option", { value: "keep", selected: true }, `${keep}（インタラクション）`) : null, choices.map(([value, text]) => h("option", { value, selected: value === current || null }, text)));
      const slides = h("select", { "aria-label": keepNames ? "移動先のスライド" : `${label}：移動先のスライド`, hidden: current !== "slide" }, slideOptions(a));
      const url = choices.some(([v]) => v === "url") ? h("input", { type: "url", placeholder: "https://… または mailto:…", value: a?.type === "url" ? a.href : "", hidden: current !== "url", "aria-label": `${label}：URL` }) : null;
      const sound = soundSelect(a, `${label}：サウンド`);
      const sync = () => { slides.hidden = kind.value !== "slide"; if (url) url.hidden = kind.value !== "url"; sound.disabled = kind.value === "keep"; };
      kind.addEventListener("change", sync);
      sync();
      const read = () => {
        const type = kind.value;
        const extra = sound.value ? { sound: sound.value } : {};
        if (type === "keep") return { value: o.action };
        if (type === "none") return { value: sound.value ? { type: "sound", sound: sound.value } : undefined };
        if (type === "slide") return { value: { type, to: app.ensureSid(Number(slides.value)), ...extra } };
        if (type === "url") return /^(https?:\/\/|mailto:)/i.test(url.value.trim()) ? { value: { type, href: url.value.trim(), ...extra } } : { error: "URLは https:// か mailto: で始めてください" };
        return { value: { type, ...extra } };
      };
      return { nodes: [kind, slides, url, sound], read };
    }
    const htmlClick = o.action && E.IX_CLICKS[o.action.type] ? E.IX_CLICKS[o.action.type] : null;
    const click = part(o.action, { label: "クリックしたときの動作", choices: ACTIONS, keep: htmlClick, keepNames: true });
    const over = part(o.overAction, { label: "マウスの通過時の動作", choices: ACTIONS.filter(([v]) => v !== "url") });
    // ヒント設定 (PowerPoint's ScreenTip): words shown when the pointer is on it in the show (the object's tip).
    const tip = h("input", { type: "text", class: "rb-link-tip", maxlength: 200, placeholder: "スクリーンヒント（マウスを乗せたときの説明・任意）", value: o.tip || "", "aria-label": "スクリーンヒント" });
    const save = () => {
      const a = click.read();
      const b = over.read();
      const error = a.error || b.error;
      if (error) { app.toast(error); return; }
      editor.apply({ action: a.value, overAction: b.value, tip: tip.value.trim() || undefined }, { ids: [o.id] });
      close?.();
      app.toast(a.value || b.value ? "動作を設定しました（発表中とHTML出力で働きます）" : "動作を外しました");
    };
    return h("div", { class: "rb-form rb-link-form" },
      h("b", {}, "クリックしたときの動作（発表中）"), ...click.nodes,
      h("b", {}, "マウスの通過時の動作（発表中）"), ...over.nodes, tip,
      h("div", { class: "rb-form-foot" }, h("button", { type: "button", class: "btn btn-primary btn-sm", onclick: save }, "設定する")));
  }

  // ---------------------------------------------------------------- the ribbon

  // PowerPoint's order: the tabs that are always there, then the ones the selection brings (図形の書式・図の形式).
  const TABS = [
    { id: "home", label: "ホーム" },
    { id: "insert", label: "挿入" },
    { id: "draw", label: "描画" },
    { id: "design", label: "デザイン" },
    { id: "transition", label: "画面切り替え" },
    { id: "animation", label: "アニメーション" },
    { id: "interact", label: "インタラクション" },
    { id: "slideshow", label: "スライド ショー" },
    { id: "review", label: "校閲" },
    { id: "view", label: "表示" },
    { id: "developer", label: "開発" },
    { id: "shape", label: "図形の書式", contextual: () => hasShape() },
    { id: "picture", label: "図の形式", contextual: () => hasImage() },
    { id: "tableDesign", label: "テーブル デザイン", contextual: () => tables.isTable() },
    { id: "tableLayout", label: "レイアウト", contextual: () => tables.isTable() },
    { id: "chartDesign", label: "グラフのデザイン", contextual: () => tables.isChart() },
    { id: "playback", label: "再生", contextual: () => media.has() },
    { id: "smartartDesign", label: "SmartArt のデザイン", contextual: () => smart.isSmartArt() },
    { id: "zoomTool", label: "ズーム", contextual: () => extras.isZoom() },
    { id: "model3d", label: "3D モデル", contextual: () => models3d.isModel() },
  ];

  function signature() {
    const ctx = TABS.filter((t) => t.contextual?.()).map((t) => t.id).join(",");
    return `${tab}|${ctx}|${ribbonMode}|${floating}|${onSlide()}|${Boolean(app.isFinal?.())}`;
  }

  // ---------------------------------------------------------------- ribbon display options
  function setRibbonMode(mode) {
    if (!RIBBON_MODES[mode]) return;
    ribbonMode = mode;
    floating = false;
    try { localStorage.setItem("hpptx-ribbon-mode", mode); } catch { /* private window */ }
    closePop(true);
    renderRibbon(true);
  }
  /** ⌘F1 / Ctrl+F1, a double-click on a tab, the ˄ button: the ribbon folds to its tabs, or comes back. */
  function toggleRibbon() { setRibbonMode(ribbonMode === "full" ? "tabs" : "full"); }
  function openFloating(id = tab) { tab = id; floating = true; renderRibbon(true); editor.draw(); }
  function closeFloating() { if (!floating) return; floating = false; closePop(true); renderRibbon(true); }
  function floatingOutside(event) { if (floating && !ribbon.contains(event.target) && !pop?.el.contains(event.target)) closeFloating(); }
  function floatingKey(event) { if (floating && !pop && event.key === "Escape") { event.stopPropagation(); closeFloating(); } }
  function fileItems() {
    return [
      { head: "ファイル" },
      { label: "新規（作成画面へ）", icon: "plus", run: () => app.goCreate() },
      { label: "開く（保存庫）…", icon: "pane", run: () => app.openLibrary() },
      { label: "上書き保存", icon: "down", keys: "⌘S", run: () => app.saveDeck() },
      { label: "名前を付けて保存（別の資料として）", icon: "duplicate", run: () => app.saveDeck({ asNew: true }) },
      "-", { head: "書き出し" },
      { label: "PowerPoint（.pptx・編集できる）", icon: "slide", run: () => app.exportPptx() },
      { label: "HTMLファイル（動きごと）", icon: "play", run: () => app.exportHtml() },
      { label: "PDF・印刷（配布資料・ノート）…", icon: "print", keys: "⌘P", run: () => app.printPdf() },
      { label: "JSONで保存", icon: "down", run: () => app.saveJson() },
      { label: "レビュー用ファイル", icon: "review", run: () => app.exportReview() },
      { label: "ビデオの作成（WebM）…", icon: "video2", run: () => app.openVideoExport() },
      { label: "画像として保存（PNG・JPEG）…", icon: "image", run: () => app.openImageExport() },
      { label: "アニメーション GIF の作成…", icon: "play", run: () => app.openGifExport() },
      { label: "HTMLファイル（パスワードで保護）…", icon: "lock", run: () => app.exportLockedHtml() },
      "-",
      { label: "JSONを読み込む…", icon: "up", run: () => app.openJson() },
      { label: "版の履歴・過去の資料…", icon: "clock", run: () => app.openHistory() },
      "-", { label: "情報（プロパティ・ドキュメント検査）…", icon: "info", run: () => app.openFileInfo() },
      { label: "共有（共同編集）…", icon: "share", run: () => app.share() },
      "-", { head: "オプション" },
      { label: "オプション…", icon: "effectOpts", run: () => optionsDialog() },
      { label: `ユーザー名：${app.userName() || "未設定"}…`, icon: "people", run: () => app.changeUserName() },
      { label: "使い方とショートカット", icon: "tip", keys: "?", run: () => app.openHelp() },
    ];
  }
  const fileMenu = () => menu(fileItems());
  /**
   * ファイル → オプション (PowerPoint's Options): the settings kept in this browser, in one place — the user name,
   * proofing (spelling, AutoCorrect), the ribbon and the Quick Access Toolbar, the slide show, PowerPoint import.
   */
  /** Which defaults this deck has (右クリック →「既定の図形に設定」など). */
  function defaultsSummary() {
    const own = app.objectDefaults() || {};
    const names = [own.shape && "図形", own.text && "テキスト ボックス", own.line && "線"].filter(Boolean);
    return names.length ? `${names.join("・")}の既定の書式が設定されています（図形を右クリック →「既定の図形に設定」など）。` : "図形・テキスト ボックス・線は標準の書式で描かれます。図形を右クリック →「既定の図形に設定」で変えられます。";
  }
  function optionsDialog() {
    const check = (key, label, on, set, hint = "") => h("label", { class: "sh-choice" }, h("input", { type: "checkbox", "data-opt": key, checked: on || null, onchange: (event) => set(event.target.checked) }), h("span", {}, label), hint ? h("small", {}, hint) : null);
    const section = (title, ...rows) => h("fieldset", { class: "opt-sec" }, h("legend", {}, title), ...rows);
    const modeSel = h("select", { "data-opt": "ribbon", "aria-label": "リボンの表示", onchange: (event) => setRibbonMode(event.target.value) }, Object.entries(RIBBON_MODES).map(([v, l]) => h("option", { value: v, selected: ribbonMode === v || null }, l)));
    const qatSel = h("select", { "data-opt": "qat", "aria-label": "クイック アクセス ツール バー", onchange: (event) => { const v = event.target.value; quick?.setPlace(v === "off" ? { show: false } : { show: true, below: v === "below" }); } },
      [["off", "表示しない"], ["above", "リボンの上（タブの左）"], ["below", "リボンの下"]].map(([v, l]) => h("option", { value: v, selected: (quick?.where() || "off") === v || null }, l)));
    const name = h("span", { class: "opt-name" }, app.userName() || "未設定");
    const dialog = h("dialog", { class: "options-dialog", "aria-label": "オプション" },
      h("div", { class: "dialog-head" }, h("h3", {}, "オプション"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body opt-body" },
        section("全般", h("div", { class: "sh-inline" }, "ユーザー名：", name, " ", h("button", { type: "button", class: "btn btn-sm", onclick: () => { dialog.close(); app.changeUserName(); } }, "変更…"))),
        section("文章校正",
          check("spell", "入力中にスペル チェックを行う", app.spellcheck(), (on) => app.setSpellcheck(on)),
          h("div", {}, h("button", { type: "button", class: "btn btn-sm opt-ac", onclick: () => { dialog.close(); app.openAutoCorrect(); } }, "オートコレクトのオプション…"))),
        section("リボンとツール バー",
          h("label", { class: "sh-inline" }, "リボンの表示 ", modeSel),
          h("label", { class: "sh-inline" }, "クイック アクセス ツール バー ", qatSel),
          h("div", {}, h("button", { type: "button", class: "btn btn-sm", onclick: () => { dialog.close(); quick?.customize(); } }, "クイック アクセス ツール バーのユーザー設定…"))),
        section("スライド ショー",
          check("presenter", "発表者ビューを使う（別のウィンドウに次のスライド・ノート・時間）", app.presenterView(), (on) => app.setPresenterView(on))),
        section("取り込み",
          check("autoHtml", "PowerPointを見た目どおりに取り込むと、HTMLの動き（おまかせ）を自動で付ける", app.autoHtml(), (on) => app.setAutoHtml(on))),
        h("p", { class: "hint" }, "ここまでの設定はこのブラウザに保存されます。"),
        // 既定の図形 (kept on this deck): what is set, and a way back to the studio's own.
        section("この資料の既定の書式",
          h("p", { class: "opt-defaults" }, defaultsSummary()),
          h("div", {}, h("button", { type: "button", class: "btn btn-sm opt-defaults-reset", disabled: !app.objectDefaults() || null, onclick: (event) => { app.clearObjectDefaults(); event.currentTarget.disabled = true; event.currentTarget.closest(".opt-sec").querySelector(".opt-defaults").textContent = defaultsSummary(); } }, "既定の図形・テキスト ボックス・線を元に戻す")))),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-primary", onclick: () => dialog.close() }, "閉じる")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
  }
  // ---------------------------------------------------------------- フォント・段落・グリッドとガイド (dialogs)

  function settingsDialog(cls, title, body, apply) {
    const dialog = h("dialog", { class: ["fmt-dialog", cls], "aria-label": title },
      h("div", { class: "dialog-head" }, h("h3", {}, title), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("form", { class: "dialog-body fmt-body", onsubmit: (event) => { event.preventDefault(); editor.batch(apply); dialog.close(); } }, body,
        h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"), h("button", { type: "submit", class: "btn btn-primary fmt-ok" }, "OK"))));
    document.body.append(dialog);
    dialog.addEventListener("close", () => { dialog.remove(); refresh(); });
    dialog.showModal();
    return dialog;
  }
  const fieldset = (legend, ...rows) => h("fieldset", { class: "opt-sec" }, h("legend", {}, legend), ...rows);
  const labeled = (label, control) => h("label", { class: "sh-inline fmt-field" }, h("span", {}, label), control);
  const select = (name, options, value) => h("select", { name, "aria-label": name }, options.map(([v, l]) => h("option", { value: String(v), selected: String(v) === String(value) || null }, l)));
  const number = (name, value, { step = 0.1, min = 0, max = 100 } = {}) => h("input", { type: "number", name, "aria-label": name, value: String(value), step: String(step), min: String(min), max: String(max), class: "fmt-num" });
  const checkbox = (name, label, on) => h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name, checked: on || null }), h("span", {}, label));

  /** フォント (⇧⌘F): the font, its style and size, the text colour, the underline, effects and spacing at once. */
  function fontDialog() {
    const st = editor.textState();
    if (!st) { app.toast("文字を入れた図形・テキスト ボックスを選んでください"); return; }
    const style = st.bold && st.italic ? "bolditalic" : st.bold ? "bold" : st.italic ? "italic" : "regular";
    const font = select("フォント", Object.entries(E.FONTS).map(([key, [label]]) => [key, label]), st.font || "body");
    const styleSel = select("スタイル", [["regular", "標準"], ["bold", "太字"], ["italic", "斜体"], ["bolditalic", "太字 斜体"]], style);
    const size = number("サイズ", st.fs, { step: 0.5, min: 6, max: 400 });
    let color = String(st.color || "").toLowerCase();
    const swatches = h("span", { class: "fp-swatches fmt-colors" }, E.PALETTE.text.map(([c, l]) => h("button", { type: "button", class: ["fp-sw", color === c ? "on" : ""], title: l, "aria-label": l, "data-color": c, style: { "--c": c }, onclick: (event) => { color = c; for (const b of swatches.children) b.classList.toggle("on", b === event.currentTarget); } })));
    const uline = select("下線のスタイル", [["none", "（下線なし）"], ["single", "一重線"], ["double", "二重線"], ["thick", "太線"], ["dotted", "点線"], ["dashed", "破線"], ["wavy", "波線"]], st.underline ? st.uline || "single" : "none");
    const strike = checkbox("strike", "取り消し線", st.strike && st.sline !== "double");
    const dstrike = checkbox("dstrike", "二重取り消し線", st.strike && st.sline === "double");
    const sup = checkbox("sup", "上付き", false);
    const sub = checkbox("sub", "下付き", false);
    if (!editor.typing) for (const el of [sup, sub]) { el.querySelector("input").disabled = true; el.title = "文字を入力しているときに、選んだ文字に付けます"; }
    const allCaps = checkbox("allCaps", "すべて大文字", st.caps === "all");
    const smallCaps = checkbox("smallCaps", "小型英大文字", st.caps === "small");
    const one = (a, b) => a.querySelector("input").addEventListener("change", (event) => { if (event.target.checked) b.querySelector("input").checked = false; });
    one(strike, dstrike); one(dstrike, strike); one(sup, sub); one(sub, sup); one(allCaps, smallCaps); one(smallCaps, allCaps);
    const spacing = select("文字間隔", CHAR_SPACING.map(([v, l]) => [v, l]), CHAR_SPACING.find(([v]) => Math.abs(v - (st.ls || 0)) < 0.001)?.[0] ?? 0);
    const body = h("div", { class: "fmt-cols" },
      fieldset("フォント", labeled("フォント", font), labeled("スタイル", styleSel), labeled("サイズ（pt）", size), labeled("フォントの色", swatches), labeled("下線のスタイル", uline)),
      fieldset("文字飾り", strike, dstrike, sup, sub, allCaps, smallCaps),
      fieldset("文字幅と間隔", labeled("間隔", spacing)),
      h("p", { class: "hint" }, "すべて大文字・小型英大文字は英字に効きます。"));
    settingsDialog("font-dialog", "フォント", body, () => {
      const isOn = (el) => el.querySelector("input").checked;
      // What reaches the words being typed first (bold, colour, size…), then what is set on the whole box.
      const [wantBold, wantItalic] = [/bold/.test(styleSel.value), /italic/.test(styleSel.value)];
      if (wantBold !== Boolean(st.bold)) editor.textFormat("bold");
      if (wantItalic !== Boolean(st.italic)) editor.textFormat("italic");
      const pt = parseFloat(size.value);
      if (pt > 0 && Math.abs(pt - st.fs) > 0.01) editor.textFormat("size", pt);
      if (color && color !== String(st.color || "").toLowerCase()) editor.textFormat("color", color);
      if (uline.value === "none") { if (st.underline) editor.textFormat("underline"); } else if (uline.value !== (st.underline ? st.uline || "single" : "none")) editor.textFormat("uline", uline.value);
      if (isOn(dstrike)) { if (!(st.strike && st.sline === "double")) editor.textFormat("sline", "double"); }
      else if (isOn(strike)) { if (!st.strike || st.sline === "double") editor.textFormat("sline", "single"); }
      else if (st.strike) editor.textFormat("strike");
      if (isOn(sup)) editor.textFormat("sup");
      if (isOn(sub)) editor.textFormat("sub");
      if (font.value !== (st.font || "body")) editor.textFormat("font", font.value);
      const caps = isOn(allCaps) ? "all" : isOn(smallCaps) ? "small" : null;
      if (caps !== (st.caps || null)) editor.textFormat("caps", caps);
      const ls = Number(spacing.value);
      if (Math.abs(ls - (st.ls || 0)) > 0.001) editor.textFormat("ls", ls);
    });
  }

  /** 段落: alignment, indents (before the text, the first line indented or hanging), spacing and line spacing. */
  function paragraphDialog() {
    const st = editor.textState();
    if (!st) { app.toast("文字を入れた図形・テキスト ボックスを選んでください"); return; }
    const ind = st.indent || { left: 0, first: 0 };
    const kind = ind.first < 0 ? "hanging" : ind.first > 0 ? "first" : "none";
    const before = Math.max(0, ind.left - (ind.first < 0 ? -ind.first : 0));
    const align = select("配置", [["left", "左揃え"], ["center", "中央揃え"], ["right", "右揃え"], ["justify", "両端揃え"], ["distributed", "均等割り付け"]], st.align || "left");
    const left = number("テキストの前", ops.toCm(before), { step: 0.1, min: 0, max: 20 });
    const firstKind = select("最初の行", [["none", "（なし）"], ["first", "字下げ"], ["hanging", "ぶら下げ"]], kind);
    const width = number("幅", ops.toCm(Math.abs(ind.first)), { step: 0.1, min: 0, max: 20 });
    const fsPt = st.fs || 18;
    const after = number("段落後", Math.round((st.psp || 0) * fsPt * 10) / 10, { step: 1, min: 0, max: 200 });
    const lineSel = select("行間", [...LINE_HEIGHTS.map(([v, l]) => [v, l]), ["custom", "倍数…"]], LINE_HEIGHTS.some(([v]) => Math.abs(v - (st.lh || 0)) < 0.01) ? LINE_HEIGHTS.find(([v]) => Math.abs(v - st.lh) < 0.01)[0] : "custom");
    const lineNum = number("倍数", st.lh || 1.35, { step: 0.05, min: 0.5, max: 5 });
    // タブ位置: "左 3, 右 9" (cm from the text's left) — one stop a word, at most eight.
    const TAB_WORDS = { l: "左", c: "中央", r: "右" };
    const tabsNow = (st.tabs || []).map((s) => `${TAB_WORDS[s.type]} ${ops.toCm(s.x)}`).join(", ");
    const tabsInput = h("input", { type: "text", name: "tabs", "aria-label": "タブ位置", value: tabsNow, placeholder: "例：左 3, 右 9（cm）" });
    const body = h("div", {},
      fieldset("全般", labeled("配置", align)),
      fieldset("インデント", labeled("テキストの前（cm）", left), labeled("最初の行", firstKind), labeled("幅（cm）", width)),
      fieldset("タブ", labeled("タブ位置", tabsInput), h("p", { class: "hint" }, "文字の編集中に Tab キーで入れたタブが、ここで決めた位置にそろいます（左・中央・右。入れなければ2.5cmごと）。")),
      fieldset("間隔", labeled("段落後（pt）", after), labeled("行間", lineSel), labeled("倍数", lineNum)),
      h("p", { class: "hint" }, "インデントは選んだ段落（入力中でなければ、選んだ図形のすべての段落）に付きます。"));
    settingsDialog("para-dialog", "段落", body, () => {
      if (align.value !== (st.align || "left")) editor.textFormat("align", align.value);
      const w = ops.fromCm(Math.max(0, Number(width.value) || 0));
      const b = ops.fromCm(Math.max(0, Number(left.value) || 0));
      const next = firstKind.value === "hanging" ? { left: Math.round((b + w) * 100) / 100, first: -w } : firstKind.value === "first" ? { left: b, first: w } : { left: b, first: 0 };
      if (Math.abs(next.left - ind.left) > 0.5 || Math.abs(next.first - ind.first) > 0.5) editor.textFormat("paraIndent", next);
      // The tab stops as typed: "左 3, 右 9" → l/r at those cm (a bare number is a left stop).
      const stops = tabsInput.value.split(/[,、，]/).map((part) => {
        const m = /^\s*(左|中央|右|l|c|r)?\s*([\d.]+)/i.exec(part.replace(/[０-９．]/g, (c) => (c === "．" ? "." : String.fromCharCode(c.charCodeAt(0) - 0xfee0))));
        if (!m || !(Number(m[2]) > 0)) return null;
        const type = { 左: "l", 中央: "c", 右: "r" }[m[1]] || (m[1] || "l").toLowerCase();
        return { type, x: Math.round(ops.fromCm(Number(m[2])) * 100) / 100 };
      }).filter(Boolean);
      const nextTabs = E.parseTabs(E.tabsText(stops));
      if (E.tabsText(nextTabs) !== E.tabsText(st.tabs || [])) editor.textFormat("paraTabs", nextTabs);
      const lh = lineSel.value === "custom" ? Number(lineNum.value) : Number(lineSel.value);
      if (lh > 0 && Math.abs(lh - (st.lh || 0)) > 0.001) editor.textFormat("lh", Math.min(5, Math.max(0.5, lh)));
      const psp = Math.round(((Number(after.value) || 0) / fsPt) * 100) / 100;
      if (Math.abs(psp - (st.psp || 0)) > 0.005) editor.textFormat("psp", Math.min(4, psp));
    });
  }

  /** グリッドとガイド: snapping to the grid and its spacing, showing the grid, the guides and the smart guides. */
  function gridDialog() {
    const st = editor.state;
    const step = select("間隔", GRID_STEPS.map((v) => [v, `${v} cm`]), st.gridStep);
    const body = h("div", {},
      fieldset("オブジェクトの位置合わせ", checkbox("snapGrid", "描画オブジェクトをグリッド線に合わせる", st.snapGrid)),
      fieldset("グリッドの設定", labeled("間隔", step), checkbox("grid", "グリッドを表示", st.grid)),
      fieldset("ガイドの設定", checkbox("guides", "ガイドを表示", st.guides), checkbox("smart", "図形の整列時にスマート ガイドを表示する", st.smart)));
    const dialog = settingsDialog("grid-dialog", "グリッドとガイド", body, () => {
      const on = (name) => dialog.querySelector(`input[name="${name}"]`).checked;
      const gridStep = Number(step.value);
      if (gridStep !== st.gridStep) editor.setView("gridStep", gridStep);
      for (const key of ["snapGrid", "grid", "guides", "smart"]) if (on(key) !== Boolean(st[key])) editor.setView(key, on(key));
    });
  }

  /** デザイン → スライドのサイズ: the slide's size (the template's wide screen) and the number the first slide carries. */
  function slideSizeDialog() {
    const first = number("firstNumber", app.firstNumber(), { step: 1, min: 0, max: 9999 });
    const body = h("div", {},
      fieldset("スライドのサイズ", h("p", { class: "hint" }, "ワイド画面（16:9・13.33×7.5インチ）です。PowerPoint の標準の大きさで書き出します。")),
      fieldset("スライド番号", labeled("開始番号", first), h("p", { class: "hint" }, "最初のスライドにつける番号です（表紙を0にして、次を1から数える、など）。ページ番号に出ます。")));
    settingsDialog("size-dialog", "スライドのサイズ", body, () => app.setFirstNumber(first.value));
  }

  function ribbonOptions(anchor) {
    openPop(anchor, menu([
      { head: "リボンの表示オプション" },
      ...Object.entries(RIBBON_MODES).map(([mode, label]) => ({ label, icon: mode === "full" ? "ribbon" : mode === "tabs" ? "chevron" : "zoomFit", on: ribbonMode === mode, keys: mode === "tabs" ? "⌘F1" : "", run: () => setRibbonMode(mode) })),
      ...(quick?.optionItems() || []),
    ]));
  }

  function renderRibbon(force = false) {
    if (!ribbon) return;
    // The ribbon is there in every view of the deck, as in PowerPoint; off the 1枚 view only what works on whole
    // slides (new slides, design, slide show, review, view) can be used.
    const visible = Boolean(app.deck());
    ribbon.hidden = !visible;
    if (!visible) { built = ""; floating = false; return; }
    const tabs = TABS.filter((t) => !t.contextual || t.contextual());
    if (!tabs.some((t) => t.id === tab)) tab = "home";
    const sig = signature();
    if (!force && sig === built) { refresh(); return; }
    built = sig;
    updaters = [];
    ribbon.dataset.mode = ribbonMode;
    ribbon.classList.toggle("rb-floating", floating);
    document.removeEventListener("pointerdown", floatingOutside, true);
    document.removeEventListener("keydown", floatingKey, true);
    if (floating) { document.addEventListener("pointerdown", floatingOutside, true); document.addEventListener("keydown", floatingKey, true); }
    // PowerPoint's ファイル: new, open, save, save as, export, history (a menu at the start of the tabs).
    const fileBtn = h("button", { type: "button", class: "rb-file", title: "ファイル：新規・開く・保存・書き出し・履歴", "aria-haspopup": "true", onclick: (event) => openPop(event.currentTarget, fileMenu()) }, "ファイル");
    const qatBar = quick?.toolbar();
    const head = h("div", { class: "rb-tabs", role: "tablist" }, quick?.where() === "above" ? qatBar : null, fileBtn,
      // Switching tabs keeps the caret (and the cells picked in a table), as in PowerPoint.
      tabs.map((t) => h("button", { type: "button", role: "tab", class: t.contextual ? "contextual" : "", "aria-selected": String(t.id === tab && (ribbonMode === "full" || floating)), "data-tab": t.id, "data-keeps-text": "", onmousedown: (event) => event.preventDefault(),
        onclick: () => {
          if (ribbonMode === "full") { tab = t.id; renderRibbon(true); editor.draw(); }
          else if (floating && tab === t.id) closeFloating();
          else openFloating(t.id);
        },
        ondblclick: () => toggleRibbon() }, t.label)),
      h("span", { class: "rb-spacer" }),
      h("span", { class: "rb-hint" }, ""),
      quick?.searchBox,
      h("button", { type: "button", class: "rb-options", title: "リボンの表示オプション（常に表示・タブのみ・自動的に非表示）", "aria-haspopup": "true", onclick: (event) => ribbonOptions(event.currentTarget) }, ico("ribbon", 15), h("span", {}, "表示オプション"), h("b", { class: "rb-caret" }, "▾")));
    const showBody = ribbonMode === "full" || floating;
    const pin = h("div", { class: "rb-end" }, floating
      ? h("button", { type: "button", class: "rb-pin", title: "リボンを固定する（常に表示）", "aria-label": "リボンを固定する", onclick: () => setRibbonMode("full") }, ico("pin", 15))
      : h("button", { type: "button", class: "rb-pin", title: "リボンを折りたたむ（⌘F1）：タブだけを表示します", "aria-label": "リボンを折りたたむ", onclick: () => setRibbonMode("tabs") }, "˄"));
    const body = showBody ? h("div", { class: "rb-body", role: "tabpanel" }, ...buildTab(tab), pin) : null;
    if (body && !onSlide()) {
      for (const g of body.querySelectorAll(":scope > .rb-group")) {
        const label = g.querySelector(":scope > .rb-label")?.textContent || "";
        if (OFF_SLIDE_TABS.has(tab) || OFF_SLIDE_GROUPS.has(label)) continue;
        g.classList.add("rb-off");
        g.inert = true;
        g.title = "1枚表示（標準）で使えます";
      }
    }
    // 最終版: only the tabs for looking and presenting work (編集する on the message bar lets the rest back).
    if (body && app.isFinal?.() && !FINAL_TABS.has(tab)) {
      for (const g of body.querySelectorAll(":scope > .rb-group")) {
        g.classList.add("rb-off");
        g.inert = true;
        g.title = "最終版です（メッセージ バーの「編集する」で編集できます）";
      }
    }
    if (pop && !pop.anchor.isConnected) closePop(true);
    const reveal = h("button", { type: "button", class: "rb-reveal", title: "リボンを表示する（自動的に非表示）", "aria-label": "リボンを表示する", onclick: () => (floating ? closeFloating() : openFloating()) }, "•••");
    const below = quick?.where() === "below" ? h("div", { class: "rb-qat-row" }, qatBar) : "";
    if (ribbonMode === "auto") ribbon.replaceChildren(reveal, floating ? h("div", { class: "rb-float" }, head, body, below) : "");
    else ribbon.replaceChildren(head, body || "", below);
    if (pop && !pop.anchor.isConnected) closePop(true);
    refresh();
    fitRibbon();
    relight();
  }
  // Off the 1枚 view: whole tabs and groups that work on slides rather than on the objects of the one on the stage.
  const OFF_SLIDE_TABS = new Set(["design", "transition", "slideshow", "review", "view", "developer"]);
  const OFF_SLIDE_GROUPS = new Set(["スライド"]);
  const FINAL_TABS = new Set(["slideshow", "view"]);
  // A narrow window squeezes the ribbon as PowerPoint does, from the right-hand groups first: small buttons
  // lose their words (the tooltip keeps them), then whole groups fold into one button that opens them;
  // past that the groups draw closer, and last of all the ribbon scrolls sideways.
  const GROUP_ICONS = { グラフィックの作成: "smartart", "SmartArt のスタイル": "fill", 文字: "font", リセット: "reset", プレビュー: "play", ブックマーク: "pin", グラフのレイアウト: "layout", オプション: "audio", "オーディオ スタイル": "audio", スライド: "slide", クリップボード: "paste", フォント: "font", 段落: "textLeft", 図形描画: "shapes", 編集: "selectAll", 画像: "image", 図: "shapes", テキスト: "textbox", 線: "line", メディア: "video", リンク: "link",
    図形の挿入: "shapes", 図形のスタイル: "style", ワードアートのスタイル: "fontColor", 配置: "front", サイズ: "zoomFit", 調整: "bright", 図のスタイル: "outline", 表示: "slide", "表示/非表示": "grid", ズーム: "zoomIn", ウィンドウ: "pane",
    プレゼンテーションの表示: "slide", "カラー/グレースケール": "grayView", 記号と日付: "symbol", コメント: "comment", デザイナー: "magic", コード: "effectOpts", 資料のデータ: "down", 書き出し: "play", テーマ: "theme", "動き（資料全体）": "motionPath", ユーザー設定: "slide", "スライド ショーの開始": "showStart", 設定: "presenter", チェック: "check", 検索: "find", ノート: "notes", AI: "magic", レビュー: "review", 描画ツール: "pen", 変換: "inkShape", 再生: "play", 音声: "mic", ズームのオプション: "zoomSlide", ズームのスタイル: "outline", 数式と記号: "equation", モニター: "presenter", キャプションと字幕: "subtitles", 文章校正: "spell", キャプション: "caption", 言語: "textCase", アクセシビリティ: "accessibility", 比較: "thumbs", インク: "pen" };
  function fold(g) {
    const items = g.querySelector(":scope > .rb-items");
    const label = g.querySelector(":scope > .rb-label")?.textContent || "";
    // Opening a folded group keeps the caret and the words picked while typing (its リンク, フォント… act on them).
    const button = h("button", { type: "button", class: "rb-btn big rb-drop rb-folded", title: label, "aria-haspopup": "true", "data-keeps-text": "",
      onmousedown: (event) => event.preventDefault(),
      onclick: (event) => { event.preventDefault(); openPop(button, () => h("div", { class: "rb-fold" }, items, h("div", { class: "rb-label" }, label))); } },
    ico(GROUP_ICONS[label] || "chevron", 26), caption(label), h("b", { class: "rb-caret" }, "▾"));
    g.classList.add("folded");
    g.prepend(h("div", { class: "rb-items rb-fold-slot" }, button));
    g.foldedItems = items;
    items.remove();
  }
  function unfold(g) {
    if (pop?.el.contains(g.foldedItems)) closePop();
    g.querySelector(":scope > .rb-fold-slot")?.replaceWith(g.foldedItems);
    g.classList.remove("folded");
    delete g.foldedItems;
  }
  function fitRibbon() {
    const body = ribbon?.querySelector(".rb-body");
    if (!body || !body.clientWidth) return;
    const groups = [...body.querySelectorAll(":scope > .rb-group")];
    for (const g of groups) { if (g.foldedItems) unfold(g); g.classList.remove("compact"); }
    ribbon.classList.remove("rb-tight");
    const over = () => body.scrollWidth > body.clientWidth + 1;
    for (let i = groups.length - 1; i >= 0 && over(); i -= 1) groups[i].classList.add("compact");
    for (let i = groups.length - 1; i >= 1 && over(); i -= 1) fold(groups[i]);
    // Folding may have made room: the groups on the left get their words back where they fit.
    for (const g of groups) {
      if (g.foldedItems || !g.classList.contains("compact")) continue;
      g.classList.remove("compact");
      if (over()) { g.classList.add("compact"); break; }
    }
    if (over()) ribbon.classList.add("rb-tight");
  }
  if (ribbon && "ResizeObserver" in window) {
    let width = 0;
    new ResizeObserver(([entry]) => { const w = Math.round(entry.contentRect.width); if (w !== width) { width = w; fitRibbon(); } }).observe(ribbon);
  }
  function refresh() {
    for (const fn of updaters) { try { fn(); } catch { /* a control without a target */ } }
    const hint = ribbon?.querySelector(".rb-hint");
    if (hint) hint.textContent = ink?.tool ? `${{ pen: "ペンで書いています", eraser: "消しゴム：消したい線をなぞる", lasso: "投げ縄：インクを囲む" }[ink.tool]}（Escでやめる）` : editor.tool ? "描画中：スライド上をドラッグ（Escでやめる）" : editor.painter ? "書式を貼り付ける図形をクリック（Escでやめる）" : extras?.dictating ? "音声入力中：話すと文字になります" : "";
  }

  function fontControls() {
    const size = h("input", { class: "rb-size", type: "text", inputmode: "decimal", list: "rbSizes", title: "フォントサイズ（pt）", "aria-label": "フォントサイズ", "data-keeps-text": "",
      onchange: (event) => { const v = parseFloat(event.target.value); if (v > 0) editor.textFormat("size", v); },
      onkeydown: (event) => { if (event.key === "Enter") { event.preventDefault(); event.target.blur(); } } });
    updaters.push(() => { const st = editor.textState(); size.disabled = !st; if (document.activeElement !== size) size.value = st ? String(st.fs) : ""; });
    const font = h("select", { class: "rb-font", title: "フォント", "aria-label": "フォント", onchange: (event) => editor.textFormat("font", event.target.value) }, Object.entries(E.FONTS).map(([key, [label]]) => h("option", { value: key }, label)));
    updaters.push(() => { const st = editor.textState(); font.disabled = !st; font.value = st?.font || "body"; });
    return [font, size, h("datalist", { id: "rbSizes" }, SIZES_PT.map((v) => h("option", { value: String(v) })))];
  }
  const tState = (key) => () => editor.textState()?.[key];

  function homeTab() {
    return [
      group("スライド",
        drop("slide", "新しい|スライド", "スライドを追加", menu([
          { label: "白紙（自由配置）", icon: "slide", run: () => app.insertBlankSlide() },
          { label: "レイアウトを選んで追加…", icon: "layout", run: () => app.openTypeDialog("insert") },
          "-",
          { label: "スライドの再利用…", icon: "duplicateSlide", run: () => app.openReuse() },
          { label: "アウトラインからスライド…", icon: "textLeft", run: () => app.openOutlineSlides() },
        ]), { big: true }),
        col(btn("layout", "レイアウト", "このスライドのレイアウトを変える", () => app.openTypeDialog("change")), btn("duplicateSlide", "複製", "このスライドを複製", () => app.duplicateSlide()),
          btn("convert", "図形に変換", "このスライドのレイアウトを図形・テキストボックス・画像に分けて、1つずつ自由に編集できるようにする（白紙のスライドになります）", () => app.convertSlide(), { enabled: () => app.canConvert() })),
        col(btn("thumbs", "セクション", "このスライドからセクションを始める（サムネイルでセクション名を右クリックすると、名前の変更・削除・移動）", () => app.addSection()))),
      group("クリップボード",
        drop("paste", "貼り付け", "貼り付け（⌘V）と貼り付けオプション", () => menu([
          { head: "貼り付けのオプション" },
          { label: "元の書式を保持", icon: "paste", keys: "⌘V", run: () => editor.pasteFromMemory("keep") },
          { label: "貼り付け先のスタイルを使用（テーマの色に合わせる）", icon: "fill", run: () => editor.pasteFromMemory("theme") },
          { label: "テキストのみ保持", icon: "textbox", keys: "⇧⌘V", run: () => editor.pasteFromMemory("text") },
          { label: "図（1枚の画像として貼り付け）", icon: "image", run: () => editor.pasteFromMemory("picture") },
        ]), { big: true }),
        col(btn("cut", "切り取り", "切り取り（⌘X）", () => { document.execCommand("cut") || cutFallback(); }, { enabled: any }),
          btn("copy", "コピー", "コピー（⌘C）", () => { document.execCommand("copy") || copyFallback(); }, { enabled: any }),
          painterBtn())),
      group("フォント",
        row(...fontControls()),
        row(btn("bold", "", "太字（⌘B）", () => editor.textFormat("bold"), { enabled: hasText, pressed: tState("bold"), keep: true }),
          btn("italic", "", "斜体（⌘I）", () => editor.textFormat("italic"), { enabled: hasText, pressed: tState("italic"), keep: true }),
          btn("underline", "", "下線（⌘U）", () => editor.textFormat("underline"), { enabled: hasText, pressed: tState("underline"), keep: true }),
          caret("下線の種類・二重取り消し線", () => menu([
            { head: "下線の種類" },
            ...[["single", "一重線"], ["double", "二重線"], ["thick", "太線"], ["dotted", "点線"], ["dashed", "破線"], ["wavy", "波線"]].map(([v, label]) => ({ label, on: editor.textState()?.uline === v, run: () => editor.textFormat("uline", v) })),
            "-", { head: "取り消し線" },
            { label: "一重取り消し線", on: editor.textState()?.sline === "single", run: () => editor.textFormat("sline", "single") },
            { label: "二重取り消し線", on: editor.textState()?.sline === "double", run: () => editor.textFormat("sline", "double") },
          ]), { enabled: hasText }),
          btn("strike", "", "取り消し線", () => editor.textFormat("strike"), { enabled: hasText, pressed: tState("strike"), keep: true }),
          btn("textShadow", "", "文字の影", () => setTextFx({ tshadow: textFxOf("tshadow") ? undefined : ops.textShadowPreset("br") }), { enabled: hasText, pressed: () => Boolean(textFxOf("tshadow")), keep: true }),
          btn("sup", "", "上付き", () => editor.textFormat("sup"), { enabled: () => editor.typing, keep: true }),
          btn("sub", "", "下付き", () => editor.textFormat("sub"), { enabled: () => editor.typing, keep: true }),
          btn("grow", "", "フォントサイズの拡大（⇧⌘>）", () => editor.textFormat("grow"), { enabled: hasText, keep: true }),
          btn("shrink", "", "フォントサイズの縮小（⇧⌘<）", () => editor.textFormat("shrink"), { enabled: hasText, keep: true }),
          drop("fontColor", "", "フォントの色", () => colors(E.PALETTE.text, textColorOf(), (c) => editor.textFormat("color", c)), { enabled: hasText, keep: true, swatch: textColorOf }),
          drop("highlight", "", "蛍光ペン（マーカー）", () => colors(E.PALETTE.highlight, "", (c) => editor.textFormat("highlight", c === "none" ? null : c), { none: "マーカーなし", custom: false }), { enabled: hasText, keep: true }),
          drop("spacing", "", "文字の間隔", () => menu(CHAR_SPACING.map(([v, label]) => ({ label, on: Math.abs((one()?.ls || 0) - v) < 0.001, run: () => editor.textFormat("ls", v) }))), { enabled: hasText, keep: true }),
          drop("textCase", "", "文字種の変換（大文字・小文字・全角・半角）", () => extras.caseMenu(), { enabled: hasText, keep: true }),
          btn("clear", "", "書式のクリア", () => editor.textFormat("clear"), { enabled: hasText, keep: true }),
          btn("font", "", "フォントの設定（フォント・スタイル・サイズ・色・下線・文字飾り・文字間隔）。⇧⌘F", () => fontDialog(), { enabled: hasText, keep: true }))),
      group("段落",
        row(btn("bullet", "", "箇条書き", () => editor.textFormat("bullet"), { enabled: hasText, pressed: () => editor.textState()?.list === "bullet", keep: true }),
          caret("箇条書きの種類（●・■・◆・➢・✓…）", listGallery("bullet"), { enabled: hasText }),
          btn("number", "", "段落番号", () => editor.textFormat("number"), { enabled: hasText, pressed: () => editor.textState()?.list === "number", keep: true }),
          caret("段落番号の種類（1.・1)・①・a.・i.・一、…）", listGallery("number"), { enabled: hasText }),
          btn("indentLess", "", "インデントを減らす", () => editor.textFormat("outdent"), { enabled: () => editor.typing, keep: true }),
          btn("indentMore", "", "インデントを増やす", () => editor.textFormat("indent"), { enabled: () => editor.typing, keep: true }),
          drop("lineSpacing", "", "行間", () => menu(LINE_HEIGHTS.map(([v, label]) => ({ label, on: Math.abs((editor.textState()?.lh ?? 0) - v) < 0.01, run: () => editor.textFormat("lh", v) }))), { enabled: hasText, keep: true })),
        row(btn("textLeft", "", "左揃え（⌘L）", () => editor.textFormat("align", "left"), { enabled: hasText, pressed: () => editor.textState()?.align === "left", keep: true }),
          btn("textCenter", "", "中央揃え（⌘E）", () => editor.textFormat("align", "center"), { enabled: hasText, pressed: () => editor.textState()?.align === "center", keep: true }),
          btn("textRight", "", "右揃え（⌘R）", () => editor.textFormat("align", "right"), { enabled: hasText, pressed: () => editor.textState()?.align === "right", keep: true }),
          btn("textJustify", "", "両端揃え（⌘J）", () => editor.textFormat("align", "justify"), { enabled: hasText, pressed: () => editor.textState()?.align === "justify", keep: true }),
          btn("textDistributed", "", "均等割り付け（⇧⌘J）", () => editor.textFormat("align", "distributed"), { enabled: hasText, pressed: () => editor.textState()?.align === "distributed", keep: true }),
          btn("indentMore", "", "段落の設定（配置・インデント（字下げ・ぶら下げ）・段落の間隔・行間）", () => paragraphDialog(), { enabled: hasText, keep: true }),
          btn("vtext", "", "縦書き", () => editor.textFormat("vertical"), { enabled: hasText, pressed: tState("vertical") }),
          caret("文字列の方向（横書き・縦書き・90度／270度回転）", () => menu([
            { head: "文字列の方向" },
            { label: "横書き", icon: "textbox", on: !tState("vertical")() && !tState("textRot")(), run: () => editor.textFormat("direction", "h") },
            { label: "縦書き", icon: "vtext", on: Boolean(tState("vertical")()), run: () => editor.textFormat("direction", "v") },
            { label: "すべてのテキストを 90 度回転", icon: "rotate", on: tState("textRot")() === 90, run: () => editor.textFormat("direction", 90) },
            { label: "すべてのテキストを 270 度回転", icon: "rotate", on: tState("textRot")() === 270, run: () => editor.textFormat("direction", 270) },
          ]), { enabled: hasText }),
          drop("valignMiddle", "", "文字の配置（上下）", () => menu([["top", "上揃え", "valignTop"], ["middle", "上下中央揃え", "valignMiddle"], ["bottom", "下揃え", "valignBottom"]].map(([v, label, icon]) => ({ label, icon, on: editor.textState()?.valign === v, run: () => editor.textFormat("valign", v) }))), { enabled: hasText }),
          drop("columns", "", "段組み：テキストを2段・3段に分ける", () => extras.columnsMenu(), { enabled: hasText }),
          drop("smartart", "", "SmartArt に変換：選んだテキストの段落（箇条書きのレベル）を図に", () => smart.convertGallery(), { enabled: () => selected().some((o) => ["shape", "text"].includes(o.kind) && o.text) }))),
      group("図形描画",
        drop("shapes", "図形", "図形を描く", () => shapeGallery(pickTool), { big: true }),
        col(drop("front", "配置", "前面・背面・整列・グループ・回転", () => arrangeMenu(), { enabled: any }),
          drop("style", "スタイル", "図形のクイックスタイル（テーマの色）", () => styleMenu(), { enabled: hasText })),
        col(drop("fill", "塗り", "図形の塗りつぶし", () => colors(E.PALETTE.fill, fillOf(), setFill, { none: "塗りつぶしなし", extra: fillExtras }), { enabled: hasText, swatch: fillOf }),
          drop("outline", "枠線", "図形の枠線", () => outlineMenu(), { enabled: () => any() && !kinds().has("icon") || kinds().has("icon"), swatch: strokeOf }))),
      group("編集",
        col(btn("find", "検索・置換", "資料全体の文字を検索・置換（⌘F）", () => app.openReplace()),
          btn("selectAll", "すべて選択", "このスライドのオブジェクトをすべて選択（⌘A）", () => editor.select(editor.objects().filter((o) => !o.hidden).map((o) => o.id))),
          btn("pane", "選択ウィンドウ", "オブジェクトの一覧・表示/非表示・順番", () => app.openPanel("format", "selection")))),
      group("音声",
        btn("mic", "ディク|テーション", "音声入力：話した言葉をカーソルの位置に入力（もう一度押すと終わります。Chrome・Edge）", () => extras.toggleDictation(), { big: true, keep: true, pressed: () => extras.dictating })),
    ];
  }

  /** テーマ (デザイン → テーマ): each theme drawn on the deck's own cover, as PowerPoint's gallery; a click switches. */
  function themeMenu() {
    return (close) => h("div", { class: "rb-themes" }, E.THEMES.map((theme) => h("button", {
      type: "button", class: ["rb-theme", app.deckTheme() === theme.id ? "on" : ""], title: `${theme.name}：${theme.desc}`, "data-theme-id": theme.id,
      onmousedown: (e) => e.preventDefault(), onclick: () => { close(); app.setDeckDesign({ theme: theme.id }); },
    }, app.themePreview(theme.id), h("b", {}, theme.name), h("small", {}, theme.desc))));
  }
  /** デザイン: the theme (colours and fonts, PowerPoint's テーマ) and the deck's motion (as the 動き dialog, one menu each). */
  function designTab() {
    const motion = () => app.deckMotion() || {};
    const choose = (head, entries, current, apply) => () => menu([{ head }, ...entries.map(([value, label]) => ({ label, on: current() === value, run: () => apply(value) }))]);
    const setMotion = (key) => (value) => app.setDeckDesign({ motion: { [key]: value } });
    return [
      group("テーマ",
        drop("theme", "テーマ", "テーマ：配色とフォント（PowerPoint のテーマとして書き出します。テーマの色で塗った部品は一緒に変わります）", () => themeMenu(), { big: true })),
      group("動き（資料全体）",
        btn("motionPath", "動きの|設定", "切り替え・登場・背景の動き・文字の動き・強調を見比べて選ぶ（資料全体）", () => app.openDesign(), { big: true }),
        col(drop("transition", "切り替え", "スライドの切り替え（資料全体。1枚ずつは「画面切り替え」タブ）", choose("スライドの切り替え", Object.entries(E.TRANSITIONS), () => motion().transition, (value) => app.setDeckDesign({ transition: value }))),
          drop("previewPlay", "登場のしかた", "スライドの見出しと中身が現れるとき", choose("登場のしかた", [...Object.entries(E.ENTRANCES), ["none", "動かさない"]], () => motion().entrance, setMotion("entrance"))),
          drop("hover", "マウスを乗せたとき", "発表中、項目にマウスを乗せたとき", choose("マウスを乗せたとき", [...Object.entries(E.HOVERS), ["none", "変化なし"]], () => motion().hover, setMotion("hover")))),
        col(drop("loop", "背景の動き", "表紙・章扉・ひと言・最後のスライドの背景", choose("背景の動き", [["none", "テンプレートの波紋だけ"], ...Object.entries(E.BACKDROPS)], () => motion().backdrop, setMotion("backdrop"))),
          drop("textbox", "文字の動き", "表紙・章扉などの大きな文字", choose("大きな文字の動き", [["none", "動かさない"], ...Object.entries(E.KINETIC)], () => motion().kinetic, setMotion("kinetic"))),
          drop("highlight", "強調の見せ方", "**語句** で囲んだ語句", choose("強調の見せ方", Object.entries(E.EMPHASES), () => motion().emphasis, setMotion("emphasis"))))),
      group("デザイナー",
        btn("magic", "デザイン|アイデア", "このスライドの見せ方の違う3つの案をAIが作り、並べて選べます（Codexに接続しているとき。右の「AIと話す」に届きます）", () => app.designIdeas(), { big: true, enabled: () => app.canAi() })),
      group("ユーザー設定",
        btn("slide", "スライドの|サイズ", "スライドのサイズ（ワイド画面 16:9）とスライド番号の開始番号", () => slideSizeDialog(), { big: true }),
        btn("hideSlide", "背景グラフィック|を表示しない", "スライド マスターに置いた図形・文字・図を、このスライドだけ隠す（スライド マスター表示で置けます）", () => app.setSlideFields({ hideMaster: app.slide()?.hideMaster ? null : true }), { big: true, pressed: () => Boolean(app.slide()?.hideMaster), enabled: () => Boolean(app.slide()) && !app.masterView() }),
        btn("background", "背景の|書式設定", "このスライドの背景（色・図と透明度・並べて表示）。すべてに適用・背景のリセットも", () => app.formatBackground(), { big: true, enabled: () => Boolean(app.slide()) }),
        col(btn("chartBar", "数字を数え上げる", "数字のカウントアップとグラフが伸びる動き（資料全体）", () => app.setDeckDesign({ motion: { numbers: !motion().numbers } }), { pressed: () => motion().numbers }),
          btn("spot", "背景をゆっくり動かす", "テーマの背景の飾りを発表中にゆっくり動かす", () => app.setDeckDesign({ motion: { ambient: !motion().ambient } }), { pressed: () => motion().ambient }),
          btn("pen", "線を描くように", "アイコン・線・マーカーを描くように見せる", () => app.setDeckDesign({ motion: { draw: !motion().draw } }), { pressed: () => motion().draw }))),
    ];
  }

  /**
   * スライド ショー: start from the beginning, from here or a custom show; the show's settings (a kiosk, looping, a
   * range, timings, no animation), hiding a slide, rehearsing, the presenter view and live subtitles.
   */
  function slideshowTab() {
    const show = () => app.showSettings();
    const customMenu = () => menu([
      { head: "目的別スライド ショー" },
      ...app.customShows().map((cs) => ({ label: `${cs.name}（${cs.sids.length}枚）`, icon: "showStart", run: () => app.presentCustom(cs.id) })),
      app.customShows().length ? "-" : null,
      { label: "目的別スライド ショー…", icon: "customShow", run: () => app.openCustomShows() },
    ]);
    return [
      group("スライド ショーの開始",
        btn("showStart", "最初から", "最初のスライドから発表する（F5）", () => app.presentFrom(0), { big: true }),
        btn("showHere", "このスライド|から", "このスライドから発表する（⇧F5）", () => app.presentFrom(app.index()), { big: true }),
        drop("customShow", "目的別|スライド ショー", "相手や時間に合わせて選んだスライドだけを、決めた順番で発表する", customMenu, { big: true }),
        btn("share", "オンライン|プレゼンテーション", "リンクを送った相手が、ブラウザで発表を見られます（視聴者の画面は発表者に合わせて進みます）", () => app.openOnline(), { big: true, pressed: () => app.onlineActive() })),
      group("設定",
        btn("presenter", "スライド ショー|の設定", "種類（自動プレゼンテーション）・繰り返し・発表するスライド（範囲・目的別）・タイミング・アニメーションなし・ペンの色", () => app.openShowSettings(), { big: true }),
        btn("hideSlide", "非表示スライド|に設定", "発表ではこのスライドを飛ばす（編集用に残ります。もう一度で戻す）", () => app.toggleHiddenSlide(), { big: true, pressed: () => Boolean(app.slide()?.hidden), enabled: () => app.index() > 0 }),
        btn("clock", "リハーサル", "最初から発表して1枚ずつの時間を計り、終わったらその時間で自動的に切り替えるようにできます", () => app.rehearse(), { big: true }),
        btn("mic", "コーチによる|リハーサル", "話す練習：マイクで聞き取り、ペース・つなぎ言葉・繰り返し・スライドの読み上げをその場とレポートで知らせます（Chrome・Edge）", () => app.rehearseWithCoach(), { big: true }),
        drop("record", "記録", "スライド ショーの記録：発表しながらマイクでナレーションを録り、スライドごとのナレーションとタイミングにする", () => menu([
          { head: "スライド ショーの記録" },
          { label: "先頭から記録", icon: "record", run: () => app.recordShow(0) },
          { label: "現在のスライドから記録", icon: "record", run: () => app.recordShow(app.index()) },
          "-", { head: "クリア" },
          { label: "現在のスライドのタイミングをクリア", icon: "clock", run: () => app.clearRecording("timings", "slide") },
          { label: "すべてのスライドのタイミングをクリア", icon: "clock", run: () => app.clearRecording("timings", "all") },
          { label: "現在のスライドのナレーションをクリア", icon: "audio", run: () => app.clearRecording("narration", "slide") },
          { label: "すべてのスライドのナレーションをクリア", icon: "audio", run: () => app.clearRecording("narration", "all") },
        ]), { big: true }),
        col(btn("check", "タイミングを使用", "リハーサル・自動で切り替えの時間で進める（オフにするとクリックで進む）", () => app.setShowSettings({ useTimings: show().useTimings === false }), { pressed: () => show().useTimings !== false }),
          btn("audio", "ナレーションの再生", "記録したナレーションを発表で流す", () => app.setShowSettings({ noNarration: !show().noNarration }), { pressed: () => !show().noNarration }),
          btn("previewPlay", "動きを確認", "閲覧表示：このスライドの動きをこの画面で再生する", () => app.previewMotion()))),
      group("モニター",
        col(btn("presenter", "発表者ツールを使用", "発表を始めると、ノート・次のスライド・経過時間の発表者ビューを別ウィンドウで開く（発表中は P）", () => app.setPresenterView(!app.presenterView()), { pressed: () => app.presenterView() }),
          btn("pen", "ペンの色", "発表中のペン（Ctrl+P）の色：スライド ショーの設定で変えられます", () => app.openShowSettings()),
          btn("clock", "自動で切り替え", "このスライドを何秒で次へ進めるか（画面切り替えタブ）", () => showTab("transition", { open: true })))),
      group("キャプションと字幕",
        btn("subtitles", "常に字幕を|使用", "発表を始めると、話した言葉を字幕で出す（発表中は J で切り替え。マイクを使います）", () => app.setShowSettings({ captions: !show().captions }), { big: true, pressed: () => Boolean(show().captions) }),
        drop("caption", "字幕の|設定", "字幕の言語", () => menu([{ head: "話す言語" }, ...[["ja-JP", "日本語"], ["en-US", "英語"], ["zh-CN", "中国語"], ["ko-KR", "韓国語"]].map(([k, label]) => ({ label, on: (show().captionLang || "ja-JP") === k, run: () => app.setShowSettings({ captionLang: k }) }))]), { big: true })),
    ];
  }

  /** 校閲: the checks, find and replace, notes, the AI's review, and review files to and from others. */
  function reviewTab() {
    return [
      group("コメント",
        btn("comment", "新しい|コメント", "このスライド（図形を選んでいればその図形）にコメントを付ける", () => app.newComment(), { big: true }),
        drop("trash", "削除", "コメントを削除", () => menu([
          { label: "このスライドのコメントをすべて削除", icon: "trash", run: () => app.clearComments("slide") },
          { label: "資料のすべてのコメントを削除", icon: "trash", run: () => app.clearComments("deck") },
        ]), { big: true }),
        col(btn("up", "前へ", "前のコメントのあるスライドへ", () => app.commentGo(-1)),
          btn("down", "次へ", "次のコメントのあるスライドへ", () => app.commentGo(1)),
          drop("eye", "コメントの表示", "スライド上のコメントとコメント ウィンドウ", () => menu([
            { label: "スライド上にコメントを表示", icon: "comment", on: app.commentPins(), run: () => app.setCommentPins(!app.commentPins()) },
            { label: "コメント ウィンドウ", icon: "taskPane", on: app.panel() === "comment", run: () => app.showPanel("comment") },
          ]))),
        btn("people", "ユーザー名", "コメントと共同編集で表示するあなたの名前", () => app.changeUserName(), { big: true })),
      group("チェック",
        btn("check", "チェック", "構成・文字のあふれを確かめる", () => app.openCheck(), { big: true }),
        btn("accessibility", "アクセシビリティ|チェック", "代替テキスト・スライド タイトル・表の見出し・リンクの文字・コントラスト・読み取り順序を確かめる（作業ウィンドウに一覧）", () => app.openAccessibility(), { big: true }),
        btn("textbox", "代替|テキスト", "選んだ部品の代替テキスト（画面読み上げが読む説明）・装飾用", () => { const o = one(); if (o) app.editAlt(o.id); else app.toast("部品を1つ選んでください"); }, { big: true, enabled: () => Boolean(one()) }),
        btn("magic", "あふれを|AIで直す", "文字が収まらないスライドをAIで順番に直す", () => app.fixOverflow(), { big: true, enabled: () => app.canFixOverflow() })),
      group("文章校正",
        btn("spell", "スペル|チェック", "入力中の文字のスペルを確かめる（ブラウザの辞書。間違いに赤い波線）", () => app.setSpellcheck(!app.spellcheck()), { big: true, pressed: () => app.spellcheck() }),
        btn("count", "文字|カウント", "資料全体のスライド数・段落・文字数（スペースあり・なし）・英数字の単語を数える（スピーカー ノートや非表示スライドを含めるかも選べます）", () => app.openWordCount(), { big: true }),
        btn("thesaurus", "類義語", "選んだ言葉の言い換えを探して入れ替える（⇧F7。辞書にない言葉はAIで探せます）", () => app.openSynonyms(), { big: true }),
        btn("textCase", "表記ゆれ|チェック", "全角・半角（ＡＩ／AI）、長音（ユーザ／ユーザー）、送り仮名（行う／行なう）のゆれを見つけて統一する", () => app.openProofing(), { big: true }),
        btn("check", "オートコレクトの|オプション", "入力中に (c) を © に、「・ 」「1. 」で始めると箇条書き・段落番号に。種類ごとにオン・オフ", () => app.openAutoCorrect(), { big: true })),
      group("音声",
        btn("readAloud", "音声|読み上げ", "選んだ文字・部品、またはスライドの文字を順に読み上げる（読んでいる所に色。もう一度押すと止まる）", () => readAloud.toggle(), { big: true, pressed: () => readAloud.active })),
      group("言語",
        drop("textCase", "翻訳", "スライドの文字をほかの言語に訳す（AI。部品・ノートも。数値と固有名詞はそのまま）", () => menu([
          { head: "このスライドを翻訳" },
          ...[["en", "英語"], ["zh", "中国語"], ["ko", "韓国語"], ["ja", "日本語"]].map(([k, l]) => ({ label: `${l}に`, run: () => app.translate(k, "slide") })),
          "-", { head: "資料全体を翻訳" },
          ...[["en", "英語"], ["zh", "中国語"], ["ko", "韓国語"], ["ja", "日本語"]].map(([k, l]) => ({ label: `資料全体を${l}に`, run: () => app.translate(k, "deck") })),
        ]), { big: true, enabled: () => app.canAi() })),
      group("検索", btn("find", "検索・|置換", "資料全体の文字を検索・置換（⌘F）", () => app.openReplace(), { big: true }),
        btn("font", "フォントの|置換", "資料で使っているフォントを、別のフォントにまとめて置き換える", () => extras.fontReplaceDialog(), { big: true })),
      group("ノート", btn("notes", "ノートを|作成", "スピーカーノートをまとめて作る（簡易・AI）", () => app.openNotes(), { big: true })),
      group("AI", btn("magic", "AIで全体を|見直す", "指示を出して資料全体をAIに見直してもらう（Codexに接続しているとき）", () => app.reviseDeck(), { big: true, enabled: () => app.canAi() })),
      group("比較",
        btn("thumbs", "比較", "この資料を、保存した版・保存庫の資料・ファイルと比べ、違うスライドを1枚ずつ採用する", () => app.openCompare(), { big: true })),
      group("インク",
        btn("pen", "インクの|非表示", "スライドに書いたインクを編集中だけ隠す（もう一度で表示）", () => { hideInk = !hideInk; document.getElementById("stageBody")?.classList.toggle("hide-ink", hideInk); refresh(); }, { big: true, pressed: () => hideInk })),
      group("レビュー",
        btn("review", "レビュー用|ファイル", "相手はブラウザで開いてスライドごとにコメントを書き、結果のファイルを返送します", () => app.exportReview(), { big: true }),
        btn("comment", "レビューを|取り込む", "コメント付きのPowerPoint、またはレビュー結果（.json）からコメントを読み込み、AIで反映します", () => app.importReview(), { big: true })),
    ];
  }

  function insertTab() {
    return [
      group("スライド", drop("slide", "新しい|スライド", "スライドを追加", menu([
        { label: "白紙（自由配置）", icon: "slide", run: () => app.insertBlankSlide() },
        { label: "レイアウトを選んで追加…", icon: "layout", run: () => app.openTypeDialog("insert") },
        "-",
        { label: "スライドの再利用…", icon: "duplicateSlide", run: () => app.openReuse() },
        { label: "アウトラインからスライド…", icon: "textLeft", run: () => app.openOutlineSlides() },
      ]), { big: true })),
      group("画像",
        drop("image", "画像", "画像を入れる", menu([
          { label: "このデバイス…", icon: "image", run: insertImages },
          { label: "内蔵の写真…", icon: "image", run: () => openPop(ribbon.querySelector('[data-rb="photos"]') || ribbon, photoGallery(insertPhoto)) },
          { label: "URLから…", icon: "link", run: () => insertFromUrl("image") },
          "-",
          { label: "フォト アルバム…", icon: "image", run: () => app.openPhotoAlbum() },
        ]), { big: true }),
        btn("screenshot", "スクリーン|ショット", "ウィンドウ・タブ・画面を撮って画像として入れる", () => extras.insertScreenshot(), { big: true }),
        h("span", { "data-rb": "photos" })),
      group("表", drop("table", "表", "表を入れる（行と列を選ぶ）", () => tables.tablePicker(), { big: true })),
      group("図",
        drop("shapes", "図形", "図形を描く", () => shapeGallery(pickTool), { big: true }),
        drop("icon", "アイコン", "アイコンを入れる", () => iconGallery(insertIcon), { big: true }),
        drop("chartBar", "グラフ", "グラフを入れる（データは表で入力）", () => tables.chartPicker(), { big: true }),
        drop("smartart", "SmartArt", "SmartArt グラフィック：リスト・手順・循環・階層構造・集合関係・マトリックス・ピラミッド", () => smart.insertGallery(), { big: true }),
        drop("model3d", "3D|モデル", "3D モデル（GLB・glTF）を入れる：このデバイス・ストック 3D モデル。回転・パンとズーム・3D のアニメーション", () => models3d.insertMenu(), { big: true })),
      group("テキスト",
        drop("textbox", "テキスト ボックス", "テキストボックスを描く", menu([
          { label: "横書きテキスト ボックス", icon: "textbox", run: () => pickTool({ kind: "text" }) },
          { label: "縦書きテキスト ボックス", icon: "vtext", run: () => pickTool({ kind: "text", vertical: true }) },
        ]), { big: true }),
        drop("wordart", "ワード|アート", "ワードアート：見出し向きの文字を入れる", () => extras.wordArtMenu(), { big: true }),
        col(drop("number", "スライド番号", "スライド番号・総数・日付のフィールド（自動で更新）", () => extras.fieldMenu(), { keep: true }))),
      group("線",
        col(btn("line", "直線", "直線を引く", () => pickTool({ kind: "line", variant: "line" })), btn("arrow", "矢印", "矢印を引く", () => pickTool({ kind: "line", variant: "arrow" }))),
        col(btn("arrows", "両矢印", "両方向の矢印", () => pickTool({ kind: "line", variant: "double" })), btn("line", "コネクタ", "図形どうしをつなぐカギ線（端を図形の点に近づけるとつながります）", () => pickTool({ kind: "line", variant: "elbow" })))),
      group("メディア",
        drop("video", "ビデオ", "動画・アニメーションを入れる", menu([
          { label: "このデバイスのビデオ・Lottie…", icon: "video", run: insertVideoFile },
          { label: "YouTube・URL…", icon: "link", run: () => insertFromUrl("video") },
        ]), { big: true }),
        drop("audio", "オーディオ", "音声を入れる：ファイル・この場で録音・URL", () => media.audioMenu(), { big: true }),
        btn("screenRec", "画面|録画", "画面・ウィンドウ・タブを録画して、ビデオとして入れる（マイクの音声も）", () => media.recorder("screen"), { big: true }),
        btn("camera", "カメオ", "カメラの映像（発表中にライブで映る。丸や角丸に切り抜けます）", () => extras.insertCamera(), { big: true })),
      group("フォーム",
        btn("poll", "アンケート", "質問と選択肢：発表中にクリックで回答、オンライン プレゼンテーションでは視聴者の回答がその場で棒グラフに", () => polls.insert(), { big: true })),
      group("数式と記号",
        btn("equation", "数式", "数式を入れる（TeXの書き方。分数・ルート・Σ…）", () => extras.equationDialog(), { big: true }),
        col(drop("symbol", "記号と特殊文字", "記号を入れる（文字の入力中はカーソルの位置に）", () => symbolGallery(), { keep: true }),
          drop("date", "日付と時刻", "今日の日付を入れる（文字の入力中はカーソルの位置に）", () => dateMenu(), { keep: true }))),
      group("リンク", zoomDrop(),
        btn("qr", "QR|コード", "文字（Webページのアドレスなど）を QR コードの図にして入れる。アドレスなら発表中にクリックして開けます", () => insertQr(), { big: true }),
        h("span", { "data-rb": "link" }, btn("link", "リンク・|動作", "選んだ文字にリンク／図形をクリックしたときの動作（スライドへ移動・Webページを開く）", editLink, { big: true, keep: true, enabled: () => Boolean(one()) || editor.typing }))),
    ];
  }

  /** 挿入 → ズーム: サマリー・スライド・セクション (the slide picker opens under the same button). */
  function zoomDrop() {
    const el = drop("zoomSlide", "ズーム", "ズーム：別のスライドへ飛んで戻ってくる、スライドの縮小版を入れる", () => extras.zoomMenu(el), { big: true });
    return el;
  }

  function arrangeMenu() {
    return menu([
      { head: "オブジェクトの順序" },
      { label: "最前面へ移動", icon: "front", keys: "⇧⌘]", run: () => editor.order("front") },
      { label: "前面へ移動", icon: "front", keys: "⌘]", run: () => editor.order("forward") },
      { label: "背面へ移動", icon: "back", keys: "⌘[", run: () => editor.order("backward") },
      { label: "最背面へ移動", icon: "back", keys: "⇧⌘[", run: () => editor.order("back") },
      "-", { head: "配置（1つならスライドに。複数は下の基準で）" },
      { label: "左揃え", icon: "alignLeft", run: () => editor.alignSelection("left") },
      { label: "左右中央揃え", icon: "alignCenter", run: () => editor.alignSelection("center") },
      { label: "右揃え", icon: "alignRight", run: () => editor.alignSelection("right") },
      { label: "上揃え", icon: "alignTop", run: () => editor.alignSelection("top") },
      { label: "上下中央揃え", icon: "alignMiddle", run: () => editor.alignSelection("middle") },
      { label: "下揃え", icon: "alignBottom", run: () => editor.alignSelection("bottom") },
      { label: "左右に整列（等間隔）", icon: "distH", run: () => editor.distributeSelection("x") },
      { label: "上下に整列（等間隔）", icon: "distV", run: () => editor.distributeSelection("y") },
      { label: "スライドの中央に置く", icon: "alignCenter", run: () => { editor.alignSelection("center", true); editor.alignSelection("middle", true); } },
      { label: "スライドに合わせて配置", on: editor.state.alignTo === "slide", run: () => editor.setView("alignTo", "slide") },
      { label: "選択したオブジェクトを揃える", on: editor.state.alignTo !== "slide", run: () => editor.setView("alignTo", "auto") },
      "-", { head: "グループ化" },
      { label: "グループ化", icon: "group", keys: "⌘G", run: () => editor.groupSelection(), disabled: editor.selection.length < 2 },
      { label: "グループ解除", icon: "ungroup", keys: "⇧⌘G", run: () => editor.ungroupSelection(), disabled: !selected().some((o) => o.group) },
      { label: "再グループ化", icon: "group", title: "グループ解除したオブジェクトをもう一度グループにします（解除したオブジェクトを選んでから）", run: () => editor.regroupSelection(), disabled: !editor.canRegroup() },
      "-", { head: "回転" },
      { label: "右へ90°回転", icon: "rotate", run: () => editor.rotateSelection(90) },
      { label: "左へ90°回転", icon: "rotate", run: () => editor.rotateSelection(-90) },
      { label: "上下反転", icon: "flipV", run: () => editor.flipSelection("y") },
      { label: "左右反転", icon: "flipH", run: () => editor.flipSelection("x") },
    ]);
  }
  function styleMenu() {
    return (close) => h("div", { class: "rb-styles rb-theme-styles" }, quickStyles(E).map(([label, style]) => {
      const sample = h("button", { type: "button", title: label, onmousedown: (e) => e.preventDefault(), onclick: () => { close(); editor.apply((o) => (["shape", "text"].includes(o.kind) ? { ...style, gradient: undefined, pattern: undefined } : null)); } },
        h("span", { class: "rb-style", style: { background: style.fill === "none" ? "transparent" : style.fill, border: style.stroke !== "none" ? `2px solid ${style.stroke}` : "1px solid transparent", color: style.color } }, "Aa"), h("small", {}, label));
      return sample;
    }));
  }
  /** 図のスタイル: each a little picture frame drawn as it would look; one ⌘Z takes it back. */
  function pictureStyleMenu() {
    return (close) => h("div", { class: "rb-styles rb-pic-styles" }, ops.PICTURE_STYLES.map(([label, look]) => {
      const radius = look.mask === "ellipse" ? "50%" : look.mask === "roundRect" ? "18%" : "0";
      const frame = h("span", { class: "rb-pic-style", style: { "border-radius": radius, border: look.stroke ? `${Math.max(2, Math.min(7, look.strokeW / 4))}px ${look.cmpd ? "double" : "solid"} ${look.stroke}` : "1px solid #d9d9d9", ...(look.mask === "snip2DiagRect" ? { "clip-path": "polygon(0 0, 85% 0, 100% 15%, 100% 100%, 15% 100%, 0 85%)" } : {}) } });
      return h("button", { type: "button", title: label, "data-picture-style": label, onmousedown: (e) => e.preventDefault(), onclick: () => { close(); editor.apply((o) => (o.kind === "image" ? look : null)); } }, frame, h("small", {}, label));
    }));
  }
  function outlineMenu() {
    return (close) => {
      const box = h("div", { class: "rb-outline" });
      box.append(colors(E.PALETTE.line, strokeOf(), setStroke, { none: "枠線なし" })(close));
      box.append(h("div", { class: "rb-menu-head" }, "太さ"), h("div", { class: "rb-menu" }, LINE_WIDTHS.map((pt) => h("button", { type: "button", onclick: () => { close(); setStrokeWidth(pt); } }, h("span", { class: "rb-weight", style: { "border-top-width": `${Math.max(1, pt)}px` } }), h("span", {}, `${pt} pt`)))));
      box.append(h("div", { class: "rb-menu-head" }, "実線/点線"), h("div", { class: "rb-menu" }, Object.entries(E.DASHES).map(([key, [label]]) => h("button", { type: "button", onclick: () => { close(); setDash(key); } }, dashSample(key), h("span", {}, label)))));
      box.append(h("div", { class: "rb-menu-head" }, "複合線"), h("div", { class: "rb-menu", "data-lines": "cmpd" }, [["", "一重線"], ...Object.entries(E.COMPOUNDS)].map(([key, label]) => h("button", { type: "button", "data-cmpd": key, onclick: () => { close(); setLineLook({ cmpd: key || undefined }); } }, lineLookSample("cmpd", key), h("span", {}, label)))));
      box.append(h("div", { class: "rb-menu-head" }, "線端"), h("div", { class: "rb-menu", "data-lines": "cap" }, Object.entries(E.LINE_CAPS).map(([key, label]) => h("button", { type: "button", "data-cap": key, onclick: () => { close(); setLineLook({ cap: key === "flat" ? undefined : key }); } }, lineLookSample("cap", key), h("span", {}, label)))));
      box.append(h("div", { class: "rb-menu-head" }, "結合点"), h("div", { class: "rb-menu", "data-lines": "join" }, Object.entries(E.LINE_JOINS).map(([key, label]) => h("button", { type: "button", "data-join": key, onclick: () => { close(); setLineLook({ join: key }); } }, lineLookSample("join", key), h("span", {}, label)))));
      return box;
    };
  }
  /** A small picture of a compound line, a line end or a join. */
  function lineLookSample(kind, key) {
    const svg = E.s("svg", { viewBox: "0 0 60 14", width: 60, height: 14, "aria-hidden": "true" });
    const put = (attrs) => svg.append(E.s("path", { fill: "none", stroke: "currentColor", ...attrs }));
    if (kind === "cmpd") {
      const rows = key === "tri" ? [2.5, 7, 11.5] : key === "dbl" ? [4, 10] : [7];
      for (const y of rows) put({ d: `M2 ${y}H58`, "stroke-width": key === "tri" ? 2 : key === "dbl" ? 3 : 2.5 });
    } else if (kind === "cap") put({ d: "M12 7H48", "stroke-width": 8, "stroke-linecap": { flat: "butt", round: "round", square: "square" }[key] });
    else put({ d: "M10 12L30 3L50 12", "stroke-width": 6, "stroke-linejoin": key });
    return svg;
  }
  function dashSample(key) {
    const svg = E.s("svg", { viewBox: "0 0 60 8", width: 60, height: 8, "aria-hidden": "true" });
    const pattern = E.DASHES[key][1];
    svg.append(E.s("path", { d: "M2 4H58", stroke: "currentColor", "stroke-width": 2, "stroke-dasharray": pattern ? pattern.map((v) => Math.max(v * 2, v === 0 ? 0.01 : 2)).join(" ") : null, "stroke-linecap": key === "roundDot" ? "round" : "butt" }));
    return svg;
  }
  function arrowMenu() {
    return (close) => h("div", { class: "rb-arrows" }, [["none", "none"], ["none", "triangle"], ["triangle", "none"], ["triangle", "triangle"], ["none", "arrow"], ["none", "stealth"], ["none", "oval"], ["oval", "triangle"], ["diamond", "triangle"], ["oval", "oval"]].map(([head, tail]) => h("button", { type: "button", onclick: () => { close(); editor.apply((o) => (o.kind === "line" ? { head: head === "none" ? undefined : head, tail: tail === "none" ? undefined : tail } : null)); } }, arrowSample(head, tail))));
  }
  function arrowSample(head, tail) {
    const svg = E.s("svg", { viewBox: "0 0 70 14", width: 70, height: 14 });
    svg.append(E.s("path", { d: "M8 7H62", stroke: "currentColor", "stroke-width": 1.6 }));
    const mark = (kind, x, dir) => {
      if (kind === "none") return;
      if (kind === "oval") svg.append(E.s("circle", { cx: x, cy: 7, r: 3.4, fill: "currentColor" }));
      else if (kind === "diamond") svg.append(E.s("path", { d: `M${x - 4} 7 L${x} 3 L${x + 4} 7 L${x} 11 Z`, fill: "currentColor" }));
      else if (kind === "arrow") svg.append(E.s("path", { d: `M${x - 6 * dir} 2 L${x} 7 L${x - 6 * dir} 12`, fill: "none", stroke: "currentColor", "stroke-width": 1.6 }));
      else svg.append(E.s("path", { d: kind === "stealth" ? `M${x} 7 L${x - 8 * dir} 2 L${x - 5 * dir} 7 L${x - 8 * dir} 12 Z` : `M${x} 7 L${x - 8 * dir} 2 L${x - 8 * dir} 12 Z`, fill: "currentColor" }));
    };
    mark(head, 6, -1);
    mark(tail, 64, 1);
    return svg;
  }

  function sizeFields() {
    const field = (label, key) => {
      const input = h("input", { type: "number", step: "0.01", min: "0.05", class: "rb-num", "aria-label": `${label}（cm）`,
        onchange: (event) => {
          const v = parseFloat(event.target.value);
          if (!(v > 0)) return;
          const o = one();
          if (!o || o.kind === "line") return;
          const px = ops.fromCm(v);
          const ratio = o.lockRatio || o.kind === "image" ? o.w / o.h : null;
          const patch = key === "w" ? { w: px, ...(ratio ? { h: px / ratio } : {}) } : { h: px, ...(ratio ? { w: px * ratio } : {}) };
          // Resize about the centre, as the boxes in PowerPoint's ribbon do.
          patch.x = o.x + (o.w - (patch.w ?? o.w)) / 2;
          patch.y = o.y + (o.h - (patch.h ?? o.h)) / 2;
          editor.apply(patch, { ids: [o.id] });
        } });
      updaters.push(() => { const o = one(); input.disabled = !o || o.kind === "line"; if (document.activeElement !== input) input.value = o && o.kind !== "line" ? String(ops.toCm(o[key])) : ""; });
      return h("label", { class: "rb-field" }, h("span", {}, label), input, h("small", {}, "cm"));
    };
    return col(field("高さ", "h"), field("幅", "w"));
  }

  function shapeTab() {
    return [
      group("図形の挿入",
        drop("shapes", "図形", "図形を描く", () => shapeGallery(pickTool), { big: true }),
        col(drop("change", "図形の変更", "選んだ図形の形を変える", () => shapeGallery((tool) => { if (tool.kind === "shape") editor.apply((o) => (["shape", "text"].includes(o.kind) ? { kind: "shape", shape: tool.shape, adj: undefined } : null)); }, { lines: false }), { enabled: hasText }),
          btn("textbox", "テキスト ボックス", "テキストボックスを描く", () => pickTool({ kind: "text" })),
          drop("union", "図形の結合", "選んだ2つ以上の図形を1つにする：接合・型抜き/合成・切り出し・重なり抽出・単純型抜き（最初に選んだ図形の書式になります）", () => mergeMenu(), { enabled: () => canMerge(selected()) }))),
      group("図形のスタイル",
        drop("style", "クイック|スタイル", "テーマの色の組み合わせ", () => styleMenu(), { big: true, enabled: hasText }),
        col(drop("fill", "塗りつぶし", "図形の塗りつぶし", () => colors(E.PALETTE.fill, fillOf(), setFill, { none: "塗りつぶしなし", extra: fillExtras }), { enabled: hasText, swatch: fillOf }),
          drop("outline", "枠線", "図形の枠線（色・太さ・実線/点線）", () => outlineMenu(), { enabled: any, swatch: strokeOf }),
          drop("arrows", "矢印", "線の始点・終点の形", () => arrowMenu(), { enabled: () => kinds().has("line") })),
        col(drop("effects", "効果", "図形の効果：影・反射・光彩・ぼかし", () => effectsMenu(), { enabled: hasEffectTarget }),
          drop("transparency", "透明度", "オブジェクトの透明度", () => menu([0, 0.15, 0.3, 0.5, 0.7].map((t) => ({ label: `${Math.round(t * 100)}%`, on: Math.abs((1 - (one()?.opacity ?? 1)) - t) < 0.01, run: () => editor.apply({ opacity: t ? Math.round((1 - t) * 100) / 100 : undefined }) }))), { enabled: any }),
          drop("line", "線の種類", "直線・カギ線・曲線", () => menu(Object.entries(E.ROUTES).map(([key, label]) => ({ label, run: () => editor.apply((o) => (o.kind === "line" ? { route: key === "straight" ? undefined : key } : null)) }))), { enabled: () => kinds().has("line") }))),
      group("ワードアートのスタイル",
        col(drop("fontColor", "文字の塗りつぶし", "文字の色", () => colors(E.PALETTE.text, textColorOf(), (c) => editor.textFormat("color", c)), { enabled: hasText, swatch: textColorOf, keep: true }),
          drop("textbox", "文字の配置", "文字の配置と余白", () => menu([["top", "上揃え", "valignTop"], ["middle", "上下中央揃え", "valignMiddle"], ["bottom", "下揃え", "valignBottom"]].map(([v, label, icon]) => ({ label, icon, on: editor.textState()?.valign === v, run: () => editor.textFormat("valign", v) }))), { enabled: hasText }),
          drop("textFx", "文字の効果", "文字の効果：文字の影・光彩", () => textEffectsMenu(), { enabled: hasText }))),
      group("アクセシビリティ", btn("textbox", "代替|テキスト", "画面読み上げが読む図形の説明（装飾用にもできます）", () => { const o = one(); if (o) app.editAlt(o.id); }, { big: true, enabled: () => Boolean(one()) })),
      arrangeGroup(),
      group("サイズ", sizeFields()),
    ];
  }
  /** 図形の結合: the five ways, in PowerPoint's order (the order of selection decides the first shape). */
  function mergeMenu() {
    return menu([
      { head: "図形の結合" },
      ...Object.entries(MERGE_MODES).map(([mode, label]) => ({ label, icon: mode, title: MERGE_HINTS[mode], run: () => mergeSelection(mode), hover: (on) => mergePreview(on ? mode : null) })),
    ]);
  }
  /** Live preview, as PowerPoint shows it while the pointer is on a choice: the result's outline over the slide. */
  let previewFor = null;
  async function mergePreview(mode) {
    previewFor = mode;
    if (!mode) { editor.draw(); return; }
    const chosen = editor.selection.map((id) => editor.objects().find((o) => o.id === id)).filter(Boolean);
    if (!canMerge(chosen)) return;
    let pc;
    try { pc = await clippingLib(); } catch { return; }
    if (previewFor !== mode) return;
    const made = mergeObjects(chosen, mode, E, pc, () => "preview") || [];
    const k = editor.scale();
    const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
    svg.setAttribute("class", "ed-merge-preview");
    svg.setAttribute("width", String(E.W * k));
    svg.setAttribute("height", String(E.H * k));
    for (const o of made) {
      const path = document.createElementNS("http://www.w3.org/2000/svg", "path");
      const d = E.geometry("custom", o.w, o.h, null, o.path).paths[0];
      path.setAttribute("d", d);
      path.setAttribute("fill-rule", "evenodd");
      path.setAttribute("transform", `translate(${o.x * k} ${o.y * k}) scale(${k})`);
      svg.append(path);
    }
    editor.draw(null, [svg]);
  }
  async function mergeSelection(mode) {
    const order = editor.selection;
    const chosen = order.map((id) => editor.objects().find((o) => o.id === id)).filter(Boolean);
    if (!canMerge(chosen)) { app.toast("図形を2つ以上選んでください（Shift＋クリック）"); return; }
    let pc;
    try { pc = await clippingLib(); } catch (error) { app.toast(error.message); return; }
    const made = mergeObjects(chosen, mode, E, pc, () => E.newObjectId());
    if (!made?.length) { app.toast(mode === "intersect" ? "重なっている部分がありません" : "結合すると何も残りません"); return; }
    const gone = new Set(chosen.filter((o) => ["shape", "text"].includes(o.kind) && !o.locked).map((o) => o.id));
    const list = editor.objects();
    // The result takes the first shape's place in the stacking order.
    const at = list.findIndex((o) => o.id === chosen[0].id);
    const next = [...list.slice(0, at).filter((o) => !gone.has(o.id)), ...made, ...list.slice(at).filter((o) => !gone.has(o.id))];
    editor.commit(next, { select: made.map((o) => o.id) });
    app.toast(`図形を${MERGE_MODES[mode]}しました${made.length > 1 ? `（${made.length}個の図形）` : ""}。頂点の編集で形を直せます`);
  }
  function arrangeGroup() {
    return group("配置",
      col(btn("front", "前面へ", "前面へ移動（最前面は ⇧⌘]）", () => editor.order("forward"), { enabled: any }), btn("back", "背面へ", "背面へ移動（最背面は ⇧⌘[）", () => editor.order("backward"), { enabled: any }), btn("pane", "選択ウィンドウ", "オブジェクトの一覧", () => app.openPanel("format", "selection"))),
      col(drop("alignLeft", "配置", "揃える・等間隔に並べる", () => arrangeMenu(), { enabled: any }), btn("group", "グループ化", "グループ化（⌘G）／もう一度でグループ解除", () => (selected().some((o) => o.group) && selected().every((o) => o.group === selected()[0].group) ? editor.ungroupSelection() : editor.groupSelection()), { enabled: () => editor.selection.length > 1 || selected().some((o) => o.group) }),
        drop("rotate", "回転", "回転・反転", () => menu([{ label: "右へ90°回転", icon: "rotate", run: () => editor.rotateSelection(90) }, { label: "左へ90°回転", icon: "rotate", run: () => editor.rotateSelection(-90) }, { label: "上下反転", icon: "flipV", run: () => editor.flipSelection("y") }, { label: "左右反転", icon: "flipH", run: () => editor.flipSelection("x") }]), { enabled: any })));
  }

  function pictureTab() {
    const img = () => selected().find((o) => o.kind === "image");
    // 図のリセット (with the size, 図とサイズのリセット): every chosen picture, one Undo.
    const resetImages = (withSize) => editor.resetPicture(selected().filter((o) => o.kind === "image").map((o) => o.id), withSize);
    return [
      group("調整",
        btn("eraser", "背景の|削除", "写真・ロゴの一色に近い背景を透明にする（周りから背景を見つけます）", () => pictureTools.background(), { big: true, enabled: hasImage }),
        drop("bright", "修整", "明るさ・コントラスト", () => menu([
          { head: "明るさ / コントラスト" },
          ...[[-0.4, 0], [-0.2, 0], [0, 0], [0.2, 0], [0.4, 0], [0, 0.2], [0, 0.4], [0.2, 0.2], [-0.2, 0.2]].map(([b, c]) => ({ label: `明るさ ${b > 0 ? "+" : ""}${Math.round(b * 100)}% / コントラスト ${c > 0 ? "+" : ""}${Math.round(c * 100)}%`, run: () => editor.apply((o) => (o.kind === "image" ? { bright: b || undefined, contrast: c || undefined } : null)) })),
        ]), { big: true, enabled: hasImage }),
        drop("magic", "アート|効果", "ぼかし・鉛筆スケッチ・線画・モザイク・セピア・ポスター（新しい画像として保存、⌘Zで戻せます）", () => menu([{ head: "アート効果" }, ...Object.entries(ART_EFFECTS).map(([k, label]) => ({ label, run: () => pictureTools.art(k) }))]), { big: true, enabled: hasImage }),
        drop("fill", "色", "色の変更（グレースケール・セピア・ウォッシュアウト・テーマの色調）と透明色", () => menu([{ head: "色の変更" }, ...Object.entries(RECOLORS).map(([k, label]) => ({ label, run: () => pictureTools.colorize(k) })), "-", { label: "透明色を指定", icon: "transparency", run: () => pictureTools.transparentColor() }]), { big: true, enabled: hasImage }),
        col(btn("gray", "グレースケール", "白黒にする（もう一度で戻す）", () => editor.apply((o) => (o.kind === "image" ? { gray: !o.gray || undefined } : null)), { enabled: hasImage, pressed: () => img()?.gray }),
          drop("transparency", "透明度", "図の透明度", () => menu([0, 0.15, 0.3, 0.5, 0.7].map((t) => ({ label: `${Math.round(t * 100)}%`, run: () => editor.apply({ opacity: t ? Math.round((1 - t) * 100) / 100 : undefined }) }))), { enabled: hasImage })),
        col(btn("transparency", "透明色を指定", "図の上でクリックした色を透明にする", () => pictureTools.transparentColor(), { enabled: hasImage }),
          btn("down", "図の圧縮", "画像を小さくして資料を軽くする（トリミング部分の削除も）", () => pictureTools.compressDialog())),
        col(btn("change", "図の変更", "画像を差し替える（大きさ・位置はそのまま）", async () => { const o = img(); if (!o) return; const [file] = await app.pickFiles("image/*", false); if (file) await app.replaceImage(o.id, file); }, { enabled: hasImage }),
          row(btn("reset", "リセット", "トリミング・修整・枠線・影を元に戻す", () => resetImages(false), { enabled: hasImage }), caret("図のリセット・図とサイズのリセット", () => menu([
            { label: "図のリセット", icon: "reset", run: () => resetImages(false) },
            { label: "図とサイズのリセット", icon: "zoomFit", run: () => resetImages(true) },
          ]), { enabled: hasImage })))),
      group("アクセシビリティ", btn("textbox", "代替|テキスト", "画面読み上げが読む図の説明（装飾用にもできます）", () => { const o = one(); if (o) app.editAlt(o.id); }, { big: true, enabled: () => Boolean(one()) })),
      group("図のスタイル",
        drop("style", "クイック|スタイル", "図のスタイル：枠線・角丸・楕円・二重線などの組み合わせ（テーマの色。図の大きさと修整はそのまま）", () => pictureStyleMenu(), { big: true, enabled: hasImage }),
        drop("outline", "図の枠線", "枠線の色・太さ", () => outlineMenu(), { enabled: hasImage, swatch: strokeOf }),
        drop("effects", "図の効果", "図の効果：影・反射・光彩・ぼかし", () => effectsMenu(), { enabled: hasImage }),
        drop("mask", "図形に合わせて|切り抜き", "画像を図形の形に切り抜く", () => (close) => h("div", { class: "rb-gallery" }, h("div", { class: "rb-gallery-grid" }, MASKS.map((key) => h("button", { type: "button", title: E.SHAPES[key].label, onclick: () => { close(); editor.apply((o) => (o.kind === "image" ? { mask: key === "rect" ? undefined : key, adj: undefined } : null)); } }, shapeThumb(key))))), { big: true, enabled: hasImage }),
        drop("smartart", "図の|レイアウト", "図のレイアウト：選んだ図をキャプション付きで並べる（横に並べる・2列のグリッド・図と説明）。図とキャプションはグループになります", () => menu(Object.entries(ops.PICTURE_LAYOUTS).map(([kind, label]) => ({ label, run: () => {
          const { list, ids } = ops.pictureLayout(editor.objects(), editor.selection, kind);
          if (!ids.length) { app.toast("並べる図を選んでください"); return; }
          editor.commit(list, { select: ids });
          app.toast("図をキャプション付きで並べました（キャプションはダブルクリックで書き換えられます。⌘Zで戻せます）");
        } }))), { big: true, enabled: hasImage })),
      arrangeGroup(),
      group("サイズ",
        btn("crop", "トリミング", "画像の端を切り取る：黒い印をドラッグ（Enterで確定）。数値は書式パネルで", () => { const o = one(); if (o?.kind === "image") crop.start(o.id); else app.openPanel("format", "picture"); }, { big: true, enabled: hasImage, pressed: () => crop.active }),
        caret("トリミングのオプション（縦横比・塗りつぶし・枠に合わせる）", () => cropMenu(), { enabled: () => one()?.kind === "image" }),
        sizeFields()),
    ];
  }

  // PowerPoint's zoom steps (100% is the slide at its real size: 13.33 inches at 96 dpi).
  const ZOOMS = [400, 300, 200, 150, 100, 66, 50, 33];
  function zoomMenu() {
    return menu([{ head: "ズーム" }, { label: "スライドを現在のウィンドウに合わせる", icon: "zoomFit", keys: "⌘0", on: editor.zoom === "fit", run: () => editor.setZoom("fit") }, "-",
      ...ZOOMS.map((z) => ({ label: `${z}%`, on: editor.zoom === z, run: () => editor.setZoom(z) }))]);
  }

  function viewTab() {
    const st = editor.state;
    return [
      group("プレゼンテーションの表示",
        btn("slide", "標準", "1枚ずつ編集", () => app.setView("single"), { big: true, pressed: () => app.state().view === "single" }),
        col(btn("grid", "スライド一覧", "スライド一覧", () => app.setView("grid"), { pressed: () => app.state().view === "grid" }),
          btn("textLeft", "構成", "見出しとキーメッセージの一覧（アウトライン）", () => app.setView("outline"), { pressed: () => app.state().view === "outline" }),
          btn("previewPlay", "閲覧表示", "このスライドの動きをこの画面で確認する", () => app.previewMotion())),
        btn("notes", "ノート|表示", "スライドとノートを1ページにした表示で、ノートを書く（←→で前後のスライド）", () => app.openNotesView(), { big: true })),
      group("マスター表示",
        btn("slide", "スライド|マスター", "スライド マスター表示：すべてのスライドに出す図形・文字・図を置く", () => (app.masterView() ? app.closeMaster() : app.openMaster()), { big: true, pressed: () => app.masterView() }),
        btn("print", "配布資料|マスター", "配布資料マスター：配布資料・アウトラインの用紙（縦・横）、ヘッダー・日付・フッター・ページ番号の場所と文字、全ページに出す図・文字を決める", () => app.openPrintMaster("handout"), { big: true }),
        btn("notes", "ノート|マスター", "ノート マスター：ノート ページの用紙、スライドとノートの場所と大きさ、ヘッダー・日付・フッター・ページ番号、全ページに出す図・文字を決める", () => app.openPrintMaster("notes"), { big: true })),
      group("表示/非表示",
        col(btn("ruler", "ルーラー", "スライドの上と左に目盛り（cm、スライドの中央が0）を表示", () => app.toggleRulers(), { pressed: () => app.rulersShown() }),
          btn("grid", "グリッド線", "グリッド線を表示（間隔はグリッドとガイドの設定で）", () => editor.setView("grid", !st.grid), { pressed: () => st.grid }),
          drop("guides", "ガイド", "ガイド：本文の領域・中央と、自分で置くガイド（ドラッグで移動・スライドの外へ出すと削除）", () => menu([
            { label: "ガイドを表示する（吸着）", icon: "guides", on: st.guides, run: () => editor.setView("guides", !st.guides) },
            "-",
            { label: "垂直ガイドを追加", icon: "alignCenter", run: () => editor.addGuide("x") },
            { label: "水平ガイドを追加", icon: "alignMiddle", run: () => editor.addGuide("y") },
            { label: "ガイドをすべて削除", icon: "trash", disabled: !(editor.customGuides().x.length || editor.customGuides().y.length), run: () => editor.clearGuides() },
          ]))),
        col(btn("notes", "ノート", "スライドの下にスピーカーノートを表示", () => app.toggleNotes(), { pressed: () => app.notesShown() }),
          btn("smart", "スマートガイド", "ほかの図形の端・中央・等間隔に吸着（Altで一時的に外す）", () => editor.setView("smart", !st.smart), { pressed: () => st.smart }),
          btn("snapGrid", "グリッドに|合わせる", "グリッド線に吸着（間隔はグリッドとガイドの設定で）", () => editor.setView("snapGrid", !st.snapGrid), { pressed: () => st.snapGrid })),
        btn("guides", "グリッドと|ガイド…", "グリッドとガイドの設定（間隔・グリッドに合わせる・グリッドの表示・ガイド・スマートガイド）", () => gridDialog(), { big: true })),
      group("カラー/グレースケール",
        btn("grayView", "グレース|ケール", "スライドを白黒で見る（白黒印刷やコピーで読めるかの確認。資料は変わりません）", () => { grayView = !grayView; document.getElementById("stageBody")?.classList.toggle("view-gray", grayView); refresh(); }, { big: true, pressed: () => grayView })),
      group("ズーム",
        drop("zoomIn", "ズーム", "表示の倍率", () => zoomMenu(), { big: true }),
        btn("zoomFit", "ウィンドウに|合わせる", "スライド全体を現在のウィンドウに合わせる（⌘0）", () => editor.setZoom("fit"), { big: true, pressed: () => editor.zoom === "fit" })),
      group("ウィンドウ",
        col(btn("thumbs", "サムネイル", "左のスライドのサムネイルを表示する", () => app.togglePane("film"), { pressed: () => app.paneOpen("film") }),
          btn("taskPane", "作業ウィンドウ", "右の作業ウィンドウ（AIと話す・編集・書式・アニメーション）を表示する", () => app.togglePane("side"), { pressed: () => app.paneOpen("side") }),
          btn("pane", "選択ウィンドウ", "オブジェクトの一覧", () => app.openPanel("format", "selection"))),
        drop("ribbon", "リボンの|表示", "リボンの表示オプション（常に表示・タブのみ・自動的に非表示。⌘F1で折りたたみ）", () => menu(Object.entries(RIBBON_MODES).map(([mode, label]) => ({ label, on: ribbonMode === mode, run: () => setRibbonMode(mode) }))), { big: true })),
    ];
  }

  /** 開発: the slide as code — its JSON (edit and apply, checked first) and the HTML it becomes; the deck's data and files. */
  function developerTab() {
    return [
      group("コード",
        btn("effectOpts", "スライドの|JSON", "このスライドのデータ（レイアウト・文字・部品・アニメーション・インタラクション）をJSONで見て直す", () => app.openSlideJson(), { big: true }),
        btn("textbox", "HTMLを|表示", "このスライドをエンジンが描いたHTMLを見る・コピーする", () => app.openSlideHtml(), { big: true })),
      group("資料のデータ",
        col(btn("down", "JSONで保存", "資料全体をJSONファイルに保存（別のPCで続きを編集できます）", () => app.saveJson()),
          btn("up", "JSONを読み込む", "JSONから資料を開く", () => app.openJson()))),
      group("書き出し",
        btn("slide", "PowerPoint|出力", "PowerPoint で編集できるファイル（.pptx）に書き出す", () => app.exportPptx(), { big: true }),
        btn("play", "HTML出力", "動きごと1つのHTMLファイルに書き出す", () => app.exportHtml(), { big: true }),
        btn("slide", "PDF", "PDFとして保存・印刷", () => app.printPdf(), { big: true })),
    ];
  }

  // ---------------------------------------------------------------- the ribbon's commands (検索・クイック アクセス)

  /**
   * Every command on the ribbon's tabs that are there now (each tab built off screen, as it would be shown), and the
   * ファイル menu's: { ref: "tab|group|title", tab, tabLabel, group, label, title, disabled, icon }.
   */
  function ribbonCommands() {
    const saved = updaters;
    const out = [];
    const seen = new Set();
    try {
      for (const t of TABS.filter((x) => !x.contextual || x.contextual())) {
        updaters = [];
        let nodes;
        try { nodes = buildTab(t.id); } catch { continue; }
        const box = h("div", {}, ...nodes);
        for (const fn of updaters) { try { fn(); } catch { /* a control without a target */ } }
        for (const g of box.querySelectorAll(".rb-group")) {
          const group = g.querySelector(":scope > .rb-label")?.textContent || "";
          const off = !onSlide() && !OFF_SLIDE_TABS.has(t.id) && !OFF_SLIDE_GROUPS.has(group);
          for (const b of g.querySelectorAll("button.rb-btn")) {
            const title = b.getAttribute("title") || "";
            const ref = `${t.id}|${group}|${title}`;
            if (!title || seen.has(ref)) continue;
            seen.add(ref);
            // The name as the search shows it: the button's words, or its tooltip's head when that says more
            // (「ウィンドウ」 → 「アニメーション ウィンドウ」; an icon-only 「太字（⌘B）」 → 「太字」).
            const words = (b.querySelector(":scope > span")?.textContent || "").trim();
            const head = title.split(/[：（(]/)[0].trim();
            const label = !words ? head || title : head && head.length <= 24 && head.includes(words) ? head : words;
            out.push({ ref, tab: t.id, tabLabel: t.label, group, label, title, disabled: b.disabled || off, icon: b.querySelector("svg")?.getAttribute("data-ico") || "" });
          }
        }
      }
    } finally { updaters = saved; }
    for (const item of fileItems()) {
      if (!item.run) continue;
      out.push({ ref: `file||${item.label}`, tab: "file", tabLabel: "ファイル", group: "", label: item.label.replace(/…$/, ""), title: item.label, disabled: false, icon: item.icon || "" });
    }
    return out;
  }

  /**
   * Runs a command of the ribbon as a click on its button would. From the 検索 box its tab is shown (a folded group
   * opened) and the button lit for a moment, so the place is learnt; from the クイック アクセス ツール バー (`from`, the
   * toolbar's button) the tab stays as it is: the command's button is built off screen, and a menu it opens comes
   * under the toolbar's button.
   */
  function runRibbonCommand(ref, { from = null } = {}) {
    const [id, group, ...rest] = String(ref || "").split("|");
    const title = rest.join("|");
    if (id === "file") {
      const item = fileItems().find((x) => x.run && x.label === title);
      if (item) item.run();
      return Boolean(item);
    }
    const t = TABS.find((x) => x.id === id);
    if (!t || (t.contextual && !t.contextual())) { app.toast("このコマンドは今は使えません（対象を選ぶと出るタブのコマンドです）"); return false; }
    if (app.isFinal?.() && !FINAL_TABS.has(id)) { app.toast("最終版です。編集するにはメッセージ バーの「編集する」を押してください"); return false; }
    closePop(true);
    if (from) {
      const saved = updaters;
      let found = null;
      try {
        updaters = [];
        const box = h("div", {}, ...buildTab(id));
        for (const fn of updaters) { try { fn(); } catch { /* a control without a target */ } }
        const g = [...box.querySelectorAll(".rb-group")].find((x) => (x.querySelector(":scope > .rb-label")?.textContent || "") === group);
        found = g && [...g.querySelectorAll("button.rb-btn")].find((b) => b.getAttribute("title") === title);
        if (found && !onSlide() && !OFF_SLIDE_TABS.has(id) && !OFF_SLIDE_GROUPS.has(group)) { app.toast("1枚表示（標準）で使えるコマンドです"); return false; }
      } finally { updaters = saved; }
      if (!found) { app.toast("コマンドが見つかりませんでした"); return false; }
      if (found.disabled) { app.toast("このコマンドは今は使えません"); return false; }
      found.click();
      if (pop && !pop.anchor.isConnected) {
        const r = from.getBoundingClientRect();
        const pr = pop.el.getBoundingClientRect();
        pop.el.style.left = `${Math.max(8, Math.min(r.left, innerWidth - pr.width - 8))}px`;
        pop.el.style.top = `${Math.min(r.bottom + 4, innerHeight - pr.height - 8)}px`;
        pop.anchor = from;
      }
      return true;
    }
    tab = id;
    if (ribbonMode === "full") renderRibbon(true); else openFloating(id);
    editor.draw();
    const g = [...ribbon.querySelectorAll(".rb-body > .rb-group")].find((x) => (x.querySelector(":scope > .rb-label")?.textContent || "") === group);
    const find = (root) => (root ? [...root.querySelectorAll("button.rb-btn")].find((b) => b.getAttribute("title") === title && !b.classList.contains("rb-folded")) : null);
    const b = g ? find(g.foldedItems) || find(g) : null;
    if (!b) { app.toast("コマンドが見つかりませんでした"); return false; }
    if (b.disabled || g.inert) { app.toast(g.inert ? "1枚表示（標準）で使えるコマンドです" : "このコマンドは今は使えません"); return false; }
    if (g.foldedItems?.contains(b)) g.querySelector(".rb-folded")?.click();
    b.scrollIntoView?.({ block: "nearest", inline: "nearest" });
    // The button stays lit for a moment even if the ribbon is drawn again meanwhile (its command may change the
    // window: the rulers, the panes, a tab coming or going).
    lit = { tab: id, group, title, until: Date.now() + 1600 };
    b.classList.add("rb-found");
    setTimeout(() => { if (lit && Date.now() >= lit.until) lit = null; ribbon.querySelectorAll(".rb-found").forEach((el) => el.classList.remove("rb-found")); }, 1600);
    b.click();
    relight();
    return true;
  }
  let lit = null;
  /** The button the search ran, lit again in a ribbon drawn since. */
  function relight() {
    if (!lit || Date.now() >= lit.until || tab !== lit.tab) return;
    const g = [...ribbon.querySelectorAll(".rb-body > .rb-group")].find((x) => (x.querySelector(":scope > .rb-label")?.textContent || "") === lit.group);
    const find = (root) => (root ? [...root.querySelectorAll("button.rb-btn")].find((x) => x.getAttribute("title") === lit.title && !x.classList.contains("rb-folded")) : null);
    (g ? find(g) || find(g.foldedItems) : null)?.classList.add("rb-found");
  }

  function buildTab(id) {
    return { home: homeTab, insert: insertTab, design: designTab, transition: anim.transitionTab, animation: anim.animationTab, interact: ix.tab, slideshow: slideshowTab, review: reviewTab, shape: shapeTab, picture: pictureTab, view: viewTab, developer: developerTab, tableDesign: tables.designTab, tableLayout: tables.layoutTab, chartDesign: tables.chartTab, playback: media.playbackTab, smartartDesign: smart.designTab, draw: ink.drawTab, zoomTool: extras.zoomTab, model3d: models3d.tab }[id]();
  }

  /**
   * Makes a tab the current one. When only the tabs show, the ribbon opens over the slide only when asked
   * (`open`: a button that says "open this tab"), not after an insertion, as in PowerPoint.
   */
  function showTab(id, { open = false } = {}) {
    tab = id;
    if (ribbonMode !== "full") { if (open) openFloating(id); else renderRibbon(true); return; }
    renderRibbon(true);
  }

  // Clipboard buttons without a keyboard event (Chrome only runs execCommand("copy") from a user gesture).
  function copyFallback() { const fake = new Event("copy"); fake.clipboardData = null; editor.onCopy(Object.assign(fake, { preventDefault() {} })); }
  function cutFallback() { const fake = new Event("cut"); editor.onCopy(Object.assign(fake, { preventDefault() {} }), true); }

  // ---------------------------------------------------------------- the 書式 pane (right panel)

  let paneFocus = null;
  function renderPane(focus = null) {
    if (!pane) return;
    if (focus) paneFocus = focus;
    const keepScroll = pane.querySelector(".fp-body")?.scrollTop ?? 0;
    const list = editor.objects();
    const chosen = selected();
    const body = h("div", { class: "fp-body" });
    const head = h("div", { class: "fp-head" });
    if (!onSlide()) {
      body.append(h("p", { class: "hint" }, "1枚表示（標準）で、スライドに図形・文字・画像を置いて書式を変えられます。"));
      pane.replaceChildren(head, body);
      return;
    }
    if (!chosen.length) {
      head.append(h("b", {}, "書式"), h("span", { class: "hint" }, "オブジェクトを選ぶと、ここで細かく設定できます"));
      body.append(h("div", { class: "fp-quick" },
        h("button", { type: "button", class: "btn", onclick: () => pickTool({ kind: "text" }) }, ico("textbox"), "テキスト ボックス"),
        h("button", { type: "button", class: "btn", onclick: (event) => openPop(event.currentTarget, shapeGallery(pickTool)) }, ico("shapes"), "図形"),
        h("button", { type: "button", class: "btn", onclick: insertImages }, ico("image"), "画像"),
        h("button", { type: "button", class: "btn", onclick: (event) => openPop(event.currentTarget, iconGallery(insertIcon)) }, ico("icon"), "アイコン"),
        h("button", { type: "button", class: "btn", onclick: () => app.insertBlankSlide() }, ico("slide"), "白紙のスライド")),
      h("p", { class: "hint fp-tips" }, "ドラッグで移動・四隅でサイズ・上の丸で回転。Shiftで比率を保つ、Altで吸着なし、Ctrlを押しながらドラッグでコピー。ダブルクリックで文字を入力。右クリックでメニュー。"));
    } else {
      const title = chosen.length === 1 ? E.objectName(chosen[0], list.indexOf(chosen[0])) : `${chosen.length}個のオブジェクト`;
      // (append writes a null as the word "null": the rename button only for one object)
      head.append(...[h("b", {}, title), chosen.length === 1 ? h("button", { type: "button", class: "btn btn-ghost btn-sm", title: "名前を変える", onclick: async () => { const name = await app.ask("オブジェクトの名前", "選択ウィンドウやアニメーションに表示される名前", E.objectName(chosen[0], list.indexOf(chosen[0]))); if (name != null) editor.rename(chosen[0].id, name); } }, "名前") : null].filter(Boolean));
      body.append(...paneSections(chosen, list));
    }
    body.append(selectionPane(list));
    pane.replaceChildren(head, body);
    if (paneFocus) {
      const target = body.querySelector(`[data-fp="${paneFocus}"]`);
      if (target) { target.open = true; requestAnimationFrame(() => target.scrollIntoView({ block: "start" })); }
      paneFocus = null;
    } else body.scrollTop = keepScroll;
  }

  const section = (key, label, open, ...content) => h("details", { class: "fp-sec", "data-fp": key, open: open || null }, h("summary", {}, label), h("div", { class: "fp-sec-body" }, ...content.flat().filter(Boolean)));
  const line = (label, ...controls) => h("div", { class: "fp-line" }, h("span", { class: "fp-label" }, label), h("div", { class: "fp-ctrl" }, ...controls.flat().filter(Boolean)));
  function swatchRow(palette, current, pick, { none = null } = {}) {
    return h("div", { class: "fp-swatches" },
      none ? h("button", { type: "button", class: ["fp-sw", "none", current === "none" ? "on" : ""], title: none, "aria-label": none, onclick: () => pick("none") }, "∅") : null,
      palette.map(([color, label]) => h("button", { type: "button", class: ["fp-sw", String(current).toLowerCase() === color ? "on" : ""], title: `${label}（${color.toUpperCase()}）`, "aria-label": label, style: { "--c": color }, onclick: () => pick(color) })));
  }
  function numberInput(value, onSet, { step = 0.1, min = null, max = null, unit = "", width = 70 } = {}) {
    return h("span", { class: "fp-num" }, h("input", { type: "number", value: value ?? "", step, min, max, style: { width: `${width}px` }, onchange: (event) => { const v = parseFloat(event.target.value); if (Number.isFinite(v)) onSet(v); }, onkeydown: (event) => { if (event.key === "Enter") event.target.blur(); } }), unit ? h("small", {}, unit) : null);
  }
  function slider(value, onSet, { min = 0, max = 100, unit = "%" } = {}) {
    const out = h("small", {}, `${Math.round(value)}${unit}`);
    return h("span", { class: "fp-range" }, h("input", { type: "range", min, max, value: Math.round(value), oninput: (event) => { out.textContent = `${event.target.value}${unit}`; }, onchange: (event) => onSet(Number(event.target.value)) }), out);
  }
  const choice = (options, value, onSet) => h("select", { onchange: (event) => onSet(event.target.value) }, options.map(([v, label]) => h("option", { value: v, selected: String(v) === String(value) || null }, label)));
  const toggle = (label, on, onSet) => h("label", { class: "fp-check" }, h("input", { type: "checkbox", checked: on || null, onchange: (event) => onSet(event.target.checked) }), label);
  const tool = (icon, title, run, pressed = false) => h("button", { type: "button", class: "fp-tool", title, "aria-label": title, "aria-pressed": String(Boolean(pressed)), onclick: run }, ico(icon));
  /** スポイト in the pane (as in the ribbon's colour menus): a colour taken from the slide. */
  const eyedrop = (pick) => tool("eyedrop", "スポイト：スライドの上の色をクリックして取る", async () => { const c = await pickColor({ root: document.getElementById("stageBody") || document.body, toast: app.toast }); if (c) pick(c); });

  function paneSections(chosen, list) {
    const o = E.withDefaults(chosen[0]);
    const single = chosen.length === 1;
    const textual = chosen.some((x) => ["shape", "text"].includes(x.kind));
    const out = [];
    if (textual) {
      const fill = o.fill;
      out.push(section("fill", "塗りつぶし", true,
        swatchRow(E.PALETTE.fill, fill, setFill, { none: "塗りつぶしなし" }),
        line("その他の色", h("input", { type: "color", value: /^#[0-9a-f]{6}$/i.test(fill) ? fill : "#4472c4", onchange: (event) => setFill(event.target.value) }), eyedrop(setFill)),
        line("図", h("button", { type: "button", class: "btn btn-sm fp-picfill", onclick: () => pictureFill() }, o.fillImg ? "図を変更…" : "図で塗りつぶし…"),
          o.fillImg ? h("label", { class: "fp-check" }, h("input", { type: "checkbox", checked: o.fillTile || null, onchange: (event) => editor.apply((x) => (x.fillImg ? { fillTile: event.target.checked || undefined } : null)) }), "並べて表示") : null,
          o.fillImg ? h("button", { type: "button", class: "btn btn-sm btn-ghost", onclick: () => editor.apply((x) => (x.fillImg ? { fillImg: undefined, fillTile: undefined } : null)) }, "外す") : null),
        line("透明度", slider((1 - (o.fillOpacity ?? 1)) * 100, (v) => editor.apply((x) => (["shape", "text"].includes(x.kind) ? { fillOpacity: v ? Math.round((1 - v / 100) * 100) / 100 : undefined } : null))))));
    }
    if (chosen.some((x) => ["shape", "text", "image", "line"].includes(x.kind))) {
      const isLine = chosen.every((x) => x.kind === "line");
      out.push(section("line", isLine ? "線" : "枠線", true,
        swatchRow(E.PALETTE.line, o.stroke, setStroke, { none: isLine ? null : "枠線なし" }),
        line("その他の色", h("input", { type: "color", value: /^#[0-9a-f]{6}$/i.test(o.stroke) ? o.stroke : "#1f3864", onchange: (event) => setStroke(event.target.value) }), eyedrop(setStroke)),
        line("幅", numberInput(ops.toPt(o.strokeW || 2), (v) => setStrokeWidth(Math.max(0.25, v)), { step: 0.25, min: 0.25, unit: "pt" })),
        line("実線/点線", choice(Object.entries(E.DASHES).map(([k, [label]]) => [k, label]), o.dash || "solid", setDash)),
        line("複合線", choice([["", "一重線"], ...Object.entries(E.COMPOUNDS)], o.cmpd || "", (v) => setLineLook({ cmpd: v || undefined }))),
        line("線端", choice(Object.entries(E.LINE_CAPS), o.cap || "flat", (v) => setLineLook({ cap: v === "flat" ? undefined : v }))),
        line("結合点", choice(Object.entries(E.LINE_JOINS), o.join || (o.kind === "line" ? "round" : "miter"), (v) => setLineLook({ join: v }))),
        chosen.some((x) => x.kind === "line") ? [
          line("線の種類", choice(Object.entries(E.ROUTES), o.route || "straight", (v) => editor.apply((x) => (x.kind === "line" ? { route: v === "straight" ? undefined : v } : null)))),
          line("始点の形", choice(Object.entries(E.ARROWHEADS), o.head || "none", (v) => editor.apply((x) => (x.kind === "line" ? { head: v === "none" ? undefined : v } : null)))),
          line("終点の形", choice(Object.entries(E.ARROWHEADS), o.tail || "none", (v) => editor.apply((x) => (x.kind === "line" ? { tail: v === "none" ? undefined : v } : null)))),
          line("矢印のサイズ", choice([[1, "小"], [2, "中"], [3, "大"]], o.tailSize || 2, (v) => editor.apply((x) => (x.kind === "line" ? { headSize: Number(v) === 2 ? undefined : Number(v), tailSize: Number(v) === 2 ? undefined : Number(v) } : null)))),
        ] : null));
    }
    if (chosen.some((x) => x.kind === "icon")) {
      out.push(section("icon", "アイコン", true,
        swatchRow(E.PALETTE.line.filter(([c]) => c !== "#ffffff"), o.color, (c) => editor.apply((x) => (x.kind === "icon" ? { color: c } : null))),
        line("線の太さ", numberInput(o.strokeW, (v) => editor.apply((x) => (x.kind === "icon" ? { strokeW: Math.max(0.5, Math.min(4, v)) } : null)), { step: 0.25, min: 0.5, max: 4 })),
        line("アイコンの変更", h("button", { type: "button", class: "btn btn-sm", onclick: (event) => openPop(event.currentTarget, iconGallery((key) => editor.apply((x) => (x.kind === "icon" ? { icon: key } : null)))) }, "選ぶ…"))));
    }
    if (textual) {
      const st = editor.textState() || {};
      out.push(section("text", "文字", true,
        line("フォント", choice(Object.entries(E.FONTS).map(([k, [label]]) => [k, label]), st.font || "body", (v) => editor.textFormat("font", v))),
        line("サイズ", numberInput(st.fs, (v) => editor.textFormat("size", v), { step: 1, min: 4, unit: "pt" }),
          tool("grow", "大きく", () => editor.textFormat("grow")), tool("shrink", "小さく", () => editor.textFormat("shrink"))),
        line("書式", tool("bold", "太字", () => editor.textFormat("bold"), st.bold), tool("italic", "斜体", () => editor.textFormat("italic"), st.italic), tool("underline", "下線", () => editor.textFormat("underline"), st.underline), tool("strike", "取り消し線", () => editor.textFormat("strike"), st.strike), tool("clear", "書式のクリア", () => editor.textFormat("clear"))),
        line("文字の色", swatchRow(E.PALETTE.text, st.color, (c) => editor.textFormat("color", c))),
        line("揃え", tool("textLeft", "左揃え", () => editor.textFormat("align", "left"), st.align === "left"), tool("textCenter", "中央揃え", () => editor.textFormat("align", "center"), st.align === "center"), tool("textRight", "右揃え", () => editor.textFormat("align", "right"), st.align === "right"), tool("textJustify", "両端揃え", () => editor.textFormat("align", "justify"), st.align === "justify"), tool("textDistributed", "均等割り付け", () => editor.textFormat("align", "distributed"), st.align === "distributed")),
        line("上下の位置", tool("valignTop", "上揃え", () => editor.textFormat("valign", "top"), st.valign === "top"), tool("valignMiddle", "上下中央", () => editor.textFormat("valign", "middle"), st.valign === "middle"), tool("valignBottom", "下揃え", () => editor.textFormat("valign", "bottom"), st.valign === "bottom")),
        line("箇条書き", tool("bullet", "箇条書き", () => editor.textFormat("bullet"), st.list === "bullet"), tool("number", "段落番号", () => editor.textFormat("number"), st.list === "number"), tool("vtext", "縦書き", () => editor.textFormat("vertical"), st.vertical)),
        line("行間", choice(LINE_HEIGHTS.map(([v, label]) => [v, `${label} 行`]), LINE_HEIGHTS.find(([v]) => Math.abs(v - (st.lh || 0)) < 0.01)?.[0] ?? st.lh, (v) => editor.textFormat("lh", Number(v)))),
        line("段落の間隔", choice([[0, "なし"], [0.3, "小"], [0.6, "中"], [1, "大"]], o.psp || 0, (v) => editor.textFormat("psp", Number(v)))),
        line("文字の間隔", choice([[0, "標準"], [-0.05, "狭く"], [0.05, "広く"], [0.12, "より広く"]], o.ls || 0, (v) => editor.textFormat("ls", Number(v)))),
        line("自動調整", choice(Object.entries(E.AUTOFIT), st.autofit || "none", (v) => editor.textFormat("autofit", v))),
        line("折り返し", toggle("図形の幅で折り返す", st.wrap !== false, (on) => editor.textFormat("wrap", on))),
        line("余白（cm）", ...["上", "右", "下", "左"].map((label, i) => h("span", { class: "fp-pad" }, h("small", {}, label), numberInput(ops.toCm((st.pad || [7, 14, 7, 14])[i]), (v) => { const pad = [...(st.pad || [7, 14, 7, 14])]; pad[i] = ops.fromCm(Math.max(0, v)); editor.textFormat("pad", pad); }, { step: 0.05, min: 0, width: 54 }))))));
    }
    if (single && chosen[0].kind === "table") {
      const t = E.withDefaults(chosen[0]);
      const setTable = (patch) => editor.changeTable((o) => ({ ...o, ...patch }));
      out.push(section("table", "表", true,
        line("スタイル", choice(Object.entries(E.TABLE_STYLES), t.style || "medium2", (v) => setTable({ style: v === "medium2" ? undefined : v }))),
        line("オプション", toggle("ヘッダー行", t.header !== false, (on) => setTable({ header: on })), toggle("縞模様（行）", t.banded !== false, (on) => setTable({ banded: on }))),
        line("", toggle("最初の列", t.firstCol, (on) => setTable({ firstCol: on })), toggle("集計行", t.lastRow, (on) => setTable({ lastRow: on }))),
        line("文字のサイズ", numberInput(ops.toPt(t.fs), (v) => editor.textFormat("size", v), { step: 1, min: 6, unit: "pt" })),
        line("フォント", choice(Object.entries(E.FONTS).map(([k, [label]]) => [k, label]), t.font || "body", (v) => editor.textFormat("font", v))),
        line("大きさ", h("span", { class: "hint" }, `${t.cells.length}行 × ${t.cols.length}列（リボンの「レイアウト」で行・列の追加・削除・結合）`))));
    }
    if (single && chosen[0].kind === "chart") {
      const c = chosen[0].chart;
      out.push(section("chart", "グラフ", true,
        line("種類", choice(Object.entries(E.CHART_KINDS), c.type, (v) => editor.apply({ chart: { ...c, type: v } }, { ids: [chosen[0].id] }))),
        line("", h("button", { type: "button", class: "btn btn-sm", onclick: () => tables.editChart(chosen[0].id) }, "データの編集…"), h("small", { class: "hint" }, `${c.labels.length}項目 × ${c.series.length}系列`))));
    }
    if (chosen.some((x) => x.kind === "image")) {
      const img = E.withDefaults(chosen.find((x) => x.kind === "image"));
      const crop = img.crop || { l: 0, t: 0, r: 0, b: 0 };
      const setCrop = (side, v) => editor.apply((x) => {
        if (x.kind !== "image") return null;
        const c = { ...(x.crop || { l: 0, t: 0, r: 0, b: 0 }), [side]: Math.max(0, Math.min(0.9, v / 100)) };
        return { crop: c.l || c.t || c.r || c.b ? c : undefined };
      });
      out.push(section("picture", "図", true,
        line("切り抜く形", choice(MASKS.map((key) => [key, E.SHAPES[key].label.replace(/^.*: /, "")]), img.mask || "rect", (v) => editor.apply((x) => (x.kind === "image" ? { mask: v === "rect" ? undefined : v, adj: undefined } : null)))),
        line("トリミング（%）", ...[["l", "左"], ["t", "上"], ["r", "右"], ["b", "下"]].map(([side, label]) => h("span", { class: "fp-pad" }, h("small", {}, label), numberInput(Math.round(crop[side] * 100), (v) => setCrop(side, v), { step: 1, min: 0, max: 90, width: 54 })))),
        line("はめ込み方", choice(Object.entries(E.FITS), img.fit || "fill", (v) => editor.apply((x) => (x.kind === "image" ? { fit: v } : null)))),
        line("明るさ", slider((img.bright || 0) * 100 + 100, (v) => editor.apply((x) => (x.kind === "image" ? { bright: (v - 100) / 100 || undefined } : null)), { min: 0, max: 200 })),
        line("コントラスト", slider((img.contrast || 0) * 100 + 100, (v) => editor.apply((x) => (x.kind === "image" ? { contrast: (v - 100) / 100 || undefined } : null)), { min: 0, max: 200 })),
        line("彩度", slider((img.sat || 0) * 100 + 100, (v) => editor.apply((x) => (x.kind === "image" ? { sat: (v - 100) / 100 || undefined } : null)), { min: 0, max: 200 })),
        line("色", toggle("グレースケール", img.gray, (on) => editor.apply((x) => (x.kind === "image" ? { gray: on || undefined } : null)))),
        line("代替テキスト", h("input", { type: "text", value: img.alt || "", placeholder: "画像の説明（読み上げ用）", onchange: (event) => editor.apply((x) => (x.kind === "image" ? { alt: event.target.value.trim() || undefined } : null)) })),
        line("", h("button", { type: "button", class: "btn btn-sm", onclick: async () => { const [file] = await app.pickFiles("image/*", false); if (file) await app.replaceImage(img.id, file); } }, "図の変更…"),
          h("button", { type: "button", class: "btn btn-sm", onclick: () => editor.apply((x) => (x.kind === "image" ? { crop: undefined, mask: undefined, adj: undefined, bright: undefined, contrast: undefined, sat: undefined, gray: undefined } : null)) }, "リセット"))));
    }
    if (chosen.some((x) => x.kind === "video" || x.kind === "lottie")) {
      const media = E.withDefaults(chosen.find((x) => x.kind === "video" || x.kind === "lottie"));
      out.push(section("media", media.kind === "lottie" ? "アニメーション" : "ビデオ", true,
        line("再生", toggle("発表で自動再生", media.autoplay !== false, (on) => editor.apply((x) => (["video", "lottie"].includes(x.kind) ? { autoplay: on } : null)))),
        line("", toggle("繰り返す", media.loop !== false, (on) => editor.apply((x) => (["video", "lottie"].includes(x.kind) ? { loop: on } : null)))),
        media.kind === "video" ? line("", toggle("音を出さない", media.muted !== false, (on) => editor.apply((x) => (x.kind === "video" ? { muted: on } : null)))) : null,
        line("はめ込み方", choice([["cover", "トリミングして埋める"], ["contain", "全体を入れる"]], media.fit || "cover", (v) => editor.apply((x) => (["video", "lottie"].includes(x.kind) ? { fit: v } : null))))));
    }
    if (chosen.some((x) => x.kind === "audio")) {
      const sound = E.withDefaults(chosen.find((x) => x.kind === "audio"));
      out.push(section("audio", "オーディオ", true,
        line("ファイル", h("span", { class: "fp-file" }, sound.fileName || "（名前なし）")),
        line("アイコンの色", swatchRow(E.PALETTE.line.filter(([c]) => c !== "#ffffff"), sound.color, (c) => editor.apply((x) => (x.kind === "audio" ? { color: c } : null)))),
        line("再生", toggle("発表で自動再生", Boolean(sound.autoplay), (on) => editor.apply((x) => (x.kind === "audio" ? { autoplay: on } : null)))),
        line("", toggle("繰り返す", Boolean(sound.loop), (on) => editor.apply((x) => (x.kind === "audio" ? { loop: on } : null)))),
        line("", h("button", { type: "button", class: "btn btn-sm", onclick: () => showTab("playback", { open: true }) }, "再生タブ（トリミング・フェード・音量）"))));
    }
    if (single && chosen[0].kind === "smartart") {
      const sa = chosen[0].smartart;
      const setSa = (patch) => editor.apply((x) => (x.kind === "smartart" ? { smartart: { ...x.smartart, ...patch } } : null));
      out.push(section("smartart", "SmartArt", true,
        line("レイアウト", choice(Object.entries(E.SMARTART_LAYOUTS).map(([k, v]) => [k, `${v.group}：${v.label}`]), sa.layout, (v) => setSa({ layout: v }))),
        line("色", choice(Object.entries(E.SMARTART_COLORS), sa.color || "blue", (v) => setSa({ color: v }))),
        line("角", choice(Object.entries(E.SMARTART_STYLES), sa.style || "round", (v) => setSa({ style: v }))),
        line("項目", h("span", { class: "fp-file" }, `${sa.items.length}項目（${sa.items.filter((it) => it.level === 0).length}つの大項目）`)),
        line("", h("button", { type: "button", class: "btn btn-sm", onclick: () => smart.openPane(chosen[0].id) }, "テキスト ウィンドウ"),
          h("button", { type: "button", class: "btn btn-sm", onclick: () => smart.toShapes(chosen[0].id) }, "図形に変換"))));
    }
    // Size and position (cm, degrees), as PowerPoint's 配置とサイズ.
    if (single) out.push(sizeSection(chosen[0], list));
    out.push(section("opacity", "透明度", false, line("透明度", slider((1 - (o.opacity ?? 1)) * 100, (v) => editor.apply({ opacity: v ? Math.round((1 - v / 100) * 100) / 100 : undefined })))));
    out.push(section("arrange", "配置", false,
      h("div", { class: "fp-tools" },
        tool("front", "最前面へ", () => editor.order("front")), tool("back", "最背面へ", () => editor.order("back")),
        tool("alignLeft", "左揃え", () => editor.alignSelection("left")), tool("alignCenter", "左右中央", () => editor.alignSelection("center")), tool("alignRight", "右揃え", () => editor.alignSelection("right")),
        tool("alignTop", "上揃え", () => editor.alignSelection("top")), tool("alignMiddle", "上下中央", () => editor.alignSelection("middle")), tool("alignBottom", "下揃え", () => editor.alignSelection("bottom")),
        tool("distH", "左右に整列", () => editor.distributeSelection("x")), tool("distV", "上下に整列", () => editor.distributeSelection("y")),
        tool("group", "グループ化", () => editor.groupSelection()), tool("ungroup", "グループ解除", () => editor.ungroupSelection()),
        tool("flipH", "左右反転", () => editor.flipSelection("x")), tool("flipV", "上下反転", () => editor.flipSelection("y")),
        tool(chosen.some((x) => x.locked) ? "unlock" : "lock", chosen.some((x) => x.locked) ? "ロックを解除" : "ロック", () => editor.setLocked(!chosen.some((x) => x.locked))),
        tool("duplicate", "複製（⌘D）", () => editor.duplicateSelection()), tool("trash", "削除", () => editor.removeSelection()))));
    if (single) out.push(section("action", "リンク・動作（発表中のクリック・マウスの通過）", Boolean(chosen[0].action || chosen[0].overAction), linkForm(chosen[0])));
    return out;
  }

  function sizeSection(o, list) {
    if (o.kind === "line") {
      const [[x1, y1], [x2, y2]] = E.lineEnds(o, list);
      const set = (patch) => editor.apply({ ...patch, from: undefined, to: undefined, x1, y1, x2, y2, ...patch }, { ids: [o.id] });
      return section("size", "サイズと位置", true,
        line("始点（横・縦）", numberInput(ops.toCm(x1), (v) => set({ x1: ops.fromCm(v) }), { step: 0.1, unit: "cm" }), numberInput(ops.toCm(y1), (v) => set({ y1: ops.fromCm(v) }), { step: 0.1, unit: "cm" })),
        line("終点（横・縦）", numberInput(ops.toCm(x2), (v) => set({ x2: ops.fromCm(v) }), { step: 0.1, unit: "cm" }), numberInput(ops.toCm(y2), (v) => set({ y2: ops.fromCm(v) }), { step: 0.1, unit: "cm" })),
        line("長さ", h("span", { class: "hint" }, `${ops.toCm(Math.hypot(x2 - x1, y2 - y1))} cm`)));
    }
    const ratio = o.w / o.h;
    const lock = o.lockRatio || (o.kind === "image" && o.lockRatio !== false);
    return section("size", "サイズと位置", true,
      line("高さ", numberInput(ops.toCm(o.h), (v) => { const hh = ops.fromCm(Math.max(0.05, v)); editor.apply(lock ? { h: hh, w: hh * ratio } : { h: hh }, { ids: [o.id] }); }, { step: 0.1, unit: "cm" })),
      line("幅", numberInput(ops.toCm(o.w), (v) => { const w = ops.fromCm(Math.max(0.05, v)); editor.apply(lock ? { w, h: w / ratio } : { w }, { ids: [o.id] }); }, { step: 0.1, unit: "cm" })),
      line("", toggle("縦横比を固定する", lock, (on) => editor.apply({ lockRatio: o.kind === "image" ? (on ? undefined : false) : on || undefined }, { ids: [o.id] }))),
      line("回転", numberInput(o.rot || 0, (v) => editor.apply({ rot: ops.normAngle(v) || undefined }, { ids: [o.id] }), { step: 1, unit: "°" })),
      line("横位置", numberInput(ops.toCm(o.x), (v) => editor.apply({ x: ops.fromCm(v) }, { ids: [o.id] }), { step: 0.1, unit: "cm" }), h("small", { class: "hint" }, "左上から")),
      line("縦位置", numberInput(ops.toCm(o.y), (v) => editor.apply({ y: ops.fromCm(v) }, { ids: [o.id] }), { step: 0.1, unit: "cm" })));
  }

  /** 選択ウィンドウ: every object, front first; click to select, eye to hide, drag to change the order. */
  function selectionPane(list) {
    const rows = [...list].map((o, i) => ({ o, i })).reverse();
    const box = h("details", { class: "fp-sec fp-selection", "data-fp": "selection", open: true },
      h("summary", {}, `選択ウィンドウ（${list.length}）`),
      h("div", { class: "fp-sec-body" },
        list.length ? h("div", { class: "fp-sel-actions" },
          h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: () => editor.setHidden(list.map((o) => o.id), false) }, "すべて表示"),
          h("button", { type: "button", class: "btn btn-ghost btn-sm", onclick: () => editor.setHidden(list.map((o) => o.id), true) }, "すべて非表示")) : h("p", { class: "hint" }, "このスライドにはまだオブジェクトがありません。「挿入」から図形・文字・画像を置けます。"),
        h("ol", { class: "fp-sel-list" }, rows.map(({ o, i }) => {
          const item = h("li", { class: [editor.selection.includes(o.id) ? "on" : "", o.hidden ? "is-hidden" : ""], draggable: "true", "data-id": o.id,
            onclick: (event) => { if (event.target.closest("button")) return; editor.select(event.shiftKey || event.metaKey || event.ctrlKey ? [...new Set([...editor.selection, o.id])] : [o.id]); },
            ondblclick: async () => { const name = await app.ask("オブジェクトの名前", "", E.objectName(o, i)); if (name != null) editor.rename(o.id, name); },
            ondragstart: (event) => { event.dataTransfer.setData("text/x-hpptx-object", o.id); event.dataTransfer.effectAllowed = "move"; },
            ondragover: (event) => { if ([...event.dataTransfer.types].includes("text/x-hpptx-object")) { event.preventDefault(); item.classList.add("drop"); } },
            ondragleave: () => item.classList.remove("drop"),
            ondrop: (event) => { event.preventDefault(); item.classList.remove("drop"); const id = event.dataTransfer.getData("text/x-hpptx-object"); if (id && id !== o.id) editor.moveInOrder(id, list.findIndex((x) => x.id === o.id)); } },
          h("span", { class: "fp-sel-kind" }, kindIcon(o)),
          h("span", { class: "fp-sel-name" }, E.objectName(o, i), o.group ? h("small", {}, " ・グループ") : null),
          o.locked ? h("span", { class: "fp-sel-lock", title: "ロック中" }, ico("lock", 14)) : null,
          h("button", { type: "button", class: "fp-sel-eye", title: o.hidden ? "表示する" : "非表示にする", "aria-label": o.hidden ? "表示する" : "非表示にする", onclick: () => editor.setHidden([o.id], !o.hidden) }, ico(o.hidden ? "eyeOff" : "eye", 15)));
          return item;
        })),
        list.length > 1 ? h("p", { class: "hint" }, "上ほど前面です。ドラッグで順番を変え、ダブルクリックで名前を変えられます。") : null));
    return box;
  }
  function kindIcon(o) {
    if (o.kind === "shape") return E.SHAPES[o.shape] ? shapeThumb(o.shape, 18, 14) : freeformThumb(o.path?.closed ? "polygon" : "curve");
    return ico({ text: "textbox", image: "image", line: "line", icon: "icon", video: "video", audio: "audio", lottie: "lottie", table: "table", chart: "chartBar", smartart: "smartart", ink: "pen", zoom: "zoomSlide", camera: "camera", equation: "equation", model: "model3d", poll: "poll" }[o.kind] || "shapes", 16);
  }

  // アニメーション・画面切り替え (anim.mjs) build their tabs and the animation pane with these same parts.
  const anim = createAnimations(editor, app, { btn, drop, group, col, row, menu, openPop, closePop, updater: (fn) => updaters.push(fn), tabNow: () => tab, refreshRibbon: () => refresh() });
  // インタラクション (interact.mjs): what only HTML does with an object — the mouse, a click, motion that keeps going.
  const ix = createInteractions(editor, app, { btn, drop, group, col, row, menu, openPop, closePop, updater: (fn) => updaters.push(fn), tabNow: () => tab, editLink: () => editLink() });
  const crop = createCrop(editor, app);
  const freeform = createFreeform(editor, app);
  // SmartArt (smartart.mjs): the gallery, the text pane and the SmartArt のデザイン tab.
  const smart = createSmartArt(editor, app, { btn, drop, group, col, row, menu, openPop, closePop, updater: (fn) => updaters.push(fn), showTab: (id) => showTab(id) });
  // 挿入 → オーディオ・画面録画 and the 再生 tab of a video or a sound (media.mjs).
  const media = createMedia(editor, app, { btn, drop, group, col, row, menu, openPop, closePop, updater: (fn) => updaters.push(fn), refreshRibbon: () => refresh() });
  const tables = createTableUi(editor, app, { btn, drop, group, col, row, menu, openPop, closePop, updater: (fn) => updaters.push(fn), colors, showTab: (id) => showTab(id) });
  // 描画 (ink.mjs): pens, highlighters, the eraser and the lasso, ink to shapes, ink replay.
  const ink = createInk(editor, app, { btn, drop, group, col, row, menu, openPop, closePop, updater: (fn) => updaters.push(fn), refreshRibbon: () => refresh() });
  // ズーム・カメオ・スクリーンショット・数式・ワードアート・フィールド・段組み・文字種・音声入力 (extras.mjs).
  const extras = createExtras(editor, app, { btn, drop, group, col, row, menu, openPop, closePop, updater: (fn) => updaters.push(fn), refreshRibbon: () => refresh() });
  // 図の形式 → 背景の削除・透明色・図の圧縮 (picture.mjs).
  const pictureTools = createPictureTools(editor, app);
  // 挿入 → 3D モデル and the 3D モデル tab (model3d.mjs).
  const models3d = createModels(editor, app, { btn, drop, group, col, row, menu, openPop, closePop, updater: (fn) => updaters.push(fn), refreshRibbon: () => refresh() });
  // 挿入 → アンケート (poll.mjs).
  const polls = createPolls(editor, app);

  // 校閲 → 音声読み上げ (readaloud.mjs).
  const readAloud = createReadAloud({ h, toast: (message) => app.toast(message), refreshRibbon: () => refresh(), selectedIds: () => selected().map((o) => o.id), index: () => app.index() });
  // 検索 (⌥Q) and the クイック アクセス ツール バー (quick.mjs).
  quick = createQuick({ h, ico, app, openPop, closePop, menu, commands: ribbonCommands, run: runRibbonCommand, render: () => renderRibbon(true), tab: () => tab, watch: (fn) => updaters.push(fn) });
  // A right-click on a ribbon button (in a folded group too) offers to add it to the toolbar.
  ribbon?.addEventListener("contextmenu", (event) => quick.contextMenu(event, tab));
  document.addEventListener("contextmenu", (event) => { if (event.target.closest?.(".rb-pop .rb-fold")) quick.contextMenu(event, tab); });

  editor.subscribe(() => { renderRibbon(); renderPane(); });

  return {
    renderRibbon, renderPane, renderAnimPane: () => anim.renderPane(), editChart: (id) => tables.editChart(id),
    editSmartart: (id, item) => { editor.select([id]); showTab("smartartDesign"); smart.openPane(id, item); }, smartartToShapes: (id) => smart.toShapes(id), startCrop: (id) => crop.start(id), editPoints: (id) => freeform.editPoints(id),
    showTab, closePop, openPop, shapeGallery, iconGallery, photoGallery, toggleRibbon, setRibbonMode, zoomMenu,
    popMenu: (anchor, items) => openPop(anchor, menu(items)),
    editEquation: (id) => { const o = editor.objects().find((x) => x.id === id); if (o?.kind === "equation") extras.equationDialog(o); },
    editPoll: (id) => polls.edit(id),
    readAloud: () => readAloud.toggle(), stopReading: () => readAloud.stop(),
    fontDialog, paragraphDialog, gridDialog,
    ribbonCommands, runRibbonCommand, focusSearch: () => { if (ribbonMode === "auto" && !floating) openFloating(); quick.focusSearch(); }, customizeQat: () => quick.customize(),
    stopInk: () => ink.stop(), get inkTool() { return ink.tool; },
    get tab() { return tab; }, get ribbonMode() { return ribbonMode; }, get floating() { return floating; },
  };
}
