// アクセシビリティ チェック (PowerPoint's Accessibility Checker): what makes a deck hard to use with a screen reader or
// hard to read — pictures and charts without alternative text, slides without a title, tables without a header row,
// links whose words say nothing, text with too little contrast, a reading order that jumps around the slide, the
// same title on several slides — each with a way to fix it. Also the 読み取り順序 pane (slide.readingOrder) and the
// 代替テキスト dialog (alt, or 装飾用 so a screen reader skips the object).

import { chromeOf, signature } from "./htmlfx.mjs";

const NEEDS_ALT = new Set(["image", "video", "chart", "smartart", "icon", "ink", "lottie", "camera", "zoom", "model"]);
const KIND_LABEL = { image: "図", video: "ビデオ", chart: "グラフ", smartart: "SmartArt", icon: "アイコン", ink: "インク", lottie: "アニメーション", camera: "カメオ", zoom: "ズーム", shape: "図形", text: "テキスト", table: "表", equation: "数式", audio: "オーディオ", line: "線" };
const VAGUE_LINK = /^(こちら|ここ|これ|クリック|ここをクリック|こちらをクリック|詳細|詳しく|リンク|link|here|click|click here|more|read more)$/i;

const strip = (html) => String(html ?? "").replace(/<[^>]*>/g, " ").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/\s+/g, " ").trim();

/** WCAG contrast between two colours (#rrggbb). */
export function contrast(a, b) {
  const lum = (hex) => {
    const [r, g, bl] = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255).map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
    return 0.2126 * r + 0.7152 * g + 0.0722 * bl;
  };
  const [x, y] = [lum(a), lum(b)].sort((m, n) => n - m);
  return (x + 0.05) / (y + 0.05);
}
/** A colour laid over white with some opacity. */
const overWhite = (hex, alpha = 1) => `#${[1, 3, 5].map((i) => Math.round(parseInt(hex.slice(i, i + 2), 16) * alpha + 255 * (1 - alpha)).toString(16).padStart(2, "0")).join("")}`;

/** Objects in the order they appear to the eye: rows from the top (boxes that overlap in height share a row), then left to right. */
export function visualOrder(objects) {
  const items = objects.map((o) => {
    const x = o.kind === "line" ? Math.min(o.x1, o.x2) : o.x;
    const y = o.kind === "line" ? Math.min(o.y1, o.y2) : o.y;
    const hh = o.kind === "line" ? Math.abs(o.y2 - o.y1) : o.h;
    return { o, x, y, mid: y + hh / 2, h: Math.max(1, hh) };
  }).sort((a, b) => a.y - b.y);
  const rows = [];
  for (const it of items) {
    const row = rows.find((r) => it.mid > r.top && it.mid < r.bottom);
    if (row) { row.items.push(it); row.bottom = Math.max(row.bottom, it.y + it.h * 0.5); } else rows.push({ top: it.y - 4, bottom: it.y + it.h, items: [it] });
  }
  return rows.flatMap((r) => r.items.sort((a, b) => a.x - b.x).map((it) => it.o));
}

/** What a screen reader reads on a slide, in its reading order (hidden and decorative objects are skipped). */
export function readable(slide, E) {
  const list = (slide?.elements || []).filter((o) => !o.hidden && !o.decorative);
  const order = E.readingOrderOf ? E.readingOrderOf(slide) : list.map((o) => o.id);
  const byId = new Map(list.map((o) => [o.id, o]));
  return order.map((id) => byId.get(id)).filter(Boolean);
}
const hasContent = (o, E) => (["shape", "text"].includes(o.kind) ? Boolean(strip(o.text)) : o.kind !== "line" && (o.kind === "table" || o.kind === "equation" || Boolean(o.alt)));

/**
 * The issues of a deck: [{ id, level: "error" | "warning" | "tip", rule, slide, el?, message, hint }], in slide order.
 * `E` is the engine (withDefaults, storyMap, readingOrderOf).
 */
export function accessibilityIssues(deck, E) {
  const slides = deck?.slides || [];
  const story = E.storyMap ? E.storyMap(slides) : { parent: [] };
  const chrome = chromeOf(slides);
  const out = [];
  const add = (level, rule, slide, message, hint, el = null) => out.push({ id: `${rule}:${slide}:${el || ""}`, level, rule, slide, el, message, hint });
  const titles = new Map();
  slides.forEach((slide, i) => {
    if (slide.hidden) return;
    // スライド タイトル
    const title = strip(slide.title);
    if (!title) add("error", "title", i, "スライド タイトルがありません", "タイトルは目次と画面読み上げで使われます。表示しないタイトルでも付けられます");
    else if (story.parent?.[i] == null) titles.set(title, [...(titles.get(title) || []), i]);
    const objects = (slide.elements || []).filter((o) => o && !o.hidden);
    for (const o of objects) {
      if (o.decorative || chrome.has(signature(o))) continue;
      const name = o.name || KIND_LABEL[o.kind] || "オブジェクト";
      // 代替テキスト
      const needsAlt = NEEDS_ALT.has(o.kind) || (o.kind === "shape" && !strip(o.text) && o.fill && o.fill !== "none" && !(o.w < 24 || o.h < 24));
      if (needsAlt && !o.alt) add("error", "alt", i, `代替テキストがありません（${name}）`, "内容を短い文で説明するか、飾りなら「装飾用にする」", o.id);
      // 表の見出し行
      if (o.kind === "table" && o.header === false) add("error", "header", i, "表に見出し行がありません", "テーブル デザイン →「見出し行」をオン（1行目を見出しにする）", o.id);
      // リンクの文字
      if (["shape", "text"].includes(o.kind) && /<a\s/i.test(o.text || "")) {
        for (const [, href, words] of String(o.text).matchAll(/<a\s[^>]*href="([^"]*)"[^>]*>(.*?)<\/a>/gis)) {
          const text = strip(words);
          if (VAGUE_LINK.test(text) || text === href || /^https?:\/\//i.test(text)) add("warning", "link", i, `リンクの文字から行き先が分かりません（「${text.slice(0, 24)}」）`, "行き先が分かる言葉にする（例：「売上の詳細（社内ポータル）」）", o.id);
        }
      }
      // 文字のコントラスト
      if (["shape", "text"].includes(o.kind) && strip(o.text) && E.withDefaults) {
        const d = E.withDefaults(o);
        const fill = d.fill && d.fill !== "none" && /^#[0-9a-f]{6}$/i.test(d.fill) ? overWhite(d.fill, (d.fillOpacity ?? 1) * (d.opacity ?? 1)) : "#ffffff";
        const colors = new Set([/^#[0-9a-f]{6}$/i.test(d.color || "") ? d.color : "#1a1a1a"]);
        for (const [, c] of String(o.text).matchAll(/color:\s*(#[0-9a-f]{6})/gi)) colors.add(c);
        const large = (d.fs || 32) >= 36 || (d.bold && (d.fs || 32) >= 28);
        const worst = Math.min(...[...colors].map((c) => contrast(c.toLowerCase(), fill)));
        if (worst < (large ? 3 : 4.5)) add("warning", "contrast", i, `文字が読みにくい色の組み合わせです（コントラスト ${worst.toFixed(1)}:1）`, `${large ? "3" : "4.5"}:1 以上に。文字を黒か濃紺に、または塗りを淡い色に`, o.id);
      }
    }
    // 読み取り順序
    const read = readable(slide, E).filter((o) => hasContent(o, E) && !chrome.has(signature(o)));
    if (read.length >= 3) {
      const eye = visualOrder(read).map((o) => o.id);
      const ids = read.map((o) => o.id);
      if (eye.join() !== ids.join()) add("warning", "order", i, "読み取り順序を確認してください", "画面読み上げは見た目と違う順に読みます。「見た目の順にそろえる」か、読み取り順序で並べ替え");
    }
    // ビデオの内容
    if (objects.some((o) => o.kind === "video" && !o.decorative && !o.captions)) add("tip", "media", i, "ビデオに字幕（キャプション）がありません", "再生タブの「キャプションの挿入」で字幕ファイル（.vtt・.srt）を付けるか、ノートや文字で内容も伝えましょう");
  });
  for (const [title, at] of titles) if (at.length > 1) for (const i of at) add("warning", "duplicate", i, `同じタイトルのスライドがあります（「${title.slice(0, 20)}」：${at.map((n) => n + 1).join("・")}枚目）`, "見分けられるタイトルにすると、目次と読み上げで迷いません");
  const rank = { error: 0, warning: 1, tip: 2 };
  return out.sort((a, b) => rank[a.level] - rank[b.level] || a.slide - b.slide);
}

const RULE_HEADS = { alt: "代替テキストがありません", title: "スライド タイトルがありません", header: "表に見出し行がありません", link: "リンクの文字が分かりにくい", contrast: "読みにくい文字のコントラスト", order: "読み取り順序を確認", duplicate: "重複するスライド タイトル", media: "ビデオの内容" };
const LEVEL_HEADS = { error: "エラー", warning: "警告", tip: "ヒント" };

export function createA11y(app) {
  const { h, E } = app;
  const pane = () => document.getElementById("a11yPane");
  const open = new Set(["error", "warning"]);

  function render() {
    const root = pane();
    if (!root || root.hidden) return;
    const deck = app.deck();
    if (!deck) { root.replaceChildren(); return; }
    const issues = accessibilityIssues(deck, E);
    const groups = ["error", "warning", "tip"].map((level) => [level, issues.filter((x) => x.level === level)]);
    const head = h("div", { class: "a11y-head" },
      h("b", {}, "アクセシビリティ"),
      h("span", { class: ["a11y-sum", issues.length ? "" : "ok"] }, issues.length ? `エラー ${groups[0][1].length}・警告 ${groups[1][1].length}・ヒント ${groups[2][1].length}` : "問題は見つかりませんでした"),
      h("button", { type: "button", class: "btn btn-ghost btn-sm", title: "もう一度確かめる", onclick: render }, "再チェック"));
    const body = h("div", { class: "a11y-list" });
    for (const [level, list] of groups) {
      if (!list.length) continue;
      const byRule = new Map();
      for (const x of list) byRule.set(x.rule, [...(byRule.get(x.rule) || []), x]);
      const sec = h("details", { class: ["a11y-level", level], open: open.has(level) || null, ontoggle: (e) => { if (e.target.open) open.add(level); else open.delete(level); } },
        h("summary", {}, `${LEVEL_HEADS[level]}（${list.length}）`),
        [...byRule].map(([rule, items]) => h("div", { class: "a11y-rule" }, h("div", { class: "a11y-rule-head" }, RULE_HEADS[rule] || rule),
          items.map((x) => h("div", { class: "a11y-item", "data-rule": x.rule, "data-slide": String(x.slide) },
            h("button", { type: "button", class: "a11y-go", title: "その場所へ移動", onclick: () => go(x) }, h("span", { class: "a11y-n" }, `${x.slide + 1}`), h("span", {}, x.message)),
            h("small", { class: "hint" }, x.hint),
            fixes(x))))));
      body.append(sec);
    }
    root.replaceChildren(head, body, readingOrderBox(), h("p", { class: "hint a11y-foot" }, "部品の右クリック →「代替テキストを編集」でも説明を付けられます。飾りの部品は「装飾用にする」と読み上げで飛ばされます。"));
  }
  function go(x) {
    app.select(x.slide);
    if (x.el) requestAnimationFrame(() => app.selectObjects([x.el]));
  }
  function fixes(x) {
    const b = (label, run, cls = "") => h("button", { type: "button", class: ["btn", "btn-xs", cls], onclick: run }, label);
    if (x.rule === "alt") return h("div", { class: "a11y-fix" }, b("説明を追加", () => { go(x); editAlt(x.slide, x.el); }, "btn-primary"), b("装飾用にする", () => { app.setObjectFields(x.slide, x.el, { decorative: true, alt: undefined }); render(); }));
    if (x.rule === "title") return h("div", { class: "a11y-fix" }, b("タイトルを付ける", async () => {
      go(x);
      const text = await app.ask("スライド タイトル", "このスライドの内容が分かるタイトル（白紙のスライドでは表示しないタイトルにもできます）", "");
      if (text?.trim()) { app.setSlideTitle(x.slide, text.trim()); render(); }
    }, "btn-primary"));
    if (x.rule === "header") return h("div", { class: "a11y-fix" }, b("見出し行をオン", () => { app.setObjectFields(x.slide, x.el, { header: true }); render(); }, "btn-primary"));
    if (x.rule === "order") return h("div", { class: "a11y-fix" }, b("見た目の順にそろえる", () => { setVisualOrder(x.slide); render(); }, "btn-primary"), b("読み取り順序…", () => { go(x); requestAnimationFrame(() => pane()?.querySelector(".a11y-order")?.scrollIntoView({ block: "nearest" })); }));
    if (x.rule === "contrast" || x.rule === "link") return h("div", { class: "a11y-fix" }, b("選んで直す", () => go(x)));
    return null;
  }

  /** 読み取り順序 of the slide on the stage: the order a screen reader reads its objects, moved with ↑ ↓. */
  function readingOrderBox() {
    const slide = app.slide();
    const box = h("details", { class: "a11y-order", open: true }, h("summary", {}, "読み取り順序（このスライド）"));
    if (!slide?.elements?.length) { box.append(h("p", { class: "hint" }, "このスライドに部品はありません。")); return box; }
    const order = E.readingOrderOf(slide);
    const byId = new Map(slide.elements.map((o) => [o.id, o]));
    const move = (k, d) => { const next = [...order]; const j = k + d; if (j < 0 || j >= next.length) return; [next[k], next[j]] = [next[j], next[k]]; app.setReadingOrder(app.index(), next); render(); };
    box.append(h("ol", { class: "a11y-read" }, order.map((id, k) => {
      const o = byId.get(id);
      if (!o) return null;
      const words = strip(o.text).slice(0, 26) || o.alt?.slice(0, 26) || "";
      return h("li", { class: [o.decorative ? "decorative" : "", o.hidden ? "hidden-obj" : ""], "data-el": id },
        h("button", { type: "button", class: "a11y-pick", onclick: () => app.selectObjects([id]) }, h("b", {}, o.name || KIND_LABEL[o.kind] || o.kind), words ? h("span", {}, ` ${words}`) : null, o.decorative ? h("small", {}, "装飾用") : null),
        h("button", { type: "button", class: "btn btn-xs", title: "前に読む", "aria-label": "前に読む", disabled: k === 0 || null, onclick: () => move(k, -1) }, "↑"),
        h("button", { type: "button", class: "btn btn-xs", title: "後に読む", "aria-label": "後に読む", disabled: k === order.length - 1 || null, onclick: () => move(k, 1) }, "↓"));
    })), h("div", { class: "a11y-fix" },
      h("button", { type: "button", class: "btn btn-xs", onclick: () => { setVisualOrder(app.index()); render(); } }, "見た目の順にそろえる"),
      h("button", { type: "button", class: "btn btn-xs", disabled: !slide.readingOrder || null, onclick: () => { app.setReadingOrder(app.index(), null); render(); } }, "重なり順に戻す")));
    return box;
  }
  function setVisualOrder(i) {
    const slide = app.deck().slides[i];
    const list = (slide?.elements || []).filter((o) => !o.hidden);
    const eye = visualOrder(list).map((o) => o.id);
    const rest = (slide?.elements || []).map((o) => o.id).filter((id) => !eye.includes(id));
    app.setReadingOrder(i, [...eye, ...rest]);
    app.toast("読み取り順序を、上から下・左から右の順にしました（⌘Zで戻せます）");
  }

  /** 代替テキスト: a description a screen reader says, or 装飾用 (skipped). */
  function editAlt(slideIndex, id) {
    const slide = app.deck()?.slides[slideIndex];
    const o = slide?.elements?.find((x) => x.id === id);
    if (!o) return;
    const text = h("textarea", { class: "alt-input", rows: 4, maxlength: "500", placeholder: "例：2026年度の店舗別売上。東日本が前年比112%で最も伸びている", "aria-label": "代替テキスト" }, o.alt || "");
    const deco = h("input", { type: "checkbox", checked: o.decorative || null });
    const sync = () => { text.disabled = deco.checked; };
    deco.addEventListener("change", sync);
    const dialog = h("dialog", { class: "alt-dialog", "aria-label": "代替テキスト" },
      h("div", { class: "dialog-head" }, h("h3", {}, "代替テキスト"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body" }, h("p", { class: "hint" }, `${o.name || KIND_LABEL[o.kind] || "オブジェクト"}の内容を、見えない人に伝える1〜2文で書きます。`), text,
        h("label", { class: "sh-choice" }, deco, h("span", {}, "装飾用にする（線や飾りなど、内容のない部品。画面読み上げは飛ばします）"))),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary alt-ok", onclick: () => {
          app.setObjectFields(slideIndex, id, deco.checked ? { decorative: true, alt: undefined } : { decorative: undefined, alt: text.value.trim() || undefined });
          dialog.close();
          render();
        } }, "OK")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
    sync();
    text.focus();
  }

  return { render, editAlt, issues: () => accessibilityIssues(app.deck(), E) };
}
