// スライド ショーの設定 and 目的別スライド ショー (PowerPoint's Set Up Slide Show and Custom Shows): how the show runs
// (a kiosk that loops by itself, looping until Esc, without animations, with or without the saved timings and
// narration, the pen colour, live subtitles) and which slides it plays (all, a range, or a named list of slides in an
// order of its own). Stored on the deck as `show` and `customShows` (slides by their sid, so moving or inserting
// slides keeps the lists right).

export const PEN_COLORS = [["#ff0000", "赤"], ["#0070c0", "青"], ["#000000", "黒"]];
export const CAPTION_LANGS = [["ja-JP", "日本語"], ["en-US", "英語"], ["zh-CN", "中国語"], ["ko-KR", "韓国語"]];
const SID = /^[A-Za-z0-9_-]{1,32}$/;
const clean = (text, max) => String(text ?? "").replace(/[\u0000-\u001f<>]/g, "").trim().slice(0, max);

/** The deck's show settings, checked (null when everything is as PowerPoint starts: all slides, by hand). */
export function showOf(value) {
  const v = value && typeof value === "object" ? value : {};
  const out = {};
  for (const key of ["kiosk", "loop", "noAnimation", "noNarration", "captions", "noControls"]) if (v[key] === true) out[key] = true;
  if (v.useTimings === false) out.useTimings = false;
  if (CAPTION_LANGS.some(([k]) => k === v.captionLang) && v.captionLang !== "ja-JP") out.captionLang = v.captionLang;
  if (PEN_COLORS.some(([c]) => c === v.penColor) && v.penColor !== "#ff0000") out.penColor = v.penColor;
  const secs = Number(v.kioskSeconds);
  if (Number.isFinite(secs) && secs >= 2 && secs <= 600) out.kioskSeconds = Math.round(secs);
  const from = Number(v.range?.from);
  const to = Number(v.range?.to);
  if (Number.isInteger(from) && Number.isInteger(to) && from >= 1 && to >= from && to <= 500) out.range = { from, to };
  if (typeof v.custom === "string" && SID.test(v.custom)) out.custom = v.custom;
  return Object.keys(out).length ? out : null;
}

/** The deck's custom shows, checked: [{ id, name, sids }] (null when there are none). */
export function customShowsOf(value) {
  if (!Array.isArray(value)) return null;
  const seen = new Set();
  const list = [];
  for (const cs of value) {
    if (!cs || typeof cs !== "object" || typeof cs.id !== "string" || !SID.test(cs.id) || seen.has(cs.id)) continue;
    const sids = [...new Set((Array.isArray(cs.sids) ? cs.sids : []).filter((sid) => typeof sid === "string" && SID.test(sid)))].slice(0, 500);
    if (!sids.length) continue;
    seen.add(cs.id);
    list.push({ id: cs.id, name: clean(cs.name, 60) || "目的別スライド ショー", sids });
    if (list.length >= 30) break;
  }
  return list.length ? list : null;
}

/** Which slides the show plays (their places in the deck, in order); null means all of them. */
export function showSlides(deck, settings = deck?.show, customId = null) {
  const slides = deck?.slides || [];
  const id = customId ?? settings?.custom;
  if (id) {
    const cs = (deck?.customShows || []).find((c) => c.id === id);
    const picked = cs ? cs.sids.map((sid) => slides.findIndex((s) => s.sid === sid)).filter((i) => i >= 0) : [];
    if (picked.length) return picked;
  }
  if (settings?.range && !customId) {
    const from = Math.max(1, settings.range.from);
    const to = Math.min(slides.length, settings.range.to);
    if (to >= from) return Array.from({ length: to - from + 1 }, (_, k) => from - 1 + k);
  }
  return null;
}

/** What the player is told (engine/motion.js createPlayer). */
export function playerOptions(settings) {
  const s = settings || {};
  return {
    loop: Boolean(s.loop), kiosk: Boolean(s.kiosk), kioskSeconds: s.kioskSeconds || 8, useTimings: s.useTimings !== false,
    static: Boolean(s.noAnimation), noAnimation: Boolean(s.noAnimation), narration: !s.noNarration, penColor: s.penColor || "#ff0000", captions: Boolean(s.captions), captionLang: s.captionLang || "ja-JP",
    // メディア コントロールの表示 (on unless turned off): a video's controls and a sound's bar during the show.
    mediaControls: !s.noControls,
  };
}

/** Ink written in the show, slide by slide, as ink objects (`objectFor(strokes)` makes one; see editor/ink.mjs). */
export function keptInk(ink, sidsAtStart, deck, objectFor) {
  const out = new Map();
  for (const { index, strokes } of Array.isArray(ink) ? ink : []) {
    const usable = (strokes || []).filter((st) => Array.isArray(st?.pts) && st.pts.length);
    if (!usable.length) continue;
    const sid = sidsAtStart?.[index];
    const at = sid ? deck.slides.findIndex((s) => s.sid === sid) : index;
    if (at < 0 || !deck.slides[at]) continue;
    out.set(at, [...(out.get(at) || []), objectFor(usable)]);
  }
  return out;
}

const newId = () => `cs${Math.random().toString(36).slice(2, 9)}`;

export function createShowTools(app) {
  const { h, E } = app;
  const slideLabel = (slide, i) => `${i + 1}. ${E.strip(slide.title || slide.message || E.TYPE_LABELS?.[slide.type] || "スライド").slice(0, 40)}`;
  const mainSlides = () => { const deck = app.deck(); const story = E.storyMap(deck.slides); return deck.slides.map((slide, i) => ({ slide, i })).filter(({ i }) => story.parent[i] == null); };

  function dialog(title, body, foot, cls = "") {
    const el = h("dialog", { class: ["show-dialog", cls], "aria-label": title },
      h("div", { class: "dialog-head" }, h("h3", {}, title), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => el.close() }, "✕")),
      h("div", { class: "dialog-body" }, body), h("div", { class: "dialog-foot" }, foot(() => el.close())));
    document.body.append(el);
    el.addEventListener("close", () => el.remove());
    el.showModal();
    return el;
  }
  const radio = (name, value, checked, label, extra = null) => h("label", { class: "sh-choice" }, h("input", { type: "radio", name, value, checked: checked || null }), h("span", {}, label), extra);
  const check = (name, checked, label) => h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name, checked: checked || null }), h("span", {}, label));

  /** スライド ショーの設定. */
  function openSettings() {
    const deck = app.deck();
    if (!deck) return;
    const cur = showOf(deck.show) || {};
    const shows = customShowsOf(deck.customShows) || [];
    const total = deck.slides.length;
    const num = (name, value, min, max) => h("input", { type: "number", name, value: String(value), min: String(min), max: String(max), class: "sh-num" });
    const from = num("from", cur.range?.from || 1, 1, total);
    const to = num("to", cur.range?.to || total, 1, total);
    const customSel = h("select", { name: "customShow", disabled: !shows.length || null }, shows.length ? shows.map((cs) => h("option", { value: cs.id, selected: cur.custom === cs.id || null }, cs.name)) : h("option", {}, "（ありません）"));
    const kioskSecs = num("kioskSeconds", cur.kioskSeconds || 8, 2, 600);
    const pen = h("select", { name: "penColor" }, PEN_COLORS.map(([c, label]) => h("option", { value: c, selected: (cur.penColor || "#ff0000") === c || null }, label)));
    const lang = h("select", { name: "captionLang" }, CAPTION_LANGS.map(([k, label]) => h("option", { value: k, selected: (cur.captionLang || "ja-JP") === k || null }, label)));
    const form = h("form", { class: "sh-form", onsubmit: (e) => e.preventDefault() },
      h("fieldset", {}, h("legend", {}, "種類"),
        radio("kind", "speaker", !cur.kiosk, "発表者として使用する（フルスクリーン表示）"),
        radio("kind", "kiosk", cur.kiosk, "自動プレゼンテーション（フルスクリーン表示）", h("small", { class: "hint" }, "クリックやキーでは進まず、Escで終わるまで繰り返します")),
        h("label", { class: "sh-inline" }, "タイミングのないスライドを表示する秒数 ", kioskSecs, " 秒")),
      h("fieldset", {}, h("legend", {}, "オプション"),
        check("loop", cur.loop || cur.kiosk, "Esc キーが押されるまで繰り返す"),
        check("noNarration", cur.noNarration, "ナレーションを付けない"),
        check("noAnimation", cur.noAnimation, "アニメーションを表示しない"),
        h("label", { class: "sh-inline" }, "ペンの色 ", pen),
        check("captions", cur.captions, "常に字幕を使用する"),
        h("label", { class: "sh-inline" }, "字幕の言語 ", lang),
        h("label", { class: "sh-choice" }, h("input", { type: "checkbox", name: "mediaControls", checked: !cur.noControls || null }), h("span", {}, "メディア コントロールを表示する"))),
      h("fieldset", {}, h("legend", {}, "スライドの表示"),
        radio("slides", "all", !cur.range && !cur.custom, "すべて"),
        radio("slides", "range", Boolean(cur.range) && !cur.custom, "", h("span", { class: "sh-inline" }, "開始 ", from, " 終了 ", to)),
        radio("slides", "custom", Boolean(cur.custom), "目的別スライド ショー ", customSel)),
      h("fieldset", {}, h("legend", {}, "スライドの切り替え"),
        radio("advance", "manual", cur.useTimings === false, "クリック時"),
        radio("advance", "timings", cur.useTimings !== false, "保存済みのタイミング（リハーサル・自動で切り替え）")));
    if (!shows.length) form.querySelector('input[value="custom"]').disabled = true;
    dialog("スライド ショーの設定", form, (close) => [
      h("button", { type: "button", class: "btn btn-ghost", onclick: close }, "キャンセル"),
      h("button", { type: "button", class: "btn btn-primary sh-ok", onclick: () => {
        const f = new FormData(form);
        const which = f.get("slides");
        const next = showOf({
          kiosk: f.get("kind") === "kiosk", loop: f.has("loop"), noNarration: f.has("noNarration"), noAnimation: f.has("noAnimation"), captions: f.has("captions"), noControls: !f.has("mediaControls"),
          useTimings: f.get("advance") !== "manual", penColor: f.get("penColor"), captionLang: f.get("captionLang"), kioskSeconds: Number(f.get("kioskSeconds")),
          range: which === "range" ? { from: Number(f.get("from")), to: Number(f.get("to")) } : null,
          custom: which === "custom" ? f.get("customShow") : null,
        });
        if (which === "range" && !next?.range) { app.toast("開始と終了のスライド番号を確かめてください"); return; }
        app.setDeckFields({ show: next });
        close();
        app.toast("スライド ショーの設定を保存しました（F5で発表すると、この設定で始まります）");
      } }, "OK"),
    ], "sh-settings");
  }

  /** 目的別スライド ショー: make, rename, reorder, copy, delete and start named lists of slides. */
  function openCustomShows(focusId = null) {
    const deck = app.deck();
    if (!deck) return;
    let shows = (customShowsOf(deck.customShows) || []).map((cs) => ({ ...cs, sids: [...cs.sids] }));
    let current = shows.find((cs) => cs.id === focusId)?.id || shows[0]?.id || null;
    const listBox = h("div", { class: "sh-shows", role: "listbox", "aria-label": "目的別スライド ショー" });
    const editor = h("div", { class: "sh-edit" });
    const sidOf = (i) => app.ensureSid(i);
    function render() {
      listBox.replaceChildren(...(shows.length ? shows.map((cs) => h("button", { type: "button", role: "option", class: ["sh-show", cs.id === current ? "on" : ""], "aria-selected": String(cs.id === current), onclick: () => { current = cs.id; render(); } }, cs.name, h("small", {}, `${cs.sids.length}枚`)))
        : [h("p", { class: "hint" }, "「新規作成」で、発表する相手や時間に合わせたスライドの組み合わせを作れます。")]));
      const cs = shows.find((c) => c.id === current);
      if (!cs) { editor.replaceChildren(); return; }
      const name = h("input", { type: "text", class: "sh-name", value: cs.name, maxlength: "60", "aria-label": "スライド ショーの名前", oninput: (e) => { cs.name = e.target.value; listBox.querySelector(".sh-show.on")?.firstChild?.replaceWith(e.target.value || "（名前なし）"); } });
      const all = mainSlides();
      const pick = h("div", { class: "sh-col" }, h("b", {}, "資料のスライド"), h("div", { class: "sh-list" }, all.map(({ slide, i }) => {
        const sid = slide.sid;
        const inShow = sid && cs.sids.includes(sid);
        return h("label", { class: ["sh-row", slide.hidden ? "hidden-slide" : ""] }, h("input", { type: "checkbox", checked: inShow || null, "data-index": String(i), onchange: (e) => {
          const s = sidOf(i);
          if (e.target.checked) { if (!cs.sids.includes(s)) cs.sids.push(s); } else cs.sids = cs.sids.filter((x) => x !== s);
          render();
        } }), h("span", {}, slideLabel(slide, i)), slide.hidden ? h("small", {}, "非表示") : null);
      })));
      const order = h("div", { class: "sh-col" }, h("b", {}, "このショーで発表する順番"), h("ol", { class: "sh-list sh-order" }, cs.sids.map((sid, k) => {
        const i = deck.slides.findIndex((s) => s.sid === sid);
        if (i < 0) return null;
        const move = (d) => { const j = k + d; if (j < 0 || j >= cs.sids.length) return; [cs.sids[k], cs.sids[j]] = [cs.sids[j], cs.sids[k]]; render(); };
        return h("li", { class: "sh-row" }, h("span", {}, slideLabel(deck.slides[i], i)),
          h("button", { type: "button", class: "btn btn-xs", title: "上へ", "aria-label": "上へ", disabled: k === 0 || null, onclick: () => move(-1) }, "↑"),
          h("button", { type: "button", class: "btn btn-xs", title: "下へ", "aria-label": "下へ", disabled: k === cs.sids.length - 1 || null, onclick: () => move(1) }, "↓"),
          h("button", { type: "button", class: "btn btn-xs", title: "このショーから外す", "aria-label": "外す", onclick: () => { cs.sids.splice(k, 1); render(); } }, "✕"));
      })));
      editor.replaceChildren(h("label", { class: "sh-inline" }, "名前 ", name), h("div", { class: "sh-cols" }, pick, order));
    }
    const actions = h("div", { class: "sh-actions" },
      h("button", { type: "button", class: "btn btn-sm sh-new", onclick: () => { const cs = { id: newId(), name: `目的別スライド ショー ${shows.length + 1}`, sids: [] }; shows.push(cs); current = cs.id; render(); editor.querySelector(".sh-name")?.select(); } }, "新規作成"),
      h("button", { type: "button", class: "btn btn-sm", onclick: () => { const src = shows.find((c) => c.id === current); if (!src) return; const cs = { id: newId(), name: `${src.name} のコピー`.slice(0, 60), sids: [...src.sids] }; shows.push(cs); current = cs.id; render(); } }, "コピー"),
      h("button", { type: "button", class: "btn btn-sm", onclick: () => { shows = shows.filter((c) => c.id !== current); current = shows[0]?.id || null; render(); } }, "削除"));
    render();
    const save = () => {
      const next = customShowsOf(shows.map((cs) => ({ ...cs, name: cs.name.trim() || "目的別スライド ショー" })));
      const settings = showOf(deck.show);
      // A show the settings pointed at is gone: back to all slides.
      const keep = settings?.custom && !next?.some((cs) => cs.id === settings.custom) ? showOf({ ...settings, custom: null }) : settings;
      app.setDeckFields({ customShows: next, show: keep });
      return next;
    };
    dialog("目的別スライド ショー", h("div", { class: "sh-custom" }, h("div", { class: "sh-side" }, listBox, actions), editor), (close) => [
      h("button", { type: "button", class: "btn btn-ghost", onclick: close }, "キャンセル"),
      h("button", { type: "button", class: "btn sh-start", onclick: () => { const id = current; const next = save(); close(); if (id && next?.some((cs) => cs.id === id)) app.presentCustom(id); else app.toast("スライドを1枚以上選んでください"); } }, "開始"),
      h("button", { type: "button", class: "btn btn-primary sh-ok", onclick: () => { const empty = shows.filter((cs) => !cs.sids.length).length; save(); close(); app.toast(empty ? "スライドのないショーは保存しませんでした" : "目的別スライド ショーを保存しました"); } }, "OK"),
    ], "sh-customs");
  }

  return { openSettings, openCustomShows };
}
