// Review comments on slides, as PowerPoint's コメント pane (校閲 → 新しいコメント): threads with the writer's name and
// time, replies, @mentions, 編集, 解決 / もう一度開く, 削除, and 前へ / 次へ to the slides that have some. A comment made
// while an object is selected is pinned to it (as PowerPoint's modern comments); the others to the slide. Pins are
// drawn on the stage (校閲 →「コメントの表示」) and open their thread.
// Comments are slide.comments ({ id, text, at, by, uid, anchor, done, doneBy, edited, replies: [{ id, text, at, by,
// uid, edited }] }): people's own, kept through AI changes, never in the slide show or the exported HTML.

import * as ops from "./ops.mjs";
import { userName, userId, setUserName, avatar, mentionNodes, mentionsIn, colorOf, initials } from "./people.mjs";

const SHOW_KEY = "hpptx-show-comments";
const FILTERS = { all: "すべて", open: "未解決", mine: "自分宛て（@メンション）" };
const newId = (prefix = "c") => `${prefix}${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;

export function createComments(app) {
  const { h, E } = app;
  const pane = document.getElementById("commentPane");
  let draft = "";
  let all = false;
  let filter = "all";
  let focus = null; // the thread to bring into view (from a pin)
  let replying = null; // thread id with an open reply box
  let replyDraft = "";
  let editing = null; // { id, rid?, text }
  let showPins = (() => { try { return localStorage.getItem(SHOW_KEY) !== "0"; } catch { return true; } })();
  const slides = () => app.deck()?.slides || [];
  const when = (at) => {
    const d = new Date(at);
    if (Number.isNaN(d.getTime())) return "";
    const pad = (n) => String(n).padStart(2, "0");
    const today = new Date();
    const same = d.toDateString() === today.toDateString();
    return same ? `今日 ${pad(d.getHours())}:${pad(d.getMinutes())}` : `${d.getMonth() + 1}/${d.getDate()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  };
  const mine = (c) => (c.uid ? c.uid === userId() : !c.by || c.by === userName());
  /** Everyone named in the deck's comments, and the people editing with us (for @ suggestions). */
  function knownNames() {
    const names = new Set(app.people?.() || []);
    for (const s of slides()) for (const c of s.comments || []) {
      if (c.by) names.add(c.by);
      for (const r of c.replies || []) if (r.by) names.add(r.by);
    }
    if (userName()) names.add(userName());
    return [...names];
  }
  const mentionsMe = (c) => {
    const me = userName();
    if (!me) return false;
    return [c, ...(c.replies || [])].some((x) => mentionsIn(x.text, knownNames()).includes(me));
  };
  function objectName(slide, id) {
    const o = (slide?.elements || []).find((x) => x.id === id);
    if (!o) return null;
    const text = String(o.text || o.html || "").replace(/<[^>]+>/g, "").trim();
    const kinds = { shape: "図形", text: "テキスト", image: "画像", line: "線", icon: "アイコン", video: "動画", audio: "オーディオ", table: "表", chart: "グラフ", smartart: "SmartArt" };
    return `${kinds[o.kind] || "部品"}${text ? `「${text.slice(0, 14)}${text.length > 14 ? "…" : ""}」` : ""}`;
  }

  /** The name to write under: asked once (and kept in this browser). */
  async function ensureName() {
    if (userName()) return userName();
    const typed = await app.ask("あなたの名前", "コメントと共同編集で表示する名前（このブラウザに保存します）", "");
    const name = setUserName(typed || "");
    if (!name) { app.toast("名前を入れると投稿できます"); return ""; }
    app.nameChanged?.(name);
    return name;
  }
  async function changeName() {
    const before = userName();
    const typed = await app.ask("ユーザー名", "コメントと共同編集で表示する名前（このブラウザに保存します）", before);
    if (typed == null) return;
    const name = setUserName(typed);
    if (!name) { app.toast("名前が空です"); return; }
    if (name === before) return;
    // This browser's comments carry the new name, as PowerPoint shows the account's current name.
    const uid = userId();
    let first = true;
    slides().forEach((slide, i) => {
      if (!(slide.comments || []).some((c) => c.uid === uid || (c.replies || []).some((r) => r.uid === uid))) return;
      const next = slide.comments.map((c) => ({
        ...c,
        ...(c.uid === uid ? { by: name } : {}),
        ...(c.doneBy && c.doneBy === before ? { doneBy: name } : {}),
        ...(c.replies ? { replies: c.replies.map((r) => (r.uid === uid ? { ...r, by: name } : r)) } : {}),
      }));
      app.setCommentsOf(i, next, { undo: first });
      first = false;
    });
    app.nameChanged?.(name);
    app.toast(`名前を「${name}」にしました`);
    render();
  }

  // ---------------------------------------------------------------- the pane

  function render() {
    if (!pane || pane.hidden) return;
    const index = app.index();
    const me = userName();
    const head = h("div", { class: "fp-head an-pane-head" }, h("b", {}, "コメント"),
      h("span", { class: "an-pane-btns" },
        h("button", { type: "button", class: "btn btn-sm", title: "前のコメントのあるスライドへ", onclick: () => go(-1) }, "‹ 前へ"),
        h("button", { type: "button", class: "btn btn-sm", title: "次のコメントのあるスライドへ", onclick: () => go(1) }, "次へ ›")));
    const body = h("div", { class: "fp-body" });
    body.append(h("div", { class: "cm-who" },
      me ? avatar(h, me, { key: me }) : h("span", { class: "pp-avatar pp-none" }, "?"),
      h("span", { class: "cm-who-name" }, me || "名前が未設定です"),
      h("button", { type: "button", class: "btn btn-sm btn-ghost cm-rename", title: "コメントと共同編集で表示する名前", onclick: () => changeName() }, me ? "名前を変更" : "名前を設定")));
    body.append(h("div", { class: "cm-tools" },
      h("label", { class: "fp-check cm-all" }, h("input", { type: "checkbox", checked: all || null, onchange: (event) => { all = event.target.checked; render(); } }), "すべてのスライド"),
      h("select", { class: "cm-filter", "aria-label": "表示するコメント", onchange: (event) => { filter = event.target.value; render(); } },
        Object.entries(FILTERS).map(([value, label]) => h("option", { value, selected: value === filter || null }, label)))));
    const list = h("ul", { class: "cm-list" });
    const targets = all ? slides().map((slide, i) => [slide, i]) : [[slides()[index], index]];
    let count = 0;
    let total = 0;
    for (const [slide, i] of targets) {
      for (const c of slide?.comments || []) {
        total += 1;
        if (filter === "open" && c.done) continue;
        if (filter === "mine" && !mentionsMe(c)) continue;
        count += 1;
        list.append(thread(slide, i, c));
      }
    }
    if (!count) {
      body.append(h("p", { class: "hint" }, total ? "条件に合うコメントはありません。"
        : all ? "この資料にはまだコメントがありません。" : "このスライドにはまだコメントがありません。下に書いて追加できます。"));
    }
    body.append(list);
    const sel = app.editor?.selection || [];
    const anchor = sel.length === 1 ? objectName(slides()[index], sel[0]) : null;
    const input = mentionBox(h("textarea", { class: "cm-input", rows: 3, "aria-label": "新しいコメント", placeholder: `${anchor ? `${anchor}に` : `スライド ${index + 1} に`}コメント（@で相手を指定できます）`, oninput: (event) => { draft = event.target.value; } }, draft), post);
    body.append(h("div", { class: "cm-new" },
      anchor ? h("div", { class: "cm-anchor-note" }, `📌 選んでいる${anchor}に付けます`) : null,
      input,
      h("div", { class: "cm-new-foot" }, h("span", { class: "hint" }, "⌘Enterで投稿"), h("button", { type: "button", class: "btn btn-primary btn-sm cm-post", onclick: () => post() }, "投稿"))));
    pane.replaceChildren(head, body);
    if (focus) {
      const el = pane.querySelector(`.cm-item[data-id="${CSS.escape(focus)}"]`);
      focus = null;
      if (el) { el.scrollIntoView({ block: "nearest" }); el.classList.add("flash"); setTimeout(() => el.classList.remove("flash"), 1200); }
    }
  }

  function thread(slide, i, c) {
    const known = knownNames();
    const me = userName();
    const anchorName = c.anchor ? objectName(slide, c.anchor) : null;
    const isEditing = (rid) => editing && editing.id === c.id && (editing.rid || null) === (rid || null);
    const message = (m, rid) => h("div", { class: "cm-msg" },
      h("div", { class: "cm-meta" },
        avatar(h, m.by || "名前なし", { key: m.by || m.uid || "", size: 22 }),
        h("b", { class: "cm-by" }, m.by || "名前なし"),
        h("span", { class: "cm-at", title: m.at || "" }, when(m.at)),
        m.edited ? h("span", { class: "cm-edited", title: `編集 ${when(m.edited)}` }, "（編集済み）") : null),
      isEditing(rid)
        ? editBox(i, c.id, rid)
        : h("div", { class: "cm-text" }, ...mentionNodes(h, m.text, known, me)));
    const ownThread = mine(c);
    const replies = (c.replies || []).map((r) => h("li", { class: "cm-reply", "data-rid": r.id }, message(r, r.id),
      mine(r) && !isEditing(r.id) ? h("div", { class: "cm-actions" },
        h("button", { type: "button", class: "btn btn-xs btn-ghost", onclick: () => { editing = { id: c.id, rid: r.id, text: r.text }; render(); } }, "編集"),
        h("button", { type: "button", class: "btn btn-xs btn-ghost btn-danger", onclick: () => removeReply(i, c.id, r.id) }, "削除")) : null));
    const replyBox = replying === c.id
      ? h("div", { class: "cm-reply-new" },
        mentionBox(h("textarea", { class: "cm-reply-input", rows: 2, "aria-label": "返信", placeholder: "返信（@で相手を指定）", oninput: (event) => { replyDraft = event.target.value; } }, replyDraft), () => reply(i, c.id)),
        h("div", { class: "cm-new-foot" },
          h("button", { type: "button", class: "btn btn-sm btn-ghost", onclick: () => { replying = null; replyDraft = ""; render(); } }, "やめる"),
          h("button", { type: "button", class: "btn btn-primary btn-sm cm-reply-post", onclick: () => reply(i, c.id) }, "返信")))
      : null;
    return h("li", { class: ["cm-item", c.done ? "done" : "", mentionsMe(c) ? "to-me" : ""], "data-id": c.id },
      h("div", { class: "cm-where" },
        all ? h("button", { type: "button", class: "cm-slide", onclick: () => app.select(i) }, `スライド ${i + 1}`) : null,
        anchorName ? h("button", { type: "button", class: "cm-anchor", title: "この部品を選ぶ", onclick: () => pick(i, c.anchor) }, `📌 ${anchorName}`) : null,
        c.done ? h("span", { class: "cm-done" }, `解決済み${c.doneBy ? `（${c.doneBy}）` : ""}`) : null),
      message(c, null),
      replies.length ? h("ul", { class: "cm-replies" }, replies) : null,
      replyBox,
      !isEditing(null) ? h("div", { class: "cm-actions" },
        !c.done ? h("button", { type: "button", class: "btn btn-sm", onclick: () => { replying = c.id; replyDraft = ""; render(); requestAnimationFrame(() => pane.querySelector(".cm-reply-input")?.focus()); } }, "返信") : null,
        h("button", { type: "button", class: "btn btn-sm", onclick: () => resolve(i, c) }, c.done ? "もう一度開く" : "✓ 解決"),
        ownThread ? h("button", { type: "button", class: "btn btn-sm btn-ghost", onclick: () => { editing = { id: c.id, rid: null, text: c.text }; render(); } }, "編集") : null,
        h("button", { type: "button", class: "btn btn-sm btn-ghost btn-danger", title: "このスレッドを削除（返信も消えます）", onclick: () => remove(i, c.id) }, "スレッドを削除")) : null);
  }

  function editBox(i, id, rid) {
    const box = mentionBox(h("textarea", { class: "cm-edit-input", rows: 3, "aria-label": "コメントを編集", oninput: (event) => { editing.text = event.target.value; } }, editing.text), () => saveEdit(i));
    requestAnimationFrame(() => { box.focus(); box.setSelectionRange(box.value.length, box.value.length); });
    return h("div", { class: "cm-edit" }, box, h("div", { class: "cm-new-foot" },
      h("button", { type: "button", class: "btn btn-sm btn-ghost", onclick: () => { editing = null; render(); } }, "やめる"),
      h("button", { type: "button", class: "btn btn-primary btn-sm cm-save", onclick: () => saveEdit(i) }, "保存")));
  }

  /** A text box that offers the known names after "@", as PowerPoint's people picker. */
  function mentionBox(input, submit) {
    let menu = null;
    let items = [];
    let active = 0;
    const close = () => { menu?.remove(); menu = null; items = []; };
    const word = () => input.value.slice(0, input.selectionStart).match(/(^|[\s(（「])@([^\s@]*)$/);
    const choose = (name) => {
      const m = word();
      if (!m) return close();
      const start = input.selectionStart - m[2].length - 1;
      input.value = `${input.value.slice(0, start)}@${name} ${input.value.slice(input.selectionStart)}`;
      const at = start + name.length + 2;
      input.setSelectionRange(at, at);
      input.dispatchEvent(new Event("input"));
      close();
      input.focus();
    };
    const show = () => {
      const m = word();
      const me = userName();
      items = m ? knownNames().filter((n) => n !== me && n.toLowerCase().includes(m[2].toLowerCase())).slice(0, 6) : [];
      if (!items.length) return close();
      active = Math.min(active, items.length - 1);
      if (!menu) { menu = h("ul", { class: "cm-mentions", role: "listbox" }); input.after(menu); }
      menu.replaceChildren(...items.map((n, k) => h("li", { role: "option", class: k === active ? "on" : "", "aria-selected": String(k === active), onmousedown: (event) => { event.preventDefault(); choose(n); } }, avatar(h, n, { size: 18 }), h("span", {}, n))));
    };
    input.addEventListener("input", show);
    input.addEventListener("blur", () => setTimeout(close, 100));
    input.addEventListener("keydown", (event) => {
      if (event.isComposing) return;
      if (menu && items.length) {
        if (event.key === "ArrowDown" || event.key === "ArrowUp") { event.preventDefault(); active = (active + (event.key === "ArrowDown" ? 1 : items.length - 1)) % items.length; show(); return; }
        if (event.key === "Enter" || event.key === "Tab") { event.preventDefault(); choose(items[active]); return; }
        if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); close(); return; }
      }
      if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) { event.preventDefault(); submit(); }
    });
    return input;
  }

  // ---------------------------------------------------------------- changes (each one undo step)

  const listOf = (i) => slides()[i]?.comments || [];
  const stamp = () => ({ by: userName(), uid: userId(), at: new Date().toISOString() });
  async function post() {
    const text = draft.trim();
    if (!text) return;
    const i = app.index();
    const slide = slides()[i];
    if (!slide || !(await ensureName())) return;
    const sel = app.editor?.selection || [];
    const anchor = sel.length === 1 && (slide.elements || []).some((o) => o.id === sel[0]) ? sel[0] : null;
    draft = "";
    app.setCommentsOf(i, [...listOf(i), { id: newId(), text: text.slice(0, 2000), ...stamp(), ...(anchor ? { anchor } : {}) }]);
    const named = mentionsIn(text, knownNames()).filter((n) => n !== userName());
    app.toast(`スライド ${i + 1} にコメントを付けました${named.length ? `（${named.map((n) => `@${n}`).join("・")}）` : ""}`);
  }
  async function reply(i, id) {
    const text = replyDraft.trim();
    if (!text || !(await ensureName())) return;
    replying = null;
    replyDraft = "";
    app.setCommentsOf(i, listOf(i).map((c) => (c.id === id ? { ...c, replies: [...(c.replies || []), { id: newId("r"), text: text.slice(0, 2000), ...stamp() }].slice(-100) } : c)));
  }
  function saveEdit(i) {
    if (!editing) return;
    const { id, rid, text } = editing;
    const clean = String(text || "").trim().slice(0, 2000);
    editing = null;
    if (!clean) { render(); return; }
    const at = new Date().toISOString();
    app.setCommentsOf(i, listOf(i).map((c) => {
      if (c.id !== id) return c;
      if (!rid) return c.text === clean ? c : { ...c, text: clean, edited: at };
      return { ...c, replies: (c.replies || []).map((r) => (r.id === rid && r.text !== clean ? { ...r, text: clean, edited: at } : r)) };
    }));
  }
  function resolve(i, c) {
    app.setCommentsOf(i, listOf(i).map((x) => {
      if (x.id !== c.id) return x;
      const { done, doneBy, ...rest } = x;
      return c.done ? rest : { ...rest, done: true, ...(userName() ? { doneBy: userName() } : {}) };
    }));
  }
  function remove(i, id) {
    app.setCommentsOf(i, listOf(i).filter((c) => c.id !== id));
  }
  function removeReply(i, id, rid) {
    app.setCommentsOf(i, listOf(i).map((c) => (c.id === id ? { ...c, replies: (c.replies || []).filter((r) => r.id !== rid) } : c)));
  }
  /** 校閲 → 削除: every comment of this slide, or of the deck (one undo step). */
  function clear(scope) {
    const which = scope === "deck" ? slides().map((s, i) => i).filter((i) => listOf(i).length) : listOf(app.index()).length ? [app.index()] : [];
    if (!which.length) { app.toast(scope === "deck" ? "この資料にコメントはありません" : "このスライドにコメントはありません"); return; }
    const n = which.reduce((sum, i) => sum + listOf(i).length, 0);
    which.forEach((i, k) => app.setCommentsOf(i, [], { undo: k === 0 }));
    app.toast(`コメント${n}件を削除しました（⌘Zで戻せます）`);
  }
  function pick(i, anchor) {
    if (app.index() !== i) app.select(i);
    requestAnimationFrame(() => app.editor?.select?.([anchor]));
  }
  /** The next (or previous) slide that has comments, wrapping around. */
  function go(dir) {
    const n = slides().length;
    for (let k = 1; k <= n; k += 1) {
      const i = (((app.index() + dir * k) % n) + n) % n;
      if (slides()[i]?.comments?.length) { app.select(i); return; }
    }
    app.toast("コメントのあるスライドはほかにありません");
  }
  const focusNew = () => requestAnimationFrame(() => pane?.querySelector(".cm-input")?.focus());

  // ---------------------------------------------------------------- pins on the stage

  /** One pin per open thread: at the top right of the object it is pinned to, or along the slide's top left. */
  function pins(list, k) {
    if (!showPins || !app.onSlide?.()) return [];
    const slide = slides()[app.index()];
    const out = [];
    let free = 0;
    for (const c of slide?.comments || []) {
      if (c.done) continue;
      const o = c.anchor ? list.find((x) => x.id === c.anchor && !x.hidden) : null;
      const b = o ? ops.bounds([o], (x) => E.lineEnds(x, list)) : null;
      // Pinned: at the object's top right; the slide's own: in a row from its top left (screen px, so they never overlap).
      const left = b ? (b.x + b.w) * k - 4 : 6 + free * 34;
      const top = b ? b.y * k - 24 : 6;
      if (!b) free += 1;
      const n = 1 + (c.replies?.length || 0);
      out.push(h("button", {
        type: "button", class: ["ed-comment-pin", mentionsMe(c) ? "to-me" : ""], "data-id": c.id,
        title: `${c.by || "名前なし"}：${c.text.slice(0, 60)}${n > 1 ? `（返信${n - 1}件）` : ""}`,
        style: { left: `${left}px`, top: `${top}px`, "--pin": colorOf(c.by || c.uid || "") },
        onpointerdown: (event) => event.stopPropagation(),
        onclick: (event) => { event.stopPropagation(); open(c.id); },
      }, initials(c.by || "?"), n > 1 ? h("small", {}, String(n)) : null));
    }
    return out;
  }
  function open(id) {
    focus = id;
    app.showPanel("comment");
    if (focus) render(); // the pane was open already (opening it renders it)
  }
  function setShowPins(on) {
    showPins = Boolean(on);
    try { localStorage.setItem(SHOW_KEY, showPins ? "1" : "0"); } catch { /* private window */ }
    app.editor?.draw();
  }
  app.editor?.overlay(pins);
  // The new-comment box says which object it will be pinned to.
  let anchorKey = "";
  app.editor?.subscribe(() => {
    const sel = app.editor.selection;
    const key = sel.length === 1 ? sel[0] : "";
    if (key !== anchorKey) { anchorKey = key; render(); }
  });

  return { render, go, focusNew, open, clear, changeName, ensureName, knownNames, showPins: () => showPins, setShowPins };
}
