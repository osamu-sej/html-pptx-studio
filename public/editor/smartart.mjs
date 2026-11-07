// SmartArt, PowerPoint style: 挿入 → SmartArt (the gallery by kind, each with a preview), the テキスト ウィンドウ
// (one line per item; Tab / Shift+Tab change the level, Enter adds, Backspace on an empty line removes), the
// contextual SmartArt のデザイン tab (add shapes before/after/above/below, bullets, levels, order, right to left,
// layouts, colours, styles, text size, reset, 図形に変換, テキストに変換), and ホーム →「SmartArt に変換」 for a text box.
// The object is { kind: "smartart", smartart: { layout, items: [{ text, level }], color, style, fsScale, rtl } };
// objects.js draws it.

const NEW_BOX = { w: 1400, h: 640 };

export function createSmartArt(editor, app, kit) {
  const { E, h } = app;
  const { btn, drop, group, col, menu, openPop, updater, showTab } = kit;
  const selected = () => editor.selectedObjects();
  const current = () => { const list = selected(); return list.length === 1 && list[0].kind === "smartart" ? list[0] : null; };
  const isSmartArt = () => Boolean(current());
  let pane = null; // { id, el, focus }
  let picked = null; // the item the commands act on (the line with the caret in the text pane)
  let typing = false; // a run of typing in one line is one undo step

  // ---------------------------------------------------------------- changing the items (one undo step per command)

  function setSmartart(id, fn, { undo = true, rerender = true } = {}) {
    const list = editor.objects();
    const o = list.find((x) => x.id === id);
    if (!o) return;
    const next = E.normalizeSmartart(fn(structuredClone(o.smartart)));
    editor.commit(list.map((x) => (x.id === id ? { ...x, smartart: next } : x)), { undo, select: [id] });
    if (rerender) { typing = false; if (pane?.id === id) renderPane(); }
  }
  const itemsOf = () => current()?.smartart?.items || [];
  const at = () => { const n = itemsOf().length; return picked != null && picked < n ? picked : n - 1; };
  /** The item's whole branch: it and the deeper items after it. */
  function branch(items, i) {
    let end = i + 1;
    while (end < items.length && items[end].level > items[i].level) end += 1;
    return [i, end];
  }
  function addShape(where) {
    const o = current();
    if (!o) return;
    const i = at();
    setSmartart(o.id, (sa) => {
      const items = sa.items;
      const base = items[i] || { level: 0 };
      const [, end] = branch(items, i);
      if (where === "after") { items.splice(end, 0, { text: "", level: base.level }); picked = end; }
      else if (where === "before") { items.splice(i, 0, { text: "", level: base.level }); picked = i; }
      else if (where === "below") { items.splice(i + 1, 0, { text: "", level: Math.min(4, base.level + 1) }); picked = i + 1; }
      else if (where === "above") {
        // A new parent: the item and its branch move one level down under it.
        const [from, to] = branch(items, i);
        for (let k = from; k < to; k += 1) items[k].level = Math.min(4, items[k].level + 1);
        items.splice(i, 0, { text: "", level: base.level });
        picked = i;
      }
      return sa;
    });
    openPane(o.id, picked);
  }
  function addBullet() {
    const o = current();
    if (!o) return;
    const i = at();
    setSmartart(o.id, (sa) => { const [, end] = branch(sa.items, i); sa.items.splice(end, 0, { text: "", level: Math.min(4, (sa.items[i]?.level ?? 0) + 1) }); picked = end; return sa; });
    openPane(o.id, picked);
  }
  function shiftLevel(delta) {
    const o = current();
    if (!o) return;
    const i = at();
    if (i <= 0 && delta > 0) { app.toast("最初の項目はレベルを下げられません"); return; }
    setSmartart(o.id, (sa) => {
      const [from, to] = branch(sa.items, i);
      for (let k = from; k < to; k += 1) sa.items[k].level = Math.max(0, Math.min(4, sa.items[k].level + delta));
      return sa;
    });
  }
  function moveItem(dir) {
    const o = current();
    if (!o) return;
    const items = itemsOf();
    const i = at();
    const [from, to] = branch(items, i);
    // Siblings swap with their branches.
    let target = null;
    if (dir < 0) { for (let k = from - 1; k >= 0; k -= 1) { if (items[k].level < items[i].level) break; if (items[k].level === items[i].level) { target = k; break; } } }
    else if (to < items.length && items[to].level === items[i].level) target = to;
    if (target == null) { app.toast(dir < 0 ? "これより上には移動できません" : "これより下には移動できません"); return; }
    setSmartart(o.id, (sa) => {
      const list = sa.items;
      if (dir < 0) { const [tf] = branch(list, target); const moved = list.splice(from, to - from); list.splice(tf, 0, ...moved); picked = tf; }
      else { const [, te] = branch(list, target); const moved = list.splice(from, to - from); list.splice(te - (to - from), 0, ...moved); picked = te - (to - from); }
      return sa;
    });
  }
  const setField = (key, value) => { const o = current(); if (o) setSmartart(o.id, (sa) => ({ ...sa, [key]: value })); };

  // ---------------------------------------------------------------- inserting and converting

  function insert(layout) {
    const box = { x: (E.W - NEW_BOX.w) / 2, y: 250, w: NEW_BOX.w, h: NEW_BOX.h };
    const o = { id: E.newObjectId(), kind: "smartart", ...box, smartart: { layout, items: E.smartartSample(layout) } };
    editor.commit([...editor.objects(), E.normalizeObject(o)], { select: [o.id] });
    showTab("smartartDesign");
    openPane(o.id, 0);
    app.toast(`SmartArt「${E.SMARTART_LAYOUTS[layout].label}」を入れました。テキスト ウィンドウで文字を直せます`);
  }
  /** The paragraphs of a text box as items (list nesting gives the level). */
  function itemsFromRich(html) {
    const box = document.createElement("div");
    box.append(E.richNodes(html || ""));
    const items = [];
    const walk = (node, depth) => {
      for (const child of node.children) {
        if (child.tagName === "UL" || child.tagName === "OL") walk(child, depth + 1);
        else if (child.tagName === "LI") {
          const own = [...child.childNodes].filter((n) => !(n.nodeType === 1 && /^(UL|OL)$/.test(n.tagName))).map((n) => n.textContent).join("").trim();
          if (own) items.push({ text: own, level: Math.max(0, depth - 1) });
          for (const sub of child.children) if (/^(UL|OL)$/.test(sub.tagName)) walk(sub, depth + 1);
        } else if (child.textContent.trim()) {
          // A paragraph indented with spaces or a tab counts its indent as levels.
          const text = child.textContent.replace(/ /g, " ");
          const indent = (text.match(/^[\t ]*/)[0].replace(/\t/g, "  ").length / 2) | 0;
          items.push({ text: text.trim().replace(/^[・•●○■◆\-–]\s*/, ""), level: Math.min(4, indent) });
        }
      }
    };
    walk(box, 0);
    return items;
  }
  function convertText(layout) {
    const o = selected().find((x) => ["shape", "text"].includes(x.kind) && E.richToText(x.text || "").trim());
    if (!o) { app.toast("文字の入ったテキスト ボックスか図形を選んでください"); return; }
    const items = itemsFromRich(o.text);
    if (!items.length) { app.toast("SmartArtにする文字がありません"); return; }
    const w = Math.max(o.w, 900);
    const hh = Math.max(o.h, 480);
    const next = E.normalizeObject({ id: o.id, kind: "smartart", x: Math.max(0, Math.min(E.W - w, o.x)), y: Math.max(0, Math.min(E.H - hh, o.y)), w, h: hh, smartart: { layout, items } });
    editor.commit(editor.objects().map((x) => (x.id === o.id ? next : x)), { select: [o.id] });
    showTab("smartartDesign");
    app.toast(`${items.length}項目のSmartArtにしました（⌘Zで戻せます）`);
  }
  /** 図形に変換: the parts become ordinary shapes and lines (grouped), keeping the place in the stacking order. */
  function toShapes(id = current()?.id) {
    const list = editor.objects();
    const o = list.find((x) => x.id === id);
    if (!o) return;
    const parts = E.smartartObjects(E.withDefaults(o));
    const k = list.indexOf(o);
    closePane();
    editor.commit([...list.slice(0, k), ...parts, ...list.slice(k + 1)], { select: parts.map((x) => x.id) });
    app.toast(`SmartArtを${parts.length}個の図形にしました（グループ化しています）`);
  }
  /** テキストに変換: a text box with the items as a bulleted list. */
  function toText() {
    const o = current();
    if (!o) return;
    const items = o.smartart.items;
    let html = "";
    let depth = -1;
    for (const it of items) {
      while (depth < it.level) { html += "<ul>"; depth += 1; }
      while (depth > it.level) { html += "</ul>"; depth -= 1; }
      html += `<li>${it.text.replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[c]) || "<br>"}</li>`;
    }
    while (depth >= 0) { html += "</ul>"; depth -= 1; }
    closePane();
    const next = E.normalizeObject({ id: o.id, kind: "text", x: o.x, y: o.y, w: o.w, h: o.h, text: E.sanitizeRich(html), fs: 32, autofit: "grow" });
    editor.commit(editor.objects().map((x) => (x.id === o.id ? next : x)), { select: [o.id] });
  }

  // ---------------------------------------------------------------- the テキスト ウィンドウ

  function openPane(id, item = null) {
    const o = editor.objects().find((x) => x.id === id);
    if (!o) return;
    if (!pane || pane.id !== id) {
      closePane();
      const el = h("div", { class: "sa-pane", role: "dialog", "aria-label": "テキスト ウィンドウ", "data-keeps-text": "" });
      document.body.append(el);
      pane = { id, el, editing: false, last: "" };
      document.addEventListener("pointerdown", outside, true);
      // The stage may re-fit (ribbon, panes, zoom): follow the graphic while the pane is open.
      const follow = () => { if (!pane || pane.el !== el) return; place(); requestAnimationFrame(follow); };
      requestAnimationFrame(follow);
    }
    if (item != null) picked = item;
    renderPane();
    place();
    requestAnimationFrame(() => { place(); pane?.el.querySelector(`input[data-i="${picked ?? 0}"]`)?.focus(); });
  }
  function closePane() {
    if (!pane) return;
    pane.el.remove();
    pane = null;
    document.removeEventListener("pointerdown", outside, true);
  }
  function outside(event) {
    if (!pane) return;
    if (pane.el.contains(event.target) || event.target.closest?.(".ribbon, .rb-pop")) return;
    // A click on the SmartArt itself keeps the pane; elsewhere it closes.
    const node = event.target.closest?.("[data-el]");
    if (node?.dataset.el === pane.id) return;
    closePane();
  }
  /** Beside the SmartArt on screen (left of it when there is room, else right). */
  function place() {
    if (!pane) return;
    const node = document.querySelector(`#stageBody .hs-obj[data-el="${CSS.escape(pane.id)}"]`);
    const r = node?.getBoundingClientRect();
    const w = 300;
    if (!r || !r.width) { pane.el.style.left = "24px"; pane.el.style.top = "160px"; return; }
    const key = `${Math.round(r.left)},${Math.round(r.top)},${Math.round(r.right)},${window.innerWidth},${window.innerHeight}`;
    if (key === pane.last) return;
    pane.last = key;
    // Left of the graphic (over the thumbnails if need be), else right of it — never over the graphic itself.
    const left = r.left - w - 12 >= 8 ? r.left - w - 12 : window.innerWidth - r.right - 12 >= w ? r.right + 12 : 8;
    pane.el.style.left = `${left}px`;
    pane.el.style.top = `${Math.max(8, Math.min(window.innerHeight - 340, r.top))}px`;
  }
  function renderPane() {
    if (!pane) return;
    const o = editor.objects().find((x) => x.id === pane.id);
    if (!o) { closePane(); return; }
    const items = o.smartart.items;
    const focusAt = pane.el.contains(document.activeElement) ? Number(document.activeElement.dataset.i) : null;
    const caret = focusAt != null ? document.activeElement.selectionStart : null;
    const commitItems = (fn, { structural = false } = {}) => {
      if (structural) { setSmartart(o.id, (sa) => { fn(sa.items); return sa; }); return; }
      setSmartart(o.id, (sa) => { fn(sa.items); return sa; }, { undo: !typing, rerender: false });
      typing = true;
    };
    const rows = items.map((it, i) => {
      const input = h("input", { type: "text", class: "sa-line", value: it.text, "data-i": String(i), "aria-label": `項目 ${i + 1}（レベル ${it.level + 1}）`, placeholder: "[テキスト]", style: { "margin-left": `${it.level * 18}px` },
        onfocus: () => { picked = i; typing = false; },
        oninput: (event) => commitItems((list) => { list[i].text = event.target.value; }),
        onkeydown: (event) => {
          if (event.isComposing) return;
          if (event.key === "Tab") { event.preventDefault(); picked = i; shiftLevel(event.shiftKey ? -1 : 1); focus(i, event.target.selectionStart); return; }
          if (event.key === "Enter") { event.preventDefault(); const pos = event.target.selectionStart; const rest = event.target.value.slice(pos); picked = i + 1; commitItems((list) => { list[i].text = list[i].text.slice(0, pos); list.splice(i + 1, 0, { text: rest, level: list[i].level }); }, { structural: true }); focus(i + 1, 0); return; }
          if (event.key === "Backspace" && !event.target.value && items.length > 1) { event.preventDefault(); picked = Math.max(0, i - 1); commitItems((list) => { list.splice(i, 1); }, { structural: true }); focus(picked, Infinity); return; }
          if (event.key === "ArrowUp" && i > 0) { event.preventDefault(); focus(i - 1, event.target.selectionStart); }
          if (event.key === "ArrowDown" && i < items.length - 1) { event.preventDefault(); focus(i + 1, event.target.selectionStart); }
          if (event.key === "Escape") { event.preventDefault(); event.stopPropagation(); closePane(); }
        } });
      return h("li", { class: ["sa-row", `lv${it.level}`] }, h("span", { class: "sa-bullet" }, it.level ? "・" : "●"), input);
    });
    pane.el.replaceChildren(
      h("div", { class: "sa-pane-head" }, h("b", {}, "ここに文字を入力してください"), h("button", { type: "button", class: "btn btn-ghost btn-icon", "aria-label": "閉じる", onclick: () => closePane() }, "✕")),
      h("ul", { class: "sa-lines" }, rows),
      h("p", { class: "hint sa-hint" }, "Enterで項目を追加・Tabでレベルを下げる・Shift+Tabで上げる"),
      h("div", { class: "sa-pane-foot" }, h("small", {}, E.SMARTART_LAYOUTS[o.smartart.layout].label)));
    if (focusAt != null) focus(Math.min(focusAt, items.length - 1), caret);
  }
  function focus(i, pos) {
    requestAnimationFrame(() => {
      const input = pane?.el.querySelector(`input[data-i="${i}"]`);
      if (!input) return;
      input.focus();
      const p = pos === Infinity || pos == null ? input.value.length : Math.min(pos, input.value.length);
      input.setSelectionRange(p, p);
    });
  }

  // ---------------------------------------------------------------- galleries and the tab

  /** A small picture of a layout (drawn by the engine with the sample items). */
  function preview(layout, { color = "blue", w = 168, hh = 92 } = {}) {
    const o = E.normalizeObject({ id: `pv${layout}`, kind: "smartart", x: 0, y: 0, w: 640, h: 360, smartart: { layout, color, items: E.smartartSample(layout) } });
    const holder = h("div", { class: "sa-preview", style: { width: `${w}px`, height: `${hh}px` } });
    const inner = h("div", { class: "sa-preview-inner hs-slide", style: { transform: `scale(${Math.min(w / 640, hh / 360)})` } });
    inner.append(E.objectNode(o, { mode: "thumb", live: false }, [o], 1));
    holder.append(inner);
    return holder;
  }
  function gallery(pick, { title = "SmartArt グラフィックの選択", currentLayout = null } = {}) {
    return (close) => {
      let kind = "すべて";
      const grid = h("div", { class: "sa-gallery-grid" });
      const fill = () => grid.replaceChildren(...Object.entries(E.SMARTART_LAYOUTS).filter(([, v]) => kind === "すべて" || v.group === kind).map(([key, v]) => h("button", { type: "button", class: ["sa-gallery-item", key === currentLayout ? "on" : ""], "data-layout": key, title: `${v.group}：${v.label}`, onclick: () => { close(); pick(key); } }, preview(key), h("span", {}, v.label))));
      const kinds = h("div", { class: "sa-gallery-kinds", role: "tablist" }, ["すべて", ...E.SMARTART_GROUPS].map((g) => h("button", { type: "button", role: "tab", "aria-selected": String(g === kind), onclick: (event) => { kind = g; for (const b of kinds.children) b.setAttribute("aria-selected", String(b === event.currentTarget)); fill(); } }, g)));
      fill();
      return h("div", { class: "sa-gallery" }, h("div", { class: "rb-menu-head" }, title), h("div", { class: "sa-gallery-body" }, kinds, grid));
    };
  }
  function colorMenu() {
    const o = current();
    const now = o?.smartart.color || "blue";
    return (close) => h("div", { class: "sa-colors" }, h("div", { class: "rb-menu-head" }, "色の変更（テーマの色）"),
      h("div", { class: "sa-color-grid" }, Object.entries(E.SMARTART_COLORS).map(([key, label]) => h("button", { type: "button", class: ["sa-color", key === now ? "on" : ""], title: label, onclick: () => { close(); setField("color", key); } }, preview(o?.smartart.layout || "blocks", { color: key, w: 120, hh: 66 }), h("span", {}, label)))));
  }

  function designTab() {
    return [
      group("グラフィックの作成",
        drop("plus", "図形の|追加", "選んでいる項目（テキスト ウィンドウのカーソルの行）の後・前・上・下に図形を追加", () => menu([
          { label: "後に図形を追加", icon: "plus", run: () => addShape("after") },
          { label: "前に図形を追加", icon: "plus", run: () => addShape("before") },
          { label: "上に図形を追加", icon: "up", run: () => addShape("above") },
          { label: "下に図形を追加", icon: "down", run: () => addShape("below") },
        ]), { big: true, enabled: isSmartArt }),
        col(btn("bullet", "箇条書きの追加", "選んでいる項目の下に箇条書きを足す", () => addBullet(), { enabled: isSmartArt }),
          btn("taskPane", "テキスト ウィンドウ", "項目を1行ずつ入力する窓を開く・閉じる", () => (pane ? closePane() : openPane(current().id)), { enabled: isSmartArt, pressed: () => Boolean(pane) }),
          btn("flipH", "右から左", "左右を入れ替える", () => setField("rtl", !current()?.smartart.rtl), { enabled: isSmartArt, pressed: () => Boolean(current()?.smartart.rtl) })),
        col(btn("indentLess", "レベル上げ", "選んでいる項目のレベルを上げる（Shift+Tab）", () => shiftLevel(-1), { enabled: isSmartArt }),
          btn("indentMore", "レベル下げ", "選んでいる項目のレベルを下げる（Tab）", () => shiftLevel(1), { enabled: isSmartArt })),
        col(btn("up", "上へ移動", "選んでいる項目を前へ", () => moveItem(-1), { enabled: isSmartArt }),
          btn("down", "下へ移動", "選んでいる項目を後へ", () => moveItem(1), { enabled: isSmartArt }))),
      group("レイアウト", layoutRow(), drop("chevron", "", "すべてのレイアウト", () => gallery((key) => setField("layout", key), { title: "レイアウト", currentLayout: current()?.smartart.layout }), { enabled: isSmartArt })),
      group("SmartArt のスタイル",
        drop("fill", "色の|変更", "SmartArtの色（淡青・淡茶・グレー・線だけ）", () => colorMenu(), { big: true, enabled: isSmartArt }),
        col(...Object.entries(E.SMARTART_STYLES).map(([key, label]) => btn(key === "square" ? "slide" : key === "soft" ? "shapes" : "layout", label, `図形の角：${label}`, () => setField("style", key), { enabled: isSmartArt, pressed: () => (current()?.smartart.style || "round") === key })))),
      group("文字",
        col(btn("grow", "文字を大きく", "SmartArtの文字を大きくする", () => setField("fsScale", Math.min(2, Math.round(((current()?.smartart.fsScale || 1) + 0.1) * 10) / 10)), { enabled: isSmartArt }),
          btn("shrink", "文字を小さく", "SmartArtの文字を小さくする", () => setField("fsScale", Math.max(0.5, Math.round(((current()?.smartart.fsScale || 1) - 0.1) * 10) / 10)), { enabled: isSmartArt }))),
      group("リセット",
        btn("reset", "グラフィックの|リセット", "色・スタイル・文字の大きさ・向きを元に戻す（文字はそのまま）", () => { const o = current(); if (o) setSmartart(o.id, (sa) => ({ layout: sa.layout, items: sa.items })); }, { big: true, enabled: isSmartArt }),
        drop("convert", "変換", "図形に変換・テキストに変換", () => menu([
          { label: "図形に変換", icon: "shapes", run: () => toShapes() },
          { label: "テキストに変換", icon: "textbox", run: () => toText() },
        ]), { big: true, enabled: isSmartArt })),
    ];
  }
  function layoutRow() {
    const row = h("div", { class: "sa-layout-row" });
    for (const key of ["blocks", "process", "chevron", "cycle", "hierarchy", "pyramid"]) {
      const b = h("button", { type: "button", class: "sa-layout-quick", "data-layout": key, title: E.SMARTART_LAYOUTS[key].label, onclick: () => setField("layout", key) }, preview(key, { w: 64, hh: 36 }));
      updater(() => { b.classList.toggle("on", current()?.smartart.layout === key); b.disabled = !isSmartArt(); });
      row.append(b);
    }
    return row;
  }

  // Keep the text pane beside the SmartArt and in step with undo / other changes.
  editor.subscribe(() => {
    if (!pane) return;
    if (!editor.objects().some((x) => x.id === pane.id)) { closePane(); return; }
    if (!pane.el.contains(document.activeElement)) renderPane();
    place();
  });
  window.addEventListener("resize", () => place());

  const insertGallery = () => gallery((key) => insert(key));
  const convertGallery = () => gallery((key) => convertText(key), { title: "SmartArt に変換" });
  return { designTab, insertGallery, convertGallery, isSmartArt, openPane, closePane, toShapes, insert, convertText, itemsFromRich };
}
