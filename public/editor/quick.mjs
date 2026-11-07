// PowerPoint's 検索 box (⌥Q / Alt+Q) and クイック アクセス ツール バー. The search finds any command of the ribbon,
// whichever tab it is on (and the ファイル menu's), and runs it as a click on the button itself would. The toolbar
// keeps the commands used most within one click: shown from the ribbon's display options, above or below the
// ribbon, with common commands to tick, any ribbon button added by a right-click, and その他のコマンド to order them.

/** Text folded for matching: full-width to half-width, katakana to hiragana, lower case, one space. */
export function fold(text) {
  return String(text || "").normalize("NFKC").toLowerCase()
    .replace(/[ァ-ヶ]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0x60))
    .replace(/\s+/g, " ").trim();
}

/**
 * The commands matching every word typed, best first: the name starting with the words, the name holding them,
 * then the description, then the tab and group. Commands that cannot be used now come after the rest.
 */
export function searchCommands(query, commands, limit = 12) {
  const words = fold(query).split(" ").filter(Boolean);
  if (!words.length) return [];
  const q = words.join(" ");
  const scored = [];
  (commands || []).forEach((c, i) => {
    const label = fold(c.label);
    const title = fold(c.title);
    const place = fold(`${c.tabLabel || ""} ${c.group || ""}`);
    if (!words.every((w) => label.includes(w) || title.includes(w) || place.includes(w))) return;
    const score = label.startsWith(q) ? 0 : label.includes(q) ? 1 : words.every((w) => label.includes(w)) ? 2 : title.includes(q) ? 3 : 4;
    scored.push([score, c.disabled ? 1 : 0, i, c]);
  });
  return scored.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]).slice(0, limit).map((x) => x[3]);
}

/** The common commands the toolbar's menu offers to tick (PowerPoint's list). */
export const QAT_BUILTINS = {
  new: { label: "新規", icon: "plus" },
  open: { label: "開く", icon: "pane" },
  save: { label: "上書き保存", icon: "save" },
  undo: { label: "元に戻す", icon: "undo" },
  redo: { label: "やり直し", icon: "redo" },
  fromStart: { label: "最初から", icon: "showStart" },
  fromHere: { label: "現在のスライドから", icon: "showHere" },
  print: { label: "印刷", icon: "print" },
  html: { label: "HTML出力", icon: "play" },
};
export const QAT_DEFAULT_ITEMS = ["save", "undo", "redo", "fromStart"];
export const qatKey = (item) => (typeof item === "string" ? item : item?.ref);

/**
 * A stored toolbar, tidied: { show, below, items } where an item is a common command's key or a ribbon command
 * { ref: "tab|group|title", label, icon }; no repeats, at most 30.
 */
// Ribbon buttons whose tooltip changed: a toolbar saved with the old one keeps working.
const QAT_RENAMED = {
  "view|表示/非表示|1cmごとの線を表示": "view|表示/非表示|グリッド線を表示（間隔はグリッドとガイドの設定で）",
  "view|表示/非表示|0.25cmごとに吸着": "view|表示/非表示|グリッド線に吸着（間隔はグリッドとガイドの設定で）",
  "design|ユーザー設定|ワイド画面（16:9・13.33×7.5インチ）：SEJテンプレートの大きさです": "design|ユーザー設定|スライドのサイズ（ワイド画面 16:9）とスライド番号の開始番号",
  "design|ユーザー設定|スライドのサイズ（ワイド画面 16:9・SEJテンプレート）とスライド番号の開始番号": "design|ユーザー設定|スライドのサイズ（ワイド画面 16:9）とスライド番号の開始番号",
  "picture|調整|トリミング・修整・枠線を元に戻す": "picture|調整|トリミング・修整・枠線・影を元に戻す",
};

export function normalizeQat(raw) {
  const out = { show: Boolean(raw?.show), below: Boolean(raw?.below), items: [] };
  const seen = new Set();
  for (const it of Array.isArray(raw?.items) ? raw.items : QAT_DEFAULT_ITEMS) {
    let item = null;
    if (typeof it === "string") item = QAT_BUILTINS[it] ? it : null;
    else if (it && typeof it.ref === "string" && /^[A-Za-z0-9]+\|[^|]*\|./.test(it.ref)) {
      item = { ref: (QAT_RENAMED[it.ref] || it.ref).slice(0, 300), label: String(it.label || "").trim().slice(0, 60) || it.ref.split("|").pop().slice(0, 40), icon: /^[A-Za-z0-9]{1,30}$/.test(it.icon || "") ? it.icon : "" };
    }
    const key = qatKey(item);
    if (!item || seen.has(key)) continue;
    seen.add(key);
    out.items.push(item);
    if (out.items.length >= 30) break;
  }
  return out;
}

/**
 * kit: { h, ico, app, openPop, closePop, menu, commands() (the ribbon's commands now), run(ref), render() (the
 * ribbon again), tab() (the tab shown), watch(fn) (kept up to date with the ribbon) }.
 */
export function createQuick(kit) {
  const { h, ico, app } = kit;
  const load = (key, fallback) => { try { return JSON.parse(localStorage.getItem(key) || "null") ?? fallback; } catch { return fallback; } };
  const store = (key, value) => { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* private window */ } };
  let qat = normalizeQat(load("hpptx-qat", null));

  function set(patch) {
    qat = normalizeQat({ ...qat, ...patch });
    store("hpptx-qat", qat);
    kit.render();
  }
  function add(item) {
    if (qat.items.some((x) => qatKey(x) === qatKey(item))) return;
    set({ show: true, items: [...qat.items, item] });
    app.toast("クイック アクセス ツール バーに追加しました");
  }
  function remove(key) { set({ items: qat.items.filter((x) => qatKey(x) !== key) }); }
  function toggle(key) { if (qat.items.includes(key)) remove(key); else set({ show: true, items: [...qat.items, key] }); }

  // ---------------------------------------------------------------- 検索
  let index = [];
  let results = [];
  let active = 0;
  const input = h("input", {
    type: "search", class: "rb-search-input", placeholder: "検索（⌥Q）", "aria-label": "コマンドを検索（Alt+Q・⌥Q）", autocomplete: "off", spellcheck: "false",
    onfocus: () => { index = kit.commands(); active = 0; show(); },
    oninput: () => { active = 0; show(); },
    onkeydown: (event) => {
      if (event.key === "ArrowDown" || event.key === "ArrowUp") {
        event.preventDefault();
        if (results.length) { active = (active + (event.key === "ArrowDown" ? 1 : -1) + results.length) % results.length; show(); }
      } else if (event.key === "Enter") { event.preventDefault(); if (results[active]) pick(results[active]); }
      else if (event.key === "Escape") { event.preventDefault(); input.value = ""; kit.closePop(true); input.blur(); }
    },
  });
  const searchBox = h("div", { class: "rb-search-box", role: "search" }, ico("find", 14), input);

  const recent = () => load("hpptx-recent-commands", []).map((ref) => index.find((c) => c.ref === ref)).filter(Boolean).slice(0, 6);
  function remember(ref) { store("hpptx-recent-commands", [ref, ...load("hpptx-recent-commands", []).filter((r) => r !== ref)].slice(0, 12)); }

  function show() {
    const q = input.value.trim();
    results = q ? searchCommands(q, index, 12) : recent();
    kit.openPop(searchBox, () => {
      const list = h("div", { class: "rb-search-results", role: "listbox", "aria-label": "コマンドの検索結果" });
      if (!results.length) list.append(h("p", { class: "rb-search-empty" }, q ? `「${q}」に合うコマンドが見つかりません` : "コマンドの名前を入れると、どのタブにあっても見つけて実行できます（例：透明度、ガイド、アニメーション）"));
      else if (!q) list.append(h("div", { class: "rb-menu-head" }, "最近使ったコマンド"));
      results.forEach((c, k) => list.append(h("button", {
        type: "button", role: "option", class: ["rb-search-item", k === active ? "on" : "", c.disabled ? "off" : ""], "aria-selected": String(k === active), title: c.disabled ? `${c.title}（今は使えません）` : c.title,
        onmousedown: (event) => event.preventDefault(), onclick: () => pick(c),
      }, ico(c.icon || "chevron", 16), h("span", { class: "rb-search-label" }, c.label), h("small", {}, [c.tabLabel, c.group].filter(Boolean).join(" › ")))));
      return list;
    });
  }
  function pick(c) {
    kit.closePop(true);
    input.value = "";
    input.blur();
    remember(c.ref);
    kit.run(c.ref);
  }
  function focusSearch() { input.focus(); input.select(); }

  // ---------------------------------------------------------------- クイック アクセス ツール バー
  const RUN = {
    new: () => app.goCreate(), open: () => app.openLibrary(), save: () => app.saveDeck(), undo: () => app.undo(), redo: () => app.redo(),
    fromStart: () => app.openPresenter(0), fromHere: () => app.openPresenter(app.index()), print: () => app.printPdf(), html: () => app.exportHtml(),
  };
  const ENABLED = { undo: () => app.canUndo(), redo: () => app.canRedo() };
  const info = (item) => (typeof item === "string" ? { ...QAT_BUILTINS[item], run: RUN[item], enabled: ENABLED[item] } : { label: item.label, icon: item.icon, run: (button) => kit.run(item.ref, { from: button }) });

  function toolbar() {
    if (!qat.show) return null;
    const bar = h("div", { class: ["rb-qat", qat.below ? "below" : ""], role: "toolbar", "aria-label": "クイック アクセス ツール バー" });
    for (const item of qat.items) {
      const it = info(item);
      const b = h("button", { type: "button", class: "rb-qat-btn", title: it.label, "aria-label": it.label, "data-qat": qatKey(item), onclick: (event) => { event.preventDefault(); it.run(event.currentTarget); } }, ico(it.icon || "chevron", 16));
      if (it.enabled) kit.watch(() => { b.disabled = !it.enabled(); });
      bar.append(b);
    }
    bar.append(h("button", { type: "button", class: "rb-qat-more", title: "クイック アクセス ツール バーのユーザー設定", "aria-label": "クイック アクセス ツール バーのユーザー設定", "aria-haspopup": "true", onclick: (event) => kit.openPop(event.currentTarget, customizeMenu()) }, "▾"));
    return bar;
  }
  const where = () => (!qat.show ? null : qat.below ? "below" : "above");
  const placeItem = () => ({ label: qat.below ? "リボンの上に表示" : "リボンの下に表示", icon: "ribbon", run: () => set({ below: !qat.below }) });

  function customizeMenu() {
    return kit.menu([
      { head: "クイック アクセス ツール バーのユーザー設定" },
      ...Object.entries(QAT_BUILTINS).map(([key, b]) => ({ label: b.label, icon: b.icon, on: qat.items.includes(key), run: () => toggle(key) })),
      "-",
      { label: "その他のコマンド…", icon: "selectAll", run: () => customize() },
      placeItem(),
      { label: "クイック アクセス ツール バーを非表示にする", icon: "eyeOff", run: () => set({ show: false }) },
    ]);
  }
  /** The ribbon's display options gain the toolbar (shown or not), as PowerPoint's. */
  const optionItems = () => ["-", { label: "クイック アクセス ツール バーを表示する", icon: "pin", on: qat.show, run: () => set({ show: !qat.show }) }];

  /** A right-click on a ribbon button adds it to the toolbar (or takes it off); on the toolbar, takes it off. */
  function contextMenu(event, tab) {
    const b = event.target.closest?.("button.rb-qat-btn, button.rb-btn:not(.rb-folded)");
    if (!b) return false;
    if (b.classList.contains("rb-qat-btn")) {
      event.preventDefault();
      const key = b.dataset.qat;
      kit.openPop(b, kit.menu([
        { label: "クイック アクセス ツール バーから削除", icon: "trash", run: () => remove(key) },
        { label: "クイック アクセス ツール バーのユーザー設定…", icon: "selectAll", run: () => customize() },
        placeItem(),
      ]));
      return true;
    }
    const title = b.getAttribute("title") || "";
    if (!title) return false;
    event.preventDefault();
    const group = b.closest(".rb-group, .rb-fold")?.querySelector(":scope > .rb-label")?.textContent || "";
    const ref = `${tab}|${group}|${title}`;
    const has = qat.items.some((x) => qatKey(x) === ref);
    const label = (b.querySelector(":scope > span")?.textContent || title).trim();
    const icon = b.querySelector("svg")?.getAttribute("data-ico") || "";
    kit.openPop(b, kit.menu([
      has ? { label: "クイック アクセス ツール バーから削除", icon: "trash", run: () => remove(ref) } : { label: "クイック アクセス ツール バーに追加", icon: "plus", run: () => add({ ref, label, icon }) },
      { label: "クイック アクセス ツール バーのユーザー設定…", icon: "selectAll", run: () => customize() },
      placeItem(),
    ]));
    return true;
  }

  /** その他のコマンド: every command (common ones first) on the left, the toolbar's on the right, in order. */
  function customize() {
    const all = kit.commands();
    let items = qat.items.slice();
    const filter = h("input", { type: "search", class: "qat-filter", placeholder: "コマンドを探す（例：配置、透明度）", "aria-label": "コマンドを探す" });
    const left = h("select", { size: 14, class: "qat-all", "aria-label": "コマンドの選択" });
    const right = h("select", { size: 14, class: "qat-items", "aria-label": "クイック アクセス ツール バーのコマンド" });
    const below = h("input", { type: "checkbox", class: "qat-below", checked: qat.below || null });
    const choices = new Map();
    const fillLeft = () => {
      const q = filter.value.trim();
      choices.clear();
      const builtins = Object.entries(QAT_BUILTINS).filter(([, b]) => !q || fold(b.label).includes(fold(q))).map(([key, b]) => [key, `${b.label}（よく使うコマンド）`, key]);
      const ribbon = (q ? searchCommands(q, all, 400) : all).map((c) => [c.ref, `${c.label}（${[c.tabLabel, c.group].filter(Boolean).join(" › ")}）`, { ref: c.ref, label: c.label, icon: c.icon }]);
      left.replaceChildren(...[...builtins, ...ribbon].map(([value, text, item]) => { choices.set(value, item); return h("option", { value }, text); }));
    };
    const fillRight = (keep = null) => {
      right.replaceChildren(...items.map((item) => h("option", { value: qatKey(item) }, info(item).label)));
      if (keep != null) right.value = keep;
    };
    const move = (d) => {
      const k = items.findIndex((x) => qatKey(x) === right.value);
      if (k < 0 || k + d < 0 || k + d >= items.length) return;
      [items[k], items[k + d]] = [items[k + d], items[k]];
      fillRight(qatKey(items[k + d]));
    };
    const addPicked = () => {
      const item = choices.get(left.value);
      if (!item || items.some((x) => qatKey(x) === qatKey(item))) return;
      items.push(item);
      fillRight(qatKey(item));
    };
    filter.addEventListener("input", fillLeft);
    left.addEventListener("dblclick", addPicked);
    const btn = (text, cls, fn) => h("button", { type: "button", class: `btn btn-sm ${cls}`, onclick: fn }, text);
    const dialog = h("dialog", { class: "qat-dialog", "aria-label": "クイック アクセス ツール バーのユーザー設定" },
      h("div", { class: "dialog-head" }, h("h3", {}, "クイック アクセス ツール バーのユーザー設定"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body" },
        h("div", { class: "qat-cols" },
          h("div", { class: "qat-col" }, h("b", {}, "コマンドの選択"), filter, left),
          h("div", { class: "qat-mid" }, btn("追加 »", "qat-add", addPicked), btn("« 削除", "qat-remove", () => { items = items.filter((x) => qatKey(x) !== right.value); fillRight(); })),
          h("div", { class: "qat-col" }, h("b", {}, "クイック アクセス ツール バー"), right,
            h("div", { class: "qat-order" }, btn("▲ 上へ", "qat-up", () => move(-1)), btn("▼ 下へ", "qat-down", () => move(1)), btn("リセット", "qat-reset", () => { items = QAT_DEFAULT_ITEMS.slice(); fillRight(); })))),
        h("label", { class: "sh-choice" }, below, h("span", {}, "クイック アクセス ツール バーをリボンの下に表示する"))),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary qat-ok", onclick: () => { set({ show: true, below: below.checked, items }); dialog.close(); } }, "OK")));
    fillLeft();
    fillRight();
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
    filter.focus();
  }

  return { searchBox, focusSearch, toolbar, where, optionItems, contextMenu, customize, remember, setPlace: (patch) => set(patch), get state() { return qat; } };
}
