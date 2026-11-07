// The ribbon's インタラクション tab: what only HTML can do with an object while presenting (and in the exported
// file) — how it answers the mouse (a lift, a zoom, a glow, a 3D tilt, the rest dimmed, a note), what a click on it
// does (a card with details, zooming in, turning over, showing and hiding other objects as tabs, a spotlight),
// and the motion it keeps (floating, a pulse, swaying, spinning, a sweep of light, bouncing). おまかせ puts the
// HTML motion on a whole slide or deck at once (htmlfx.mjs). The engine draws and plays them (engine/objects.js,
// engine/motion.js).

export function createInteractions(editor, app, kit) {
  const { E, h } = app;
  const { btn, drop, group, col, menu, openPop, updater, tabNow } = kit;

  const selected = () => editor.selectedObjects();
  const one = () => (selected().length === 1 ? selected()[0] : null);
  const any = () => selected().length > 0;
  const ids = () => selected().map((o) => o.id);
  const same = (key) => { const list = selected(); const v = list[0]?.[key]; return list.length && list.every((o) => o[key] === v) ? v : undefined; };
  const apply = (patch, message) => { if (!any()) { app.toast("部品を選んでください"); return; } editor.apply(patch, { ids: ids() }); if (message) app.toast(message); };

  // ---------------------------------------------------------------- the mouse

  function hoverMenu() {
    const now = same("hover");
    return menu([
      { head: "マウスを乗せたとき（発表中・HTML出力）" },
      { label: "なし", on: !now, run: () => apply({ hover: undefined }) },
      ...Object.entries(E.IX_HOVERS).map(([key, label]) => ({ label, on: now === key, run: () => apply({ hover: key }, `マウスを乗せると「${label}」ようにしました（▶ 試すで確かめられます）`) })),
    ]);
  }
  async function editTip() {
    const o = one();
    if (!o) { app.toast("説明を付ける部品を1つ選んでください"); return; }
    const value = await app.ask("マウスを乗せたときの説明", "発表中、この部品にマウスを乗せると吹き出しで出ます（空にすると外します）", o.tip || "");
    if (value == null) return;
    editor.apply({ tip: value.trim() ? value.trim().slice(0, 200) : undefined }, { ids: [o.id] });
    app.toast(value.trim() ? "マウスを乗せると説明が出るようにしました" : "説明を外しました");
  }

  // ---------------------------------------------------------------- a click

  const actionOf = () => one()?.action;
  function clickMenu() {
    const o = one();
    const type = o?.action?.type;
    if (!o) return menu([{ head: "部品を1つ選ぶと、クリックしたときの動きを選べます" }]);
    return menu([
      { head: "クリックしたとき（発表中・HTML出力）" },
      { label: "なし", on: !type, run: () => setAction(undefined) },
      { label: "詳細を開く…（カード・根拠パネル）", icon: "popup", on: type === "popup", run: () => openForm(popupForm) },
      { label: "拡大して見せる", icon: "zoomClick", on: type === "zoom", run: () => setAction({ type: "zoom" }, "クリックすると拡大して見せます（もう一度クリックで戻る）") },
      { label: "裏返す…（裏の文字）", icon: "cardFlip", on: type === "flip", run: () => openForm(flipForm) },
      { label: "ほかの部品を表示・非表示…（タブ）", icon: "reveal", on: type === "reveal", run: () => openForm(revealForm) },
      { label: "スポットライトを当てる", icon: "spot", on: type === "spot", run: () => setAction({ type: "spot" }, "クリックすると、ほかを暗くしてこの部品に注目させます") },
      "-",
      { label: "スライドへ移動・Webページを開く…", icon: "link", on: ["next", "prev", "first", "last", "slide", "url", "end"].includes(type), run: () => setTimeout(() => kit.editLink(), 0) },
    ]);
  }
  function setAction(action, message) {
    const o = one();
    if (!o) return;
    editor.apply({ action }, { ids: [o.id] });
    app.toast(message || (action ? "クリックしたときの動きを設定しました" : "クリックしたときの動きを外しました"));
  }
  const anchor = () => document.querySelector('[data-rb="ixClick"]') || document.getElementById("ribbon");
  function openForm(build) { const o = one(); if (o) setTimeout(() => openPop(anchor(), (close) => build(o, close)), 0); }
  const foot = (save, close) => h("div", { class: "rb-form-foot" }, h("button", { type: "button", class: "btn btn-sm", onclick: () => close() }, "やめる"), h("button", { type: "button", class: "btn btn-primary btn-sm", onclick: save }, "設定する"));

  /** 詳細を開く: a title, the words, a breakdown (label,value per line) and where the figures come from. */
  function popupForm(o, close) {
    const cur = o.action?.type === "popup" ? o.action : {};
    const title = h("input", { type: "text", placeholder: "見出し（例：内訳）", value: cur.title || "", maxlength: 80 });
    const text = h("textarea", { rows: 3, placeholder: "カードに出す文" }, cur.text ? E.richToText(cur.text) : "");
    const rows = h("textarea", { rows: 3, placeholder: "内訳（1行に「ラベル,値」。例：東日本,42%）" }, (cur.rows || []).map((r) => `${r.label},${r.value}`).join("\n"));
    const source = h("input", { type: "text", placeholder: "出所（任意）", value: cur.source || "", maxlength: 120 });
    const save = () => {
      const list = rows.value.split("\n").map((line) => line.split(/[,，\t]/)).filter(([label]) => label && label.trim()).map(([label, ...rest]) => ({ label: label.trim(), value: rest.join(",").trim() }));
      const action = { type: "popup", title: title.value.trim() || undefined, text: text.value.trim() ? E.textToRich(text.value.trim()) : undefined, rows: list.length ? list : undefined, source: source.value.trim() || undefined };
      if (!action.title && !action.text && !action.rows) { app.toast("見出しか文か内訳を入れてください"); return; }
      editor.apply({ action }, { ids: [o.id] });
      close();
      app.toast(list.length || action.source ? "クリックすると右から根拠パネルが開きます" : "クリックすると詳細カードが開きます");
    };
    return h("div", { class: "rb-form ix-form" }, h("b", {}, "クリックで開く詳細"), title, text, rows, source,
      h("p", { class: "rb-note" }, "内訳か出所を入れると、右から根拠パネルとして開きます。"), foot(save, close));
  }

  /** 裏返す: the words on the back and its colour. */
  function flipForm(o, close) {
    const cur = o.action?.type === "flip" ? o.action : {};
    const back = h("textarea", { rows: 4, placeholder: "裏の文字（例：答え・詳しい数字）" }, cur.back ? E.richToText(cur.back) : "");
    let fill = cur.fill || null;
    const swatches = h("div", { class: "rb-swatches ix-swatches" }, E.PALETTE.fill.filter(([c]) => c !== "#1f3864" && c !== "#808080").map(([c, label]) => {
      const b = h("button", { type: "button", class: ["rb-sw", fill === c ? "on" : ""], title: label, style: { "--c": c }, onclick: () => { fill = c; for (const x of swatches.children) x.classList.toggle("on", x === b); } });
      return b;
    }));
    const save = () => {
      if (!back.value.trim()) { app.toast("裏の文字を入れてください"); return; }
      editor.apply({ action: { type: "flip", back: E.textToRich(back.value.trim()), fill: fill || undefined } }, { ids: [o.id] });
      close();
      app.toast("クリックすると裏返ります（もう一度クリックで表に戻る）");
    };
    return h("div", { class: "rb-form ix-form" }, h("b", {}, "クリックで裏返す"), back, h("span", { class: "rb-note" }, "裏の色（選ばなければ表と同じ）"), swatches, foot(save, close));
  }

  /** ほかの部品を表示・非表示: the objects this one shows; "tabs" hides what the slide's other buttons showed. */
  function revealForm(o, close) {
    const cur = o.action?.type === "reveal" ? o.action : { targets: [] };
    const list = editor.objects().filter((x) => x.id !== o.id);
    const picked = new Set(cur.targets);
    // Objects picked together with the button are offered first.
    for (const x of selected()) if (x.id !== o.id && !cur.targets.length) picked.add(x.id);
    const boxes = h("div", { class: "ix-targets" }, list.map((x) => h("label", {}, h("input", { type: "checkbox", value: x.id, checked: picked.has(x.id) || null, onchange: (event) => { if (event.target.checked) picked.add(x.id); else picked.delete(x.id); } }), h("span", {}, E.objectName(x, editor.objects().indexOf(x))), x.group ? h("small", {}, "（グループ）") : null)));
    const only = h("input", { type: "checkbox", checked: cur.only || null });
    const save = () => {
      if (!picked.size) { app.toast("表示する部品を選んでください"); return; }
      editor.apply({ action: { type: "reveal", targets: [...picked], only: only.checked || undefined } }, { ids: [o.id] });
      close();
      app.toast("発表中、選んだ部品は隠れていて、この部品のクリックで出ます");
    };
    return h("div", { class: "rb-form ix-form" }, h("b", {}, "クリックで表示する部品（もう一度で隠す）"), boxes,
      h("label", { class: "ix-check" }, only, h("span", {}, "タブのように：ほかのボタンが出したものは隠す")), foot(save, close));
  }

  // ---------------------------------------------------------------- motion that keeps going

  function loopMenu() {
    const now = same("loop");
    return menu([
      { head: "スライドを表示している間ずっと（発表中・HTML出力）" },
      { label: "なし", on: !now, run: () => apply({ loop: undefined }) },
      ...Object.entries(E.IX_LOOPS).map(([key, label]) => ({ label, on: now === key, run: () => apply({ loop: key }, `「${label}」動きを付けました（▶ 試すで確かめられます）`) })),
    ]);
  }

  // ---------------------------------------------------------------- the tab

  function tab() {
    const status = h("div", { class: "ix-status" });
    updater(() => {
      const o = one();
      const parts = [];
      if (o?.hover) parts.push(`マウス：${E.IX_HOVERS[o.hover]}`);
      if (o?.tip) parts.push("説明あり");
      if (o?.action && E.IX_CLICKS[o.action.type]) parts.push(`クリック：${E.IX_CLICKS[o.action.type]}`);
      if (o?.loop) parts.push(`ずっと：${E.IX_LOOPS[o.loop]}`);
      if (o?.overAction) parts.push("マウスの通過で動作");
      status.textContent = !any() ? "部品を選ぶと設定できます" : selected().length > 1 ? `${selected().length}個を選択中` : parts.join("・") || "まだ設定していません";
    });
    const clickBtn = drop("popup", "クリック|したとき", "クリックしたとき：詳細を開く・拡大・裏返す・表示の切り替え（タブ）・スポットライト", () => clickMenu(), { big: true, enabled: () => Boolean(one()) });
    clickBtn.dataset.rb = "ixClick";
    return [
      group("おまかせ",
        btn("magic", "HTMLの動き|を付ける", "このスライドに、HTMLならではの動きをおまかせで付ける（タイトル・グラフ・数字・線・カード・写真）", () => app.enhance("slide"), { big: true }),
        col(btn("applyAll", "すべてのスライドに", "資料のすべてのスライドに、おまかせで付ける", () => app.enhance("all")),
          btn("check", "取り込み時に自動", "PowerPointを見た目どおりに取り込んだとき、HTMLの動きを自動で付ける（⌘Zで外せます）", () => app.setAutoHtml(!app.autoHtml()), { pressed: () => app.autoHtml() }),
          btn("play", "▶ 試す", "このスライドから発表して、マウスやクリックの動きを確かめる", () => app.present()))),
      group("マウスを乗せたとき",
        drop("hover", "反応", "マウスを乗せたとき：浮き上がる・拡大・光る・3Dで傾く・ほかを薄くする", () => hoverMenu(), { big: true, enabled: any }),
        col(btn("tip", "説明を出す…", "マウスを乗せると吹き出しで説明を出す", editTip, { enabled: () => Boolean(one()), pressed: () => Boolean(one()?.tip) }))),
      group("クリックしたとき", clickBtn,
        col(btn("trash", "外す", "クリックしたときの動きを外す", () => setAction(undefined), { enabled: () => Boolean(actionOf()) }))),
      group("ずっと動く",
        drop("loop", "動き", "スライドを表示している間ずっと：ふわふわ・鼓動・ゆらゆら・回転・光が走る・弾む", () => loopMenu(), { big: true, enabled: any })),
      group("いまの設定", status),
    ];
  }

  /** Marks on the stage (while this tab is open): what each object does under the mouse, on a click, all along. */
  function badges(list, k) {
    if (tabNow() !== "interact" || app.state().view !== "single") return [];
    const out = [];
    for (const o of list) {
      if (o.hidden) continue;
      const marks = [o.hover ? "マウス" : "", o.tip ? "説明" : "", o.action && E.IX_CLICKS[o.action.type] ? "クリック" : "", o.loop ? "ずっと" : "", o.overAction ? "通過" : ""].filter(Boolean);
      const targetOf = list.filter((x) => x.action?.type === "reveal" && x.action.targets.includes(o.id));
      if (targetOf.length) marks.push("クリックで出る");
      if (!marks.length) continue;
      const b = E.bounds(o, list);
      out.push(h("div", { class: "ed-ix-badge", style: { left: `${b.x * k + b.w * k - 6}px`, top: `${b.y * k - 6}px` } }, marks.join("・")));
    }
    return out;
  }

  editor.overlay(badges);
  return { tab };
}
