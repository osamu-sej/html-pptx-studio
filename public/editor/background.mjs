// デザイン → 背景の書式設定 (PowerPoint's Format Background): this slide's background — none (the theme's), a colour
// (the theme's colours or any other), or a picture with its transparency (filling the slide or tiled) — changed as you
// watch, then 「すべてに適用」 or 「背景のリセット」.

export function createBackground(app) {
  const { h, E } = app;
  let dialog = null;
  let first = true;

  const current = () => E.normalizeBackground(app.slide()?.background) || {};
  /** Apply a change to this slide (one undo step for the whole time the pane is open). */
  function apply(patch, target = app.index()) {
    const next = { ...current(), ...patch };
    for (const [k, v] of Object.entries(next)) if (v == null || v === false || v === "") delete next[k];
    app.setBackground(target, next, { undo: first });
    first = false;
    refresh();
  }

  function refresh() {
    if (!dialog) return;
    const bg = current();
    const kind = bg.image ? "image" : bg.color ? "color" : "none";
    for (const input of dialog.querySelectorAll("[name=bg-kind]")) input.checked = input.value === kind;
    for (const sw of dialog.querySelectorAll(".bg-sw")) sw.classList.toggle("on", sw.dataset.color === (bg.color || "#ffffff"));
    dialog.querySelector(".bg-colors").classList.toggle("is-off", kind === "image");
    dialog.querySelector(".bg-picture").classList.toggle("is-off", kind !== "image");
    const range = dialog.querySelector(".bg-transparency");
    range.value = String(Math.round((bg.transparency || 0) * 100));
    dialog.querySelector(".bg-transparency-v").textContent = `${range.value}%`;
    dialog.querySelector(".bg-tile").checked = Boolean(bg.tile);
    dialog.querySelector(".bg-where").textContent = `${app.index() + 1}枚目のスライド`;
  }

  async function pickPicture() {
    const [file] = await app.pickFiles("image/png,image/jpeg,image/webp,image/gif,image/svg+xml");
    if (!file) return;
    try {
      const { src } = await app.storePicture(file);
      apply({ image: src, color: undefined });
    } catch (error) {
      app.toast(`画像を読み込めませんでした（${error.message}）`);
    }
  }

  function open() {
    if (!app.slide()) return;
    if (dialog) { dialog.querySelector("button")?.focus(); return; }
    first = true;
    const swatches = E.PALETTE.fill.map(([color, name]) => h("button", { type: "button", class: "bg-sw", title: name, "aria-label": name, "data-color": color, style: { background: color }, onclick: () => apply({ color: color === E.themeColors(E.PALETTE.sheet?.themeId).bg1 ? undefined : color, image: undefined, transparency: undefined, tile: undefined }) }));
    const kind = (value, label) => h("label", { class: "sh-choice" }, h("input", { type: "radio", name: "bg-kind", value, onchange: () => {
      if (value === "none") apply({ color: undefined, image: undefined, transparency: undefined, tile: undefined });
      else if (value === "color") apply({ image: undefined, transparency: undefined, tile: undefined, color: current().color || E.themeColors(E.PALETTE.sheet?.themeId).bg2 });
      else if (!current().image) pickPicture().then(refresh);
    } }), h("span", {}, label));
    dialog = h("dialog", { class: "bg-dialog", "aria-label": "背景の書式設定" },
      h("div", { class: "dialog-head" }, h("h3", {}, "背景の書式設定"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body sh-form-col" },
        h("p", { class: "hint bg-where" }),
        h("b", {}, "塗りつぶし"),
        kind("none", "塗りつぶしなし（テーマの背景）"),
        kind("color", "塗りつぶし（単色）"),
        h("div", { class: "bg-colors" }, ...swatches),
        kind("image", "塗りつぶし（図）"),
        h("div", { class: "bg-picture" },
          h("button", { type: "button", class: "btn btn-sm bg-file", onclick: pickPicture }, "ファイルから図を選ぶ…"),
          h("label", { class: "sh-inline" }, "透明度 ", h("input", { type: "range", min: "0", max: "95", step: "5", class: "bg-transparency", oninput: (event) => apply({ transparency: Number(event.target.value) / 100 }) }), h("span", { class: "bg-transparency-v" })),
          h("label", { class: "sh-choice" }, h("input", { type: "checkbox", class: "bg-tile", onchange: (event) => apply({ tile: event.target.checked }) }), h("span", {}, "図をテクスチャとして並べる"))),
        h("p", { class: "hint" }, "テーマの色はテーマを変えると一緒に変わります。")),
      h("div", { class: "dialog-foot" },
        h("button", { type: "button", class: "btn btn-ghost bg-reset", onclick: () => apply({ color: undefined, image: undefined, transparency: undefined, tile: undefined }) }, "背景のリセット"),
        h("button", { type: "button", class: "btn bg-all", onclick: () => { app.setBackground("all", current(), { undo: true }); app.toast("すべてのスライドに同じ背景を適用しました"); refresh(); } }, "すべてに適用"),
        h("button", { type: "button", class: "btn btn-primary", onclick: () => dialog.close() }, "閉じる")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => { dialog.remove(); dialog = null; });
    // Beside the slide, not over it (the slide changes as you choose).
    dialog.show();
    refresh();
  }

  return { open, refresh, isOpen: () => Boolean(dialog) };
}
