// More of PowerPoint's 挿入 and ホーム: ズーム (slide, summary, section), カメオ (live camera), スクリーンショット,
// 数式 (TeX, drawn as MathML), ワードアート (text styles in the theme's colours), スライド番号 and 日付 fields, 段組み,
// 文字種の変換 (全角/半角, 大文字/小文字) and 音声入力 (the browser's speech recognition) — and the 「ズーム」
// contextual tab of a zoom object.

/** ワードアートのスタイル in the theme: テキスト 1 and the accents, bold, marked, underlined or spaced. */
function wordArtStyles(E) {
  const c = E.themeColors(E.PALETTE.sheet?.themeId);
  return [
    { id: "accentBold", label: "アクセント 1・太字（見出し）", text: { color: c.accent1, bold: true, fs: 88 } },
    { id: "textBold", label: "テキスト 1・太字", text: { color: c.tx1, bold: true, fs: 80 } },
    { id: "mutedLight", label: "テキスト 2・細め（補足）", text: { color: c.tx2, fs: 64 } },
    { id: "marker", label: "テキスト 1・アクセントのマーカー", text: { color: c.tx1, bold: true, fs: 72 }, marker: E.lumColor(c.accent1, 40000, 60000) },
    { id: "underline", label: "アクセント 2・下線", text: { color: c.accent2, bold: true, underline: true, fs: 72 } },
    { id: "spaced", label: "アクセント 1・字間を広く", text: { color: c.accent1, bold: true, fs: 72, ls: 0.2 } },
  ];
}
const EQUATIONS = [
  ["二次方程式の解", "x = \\frac{-b \\pm \\sqrt{b^2 - 4ac}}{2a}"],
  ["合計", "\\sum_{i=1}^{n} x_i"],
  ["平均", "\\bar{x} = \\frac{1}{n} \\sum_{i=1}^{n} x_i"],
  ["前年比", "\\text{前年比} = \\frac{\\text{今年}}{\\text{前年}} \\times 100"],
  ["三平方の定理", "a^2 + b^2 = c^2"],
  ["面積", "S = \\pi r^2"],
];

const escapeText = (text) => String(text ?? "").replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[c]);

export function createExtras(editor, app, kit) {
  const { E, h } = app;
  const { btn, drop, group, col, menu, updater, openPop } = kit;
  const selected = () => editor.selectedObjects();
  const zoomSel = () => { const list = selected(); return list.length === 1 && list[0].kind === "zoom" ? list[0] : null; };
  const slides = () => app.deck()?.slides || [];
  const center = (w, hh) => ({ x: Math.round((E.W - w) / 2), y: Math.round((E.H - hh) / 2) + 20, w, h: hh });

  // ---------------------------------------------------------------- ズーム

  /** A grid of thumbnails to pick slides from (one or several). */
  function slidePicker({ title, multiple = false, onPick, filter = () => true }) {
    return (close) => {
      const chosen = new Set();
      const grid = h("div", { class: "zm-grid" });
      slides().forEach((slide, i) => {
        if (!filter(slide, i)) return;
        const el = E.render(slide, { ...app.renderOptions(), deck: app.deck(), index: i, mode: "thumb" });
        const cell = h("button", { type: "button", class: "zm-cell", "data-index": String(i), title: `${i + 1}. ${E.strip(slide.title || "")}`,
          onclick: () => {
            if (!multiple) { close(); onPick([i]); return; }
            if (chosen.has(i)) chosen.delete(i); else chosen.add(i);
            cell.classList.toggle("on", chosen.has(i));
            ok.disabled = !chosen.size;
            ok.textContent = chosen.size ? `${chosen.size}枚で作る` : "スライドを選んでください";
          } }, E.mount(el), h("span", {}, `${i + 1}. ${E.strip(slide.title || slide.message || "").slice(0, 18)}`));
        grid.append(cell);
      });
      const ok = h("button", { type: "button", class: "btn btn-primary btn-sm zm-ok", disabled: true, onclick: () => { close(); onPick([...chosen].sort((a, b) => a - b)); } }, "スライドを選んでください");
      return h("div", { class: "zm-picker" }, h("div", { class: "rb-menu-head" }, title), grid, multiple ? h("div", { class: "zm-foot" }, ok) : null);
    };
  }
  const zoomObject = (i, box) => ({ id: E.newObjectId(), kind: "zoom", ...box, target: app.ensureSid(i) });
  /** スライド ズーム: a picture of one slide on this one; clicking it in the show goes there and back. */
  function insertSlideZoom(anchor) {
    openPop(anchor, slidePicker({ title: "スライド ズーム：ズームするスライドを選ぶ", onPick: ([i]) => {
      const [o] = editor.insert([zoomObject(i, center(480, 270))]);
      if (o) app.toast(`スライド ${i + 1} へのズームを入れました（発表でクリックするとそのスライドへ移り、終わると戻ります）`);
    }, filter: (_, i) => i !== app.index() }));
  }
  /** サマリー ズーム: a new slide with a zoom to each chosen slide, laid out in a grid. */
  function insertSummaryZoom(anchor) {
    openPop(anchor, slidePicker({ title: "サマリー ズーム：まとめに入れるスライドを選ぶ", multiple: true, onPick: (picked) => {
      if (!picked.length) return;
      const cols = picked.length <= 2 ? picked.length : picked.length <= 4 ? 2 : 3;
      const rows = Math.ceil(picked.length / cols);
      const gap = 40;
      const areaW = 1640;
      const areaH = 720;
      const w = Math.min((areaW - gap * (cols - 1)) / cols, ((areaH - gap * (rows - 1)) / rows) * (16 / 9));
      const hh = w * (9 / 16);
      const x0 = (E.W - (w * cols + gap * (cols - 1))) / 2;
      const y0 = 250 + (areaH - (hh * rows + gap * (rows - 1))) / 2;
      const elements = picked.map((i, k) => zoomObject(i, { x: Math.round(x0 + (k % cols) * (w + gap)), y: Math.round(y0 + Math.floor(k / cols) * (hh + gap)), w: Math.round(w), h: Math.round(hh) }));
      app.insertSlideWith({ type: "blank", title: "サマリー", elements });
      app.toast(`${picked.length}枚へのサマリー ズームのスライドを作りました`);
    } }));
  }
  /** セクション ズーム: zooms to the first slide of each section. */
  function insertSectionZoom() {
    const starts = slides().map((slide, i) => (slide.section ? i : -1)).filter((i) => i >= 0);
    if (!starts.length) { app.toast("セクションがありません（ホーム →「セクション」で作れます）"); return; }
    const w = 420;
    const hh = 236;
    const cols = Math.min(3, starts.length);
    const gap = 40;
    const x0 = (E.W - (w * cols + gap * (cols - 1))) / 2;
    const made = starts.map((i, k) => zoomObject(i, { x: Math.round(x0 + (k % cols) * (w + gap)), y: 280 + Math.floor(k / cols) * (hh + gap), w, h: hh }));
    editor.commit([...editor.objects(), ...made], { select: made.map((o) => o.id) });
    app.toast(`${starts.length}つのセクションへのズームを入れました`);
  }
  function zoomTab() {
    return [
      group("ズームのオプション",
        btn("reset", "ズームに|戻る", "ズーム先のスライドが終わったら、このスライドに戻る", () => editor.apply((o) => (o.kind === "zoom" ? { back: o.back === false ? undefined : false } : null)), { big: true, pressed: () => zoomSel()?.back !== false }),
        btn("change", "ズーム先の|変更", "ズームするスライドを選び直す", (event) => {
          const o = zoomSel();
          if (!o) return;
          openPop(event.currentTarget.closest("button") || event.currentTarget, slidePicker({ title: "ズーム先のスライド", onPick: ([i]) => editor.apply((x) => (x.id === o.id ? { target: app.ensureSid(i) } : null)) }));
        }, { big: true, enabled: () => Boolean(zoomSel()) }),
        btn("slide", "ズーム先へ|移動", "編集画面でズーム先のスライドを開く", () => { const o = zoomSel(); const i = slides().findIndex((s) => s.sid === o?.target); if (i >= 0) app.select(i); }, { big: true, enabled: () => Boolean(zoomSel()) })),
      group("ズームのスタイル",
        drop("outline", "枠線", "ズームの枠の色", () => menu([[E.themeColors(E.PALETTE.sheet?.themeId).accent1, "アクセント 1"], [E.themeColors(E.PALETTE.sheet?.themeId).tx1, "テキスト 1"], ["#d9d9d9", "薄いグレー"], ["none", "なし"]].map(([c, label]) => ({ label, on: (zoomSel()?.stroke ?? "#b7c3da") === c, run: () => editor.apply((o) => (o.kind === "zoom" ? { stroke: c } : null)) }))), { big: true, enabled: () => Boolean(zoomSel()) })),
    ];
  }

  // ---------------------------------------------------------------- カメオ・スクリーンショット

  function insertCamera() {
    const box = { x: 1480, y: 640, w: 320, h: 320 };
    const [o] = editor.insert([{ id: E.newObjectId(), kind: "camera", ...box, mask: "ellipse" }], { at: [box.x + box.w / 2, box.y + box.h / 2] });
    if (o) app.toast("カメオを入れました。発表を始めるとカメラの映像になります（ブラウザでカメラを許可してください）");
  }
  /** スクリーンショット: a picture of a window, a tab or the screen (chosen in the browser's own dialog). */
  async function insertScreenshot() {
    let stream;
    try { stream = await navigator.mediaDevices.getDisplayMedia({ video: true, audio: false }); } catch (error) { if (error?.name !== "NotAllowedError") app.toast(`画面を取り込めません（${error.message || error.name}）`); return; }
    try {
      const video = document.createElement("video");
      video.srcObject = stream;
      video.muted = true;
      await video.play();
      await new Promise((resolve) => (video.readyState >= 2 ? resolve() : video.addEventListener("loadeddata", resolve, { once: true })));
      await new Promise((resolve) => setTimeout(resolve, 250));
      const canvas = document.createElement("canvas");
      canvas.width = video.videoWidth;
      canvas.height = video.videoHeight;
      canvas.getContext("2d").drawImage(video, 0, 0);
      const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
      if (!blob) throw new Error("画像にできません");
      await app.insertFiles([new File([blob], `スクリーンショット ${new Date().toLocaleString("ja-JP")}.png`, { type: "image/png" })]);
      app.toast("スクリーンショットを入れました");
    } catch (error) {
      app.toast(`スクリーンショットを撮れませんでした（${error.message}）`);
    } finally {
      stream.getTracks().forEach((t) => t.stop());
    }
  }

  // ---------------------------------------------------------------- 数式

  function equationDialog(existing = null) {
    const preview = h("div", { class: "eq-preview" });
    const input = h("textarea", { class: "eq-input", rows: 3, spellcheck: "false", "aria-label": "数式（TeX）", placeholder: "例：\\frac{a}{b}、x^2、\\sqrt{x}、\\sum_{i=1}^{n}" }, existing?.tex || "");
    const draw = () => { preview.replaceChildren(input.value.trim() ? E.texToMathML(input.value) : h("span", { class: "hint" }, "ここに数式が表示されます")); };
    input.addEventListener("input", draw);
    const samples = h("div", { class: "eq-samples" }, EQUATIONS.map(([label, tex]) => h("button", { type: "button", class: "btn btn-sm", title: tex, onclick: () => { input.value = tex; draw(); input.focus(); } }, label)));
    const symbols = h("div", { class: "eq-samples" }, [["分数", "\\frac{}{}"], ["上付き", "^{}"], ["下付き", "_{}"], ["ルート", "\\sqrt{}"], ["Σ", "\\sum_{}^{}"], ["∫", "\\int_{}^{}"], ["×", "\\times "], ["÷", "\\div "], ["±", "\\pm "], ["≤", "\\le "], ["≥", "\\ge "], ["≠", "\\ne "], ["≈", "\\approx "], ["→", "\\to "], ["π", "\\pi "], ["α", "\\alpha "], ["文字", "\\text{}"]]
      .map(([label, snippet]) => h("button", { type: "button", class: "btn btn-xs", title: snippet, onmousedown: (event) => event.preventDefault(), onclick: () => {
        const at = input.selectionStart;
        input.value = input.value.slice(0, at) + snippet + input.value.slice(input.selectionEnd);
        const brace = snippet.indexOf("{}");
        const caret = at + (brace >= 0 ? brace + 1 : snippet.length);
        input.setSelectionRange(caret, caret);
        draw();
        input.focus();
      } }, label)));
    const dialog = h("dialog", { class: "eq-dialog", "aria-label": "数式" },
      h("div", { class: "dialog-head" }, h("h3", {}, existing ? "数式の編集" : "数式の挿入"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body" }, h("p", { class: "hint" }, "TeXの書き方で入力します（\\frac{分子}{分母}・x^2・x_i・\\sqrt{x}・\\sum_{i=1}^{n}・\\text{日本語}）。"), samples, symbols, input, preview),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "やめる"),
        h("button", { type: "button", class: "btn btn-primary eq-ok", onclick: () => {
          const tex = input.value.trim();
          if (!tex) return;
          if (existing) editor.apply((o) => (o.id === existing.id ? { tex } : null));
          else editor.insert([{ id: E.newObjectId(), kind: "equation", ...center(900, 220), tex, fs: 56 }]);
          dialog.close();
        } }, existing ? "更新" : "挿入")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
    draw();
    input.focus();
  }

  // ---------------------------------------------------------------- ワードアート・フィールド

  function insertWordArt(style) {
    const words = "ここに文字を入力";
    const html = style.marker ? `<p><span style="background-color:${style.marker}">${words}</span></p>` : `<p>${words}</p>`;
    const [o] = editor.insert([{ id: E.newObjectId(), kind: "text", ...center(1100, 160), text: html, align: "center", valign: "middle", autofit: "grow", ...style.text }]);
    if (o) { editor.select([o.id]); requestAnimationFrame(() => editor.startTyping(o.id, { caret: "all" })); }
  }
  /** スライド番号・総数・日付 as fields: at the caret while typing, else in a new text box. */
  function insertField(kind) {
    const label = { slideno: "#", total: "#", date: "日付" }[kind];
    const html = `<span data-field="${kind}">${label}</span>`;
    if (editor.typing && editor.restoreRange()) { document.execCommand("insertHTML", false, html); return; }
    editor.insert([{ id: E.newObjectId(), kind: "text", ...center(kind === "date" ? 520 : 260, 90), text: `<p>${html}</p>`, fs: 36, align: "center" }]);
    app.toast(kind === "date" ? "今日の日付を入れました（表示するたびに更新されます）" : "スライド番号を入れました（並べ替えても自動で直ります）");
  }

  // ---------------------------------------------------------------- 段組み・文字種の変換・音声入力

  const COLS = [[1, "1段"], [2, "2段"], [3, "3段"], [4, "4段"]];
  const columnsMenu = () => menu(COLS.map(([n, label]) => ({ label, on: (selected().find((o) => ["shape", "text"].includes(o.kind))?.cols || 1) === n, run: () => editor.apply((o) => (["shape", "text"].includes(o.kind) ? { cols: n > 1 ? n : undefined } : null)) })));
  const CASES = {
    upper: ["すべて大文字にする", (t) => t.toUpperCase()],
    lower: ["すべて小文字にする", (t) => t.toLowerCase()],
    title: ["各単語の先頭文字を大文字にする", (t) => t.replace(/\b([a-z])/g, (m) => m.toUpperCase())],
    full: ["全角にする", (t) => t.replace(/[!-~]/g, (c) => String.fromCharCode(c.charCodeAt(0) + 0xfee0)).replace(/ /g, "　")],
    half: ["半角にする", (t) => t.replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/　/g, " ")],
  };
  /** 文字種の変換: the selected words while typing, else every text node of the selected objects. */
  function changeCase(kind) {
    const fn = CASES[kind][1];
    if (editor.typing && editor.restoreRange()) {
      const sel = window.getSelection();
      const text = sel?.toString() || "";
      if (text) { document.execCommand("insertText", false, fn(text)); return; }
    }
    const targets = selected().filter((o) => ["shape", "text"].includes(o.kind) && o.text);
    if (!targets.length) { app.toast("文字を選ぶか、文字の入った図形を選んでください"); return; }
    editor.apply((o) => {
      if (!targets.some((x) => x.id === o.id)) return null;
      const box = document.createElement("div");
      box.append(E.richNodes(o.text));
      const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
      while (walker.nextNode()) walker.currentNode.nodeValue = fn(walker.currentNode.nodeValue);
      return { text: E.sanitizeRich(box.innerHTML) };
    });
  }
  let dictation = null;
  /** 音声入力: what is said is typed at the caret (or into a new text box). */
  function toggleDictation() {
    if (dictation) { dictation.stop(); return; }
    const Recognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!Recognition) { app.toast("このブラウザは音声入力に対応していません（Chrome・Edgeで使えます）"); return; }
    if (!editor.typing) {
      const target = selected().find((o) => ["shape", "text"].includes(o.kind));
      if (target) editor.startTyping(target.id, { caret: "end" });
      else { const [o] = editor.insert([{ id: E.newObjectId(), kind: "text", ...center(1200, 120), text: "<p><br></p>", fs: 40 }]); if (o) editor.startTyping(o.id); }
    }
    const rec = new Recognition();
    rec.lang = "ja-JP";
    rec.continuous = true;
    rec.interimResults = false;
    rec.onresult = (event) => {
      for (let i = event.resultIndex; i < event.results.length; i += 1) {
        if (!event.results[i].isFinal) continue;
        const text = event.results[i][0].transcript;
        if (editor.typing && editor.restoreRange()) document.execCommand("insertText", false, text);
      }
    };
    rec.onerror = (event) => { if (["not-allowed", "service-not-allowed"].includes(event.error)) app.toast("マイクを使えないため音声入力できません"); };
    rec.onend = () => { if (dictation === rec) { dictation = null; kit.refreshRibbon?.(); app.toast("音声入力を終えました"); } };
    try { rec.start(); } catch { return; }
    dictation = rec;
    kit.refreshRibbon?.();
    app.toast("音声入力：話すと文字になります（もう一度押すと終わります）");
  }

  // ---------------------------------------------------------------- フォントの置換

  /** The fonts the deck's text uses: [{ key, label, count }] (a key is "font:<name>" or "face:<name>"). */
  function fontsInUse() {
    const counts = new Map();
    for (const slide of app.deck()?.slides || []) for (const o of slide.elements || []) {
      if (!["shape", "text"].includes(o.kind) || !o.text) continue;
      const key = o.fontFace ? `face:${o.fontFace}` : `font:${o.font || "body"}`;
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    return [...counts].map(([key, count]) => ({ key, count, label: key.startsWith("face:") ? key.slice(5) : (E.FONTS[key.slice(5)]?.[0] || key.slice(5)) }));
  }
  /** フォントの置換: every text in one font set in another (one of the studio's fonts). */
  function fontReplaceDialog() {
    const used = fontsInUse();
    if (!used.length) { app.toast("文字の入った図形・テキスト ボックスがありません"); return; }
    const from = h("select", { "aria-label": "置換前のフォント" }, used.map((f) => h("option", { value: f.key }, `${f.label}（${f.count}か所）`)));
    const to = h("select", { "aria-label": "置換後のフォント" }, Object.entries(E.FONTS).map(([key, [label]]) => h("option", { value: key }, label)));
    const dialog = h("dialog", { class: "font-dialog", "aria-label": "フォントの置換" },
      h("div", { class: "dialog-head" }, h("h3", {}, "フォントの置換"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body fr-body" }, h("label", {}, "置換前のフォント", from), h("label", {}, "置換後のフォント", to),
        h("p", { class: "hint" }, "資料のすべての図形・テキスト ボックスの文字を置き換えます（⌘Zで戻せます）。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "閉じる"),
        h("button", { type: "button", class: "btn btn-primary fr-ok", onclick: () => {
          const changes = [];
          (app.deck()?.slides || []).forEach((slide, i) => (slide.elements || []).forEach((o) => {
            if (!["shape", "text"].includes(o.kind) || !o.text) return;
            const key = o.fontFace ? `face:${o.fontFace}` : `font:${o.font || "body"}`;
            if (key === from.value) changes.push({ slide: i, id: o.id, patch: { font: to.value === "body" ? undefined : to.value, fontFace: undefined } });
          }));
          if (!changes.length || from.value === `font:${to.value}`) { app.toast("置き換える文字がありません"); return; }
          app.patchObjects(changes);
          dialog.close();
          app.toast(`${changes.length}か所のフォントを「${E.FONTS[to.value][0]}」に置き換えました（⌘Zで戻せます）`);
        } }, "置換")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
  }

  // ---------------------------------------------------------------- menus used by the ribbon

  const zoomMenu = (anchor) => menu([
    { label: "サマリー ズーム…", icon: "zoomSlide", run: () => insertSummaryZoom(anchor) },
    { label: "スライド ズーム…", icon: "zoomSlide", run: () => insertSlideZoom(anchor) },
    { label: "セクション ズーム", icon: "zoomSlide", run: () => insertSectionZoom() },
  ]);
  const wordArtMenu = () => (close) => h("div", { class: "wa-gallery" }, h("div", { class: "rb-menu-head" }, "ワードアート（テーマの色）"),
    h("div", { class: "wa-grid" }, wordArtStyles(E).map((st) => h("button", { type: "button", class: "wa-item", title: st.label, onclick: () => { close(); insertWordArt(st); } },
      h("span", { style: { color: st.text.color, "font-weight": st.text.bold ? "700" : "400", "text-decoration": st.text.underline ? "underline" : "none", "letter-spacing": st.text.ls ? `${st.text.ls}em` : null, background: st.marker ? `linear-gradient(transparent 55%, ${st.marker} 55%)` : null } }, "あア"), h("small", {}, st.label)))));
  const fieldMenu = () => menu([
    { label: "スライド番号", icon: "slide", run: () => insertField("slideno") },
    { label: "スライドの総数", icon: "slide", run: () => insertField("total") },
    { label: "日付（自動更新）", icon: "date", run: () => insertField("date") },
  ]);
  const caseMenu = () => menu(Object.entries(CASES).map(([k, [label]]) => ({ label, run: () => changeCase(k) })));

  return {
    zoomTab, zoomMenu, isZoom: () => Boolean(zoomSel()), insertCamera, insertScreenshot, equationDialog, wordArtMenu, fieldMenu, columnsMenu, caseMenu,
    toggleDictation, get dictating() { return Boolean(dictation); }, escapeText, insertSlideZoom, insertSummaryZoom, fontReplaceDialog, fontsInUse,
  };
}
