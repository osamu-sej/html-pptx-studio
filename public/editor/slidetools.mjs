// More ways to add slides, as PowerPoint's 新しいスライド and 挿入: スライドの再利用 (slides from a saved deck, a
// PowerPoint file or a JSON file, keeping their own look or taking this deck's theme), アウトラインからスライド (a
// heading per slide and its bullet points, from typed or pasted text, Markdown or a .txt file) and フォト アルバム
// (pictures laid one, two or four to a slide, with their names as captions if wanted).

/** Indented text or Markdown → slides: a line at the left edge (or "# ") is a title, the lines under it its points. */
export function outlineToSlides(text) {
  const slides = [];
  for (const raw of String(text || "").replace(/\r/g, "").split("\n")) {
    if (!raw.trim()) continue;
    const indent = /^(\t| {2,}|　)/.test(raw);
    const line = raw.trim();
    const bullet = /^([-*・•●○◦▪]|\d+[.)．、])\s*/.test(line);
    const heading = /^#{1,6}\s+/.test(line);
    const words = line.replace(/^#{1,6}\s+/, "").replace(/^([-*・•●○◦▪]|\d+[.)．、])\s*/, "").trim();
    if (!words) continue;
    if (heading || (!indent && !bullet) || !slides.length) slides.push({ title: words.slice(0, 90), points: [] });
    else slides.at(-1).points.push(words.slice(0, 200));
  }
  return slides.map((s) => (s.points.length ? { type: "content", title: s.title, points: s.points.slice(0, 8) } : { type: "section", title: s.title }));
}

/** Where pictures go on album slides: [{ x, y, w, h }] boxes inside the content area, for 1, 2 or 4 a slide. */
export function albumBoxes(per, { captions = false } = {}) {
  const area = { x: 120, y: 230, w: 1680, h: 720 };
  const capH = captions ? 56 : 0;
  const gap = 40;
  const cols = per === 1 ? 1 : 2;
  const rows = per === 4 ? 2 : 1;
  const w = (area.w - gap * (cols - 1)) / cols;
  const hh = (area.h - gap * (rows - 1)) / rows - capH;
  return Array.from({ length: per }, (_, k) => ({ x: Math.round(area.x + (k % cols) * (w + gap)), y: Math.round(area.y + Math.floor(k / cols) * (hh + capH + gap)), w: Math.round(w), h: Math.round(hh) }));
}
/** A picture of natural size `nw` × `nh` fitted (centred) in a box. */
export function fitIn(box, nw, nh) {
  const k = Math.min(box.w / Math.max(1, nw), box.h / Math.max(1, nh));
  const w = Math.round(nw * k);
  const hh = Math.round(nh * k);
  return { x: Math.round(box.x + (box.w - w) / 2), y: Math.round(box.y + (box.h - hh) / 2), w, h: hh };
}

export function createSlideTools(app) {
  const { h, E } = app;
  function dialog(title, body, foot, cls) {
    const el = h("dialog", { class: ["st-dialog", cls], "aria-label": title },
      h("div", { class: "dialog-head" }, h("h3", {}, title), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => el.close() }, "✕")),
      h("div", { class: "dialog-body" }, body), h("div", { class: "dialog-foot" }, foot(() => el.close())));
    document.body.append(el);
    el.addEventListener("close", () => el.remove());
    el.showModal();
    return el;
  }

  // ---------------------------------------------------------------- スライドの再利用

  async function openReuse() {
    const list = h("div", { class: "st-sources" });
    const grid = h("div", { class: "st-grid" }, h("p", { class: "hint" }, "左から資料を選ぶと、スライドが並びます。"));
    const keep = h("input", { type: "checkbox", checked: true });
    const count = h("span", { class: "hint st-count" });
    let source = null;
    const picked = new Set();
    const showDeck = (deck, label) => {
      source = deck;
      picked.clear();
      grid.replaceChildren(h("div", { class: "rb-menu-head" }, `${label}（${deck.slides.length}枚）`), ...deck.slides.map((slide, i) => {
        const el = E.render(slide, { ...app.renderOptions(), deck, index: i, mode: "thumb" });
        const cell = h("button", { type: "button", class: "zm-cell", "data-index": String(i), title: `${i + 1}. ${E.strip(slide.title || "")}`, onclick: () => { if (picked.has(i)) picked.delete(i); else picked.add(i); cell.classList.toggle("on", picked.has(i)); sync(); } }, E.mount(el), h("span", {}, `${i + 1}. ${E.strip(slide.title || slide.message || "").slice(0, 18)}`));
        return cell;
      }));
      sync();
    };
    const sync = () => { count.textContent = source ? (picked.size ? `${picked.size}枚を選んでいます` : "クリックで選ぶ（選ばずに「すべて挿入」も）") : ""; };
    const records = await app.savedDecks().catch(() => []);
    list.append(h("div", { class: "rb-menu-head" }, "保存庫の資料"),
      ...(records.length ? records.slice(0, 60).map((r) => h("button", { type: "button", class: "st-source", onclick: async (e) => {
        list.querySelectorAll(".st-source").forEach((b) => b.classList.toggle("on", b === e.currentTarget));
        const rec = await app.savedDeck(r.id);
        if (rec?.deck) showDeck(rec.deck, r.title || "資料");
      } }, h("b", {}, r.title || "無題"), h("small", {}, new Date(r.updatedAt || r.createdAt || Date.now()).toLocaleDateString("ja-JP")))) : [h("p", { class: "hint" }, "保存した資料はありません。")]),
      h("div", { class: "rb-menu-head" }, "ファイルから"),
      h("button", { type: "button", class: "btn btn-sm st-file", onclick: async () => {
        const [file] = await app.pickFiles(".pptx,.json,application/json", false);
        if (!file) return;
        grid.replaceChildren(h("p", { class: "hint" }, `${file.name} を読み込んでいます…`));
        try { showDeck(await app.deckFromFile(file), file.name); } catch (error) { grid.replaceChildren(h("p", { class: "hint" }, `読み込めませんでした：${error.message}`)); }
      } }, "PowerPoint・JSON を開く…"));
    const insert = (all) => {
      if (!source) { app.toast("資料を選んでください"); return false; }
      const indices = all ? source.slides.map((_, i) => i) : [...picked].sort((a, b) => a - b);
      if (!indices.length) { app.toast("スライドを選んでください（またはすべて挿入）"); return false; }
      const slides = indices.map((i) => {
        const slide = JSON.parse(JSON.stringify(source.slides[i]));
        delete slide.sid;
        delete slide.drillOf;
        delete slide.comments;
        if (!keep.checked && slide.master === "source") { delete slide.master; delete slide.sourceViewport; }
        return slide;
      });
      app.insertSlides(slides);
      app.toast(`${slides.length}枚のスライドを再利用しました（⌘Zで戻せます）`);
      return true;
    };
    dialog("スライドの再利用", h("div", { class: "st-reuse" }, list, h("div", { class: "st-main" }, grid, h("label", { class: "sh-choice" }, keep, h("span", {}, "元の書式を保持する（PowerPointから取り込んだスライドは元のマスターのまま。外すとSEJのマスター）")))),
      (close) => [count, h("span", { style: { flex: 1 } }), h("button", { type: "button", class: "btn btn-ghost", onclick: close }, "キャンセル"),
        h("button", { type: "button", class: "btn st-all", onclick: () => { if (insert(true)) close(); } }, "すべて挿入"),
        h("button", { type: "button", class: "btn btn-primary st-ok", onclick: () => { if (insert(false)) close(); } }, "選んだスライドを挿入")], "st-reuse-dialog");
  }

  // ---------------------------------------------------------------- アウトラインからスライド

  function openOutline() {
    const text = h("textarea", { class: "st-outline", rows: 12, placeholder: "例：\n売上の現状\n  前年比112%\n  東日本が伸びている\n来期の打ち手\n  - 品揃えの見直し\n  - 発注の自動化\n\n（左端の行が見出し、字下げ・「-」「・」の行が箇条書き。# 見出し の Markdown も使えます）" });
    const preview = h("p", { class: "hint st-outline-count" });
    const update = () => { const slides = outlineToSlides(text.value); preview.textContent = slides.length ? `${slides.length}枚のスライドになります（箇条書きのある見出しは本文のスライド、ないものは章扉）` : ""; };
    text.addEventListener("input", update);
    dialog("アウトラインからスライド", h("div", { class: "st-outline-box" }, text, h("div", { class: "st-row" }, h("button", { type: "button", class: "btn btn-sm", onclick: async () => {
      const [file] = await app.pickFiles(".txt,.md,text/plain,text/markdown", false);
      if (!file) return;
      text.value = await file.text();
      update();
    } }, "テキスト・Markdown ファイルを開く…"), preview)),
    (close) => [h("button", { type: "button", class: "btn btn-ghost", onclick: close }, "キャンセル"), h("button", { type: "button", class: "btn btn-primary st-ok", onclick: () => {
      const slides = outlineToSlides(text.value);
      if (!slides.length) { app.toast("見出しを1行以上書いてください"); return; }
      app.insertSlides(slides);
      app.toast(`アウトラインから${slides.length}枚のスライドを作りました（⌘Zで戻せます）`);
      close();
    } }, "スライドを作る")], "st-outline-dialog");
    setTimeout(() => text.focus(), 0);
  }

  // ---------------------------------------------------------------- フォト アルバム

  async function openPhotoAlbum() {
    const files = await app.pickFiles("image/*", true);
    if (!files.length) return;
    const per = h("select", { "aria-label": "1枚のスライドの写真の数" }, [[1, "スライドに合わせる（1枚）"], [2, "2枚"], [4, "4枚"]].map(([v, l]) => h("option", { value: String(v) }, l)));
    const captions = h("input", { type: "checkbox" });
    const titles = h("input", { type: "checkbox", checked: true });
    const name = h("input", { type: "text", value: "フォト アルバム", maxlength: "60", class: "sh-name" });
    dialog("フォト アルバム", h("div", { class: "st-album" },
      h("p", {}, `${files.length}枚の写真を、新しいスライドに並べます。`),
      h("label", { class: "sh-inline" }, "写真のレイアウト ", per),
      h("label", { class: "sh-choice" }, captions, h("span", {}, "すべての写真の下にキャプション（ファイル名）を付ける")),
      h("label", { class: "sh-choice" }, titles, h("span", {}, "スライドにタイトルを付ける"), name)),
    (close) => [h("button", { type: "button", class: "btn btn-ghost", onclick: close }, "キャンセル"), h("button", { type: "button", class: "btn btn-primary st-ok", onclick: async () => {
      close();
      const n = Number(per.value);
      const boxes = albumBoxes(n, { captions: captions.checked });
      const pics = [];
      for (const file of files) { try { pics.push({ file, ...(await app.storePicture(file)) }); } catch { /* not a picture */ } }
      if (!pics.length) { app.toast("写真を読み込めませんでした"); return; }
      const slides = [];
      for (let i = 0; i < pics.length; i += n) {
        const elements = [];
        pics.slice(i, i + n).forEach((p, k) => {
          const box = fitIn(boxes[k], p.w, p.h);
          const label = p.file.name.replace(/\.[a-z0-9]+$/i, "").slice(0, 60);
          elements.push({ id: E.newObjectId(), kind: "image", ...box, src: p.src, fileName: p.file.name, alt: label });
          if (captions.checked) elements.push({ id: E.newObjectId(), kind: "text", x: boxes[k].x, y: boxes[k].y + boxes[k].h + 6, w: boxes[k].w, h: 48, text: `<p>${label.replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[c])}</p>`, fs: 26, align: "center", color: "#808080" });
        });
        slides.push({ type: "blank", title: titles.checked ? `${name.value.trim() || "フォト アルバム"}${pics.length > n ? `（${i / n + 1}）` : ""}` : "", elements });
      }
      app.insertSlides(slides);
      app.toast(`${pics.length}枚の写真で${slides.length}枚のスライドを作りました（⌘Zで戻せます）`);
    } }, "作成")], "st-album-dialog");
  }

  return { openReuse, openOutline, openPhotoAlbum };
}
