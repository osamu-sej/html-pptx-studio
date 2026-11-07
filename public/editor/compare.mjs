// 校閲 → 比較 (PowerPoint's Compare): this deck side by side with another version of it — a saved version (版の履歴),
// a deck in the library or a file. Slides are matched by their id (else by title); what the other version adds,
// removes or changes is listed slide by slide, and each change (or all of them) can be taken over.

const VOLATILE = new Set(["comments", "sid", "readingOrder"]);
const fingerprint = (slide) => JSON.stringify(Object.keys(slide || {}).sort().filter((k) => !VOLATILE.has(k)).map((k) => [k, slide[k]]));
const strip = (t) => String(t ?? "").replace(/<[^>]*>/g, "").replace(/\*\*/g, "").trim();

/** What differs between two versions of a slide, in words. */
export function slideDiff(a, b) {
  const out = [];
  if (strip(a.title) !== strip(b.title)) out.push("タイトル");
  if (a.type !== b.type) out.push("レイアウト");
  const text = (s) => JSON.stringify(Object.entries(s).filter(([k]) => !["title", "type", "elements", "timeline", "notes", "comments", "sid", "readingOrder"].includes(k)));
  if (text(a) !== text(b)) out.push("本文");
  const ea = a.elements || [];
  const eb = b.elements || [];
  if (JSON.stringify(ea) !== JSON.stringify(eb)) out.push(ea.length !== eb.length ? `部品（${ea.length}→${eb.length}個）` : "部品");
  if (JSON.stringify(a.timeline || []) !== JSON.stringify(b.timeline || [])) out.push("アニメーション");
  if (strip(a.notes) !== strip(b.notes)) out.push("ノート");
  return out;
}

/**
 * The changes from `mine` to `theirs` (both decks): [{ kind: "changed" | "added" | "removed", mine?, theirs?, what? }],
 * where mine / theirs are slide indexes.
 */
export function compareDecks(mine, theirs) {
  const a = mine?.slides || [];
  const b = theirs?.slides || [];
  const used = new Set();
  const pairOf = new Map();
  const pair = (j, i) => { used.add(i); pairOf.set(j, i); };
  const free = (test) => b.forEach((slide, j) => {
    if (pairOf.has(j)) return;
    const i = a.findIndex((s, k) => !used.has(k) && test(s, slide, k, j));
    if (i >= 0) pair(j, i);
  });
  // The same slide by its id; then the very same content; then the same title (slides from elsewhere have no ids).
  free((s, slide) => slide.sid && s.sid === slide.sid);
  free((s, slide) => (!s.sid || !slide.sid) && fingerprint(s) === fingerprint(slide));
  free((s, slide) => (!s.sid || !slide.sid) && strip(slide.title) && strip(s.title) === strip(slide.title));
  // Last, a slide left over on each side in the same place between matched ones, with the same layout and body
  // (its title or notes rewritten): one slide changed.
  b.forEach((slide, j) => {
    if (pairOf.has(j)) return;
    const prevMine = Math.max(-1, ...[...pairOf.entries()].filter(([jj]) => jj < j).map(([, i]) => i));
    const nextMine = Math.min(a.length, ...[...pairOf.entries()].filter(([jj]) => jj > j).map(([, i]) => i));
    const body = (x) => JSON.stringify(Object.entries(x).filter(([k]) => !["title", "sid", "comments", "readingOrder", "notes"].includes(k)).sort());
    const i = a.findIndex((s, k) => !used.has(k) && k > prevMine && k < nextMine && s.type === slide.type && (!s.sid || !slide.sid) && body(s) === body(slide));
    if (i >= 0) pair(j, i);
  });
  const out = [];
  b.forEach((slide, j) => {
    const i = pairOf.get(j);
    if (i == null) out.push({ kind: "added", theirs: j });
    else if (fingerprint(a[i]) !== fingerprint(slide)) out.push({ kind: "changed", mine: i, theirs: j, what: slideDiff(a[i], slide) });
  });
  a.forEach((_, i) => { if (!used.has(i)) out.push({ kind: "removed", mine: i }); });
  const at = (c) => (c.mine ?? (() => { const prev = [...pairOf.entries()].filter(([j]) => j < c.theirs).map(([, i]) => i); return prev.length ? Math.max(...prev) + 0.5 : -0.5; })());
  return out.sort((x, y) => at(x) - at(y));
}

/** The slides after taking over some changes (each a change from compareDecks). */
export function accept(mine, theirs, changes) {
  const slides = mine.slides.map((s) => ({ slide: s, keep: true }));
  const inserts = [];
  for (const c of changes) {
    if (c.kind === "changed") slides[c.mine] = { slide: { ...JSON.parse(JSON.stringify(theirs.slides[c.theirs])), sid: mine.slides[c.mine].sid || theirs.slides[c.theirs].sid, comments: mine.slides[c.mine].comments }, keep: true };
    else if (c.kind === "removed") slides[c.mine].keep = false;
    else if (c.kind === "added") {
      // After the slide that comes before it in their version (matched in ours), else at the start.
      let anchor = -1;
      for (let j = c.theirs - 1; j >= 0; j -= 1) {
        const sid = theirs.slides[j].sid;
        const i = sid ? mine.slides.findIndex((s) => s.sid === sid) : mine.slides.findIndex((s) => strip(s.title) && strip(s.title) === strip(theirs.slides[j].title));
        if (i >= 0) { anchor = i; break; }
      }
      inserts.push({ after: anchor, order: c.theirs, slide: JSON.parse(JSON.stringify(theirs.slides[c.theirs])) });
    }
  }
  const out = [];
  inserts.filter((x) => x.after < 0).sort((p, q) => p.order - q.order).forEach((x) => out.push(x.slide));
  slides.forEach((item, i) => {
    if (item.keep) out.push(item.slide);
    inserts.filter((x) => x.after === i).sort((p, q) => p.order - q.order).forEach((x) => out.push(x.slide));
  });
  return out;
}

const LABEL = { changed: "変更", added: "追加", removed: "削除" };

export function createCompare(app) {
  const { h, E } = app;
  async function open() {
    const mine = app.deck();
    if (!mine) return;
    const sources = h("div", { class: "st-sources" });
    const list = h("div", { class: "cmp-list" }, h("p", { class: "hint" }, "左から比べる版を選んでください。"));
    let theirs = null;
    let changes = [];
    const thumb = (deck, i) => {
      const el = E.render(deck.slides[i], { ...app.renderOptions(), deck, index: i, mode: "thumb" });
      return h("div", { class: "cmp-thumb" }, E.mount(el));
    };
    function render(label) {
      const now = app.deck();
      changes = compareDecks(now, theirs);
      if (!changes.length) { list.replaceChildren(h("p", { class: "proof-ok" }, `「${label}」とこの資料に違いはありません。`)); return; }
      list.replaceChildren(h("div", { class: "cmp-head" }, h("b", {}, `「${label}」との違い：${changes.length}か所`),
        h("button", { type: "button", class: "btn btn-sm cmp-all", onclick: () => { app.replaceSlides(accept(app.deck(), theirs, changes)); app.toast(`${changes.length}か所の変更をすべて採用しました（⌘Zで戻せます）`); render(label); } }, "すべて採用")),
      ...changes.map((c) => h("div", { class: ["cmp-item", c.kind], "data-kind": c.kind },
        h("div", { class: "cmp-tag" }, LABEL[c.kind]),
        h("div", { class: "cmp-pair" },
          c.mine != null ? h("div", {}, h("small", {}, `この資料 ${c.mine + 1}枚目`), thumb(now, c.mine)) : h("div", { class: "cmp-none" }, "（この資料にはない）"),
          c.theirs != null ? h("div", {}, h("small", {}, `比べる版 ${c.theirs + 1}枚目`), thumb(theirs, c.theirs)) : h("div", { class: "cmp-none" }, "（比べる版にはない）")),
        h("div", { class: "cmp-what" }, c.kind === "changed" ? `違い：${c.what.join("・") || "細かな設定"}` : c.kind === "added" ? "比べる版で増えたスライド" : "比べる版でなくなったスライド"),
        h("button", { type: "button", class: "btn btn-sm btn-primary cmp-take", onclick: () => { app.replaceSlides(accept(app.deck(), theirs, [c])); app.toast(`${LABEL[c.kind]}を採用しました（⌘Zで戻せます）`); render(label); } }, c.kind === "removed" ? "削除を採用" : c.kind === "added" ? "追加を採用" : "変更を採用"))));
    }
    const choose = (deck, label, btn) => {
      sources.querySelectorAll(".st-source").forEach((b) => b.classList.toggle("on", b === btn));
      theirs = deck;
      render(label);
    };
    const versions = await app.versions().catch(() => []);
    const records = await app.savedDecks().catch(() => []);
    sources.append(h("div", { class: "rb-menu-head" }, "版の履歴"),
      ...(versions.length ? versions.slice(0, 30).map((v) => { const b = h("button", { type: "button", class: "st-source", onclick: () => choose(JSON.parse(v.data), v.label, b) }, h("b", {}, v.label), h("small", {}, `${new Date(v.at).toLocaleString("ja-JP", { month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" })}・${v.count}枚`)); return b; }) : [h("p", { class: "hint" }, "版はまだありません。")]),
      h("div", { class: "rb-menu-head" }, "保存庫の資料"),
      ...records.slice(0, 30).map((r) => { const b = h("button", { type: "button", class: "st-source", onclick: async () => { const rec = await app.savedDeck(r.id); if (rec?.deck) choose(rec.deck, r.title || "資料", b); } }, h("b", {}, r.title || "無題"), h("small", {}, new Date(r.updatedAt || Date.now()).toLocaleDateString("ja-JP"))); return b; }),
      h("div", { class: "rb-menu-head" }, "ファイルから"),
      h("button", { type: "button", class: "btn btn-sm", onclick: async () => {
        const [file] = await app.pickFiles(".pptx,.json,application/json", false);
        if (!file) return;
        list.replaceChildren(h("p", { class: "hint" }, `${file.name} を読み込んでいます…`));
        try { choose(await app.deckFromFile(file), file.name, null); } catch (error) { list.replaceChildren(h("p", { class: "hint" }, `読み込めませんでした：${error.message}`)); }
      } }, "PowerPoint・JSON を開く…"));
    const dialog = h("dialog", { class: "compare-dialog", "aria-label": "比較" },
      h("div", { class: "dialog-head" }, h("h3", {}, "比較"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body st-reuse" }, sources, list),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn", onclick: () => dialog.close() }, "閉じる")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
  }
  return { open };
}
