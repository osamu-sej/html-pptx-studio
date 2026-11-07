// 校閲 → 音声読み上げ (PowerPoint's Read Aloud): the slide's words read out by the browser's voice, a block at a time
// and lit on the slide as it is read, with PowerPoint's small bar (前へ・再生/一時停止・次へ・速度・声・閉じる).
// It reads the words picked while typing, else the objects selected, else the whole slide in reading order (decoration
// and what is hidden from screen readers are not read).

const BLOCK = /^(block|flex|grid|list-item|table|table-cell|table-row|flow-root|-webkit-box)$/;
const BLOCK_TAGS = new Set(["P", "DIV", "LI", "UL", "OL", "H1", "H2", "H3", "H4", "H5", "H6", "TD", "TH", "TR", "TABLE", "SECTION", "ARTICLE", "FIGCAPTION", "BLOCKQUOTE", "DT", "DD"]);
/** A block holds its words together (the page's own layout says; the tag where there is none). */
function isBlock(el, view) {
  const display = view?.getComputedStyle?.(el)?.display;
  return display ? BLOCK.test(display) : BLOCK_TAGS.has(el.tagName);
}

/** The blocks of words in `root`, in order: [{ text, el }] (an element's words once, the master's marks left out). */
export function readingChunks(root, { skip = ".hs-decor, [aria-hidden='true'], .ed-overlay, .ed-handle" } = {}) {
  const chunks = [];
  const byBlock = new Map();
  const walker = root.ownerDocument.createTreeWalker(root, 4 /* NodeFilter.SHOW_TEXT */);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.textContent.replace(/\s+/g, " ").trim();
    const parent = node.parentElement;
    if (!text || !parent || parent.closest(skip)) continue;
    const view = root.ownerDocument.defaultView;
    if (parent.closest("[hidden]") || view?.getComputedStyle?.(parent)?.visibility === "hidden") continue;
    // The nearest block holds the words (a 太字 run stays in its sentence).
    let block = parent;
    while (block !== root && block.parentElement && !isBlock(block, view)) block = block.parentElement;
    let chunk = byBlock.get(block);
    if (!chunk) { chunk = { text: "", el: block }; byBlock.set(block, chunk); chunks.push(chunk); }
    chunk.text = `${chunk.text}${chunk.text && !/[　-ヿ一-鿿]$/.test(chunk.text) ? " " : ""}${text}`;
  }
  return chunks.filter((c) => c.text);
}

export function createReadAloud(app) {
  const { h } = app;
  const synth = () => window.speechSynthesis;
  let chunks = [];
  let at = 0;
  let playing = false;
  let token = 0;
  let bar = null;
  let slideAt = -1; // the slide being read: moving to another one stops the reading
  let rate = (() => { try { return Number(localStorage.getItem("hpptx-read-rate")) || 1; } catch { return 1; } })();
  let voiceName = (() => { try { return localStorage.getItem("hpptx-read-voice") || ""; } catch { return ""; } })();

  const voices = () => (synth()?.getVoices?.() || []).filter((v) => /^ja/i.test(v.lang));
  function light(chunk) {
    for (const old of document.querySelectorAll(".ra-reading")) old.classList.remove("ra-reading");
    if (!chunk) return;
    // The stage draws its slide again after an edit: the same words are found on the new drawing.
    if (!chunk.el?.isConnected) {
      const slide = document.querySelector("#stageBody .slide-wrap .hs-slide");
      const twin = slide && readingChunks(slide).find((c) => c.text === chunk.text);
      if (twin) chunk.el = twin.el;
    }
    if (chunk.el?.isConnected) chunk.el.classList.add("ra-reading");
  }
  function speak(k) {
    const my = ++token;
    synth()?.cancel();
    if (app.index?.() !== slideAt) { stop(); return; }
    if (k >= chunks.length) { stop(); app.toast("読み上げが終わりました"); return; }
    at = Math.max(0, k);
    playing = true;
    light(chunks[at]);
    paint();
    const u = new SpeechSynthesisUtterance(chunks[at].text);
    u.lang = "ja-JP";
    u.rate = rate;
    const voice = voices().find((v) => v.name === voiceName);
    if (voice) u.voice = voice;
    u.onend = () => { if (my === token && playing) speak(at + 1); };
    u.onerror = (event) => { if (my === token && event.error !== "interrupted" && event.error !== "canceled") { stop(); app.toast("読み上げられませんでした（このブラウザの音声）"); } };
    synth().speak(u);
  }
  function pause() { token += 1; playing = false; synth()?.cancel(); paint(); }
  function stop() {
    token += 1;
    playing = false;
    synth()?.cancel();
    light(null);
    bar?.remove();
    bar = null;
    app.refreshRibbon?.();
  }

  /** What to read: the words picked while typing, else the selected objects, else the slide. */
  function gather() {
    const slide = document.querySelector("#stageBody .slide-wrap .hs-slide");
    if (!slide) return [];
    const sel = window.getSelection();
    const picked = sel && !sel.isCollapsed && slide.contains(sel.anchorNode) ? sel.toString().replace(/\s+/g, " ").trim() : "";
    if (picked) return [{ text: picked, el: sel.anchorNode.parentElement?.closest(".hs-obj") || null }];
    const ids = app.selectedIds?.() || [];
    if (ids.length) {
      const own = ids.flatMap((id) => readingChunks(slide).filter((c) => c.el.closest(`[data-el="${CSS.escape(id)}"]`)));
      if (own.length) return own;
    }
    return readingChunks(slide);
  }

  function paint() {
    if (!bar) return;
    bar.querySelector(".ra-play").textContent = playing ? "⏸" : "▶";
    bar.querySelector(".ra-play").title = playing ? "一時停止" : "再生";
    bar.querySelector(".ra-where").textContent = chunks.length ? `${Math.min(at + 1, chunks.length)} / ${chunks.length}` : "";
  }
  function openBar() {
    const speed = h("select", { class: "ra-rate", "aria-label": "読み上げの速度", onchange: (event) => { rate = Number(event.target.value); try { localStorage.setItem("hpptx-read-rate", String(rate)); } catch { /* private */ } if (playing) speak(at); } },
      [0.75, 1, 1.25, 1.5, 2].map((r) => h("option", { value: String(r), selected: r === rate || null }, `${r}×`)));
    const voice = h("select", { class: "ra-voice", "aria-label": "声" }, h("option", { value: "" }, "標準の声"));
    const fillVoices = () => voice.replaceChildren(h("option", { value: "" }, "標準の声"), ...voices().map((v) => h("option", { value: v.name, selected: v.name === voiceName || null }, v.name)));
    fillVoices();
    synth()?.addEventListener?.("voiceschanged", fillVoices, { once: true });
    voice.addEventListener("change", () => { voiceName = voice.value; try { localStorage.setItem("hpptx-read-voice", voiceName); } catch { /* private */ } if (playing) speak(at); });
    const b = (cls, label, title, fn) => h("button", { type: "button", class: `ra-btn ${cls}`, title, "aria-label": title, onclick: fn }, label);
    bar = h("div", { class: "ra-bar", role: "toolbar", "aria-label": "音声読み上げ" },
      b("ra-prev", "⏮", "前へ", () => speak(Math.max(0, at - 1))),
      b("ra-play", "⏸", "一時停止", () => (playing ? pause() : speak(at))),
      b("ra-next", "⏭", "次へ", () => speak(at + 1)),
      h("span", { class: "ra-where" }), speed, voice,
      b("ra-close", "✕", "閉じる（読み上げを止める）", () => stop()));
    // On the stage (its slide is drawn again and again; the bar stays).
    (document.querySelector(".stage") || document.body).append(bar);
  }

  function toggle() {
    if (bar) { stop(); return; }
    if (!window.speechSynthesis || typeof window.SpeechSynthesisUtterance !== "function") { app.toast("このブラウザは音声読み上げに対応していません"); return; }
    chunks = gather();
    slideAt = app.index?.() ?? -1;
    if (!chunks.length) { app.toast("読み上げる文字がありません"); return; }
    openBar();
    speak(0);
    app.refreshRibbon?.();
  }

  return { toggle, stop, get active() { return Boolean(bar); }, get chunks() { return chunks; }, get at() { return at; } };
}
