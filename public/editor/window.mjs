// The editing window around the slide, laid out as PowerPoint's (index.html #editView): the splitters of the
// thumbnails, the task pane and the notes; closing and reopening the thumbnails and the task pane; the notes under
// the slide; the one-line message bar under the ribbon; and the status bar (where you are, the notes, the zoom).
// The sizes and what is open are remembered in this browser.

import { createRulers } from "./rulers.mjs";
import { createNotesEditor } from "./notesedit.mjs";

const KEY = "hpptx-window";
const LIMITS = { film: [120, 420], side: [300, 760], notes: [60, 520] };

export function createShell(app) {
  const { h, editor } = app;
  const $ = (id) => document.getElementById(id);
  const view = $("editView");
  const prefs = { film: null, side: null, notes: null, filmClosed: false, sideClosed: false, notesShown: false, rulers: false };
  try { Object.assign(prefs, JSON.parse(localStorage.getItem(KEY) || "{}")); } catch { /* a fresh start */ }
  const save = () => { try { localStorage.setItem(KEY, JSON.stringify(prefs)); } catch { /* private window */ } };
  const clamp = (key, value) => Math.round(Math.max(LIMITS[key][0], Math.min(LIMITS[key][1], value)));

  // ---------------------------------------------------------------- the panes and their splitters

  function applyPanes() {
    view.classList.toggle("film-closed", prefs.filmClosed);
    view.classList.toggle("side-closed", prefs.sideClosed);
    // A width someone dragged wins over the window's default (and over the narrow-window rules).
    const width = (name, key, closed) => {
      if (closed) view.style.removeProperty(name);
      else if (prefs[key]) view.style.setProperty(name, `${prefs[key]}px`);
      else view.style.removeProperty(name);
    };
    width("--film-w", "film", prefs.filmClosed);
    width("--side-w", "side", prefs.sideClosed);
    if (prefs.notes) $("stageBody").parentElement.style.setProperty("--notes-h", `${prefs.notes}px`);
    $("filmReopen").hidden = !prefs.filmClosed;
    $("sideClose").title = prefs.sideClosed ? "作業ウィンドウを開く" : "作業ウィンドウを閉じる（タブを押すと開きます）";
    $("sideClose").setAttribute("aria-label", prefs.sideClosed ? "作業ウィンドウを開く" : "作業ウィンドウを閉じる");
    app.ribbonChanged?.();
  }
  /** "film" (the thumbnails) or "side" (the task pane). */
  const paneOpen = (which) => !(which === "film" ? prefs.filmClosed : prefs.sideClosed);
  function setPane(which, open) {
    if (which === "film") prefs.filmClosed = !open; else prefs.sideClosed = !open;
    save();
    applyPanes();
  }
  const togglePane = (which) => setPane(which, !paneOpen(which));

  /** A splitter: drag to resize (pointer capture keeps the drag even over the slide). */
  function drag(el, onMove, onEnd) {
    el.addEventListener("pointerdown", (down) => {
      if (down.button !== 0) return;
      down.preventDefault();
      el.setPointerCapture(down.pointerId);
      el.classList.add("dragging");
      document.body.classList.add("resizing");
      const move = (event) => onMove(event);
      const up = () => {
        el.classList.remove("dragging");
        document.body.classList.remove("resizing");
        el.removeEventListener("pointermove", move);
        el.removeEventListener("pointerup", up);
        el.removeEventListener("pointercancel", up);
        save();
        onEnd?.();
      };
      el.addEventListener("pointermove", move);
      el.addEventListener("pointerup", up);
      el.addEventListener("pointercancel", up);
    });
  }
  const filmSplit = view.querySelector('.pane-split[data-split="film"]');
  const sideSplit = view.querySelector('.pane-split[data-split="side"]');
  const notesSplit = view.querySelector('.notes-split[data-split="notes"]');
  // Dragging the thumbnails narrower than they can show closes them, as in PowerPoint; reopened, they come back
  // as wide as they were before that drag.
  let filmBefore = null;
  filmSplit.addEventListener("pointerdown", () => { filmBefore = prefs.film; }, true);
  drag(filmSplit, (event) => {
    const width = event.clientX - view.getBoundingClientRect().left;
    if (width < 80) { if (!prefs.filmClosed) { prefs.filmClosed = true; prefs.film = filmBefore; applyPanes(); } return; }
    prefs.filmClosed = false;
    prefs.film = clamp("film", width);
    applyPanes();
  });
  filmSplit.addEventListener("dblclick", () => togglePane("film"));
  drag(sideSplit, (event) => {
    if (prefs.sideClosed) return;
    const r = view.getBoundingClientRect();
    prefs.side = clamp("side", Math.min(r.right - event.clientX, r.width * 0.55));
    applyPanes();
  });
  drag(notesSplit, (event) => {
    const r = $("stageBody").parentElement.getBoundingClientRect();
    prefs.notes = clamp("notes", Math.min(r.bottom - event.clientY, r.height * 0.6));
    applyPanes();
  });
  $("filmReopen").addEventListener("click", () => setPane("film", true));
  $("sideClose").addEventListener("click", () => togglePane("side"));
  // A tab of the closed task pane opens it on that tab.
  for (const tab of view.querySelectorAll(".side-tab")) tab.addEventListener("click", () => { if (prefs.sideClosed) setPane("side", true); });

  // ---------------------------------------------------------------- the notes under the slide

  // ノートの書式: the box has a small toolbar (bold, italic, underline, strike, bullets, numbers, links).
  const notesEditor = createNotesEditor(app, { id: "notesInput" });
  $("notesPane").append(notesEditor.el);
  const notesShown = () => prefs.notesShown;
  function renderNotes() {
    const slide = app.slide();
    const shown = prefs.notesShown && app.view() === "single" && Boolean(slide);
    $("notesPane").hidden = !shown;
    notesSplit.hidden = !shown;
    if (shown) notesEditor.show(slide);
    const toggle = $("notesToggle");
    toggle.setAttribute("aria-pressed", String(prefs.notesShown));
    toggle.classList.toggle("has-notes", Boolean(String(slide?.notes ?? "").trim()));
    toggle.title = prefs.notesShown ? "ノートを隠す" : `ノート：スライドの下にスピーカーノートを表示する${slide?.notes ? "（このスライドにはノートがあります）" : ""}`;
  }
  function toggleNotes(show = !prefs.notesShown) {
    prefs.notesShown = show;
    save();
    renderNotes();
    if (show && app.view() === "single") requestAnimationFrame(() => notesEditor.focus());
    app.ribbonChanged?.();
  }
  $("notesToggle").addEventListener("click", () => toggleNotes());

  // ---------------------------------------------------------------- the message bar

  /** One line under the ribbon (PowerPoint's message bar): the whole text on 「詳しく」. */
  function renderMessage(content) {
    const bar = $("msgBar");
    if (!content) { bar.hidden = true; bar.classList.remove("open"); bar.replaceChildren(); return; }
    const text = content.querySelector(".text");
    if (text) {
      text.title = text.textContent;
      text.after(h("button", { type: "button", class: "msg-more", onclick: () => { bar.classList.toggle("open"); } }, "詳しく"));
    }
    bar.hidden = false;
    bar.replaceChildren(content);
  }

  // ---------------------------------------------------------------- the status bar

  const slider = $("zoomSlider");
  function renderZoom() {
    rulers?.refresh();
    const single = app.view() === "single";
    const pct = editor.zoomPercent();
    if (document.activeElement !== slider) slider.value = String(pct);
    $("zoomPctBtn").textContent = `${pct}%`;
    $("zoomFitBtn").setAttribute("aria-pressed", String(editor.zoom === "fit"));
    for (const id of ["zoomOutBtn", "zoomInBtn", "zoomPctBtn", "zoomFitBtn"]) $(id).disabled = !single;
    slider.disabled = !single;
  }
  function renderStatus() {
    const info = app.slideStatus();
    const pos = $("slidePos");
    if (!info) { pos.replaceChildren(); return; }
    pos.replaceChildren(info.pos, info.detail ? h("span", { class: "sb-detail" }, `　${info.detail}`) : "");
    pos.title = [info.pos, info.detail].filter(Boolean).join("　");
    renderNotes();
    renderZoom();
  }
  slider.addEventListener("input", () => editor.setZoom(Number(slider.value)));
  $("zoomOutBtn").addEventListener("click", () => editor.zoomStep(-1));
  $("zoomInBtn").addEventListener("click", () => editor.zoomStep(1));
  $("zoomFitBtn").addEventListener("click", () => editor.setZoom("fit"));
  $("zoomPctBtn").addEventListener("click", (event) => app.openZoomMenu(event.currentTarget));
  // ⌘＋ホイール (and a trackpad pinch) zooms, as in PowerPoint; small steps add up to one zoom step.
  let wheel = 0;
  $("stageBody").addEventListener("wheel", (event) => {
    if (!(event.ctrlKey || event.metaKey) || app.view() !== "single") return;
    event.preventDefault();
    wheel += event.deltaY;
    if (Math.abs(wheel) < 30) return;
    editor.zoomStep(wheel < 0 ? 1 : -1);
    wheel = 0;
  }, { passive: false });

  /** ⌘0 fits, ⌘＋ / ⌘− zoom (only on the 1枚 view, outside text fields). Returns true when the key was taken. */
  function zoomKey(event) {
    const meta = event.metaKey || event.ctrlKey;
    if (!meta || event.altKey || app.view() !== "single") return false;
    if (event.key === "0") { editor.setZoom("fit"); return true; }
    if (event.key === "=" || event.key === "+" || event.key === ";") { editor.zoomStep(1); return true; }
    if (event.key === "-") { editor.zoomStep(-1); return true; }
    return false;
  }

  // ---------------------------------------------------------------- the rulers (表示 → ルーラー)

  const rulers = createRulers(editor, { E: app.E, stage: $("stageBody").parentElement, body: $("stageBody") });
  const rulersShown = () => prefs.rulers;
  function toggleRulers(show = !prefs.rulers) {
    prefs.rulers = show;
    save();
    rulers.show(show && app.view() === "single");
    app.ribbonChanged?.();
  }

  applyPanes();
  rulers.show(prefs.rulers);
  return {
    renderStatus, renderZoom, renderNotes, renderMessage, toggleNotes, notesShown, paneOpen, setPane, togglePane, zoomKey, toggleRulers, rulersShown,
    viewChanged: () => rulers.show(prefs.rulers && app.view() === "single"),
  };
}
