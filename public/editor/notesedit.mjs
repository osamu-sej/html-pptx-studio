// ノートの書式: the speaker notes as a small rich-text box — bold, italic, underline, strike, bullets, numbers and links — with
// a toolbar of its own. The notes under the slide and the 表示 → ノート表示 dialog share it. What is typed is kept on the slide
// two ways: `notes` (the plain words, so everything that reads the notes keeps working) and `notesRich` (the formatted words,
// kept only while something is formatted). objects.js (`noteFields`, `noteOf`) decides both; this file is only the box.

const TOOLS = [
  ["bold", "太字（⌘B）", "B", "tool-bold"],
  ["italic", "斜体（⌘I）", "I", "tool-italic"],
  ["underline", "下線（⌘U）", "U", "tool-underline"],
  ["strikeThrough", "取り消し線", "S", "tool-strike"],
  "-",
  ["insertUnorderedList", "箇条書き", "•", "tool-bullet"],
  ["insertOrderedList", "段落番号", "1.", "tool-number"],
  "-",
  ["link", "リンク", "🔗", "tool-link"],
  ["clear", "書式をクリア", "Tx", "tool-clear"],
];

/** A notes box: `{ el, box, show(slide, { force }), focus() }`. `el` is the toolbar with the box under it. */
export function createNotesEditor(app, { id = null, label = "ノート" } = {}) {
  const { h, E } = app;
  const box = h("div", { class: "notes-input", contenteditable: "true", role: "textbox", "aria-multiline": "true", "aria-label": label, "data-placeholder": "ここにノート（話す内容）を入力（発表者ビューに出ます）", spellcheck: "true" });
  if (id) box.id = id;
  let saved = null;
  const remember = () => {
    const sel = window.getSelection();
    if (sel?.rangeCount && box.contains(sel.anchorNode)) saved = sel.getRangeAt(0).cloneRange();
  };
  const restore = () => {
    box.focus({ preventScroll: true });
    if (!saved) return;
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(saved);
  };
  const changed = () => app.setNotesRich(box.innerHTML);
  const buttons = new Map();

  async function link() {
    remember();
    const range = saved;
    if (!range || range.collapsed) { app.toast("リンクにする文字を選んでください"); return; }
    const url = await app.ask("リンク", "https:// から始まるアドレス", "https://");
    if (!url?.trim()) return;
    restore();
    document.execCommand("createLink", false, url.trim());
    changed();
  }
  function run(command) {
    if (command === "link") { link(); return; }
    restore();
    // Tags (<b>, <i>…) rather than styles, whatever another part of the studio left the editor set to.
    document.execCommand("styleWithCSS", false, false);
    if (command === "clear") { document.execCommand("removeFormat"); document.execCommand("unlink"); } else document.execCommand(command);
    remember();
    changed();
    refreshState();
  }
  function refreshState() {
    for (const [command, button] of buttons) {
      let on = false;
      try { on = !["link", "clear"].includes(command) && document.queryCommandState(command); } catch { on = false; }
      button.setAttribute("aria-pressed", String(on));
    }
  }

  const bar = h("div", { class: "notes-tools", role: "toolbar", "aria-label": "ノートの書式" },
    TOOLS.map((tool) => {
      if (tool === "-") return h("span", { class: "notes-tools-sep", "aria-hidden": "true" });
      const [command, title, text, cls] = tool;
      const button = h("button", { type: "button", class: ["notes-tool", cls], title, "aria-label": title, "aria-pressed": "false", onmousedown: (event) => { event.preventDefault(); remember(); }, onclick: () => run(command) }, text);
      buttons.set(command, button);
      return button;
    }));

  box.addEventListener("input", changed);
  box.addEventListener("blur", remember);
  box.addEventListener("keyup", refreshState);
  box.addEventListener("mouseup", refreshState);
  // Pasted words come in as the notes' own formatting only (bold, lists, links), never colours or sizes.
  box.addEventListener("paste", (event) => {
    event.preventDefault();
    const html = event.clipboardData?.getData("text/html");
    const text = event.clipboardData?.getData("text/plain") || "";
    const clean = html ? E.noteHtml(html) : E.textToRich(text);
    document.execCommand("insertHTML", false, clean || E.textToRich(text));
    changed();
  });
  // Tab keeps its place (it moves on to the next control); ⌘B / ⌘I / ⌘U are the box's own.

  const el = h("div", { class: "notes-editor" }, bar, box);
  return {
    el,
    box,
    /** Shows a slide's notes (not while the box is being typed in, unless forced). */
    show(slide, { force = false } = {}) {
      if (!force && (document.activeElement === box || box.contains(document.activeElement))) return;
      const html = E.noteOf(slide).rich;
      if (box.innerHTML !== html) box.innerHTML = html;
      saved = null;
    },
    focus() { box.focus({ preventScroll: true }); },
  };
}
