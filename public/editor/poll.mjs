// 挿入 → アンケート (PowerPoint Live's polls / Microsoft Forms in a slide): a question and its choices. In a slide
// show each choice can be clicked to answer; presented online, the viewers answer from their own screens and the
// bars fill as the answers come in (motion.js, server/rooms.mjs). Double-click (or 編集) to change the question and
// the choices.

export const POLL_DEFAULT = { question: "どの案がよいと思いますか？", options: ["A案：店舗から始める", "B案：本部から始める", "C案：同時に始める"] };

/** The choices typed one per line: tidy, two to eight. */
export function parseOptions(text) {
  return String(text || "").split(/\r?\n/).map((line) => line.replace(/^\s*(?:[-・*]|\d+[.)．])\s*/, "").trim()).filter(Boolean).slice(0, 8);
}

export function createPolls(editor, app) {
  const { h, E } = app;

  function insert() {
    const o = { id: E.newObjectId(), kind: "poll", x: 0, y: 0, w: 1100, h: 520, ...POLL_DEFAULT };
    const [placed] = editor.insert([o]);
    edit(placed?.id || o.id);
  }

  /** The question and the choices (one per line). */
  function edit(id) {
    const o = editor.objects().find((x) => x.id === id && x.kind === "poll");
    if (!o) return;
    const question = h("input", { type: "text", class: "pl-question", maxlength: 200, value: o.question, "aria-label": "質問" });
    const options = h("textarea", { class: "pl-options", rows: 6, "aria-label": "選択肢（1行に1つ）" });
    options.value = o.options.join("\n");
    const note = h("p", { class: "hint" });
    const check = () => { const n = parseOptions(options.value).length; note.textContent = n < 2 ? "選択肢は2つ以上入れてください" : `選択肢 ${n}個（8個まで）`; ok.disabled = n < 2 || !question.value.trim(); };
    const ok = h("button", { type: "button", class: "btn btn-primary pl-ok", onclick: () => {
      editor.apply((x) => (x.id === id ? { question: question.value.trim(), options: parseOptions(options.value) } : null), { ids: [id] });
      dialog.close();
    } }, "OK");
    options.addEventListener("input", check);
    question.addEventListener("input", check);
    const dialog = h("dialog", { class: "poll-dialog", "aria-label": "アンケート" },
      h("div", { class: "dialog-head" }, h("h3", {}, "アンケート"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body sh-form-col" },
        h("label", { class: "pl-field" }, h("span", {}, "質問"), question),
        h("label", { class: "pl-field" }, h("span", {}, "選択肢（1行に1つ）"), options),
        note,
        h("p", { class: "hint" }, "発表中は選択肢をクリックして回答できます。オンライン プレゼンテーションでは、視聴者が自分の画面から答え、結果がその場で棒グラフに出ます。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"), ok));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
    check();
    question.focus();
    question.select();
  }

  return { insert, edit };
}
