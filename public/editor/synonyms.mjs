// 校閲 → 類義語 (⇧F7): the words that can stand for the selected one. The built-in thesaurus answers at once; the AI
// (/api/decks/synonyms) is asked for the words that are not in it. A word picked takes the place of the selected words
// while typing, or is copied when nothing was selected.

import { synonymsOf } from "./thesaurus.mjs";

export function createSynonyms(app, editor) {
  const { h } = app;

  function open(initial) {
    const typed = editor.typing ? editor.selectedWords().trim() : "";
    const canReplace = Boolean(typed);
    // data-keeps-text: looking a word up does not end the typing it came from (the pick takes the words' place there).
    const dialog = h("dialog", { class: "syn-dialog", "aria-label": "類義語", "data-keeps-text": "" });
    const input = h("input", { type: "text", name: "word", "aria-label": "調べる言葉", maxlength: "40", placeholder: "調べたい言葉", value: (initial ?? typed).trim().slice(0, 40) });
    const list = h("div", { class: "syn-list", "aria-live": "polite" });
    const status = h("p", { class: "hint syn-status" });
    const aiButton = h("button", { type: "button", class: "btn btn-sm syn-ai", title: "辞書にない言葉をAIで探す（Codexに接続しているとき）", disabled: app.canAi?.() ? null : "", onclick: () => ask() }, "AIで探す");
    let senses = [];
    let aiSenses = [];

    const choose = async (word) => {
      if (canReplace) {
        dialog.close();
        if (editor.replaceWords(word)) app.toast(`「${typed}」を「${word}」にしました`);
      } else {
        try { await navigator.clipboard?.writeText(word); app.toast(`「${word}」をコピーしました`); } catch { app.toast(`「${word}」を選びました`); }
      }
    };
    const section = (sense, source) => h("section", { class: "syn-sense", "data-source": source },
      h("h4", {}, sense.label, source === "ai" ? h("small", {}, "AI") : null),
      h("div", { class: "syn-words" }, sense.words.map((word) => h("button", { type: "button", class: "btn btn-sm syn-word", onclick: () => choose(word) }, word))));
    const render = () => {
      list.replaceChildren(...senses.map((s) => section(s, "dictionary")), ...aiSenses.map((s) => section(s, "ai")));
      const word = input.value.trim();
      status.textContent = !word ? "言葉を入れてください。" : senses.length || aiSenses.length ? (canReplace ? "選ぶと、選んでいた言葉と入れ替わります。" : "選ぶとコピーします（文字を選んでから開くと入れ替えられます）。") : "辞書に見つかりません。「AIで探す」で探せます。";
    };
    const search = () => { senses = synonymsOf(input.value); aiSenses = []; render(); };
    async function ask() {
      const word = input.value.trim();
      if (!word) return;
      aiButton.disabled = true;
      status.textContent = "AIで探しています…";
      try {
        const job = await app.aiRun("/api/decks/synonyms", { word, context: editor.typingText?.().slice(0, 300) || undefined });
        aiSenses = job.senses || [];
        render();
        if (!aiSenses.length) status.textContent = "AIも言い換えを見つけられませんでした。";
      } catch (error) {
        status.textContent = `AIで探せませんでした（${error.message}）`;
      } finally { aiButton.disabled = app.canAi?.() ? false : true; }
    }
    input.addEventListener("input", search);
    input.addEventListener("keydown", (event) => { if (event.key === "Enter") { event.preventDefault(); search(); } });
    dialog.append(
      h("div", { class: "dialog-head" }, h("h3", {}, "類義語"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body syn-body" }, h("div", { class: "syn-search" }, input, aiButton), status, list),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "閉じる")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    search();
    dialog.showModal();
    input.focus();
    input.select();
  }

  return { open };
}
