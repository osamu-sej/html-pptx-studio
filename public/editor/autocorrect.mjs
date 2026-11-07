// オートコレクト (PowerPoint's AutoCorrect while typing): "(c)" becomes ©, "-->" an arrow, and a paragraph begun with
// "・ " or "1. " becomes a bulleted or numbered list (in the marker's style). ⌘Z takes a correction back (it is typed
// like any other text), and 校閲 → オートコレクトのオプション turns each kind off (kept in localStorage).

export const AUTOCORRECT_KEY = "hpptx-autocorrect";
export const AUTOCORRECT_DEFAULTS = { replace: true, bullets: true, numbers: true };

/** 入力中に自動修正する文字列 (PowerPoint's list, those that fit Japanese slides). */
export const REPLACEMENTS = [
  ["(c)", "©"], ["(C)", "©"], ["(r)", "®"], ["(R)", "®"], ["(tm)", "™"], ["(TM)", "™"], ["(e)", "€"],
  ["...", "…"], ["-->", "→"], ["<--", "←"], ["==>", "⇒"], ["<==", "⇐"], ["<=>", "⇔"], ["<->", "↔"],
];

/** Markers that start a bulleted paragraph, and the bullet style each becomes. */
const BULLETS = [["・", "disc"], ["-", "dash"], ["*", "disc"], ["●", "disc"], ["○", "circle"], ["■", "square"], ["◆", "diamond"], [">", "arrow"], ["➢", "arrow"], ["✓", "check"], ["★", "star"]];
/** Markers that start a numbered one (the first number only), and the numbering style. */
const NUMBERS = [
  [/^1[.．]$/, "decimal"], [/^1[)）]$/, "paren"], [/^[(（]1[)）]$/, "paren"], [/^①$/, "circled"],
  [/^a[.)]$/, "lower-alpha"], [/^A[.)]$/, "upper-alpha"], [/^i[.)]$/, "lower-roman"], [/^I[.)]$/, "upper-roman"], [/^一、$/, "kanji"],
];

/** The options kept (each on unless turned off). */
export function autoCorrectOptions(storage = globalThis.localStorage) {
  try {
    const saved = JSON.parse(storage?.getItem(AUTOCORRECT_KEY) || "{}");
    return Object.fromEntries(Object.entries(AUTOCORRECT_DEFAULTS).map(([k, v]) => [k, typeof saved[k] === "boolean" ? saved[k] : v]));
  } catch { return { ...AUTOCORRECT_DEFAULTS }; }
}
export function saveAutoCorrectOptions(options, storage = globalThis.localStorage) {
  const kept = Object.fromEntries(Object.keys(AUTOCORRECT_DEFAULTS).map((k) => [k, options?.[k] !== false]));
  try { storage?.setItem(AUTOCORRECT_KEY, JSON.stringify(kept)); } catch { /* private window */ }
  return kept;
}

/**
 * What to correct now that `typed` was typed, given the paragraph's text up to the caret (typed included) and
 * whether the paragraph is already in a list:
 *   { replace: n, with: "©" }          — the last n characters become the symbol;
 *   { list: "bullet"|"number", style, remove: n } — the paragraph becomes a list and its first n characters go;
 *   null                                — nothing.
 */
export function autoCorrect(before, typed, { inList = false, options = AUTOCORRECT_DEFAULTS } = {}) {
  if (typeof before !== "string" || !typed) return null;
  // A marker and a space at the paragraph's start (a full-width space too).
  if (!inList && /^[ 　]$/.test(typed)) {
    const marker = before.slice(0, -1);
    if (options.bullets !== false) {
      const bullet = BULLETS.find(([m]) => m === marker);
      if (bullet) return { list: "bullet", style: bullet[1], remove: before.length };
    }
    if (options.numbers !== false) {
      const number = NUMBERS.find(([re]) => re.test(marker));
      if (number) return { list: "number", style: number[1], remove: before.length };
    }
  }
  if (options.replace !== false) {
    for (const [from, to] of REPLACEMENTS) {
      if (before.endsWith(from) && from.endsWith(typed.slice(-1))) return { replace: from.length, with: to };
    }
  }
  return null;
}

export function createAutoCorrect(app) {
  const { h } = app;
  /** 校閲 → オートコレクトのオプション. */
  function openOptions() {
    const now = autoCorrectOptions();
    const box = (key, label) => {
      const input = h("input", { type: "checkbox", checked: now[key], "data-ac": key });
      return h("label", { class: "sh-choice" }, input, h("span", {}, label));
    };
    const rows = [box("replace", "入力中に自動修正する（下の一覧）"), box("bullets", "入力中に自動で箇条書きにする（「・」「-」「●」などとスペースで始めたとき）"), box("numbers", "入力中に自動で段落番号を付ける（「1.」「(1)」「①」「a.」などとスペースで始めたとき）")];
    const list = h("table", { class: "ac-list" }, h("thead", {}, h("tr", {}, h("th", {}, "修正文字列"), h("th", {}, "修正後の文字列"))),
      h("tbody", {}, ...REPLACEMENTS.filter(([from]) => from === from.toLowerCase() || !REPLACEMENTS.some(([f]) => f === from.toLowerCase())).map(([from, to]) => h("tr", {}, h("td", {}, from), h("td", {}, to)))));
    const dialog = h("dialog", { class: "ac-dialog", "aria-label": "オートコレクト" },
      h("div", { class: "dialog-head" }, h("h3", {}, "オートコレクト"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body sh-form-col" }, ...rows, list, h("p", { class: "hint" }, "自動で直された直後に ⌘Z（Ctrl+Z）を押すと、入力したとおりに戻ります。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary ac-ok", onclick: () => {
          const next = Object.fromEntries([...dialog.querySelectorAll("[data-ac]")].map((el) => [el.dataset.ac, el.checked]));
          saveAutoCorrectOptions(next);
          dialog.close();
          app.toast("オートコレクトの設定を保存しました");
        } }, "OK")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
  }
  return { openOptions };
}
