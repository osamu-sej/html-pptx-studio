// 校閲 → 表記ゆれチェック (Japanese PowerPoint's 表記ゆれ): the same word written two ways in one deck — full-width and
// half-width letters and digits (ＡＩ／AI), half-width katakana (ｻｰﾋﾞｽ／サービス), a katakana word with and without its
// last long vowel (ユーザ／ユーザー), and okurigana (行う／行なう, 打合せ／打ち合わせ) — counted, and made one with a click.

const ALNUM = "A-Za-z0-9Ａ-Ｚａ-ｚ０-９";
const KANA = "ァ-ヶー";
const OKURIGANA = [
  ["行う", "行なう"], ["表す", "表わす"], ["現す", "現わす"], ["終わる", "終る"], ["変わる", "変る"], ["押さえる", "押える"],
  ["取り組み", "取組み", "取組"], ["申し込み", "申込み", "申込"], ["打ち合わせ", "打合せ", "打ち合せ"], ["問い合わせ", "問合せ", "問い合せ"],
  ["引き継ぎ", "引継ぎ", "引継"], ["売り上げ", "売上げ"], ["受け付け", "受付け"], ["見積もり", "見積り"], ["組み合わせ", "組合せ", "組み合せ"],
  ["割り引き", "割引き"], ["取り扱い", "取扱い"], ["手続き", "手続"], ["話し合い", "話合い"], ["届け出", "届出"],
];
const escape = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const widthKey = (s) => s.normalize("NFKC");

/** Every way the deck writes a word, where it writes it more than one way: [{ key, kind, forms: [{ text, count }] }]. */
export function variantGroups(texts) {
  const all = texts.map((t) => String(t ?? "").replace(/<[^>]*>/g, " ")).join("\n");
  const groups = new Map();
  const add = (kind, key, text) => {
    const id = `${kind}:${key}`;
    if (!groups.has(id)) groups.set(id, { key, kind, forms: new Map() });
    const forms = groups.get(id).forms;
    forms.set(text, (forms.get(text) || 0) + 1);
  };
  // Letters and digits: full-width or half-width.
  for (const m of all.matchAll(new RegExp(`[${ALNUM}]+`, "g"))) if (/[Ａ-Ｚａ-ｚ０-９]/.test(m[0]) || /[A-Za-z]{2,}|\d{2,}/.test(m[0])) add("width", widthKey(m[0]), m[0]);
  // Half-width katakana words.
  for (const m of all.matchAll(/[ｦ-ﾟ]{2,}/g)) add("kana", widthKey(m[0]).replace(/ー$/, ""), m[0]);
  // Katakana words: the last long vowel.
  for (const m of all.matchAll(new RegExp(`[${KANA}]{3,}`, "g"))) add("kana", m[0].replace(/ー$/, ""), m[0]);
  // Okurigana.
  for (const forms of OKURIGANA) {
    const re = new RegExp(forms.slice().sort((a, b) => b.length - a.length).map(escape).join("|"), "g");
    for (const m of all.matchAll(re)) add("okurigana", forms[0], m[0]);
  }
  return [...groups.values()].filter((g) => g.forms.size > 1).map((g) => ({ ...g, forms: [...g.forms].map(([text, count]) => ({ text, count })).sort((a, b) => b.count - a.count) }))
    .sort((a, b) => b.forms.reduce((s, f) => s + f.count, 0) - a.forms.reduce((s, f) => s + f.count, 0));
}

/** The text with every form of the group written as `chosen` (whole words only; the longest form wins where two overlap). */
export function unify(text, group, chosen) {
  const forms = group.forms.map((f) => f.text).sort((a, b) => b.length - a.length);
  const edge = group.kind === "width" ? ALNUM : group.kind === "kana" ? `${KANA}ｦ-ﾟ` : null;
  const re = new RegExp(edge ? `(?<![${edge}])(?:${forms.map(escape).join("|")})(?![${edge}])` : forms.map(escape).join("|"), "g");
  return String(text ?? "").replace(re, (m) => (m === chosen ? m : chosen));
}

const KIND_LABEL = { width: "全角・半角", kana: "カタカナ（長音・半角）", okurigana: "送り仮名" };

export function createProofing(app) {
  const { h } = app;
  function open() {
    const dialog = h("dialog", { class: "proof-dialog", "aria-label": "表記ゆれチェック" });
    const body = h("div", { class: "dialog-body" });
    function render() {
      const groups = variantGroups(app.allTexts());
      body.replaceChildren(groups.length ? h("p", { class: "hint" }, `${groups.length}組の表記ゆれがあります。使う表記を選んで「統一」を押すと、資料全体（本文・部品・ノート・詳細）で置き換えます。`) : h("p", { class: "proof-ok" }, "表記ゆれは見つかりませんでした。"),
        h("div", { class: "proof-list" }, groups.map((g, k) => {
          const name = `pf${k}`;
          return h("div", { class: "proof-group", "data-kind": g.kind },
            h("div", { class: "proof-kind" }, KIND_LABEL[g.kind]),
            h("div", { class: "proof-forms" }, g.forms.map((f, i) => h("label", { class: "sh-choice" }, h("input", { type: "radio", name, value: f.text, checked: i === 0 || null }), h("span", { class: "proof-word" }, f.text), h("small", {}, `${f.count}か所`)))),
            h("button", { type: "button", class: "btn btn-sm proof-fix", onclick: () => {
              const chosen = body.querySelector(`input[name="${name}"]:checked`)?.value;
              if (!chosen) return;
              app.transformAllText((text) => unify(text, g, chosen));
              app.toast(`「${chosen}」に統一しました（⌘Zで戻せます）`);
              render();
            } }, "統一"));
        })));
    }
    dialog.append(
      h("div", { class: "dialog-head" }, h("h3", {}, "表記ゆれチェック"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      body,
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "閉じる")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    render();
    dialog.showModal();
  }
  return { open };
}
