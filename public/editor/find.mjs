// 検索・置換 (Ctrl+F / Ctrl+H), PowerPoint style: 大文字と小文字を区別する, 完全に一致する単語だけ, and (as in Word)
// 全角と半角を区別する. Everything here is a pure function on one text at a time; the dialog in app.js walks the deck.

const HALF_KANA = /[ｦ-ﾝ]/;
const KANA_MARK = /[ﾞﾟ]/;
const WORD_CHAR = /[A-Za-z0-9_]/;

/** A text folded for comparing: width folded (全角の英数字・半角のカナ become the usual ones) and/or lower-cased.
 *  `from`/`to` give, for each character of `folded`, the span of the original text it came from. */
export function foldText(text, { width = true, caseSensitive = false } = {}) {
  const s = String(text ?? "");
  let folded = "";
  const from = [];
  const to = [];
  let i = 0;
  while (i < s.length) {
    let len = s.codePointAt(i) > 0xffff ? 2 : 1;
    // A half-width kana and its voiced mark (ｶﾞ) are one letter (ガ).
    if (width && HALF_KANA.test(s[i]) && KANA_MARK.test(s[i + 1] || "")) len += 1;
    let piece = s.slice(i, i + len);
    if (width) piece = piece.normalize("NFKC");
    if (!caseSensitive) piece = piece.toLowerCase();
    for (let k = 0; k < piece.length; k += 1) { folded += piece[k]; from.push(i); to.push(i + len); }
    i += len;
  }
  return { folded, from, to };
}

/**
 * What a search for `needle` means: `matcher({ needle, caseSensitive, wholeWord, width })` → null for an empty needle, else
 * { ranges(text) → [[start, end], …] in the original text, count(text), replace(text, replacement) }.
 * `width: true` (the default, like Word) treats ＤＸ and DX, ｱ and ア alike; `wholeWord` wants a word of ASCII letters
 * and digits to stand alone (DX in "DX化" does; DX in "ADX" does not).
 */
export function matcher({ needle, caseSensitive = false, wholeWord = false, width = true } = {}) {
  const wanted = String(needle ?? "");
  if (!wanted) return null;
  const key = foldText(wanted, { width, caseSensitive }).folded;
  if (!key) return null;
  const ranges = (text) => {
    const out = [];
    const t = String(text ?? "");
    if (!t) return out;
    const { folded, from, to } = foldText(t, { width, caseSensitive });
    let at = folded.indexOf(key);
    while (at >= 0) {
      const end = at + key.length;
      const before = folded[at - 1];
      const after = folded[end];
      const alone = !wholeWord || ((!WORD_CHAR.test(key[0]) || !before || !WORD_CHAR.test(before)) && (!WORD_CHAR.test(key.at(-1)) || !after || !WORD_CHAR.test(after)));
      if (alone) {
        out.push([from[at], to[end - 1]]);
        at = folded.indexOf(key, end);
      } else at = folded.indexOf(key, at + 1);
    }
    return out;
  };
  return {
    ranges,
    count: (text) => ranges(text).length,
    replace: (text, replacement) => {
      const t = String(text ?? "");
      const found = ranges(t);
      if (!found.length) return t;
      let out = "";
      let last = 0;
      for (const [start, end] of found) { out += t.slice(last, start) + replacement; last = end; }
      return out + t.slice(last);
    },
  };
}

/** A few words round a match, for the list of hits. */
export function snippet(text, range, around = 14) {
  const t = String(text ?? "");
  if (!range) return t.replace(/\s+/g, " ").slice(0, around * 2);
  const start = Math.max(0, range[0] - around);
  const end = Math.min(t.length, range[1] + around);
  return `${start > 0 ? "…" : ""}${t.slice(start, end).replace(/\s+/g, " ")}${end < t.length ? "…" : ""}`;
}
