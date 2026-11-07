// Who is working on the deck: the name on review comments (校閲) and in a shared editing session (共有), as
// PowerPoint takes the author from the signed-in account. The name is typed once and kept in this browser; a random
// id tells this browser's comments apart even after the name changes.

const NAME_KEY = "hpptx-user-name";
const ID_KEY = "hpptx-user-id";
const COLORS = ["#2b6be0", "#c2410c", "#0f766e", "#7c3aed", "#b45309", "#be185d", "#15803d", "#4338ca", "#a16207", "#0e7490"];

const read = (key) => { try { return localStorage.getItem(key) || ""; } catch { return ""; } };
const write = (key, value) => { try { localStorage.setItem(key, value); } catch { /* private window */ } };

/** The name typed for this browser ("" until it is set). */
export function userName() {
  return cleanName(read(NAME_KEY));
}
export function setUserName(name) {
  const clean = cleanName(name);
  if (clean) write(NAME_KEY, clean);
  return clean;
}
export function cleanName(name) {
  return String(name || "").replace(/[\u0000-\u001f<>]/g, "").replace(/\s+/g, " ").trim().slice(0, 40);
}
/** A random id for this browser, made the first time it is asked for. */
export function userId() {
  let id = read(ID_KEY);
  if (!/^[a-z0-9]{6,24}$/.test(id)) {
    id = (globalThis.crypto?.randomUUID?.() || `${Date.now().toString(36)}${Math.random().toString(36).slice(2)}`).replace(/[^a-z0-9]/gi, "").toLowerCase().slice(0, 16);
    write(ID_KEY, id);
  }
  return id;
}
/** One or two letters for an avatar: the first letters of two words, or the first two characters of a CJK name. */
export function initials(name) {
  const n = cleanName(name);
  if (!n) return "?";
  const words = n.split(" ").filter(Boolean);
  if (words.length > 1 && /^[A-Za-z]/.test(words[0]) && /^[A-Za-z]/.test(words[1])) return (words[0][0] + words[1][0]).toUpperCase();
  const chars = [...n];
  return /^[A-Za-z]/.test(chars[0]) ? chars[0].toUpperCase() : chars.slice(0, 2).join("");
}
/** A steady color for a person (the same name always gets the same one). */
export function colorOf(key) {
  let hash = 0;
  for (const ch of String(key || "")) hash = (hash * 31 + ch.codePointAt(0)) >>> 0;
  return COLORS[hash % COLORS.length];
}
/** An avatar circle (h: the app's element helper). */
export function avatar(h, name, { size = 22, key = name, title = name } = {}) {
  const text = initials(name);
  const scale = [...text].length > 1 && !/^[A-Z]+$/.test(text) ? 0.38 : 0.45; // two CJK characters need a smaller size
  return h("span", { class: "pp-avatar", title: title || "名前なし", style: { width: `${size}px`, height: `${size}px`, background: colorOf(key || name), fontSize: `${Math.round(size * scale)}px` } }, text);
}

/** The @names in a text: "@" followed by a known name, or by any run of non-space characters. */
export function mentionsIn(text, known = []) {
  const out = new Set();
  const names = [...known].filter(Boolean).sort((a, b) => b.length - a.length);
  const src = String(text || "");
  for (let i = src.indexOf("@"); i >= 0; i = src.indexOf("@", i + 1)) {
    if (i > 0 && /[A-Za-z0-9._%+-]/.test(src[i - 1])) continue; // an e-mail address, not a mention
    const rest = src.slice(i + 1);
    const hit = names.find((n) => rest.startsWith(n));
    if (hit) { out.add(hit); continue; }
    const word = rest.match(/^[^\s@、。,.!?！？「」()（）]+/);
    if (word) out.add(word[0]);
  }
  return [...out];
}
/** The text as nodes, with @mentions marked (never as HTML: comments are people's plain text). */
export function mentionNodes(h, text, known = [], me = "") {
  const src = String(text || "");
  const names = [...known].filter(Boolean).sort((a, b) => b.length - a.length);
  const nodes = [];
  let from = 0;
  for (let i = src.indexOf("@"); i >= 0; i = src.indexOf("@", i + 1)) {
    if (i > 0 && /[A-Za-z0-9._%+-]/.test(src[i - 1])) continue;
    const rest = src.slice(i + 1);
    const name = names.find((n) => rest.startsWith(n)) || rest.match(/^[^\s@、。,.!?！？「」()（）]+/)?.[0];
    if (!name) continue;
    if (i > from) nodes.push(src.slice(from, i));
    nodes.push(h("span", { class: ["pp-mention", me && name === me ? "me" : ""] }, `@${name}`));
    from = i + 1 + name.length;
    i = from - 1;
  }
  if (from < src.length) nodes.push(src.slice(from));
  return nodes;
}
