// Excel スプレッドシート (挿入 → 表 → Excel スプレッドシート): a table whose cells may hold formulas (`cell.f`, "=SUM(B2:B4)").
// Everything here is a pure function on a table's cells. The words a formula cell shows (`cell.text`) are what the formula
// works out to, so drawing, presenting and exporting never need this file; `recalc` keeps them in step, and the
// engine calls it (E.sheetCalc) whenever a table that has `sheet: true` is normalized.

export const MAX_ROWS = 500;
export const MAX_COLS = 100;

export class SheetError extends Error {
  constructor(code, { unknown = false } = {}) {
    super(code);
    this.code = code;
    this.unknown = unknown;
  }
}
const fail = (code) => { throw new SheetError(code); };
const syntax = () => { throw new SheetError("#ERROR!", { unknown: true }); };

/** What each error means, in the words of the screen. */
export const ERROR_HINTS = {
  "#DIV/0!": "0で割っています",
  "#VALUE!": "数値でないものを計算しています",
  "#REF!": "消えたセルを指しています",
  "#NAME?": "知らない関数・名前です",
  "#N/A": "見つかりませんでした",
  "#NUM!": "計算できない数です",
  "#CIRC!": "セルが自分自身をたどっています（循環参照）",
  "#ERROR!": "数式の書き方が正しくありません",
};

// ---------------------------------------------------------------- addresses

export function colName(c) {
  let n = c + 1;
  let out = "";
  while (n > 0) { const m = (n - 1) % 26; out = String.fromCharCode(65 + m) + out; n = Math.floor((n - 1) / 26); }
  return out;
}
export function colIndex(name) {
  let n = 0;
  for (const ch of String(name).toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
export const cellName = (r, c) => `${colName(c)}${r + 1}`;
/** "$B$3" → { r: 2, c: 1, ra: true, ca: true } (null when it is not an address). */
export function parseAddress(text) {
  const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,4})$/.exec(String(text).trim());
  if (!m) return null;
  const r = Number(m[4]) - 1;
  if (r < 0) return null;
  return { r, c: colIndex(m[2]), ra: m[3] === "$", ca: m[1] === "$" };
}
const addressText = ({ r, c, ra, ca }) => `${ca ? "$" : ""}${colName(c)}${ra ? "$" : ""}${r + 1}`;

// ---------------------------------------------------------------- words ⇄ numbers

/** A cell's rich text as plain words (a line per paragraph). */
export function plainOf(html) {
  return String(html ?? "")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/(p|li|div)>\s*(?=<)/gi, "\n")
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, " ").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&amp;/g, "&")
    .trim();
}
const escapeText = (s) => String(s).replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" })[c]);
/** Plain words as a cell's rich text. */
export const richOf = (text) => String(text).split(/\r?\n/).map((line) => `<p>${escapeText(line) || "<br>"}</p>`).join("");

/**
 * The number a text means, or null: 1,200 ・ ¥1,200 ・ 12% ・ （1,200） ・ △1,200 ・ ▲3 ・ １２（full width）・ 5円.
 * Percentages are fractions (12% → 0.12).
 */
export function parseNumber(text) {
  let t = String(text ?? "").normalize("NFKC").trim();
  if (!t) return null;
  let sign = 1;
  if (/^[▲△]/.test(t)) { sign = -1; t = t.slice(1).trim(); }
  const paren = /^\((.*)\)$/.exec(t);
  if (paren) { sign = -sign; t = paren[1].trim(); }
  if (/^[-−]/.test(t)) { sign = -sign; t = t.slice(1).trim(); } else if (t[0] === "+") t = t.slice(1).trim();
  t = t.replace(/^[¥$]\s*/, "");
  let pct = false;
  if (/%$/.test(t)) { pct = true; t = t.slice(0, -1).trim(); }
  if (/円$/.test(t)) t = t.slice(0, -1).trim();
  if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) t = t.replace(/,/g, "");
  if (!/^(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) return null;
  const n = Number(t) * sign / (pct ? 100 : 1);
  return Number.isFinite(n) ? n : null;
}

/** A date written as 2024/4/1, 2024-04-01 or 2024年4月1日 → Excel's serial number (days since 1899-12-30), or null. */
export function parseDate(text) {
  const m = /^(\d{4})[/\-.年](\d{1,2})[/\-.月](\d{1,2})日?$/.exec(String(text ?? "").normalize("NFKC").trim());
  if (!m) return null;
  const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
  const at = new Date(Date.UTC(y, mo - 1, d));
  if (at.getUTCFullYear() !== y || at.getUTCMonth() !== mo - 1 || at.getUTCDate() !== d) return null;
  return Math.round((at.getTime() - Date.UTC(1899, 11, 30)) / 86400000);
}

const groups = (digits) => digits.replace(/\B(?=(\d{3})+(?!\d))/g, ",");
function roundHalfAway(x, digits) {
  const f = 10 ** digits;
  const y = Number((Math.abs(x) * f).toPrecision(15));
  return (Math.sign(x) * Math.round(y)) / f;
}
/** A number as words: up to 12 significant digits, or `dec` decimals; `sep` puts commas in the thousands, `pct` shows a fraction as a percentage, `cur` a ¥ or $ in front. */
export function formatNumber(value, { dec = null, sep = false, pct = false, cur = "" } = {}) {
  let n = pct ? value * 100 : value;
  if (!Number.isFinite(n)) return String(n);
  let s;
  if (dec == null) {
    n = Number(n.toPrecision(12));
    s = String(n);
    if (/e/i.test(s)) return `${cur}${s}${pct ? "%" : ""}`;
  } else {
    n = roundHalfAway(n, dec);
    s = Math.abs(n).toFixed(dec);
    if (n < 0 && Number(s) !== 0) s = `-${s}`;
  }
  const neg = s.startsWith("-");
  let [whole, frac] = (neg ? s.slice(1) : s).split(".");
  if (sep) whole = groups(whole);
  return `${neg ? "-" : ""}${cur}${whole}${frac != null ? `.${frac}` : ""}${pct ? "%" : ""}`;
}

// ---------------------------------------------------------------- scanning a formula

// The same scanner reads a formula to evaluate it and finds its addresses to move them (copy, insert, delete).
const REF = String.raw`\$?[A-Z]{1,3}\$?\d{1,4}`;
const TOKEN = new RegExp([
  String.raw`(?<space>\s+)`,
  String.raw`(?<str>"(?:[^"]|"")*")`,
  String.raw`(?<num>(?:\d+\.?\d*|\.\d+)(?:E[+-]?\d+)?)(?![A-Z_])`,
  String.raw`(?<err>#REF!|#N/A|#DIV/0!|#VALUE!|#NAME\?|#NUM!)`,
  String.raw`(?<ref>(?<a>${REF})(?::(?<b>${REF}))?)(?![A-Z0-9_.(])`,
  String.raw`(?<name>[A-Z_][A-Z0-9_.]*)`,
  String.raw`(?<op><>|<=|>=|[-+*/^&=<>%(),:])`,
].join("|"), "y");

/** The formula the way it is kept: full-width letters and signs (＝ＳＵＭ（Ａ１：Ａ３）) made half-width, letters upper-case (outside "quotes"). */
export function foldFormula(f) {
  const text = String(f ?? "").trim();
  let out = "";
  const parts = text.split(/("(?:[^"]|"")*")/);
  for (const [i, part] of parts.entries()) {
    out += i % 2 ? part : part.normalize("NFKC").replace(/、/g, ",").replace(/[ー−]/g, "-").toUpperCase();
  }
  return out;
}

function scan(src) {
  const tokens = [];
  let at = 0;
  TOKEN.lastIndex = 0;
  while (at < src.length) {
    TOKEN.lastIndex = at;
    const m = TOKEN.exec(src);
    if (!m) { tokens.push({ type: "bad", text: src[at], at }); at += 1; continue; }
    const g = m.groups;
    const type = g.space != null ? "space" : g.str != null ? "str" : g.num != null ? "num" : g.err != null ? "err" : g.ref != null ? "ref" : g.name != null ? "name" : "op";
    tokens.push({ type, text: m[0], at, a: g.a, b: g.b });
    at = TOKEN.lastIndex;
  }
  return tokens;
}
const body = (f) => { const t = foldFormula(f); return t.startsWith("=") ? t.slice(1) : t; };

/** Rewrite every address of a formula: `fn(first, second | null)` → the new text ("#REF!" for a lost address). */
function mapAddresses(f, fn) {
  const tokens = scan(body(f));
  return `=${tokens.map((t) => {
    if (t.type !== "ref") return t.text;
    const a = parseAddress(t.a);
    const b = t.b ? parseAddress(t.b) : null;
    return fn(a, b) ?? "#REF!";
  }).join("")}`;
}
const refText = (a, b) => (b ? `${addressText(a)}:${addressText(b)}` : addressText(a));

/** The formula moved by dr rows and dc columns (copying it down or across): relative addresses move, $ ones stay. */
export function shiftFormula(f, dr, dc) {
  return mapAddresses(f, (a, b) => {
    const move = (x) => ({ ...x, r: x.ra ? x.r : x.r + dr, c: x.ca ? x.c : x.c + dc });
    const [p, q] = [move(a), b ? move(b) : null];
    if (p.r < 0 || p.c < 0 || (q && (q.r < 0 || q.c < 0))) return null;
    return refText(p, q);
  });
}
/** The formula after `count` rows (axis "row") or columns were inserted before index `at`. */
export function insertRefs(f, axis, at, count = 1) {
  const key = axis === "row" ? "r" : "c";
  const bump = (x) => (x[key] >= at ? { ...x, [key]: x[key] + count } : x);
  return mapAddresses(f, (a, b) => refText(bump(a), b ? bump(b) : null));
}
/** The formula after rows (or columns) `from`…`to` were deleted: addresses in them are lost, ranges shrink. */
export function deleteRefs(f, axis, from, to) {
  const key = axis === "row" ? "r" : "c";
  const n = to - from + 1;
  return mapAddresses(f, (a, b) => {
    if (!b) {
      if (a[key] >= from && a[key] <= to) return null;
      return refText(a[key] > to ? { ...a, [key]: a[key] - n } : a, null);
    }
    const [lo, hi] = a[key] <= b[key] ? [a, b] : [b, a];
    let s = lo[key];
    let e = hi[key];
    s = s >= from && s <= to ? from : s > to ? s - n : s;
    e = e >= from && e <= to ? from - 1 : e > to ? e - n : e;
    if (e < s) return null;
    return refText({ ...lo, [key]: s }, { ...hi, [key]: e });
  });
}
/** Every cell a formula reads, as [r, c] (ranges spelled out, up to a limit). */
export function formulaCells(f) {
  const out = [];
  for (const t of scan(body(f))) {
    if (t.type !== "ref") continue;
    const a = parseAddress(t.a);
    const b = t.b ? parseAddress(t.b) : a;
    if (!a || !b) continue;
    for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r) && out.length < 5000; r += 1) for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c) && out.length < 5000; c += 1) out.push([r, c]);
  }
  return out;
}

// ---------------------------------------------------------------- parsing

function parse(src) {
  const tokens = scan(src).filter((t) => t.type !== "space");
  let i = 0;
  const peek = () => tokens[i];
  const next = () => tokens[i++];
  const isOp = (v) => peek()?.type === "op" && peek().text === v;
  const expect = (v) => { if (!isOp(v)) syntax(); i += 1; };

  const comparison = () => {
    let left = concat();
    while (peek()?.type === "op" && ["=", "<>", "<", ">", "<=", ">="].includes(peek().text)) left = { n: "bin", op: next().text, left, right: concat() };
    return left;
  };
  const concat = () => {
    let left = additive();
    while (isOp("&")) { next(); left = { n: "bin", op: "&", left, right: additive() }; }
    return left;
  };
  const additive = () => {
    let left = multiplicative();
    while (isOp("+") || isOp("-")) left = { n: "bin", op: next().text, left, right: multiplicative() };
    return left;
  };
  const multiplicative = () => {
    let left = power();
    while (isOp("*") || isOp("/")) left = { n: "bin", op: next().text, left, right: power() };
    return left;
  };
  const power = () => {
    let left = unary();
    while (isOp("^")) { next(); left = { n: "bin", op: "^", left, right: unary() }; }
    return left;
  };
  // As in Excel, a minus sign binds tighter than ^ (-2^2 is 4).
  const unary = () => {
    if (isOp("-")) { next(); return { n: "neg", value: unary() }; }
    if (isOp("+")) { next(); return unary(); }
    return postfix();
  };
  const postfix = () => {
    let node = primary();
    while (isOp("%")) { next(); node = { n: "pct", value: node }; }
    return node;
  };
  const primary = () => {
    const t = next();
    if (!t) syntax();
    if (t.type === "num") return { n: "num", v: Number(t.text) };
    if (t.type === "str") return { n: "str", v: t.text.slice(1, -1).replace(/""/g, '"') };
    if (t.type === "err") return { n: "err", code: t.text };
    if (t.type === "ref") {
      const a = parseAddress(t.a);
      if (!a) fail("#REF!");
      if (t.b) { const b = parseAddress(t.b); if (!b) fail("#REF!"); return { n: "range", a, b }; }
      return { n: "ref", a };
    }
    if (t.type === "name") {
      if (isOp("(")) {
        next();
        const args = [];
        if (!isOp(")")) {
          for (;;) {
            args.push(isOp(",") || isOp(")") ? { n: "empty" } : comparison());
            if (isOp(",")) { next(); continue; }
            break;
          }
        }
        expect(")");
        return { n: "call", name: t.text, args };
      }
      if (t.text === "TRUE" || t.text === "FALSE") return { n: "bool", v: t.text === "TRUE" };
      return { n: "name", name: t.text };
    }
    if (t.type === "op" && t.text === "(") { const inner = comparison(); expect(")"); return inner; }
    return syntax();
  };
  const tree = comparison();
  if (i < tokens.length) syntax();
  return tree;
}

// ---------------------------------------------------------------- values

class Range {
  constructor(rows) { this.rows = rows; }
  get height() { return this.rows.length; }
  get width() { return this.rows[0]?.length || 0; }
  flat() { return this.rows.flat(); }
}
const isErr = (v) => v instanceof SheetError;

function toNum(v) {
  if (v instanceof Range) { if (v.height * v.width === 1) return toNum(v.rows[0][0]); fail("#VALUE!"); }
  if (isErr(v)) throw v;
  if (typeof v === "number") return v;
  if (v == null || v === "") return 0;
  if (typeof v === "boolean") return v ? 1 : 0;
  const n = parseNumber(v);
  if (n == null) fail("#VALUE!");
  return n;
}
function toText(v) {
  if (v instanceof Range) { if (v.height * v.width === 1) return toText(v.rows[0][0]); fail("#VALUE!"); }
  if (isErr(v)) throw v;
  if (v == null) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  if (typeof v === "number") return formatNumber(v);
  return String(v);
}
function toBool(v) {
  if (v instanceof Range) { if (v.height * v.width === 1) return toBool(v.rows[0][0]); fail("#VALUE!"); }
  if (isErr(v)) throw v;
  if (typeof v === "boolean") return v;
  if (typeof v === "number") return v !== 0;
  if (v == null || v === "") return false;
  const t = String(v).trim().toUpperCase();
  if (t === "TRUE") return true;
  if (t === "FALSE") return false;
  const n = parseNumber(v);
  if (n == null) fail("#VALUE!");
  return n !== 0;
}
const scalar = (v) => (v instanceof Range ? (v.height * v.width === 1 ? v.rows[0][0] : fail("#VALUE!")) : v);

/** Every value of the arguments, ranges spelled out; `numbersOnly` ignores the words in ranges (as SUM does). */
function collect(args, { strict = true } = {}) {
  const out = [];
  for (const a of args) {
    if (a instanceof Range) {
      for (const v of a.flat()) {
        if (isErr(v)) throw v;
        if (typeof v === "number") out.push(v);
        else if (typeof v === "boolean") { /* ignored inside a range */ } else if (typeof v === "string" && v !== "") { const n = parseNumber(v); if (n != null && !strict) out.push(n); }
      }
    } else if (isErr(a)) throw a;
    else out.push(toNum(a));
  }
  return out;
}

const compareValues = (a, b) => {
  a = scalar(a); b = scalar(b);
  if (isErr(a)) throw a;
  if (isErr(b)) throw b;
  if (a == null) a = typeof b === "string" ? "" : typeof b === "boolean" ? false : 0;
  if (b == null) b = typeof a === "string" ? "" : typeof a === "boolean" ? false : 0;
  const rank = (v) => (typeof v === "number" ? 0 : typeof v === "string" ? 1 : 2);
  // Numbers typed as words (1,200) compare as numbers.
  if (typeof a === "string" && typeof b === "number" && parseNumber(a) != null) a = parseNumber(a);
  if (typeof b === "string" && typeof a === "number" && parseNumber(b) != null) b = parseNumber(b);
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  if (typeof a === "string") { const x = a.toLowerCase(); const y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
  return a < b ? -1 : a > b ? 1 : 0;
};

/** What SUMIF / COUNTIF mean by a condition: 100, ">=100", "<>東京", "東*". */
function criterion(c) {
  c = scalar(c);
  if (typeof c === "number") return (v) => typeof v === "number" && v === c;
  let text = toText(c);
  let op = "=";
  const m = /^(>=|<=|<>|>|<|=)/.exec(text);
  if (m) { op = m[1]; text = text.slice(op.length); }
  const n = parseNumber(text);
  const wild = /[*?]/.test(text) ? new RegExp(`^${text.replace(/[.+^${}()|[\]\\]/g, "\\$&").replace(/\*/g, ".*").replace(/\?/g, ".")}$`, "i") : null;
  return (v) => {
    if (isErr(v)) return false;
    if (n != null && (typeof v === "number" || (typeof v === "string" && parseNumber(v) != null))) {
      const x = typeof v === "number" ? v : parseNumber(v);
      return op === "=" ? x === n : op === "<>" ? x !== n : op === ">" ? x > n : op === "<" ? x < n : op === ">=" ? x >= n : x <= n;
    }
    if (n != null && op !== "=" && op !== "<>") return false;
    const s = v == null ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v);
    const same = wild ? wild.test(s) : s.toLowerCase() === text.toLowerCase();
    if (op === "=") return same;
    if (op === "<>") return !same;
    return op === ">" ? s.toLowerCase() > text.toLowerCase() : op === "<" ? s.toLowerCase() < text.toLowerCase() : op === ">=" ? s.toLowerCase() >= text.toLowerCase() : s.toLowerCase() <= text.toLowerCase();
  };
}
const rangeOf = (v) => (v instanceof Range ? v : new Range([[v]]));
const sameShape = (a, b) => a.height === b.height && a.width === b.width;
/** The cells of the conditions (range, criterion pairs) all hold, as a list of positions. */
function matching(pairs) {
  const ranges = pairs.map(([r]) => rangeOf(r));
  if (!ranges.length || !ranges.every((r) => sameShape(r, ranges[0]))) fail("#VALUE!");
  const tests = pairs.map(([, c]) => criterion(c));
  const hits = [];
  ranges[0].rows.forEach((row, i) => row.forEach((_, j) => { if (ranges.every((r, k) => tests[k](r.rows[i][j]))) hits.push([i, j]); }));
  return hits;
}

const EPS = 1e-12;
const FUNCTIONS = {
  SUM: (a) => collect(a).reduce((s, v) => s + v, 0),
  PRODUCT: (a) => collect(a).reduce((s, v) => s * v, 1),
  AVERAGE: (a) => { const v = collect(a); return v.length ? v.reduce((s, x) => s + x, 0) / v.length : fail("#DIV/0!"); },
  MIN: (a) => { const v = collect(a); return v.length ? Math.min(...v) : 0; },
  MAX: (a) => { const v = collect(a); return v.length ? Math.max(...v) : 0; },
  MEDIAN: (a) => { const v = collect(a).sort((x, y) => x - y); if (!v.length) fail("#NUM!"); const h = Math.floor(v.length / 2); return v.length % 2 ? v[h] : (v[h - 1] + v[h]) / 2; },
  COUNT: (a) => a.reduce((n, x) => n + (x instanceof Range ? x.flat().filter((v) => typeof v === "number").length : typeof scalarNumber(x) === "number" ? 1 : 0), 0),
  COUNTA: (a) => a.reduce((n, x) => n + (x instanceof Range ? x.flat().filter((v) => v != null && v !== "").length : x != null && x !== "" ? 1 : 0), 0),
  COUNTBLANK: (a) => a.reduce((n, x) => n + rangeOf(x).flat().filter((v) => v == null || v === "").length, 0),
  LARGE: ([r, k]) => { const v = collect([r]).sort((x, y) => y - x); const i = toNum(k); return i >= 1 && i <= v.length ? v[i - 1] : fail("#NUM!"); },
  SMALL: ([r, k]) => { const v = collect([r]).sort((x, y) => x - y); const i = toNum(k); return i >= 1 && i <= v.length ? v[i - 1] : fail("#NUM!"); },
  SUMPRODUCT: (a) => {
    const ranges = a.map(rangeOf);
    if (!ranges.every((r) => sameShape(r, ranges[0]))) fail("#VALUE!");
    return ranges[0].flat().reduce((sum, _, i) => sum + ranges.reduce((p, r) => { const v = r.flat()[i]; if (isErr(v)) throw v; return p * (typeof v === "number" ? v : 0); }, 1), 0);
  },
  ROUND: ([x, n = 0]) => roundHalfAway(toNum(x), Math.trunc(toNum(n))),
  ROUNDUP: ([x, n = 0]) => { const f = 10 ** Math.trunc(toNum(n)); const v = toNum(x); return (Math.sign(v) * Math.ceil(Number((Math.abs(v) * f).toPrecision(15)) - EPS)) / f; },
  ROUNDDOWN: ([x, n = 0]) => { const f = 10 ** Math.trunc(toNum(n)); const v = toNum(x); return (Math.sign(v) * Math.floor(Number((Math.abs(v) * f).toPrecision(15)) + EPS)) / f; },
  INT: ([x]) => Math.floor(toNum(x)),
  ABS: ([x]) => Math.abs(toNum(x)),
  SQRT: ([x]) => { const v = toNum(x); return v < 0 ? fail("#NUM!") : Math.sqrt(v); },
  POWER: ([x, y]) => { const v = toNum(x) ** toNum(y); return Number.isFinite(v) ? v : fail("#NUM!"); },
  MOD: ([x, y]) => { const d = toNum(y); if (d === 0) fail("#DIV/0!"); const v = toNum(x); return v - d * Math.floor(v / d); },
  PI: () => Math.PI,
  SIGN: ([x]) => Math.sign(toNum(x)),
  AND: (a) => collectBools(a).every(Boolean),
  OR: (a) => collectBools(a).some(Boolean),
  NOT: ([x]) => !toBool(x),
  CONCAT: (a) => a.map((x) => (x instanceof Range ? x.flat().map(toText).join("") : toText(x))).join(""),
  CONCATENATE: (a) => a.map(toText).join(""),
  LEN: ([x]) => [...toText(x)].length,
  LEFT: ([x, n = 1]) => [...toText(x)].slice(0, Math.max(0, toNum(n))).join(""),
  RIGHT: ([x, n = 1]) => { const s = [...toText(x)]; const k = Math.max(0, toNum(n)); return k ? s.slice(-k).join("") : ""; },
  MID: ([x, start, n]) => [...toText(x)].slice(Math.max(1, toNum(start)) - 1, Math.max(1, toNum(start)) - 1 + Math.max(0, toNum(n))).join(""),
  UPPER: ([x]) => toText(x).toUpperCase(),
  LOWER: ([x]) => toText(x).toLowerCase(),
  TRIM: ([x]) => toText(x).trim().replace(/\s+/g, " "),
  VALUE: ([x]) => { const n = parseNumber(toText(x)); return n == null ? fail("#VALUE!") : n; },
  TEXT: ([x, fmt]) => textFormat(toNum(x), toText(fmt)),
  ISNUMBER: ([x]) => typeof scalar(x) === "number",
  ISTEXT: ([x]) => typeof scalar(x) === "string" && scalar(x) !== "",
  ISBLANK: ([x]) => { const v = scalar(x); return v == null || v === ""; },
  SUMIF: ([range, crit, sumRange]) => {
    const base = rangeOf(range);
    const sums = sumRange == null ? base : rangeOf(sumRange);
    if (!sameShape(base, sums)) fail("#VALUE!");
    return matching([[base, crit]]).reduce((s, [i, j]) => { const v = sums.rows[i][j]; if (isErr(v)) throw v; return s + (typeof v === "number" ? v : 0); }, 0);
  },
  SUMIFS: ([sumRange, ...rest]) => {
    const sums = rangeOf(sumRange);
    const pairs = []; for (let i = 0; i + 1 < rest.length; i += 2) pairs.push([rest[i], rest[i + 1]]);
    if (!pairs.length) fail("#VALUE!");
    if (!sameShape(sums, rangeOf(pairs[0][0]))) fail("#VALUE!");
    return matching(pairs).reduce((s, [i, j]) => { const v = sums.rows[i][j]; if (isErr(v)) throw v; return s + (typeof v === "number" ? v : 0); }, 0);
  },
  COUNTIF: ([range, crit]) => matching([[range, crit]]).length,
  COUNTIFS: (a) => { const pairs = []; for (let i = 0; i + 1 < a.length; i += 2) pairs.push([a[i], a[i + 1]]); return pairs.length ? matching(pairs).length : fail("#VALUE!"); },
  AVERAGEIF: ([range, crit, avgRange]) => {
    const base = rangeOf(range);
    const avg = avgRange == null ? base : rangeOf(avgRange);
    if (!sameShape(base, avg)) fail("#VALUE!");
    const v = matching([[base, crit]]).map(([i, j]) => avg.rows[i][j]).filter((x) => typeof x === "number");
    return v.length ? v.reduce((s, x) => s + x, 0) / v.length : fail("#DIV/0!");
  },
  VLOOKUP: ([key, table, col, approx = true]) => {
    const t = rangeOf(table);
    const k = Math.trunc(toNum(col));
    if (k < 1 || k > t.width) fail("#REF!");
    const row = lookupRow(scalar(key), t.rows.map((r) => r[0]), toBool(approx) ? 1 : 0);
    return row < 0 ? fail("#N/A") : t.rows[row][k - 1];
  },
  HLOOKUP: ([key, table, rowNo, approx = true]) => {
    const t = rangeOf(table);
    const k = Math.trunc(toNum(rowNo));
    if (k < 1 || k > t.height) fail("#REF!");
    const col = lookupRow(scalar(key), t.rows[0], toBool(approx) ? 1 : 0);
    return col < 0 ? fail("#N/A") : t.rows[k - 1][col];
  },
  MATCH: ([key, range, type = 1]) => {
    const r = rangeOf(range);
    if (r.height > 1 && r.width > 1) fail("#N/A");
    const at = lookupRow(scalar(key), r.flat(), Math.sign(toNum(type)));
    return at < 0 ? fail("#N/A") : at + 1;
  },
  INDEX: ([range, row = 1, col]) => {
    const r = rangeOf(range);
    let i = Math.trunc(toNum(row));
    let j = col == null ? 1 : Math.trunc(toNum(col));
    if (col == null && r.height === 1) { j = i; i = 1; }
    if (i < 1 || j < 1 || i > r.height || j > r.width) fail("#REF!");
    return r.rows[i - 1][j - 1];
  },
  ROWS: ([x]) => rangeOf(x).height,
  COLUMNS: ([x]) => rangeOf(x).width,
};
function scalarNumber(x) { return typeof x === "number" ? x : typeof x === "string" ? parseNumber(x) : undefined; }
function collectBools(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof Range) { for (const v of a.flat()) { if (isErr(v)) throw v; if (v == null || v === "") continue; out.push(toBool(v)); } } else out.push(toBool(a));
  }
  if (!out.length) fail("#VALUE!");
  return out;
}
/** The position of `key` among `list`: exactly (type 0), the last one not above it (1), the last one not below it (-1). */
function lookupRow(key, list, type) {
  if (type === 0) return list.findIndex((v) => v != null && compareValues(v, key) === 0 && (typeof v === typeof key || (typeof v === "string" && typeof key === "number" && parseNumber(v) === key)));
  let best = -1;
  list.forEach((v, i) => {
    if (v == null || v === "") return;
    const d = compareValues(v, key);
    if ((type > 0 && d <= 0) || (type < 0 && d >= 0)) best = i;
  });
  return best;
}
/** TEXT(value, "0.0") for the formats that matter on a slide: 0 ・ 0.00 ・ #,##0 ・ #,##0.00 ・ 0% ・ 0.0%. */
function textFormat(n, fmt) {
  const m = /^(¥|\$)?(#,##)?([0#]+)(?:\.([0#]+))?(%)?$/.exec(fmt.trim());
  if (!m) return toText(n);
  return formatNumber(n, { dec: m[4] ? m[4].length : 0, sep: Boolean(m[2]), pct: Boolean(m[5]), cur: m[1] || "" });
}
const SPECIAL = new Set(["IF", "IFERROR", "IFNA", "ISERROR"]);

// ---------------------------------------------------------------- evaluating a table

/**
 * A sheet over a table's cells: `cell(r, c)` works out one cell (memoized, loops caught), `values()` every cell.
 * A formula that uses a function this file does not know keeps the words the cell already shows (a sheet from Excel).
 */
export function makeSheet(cells) {
  const rows = cells.length;
  const cols = rows ? cells[0].length : 0;
  const memo = new Map();
  const state = new Map();
  const kept = new Set();
  const key = (r, c) => r * 1024 + c;
  const constant = (cell) => {
    const text = plainOf(cell.text);
    if (!text) return null;
    const n = parseNumber(text);
    if (n != null) return n;
    // Dates take part in sums and differences (2024/4/30 − 2024/4/1) as the number of days, as in Excel.
    const day = parseDate(text);
    return day != null ? day : text;
  };
  function cell(r, c) {
    if (r < 0 || c < 0 || r >= rows || c >= cols) return null;
    const k = key(r, c);
    if (memo.has(k)) return memo.get(k);
    const own = cells[r][c];
    if (!own || own.merged) return null;
    if (typeof own.f !== "string" || !own.f) { const v = constant(own); memo.set(k, v); return v; }
    if (state.get(k)) return new SheetError("#CIRC!");
    state.set(k, true);
    let value;
    try {
      value = evaluate(parse(body(own.f)));
    } catch (e) {
      if (!(e instanceof SheetError)) throw e;
      // An unknown function: what the cell shows stays (an imported sheet), the way Excel keeps the value it saved.
      if (e.unknown && plainOf(own.text)) { value = constant(own); kept.add(k); } else value = e;
    }
    state.set(k, false);
    if (value == null) value = 0;
    if (value instanceof Range) value = value.height * value.width === 1 ? value.rows[0][0] : new SheetError("#VALUE!");
    memo.set(k, value);
    return value;
  }
  function at(r, c) { const v = cell(r, c); if (isErr(v)) throw v; return v; }
  function evaluate(node) {
    switch (node.n) {
      case "num": case "str": case "bool": return node.v;
      case "empty": return null;
      case "err": return new SheetError(node.code);
      case "ref": return at(node.a.r, node.a.c);
      case "name": throw new SheetError("#NAME?", { unknown: true });
      case "range": {
        const [r0, r1] = [Math.min(node.a.r, node.b.r), Math.max(node.a.r, node.b.r)];
        const [c0, c1] = [Math.min(node.a.c, node.b.c), Math.max(node.a.c, node.b.c)];
        if ((r1 - r0 + 1) * (c1 - c0 + 1) > 50000) fail("#REF!");
        const out = [];
        for (let r = r0; r <= r1; r += 1) { const row = []; for (let c = c0; c <= c1; c += 1) row.push(cell(r, c)); out.push(row); }
        return new Range(out);
      }
      case "neg": return -toNum(evaluate(node.value));
      case "pct": return toNum(evaluate(node.value)) / 100;
      case "bin": return binary(node);
      case "call": return call(node);
      default: return fail("#ERROR!");
    }
  }
  function binary({ op, left, right }) {
    const a = evaluate(left);
    const b = evaluate(right);
    switch (op) {
      case "+": return toNum(a) + toNum(b);
      case "-": return toNum(a) - toNum(b);
      case "*": return toNum(a) * toNum(b);
      case "/": { const d = toNum(b); if (d === 0) fail("#DIV/0!"); return toNum(a) / d; }
      case "^": { const v = toNum(a) ** toNum(b); return Number.isFinite(v) ? v : fail("#NUM!"); }
      case "&": return toText(a) + toText(b);
      case "=": return compareValues(a, b) === 0;
      case "<>": return compareValues(a, b) !== 0;
      case "<": return compareValues(a, b) < 0;
      case ">": return compareValues(a, b) > 0;
      case "<=": return compareValues(a, b) <= 0;
      case ">=": return compareValues(a, b) >= 0;
      default: return fail("#ERROR!");
    }
  }
  function call({ name, args }) {
    if (SPECIAL.has(name)) {
      if (name === "IF") {
        if (args.length < 2 || args.length > 3) fail("#ERROR!");
        if (toBool(evaluate(args[0]))) return evaluate(args[1]);
        return args[2] ? evaluate(args[2]) : false;
      }
      if (name === "ISERROR") {
        if (args.length !== 1) fail("#ERROR!");
        try { return evaluate(args[0]) instanceof SheetError; } catch (e) { if (e instanceof SheetError) return true; throw e; }
      }
      if (args.length !== 2) fail("#ERROR!");
      try {
        const v = evaluate(args[0]);
        if (name === "IFNA" && v instanceof SheetError && v.code !== "#N/A") return v;
        return v;
      } catch (e) {
        if (!(e instanceof SheetError)) throw e;
        if (name === "IFNA" && e.code !== "#N/A") throw e;
        return evaluate(args[1]);
      }
    }
    const fn = FUNCTIONS[name];
    if (!fn) throw new SheetError("#NAME?", { unknown: true });
    return fn(args.map((a) => {
      if (a.n === "ref") return cell(a.a.r, a.a.c);
      return evaluate(a);
    }));
  }
  return { rows, cols, cell, kept: (r, c) => kept.has(key(r, c)) };
}

/** The words a value is shown as in a cell, by the cell's own number settings. */
export function showValue(value, own = {}) {
  if (value instanceof SheetError) return value.code;
  if (value == null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return formatNumber(value, { dec: own.dec ?? null, sep: Boolean(own.sep), pct: Boolean(own.pct), cur: own.cur || "" });
  return String(value);
}

/** One cell's value, as a number, words, TRUE/FALSE, null (empty) or a SheetError. */
export function valueAt(cells, r, c) { return makeSheet(cells).cell(r, c); }

const OK_CUR = new Set(["¥", "$"]);
/** The cell's number settings as they are kept: `dec` 0–6, `sep`, `pct`, `cur`. */
export function numberSettings(src) {
  const out = {};
  const dec = Math.round(Number(src?.dec));
  if (src?.dec != null && Number.isFinite(dec) && dec >= 0 && dec <= 6) out.dec = dec;
  if (src?.sep === true) out.sep = true;
  if (src?.pct === true) out.pct = true;
  if (OK_CUR.has(src?.cur)) out.cur = src.cur;
  return out;
}
/** A formula as it is kept: folded, starting with "=", short enough; "" for anything else. */
export function cleanFormula(f) {
  if (typeof f !== "string") return "";
  const t = foldFormula(f);
  return t.startsWith("=") && t.length > 1 && t.length <= 500 ? t : "";
}

/**
 * The table with its formula cells worked out: each shows its value (`text`). The same table comes back when nothing
 * changes. `inPlace` writes into the cells (the engine's own copy, while it normalizes a table).
 */
export function recalc(table, { inPlace = false } = {}) {
  if (!table?.sheet || !Array.isArray(table.cells)) return table;
  if (!table.cells.some((row) => row.some((cell) => cell && typeof cell.f === "string"))) return table;
  let cells = table.cells;
  if (!inPlace) cells = cells.map((row) => row.map((cell) => ({ ...cell })));
  // Formulas are kept folded and short; one that is not valid loses its `f` (the words stay).
  for (const row of cells) for (const cell of row) {
    if (cell && cell.f !== undefined) { const f = cleanFormula(cell.f); if (f) cell.f = f; else delete cell.f; }
  }
  const sheet = makeSheet(cells);
  let changed = false;
  cells.forEach((row, r) => row.forEach((cell, c) => {
    // An imported formula this file cannot work out keeps the words it came with.
    if (!cell || cell.merged || !cell.f) return;
    const value = sheet.cell(r, c);
    if (sheet.kept(r, c)) return;
    const words = showValue(value, cell);
    const rich = words === "" ? undefined : richOf(words);
    if (plainOf(cell.text) !== words || (cell.text === undefined) !== (rich === undefined)) {
      if (rich === undefined) delete cell.text; else cell.text = rich;
      changed = true;
    }
  }));
  if (inPlace) return table;
  const was = JSON.stringify(table.cells);
  return changed || was !== JSON.stringify(cells) ? { ...table, cells } : table;
}

// ---------------------------------------------------------------- editing helpers (table in → table out)

/** What a cell shows while it is being edited: its formula, or its words. */
export function editText(cell) {
  return cell?.f ? cell.f : plainOf(cell?.text);
}
/** The cell after the words typed into it: "=…" is a formula (in a sheet), anything else is words. */
export function enterText(cell, typed, { sheet = true } = {}) {
  const text = String(typed ?? "").replace(/ /g, " ");
  const out = { ...cell };
  const f = sheet ? cleanFormula(text.trim()) : "";
  if (f) { out.f = f; delete out.text; return out; }
  delete out.f;
  if (text.trim()) out.text = richOf(text.trim()); else delete out.text;
  return out;
}
/** `enterText` as a patch for ops.tableCells (what the typed words take away is set to undefined). */
export function cellPatch(cell, typed, options) {
  const made = enterText(cell, typed, options);
  const patch = { ...made };
  for (const k of Object.keys(cell)) if (!(k in made)) patch[k] = undefined;
  return patch;
}
/** Copy a formula cell to another position: its relative addresses move with it. */
export function moveCell(cell, dr, dc) {
  if (!cell?.f) return { ...cell };
  return { ...cell, f: shiftFormula(cell.f, dr, dc) };
}
/** セルのコピー下へ・右へ: the top row (left column) of the block copied over the rest, formulas moving with them. */
export function fillBlock(cells, r0, c0, r1, c1, axis) {
  const out = cells.map((row) => row.map((cell) => ({ ...cell })));
  for (let r = r0; r <= r1; r += 1) for (let c = c0; c <= c1; c += 1) {
    if (axis === "down" ? r === r0 : c === c0) continue;
    const from = axis === "down" ? out[r0][c] : out[r][c0];
    const to = out[r]?.[c];
    if (!from || !to || from.merged || to.merged) continue;
    // The whole cell (its look too) is copied, as Excel's fill does; its formula's relative addresses move with it.
    const moved = moveCell(from, axis === "down" ? r - r0 : 0, axis === "down" ? 0 : c - c0);
    delete moved.rs; delete moved.cs; delete moved.merged;
    for (const k of Object.keys(to)) delete to[k];
    Object.assign(to, moved);
  }
  return out;
}
/** オートSUM: the formula for the cell at (r, c): the numbers right above it, else the numbers to its left. */
export function autoSum(cells, r, c) {
  const sheet = makeSheet(cells);
  const num = (i, j) => typeof sheet.cell(i, j) === "number";
  let top = r;
  while (top - 1 >= 0 && num(top - 1, c)) top -= 1;
  if (top < r) return `=SUM(${cellName(top, c)}:${cellName(r - 1, c)})`;
  let left = c;
  while (left - 1 >= 0 && num(r, left - 1)) left -= 1;
  if (left < c) return `=SUM(${cellName(r, left)}:${cellName(r, c - 1)})`;
  return "";
}

// ---------------------------------------------------------------- number settings of a cell

/** What a typed number tells about its own format: 1,200 → sep; ¥3,500.5 → ¥ sep 1 decimal; 12% → pct. */
export function settingsFromWords(text) {
  const t = String(text ?? "").normalize("NFKC").trim();
  const out = {};
  const m = /\.(\d+)/.exec(t.replace(/%$/, ""));
  if (m) out.dec = Math.min(6, m[1].length);
  if (/\d,\d{3}/.test(t)) { out.sep = true; if (out.dec == null) out.dec = 0; }
  if (/%$/.test(t)) { out.pct = true; if (out.dec == null) out.dec = 0; }
  if (/^[-▲△(]*\s*¥/.test(t)) out.cur = "¥";
  else if (/^[-▲△(]*\s*\$/.test(t)) out.cur = "$";
  return out;
}
/** The decimals a number is shown with (1,234.50 → 2). */
const decimalsOf = (text) => { const m = /\.(\d+)/.exec(String(text).replace(/%$/, "")); return m ? m[1].length : 0; };
/**
 * 桁数を増やす・減らす, 桁区切り, パーセント, 通貨 on one cell: a patch for ops.tableCells, or null when the cell holds no number.
 * A formula's result takes the settings (`dec`, `sep`, `pct`, `cur`); a number typed in is rewritten as the settings show it.
 */
export function restyleNumber(cell, op) {
  const formula = Boolean(cell?.f);
  const words = plainOf(cell?.text);
  const n = formula ? parseNumber(words) ?? (words === "" ? 0 : null) : parseNumber(words);
  if (!formula && n == null) return null;
  if (formula && words && n == null) return null;
  const now = formula ? { ...numberSettings(cell), ...(cell.dec == null ? { dec: decimalsOf(words) } : {}) } : settingsFromWords(words);
  const next = { ...now };
  if (op === "decMore") next.dec = Math.min(6, (now.dec ?? decimalsOf(words)) + 1);
  else if (op === "decLess") next.dec = Math.max(0, (now.dec ?? decimalsOf(words)) - 1);
  else if (op === "sep") { if (next.sep) delete next.sep; else { next.sep = true; if (next.dec == null) next.dec = decimalsOf(words); } }
  // As in Excel: % shows whole percents (26%), and taking it off goes back to the plain number.
  else if (op === "pct") { if (next.pct) { delete next.pct; delete next.dec; } else { next.pct = true; next.dec = 0; } }
  else if (op === "yen") { if (next.cur === "¥") delete next.cur; else next.cur = "¥"; }
  else return null;
  if (formula) return { dec: next.dec, sep: next.sep || undefined, pct: next.pct || undefined, cur: next.cur || undefined };
  return { text: richOf(formatNumber(n, next)) };
}

// ---------------------------------------------------------------- CSV

/** Comma, tab or semicolon separated text → rows of words (quotes, doubled quotes, CRLF and a BOM handled). */
export function parseCsv(text) {
  const src = String(text ?? "").replace(/^﻿/, "");
  const first = src.split(/\r?\n/, 1)[0] || "";
  const count = (ch) => (first.match(new RegExp(ch === "\t" ? "\t" : `\\${ch}`, "g")) || []).length;
  const delim = [",", "\t", ";"].map((d) => [d, count(d)]).sort((a, b) => b[1] - a[1])[0];
  const sep = delim[1] ? delim[0] : ",";
  const rows = [];
  let row = [];
  let cur = "";
  let quoted = false;
  for (let i = 0; i < src.length; i += 1) {
    const ch = src[i];
    if (quoted) {
      if (ch === '"') { if (src[i + 1] === '"') { cur += '"'; i += 1; } else quoted = false; } else cur += ch;
    } else if (ch === '"' && !cur) quoted = true;
    else if (ch === sep) { row.push(cur); cur = ""; } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && src[i + 1] === "\n") i += 1;
      row.push(cur); rows.push(row); row = []; cur = "";
    } else cur += ch;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  while (rows.length && rows[rows.length - 1].every((v) => !v.trim())) rows.pop();
  return rows;
}
/** A grid of words as CSV text (quoted where needed; words Excel would run as a formula get a ' in front). */
export function toCsv(rows) {
  const quote = (v) => {
    let s = String(v ?? "");
    if (/^[=@+\-\t\r]/.test(s) && parseNumber(s) == null) s = `'${s}`;
    return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
  };
  return rows.map((row) => row.map(quote).join(",")).join("\r\n");
}
/** A table's words as rows (a formula cell gives what it works out to). */
export function tableWords(table) {
  return table.cells.map((row) => row.map((cell) => (cell.merged ? "" : plainOf(cell.text))));
}
/** Rows of words as table cells: "=…" is a formula, 12% and 1,200 stay as typed. */
export function cellsFromWords(rows, { sheet = true } = {}) {
  const cols = Math.min(MAX_COLS, Math.max(1, ...rows.map((row) => row.length)));
  return rows.slice(0, MAX_ROWS).map((row) => Array.from({ length: cols }, (_, c) => enterText({}, row[c] ?? "", { sheet })));
}

// ---------------------------------------------------------------- 構造の変更（行・列の挿入と削除）

/** Every formula of the table rewritten by `fn(formula)` (a table without formulas comes back as it is). */
export function mapFormulas(table, fn) {
  if (!table?.cells?.some((row) => row.some((cell) => cell?.f))) return table;
  return { ...table, cells: table.cells.map((row) => row.map((cell) => (cell?.f ? { ...cell, f: fn(cell.f) } : cell))) };
}

// ---------------------------------------------------------------- 並べ替え and フィルター (the rows of a sheet)

export const FILTER_OPS = [
  ["eq", "と等しい"], ["ne", "と等しくない"], ["contains", "を含む"], ["gt", "より大きい"], ["ge", "以上"], ["lt", "より小さい"], ["le", "以下"], ["nonblank", "空白でない"], ["blank", "空白"],
];
const SORT_RANK = (v) => (v === null || v === undefined || v === "" ? 4 : v instanceof SheetError ? 3 : typeof v === "boolean" ? 2 : typeof v === "number" ? 0 : 1);

/** The row a sort of this table starts and ends at when one cell is picked: under the header row, above a totals row. */
export function sortSpan(table) {
  const first = table.header === false ? 0 : 1;
  const last = table.cells.length - 1 - (table.lastRow ? 1 : 0);
  return [first, last];
}

/**
 * 並べ替え: the rows r0…r1 in the order of one column — 昇順 (numbers, then words, then TRUE/FALSE, then errors) or 降順,
 * blanks last either way. A formula moves with its row and its references to other cells in that row follow it, as in
 * Excel (a reference written with $ stays). Returns the new cells, or null when there is nothing to sort or a merged cell is in the way.
 */
export function sortRows(cells, { col, desc = false, r0, r1 }) {
  if (!Array.isArray(cells) || !(r1 > r0) || col < 0 || r0 < 0 || r1 >= cells.length) return null;
  for (let r = r0; r <= r1; r += 1) if (cells[r].some((cell) => cell && (cell.merged || cell.rs > 1 || cell.cs > 1))) return null;
  const sheet = makeSheet(cells);
  const entries = [];
  for (let r = r0; r <= r1; r += 1) entries.push({ r, v: sheet.cell(r, col) });
  entries.sort((a, b) => {
    const ra = SORT_RANK(a.v);
    const rb = SORT_RANK(b.v);
    if (ra === 4 || rb === 4) return ra === rb ? 0 : ra === 4 ? 1 : -1;
    if (ra !== rb) return desc ? rb - ra : ra - rb;
    const d = ra === 0 ? a.v - b.v : ra === 1 ? String(a.v).localeCompare(String(b.v), "ja") : ra === 2 ? Number(a.v) - Number(b.v) : String(a.v.code).localeCompare(String(b.v.code));
    return desc ? -d : d;
  });
  if (entries.every((e, i) => e.r === r0 + i)) return cells;
  const moved = entries.map((e, i) => cells[e.r].map((cell) => (cell?.f && e.r !== r0 + i ? { ...cell, f: shiftFormula(cell.f, r0 + i - e.r, 0) } : cell)));
  return [...cells.slice(0, r0), ...moved, ...cells.slice(r1 + 1)];
}

/** Whether a cell's words meet a filter's condition (numbers compare as numbers when both are, else as words). */
function filterMatch(words, { op, value }) {
  const text = String(words ?? "").trim();
  const want = String(value ?? "").trim();
  if (op === "blank") return text === "";
  if (op === "nonblank") return text !== "";
  if (op === "contains") return want === "" || text.toLowerCase().includes(want.toLowerCase());
  // A blank cell meets no comparison (but is "not equal" to anything written).
  if (text === "") return op === "ne" && want !== "";
  const a = parseNumber(text);
  const b = parseNumber(want);
  const numbers = a != null && b != null;
  const cmp = numbers ? a - b : text.localeCompare(want, "ja");
  return { eq: cmp === 0, ne: cmp !== 0, gt: cmp > 0, ge: cmp >= 0, lt: cmp < 0, le: cmp <= 0 }[op] ?? true;
}
/**
 * フィルター: the rows a table's `filter` ({ col, op, value }) hides — the data rows that fail it (the header row and a totals
 * row stay). A table with merged cells is not filtered (a hidden row could cut a merged cell in two).
 */
export function filterRows(table) {
  const f = table?.filter;
  if (!f || !Array.isArray(table.cells) || !FILTER_OPS.some(([op]) => op === f.op) || !(f.col >= 0)) return [];
  if (table.cells.some((row) => row.some((cell) => cell && (cell.merged || cell.rs > 1 || cell.cs > 1)))) return [];
  const [first, last] = sortSpan(table);
  const hide = [];
  for (let r = first; r <= last; r += 1) if (!filterMatch(plainOf(table.cells[r][f.col]?.text), f)) hide.push(r);
  return hide;
}
/** Fills `table.hide` from `table.filter` in place (the hook the engine calls whenever a sheet is normalised). */
export function applyFilter(table) {
  if (!table) return table;
  const hide = filterRows(table);
  if (hide.length) table.hide = hide; else delete table.hide;
  return table;
}
