// スライド ショー → コーチによるリハーサル (PowerPoint's Rehearse with Coach): present from the start while the
// browser's speech recognition listens; the coach shows the pace as you go and, at the end, a report: how fast you
// spoke, the filler words, words said twice in a row, slides whose text you read out word for word, and the time on
// each slide. Nothing is recorded or kept; the words stay in this browser.

/** Filler words the coach counts (Japanese and English), longest first so "えーと" is not also counted as "えー". */
export const FILLERS = ["えーっと", "えーと", "ええと", "えっと", "えー", "あのー", "あの", "そのー", "うーん", "まあ", "なんか", "um", "uh", "you know"];
// Characters a minute (Japanese speech): below SLOW feels slow, above FAST is hard to follow.
export const PACE = { slow: 250, fast: 420 };

const strip = (t) => String(t ?? "").replace(/<[^>]*>/g, " ").replace(/\*\*/g, "");
const chars = (t) => strip(t).replace(/[\s、。，．,.!?！？「」『』（）()・…ー-]/g, "").length;

/** Every filler in a text, counted (a filler inside a longer one is counted once, as the longer one). */
export function countFillers(text) {
  let rest = ` ${String(text || "").toLowerCase()} `;
  const counts = new Map();
  for (const word of FILLERS) {
    const re = /^[a-z ]+$/.test(word) ? new RegExp(`(?<![a-z])${word}(?![a-z])`, "g") : new RegExp(word, "g");
    const found = rest.match(re)?.length || 0;
    if (found) { counts.set(word, found); rest = rest.replace(re, " "); }
  }
  return counts;
}

/**
 * Words (or short phrases) said twice in a row with a pause between: "その、その" or "the the". (Without a pause it
 * is often a word of its own: いろいろ, ますます.)
 */
export function repeats(text) {
  const found = new Map();
  const re = /(?<![\p{L}\p{N}])([\p{L}\p{N}]{2,8})[\s、,]+\1/gu;
  for (const m of String(text || "").matchAll(re)) {
    const word = m[1];
    if (FILLERS.includes(word.toLowerCase())) continue;
    found.set(word, (found.get(word) || 0) + 1);
  }
  return found;
}

/** How much of what was said on a slide is its own text read out (0–1): shared 4-character pieces. */
export function readingShare(spoken, slideText) {
  const said = strip(spoken).replace(/\s+/g, "");
  const shown = strip(slideText).replace(/\s+/g, "");
  if (said.length < 12 || shown.length < 12) return 0;
  const pieces = new Set();
  for (let i = 0; i + 4 <= shown.length; i += 1) pieces.add(shown.slice(i, i + 4));
  let hit = 0;
  let all = 0;
  for (let i = 0; i + 4 <= said.length; i += 2) { all += 1; if (pieces.has(said.slice(i, i + 4))) hit += 1; }
  return all ? hit / all : 0;
}

/** The words a slide shows (titles, body, objects' text), for readingShare. */
export function slideWords(slide, E) {
  const parts = [];
  const walk = (value, key = "") => {
    if (["notes", "sid", "type", "elements", "timeline", "comments", "readingOrder"].includes(key)) return;
    if (typeof value === "string") { if (!/^(#|data:|idb:|https?:|[a-z]+$)/i.test(value)) parts.push(value); }
    else if (Array.isArray(value)) value.forEach((v) => walk(v));
    else if (value && typeof value === "object") Object.entries(value).forEach(([k, v]) => walk(v, k));
  };
  walk(slide);
  for (const o of slide?.elements || []) { const t = E?.objectText?.(o) ?? o.text; if (t) parts.push(strip(t)); }
  return parts.join(" ");
}

/**
 * The report for one rehearsal. `said` is [{ text, slide, at }] (final results, `at` in ms from the start); `times`
 * maps a slide index to the ms spent on it; `deck` gives the slides' own words.
 */
export function coachReport({ said = [], times = new Map(), deck = null, E = null } = {}) {
  const all = said.map((s) => s.text).join(" ");
  const total = [...times.values()].reduce((a, b) => a + b, 0);
  const n = chars(all);
  const cpm = total >= 5000 ? Math.round(n / (total / 60000)) : null;
  const pace = cpm == null ? "unknown" : cpm < PACE.slow ? "slow" : cpm > PACE.fast ? "fast" : "good";
  const fillers = [...countFillers(all)].map(([word, count]) => ({ word, count })).sort((a, b) => b.count - a.count);
  const twice = [...repeats(all)].map(([word, count]) => ({ word, count })).sort((a, b) => b.count - a.count);
  const slides = [...times.keys()].sort((a, b) => a - b).map((i) => {
    const text = said.filter((s) => s.slide === i).map((s) => s.text).join(" ");
    const ms = times.get(i) || 0;
    const share = deck?.slides[i] ? readingShare(text, slideWords(deck.slides[i], E)) : 0;
    return { slide: i, ms, chars: chars(text), cpm: ms >= 5000 ? Math.round(chars(text) / (ms / 60000)) : null, reading: share >= 0.55 };
  });
  const tips = [];
  if (pace === "fast") tips.push("話すのが速めです。区切りで一呼吸おくと、聞き手がついてきやすくなります。");
  if (pace === "slow") tips.push("ややゆっくりです。要点はそのままに、つなぎの言葉を減らすとテンポが上がります。");
  const fillerCount = fillers.reduce((a, f) => a + f.count, 0);
  if (fillerCount >= 3) tips.push(`「${fillers[0].word}」などのつなぎ言葉が${fillerCount}回ありました。言葉に詰まったら、黙って間をとっても大丈夫です。`);
  if (twice.length) tips.push(`同じ言葉の繰り返し（「${twice[0].word}」など）がありました。`);
  const reading = slides.filter((s) => s.reading);
  if (reading.length) tips.push(`${reading.map((s) => s.slide + 1).join("・")}枚目は、スライドの文字をそのまま読み上げていました。見れば分かる文字は読まず、補足や理由を話しましょう。`);
  if (!tips.length && n) tips.push("よいペースで、つなぎ言葉も少なく話せています。");
  return { chars: n, ms: total, cpm, pace, fillers, fillerCount, repeats: twice, slides, tips, words: all };
}

const PACE_LABEL = { slow: "ゆっくり", good: "ちょうどよい", fast: "速い", unknown: "—" };
const fmt = (ms) => { const sec = Math.round(ms / 1000); return `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, "0")}`; };

export function createCoach(app) {
  const { h } = app;

  /** コーチによるリハーサル: from the first slide, listening; the report comes when the show ends. */
  async function start() {
    const Recognition = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!Recognition) { app.toast("このブラウザでは音声認識を使えません（Chrome・Edgeで使えます）"); return; }
    if (!app.deck() || app.player()) return;
    const said = [];
    const times = new Map();
    let at = 0;
    let since = Date.now();
    const started = since;
    let interim = "";
    let stopped = false;
    let timer = 0;
    const rec = new Recognition();
    rec.lang = app.showSettings()?.captionLang || "ja-JP";
    rec.continuous = true;
    rec.interimResults = true;
    const meter = h("div", { class: "coach-meter", role: "status" });
    const tick = () => {
      const player = app.player();
      if (!player) return;
      const now = Date.now();
      if (player.index !== at) { times.set(at, (times.get(at) || 0) + now - since); at = player.index; since = now; }
      const live = coachReport({ said, times: new Map([...times, [at, (times.get(at) || 0) + now - since]]) });
      meter.replaceChildren(
        h("b", {}, "コーチ"),
        h("span", { class: ["coach-pace", live.pace] }, `ペース：${PACE_LABEL[live.pace]}${live.cpm ? `（${live.cpm}字/分）` : ""}`),
        h("span", {}, `つなぎ言葉 ${live.fillerCount}`),
        h("span", { class: "coach-heard" }, interim || said.at(-1)?.text || "（話し始めてください）"));
    };
    rec.onresult = (event) => {
      interim = "";
      for (let i = event.resultIndex; i < event.results.length; i += 1) {
        const text = event.results[i][0].transcript;
        if (event.results[i].isFinal) said.push({ text, slide: app.player()?.index ?? at, at: Date.now() - started });
        else interim += text;
      }
      tick();
    };
    // Recognition stops by itself after a silence: listen again until the show ends.
    rec.onend = () => { if (!stopped) { try { rec.start(); } catch { /* starting */ } } };
    rec.onerror = (event) => { if (event.error === "not-allowed") { stopped = true; app.toast("マイクを使えません（ブラウザでマイクを許可してください）"); } };
    app.onShowClosed(() => {
      stopped = true;
      clearInterval(timer);
      try { rec.stop(); } catch { /* stopped */ }
      times.set(at, (times.get(at) || 0) + Date.now() - since);
      setTimeout(() => report(coachReport({ said, times, deck: app.deck(), E: app.E })), 0);
    });
    await app.presentFrom(0);
    if (!app.player()) { stopped = true; return; }
    app.presenterHost().append(meter);
    at = app.player().index;
    since = Date.now();
    try { rec.start(); } catch { /* already */ }
    timer = setInterval(tick, 400);
    tick();
  }

  /** The report at the end (PowerPoint's リハーサル レポート). */
  function report(r) {
    const row = (label, value, extra = null) => h("div", { class: "coach-row" }, h("span", {}, label), h("b", {}, value), extra);
    const dialog = h("dialog", { class: "coach-dialog", "aria-label": "リハーサル レポート" },
      h("div", { class: "dialog-head" }, h("h3", {}, "リハーサル レポート"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body coach-body" },
        h("div", { class: "coach-summary" },
          row("時間", fmt(r.ms)),
          row("ペース", r.cpm ? `${r.cpm} 字/分` : "—", h("small", { class: ["coach-pace", r.pace] }, PACE_LABEL[r.pace])),
          row("つなぎ言葉", `${r.fillerCount} 回`, h("small", {}, r.fillers.slice(0, 4).map((f) => `${f.word}×${f.count}`).join("　"))),
          row("繰り返し", `${r.repeats.reduce((a, x) => a + x.count, 0)} 回`, h("small", {}, r.repeats.slice(0, 3).map((x) => x.word).join("・")))),
        h("ul", { class: "coach-tips" }, r.tips.map((t) => h("li", {}, t))),
        h("table", { class: "coach-slides" },
          h("thead", {}, h("tr", {}, h("th", {}, "スライド"), h("th", {}, "時間"), h("th", {}, "字/分"), h("th", {}, "読み上げ"))),
          h("tbody", {}, r.slides.map((s) => h("tr", {}, h("td", {}, `${s.slide + 1}`), h("td", {}, fmt(s.ms)), h("td", {}, s.cpm ?? "—"), h("td", {}, s.reading ? "スライドの文字を読んでいます" : ""))))),
        h("p", { class: "hint" }, "話した言葉はこのブラウザの中だけで使い、保存しません。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-primary", onclick: () => dialog.close() }, "閉じる")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
    return dialog;
  }

  return { start, report };
}
