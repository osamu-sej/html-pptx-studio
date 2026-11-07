import { buildSearchIndex, deleteSavedDeck, getSavedDeck, listSavedDecks, putSavedDeck, searchSavedDecks, slideExcerpt } from "./saved-library.js";
import { hasSlidePicture, picturePlan } from "./auto-images.mjs?v=__APP_VERSION__";
import { LOOK_ADVICE, LOOKS, lookOf, varietyIssues } from "./layout-looks.mjs?v=__APP_VERSION__";
import { createCanvas } from "./editor/canvas.mjs?v=__APP_VERSION__";
import { createEditorUi } from "./editor/ui.mjs?v=__APP_VERSION__";
import * as ops from "./editor/ops.mjs?v=__APP_VERSION__";
import { createConverter } from "./editor/convert.mjs?v=__APP_VERSION__";
import { resetDeckActions } from "./reset-actions.mjs?v=__APP_VERSION__";
import { chromeOf, describeAdded, enhanceSlide } from "./editor/htmlfx.mjs?v=__APP_VERSION__";
import { createShell } from "./editor/window.mjs?v=__APP_VERSION__";
import { createComments } from "./editor/comments.mjs?v=__APP_VERSION__";
import { userName } from "./editor/people.mjs?v=__APP_VERSION__";
import { createCoedit } from "./editor/coedit.mjs?v=__APP_VERSION__";
import { newSid, uniqueSids } from "./editor/coedit-merge.mjs?v=__APP_VERSION__";
import { inkObject } from "./editor/ink.mjs?v=__APP_VERSION__";
import { createShowTools, customShowsOf, keptInk, playerOptions, showOf, showSlides } from "./editor/show.mjs?v=__APP_VERSION__";
import { createA11y } from "./editor/a11y.mjs?v=__APP_VERSION__";
import { createPrinter } from "./editor/print.mjs?v=__APP_VERSION__";
import { matcher as findMatcher, snippet as findSnippet } from "./editor/find.mjs?v=__APP_VERSION__";
import { createSynonyms } from "./editor/synonyms.mjs?v=__APP_VERSION__";
import { createProofing } from "./editor/proof.mjs?v=__APP_VERSION__";
import { createFileInfo, infoOf, isFinal, FINAL_STATUS, strip as stripDeckData } from "./editor/fileinfo.mjs?v=__APP_VERSION__";
import { createSlideTools } from "./editor/slidetools.mjs?v=__APP_VERSION__";
import { createVideoExport } from "./editor/video.mjs?v=__APP_VERSION__";
import { createCompare } from "./editor/compare.mjs?v=__APP_VERSION__";
import { createOnline } from "./editor/online.mjs?v=__APP_VERSION__";
import { createCoach } from "./editor/coach.mjs?v=__APP_VERSION__";
import { createImageExport } from "./editor/imagexport.mjs?v=__APP_VERSION__";
import { createPictureSaver } from "./editor/picsave.mjs?v=__APP_VERSION__";
import { createGifExport } from "./editor/gif.mjs?v=__APP_VERSION__";
import { lockHtml } from "./editor/protect.mjs?v=__APP_VERSION__";
import { createBackground } from "./editor/background.mjs?v=__APP_VERSION__";
import { createAutoCorrect } from "./editor/autocorrect.mjs?v=__APP_VERSION__";
import { applyFilter as applySheetFilter, recalc as sheetRecalc } from "./editor/sheet.mjs?v=__APP_VERSION__";
import { createPrintMasters, masterSlides, normalizeMasters } from "./editor/printmaster.mjs?v=__APP_VERSION__";

/*
 * HTML PPTX Studio — the editor.
 * Brief → outline → slides (Codex), then edit by chat or by hand, present with motion, export one HTML file.
 * Slides are drawn by SlideEngine (engine/engine.js) and played by engine/motion.js.
 */
const $ = (id) => document.getElementById(id);
const E = window.SlideEngine;
// Excel スプレッドシート: a table with `sheet` has its formulas worked out whenever the engine normalizes it.
E.sheetCalc = (table) => { sheetRecalc(table, { inPlace: true }); applySheetFilter(table); };
E.lottieUrl = "/vendor/lottie.js";
const APP_VERSION = "__APP_VERSION__";
const STORAGE = {
  current: "hpptx-studio-current-v1",
  history: "hpptx-studio-history-v1",
  chat: "hpptx-studio-chat-v1",
  panel: "hpptx-studio-panel",
};

// ---------------------------------------------------------------- catalog

const TYPE_DESC = {
  title: "資料名・日付・対象者", section: "章の区切り", closing: "次のアクション", hero: "写真いっぱいにひと言", statement: "大きな一文で言い切る",
  content: "要点を順に説明", agenda: "本日の流れ", executiveSummary: "結論・根拠・依頼", kpi: "主要な数値を大きく", dashboard: "数値＋グラフ",
  imageText: "グラフや写真＋説明", statsCompare: "前後の数値を並べる", compare: "2つの案・状態を比較", beforeAfter: "変化を示す", table: "複数項目を整理",
  process: "横並びの工程", processList: "縦並びの手順", flowChart: "1〜2本の流れ", timeline: "時期と出来事", roadmap: "段階的な展開",
  gantt: "期間とタスク", waterfall: "合計の内訳・増減", logicTree: "原因・論点の分解", cards: "並列の要点", headerCards: "見出し付きの要点",
  bulletCards: "説明付きの要点", headerTwoColumn: "見出し付き2列", headerThreeSummary: "3論点から結論へ", grid2x2: "4つの論点", matrix: "2軸で整理",
  swot: "強み・弱み・機会・脅威", diagram: "役割ごとの流れ", cycle: "循環する取り組み", pyramid: "階層構造", funnel: "段階的な絞り込み",
  stepUp: "段階的な成長", triangle: "3要素の関係", venn: "重なり", orgChart: "体制・役割分担", checklist: "確認項目", faq: "想定問答", quote: "声・メッセージ",
  simulator: "条件を動かすと結果を計算し直す", gap: "打ち手で目標との差を埋める", blank: "図形・文字・画像を自由に置く",
};
const TYPE_INFO = Object.fromEntries(Object.keys(TYPE_DESC).map((type) => [type, [E.TYPE_LABELS[type] ?? type, TYPE_DESC[type]]]));
const SLIDE_TYPES = Object.keys(TYPE_INFO);
const typeLabel = (type) => TYPE_INFO[type]?.[0] ?? type;
const TITLED = (type) => !["title", "section", "closing"].includes(type);
const THEME_IDS = new Set(E.THEMES.map((theme) => theme.id));
const themeMeta = (id) => E.THEMES.find((theme) => theme.id === id) ?? E.THEMES[0];
// A deck's design is one of the engine's themes (each is also a PowerPoint theme: colours and fonts). A new deck
// takes the theme last chosen in this browser.
const THEME_KEY = "hpptx-theme";
function preferredTheme() {
  try { const id = localStorage.getItem(THEME_KEY); return THEME_IDS.has(id) ? id : E.DEFAULT_THEME; } catch { return E.DEFAULT_THEME; }
}
function rememberTheme(id) {
  try { localStorage.setItem(THEME_KEY, id); } catch { /* private window */ }
  if (document.getElementById("createThemes")) renderCreateThemes();
}
/** The create view's themes: the one a new deck takes (each chip drawn from the theme's colours). */
function renderCreateThemes() {
  const box = document.getElementById("createThemes");
  if (!box) return;
  const current = preferredTheme();
  box.replaceChildren(...E.THEMES.map((theme) => {
    const [bg, ink, accent, accent2] = theme.swatch;
    return h("button", { type: "button", class: "theme-chip", role: "radio", "aria-checked": String(current === theme.id), "data-theme-id": theme.id, title: theme.desc, onclick: () => rememberTheme(theme.id) },
      h("span", { class: "art", style: { background: bg } },
        h("i", { class: "bar", style: { background: ink } }), h("i", { class: "sub", style: { background: ink } }),
        h("i", { class: "dot", style: { background: accent, "box-shadow": `-14px -10px 0 -4px ${accent2}` } })),
      h("b", {}, theme.name));
  }));
}
/** The deck the studio works on: a theme the engine knows, and the editor's colours are that theme's. */
function adoptDeck(deck) {
  withTheme(deck);
  if (deck) E.useTheme(deck.theme);
  return deck;
}
/** A deck with a theme the engine knows (in place). */
function withTheme(deck) {
  if (!deck) return deck;
  if (!THEME_IDS.has(deck.theme)) deck.theme = E.DEFAULT_THEME;
  return deck;
}
// Layouts that give a photo its own place; anywhere else a photo or video floats where you put it.
const SLOTTED = new Set(["title", "section", "closing", "hero", "statement", "content", "quote", "imageText"]);
const ICON_TYPES = new Set(["cards", "headerCards", "bulletCards", "triangle", "orgChart", "grid2x2", "headerTwoColumn", "headerThreeSummary"]);
const DEFAULT_PLACEMENT = { x: 0.6, y: 0.3, w: 0.32, h: 0.46 };
const DEFAULT_MOTION = { entrance: "rise", hover: "lift", numbers: true, ambient: true, kinetic: "mask", backdrop: "none", emphasis: "marker", draw: true };
// Every kind of motion comes from the engine's catalogs (public/engine/engine.js), so the studio, the AI and
// the player always agree on what exists.
const TRANSITIONS = Object.keys(E.TRANSITIONS);
// Motion graphics: the big lines' motion (kinetic) and the moving graphic behind a slide (backdrop).
const KINETIC_INFO = { auto: "おまかせ（資料の設定）", none: "動かさない", ...E.KINETIC };
const BACKDROP_INFO = { auto: "おまかせ（資料の設定）", none: "なし", ...E.BACKDROPS };
// What one slide can choose for itself ("auto" follows the deck).
const ENTRANCE_INFO = { auto: "おまかせ（資料の設定）", ...E.ENTRANCES, none: "動かさない" };
const EMPHASIS_INFO = { auto: "おまかせ（資料の設定）", ...E.EMPHASES };
const SLIDE_TRANSITION_INFO = { auto: "おまかせ（資料の設定）", ...E.TRANSITIONS };
const BUILD_INFO = {
  auto: ["おまかせ", "レイアウトに合わせて自動で選びます"],
  none: ["なし", "最初からすべて表示します"],
  fade: ["まとめて", "スライドが出るとき、中身がまとめて現れます"],
  cascade: ["順番に", "項目が少しずつ時間差で現れます"],
  click: ["クリックで", "クリック（→キー）のたびに1項目ずつ出ます。説明しながら見せたいときに"],
  spotlight: ["順に注目", "全部見せたまま、クリックのたびに1項目を強調し、ほかを薄くします。一覧を見せながら1つずつ話すときに"],
};
const PHOTO_MOTIONS = [["none", "動かさない"], ...Object.entries(E.PHOTO_MOTIONS)];
const DENSITY_LABEL = { light: "少なめ", standard: "標準", rich: "多め" };

const SAMPLES = [
  {
    label: "生成AIの全社展開（役員報告）",
    audience: "役員", purpose: "意思決定・報告", slideCount: 10, density: "standard", tone: "executive",
    brief: "生成AI活用の試行結果を役員に報告し、全社展開の予算承認をもらいたい。5部門120人で3か月試行し、月1,440時間の作業を削減（資料作成620時間、議事録380時間、データ集計290時間、その他150時間）。利用者満足度は4.3点（5点満点）。一方で、情報漏えいへの不安と、使いこなしの個人差が課題。対策として社内ガイドラインと入力制限、部門ごとの推進役を置く。展開は10月から3段階（本社→支社→全拠点）。初年度予算2,400万円、年間の削減効果は約1億円を見込む。",
  },
  {
    label: "新入社員向け 生成AI活用研修",
    audience: "新入社員", purpose: "研修・勉強会", slideCount: 8, density: "standard", tone: "friendly",
    brief: "新入社員向けの生成AI活用研修。目的は日常業務（議事録要約、メール下書き、データ整理）で安全に生成AIを使えるようにすること。良いプロンプトの5要素：役割、背景、指示、制約、出力形式。社外秘情報や個人情報は入力しない。回答は必ず人が事実確認する。演習：会議メモから要約と次アクションを作る。最後に明日から試す3つの行動を宣言する。",
  },
  {
    label: "新サービスのキックオフ（チーム共有）",
    audience: "チームメンバー", purpose: "情報共有・周知", slideCount: 9, density: "standard", tone: "friendly",
    brief: "来春リリースする会員アプリ新機能「おすすめ通知」のキックオフ。狙いは再来店率を3か月で5ポイント上げること。対象は直近90日に2回以上来店した会員約40万人。機能は購買履歴から好みを推定し、週1回まで通知する。開発は4フェーズ（要件定義11月、開発12〜2月、テスト3月、リリース4月）。体制は企画2名、開発6名、デザイン2名、データ分析2名。リスクは通知の送りすぎによる解除増加で、送信上限とABテストで抑える。",
  },
];

// ---------------------------------------------------------------- field specs (the "編集" form)

const T = (key, label, o = {}) => ({ key, label, kind: "text", ...o });
const A = (key, label, o = {}) => ({ key, label, kind: "area", ...o });
const L = (key, label, o = {}) => ({ key, label, kind: "list", ...o });
const S = (key, label, options, o = {}) => ({ key, label, kind: "select", options, ...o });
const C = (key, label) => ({ key, label, kind: "check" });
const N = (key, label, o = {}) => ({ key, label, kind: "number", ...o });
const G = (key, label, fields, o = {}) => ({ key, label, kind: "group", fields, ...o });
const I = (key, label) => ({ key, label, kind: "icon" });
const CARD_FIELDS = (descReq = false, icon = true) => [T("title", "見出し", { req: true, max: 20 }), A("desc", "説明", { max: 50, req: descReq, rows: 2 }), icon ? I("icon", "アイコン") : null].filter(Boolean);
const CARDS = (label, min, max, o = {}) => G("items", label, CARD_FIELDS(o.descReq, o.icon !== false), { min, max, newItem: () => ({ title: "新しい項目", desc: "説明を入力" }), ...o });
const STATES = [["done", "完了"], ["next", "次に実施"], ["todo", "予定"]];
const TRENDS = [["neutral", "—"], ["up", "↑ 上昇"], ["down", "↓ 低下"]];
const STATUS = [["neutral", "中立"], ["good", "良い"], ["bad", "要注意"]];

const SPEC = {
  title: [T("title", "資料タイトル", { req: true, max: 30 }), T("subtitle", "サブタイトル", { max: 50 }), T("date", "日付", { max: 20, placeholder: "例：2026年9月" })],
  section: [T("title", "章タイトル", { req: true, max: 24 }), A("takeaway", "補足（任意）", { max: 50 })],
  closing: [T("title", "見出し", { max: 24, placeholder: "次のアクション" }), A("message", "次のアクション（誰が・いつまでに・何を。改行で2〜3項目に分けると番号付きで並びます）", { max: 90, rows: 3 })],
  hero: [],
  statement: [A("text", "大きく見せる一文（**語句** で強調）", { req: true, max: 60, rows: 3 })],
  content: [L("points", "箇条書き", { max: 10, maxChars: 50, hint: "1行に1項目。「要素名：説明」の形で2〜5項目にすると番号付きで大きく表示されます。" }), C("twoColumn", "2列で表示")],
  agenda: [L("items", "項目", { min: 1, max: 10, hint: "1行に1項目。「章名：概要」で2段表示。" })],
  compare: [T("leftTitle", "左の見出し", { max: 16 }), L("leftItems", "左の項目", { max: 8, maxChars: 45 }), T("rightTitle", "右の見出し", { max: 16 }), L("rightItems", "右の項目", { max: 8, maxChars: 45 })],
  beforeAfter: [T("leftTitle", "Beforeの見出し", { max: 16 }), L("leftItems", "Beforeの項目", { max: 7, maxChars: 45 }), T("rightTitle", "Afterの見出し", { max: 16 }), L("rightItems", "Afterの項目", { max: 7, maxChars: 45 })],
  process: [L("steps", "工程", { min: 2, max: 6, hint: "1行に1工程。「工程名：説明」で説明付きになります。" })],
  processList: [L("steps", "工程", { min: 2, max: 8, hint: "1行に1工程。「工程名：説明」で2段表示。" })],
  flowChart: [G("flows", "フロー", [L("steps", "工程", { min: 2, max: 5 })], { min: 1, max: 2, newItem: () => ({ steps: ["開始", "処理", "完了"] }) })],
  timeline: [G("milestones", "マイルストーン", [T("date", "時期", { max: 12 }), T("label", "内容", { req: true, max: 30 }), S("state", "状態", STATES)], { min: 2, max: 8, newItem: () => ({ date: "時期", label: "出来事", state: "todo" }) })],
  diagram: [G("lanes", "レーン", [T("title", "レーン名", { req: true, max: 10 }), L("items", "項目", { max: 6 })], { min: 2, max: 5, newItem: () => ({ title: "担当", items: ["作業"] }) })],
  cycle: [G("items", "要素", [T("label", "名前", { req: true, max: 10 }), T("subLabel", "補足", { max: 24 })], { min: 3, max: 6, newItem: () => ({ label: "要素" }) }), T("centerText", "中央の言葉", { max: 10 })],
  cards: [S("columns", "列数", [["", "自動"], ["2", "2列"], ["3", "3列"]], { number: true }), CARDS("カード", 1, 6)],
  headerCards: [S("columns", "列数", [["", "自動"], ["2", "2列"], ["3", "3列"]], { number: true }), CARDS("カード", 1, 6)],
  bulletCards: [CARDS("カード", 1, 4, { descReq: true })],
  grid2x2: [CARDS("項目", 1, 4)],
  headerTwoColumn: [CARDS("列", 2, 2)],
  headerThreeSummary: [CARDS("列", 3, 3), A("summary", "まとめ", { max: 50 })],
  triangle: [CARDS("要素", 3, 3)],
  venn: [CARDS("円", 2, 3, { icon: false })],
  stepUp: [CARDS("段階", 2, 5, { descReq: true, icon: false })],
  roadmap: [CARDS("フェーズ", 2, 6, { icon: false })],
  orgChart: [T("root", "上位の組織・役割", { max: 20 }), CARDS("部署・役割", 2, 5)],
  checklist: [G("items", "項目", [T("title", "項目", { req: true, max: 18 }), A("desc", "説明", { max: 45, rows: 2 }), C("done", "完了済み")], { min: 1, max: 7, newItem: () => ({ title: "確認項目", desc: "" }) })],
  matrix: [T("yLabel", "縦軸の名前", { max: 12 }), T("xLabel", "横軸の名前", { max: 12 }), CARDS("象限（左上→右上→左下→右下）", 4, 4, { icon: false })],
  swot: [CARDS("象限（強み→弱み→機会→脅威）", 4, 4, { icon: false })],
  gantt: [L("periods", "期間（1行に1つ）", { max: 8 }), G("items", "タスク", [T("title", "タスク", { req: true, max: 14 }), T("desc", "バーの文言", { max: 14 }), N("start", "開始（0始まり）", { min: 0, max: 7 }), N("span", "期間数", { min: 1, max: 8 })], { min: 2, max: 7, newItem: () => ({ title: "タスク", desc: "", start: 0, span: 1 }) }),
    N("now", "「いま」の線（期間の位置。例：1.5＝2つ目の期間の真ん中。空欄なら出さない）", { free: true, optional: true })],
  simulator: [
    G("inputs", "条件（発表中はスライダーで動かせます。式では上から a・b・c）", [T("label", "条件の名前", { req: true, max: 14 }), N("value", "いまの値", { free: true }), N("min", "最小", { free: true }), N("max", "最大", { free: true }), N("step", "刻み（空欄なら自動）", { free: true, optional: true }), T("unit", "単位", { max: 8 })], { min: 1, max: 3, newItem: () => ({ label: "条件", value: 10, min: 0, max: 20, unit: "" }) }),
    T("formula", "計算式（a・b・c と ＋ − × ÷ ( ) と数字。例：a × b × c ÷ 100）", { req: true, max: 60 }),
    T("resultLabel", "結果の名前", { max: 16, placeholder: "例：月の削減時間" }), T("resultUnit", "結果の単位", { max: 8 }),
    S("digits", "小数点以下", [["", "自動"], ["0", "0桁"], ["1", "1桁"], ["2", "2桁"]], { number: true }),
    T("compareLabel", "比べる値の名前（目標・現状など。任意）", { max: 12 }), N("compareValue", "比べる値（空欄なら出さない）", { free: true, optional: true }),
  ],
  gap: [
    T("unit", "単位", { max: 8, placeholder: "例：時間・億円" }),
    T("targetLabel", "目標の名前", { max: 12 }), N("target", "目標の値", { free: true }),
    T("currentLabel", "現状の名前", { max: 12 }), N("current", "現状の値", { free: true }),
    G("measures", "打ち手（発表中はクリックのたびに1つずつオン。クリックでオン・オフも）", [T("title", "打ち手", { req: true, max: 16 }), N("value", "上乗せする量", { free: true }), A("desc", "説明", { max: 30, rows: 2 })], { min: 1, max: 5, newItem: () => ({ title: "打ち手", value: 10, desc: "" }) }),
  ],
  waterfall: [T("unit", "単位", { max: 6, placeholder: "例：時間・億円" }), G("items", "棒（左から）", [T("label", "名前", { req: true, max: 8 }), N("value", "値（減少はマイナス）", { free: true }), C("total", "合計の棒（0から立てる）")], { min: 2, max: 8, newItem: () => ({ label: "要因", value: 10 }) })],
  logicTree: [T("root", "分解する課題・論点", { req: true, max: 20 }), G("branches", "枝", [T("title", "枝（要因・論点）", { req: true, max: 14 }), L("items", "葉（1行に1つ）", { max: 3, maxChars: 18 }), C("highlight", "強調する（手を打つ枝）")], { min: 2, max: 4, newItem: () => ({ title: "要因", items: ["具体的な事象"] }) })],
  table: [{ kind: "table", key: "table", label: "表" }],
  quote: [A("text", "引用文", { req: true, max: 90, rows: 3 }), T("author", "出典・発言者", { max: 30 })],
  kpi: [G("items", "指標", [T("value", "数値", { req: true, max: 10 }), T("label", "指標名", { req: true, max: 14 }), T("change", "補足", { max: 24 }), S("status", "評価", STATUS)], { min: 1, max: 4, newItem: () => ({ value: "—", label: "指標", change: "", status: "neutral" }) })],
  dashboard: [G("items", "指標", [T("value", "数値", { req: true, max: 10 }), T("label", "指標名", { req: true, max: 14 }), T("change", "補足", { max: 24 })], { min: 2, max: 4, newItem: () => ({ value: "—", label: "指標", change: "" }) }), { kind: "chart", key: "image", label: "グラフ" }],
  statsCompare: [T("leftTitle", "左列の見出し", { max: 12 }), T("rightTitle", "右列の見出し", { max: 12 }), G("stats", "指標", [T("label", "指標", { req: true, max: 14 }), T("leftValue", "左の値", { max: 12 }), T("rightValue", "右の値", { max: 12 }), S("trend", "傾向", TRENDS)], { min: 1, max: 6, newItem: () => ({ label: "指標", leftValue: "—", rightValue: "—", trend: "neutral" }) })],
  faq: [G("items", "質問", [T("q", "質問", { req: true, max: 30 }), A("a", "回答", { req: true, max: 60, rows: 2 })], { min: 1, max: 4, newItem: () => ({ q: "質問", a: "回答" }) })],
  pyramid: [G("levels", "段（上から）", [T("title", "名前", { req: true, max: 12 }), A("description", "説明", { max: 40, rows: 2 })], { min: 3, max: 5, newItem: () => ({ title: "段", description: "" }) })],
  funnel: [G("levels", "段（上から）", [T("title", "名前", { req: true, max: 12 }), A("description", "説明", { max: 40, rows: 2 })], { min: 3, max: 5, newItem: () => ({ title: "段", description: "" }) })],
  imageText: [L("points", "説明（箇条書き）", { max: 8, maxChars: 45 }), { kind: "chart", key: "image", label: "グラフ・画像", allowPhoto: true }, S("imagePosition", "図の位置", [["left", "左"], ["right", "右"]]), T("imageCaption", "キャプション", { max: 30 })],
  executiveSummary: [A("conclusion", "結論", { max: 50, rows: 2 }), CARDS("根拠・要点", 2, 3), A("action", "次のアクション", { max: 50, rows: 2 })],
};
const COMMON_TOP = [
  T("title", "タイトル（論点）", { req: true, max: 30 }),
  A("takeaway", "キーメッセージ（結論を一文で。**語句** で囲むとその語句だけ強調）", { max: 60, rows: 2 }),
  T("subhead", "小見出し（任意。空欄なら章の名前を表示）", { max: 24 }),
  T("source", "出所（任意。ページ下とグラフの吹き出しに出ます）", { max: 60, placeholder: "例：総務省「労働力調査」2025年" }),
];

function specFor(type) {
  // 白紙: the title is optional; everything else on the page is objects placed by hand.
  if (type === "blank") return [T("title", "タイトル（任意。空欄なら出しません）", { max: 30 }), ...COMMON_TOP.slice(1)];
  return TITLED(type) ? [...COMMON_TOP, ...(SPEC[type] ?? [])] : SPEC[type] ?? [];
}

// ---------------------------------------------------------------- state

const state = {
  deck: null,
  selected: 0,
  view: "single",
  mode: "create",
  undo: [],
  redo: [],
  editBurst: null,
  aiBusy: false,
  codexAuthorized: false,
  codexImageAuthorized: false,
  historyId: null,
  savedDeckId: null,
  chat: { messages: [], busy: false, progress: "", attachment: null },
  outline: null,
  createBusy: false,
  panel: "chat",
  imported: null,
  player: null,
  motionPreview: null,
  inline: null,
};
let autoImageRun = null;

// ---------------------------------------------------------------- small helpers

const h = E.h;
const clone = (value) => JSON.parse(JSON.stringify(value));
const keepEmphasis = (value = "") => String(value ?? "").replace(/\[\[([^\]]+)\]\]/g, "**$1**").trim();
const strip = (value = "") => String(value ?? "").replace(/\*\*([^*]+)\*\*/g, "$1").replace(/\[\[([^\]]+)\]\]/g, "$1").trim();
const charCount = (value = "") => [...String(value)].reduce((sum, ch) => sum + (/[\x20-\x7e]/.test(ch) ? 0.5 : 1), 0);
const squash = (text) => strip(String(text ?? "")).replace(/\s+/g, "");
const esc = (text) => String(text ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);
const safeJson = (value) => JSON.stringify(value).replace(/</g, "\\u003c").replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029");
const fileSafe = (text, fallback = "資料") => String(text ?? "").replace(/[\\/:*?"<>|]/g, "_").trim().slice(0, 60) || fallback;
const today = () => { const d = new Date(); return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}`; };
const wait = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

/** A small inline SVG icon from the engine's icon set, for the studio chrome. */
function uiIcon(name, cls = "ui-icon") {
  const entry = E.icons[name];
  if (!entry) return null;
  const el = E.s("svg", { class: cls, viewBox: "0 0 24 24", "aria-hidden": "true" });
  el.innerHTML = entry.svg;
  return el;
}

let toastTimer = null;
function toast(message) {
  const el = $("toast");
  el.textContent = message;
  el.classList.add("show");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove("show"), 3200);
}

function showStatus(id, text, type = "info") {
  const el = $(id);
  el.textContent = text;
  el.className = `status show ${type}`;
}

function clearStatus(id) {
  $(id).className = "status";
}

function setPill(text, kind = "") {
  const pill = $("hostStatus");
  pill.className = `pill ${kind}`;
  pill.querySelector(".label").textContent = text;
  pill.title = text;
}

async function jsonFetch(url, options = {}) {
  const response = await fetch(url, {
    credentials: "same-origin",
    ...options,
    headers: { "content-type": "application/json", ...(options.headers || {}) },
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw Object.assign(new Error(payload.error || `HTTP ${response.status}`), { body: payload, status: response.status });
  return payload;
}

function getPath(object, path) {
  return path.reduce((value, key) => (value == null ? undefined : value[key]), object);
}

function setPath(object, path, value) {
  let target = object;
  for (let i = 0; i < path.length - 1; i += 1) {
    if (target[path[i]] == null) target[path[i]] = typeof path[i + 1] === "number" ? [] : {};
    target = target[path[i]];
  }
  const last = path.at(-1);
  if (value === undefined || value === "") delete target[last];
  else target[last] = value;
}

const pathKey = (path) => path.reduce((key, part) => (typeof part === "number" ? `${key}[${part}]` : key ? `${key}.${part}` : part), "");
const parseField = (field) => {
  const path = [];
  String(field).replace(/([^.[\]]+)|\[(\d+)\]/g, (_, key, index) => { path.push(index != null ? Number(index) : key); return ""; });
  return path;
};

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error);
    reader.readAsDataURL(blob);
  });
}

async function dataUrlToBlob(dataUrl) {
  return (await fetch(dataUrl)).blob();
}

// ---------------------------------------------------------------- Codex connection

let loginPollTimer = null;

function setCodexConnection(kind, title, detail) {
  const card = $("standaloneConnection");
  card.classList.remove("ready", "error", "hidden");
  if (kind) card.classList.add(kind);
  $("codexConnectionTitle").textContent = title;
  $("codexConnectionDetail").textContent = detail;
}

/** Back to the connection state after a job: Codex, or the stand-by AI when that is what answers. */
function restorePill() {
  const codexOff = state.localAi?.available && !state.codexImageAuthorized;
  setPill(codexOff ? "予備AIで動作中" : state.codexAuthorized ? "Codex接続済み" : "Codex未接続", codexOff ? "busy" : state.codexAuthorized ? "ok" : "");
}

/** The stand-by AI (Gemma on the machine running the server): shown next to the Codex status. */
function showLocalAi(fallback) {
  state.localAi = fallback || null;
  const pill = $("localAiStatus");
  if (!fallback) { pill.hidden = true; return; }
  pill.hidden = false;
  pill.className = `pill ${fallback.available ? "ok" : ""}`;
  pill.querySelector(".label").textContent = fallback.available ? `予備AI：${fallback.model}` : fallback.web ? "予備AI：なし（Web版）" : "予備AI：なし";
  pill.title = fallback.available
    ? `Codexが使えないとき（利用上限・接続切れなど）は、自動で ${fallback.model} に切り替えて作ります。品質と速度はCodexより落ちます。`
    : fallback.reason || "";
}

async function checkCodexStatus() {
  try {
    const status = await jsonFetch("/api/codex/status", { method: "GET" });
    showLocalAi(status.fallback);
    const standBy = Boolean(status.fallback?.available);
    state.codexAuthorized = Boolean(status.authorized || status.authenticated) || standBy;
    state.codexImageAuthorized = Boolean(status.authorized || status.authenticated);
    const needsPasscode = Boolean(status.team && !status.member);
    $("passcodePanel").classList.toggle("hidden", !needsPasscode);
    $("connectCodexBtn").textContent = status.pendingLogin ? "コードを再発行" : "Codexに接続";
    if (needsPasscode) {
      setCodexConnection("", "合言葉を入力してください", "このスタジオはチームで共有されています。管理者から聞いた合言葉を入れると、AIで構成を作れます。雛形から作る・JSON読込は合言葉なしで使えます。");
      setPill("未入室");
      $("connectCodexBtn").hidden = true;
    } else if (state.codexImageAuthorized) {
      setCodexConnection("ready", "Codexに接続済み", `${status.planType || status.accountType || "ChatGPTアカウント"}で構成を生成できます。${standBy ? `Codexが使えなくなったときは${status.fallback.model}に自動で切り替えます。` : ""}`);
      setPill("Codex接続済み", "ok");
      $("connectCodexBtn").hidden = true;
      $("deviceAuthPanel").classList.add("hidden");
      clearInterval(loginPollTimer);
      loginPollTimer = null;
    } else if (status.pendingLogin) {
      $("deviceCode").textContent = status.pendingLogin.userCode;
      $("deviceLink").href = /^https:\/\//.test(status.pendingLogin.verificationUrl || "") ? status.pendingLogin.verificationUrl : "https://chatgpt.com/";
      $("deviceAuthPanel").classList.remove("hidden");
      setCodexConnection("", "認証コードを入力してください", "認証ページでコードを入力すると、自動で接続されます。コードが使えなくなった場合は再発行できます。");
      setPill(standBy ? "予備AIで動作中" : "Codex認証待ち", standBy ? "busy" : "");
      $("connectCodexBtn").hidden = false;
      if (!loginPollTimer) loginPollTimer = setInterval(checkCodexStatus, 2500);
    } else if (status.loginError) {
      $("deviceAuthPanel").classList.add("hidden");
      setCodexConnection("error", "Codexへ接続できません", status.loginError);
      setPill(standBy ? "予備AIで動作中" : "Codex接続エラー", standBy ? "busy" : "err");
      $("connectCodexBtn").hidden = false;
      clearInterval(loginPollTimer);
      loginPollTimer = null;
    } else if (standBy) {
      $("deviceAuthPanel").classList.add("hidden");
      setCodexConnection("ready", `予備AI（${status.fallback.model}）で動作中`, `${status.available === false ? `Codexを起動できません（${status.error}）。` : "Codexに接続していません。"}手元の${status.fallback.model}で作ります（Codexより時間がかかり、仕上がりも簡素になります）。Codexに接続すると通常どおり使えます。`);
      setPill("予備AIで動作中", "busy");
      $("connectCodexBtn").hidden = status.available === false;
    } else {
      $("deviceAuthPanel").classList.add("hidden");
      setCodexConnection("", "Codexへの接続が必要です", `ChatGPTアカウントでデバイス認証します（APIキーは不要）。雛形から作る・JSON読込・サンプルは接続なしで使えます。${status.fallback?.web ? "（Web版では予備AIのGemmaは使えません）" : ""}`);
      setPill("Codex未接続");
      $("connectCodexBtn").hidden = false;
      clearInterval(loginPollTimer);
      loginPollTimer = null;
    }
  } catch (error) {
    state.codexAuthorized = false;
    state.codexImageAuthorized = false;
    showLocalAi(error.body?.fallback);
    $("deviceAuthPanel").classList.add("hidden");
    setCodexConnection("error", "Codexを起動できません", `${error.message}${error.body?.fallback?.web ? "（Web版では予備AIのGemmaは使えないため、AI機能は止まっています）" : ""}`);
    setPill("Codex接続エラー", "err");
    $("connectCodexBtn").hidden = false;
    clearInterval(loginPollTimer);
    loginPollTimer = null;
  }
  renderInspector();
  renderChat();
  updateTopbar();
  showImageCoverage();
}

async function submitPasscode(event) {
  event.preventDefault();
  const input = $("passcodeInput");
  if (!input.value) return input.focus();
  $("passcodeBtn").disabled = true;
  try {
    await jsonFetch("/api/auth/passcode", { method: "POST", body: JSON.stringify({ passcode: input.value }) });
    input.value = "";
    await checkCodexStatus();
    toast("入室しました");
  } catch (error) {
    setCodexConnection("error", "入室できません", error.message);
  } finally {
    $("passcodeBtn").disabled = false;
  }
}

async function startCodexLogin() {
  const button = $("connectCodexBtn");
  button.disabled = true;
  button.textContent = "準備中…";
  try {
    const login = await jsonFetch("/api/codex/login", { method: "POST", body: "{}" });
    if (login.alreadyConnected) return checkCodexStatus();
    await checkCodexStatus();
  } catch (error) {
    setCodexConnection("error", "Codexへ接続できません", error.message);
  } finally {
    button.disabled = false;
    button.textContent = $("deviceAuthPanel").classList.contains("hidden") ? "Codexに接続" : "コードを再発行";
  }
}

async function copyDeviceCode() {
  const code = $("deviceCode").textContent.trim();
  if (!code) return;
  try {
    if (!navigator.clipboard?.writeText) throw new Error("Clipboard API unavailable");
    await navigator.clipboard.writeText(code);
    toast("認証コードをコピーしました");
  } catch {
    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents($("deviceCode"));
    selection?.removeAllRanges();
    selection?.addRange(range);
    toast("コードを選択しました。⌘C または Ctrl+C でコピーしてください");
  }
}

// ---------------------------------------------------------------- normalization

const EMPHASIS_KEYS = new Set(["title", "takeaway", "conclusion", "text", "message", "summary", "action"]);

function normalizeMedia(media) {
  if (!media || typeof media !== "object" || typeof media.src !== "string" || !media.src.trim()) return undefined;
  const out = { src: media.src.trim() };
  if (["image", "video", "lottie"].includes(media.kind)) out.kind = media.kind;
  if (media.name) out.name = String(media.name).slice(0, 200);
  if (media.fit === "contain" || (media.kind === "lottie" && media.fit === "cover")) out.fit = media.fit;
  for (const key of ["autoplay", "loop", "muted"]) if (typeof media[key] === "boolean") out[key] = media[key];
  const p = media.placement;
  if (p && [p.x, p.y, p.w, p.h].every((v) => Number.isFinite(v))) out.placement = clampPlacement(p);
  return out;
}

function clampPlacement(p) {
  const w = Math.max(0.05, Math.min(1, p.w));
  const hh = Math.max(0.05, Math.min(1, p.h));
  return { x: Math.max(0, Math.min(1 - w, p.x)), y: Math.max(0, Math.min(1 - hh, p.y)), w, h: hh };
}

function normalizeSlide(raw, index, total) {
  const source = raw && typeof raw === "object" ? clone(raw) : {};
  let type = String(source.type || (index === 0 ? "title" : "content"));
  const aliases = { conclusion: "closing", progress: "content", barCompare: "statsCompare", fullImage: "hero", integratedContent: "imageText", splitImage: "imageText", integratedTitle: "title" };
  type = aliases[type] ?? type;
  if (!SLIDE_TYPES.includes(type)) throw new Error(`${index + 1}枚目: 未対応のレイアウト「${type}」です。`);
  const slide = { ...source, type };
  for (const key of ["title", "subtitle", "subhead", "takeaway", "message", "notes", "text", "author", "summary", "conclusion", "action", "root", "centerText", "source"]) {
    if (typeof slide[key] === "string") slide[key] = EMPHASIS_KEYS.has(key) ? keepEmphasis(slide[key]) : strip(slide[key]);
  }
  if (!slide.takeaway && TITLED(type) && source.keyMessage) slide.takeaway = keepEmphasis(source.keyMessage);
  delete slide.keyMessage;
  delete slide.illustration;
  if (type === "content" && !Array.isArray(slide.points)) {
    slide.points = Array.isArray(source.items) ? source.items.map((item) => (typeof item === "string" ? item : [item.headline || item.title, item.detail || item.desc].filter(Boolean).join("："))) : [];
    delete slide.items;
  }
  if (["process", "processList"].includes(type) && !Array.isArray(slide.steps)) slide.steps = (source.items || []).map((item) => (typeof item === "string" ? item : item.title || item.headline || "")).filter(Boolean);
  if (type === "kpi" && !Array.isArray(slide.items) && Array.isArray(source.metrics)) slide.items = source.metrics;
  if (type === "statement" && !strip(slide.text)) slide.text = slide.takeaway || slide.title || "メッセージ";
  if (typeof slide.visualAsset === "string" && !E.PHOTOS[slide.visualAsset] && slide.visualAsset !== "none") delete slide.visualAsset;
  if (Array.isArray(slide.items)) slide.items = slide.items.map((item) => (item && typeof item === "object" && item.icon && !E.icons[item.icon] ? { ...item, icon: undefined } : item));
  const media = normalizeMedia(slide.media);
  if (media) slide.media = media; else delete slide.media;
  if (slide.animation && !E.BUILDS.includes(slide.animation)) delete slide.animation;
  if (slide.animation === "auto") delete slide.animation;
  if (slide.photoMotion && !PHOTO_MOTIONS.some(([value]) => value === slide.photoMotion)) delete slide.photoMotion;
  // A slide's own motion: unknown values and "auto" (follow the deck) are dropped.
  for (const [key, info] of [["kinetic", KINETIC_INFO], ["backdrop", BACKDROP_INFO], ["entrance", ENTRANCE_INFO], ["emphasis", EMPHASIS_INFO], ["transition", SLIDE_TRANSITION_INFO]]) {
    if (slide[key] != null && (!Object.hasOwn(info, slide[key]) || slide[key] === "auto")) delete slide[key];
  }
  if (Array.isArray(slide.details)) {
    // Evidence: the judgement (text), and optionally a breakdown, where the figures come from, and assumptions.
    const rowsOf = (rows) => (Array.isArray(rows) ? rows : []).filter((row) => row && strip(row.label)).map((row) => ({ label: strip(row.label).slice(0, 40), value: strip(row.value ?? "").slice(0, 30) })).slice(0, 8);
    slide.details = slide.details.filter((d) => d && typeof d.target === "string" && strip(d.text)).map((d) => ({
      target: d.target.slice(0, 30), ...(strip(d.title) ? { title: strip(d.title).slice(0, 60) } : {}), text: String(d.text).slice(0, 400),
      ...(rowsOf(d.rows).length ? { rows: rowsOf(d.rows) } : {}), ...(strip(d.source) ? { source: strip(d.source).slice(0, 120) } : {}), ...(strip(d.note) ? { note: String(d.note).trim().slice(0, 200) } : {}),
    })).slice(0, 12);
    if (!slide.details.length) delete slide.details;
  } else delete slide.details;
  if (typeof slide.drillOf === "string" && slide.drillOf.trim() && index > 0 && TITLED(type)) slide.drillOf = slide.drillOf.trim().slice(0, 30);
  else delete slide.drillOf;
  // Objects placed by hand, their animations, and the id that links to this slide point at.
  const objects = E.normalizeObjects(slide.elements);
  if (objects.length) slide.elements = objects; else delete slide.elements;
  const timeline = E.normalizeTimeline(slide.timeline, slide).slice(0, 5000);
  if (timeline.length) slide.timeline = timeline; else delete slide.timeline;
  // 読み取り順序: only ids of objects that exist, each once (none set means the stacking order).
  const reading = Array.isArray(slide.readingOrder) && objects.length ? [...new Set(slide.readingOrder.filter((id) => objects.some((o) => o.id === id)))] : [];
  if (reading.length > 1) slide.readingOrder = reading; else delete slide.readingOrder;
  // 画面切り替え: how long the way in takes (ms) and moving on by itself after some seconds.
  if (Number(slide.transitionDur) >= 100 && Number(slide.transitionDur) <= 10000) slide.transitionDur = Math.round(Number(slide.transitionDur)); else delete slide.transitionDur;
  if (!Object.hasOwn(E.TRANSITION_SOUNDS, slide.transitionSound)) delete slide.transitionSound;
  if (!Object.values(E.TRANSITION_OPTIONS).some((list) => list.some(([key]) => key === slide.transitionDir))) delete slide.transitionDir;
  if (slide.advance != null && Number.isFinite(Number(slide.advance)) && Number(slide.advance) >= 0) slide.advance = Math.min(600, Math.round(Number(slide.advance) * 10) / 10); else delete slide.advance;
  if (!(typeof slide.sid === "string" && /^[A-Za-z0-9_-]{1,32}$/.test(slide.sid))) delete slide.sid;
  // 背景の書式設定: the slide's own background (a colour, a picture).
  const background = E.normalizeBackground?.(slide.background);
  if (background) slide.background = background; else delete slide.background;
  // 背景グラフィックを表示しない: the deck's master objects stay off this slide.
  if (slide.hideMaster === true) slide.hideMaster = true; else delete slide.hideMaster;
  // ノートの書式: the formatted notes stay only while they say what the plain notes say.
  if (slide.notesRich != null) { if (E.noteOf(slide).formatted) slide.notesRich = E.noteHtml(slide.notesRich); else delete slide.notesRich; }
  // A page brought over from PowerPoint as it looked keeps its title as an object and may use the cover's master.
  if (type === "blank") {
    if (slide.hideTitle === true) slide.hideTitle = true; else delete slide.hideTitle;
    if (slide.master !== "title" && slide.master !== "content" && slide.master !== "source") delete slide.master;
    if (slide.master === "source") {
      const v = slide.sourceViewport;
      if (!v || ![v.x, v.y, v.w, v.h].every(Number.isFinite)) delete slide.sourceViewport;
    } else delete slide.sourceViewport;
  } else { delete slide.hideTitle; delete slide.master; }
  if (slide.hidden !== true) delete slide.hidden;
  // PowerPoint's sections: the first slide of each carries the section's name.
  if (typeof slide.section === "string" && slide.section.trim()) slide.section = slide.section.trim().slice(0, 40); else delete slide.section;
  // Review comments on the slide (校閲 → 新しいコメント).
  const comments = (Array.isArray(slide.comments) ? slide.comments : []).filter((c) => c && typeof c.text === "string" && c.text.trim()).slice(0, 200)
    .map((c, i) => normalizeComment(c, `c${Date.now().toString(36)}${i}`, true));
  if (comments.length) slide.comments = comments; else delete slide.comments;
  // The first page is the cover: a title page, or a cover brought over from PowerPoint (a 白紙 page).
  if (index === 0 && total > 1 && type !== "title" && slide.master !== "source" && !(type === "blank" && slide.elements?.length)) return { type: "title", title: strip(slide.title) || "無題の資料" };
  return slide;
}

/** A review comment (or, thread = false, a reply): text, who and when, kept as people wrote it. */
function normalizeComment(c, fallbackId, thread = false) {
  const str = (v, max) => (typeof v === "string" && v.trim() ? v.trim().slice(0, max) : null);
  const out = { id: /^[A-Za-z0-9_-]{1,32}$/.test(c.id || "") ? c.id : fallbackId, text: c.text.trim().slice(0, 2000) };
  if (str(c.at, 40)) out.at = str(c.at, 40);
  if (str(c.by, 40)) out.by = str(c.by, 40).replace(/[<>]/g, "");
  if (/^[a-z0-9]{6,24}$/.test(c.uid || "")) out.uid = c.uid;
  if (str(c.edited, 40)) out.edited = str(c.edited, 40);
  if (!thread) return out;
  if (/^[A-Za-z0-9_-]{1,40}$/.test(c.anchor || "")) out.anchor = c.anchor;
  if (c.done === true) {
    out.done = true;
    if (str(c.doneBy, 40)) out.doneBy = str(c.doneBy, 40).replace(/[<>]/g, "");
  }
  const replies = (Array.isArray(c.replies) ? c.replies : []).filter((r) => r && typeof r.text === "string" && r.text.trim()).slice(-100)
    .map((r, i) => normalizeComment(r, `${out.id.slice(0, 24)}r${i}`));
  if (replies.length) out.replies = replies;
  return out;
}

function normalizeMotion(motion = {}) {
  return {
    entrance: motion.entrance === "none" || Object.hasOwn(E.ENTRANCES, motion.entrance ?? "") ? motion.entrance : "rise",
    hover: motion.hover === "none" || Object.hasOwn(E.HOVERS, motion.hover ?? "") ? motion.hover : "lift",
    numbers: motion.numbers !== false,
    ambient: motion.ambient !== false,
    kinetic: motion.kinetic === "none" || E.KINETIC[motion.kinetic] ? motion.kinetic : "mask",
    backdrop: E.BACKDROPS[motion.backdrop] ? motion.backdrop : "none",
    emphasis: Object.hasOwn(E.EMPHASES, motion.emphasis ?? "") ? motion.emphasis : "marker",
    draw: motion.draw !== false,
  };
}

/** Any deck-like JSON (studio decks, generated decks, imports) → the studio's deck. `base` lends its design. */
function normalizeDeck(value, base = null) {
  const source = value?.structuredContent || value;
  let slides;
  let meta = {};
  if (Array.isArray(source)) slides = source;
  else if (Array.isArray(source?.slideData)) { slides = source.slideData; meta = { ...source, title: source.deckTitle ?? source.title }; }
  else if (Array.isArray(source?.slides)) { slides = source.slides; meta = source; }
  else if (Array.isArray(source?.deck?.slides)) { slides = source.deck.slides; meta = source.deck; }
  else throw new Error("slideData 配列が見つかりません。");
  if (slides.length < 1 || slides.length > 500) throw new Error("スライド枚数は1〜500枚にしてください。");
  const normalized = slides.map((slide, index) => normalizeSlide(slide, index, slides.length));
  if (meta.fidelity === "exact") normalized.forEach((slide, index) => {
    const source = slides[index];
    const original = Array.isArray(source?.elements) ? source.elements : [];
    if ((slide.elements?.length ?? 0) !== original.length) throw new Error(`${index + 1}枚目の部品が一部取り込めません。原本を確認してください。`);
    if ((slide.timeline?.length ?? 0) !== (source.timeline?.length ?? 0)) throw new Error(`${index + 1}枚目のアニメーションが一部取り込めません。`);
    original.forEach((object, i) => {
      const kept = slide.elements[i];
      if (object.gradient && kept.gradient?.stops?.length !== object.gradient.stops?.length)
        throw new Error(`${index + 1}枚目のグラデーションが一部取り込めません。`);
      if (object.kind === "table" && (kept.cells.length !== object.cells.length || kept.cells[0]?.length !== object.cells[0]?.length))
        throw new Error(`${index + 1}枚目の表が一部取り込めません。`);
      if (object.kind === "chart" && (kept.chart.labels.length !== object.chart.labels.length || kept.chart.series.length !== object.chart.series.length))
        throw new Error(`${index + 1}枚目のグラフが一部取り込めません。`);
    });
  });
  const transition = TRANSITIONS.includes(meta.transition) ? meta.transition : base?.transition ?? "fade";
  // A design the engine does not know opens in the default theme.
  return withTheme({
    title: strip(meta.title || meta.deckTitle || normalized[0]?.title || "無題の資料").slice(0, 100),
    purpose: strip(meta.purpose ?? $("purposeInput").value ?? "").slice(0, 180),
    audience: strip(meta.audience ?? $("audienceInput").value ?? "").slice(0, 80),
    theme: THEME_IDS.has(meta.theme) ? meta.theme : preferredTheme(),
    transition,
    motion: normalizeMotion(meta.motion ?? base?.motion ?? DEFAULT_MOTION),
    memo: String(meta.memo ?? base?.memo ?? "").slice(0, 2000),
    ...(guidesOf(meta.guides ?? base?.guides) ? { guides: guidesOf(meta.guides ?? base?.guides) } : {}),
    // スライド マスター: the objects every slide shows.
    ...(masterObjectsOf(meta.masterObjects ?? base?.masterObjects) ? { masterObjects: masterObjectsOf(meta.masterObjects ?? base?.masterObjects) } : {}),
    // 配布資料マスター・ノート マスター.
    ...(printMastersOf(meta.printMasters ?? base?.printMasters) ? { printMasters: printMastersOf(meta.printMasters ?? base?.printMasters) } : {}),
    // スライド番号の開始番号.
    ...(firstNumberOf(meta.firstNumber ?? base?.firstNumber) !== 1 ? { firstNumber: firstNumberOf(meta.firstNumber ?? base?.firstNumber) } : {}),
    // 既定の図形・テキスト ボックス・線.
    ...(objectDefaultsIn(meta.objectDefaults ?? base?.objectDefaults) ? { objectDefaults: objectDefaultsIn(meta.objectDefaults ?? base?.objectDefaults) } : {}),
    // スライド ショーの設定・目的別スライド ショー.
    ...(showOf(meta.show ?? base?.show) ? { show: showOf(meta.show ?? base?.show) } : {}),
    ...(customShowsOf(meta.customShows ?? base?.customShows) ? { customShows: customShowsOf(meta.customShows ?? base?.customShows) } : {}),
    // ファイル → 情報: the deck's properties.
    ...(infoOf(meta.info ?? base?.info) ? { info: infoOf(meta.info ?? base?.info) } : {}),
    slides: normalized,
  });
}

/** The deck's guides (表示 → ガイド): positions in slide px, up to 20 each way. */
function guidesOf(value) {
  const list = (axis, max) => (Array.isArray(value?.[axis]) ? value[axis] : []).map(Number).filter((v) => Number.isFinite(v) && v >= 0 && v <= max).slice(0, 20).map((v) => Math.round(v * 2) / 2);
  const guides = { x: list("x", E.W), y: list("y", E.H) };
  return guides.x.length || guides.y.length ? guides : null;
}

/** スライド マスター (deck.masterObjects): objects checked as a slide's are, up to 500. */
function masterObjectsOf(value) {
  const list = E.normalizeObjects(Array.isArray(value) ? value : []).slice(0, 500);
  return list.length ? list : null;
}
/** The deck's slides and, when it has master objects, the master as one more "slide" (so media and text are found there too). */
/** 配布資料マスター・ノート マスター (deck.printMasters): each master's page and the words and pictures on it, checked as the studio keeps them. */
function printMastersOf(value) {
  return normalizeMasters(value, { objects: (list) => E.normalizeObjects(list).slice(0, 30) });
}
// Pictures of the 配布資料マスター・ノート マスター are found and carried like the slide master's.
const withMaster = (deck) => [...(deck?.slides || []), ...(deck?.masterObjects?.length ? [{ elements: deck.masterObjects }] : []), ...masterSlides(deck?.printMasters)];

/** スライド番号の開始番号 (deck.firstNumber): a whole number from 0 to 9999; 1 unless set. */
function firstNumberOf(value) {
  const n = Number(value);
  return Number.isInteger(n) && n >= 0 && n <= 9999 ? n : 1;
}

/** 既定の図形・テキスト ボックス・線 (deck.objectDefaults): each look checked as an object's would be. */
function objectDefaultsIn(value) {
  return ops.objectDefaultsOf(value, (kind, style) => {
    const probe = kind === "line" ? { id: "dflt", kind: "line", x1: 0, y1: 0, x2: 100, y2: 0, ...style } : { id: "dflt", kind, x: 0, y: 0, w: 100, h: 100, ...(kind === "shape" ? { shape: "rect" } : {}), text: "<p>a</p>", ...style };
    const o = E.normalizeObject(probe);
    return o ? ops.snapToPalette(o, E.PALETTE) : null;
  });
}

/** 既定の図形に設定 (and テキスト ボックス・線): new ones in this deck start with this object's look (one ⌘Z). */
function setObjectDefault(kind, o) {
  if (!state.deck || isFinal(state.deck)) return;
  const next = objectDefaultsIn({ ...(state.deck.objectDefaults || {}), [kind]: ops.defaultStyleOf(kind, E.withDefaults(o)) });
  pushUndo();
  if (next) state.deck.objectDefaults = next; else delete state.deck.objectDefaults;
  saveCurrent();
  scheduleVersion();
  toast({ shape: "既定の図形に設定しました。この資料で新しく描く図形がこの書式になります", text: "既定のテキスト ボックスに設定しました。この資料で新しく置くテキスト ボックスがこの書式になります", line: "既定の線に設定しました。この資料で新しく引く線がこの書式になります" }[kind]);
}

/** 背景の保存: the slide's background picture as a file. */
async function saveBackground() {
  const bg = state.deck?.slides[state.selected]?.background;
  const blob = bg?.image ? await mediaBlob(bg.image).catch(() => null) : null;
  if (!blob) { toast("保存できる背景の図がありません"); return; }
  const ext = { "image/png": "png", "image/jpeg": "jpg", "image/webp": "webp", "image/gif": "gif", "image/svg+xml": "svg" }[blob.type] || "png";
  await downloadBlob(blob, `背景.${ext}`);
  toast("背景の図を保存しました");
}

/** Make the deck acceptable to the server schema without changing what the user sees. */
function sanitizeSlide(slide) {
  const out = clone(slide);
  const fill = (value) => (strip(value) ? String(value).trim() : "—");
  const fixFields = (target, fields) => {
    for (const field of fields) {
      const value = target[field.key];
      if (field.kind === "text" || field.kind === "area") {
        if (field.req) target[field.key] = fill(value);
        else if (value != null && !strip(value)) delete target[field.key];
      } else if (field.kind === "list") {
        let list = (Array.isArray(value) ? value : []).map((line) => String(line ?? "").trim()).filter((line) => strip(line));
        if (field.max) list = list.slice(0, field.max);
        while (list.length < (field.min ?? 0)) list.push("—");
        if (list.length || field.min || ["points", "leftItems", "rightItems"].includes(field.key)) target[field.key] = list;
        else delete target[field.key];
      } else if (field.kind === "group") {
        let items = (Array.isArray(value) ? value : []).filter((item) => item && typeof item === "object");
        if (field.max) items = items.slice(0, field.max);
        while (items.length < (field.min ?? 0)) items.push(field.newItem ? field.newItem() : {});
        items.forEach((item) => fixFields(item, field.fields));
        target[field.key] = items;
      } else if (field.kind === "number" && field.free) {
        const empty = value === "" || value == null || Number.isNaN(Number(value));
        if (empty && field.optional) delete target[field.key];
        else target[field.key] = empty ? 0 : Number(value);
      } else if (field.kind === "number") {
        if (value === "" || value == null || Number.isNaN(Number(value))) delete target[field.key];
        else target[field.key] = Math.max(field.min ?? 0, Math.min(field.max ?? 99, Math.round(Number(value))));
      } else if (field.kind === "select" && field.number) {
        if (value === "" || value == null) delete target[field.key];
        else target[field.key] = Number(value);
      }
    }
  };
  fixFields(out, specFor(out.type));
  if (out.type === "table") {
    const headers = (out.headers || []).map((cell) => String(cell ?? "").trim()).slice(0, 6);
    while (headers.length < 2) headers.push("—");
    out.headers = headers;
    const rows = (out.rows || []).map((row) => headers.map((_, index) => String(row?.[index] ?? "").trim())).slice(0, 8);
    out.rows = rows.length ? rows : [headers.map(() => "")];
  }
  if (out.type === "closing" && !strip(out.title)) delete out.title;
  if (out.image && typeof out.image === "object" && !out.image.data) delete out.image;
  if (out.type === "cards" && Array.isArray(out.items)) out.items = out.items.map((item) => (typeof item === "string" ? item : { title: fill(item.title), ...(strip(item.desc) ? { desc: String(item.desc).trim() } : {}), ...(item.icon ? { icon: item.icon } : {}) }));
  return out;
}

function serverDeck(deck = state.deck) {
  return {
    title: strip(deck.title) || "無題の資料",
    purpose: deck.purpose || "",
    audience: deck.audience || "",
    schemaVersion: "3.0",
    theme: deck.theme,
    ...(deck.accent ? { accent: deck.accent } : {}),
    transition: deck.transition || "fade",
    motion: deck.motion,
    memo: deck.memo || "",
    slides: deck.slides.map(sanitizeSlide),
  };
}

// ---------------------------------------------------------------- default slides & conversion

function defaultSlide(type) {
  const base = TITLED(type) ? { type, title: "スライドタイトル", takeaway: "このスライドで伝える結論" } : { type };
  const card = (n) => Array.from({ length: n }, (_, i) => ({ title: `要点${"ABCDEF"[i]}`, desc: "具体的な説明" }));
  switch (type) {
    case "title": return { type, title: state.deck?.title || "資料タイトル", subtitle: "", date: new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long" }) };
    case "section": return { type, title: "章タイトル" };
    case "blank": return { type, title: "" };
    case "closing": return { type, title: "次のアクション", message: "誰が・いつまでに・何をするか" };
    case "hero": return { type, title: "伝えたいひと言", takeaway: "写真に重ねて見せる補足の一文", visualAsset: "transformationRoadmap", photoMotion: "zoom" };
    case "statement": return { type, title: "ポイント", text: "いちばん伝えたい**ひと言**を大きく", takeaway: "" };
    case "content": return { ...base, points: ["結論を支える根拠", "実行時の重要ポイント", "期待できる効果"] };
    case "agenda": return { ...base, title: "本日のアジェンダ", items: ["背景：なぜ今取り組むか", "現状：何が起きているか", "提案：何をするか", "次のアクション"] };
    case "compare": return { ...base, leftTitle: "現状", rightTitle: "目指す姿", leftItems: ["現状の課題"], rightItems: ["目指す状態"] };
    case "beforeAfter": return { ...base, leftTitle: "Before", rightTitle: "After", leftItems: ["変える前"], rightItems: ["変えた後"] };
    case "process": return { ...base, steps: ["準備：目的と対象を決める", "実行：手順どおりに進める", "確認：結果を振り返る"] };
    case "processList": return { ...base, steps: ["準備：目的を共有", "実行：手順を実施", "定着：振り返り"] };
    case "flowChart": return { ...base, flows: [{ steps: ["開始", "処理", "判断", "完了"] }] };
    case "timeline": return { ...base, milestones: [{ date: "4月", label: "開始", state: "done" }, { date: "7月", label: "展開", state: "next" }, { date: "10月", label: "定着", state: "todo" }] };
    case "diagram": return { ...base, lanes: [{ title: "現場", items: ["受付", "対応"] }, { title: "本部", items: ["分析", "支援"] }] };
    case "cycle": return { ...base, centerText: "改善", items: [{ label: "計画" }, { label: "実行" }, { label: "評価" }, { label: "改善" }] };
    case "table": return { ...base, headers: ["項目", "現状", "目標"], rows: [["指標A", "—", "—"], ["指標B", "—", "—"]] };
    case "quote": return { ...base, text: "引用したい声やメッセージ", author: "発言者" };
    case "kpi": return { ...base, items: [{ value: "—", label: "指標A", change: "補足", status: "neutral" }, { value: "—", label: "指標B", change: "補足", status: "neutral" }, { value: "—", label: "指標C", change: "補足", status: "neutral" }] };
    case "dashboard": return { ...base, items: [{ value: "—", label: "指標A", change: "" }, { value: "—", label: "指標B", change: "" }], image: { chartType: "bar", data: { title: "推移", items: [{ label: "4月", value: 10 }, { label: "5月", value: 12 }, { label: "6月", value: 15 }] } } };
    case "statsCompare": return { ...base, leftTitle: "導入前", rightTitle: "導入後", stats: [{ label: "指標", leftValue: "—", rightValue: "—", trend: "neutral" }] };
    case "faq": return { ...base, title: "想定される質問", items: [{ q: "質問", a: "回答" }, { q: "質問", a: "回答" }] };
    case "pyramid": case "funnel": return { ...base, levels: [{ title: "段1", description: "" }, { title: "段2", description: "" }, { title: "段3", description: "" }] };
    case "imageText": return { ...base, points: ["図から読み取れる要点", "意思決定への示唆"], imagePosition: "left", image: { chartType: "bar", data: { title: "グラフタイトル", items: [{ label: "A", value: 10 }, { label: "B", value: 14 }, { label: "C", value: 12 }] } } };
    case "executiveSummary": return { ...base, conclusion: "結論を一文で", items: card(3), action: "お願いしたいこと" };
    case "headerTwoColumn": return { ...base, items: card(2) };
    case "headerThreeSummary": return { ...base, items: card(3), summary: "3つの要点から導くまとめ" };
    case "triangle": return { ...base, items: card(3) };
    case "venn": return { ...base, items: card(2) };
    case "matrix": case "swot": case "grid2x2": return { ...base, items: type === "swot" ? [{ title: "強み", desc: "" }, { title: "弱み", desc: "" }, { title: "機会", desc: "" }, { title: "脅威", desc: "" }] : card(4) };
    case "gantt": return { ...base, periods: ["4月", "5月", "6月", "7月"], items: [{ title: "設計", desc: "", start: 0, span: 1 }, { title: "開発", desc: "", start: 1, span: 2 }, { title: "展開", desc: "", start: 3, span: 1 }] };
    case "orgChart": return { ...base, root: "プロジェクト責任者", items: card(3) };
    case "waterfall": return { ...base, unit: "", items: [{ label: "前期", value: 100, total: true }, { label: "増加要因", value: 30 }, { label: "減少要因", value: -10 }, { label: "今期", value: 120, total: true }] };
    case "logicTree": return { ...base, root: "分解する課題", branches: [{ title: "要因A", items: ["具体的な事象"] }, { title: "要因B", items: ["具体的な事象"] }, { title: "要因C", items: ["具体的な事象"] }] };
    case "checklist": return { ...base, items: card(3).map((item) => ({ ...item, done: false })) };
    case "simulator": return { ...base, title: "条件を変えたときの試算", takeaway: "条件によって結果がどう変わるか", inputs: [{ label: "対象人数", value: 100, min: 10, max: 500, step: 10, unit: "人" }, { label: "1人あたりの効果", value: 5, min: 1, max: 20, unit: "時間" }], formula: "a × b", resultLabel: "合計の効果", resultUnit: "時間", compareLabel: "目標", compareValue: 1000 };
    case "gap": return { ...base, title: "目標までの不足と打ち手", takeaway: "打ち手を重ねると目標に届く", unit: "", targetLabel: "目標", target: 100, currentLabel: "現状", current: 60, measures: [{ title: "打ち手A", value: 20, desc: "" }, { title: "打ち手B", value: 15, desc: "" }, { title: "打ち手C", value: 10, desc: "" }] };
    default: return { ...base, items: card(3) };
  }
}

/** Pull the textual units out of any slide so switching layouts keeps the content. */
function extractUnits(slide) {
  const units = [];
  const push = (title, desc = "") => { if (strip(title) || strip(desc)) units.push({ title: strip(title), desc: strip(desc) }); };
  const split = (text) => E.splitLabel(text);
  for (const key of ["points", "steps", "leftItems", "rightItems"]) (slide[key] || []).forEach((item) => push(...split(item)));
  if (Array.isArray(slide.items)) slide.items.forEach((item) => (typeof item === "string" ? push(...split(item)) : push(item.title ?? item.label ?? item.q ?? item.value, item.desc ?? item.subLabel ?? item.a ?? item.change ?? "")));
  (slide.levels || []).forEach((item) => push(item.title, item.description));
  (slide.milestones || []).forEach((item) => push(item.label, item.date));
  (slide.lanes || []).forEach((lane) => push(lane.title, (lane.items || []).join("、")));
  (slide.stats || []).forEach((item) => push(item.label, [item.leftValue, item.rightValue].filter(Boolean).join(" → ")));
  (slide.flows || []).forEach((flow) => (flow.steps || []).forEach((step) => push(...split(step))));
  (slide.rows || []).forEach((row) => push(row[0], row.slice(1).join(" / ")));
  (slide.branches || []).forEach((branch) => push(branch.title, (branch.items || []).join("、")));
  (slide.measures || []).forEach((m) => push(m.title, m.desc));
  (slide.inputs || []).forEach((input) => push(input.label, `${input.value ?? ""}${input.unit ?? ""}`));
  if (slide.type === "statement" && slide.text) push(slide.text);
  return units;
}

function convertSlide(slide, type) {
  const next = defaultSlide(type);
  for (const key of ["title", "takeaway", "subhead", "source", "notes", "notesRich", "visualAsset", "customImage", "imagePlacement", "media", "photoMotion", "kinetic", "backdrop", "entrance", "emphasis", "transition", "drillOf", "elements", "timeline", "sid", "transitionDur", "transitionSound", "transitionDir", "advance", "hidden", "section", "comments", "background"]) {
    if (slide[key] != null && slide[key] !== "" && (key !== "takeaway" || TITLED(type))) next[key] = clone(slide[key]);
  }
  if (type === "closing" && !next.message) next.message = slide.takeaway || slide.message || "";
  if (type === "statement") next.text = strip(slide.takeaway) ? slide.takeaway : next.text;
  const units = extractUnits(slide);
  if (!units.length) return next;
  const line = (u) => (u.desc ? `${u.title}：${u.desc}` : u.title);
  const cards = (min, max) => { const list = units.slice(0, max).map((u) => ({ title: u.title || "項目", desc: u.desc })); while (list.length < min) list.push({ title: "項目", desc: "" }); return list; };
  switch (type) {
    case "content": next.points = units.slice(0, 10).map(line); break;
    case "agenda": next.items = units.slice(0, 10).map(line); break;
    case "process": case "processList": next.steps = units.slice(0, type === "process" ? 6 : 8).map(line); while (next.steps.length < 2) next.steps.push("工程"); break;
    case "compare": case "beforeAfter": { const half = Math.ceil(units.length / 2); next.leftItems = units.slice(0, half).map(line); next.rightItems = units.slice(half).map(line); break; }
    case "flowChart": next.flows = [{ steps: units.slice(0, 5).map((u) => u.title) }]; while (next.flows[0].steps.length < 2) next.flows[0].steps.push("工程"); break;
    case "timeline": next.milestones = units.slice(0, 8).map((u, i) => ({ date: u.desc.length <= 12 ? u.desc : `STEP ${i + 1}`, label: u.title, state: "todo" })); while (next.milestones.length < 2) next.milestones.push({ date: "", label: "出来事", state: "todo" }); break;
    case "diagram": next.lanes = units.slice(0, 5).map((u) => ({ title: u.title.slice(0, 10), items: u.desc ? u.desc.split(/[、,]/).slice(0, 6) : [] })); while (next.lanes.length < 2) next.lanes.push({ title: "担当", items: [] }); break;
    case "cycle": next.items = units.slice(0, 6).map((u) => ({ label: u.title.slice(0, 10), subLabel: u.desc.slice(0, 20) })); while (next.items.length < 3) next.items.push({ label: "要素" }); break;
    case "table": next.headers = ["項目", "内容"]; next.rows = units.slice(0, 8).map((u) => [u.title, u.desc]); break;
    case "kpi": case "dashboard": next.items = units.slice(0, 4).map((u) => ({ value: u.desc.slice(0, 10) || "—", label: u.title.slice(0, 12), change: "" })); while (next.items.length < (type === "dashboard" ? 2 : 1)) next.items.push({ value: "—", label: "指標" }); break;
    case "statsCompare": next.stats = units.slice(0, 6).map((u) => ({ label: u.title, leftValue: "—", rightValue: u.desc.slice(0, 12) || "—", trend: "neutral" })); break;
    case "faq": next.items = units.slice(0, 4).map((u) => ({ q: u.title, a: u.desc || "回答" })); break;
    case "pyramid": case "funnel": next.levels = cards(3, 5).map((c) => ({ title: c.title, description: c.desc })); break;
    case "imageText": next.points = units.slice(0, 8).map(line); break;
    case "quote": next.text = units.map(line).join(" ").slice(0, 90); break;
    case "statement": if (!strip(slide.takeaway)) next.text = units.map((u) => u.title).join("・").slice(0, 60); break;
    case "gantt": next.items = units.slice(0, 7).map((u, i) => ({ title: u.title.slice(0, 14), desc: "", start: Math.min(i, 3), span: 1 })); while (next.items.length < 2) next.items.push({ title: "タスク", start: 0, span: 1 }); break;
    case "checklist": next.items = cards(1, 7).map((c) => ({ ...c, done: false })); break;
    case "headerTwoColumn": next.items = cards(2, 2); break;
    case "headerThreeSummary": case "triangle": next.items = cards(3, 3); break;
    case "venn": next.items = cards(2, 3); break;
    case "matrix": case "swot": next.items = cards(4, 4); break;
    case "executiveSummary": next.items = cards(2, 3); break;
    case "bulletCards": next.items = cards(1, 4).map((c) => ({ ...c, desc: c.desc || "説明" })); break;
    case "stepUp": next.items = cards(2, 5).map((c) => ({ ...c, desc: c.desc || "説明" })); break;
    case "roadmap": next.items = cards(2, 6); break;
    case "orgChart": next.items = cards(2, 5); break;
    case "logicTree": next.branches = units.slice(0, 4).map((u) => ({ title: u.title.slice(0, 14), items: u.desc ? u.desc.split(/[、,]/).map((x) => x.trim()).filter(Boolean).slice(0, 3) : [] })); while (next.branches.length < 2) next.branches.push({ title: "要因", items: [] }); break;
    case "waterfall": { const found = units.map((u) => ({ label: u.title.slice(0, 8), value: Number(String(u.desc).replace(/[,，]/g, "").match(/[-−]?\d+(?:\.\d+)?/)?.[0]?.replace("−", "-")) })).filter((item) => Number.isFinite(item.value)).slice(0, 8); if (found.length >= 2) next.items = found; break; }
    case "cards": case "headerCards": case "grid2x2": next.items = cards(1, type === "grid2x2" ? 4 : 6); break;
    case "gap": next.measures = cards(1, 5).map((c) => ({ title: c.title.slice(0, 16), value: Number(String(c.desc).replace(/[,，]/g, "").match(/\d+(?:\.\d+)?/)?.[0]) || 10, ...(c.desc ? { desc: c.desc.slice(0, 30) } : {}) })); break;
    default: break;
  }
  return next;
}

// ---------------------------------------------------------------- photos & videos kept in this browser (IndexedDB)

const mediaUrls = {};

function openDb() {
  if (typeof indexedDB === "undefined") return Promise.resolve(null);
  return new Promise((resolve, reject) => {
    const request = indexedDB.open("html-pptx-studio", 1);
    request.onupgradeneeded = () => {
      for (const name of ["media", "versions"]) if (!request.result.objectStoreNames.contains(name)) request.result.createObjectStore(name, { keyPath: "id" });
    };
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function idb(mode, fn, storeName) {
  const db = await openDb();
  if (!db) return null;
  return new Promise((resolve, reject) => {
    const tx = db.transaction(storeName, mode);
    const result = fn(tx.objectStore(storeName));
    tx.oncomplete = () => { db.close(); resolve(result?.result ?? null); };
    tx.onerror = () => { db.close(); reject(tx.error); };
  });
}

async function putMedia(blob, name = "") {
  const id = `${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
  await idb("readwrite", (store) => store.put({ id, blob, type: blob.type, name, size: blob.size, savedAt: Date.now() }), "media");
  const src = `idb:${id}`;
  mediaUrls[src] = URL.createObjectURL(blob);
  return src;
}

/** A file from a shared room, kept in this browser under the same id the deck uses. */
async function putMediaAs(src, blob) {
  if (!src?.startsWith("idb:")) return;
  const id = src.slice(4);
  await idb("readwrite", (store) => store.put({ id, blob, type: blob.type, name: "", size: blob.size, savedAt: Date.now() }), "media");
  mediaUrls[src] = URL.createObjectURL(blob);
}

async function mediaBlob(src) {
  if (src?.startsWith("idb:")) return (await idb("readonly", (store) => store.get(src.slice(4)), "media"))?.blob ?? null;
  if (src?.startsWith("data:")) return dataUrlToBlob(src);
  if (src?.startsWith("asset:")) {
    const file = E.PHOTOS[src.slice(6)]?.[0];
    return file ? (await fetch(`/assets/${file}`)).blob() : null;
  }
  return null;
}

/** Every photo, video and animation a slide uses: its own media and the objects placed on it. */
const mediaSources = (slide) => [slide.media?.src, slide.background?.image, ...(slide.elements || []).flatMap((o) => [o?.src, o?.fillImg])].filter((src) => typeof src === "string" && src);
// Bumped when a browser-kept file becomes available, so thumbnails drawn without it are drawn again.
let mediaEpoch = 0;

/** Load the browser-kept photos and videos a deck refers to; true when something new arrived. */
async function ensureMedia(deck = state.deck) {
  if (!deck) return false;
  let added = false;
  for (const slide of withMaster(deck)) {
    for (const src of mediaSources(slide)) {
      if (!src.startsWith("idb:") || mediaUrls[src]) continue;
      try {
        const blob = await mediaBlob(src);
        if (blob) { mediaUrls[src] = URL.createObjectURL(blob); added = true; }
      } catch { /* missing on this browser */ }
    }
  }
  if (added) mediaEpoch += 1;
  return added;
}

/** Pictures pasted into JSON (data: URLs) move into the browser store so the deck itself stays small. */
async function storeInlineMedia(deck) {
  for (const slide of withMaster(deck)) {
    if (slide.media?.src?.startsWith("data:")) {
      try { slide.media.src = await putMedia(await dataUrlToBlob(slide.media.src), slide.media.name || ""); } catch { /* keep the data URL */ }
    }
    for (const o of slide.elements || []) {
      if (o?.src?.startsWith("data:")) {
        try { o.src = await putMedia(await dataUrlToBlob(o.src), o.fileName || ""); } catch { /* keep the data URL */ }
      }
      if (o?.fillImg?.startsWith("data:")) {
        try { o.fillImg = await putMedia(await dataUrlToBlob(o.fillImg), "図形の塗りつぶし"); } catch { /* keep the data URL */ }
      }
    }
    if (slide.background?.image?.startsWith("data:")) {
      try { slide.background.image = await putMedia(await dataUrlToBlob(slide.background.image), "背景"); } catch { /* keep the data URL */ }
    }
  }
  return deck;
}

async function downscaleImage(file, maxSide = 1920, type = "image/jpeg") {
  const url = URL.createObjectURL(file);
  try {
    const img = await new Promise((resolve, reject) => { const image = new Image(); image.onload = () => resolve(image); image.onerror = reject; image.src = url; });
    const scale = Math.min(1, maxSide / Math.max(img.width, img.height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(img.width * scale);
    canvas.height = Math.round(img.height * scale);
    const ctx = canvas.getContext("2d");
    if (type === "image/jpeg") { ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, canvas.width, canvas.height); }
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
    return await new Promise((resolve) => canvas.toBlob(resolve, type, 0.88));
  } finally {
    URL.revokeObjectURL(url);
  }
}

// ---------------------------------------------------------------- rendering (engine), caching and text fitting

const loadedFonts = new Set();
function useFonts(themeIds) {
  let changed = false;
  for (const id of themeIds) if (THEME_IDS.has(id) && !loadedFonts.has(id)) { loadedFonts.add(id); changed = true; }
  if (!changed) return;
  const href = E.fontHref([...loadedFonts]);
  if (href) $("slideFonts").href = href;
}

function renderOptions(extra = {}) {
  return { deck: state.deck, assetBase: "/assets/", mediaUrls, ...extra };
}

/** Everything a slide's picture depends on besides the slide itself: design, position, chapter. */
function contextKey(index, deck = state.deck) {
  const chapters = deck.slides.slice(0, index + 1).filter((slide) => slide.type === "section").map((slide) => `${slide.sectionNo ?? ""}${slide.title}`);
  // Page numbers count the story; a slide shows arrows for its deep-dive pages; a deep-dive page names its slide.
  const story = E.storyMap(deck.slides);
  const parent = story.parent[index];
  const drills = (story.drills[index] ?? []).map((drill) => `${drill.index}:${drill.target}`);
  return JSON.stringify([deck.theme, deck.accent ?? "", deck.motion, deck.title, deck.purpose, deck.audience, index, deck.slides.length, chapters, story.no[index], story.order.length, drills, parent != null ? deck.slides[parent]?.title ?? "" : null]);
}

const fitState = { byKey: new Map(), epoch: 0, running: false, timer: null };
const slideKey = (index, deck = state.deck) => `${fitState.epoch}|${contextKey(index, deck)}|${JSON.stringify(deck.slides[index])}`;

function fitFor(index) {
  if (!state.deck?.slides[index]) return null;
  return fitState.byKey.get(slideKey(index)) ?? null;
}

const measured = (index) => Boolean(fitFor(index));

/** Measure every slide that changed (off screen, at full size): text scale and what still overflows. */
function scheduleMeasure(delay = 60) {
  clearTimeout(fitState.timer);
  fitState.timer = setTimeout(runMeasure, delay);
}

async function runMeasure() {
  if (fitState.running || !state.deck) return;
  fitState.running = true;
  try {
    await document.fonts?.ready;
    let dirty = false;
    let started = performance.now();
    for (let index = 0; state.deck && index < state.deck.slides.length; index += 1) {
      if (measured(index)) continue;
      measureSlide(index);
      dirty = true;
      if (performance.now() - started > 40) {
        await wait(0);
        started = performance.now();
      }
    }
    if (fitState.byKey.size > 400) {
      const keep = new Set(state.deck.slides.map((_, index) => slideKey(index)));
      for (const key of fitState.byKey.keys()) if (!keep.has(key)) fitState.byKey.delete(key);
    }
    if (dirty) {
      renderFilmstrip();
      if (state.view === "grid") renderStage();
      renderIssueSummary();
      refreshInspectorIssues();
    }
  } finally {
    fitState.running = false;
  }
  if (state.deck && state.deck.slides.some((_, index) => !measured(index))) scheduleMeasure(30);
}

function measureSlide(index) {
  const root = $("measureRoot");
  const el = E.render(state.deck.slides[index], renderOptions({ index, mode: "thumb" }));
  root.replaceChildren(el);
  const result = E.fit(el);
  fitState.byKey.set(slideKey(index), { fs: result.fs, ts: result.ts, ...(result.objs ? { objs: result.objs } : {}), issues: result.issues.map((issue) => ({ ...issue, slide: index })) });
  root.replaceChildren();
}

async function measureAll() {
  if (!state.deck) return;
  await document.fonts?.ready;
  for (let index = 0; index < state.deck.slides.length; index += 1) if (!measured(index)) measureSlide(index);
}

// Web fonts change line breaks: measure again once they arrive.
document.fonts?.addEventListener?.("loadingdone", () => {
  clearTimeout(fitState.fontTimer);
  fitState.fontTimer = setTimeout(() => {
    if (!state.deck) return;
    fitState.epoch += 1;
    fitState.byKey.clear();
    thumbCache.clear();
    scheduleMeasure(0);
    if (!state.inline && !state.motionPreview) renderStage();
  }, 250);
});

// Rendered thumbnails are reused until the slide (or its design) changes.
const thumbCache = new Map();
/** The master's objects as a string, so a change to them draws every thumbnail again. */
const masterStamp = () => (state.deck?.masterObjects?.length ? JSON.stringify(state.deck.masterObjects) : "");
function thumb(index, variant = "film") {
  const fit = fitFor(index);
  const key = `${variant}|${slideKey(index)}|${masterStamp()}|${fit ? `${fit.fs},${fit.ts},${JSON.stringify(fit.objs ?? "")}` : "-"}|${mediaSources(state.deck.slides[index] || {}).filter((src) => mediaUrls[src]).length}|${mediaEpoch}`;
  const cached = thumbCache.get(key);
  if (cached) return cached;
  const el = E.render(state.deck.slides[index], renderOptions({ index, mode: "thumb", fit: fit ?? undefined }));
  const scaler = E.mount(el);
  thumbCache.set(key, scaler);
  if (thumbCache.size > 500) for (const old of [...thumbCache.keys()].slice(0, 200)) thumbCache.delete(old);
  return scaler;
}

/** A one-off picture of any slide (proposals, variants, previews). */
function slidePicture(slide, index, deck = state.deck) {
  const el = E.render(slide, { ...renderOptions(), deck, index, mode: "thumb" });
  return E.mount(el);
}

// ---------------------------------------------------------------- deck lifecycle

function loadDeck(deck, { source = "", keepUndo = false, imported = null, savedDeckId = null, selected = 0 } = {}) {
  state.masterEdit = false;
  // Opening another deck leaves a shared one (its room keeps the shared deck for the others).
  if (coedit?.active && source !== "共同編集") coedit.disconnect();
  if (autoImageRun) autoImageRun.cancelled = true;
  autoImageRun = null;
  stopMotionPreview({ render: false });
  state.imported = imported;
  if (!keepUndo && state.deck) pushUndo();
  state.deck = adoptDeck(deck);
  showImageCoverage();
  state.selected = keepUndo ? Math.min(state.selected, deck.slides.length - 1) : Math.max(0, Math.min(selected, deck.slides.length - 1));
  state.savedDeckId = savedDeckId;
  $("deckTitleInput").value = deck.title;
  $("editTab").disabled = false;
  $("resumeEditBtn").classList.remove("hidden");
  useFonts([deck.theme]);
  state.historyId = saveHistory();
  saveCurrent();
  setMode("edit");
  renderAll();
  scheduleMeasure(0);
  loadChat();
  saveVersion(source ? `${source}で作成` : "読み込み");
  ensureMedia(deck).then((added) => { if (added) { thumbCache.clear(); renderFilmstrip(); renderStage(); } });
  if (source) toast(`${deck.slides.length}枚の資料を開きました`);
}

function markChanged({ structural = false } = {}) {
  saveCurrent();
  coedit?.changed();
  scheduleMeasure();
  scheduleVersion();
  if (structural) {
    renderFilmstrip();
    renderStage();
    renderInspector();
    if (state.panel === "anim") editorUi.renderAnimPane();
    if (state.panel === "comment") comments.render();
    if (state.panel === "a11y") a11y.render();
  } else {
    // Typing: refresh the picture without rebuilding the form under the caret.
    clearTimeout(state.stageTimer);
    state.stageTimer = setTimeout(() => { renderStage(); renderFilmstrip(); }, 120);
  }
  renderIssueSummary();
  updateDesignButton();
}

function pushUndo() {
  if (!state.deck) return;
  state.undo.push(JSON.stringify({ deck: state.deck, selected: state.selected }));
  if (state.undo.length > 80) state.undo.shift();
  state.redo = [];
  updateTopbar();
}

/** Group keystrokes into one undo step (snapshot at the start of a typing burst). */
function beginEdit() {
  if (!state.editBurst) pushUndo();
  clearTimeout(state.editBurst);
  state.editBurst = setTimeout(() => { state.editBurst = null; }, 900);
}

function undoRedo(direction) {
  const from = direction === "undo" ? state.undo : state.redo;
  const to = direction === "undo" ? state.redo : state.undo;
  if (!from.length || !state.deck) return;
  to.push(JSON.stringify({ deck: state.deck, selected: state.selected }));
  const snapshot = JSON.parse(from.pop());
  state.deck = adoptDeck(snapshot.deck);
  state.selected = Math.min(snapshot.selected, state.deck.slides.length - 1);
  clearTimeout(state.editBurst);
  state.editBurst = null;
  $("deckTitleInput").value = state.deck.title;
  useFonts([state.deck.theme]);
  markChanged({ structural: true });
  updateTopbar();
}

function replaceSlide(index, slide) {
  pushUndo();
  state.deck.slides[index] = normalizeSlide(slide, index, state.deck.slides.length);
  state.selected = index;
  markChanged({ structural: true });
}

/** Returns where the slide went: a story slide never lands between a slide and its deep-dive pages. */
function insertSlide(index, slide) {
  if (state.deck.slides.length >= 500) { toast("スライドは500枚までです"); return -1; }
  if (!slide.drillOf) { const story = storyOf(); while (story.parent[index] != null) index += 1; }
  pushUndo();
  state.deck.slides.splice(index, 0, slide);
  state.selected = index;
  markChanged({ structural: true });
  return index;
}

/** Several slides after the current one in one undo step (スライドの再利用・アウトライン・フォト アルバム). */
function insertSlides(raw) {
  if (!state.deck || !raw?.length) return;
  if (state.deck.slides.length + raw.length > 500) { toast("スライドは500枚までです"); return; }
  let at = Math.min(state.selected + 1, state.deck.slides.length - 1);
  const story = storyOf();
  while (story.parent[at] != null) at += 1;
  const slides = raw.map((slide, k) => normalizeSlide(slide, at + k, state.deck.slides.length + raw.length));
  // Slides taken from another copy of this deck (a saved version, the library) bring the same ids: they get new ones,
  // and links between the slides brought in follow them.
  const used = new Set(state.deck.slides.map((slide) => slide.sid).filter(Boolean));
  const renamed = new Map();
  for (const slide of slides) {
    if (!slide.sid) continue;
    if (used.has(slide.sid)) { const sid = newSid(used); renamed.set(slide.sid, sid); slide.sid = sid; }
    used.add(slide.sid);
  }
  if (renamed.size) {
    for (const o of slides.flatMap((slide) => slide.elements || [])) {
      if (o.action?.type === "slide" && renamed.has(o.action.to)) o.action = { ...o.action, to: renamed.get(o.action.to) };
      if (o.overAction?.type === "slide" && renamed.has(o.overAction.to)) o.overAction = { ...o.overAction, to: renamed.get(o.overAction.to) };
      if (o.kind === "zoom" && renamed.has(o.target)) o.target = renamed.get(o.target);
    }
  }
  pushUndo();
  state.deck.slides.splice(at, 0, ...slides);
  state.selected = at;
  markChanged({ structural: true });
}
/** A deck to take slides from: a JSON file, or a PowerPoint file brought over as it looks. */
async function deckFromFile(file) {
  if (/\.json$/i.test(file.name) || file.type === "application/json") return adoptDeckMedia(normalizeDeck(JSON.parse(await file.text())));
  if (file.size > 30_000_000) throw new Error("ファイルが大きすぎます（30MBまで）");
  const response = await fetch(`/api/import?name=${encodeURIComponent(file.name)}&mode=exact`, { method: "POST", credentials: "same-origin", headers: { "content-type": "application/octet-stream" }, body: file });
  const result = await response.json();
  if (!response.ok) throw new Error(result.error || `HTTP ${response.status}`);
  return adoptDeckMedia(normalizeDeck(result));
}

/** Deleting a slide deletes its deep-dive pages with it. */
function deleteSlide(index) {
  const story = storyOf();
  const count = story.parent[index] == null ? groupEnd(index, story) - index : 1;
  if (state.deck.slides.length - count < 2) return toast("スライドは2枚以上必要です");
  pushUndo();
  // A section starting here goes on with the slide after the deleted ones (PowerPoint keeps the section).
  const section = state.deck.slides[index].section;
  state.deck.slides.splice(index, count);
  if (section && state.deck.slides[index] && !state.deck.slides[index].section && !state.deck.slides[index].drillOf) state.deck.slides[index].section = section;
  state.selected = Math.min(index, state.deck.slides.length - 1);
  markChanged({ structural: true });
  toast(count > 1 ? `${index + 1}枚目と、その深掘りページ${count - 1}枚を削除しました（⌘Zで元に戻せます）` : `${index + 1}枚目を削除しました（⌘Zで元に戻せます）`);
}

/** Move a slide (with its deep-dive pages) in front of the slide at `before`. Deep-dive pages move with their slide. */
function moveSlideBefore(from, before) {
  const slides = state.deck.slides;
  const last = slides.length - 1;
  const story = storyOf();
  if (story.parent[from] != null) return;
  const end = groupEnd(from, story);
  while (story.parent[before] != null) before += 1;
  if (from <= 0 || from >= last || before <= 0 || before > last || (before >= from && before <= end)) return;
  pushUndo();
  const group = slides.splice(from, end - from);
  const at = before > from ? before - group.length : before;
  slides.splice(at, 0, ...group);
  state.selected = at;
  markChanged({ structural: true });
}

const slideCount = () => state.deck?.slides.length ?? 0;
const storyOf = (deck = state.deck) => E.storyMap(deck.slides);

/** Where a slide's group (the slide and its deep-dive pages) ends. */
function groupEnd(index, story = storyOf()) {
  const head = story.parent[index] ?? index;
  let end = head + 1;
  while (story.parent[end] === head) end += 1;
  return end;
}

/** A copy of a slide as a story slide (a copied deep-dive page stands on its own). */
function copyOf(slide) {
  const copy = clone(slide);
  delete copy.drillOf;
  // A copy is another slide: links to the original keep pointing at the original; it joins the section it lands in.
  delete copy.sid;
  delete copy.section;
  return copy;
}

function select(index) {
  if (!state.deck) return;
  state.masterEdit = false;
  stopMotionPreview({ render: false });
  if (state.inline) finishInlineEdit(true);
  if (editor.typing) editor.stopTyping(true);
  state.selected = Math.max(0, Math.min(index, slideCount() - 1));
  // A slide picked inside a folded section unfolds it.
  const within = sectionList().find((sec) => state.selected >= sec.start && state.selected < sec.end);
  if (within) state.collapsedSections?.delete(within.key);
  renderFilmstrip();
  renderStage();
  renderInspector();
  if (state.panel === "comment") comments.render();
  if (state.panel === "a11y") a11y.render();
  document.querySelector(".film-item.selected")?.scrollIntoView({ block: "nearest" });
  renderChatContext();
  coedit?.sendPresence();
}

// ---------------------------------------------------------------- persistence

function formState() {
  return {
    brief: $("briefInput").value,
    audience: $("audienceInput").value,
    purpose: $("purposeInput").value,
    tone: $("toneInput").value,
    slideCount: Number($("slideCount").value) || 8,
    density: document.querySelector("input[name=density]:checked")?.value || "standard",
  };
}

// Audience and purpose: the usual choices, and "その他" for anything else. The text input holds the value either way.
const CHOICES = [["audienceSelect", "audienceInput"], ["purposeSelect", "purposeInput"]];

function showChoice(inputId) {
  const [selectId] = CHOICES.find(([, id]) => id === inputId);
  const select = $(selectId);
  const input = $(inputId);
  const known = [...select.options].some((option) => option.value === input.value && option.value !== "__other");
  select.value = known ? input.value : "__other";
  input.hidden = known;
}

function wireChoices() {
  for (const [selectId, inputId] of CHOICES) {
    $(selectId).addEventListener("change", () => {
      const input = $(inputId);
      const other = $(selectId).value === "__other";
      input.hidden = !other;
      if (other) { input.value = ""; input.focus(); } else input.value = $(selectId).value;
      input.dispatchEvent(new Event("change"));
    });
    showChoice(inputId);
  }
}

function applyForm(form = {}) {
  if (form.brief != null) $("briefInput").value = form.brief;
  if (form.audience != null) { $("audienceInput").value = form.audience; showChoice("audienceInput"); }
  if (form.purpose != null) { $("purposeInput").value = form.purpose; showChoice("purposeInput"); }
  if (form.tone) $("toneInput").value = form.tone;
  if (form.slideCount) $("slideCount").value = form.slideCount;
  if (form.density) { const radio = document.querySelector(`input[name=density][value=${form.density}]`); if (radio) radio.checked = true; }
  updateBriefCount();
}

let saveTimer = null;
function saveCurrent() {
  clearTimeout(saveTimer);
  saveTimer = setTimeout(() => {
    try {
      localStorage.setItem(STORAGE.current, JSON.stringify({ deck: state.deck, selected: state.selected, form: formState(), outline: state.outline, historyId: state.historyId, savedDeckId: state.savedDeckId, savedAt: new Date().toISOString() }));
      state.autosaveWarned = false;
    } catch {
      if (!state.autosaveWarned) toast("自動保存できませんでした（ブラウザの容量不足）。「その他」→「JSONで保存」で保存してください");
      state.autosaveWarned = true;
    }
    if (state.historyId) saveHistory({ quiet: true });
  }, 400);
}

function historyEntries() {
  try { return JSON.parse(localStorage.getItem(STORAGE.history) || "[]"); } catch { return []; }
}

function saveHistory({ exported = null, quiet = false } = {}) {
  if (!state.deck) return null;
  const entries = historyEntries();
  const id = state.historyId || `${Date.now()}`;
  const existing = entries.find((entry) => entry.id === id);
  const entry = {
    id,
    savedAt: new Date().toISOString(),
    title: state.deck.title,
    slideCount: state.deck.slides.length,
    theme: state.deck.theme,
    deck: state.deck,
    form: formState(),
    exported: exported ?? existing?.exported ?? "",
  };
  const next = [entry, ...entries.filter((item) => item.id !== id)].slice(0, 20);
  try { localStorage.setItem(STORAGE.history, JSON.stringify(next)); }
  catch { if (!quiet) toast("履歴を保存できませんでした（ブラウザの容量不足）"); }
  return id;
}

function renderHistory() {
  const list = $("historyList");
  const entries = historyEntries();
  list.replaceChildren(...(entries.length ? entries.map((entry) => h("div", { class: "history-item" },
    h("div", {}, h("b", {}, entry.title || "無題の資料"), h("span", {}, `${new Date(entry.savedAt).toLocaleString("ja-JP", { month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" })}・${entry.slideCount}枚${entry.exported ? "・HTML出力済み" : ""}`)),
    h("div", { class: "actions" },
      h("button", { class: "btn", type: "button", onclick: () => restoreHistory(entry.id) }, "開く"),
      h("button", { class: "btn btn-ghost btn-danger", type: "button", title: "削除", onclick: () => deleteHistory(entry.id) }, "削除")),
  )) : [h("div", { class: "empty-stage", style: { "min-height": "120px" } }, "まだ履歴はありません。")]));
}

function restoreHistory(id) {
  const entry = historyEntries().find((item) => item.id === id);
  if (!entry) return;
  applyForm(entry.form);
  $("historyDialog").close();
  const deck = normalizeDeck(entry.deck);
  state.historyId = entry.id;
  loadDeck(deck, {});
  toast(`「${deck.title}」を開きました`);
}

function deleteHistory(id) {
  try { localStorage.setItem(STORAGE.history, JSON.stringify(historyEntries().filter((item) => item.id !== id))); } catch { /* ignore */ }
  if (state.historyId === id) state.historyId = null;
  renderHistory();
}

// Versions: a snapshot before every AI change and at quiet moments, restorable from 履歴.
async function saveVersion(label) {
  if (!state.deck) return;
  // Slides carry their ids into the version, so 比較 can match them later.
  uniqueSids(state.deck.slides);
  const record = { id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, deck: chatKey(), at: Date.now(), label, title: state.deck.title, count: state.deck.slides.length, data: JSON.stringify(state.deck) };
  state.lastVersionAt = Date.now();
  state.lastVersionJson = record.data;
  try {
    await idb("readwrite", (store) => store.put(record), "versions");
    const old = (await listVersions()).slice(40);
    if (old.length) await idb("readwrite", (store) => { for (const item of old) store.delete(item.id); }, "versions");
  } catch { /* versions are best effort */ }
}

async function listVersions() {
  const key = chatKey();
  try {
    const all = (await idb("readonly", (store) => store.getAll(), "versions")) || [];
    return all.filter((item) => item.deck === key).sort((a, b) => b.at - a.at);
  } catch { return []; }
}

function scheduleVersion() {
  clearTimeout(state.versionTimer);
  state.versionTimer = setTimeout(() => {
    if (!state.deck) return;
    const data = JSON.stringify(state.deck);
    if (data !== state.lastVersionJson && Date.now() - (state.lastVersionAt || 0) > 3 * 60_000) saveVersion("編集の区切り");
  }, 60_000);
}

async function renderVersions() {
  const list = $("versionList");
  $("versionSection").hidden = !state.deck;
  if (!state.deck) return;
  const versions = await listVersions();
  list.replaceChildren(...(versions.length ? versions.map((version) => h("div", { class: "history-item" },
    h("div", {}, h("b", {}, version.label), h("span", {}, `${new Date(version.at).toLocaleString("ja-JP", { month: "numeric", day: "numeric", hour: "2-digit", minute: "2-digit" })}・${version.count}枚・${version.title}`)),
    h("div", { class: "actions" }, h("button", { class: "btn", type: "button", onclick: () => restoreVersion(version) }, "この版に戻す")),
  )) : [h("div", { class: "hint" }, "まだ版はありません。AIの提案を採用したときなどに自動で保存されます。")]));
}

async function restoreVersion(version) {
  await saveVersion("版を戻す前");
  pushUndo();
  state.deck = adoptDeck(JSON.parse(version.data));
  state.selected = Math.min(state.selected, state.deck.slides.length - 1);
  $("deckTitleInput").value = state.deck.title;
  $("historyDialog").close();
  useFonts([state.deck.theme]);
  markChanged({ structural: true });
  toast(`「${version.label}」の版に戻しました（⌘Zでやり直せます）`);
}

function openHistory() {
  renderHistory();
  renderVersions();
  $("historyDialog").showModal();
}

// ---------------------------------------------------------------- explicitly saved decks (separate from the rolling autosave/history)

let libraryRecords = [];

async function assertLibraryMedia(deck) {
  for (const slide of deck.slides) {
    const src = slide.media?.src;
    if (src?.startsWith("idb:") && !(await mediaBlob(src))) {
      throw new Error("この資料の写真・動画がブラウザ内に見つかりません。欠けた素材を入れ直してから保存してください。");
    }
  }
}

async function saveToLibrary({ asNew = false } = {}) {
  if (!state.deck) return;
  if (state.inline) finishInlineEdit(true);
  const button = $("saveDeckBtn");
  const newButton = $("saveAsNewBtn");
  button.disabled = true;
  newButton.disabled = true;
  try {
    uniqueSids(state.deck.slides);
    const deck = JSON.parse(JSON.stringify(state.deck));
    await assertLibraryMedia(deck);
    const existing = !asNew && state.savedDeckId ? await getSavedDeck(state.savedDeckId) : null;
    const id = existing?.id || crypto.randomUUID();
    const record = {
      id, title: deck.title, createdAt: existing?.createdAt || Date.now(), updatedAt: Date.now(),
      selected: state.selected, deck, searchIndex: buildSearchIndex(deck),
    };
    await putSavedDeck(record);
    state.savedDeckId = id;
    saveCurrent();
    updateTopbar();
    if ($("libraryDialog").open) {
      libraryRecords = [record, ...libraryRecords.filter((item) => item.id !== id)];
      renderLibrary();
    }
    toast(existing ? "資料を上書き保存しました" : "資料を保存しました");
  } catch (error) {
    toast(`保存できませんでした：${error.message}`);
  } finally {
    button.disabled = false;
    newButton.disabled = !state.deck;
  }
}

async function currentDeckCanBeReplaced() {
  if (!state.deck) return true;
  const saved = state.savedDeckId ? await getSavedDeck(state.savedDeckId).catch(() => null) : null;
  if (saved && JSON.stringify(saved.deck) === JSON.stringify(state.deck)) return true;
  return window.confirm("編集中の資料に未保存の変更があります。保存庫から別の資料を開きますか？ 必要なら先に「資料を保存」を押してください。");
}

async function openSavedDeck(record, selected = 0) {
  try {
    if (!(await currentDeckCanBeReplaced())) return;
    const latest = await getSavedDeck(record.id);
    if (!latest) throw new Error("この資料は保存庫から削除されています。");
    await assertLibraryMedia(latest.deck);
    $("libraryDialog").close();
    state.historyId = null;
    loadDeck(normalizeDeck(latest.deck), { source: "保存庫", savedDeckId: latest.id, selected });
  } catch (error) {
    showStatus("libraryStatus", `開けませんでした：${error.message}`, "error");
  }
}

async function insertSavedSlide(record, index) {
  if (!state.deck || state.deck.slides.length >= 500) return;
  try {
    const latest = await getSavedDeck(record.id);
    const source = latest?.deck.slides[index];
    if (!source) throw new Error("スライドが見つかりません。");
    await assertLibraryMedia({ slides: [source] });
    const at = Math.min(state.selected + 1, state.deck.slides.length - 1);
    insertSlide(at, copyOf(source));
    $("libraryDialog").close();
    ensureMedia({ slides: [source] }).then((added) => { if (added) { thumbCache.clear(); renderFilmstrip(); renderStage(); } });
    toast(`${index + 1}枚目を現在の資料に追加しました`);
  } catch (error) {
    showStatus("libraryStatus", `追加できませんでした：${error.message}`, "error");
  }
}

async function removeSavedDeck(record) {
  if (!window.confirm(`「${record.title}」を保存庫から削除しますか？ この操作は取り消せません。`)) return;
  try {
    await deleteSavedDeck(record.id);
    libraryRecords = libraryRecords.filter((item) => item.id !== record.id);
    if (state.savedDeckId === record.id) { state.savedDeckId = null; saveCurrent(); updateTopbar(); }
    renderLibrary();
  } catch (error) {
    showStatus("libraryStatus", `削除できませんでした：${error.message}`, "error");
  }
}

function renderLibrary() {
  const query = $("librarySearch").value.trim();
  const matches = searchSavedDecks(libraryRecords, query);
  $("saveAsNewBtn").disabled = !state.deck;
  showStatus("libraryStatus", query ? `${matches.length}件の資料が見つかりました` : `${matches.length}件の資料を保存しています`);
  $("libraryList").replaceChildren(...(matches.length ? matches.map(({ record, slideMatches }) => {
    const slideRows = slideMatches.map((index) => {
      const slide = record.deck.slides[index];
      return h("div", { class: "library-slide" },
        h("div", { class: "text" }, h("b", {}, `${index + 1}枚目｜${slide.title || typeLabel(slide.type)}`), h("span", {}, slideExcerpt(slide, query))),
        h("div", { class: "library-slide-actions" },
          h("button", { class: "btn", type: "button", onclick: () => openSavedDeck(record, index) }, "この1枚で開く"),
          h("button", { class: "btn", type: "button", disabled: !state.deck || state.deck.slides.length >= 500, onclick: () => insertSavedSlide(record, index) }, "今の資料に追加")));
    });
    return h("div", { class: "library-card" },
      h("div", { class: "library-card-head" },
        h("div", {}, h("b", {}, record.title || "無題の資料"), h("span", { class: "hint" }, `${record.deck.slides.length}枚・${new Date(record.updatedAt).toLocaleString("ja-JP")}`)),
        h("div", { class: "library-actions" },
          h("button", { class: "btn btn-primary", type: "button", onclick: () => openSavedDeck(record) }, "資料全体を開く"),
          h("button", { class: "btn btn-ghost btn-danger", type: "button", onclick: () => removeSavedDeck(record) }, "削除"))),
      h("details", { open: Boolean(query && slideRows.length) }, h("summary", {}, `スライドを表示（${slideRows.length}枚）`),
        slideRows.length ? slideRows : h("p", { class: "hint" }, "資料内の複数箇所にキーワードが見つかりました。")));
  }) : [h("p", { class: "hint" }, query ? "一致する資料はありません。別のキーワードを試してください。" : "まだ保存した資料はありません。編集中の資料で「資料を保存」を押してください。")]));
}

async function openLibrary() {
  $("librarySearch").value = "";
  $("libraryDialog").showModal();
  $("librarySearch").focus();
  try {
    libraryRecords = await listSavedDecks();
    renderLibrary();
  } catch (error) {
    showStatus("libraryStatus", `保存庫を開けませんでした：${error.message}`, "error");
  }
}

// ---------------------------------------------------------------- deck lint ("構成チェック")

const PLACEHOLDERS = new Set([
  "スライドタイトル", "このスライドで伝える結論", "具体的な説明", "説明を入力", "新しい項目", "要点A", "要点B", "要点C", "要点D", "要点E", "要点F",
  "結論を支える根拠", "実行時の重要ポイント", "期待できる効果", "誰が・いつまでに・何をするか", "資料タイトル", "章タイトル", "引用したい声やメッセージ", "発言者",
  "指標A", "指標B", "指標C", "指標", "—", "質問", "回答", "グラフタイトル", "図から読み取れる要点", "意思決定への示唆", "結論を一文で", "お願いしたいこと",
  "3つの要点から導くまとめ", "タスク", "確認項目", "段1", "段2", "段3", "出来事", "時期", "担当", "作業", "要素", "現状の課題", "目指す状態", "変える前", "変えた後",
  "工程", "項目", "補足", "説明", "伝えたいひと言", "写真に重ねて見せる補足の一文", "ポイント", "いちばん伝えたい**ひと言**を大きく",
]);
const FULLWIDTH_NUMBER = /[０-９％．，]/;
const NON_TEXT_KEYS = new Set(["type", "visualAsset", "imagePosition", "state", "trend", "status", "chartType", "customImage", "icon", "animation", "photoMotion", "kinetic", "backdrop", "entrance", "emphasis", "transition", "media", "imagePlacement", "target", "notes", "formula", "elements", "timeline", "sid", "transitionDur", "transitionSound", "transitionDir", "advance", "background"]);

function textEntries(value, path = [], out = []) {
  if (typeof value === "string") out.push([path, value]);
  else if (Array.isArray(value)) value.forEach((item, i) => textEntries(item, [...path, i], out));
  else if (value && typeof value === "object") {
    for (const [key, child] of Object.entries(value)) if (!NON_TEXT_KEYS.has(key) && key !== "details") textEntries(child, [...path, key], out);
  }
  return out;
}

function eachText(value, visit, key = "") {
  if (typeof value === "string") return NON_TEXT_KEYS.has(key) && key !== "notes" || value.startsWith("data:") || value.startsWith("idb:") ? value : visit(value);
  if (Array.isArray(value)) return value.map((item) => eachText(item, visit, key));
  // Objects keep their words in rich text: they are visited by eachObjectText, never as markup.
  if (value && typeof value === "object") return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, ["media", "elements", "timeline", "sid", "readingOrder"].includes(k) ? v : eachText(v, visit, k)]));
  return value;
}

/** Visit the words of every object's text (not its markup); returns how many objects changed. */
function eachObjectText(slides, visit) {
  let changed = 0;
  // One rich text: its words visited, the markup kept. Returns the new text, or null when nothing changed.
  const visitRich = (html) => {
    const box = document.createElement("div");
    box.append(E.richFragment(html));
    const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
    let dirty = false;
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const next = visit(node.data);
      if (next !== node.data) { node.data = next; dirty = true; }
    }
    return dirty ? E.sanitizeRich(box.innerHTML) : null;
  };
  for (const slide of slides) {
    for (const o of slide.elements || []) {
      let dirty = false;
      if (o?.text) { const next = visitRich(o.text); if (next != null) { o.text = next; dirty = true; } }
      // A table's cells and a chart's words are found and replaced too.
      if (o?.kind === "table") for (const cell of o.cells.flat()) if (cell.text) { const next = visitRich(cell.text); if (next != null) { cell.text = next; dirty = true; } }
      if (o?.kind === "chart") {
        if (o.chart.title) { const next = visit(o.chart.title); if (next !== o.chart.title) { o.chart.title = next; dirty = true; } }
        o.chart.labels = o.chart.labels.map((label) => { const next = visit(label); if (next !== label) dirty = true; return next; });
        for (const serie of o.chart.series) { const next = visit(serie.name); if (next !== serie.name) { serie.name = next; dirty = true; } }
      }
      if (dirty) changed += 1;
    }
  }
  return changed;
}

function normalizeNumbers() {
  pushUndo();
  const fix = (text) => text.replace(/[０-９]/g, (d) => String.fromCharCode(d.charCodeAt(0) - 0xfee0)).replace(/％/g, "%").replace(/(?<=\d)．(?=\d)/g, ".").replace(/(?<=\d)，(?=\d{3})/g, ",");
  state.deck.slides = eachText(state.deck.slides, fix);
  markChanged({ structural: true });
  toast("数字を半角に揃えました（⌘Zで元に戻せます）");
}

let lintMemo = { key: "", result: [] };
function lintDeck() {
  const deck = state.deck;
  if (!deck) return [];
  const key = JSON.stringify(deck.slides);
  if (lintMemo.key === key) return lintMemo.result;
  const result = deck.slides.map(() => []);
  const add = (i, field, message, severity = "warning", fix = null, plain = false) => result[i].push({ slide: i, field, kind: "lint", severity, message, fix, plain });
  const titles = new Map();
  const story = E.storyMap(deck.slides);
  const itemsOf = new Map();
  deck.slides.forEach((slide, i) => {
    const parent = story.parent[i];
    if (parent != null) {
      if (!itemsOf.has(parent)) itemsOf.set(parent, new Set(itemKeys(deck.slides[parent], parent).map((entry) => entry.key)));
      if (!itemsOf.get(parent).has(slide.drillOf)) add(i, "title", `${parent + 1}枚目に、このページを開く項目が見つかりません（項目が変わった可能性があります）。右の「深掘りページ」で開く項目を選び直してください`, "warning", null, true);
    }
    const texts = textEntries(slide);
    // Table headers ("項目") and "—" in a table cell are real content; so is a title like "お願いしたいこと".
    const isPlaceholder = ([path, text]) => {
      const value = text.trim();
      if (/【[^】]+】/.test(value)) return true;
      if (!PLACEHOLDERS.has(value)) return false;
      return !(path[0] === "headers" || (path[0] === "rows" && value === "—") || (path[0] === "title" && value === "お願いしたいこと"));
    };
    const placeholders = texts.filter(isPlaceholder);
    placeholders.slice(0, 3).forEach(([path]) => add(i, pathKey(path), "仮の文言のままです"));
    if (placeholders.length > 3) add(i, pathKey(placeholders[3][0]), `ほかにも仮の文言が${placeholders.length - 3}か所あります`, "warning", null, true);
    if (TITLED(slide.type) && !["hero", "statement", "blank"].includes(slide.type) && !strip(slide.takeaway)) add(i, "takeaway", "キーメッセージ（このスライドの結論）がありません");
    if (slide.type === "closing" && !strip(slide.message)) add(i, "message", "次のアクションが書かれていません");
    if (slide.type === "hero" && !slide.media && !slide.visualAsset && !slide.customImage) add(i, "title", "全面写真のスライドに写真がありません（「編集」→「写真・動画」で選べます）", "warning", null, true);
    if (slide.media?.src?.startsWith("idb:") && !mediaUrls[slide.media.src] && state.mediaChecked) add(i, "title", "このブラウザに動画・写真のデータがありません（別のPCで保存した資料です）", "warning", null, true);
    if (texts.some(([, text]) => FULLWIDTH_NUMBER.test(text))) add(i, "title", "全角の数字・記号があります（半角に揃えると読みやすくなります）", "info", normalizeNumbers, true);
    const chars = texts.filter(([path]) => !["title", "takeaway", "subhead"].includes(String(path[0]))).reduce((sum, [, text]) => sum + strip(text).length, 0);
    const limit = slide.type === "table" ? 320 : 240;
    if (TITLED(slide.type) && chars > limit) {
      add(i, "title", `本文が${chars}字あります。1枚${limit}字までを目安に、要点を絞るか、補足は「クリックで開く詳細」に回すと伝わりやすくなります`, "warning", null, true);
    }
    const title = strip(slide.title);
    if (title && slide.type !== "title") titles.set(title, [...(titles.get(title) ?? []), i]);
    if (TITLED(slide.type) && charCount(title) > 30) add(i, "title", `タイトルが${Math.ceil(charCount(title))}字あります。論点を30字以内にすると2行に収まります`, "info", null, true);
    if (TITLED(slide.type) && title && strip(slide.takeaway) === title) add(i, "takeaway", "キーメッセージがタイトルと同じです。タイトルは論点、キーメッセージは結論の一文にします", "info");
    if ((String(slide.takeaway ?? "").match(/\*\*[^*]+\*\*/g) ?? []).length > 2) add(i, "takeaway", "強調（**〜**）は1〜2か所までにすると、伝えたい語句が目立ちます", "info");
    // Layouts that look alike are caught by look below; this catches three of the same chart or shape.
    if (i >= 2 && !LOOKS[lookOf(slide)]?.alike && slide.type !== "section" && deck.slides[i - 1].type === slide.type && deck.slides[i - 2].type === slide.type) {
      add(i, "title", `「${typeLabel(slide.type)}」が3枚続いています。別のレイアウトにすると単調になりません`, "info", null, true);
    }
  });
  for (const issue of varietyIssues(deck.slides)) {
    add(issue.slide, "title", issue.message, issue.rule === "adjacent" ? "info" : "warning", null, true);
    result[issue.slide].at(-1).ai = varietyRequest(issue);
  }
  for (const [title, indices] of titles) {
    if (indices.length < 2) continue;
    for (const i of indices) add(i, "title", `「${title}」と同じタイトルのスライドがあります（${indices.filter((j) => j !== i).map((j) => `${j + 1}枚目`).join("・")}）`, "info", null, true);
  }
  lintMemo = { key, result };
  return result;
}

/** What to ask the AI when slides look alike (the answer comes back as a proposal to compare and accept). */
function varietyRequest(issue) {
  const at = issue.slides.map((i) => `@${i + 1}`).join(" ");
  const how = `内容に合う別の見せ方（${LOOK_ADVICE.join("／")}）`;
  if (issue.rule === "crowded") return `${at} はどれも「${LOOKS[issue.look].label}」の見た目で単調です。内容に最も合う${issue.keep}枚だけ残し、ほかは${how}に変えてください。主張と数値は変えないでください`;
  if (issue.rule === "adjacent") return `${at} が同じ「${LOOKS[issue.look].label}」の見た目で続いています。内容に合わないほうの1枚を、${how}に変えてください。主張と数値は変えないでください`;
  return `本文の見た目が単調です。数字・図形・大きな一文など、内容に合う別の見せ方を取り入れて、見た目を4種類以上にしてください（${LOOK_ADVICE.join("／")}）。主張と数値は変えないでください`;
}

function overflowFor(index) {
  return fitFor(index)?.issues ?? [];
}

function issuesFor(index) {
  return [...overflowFor(index), ...(lintDeck()[index] ?? [])];
}

// ---------------------------------------------------------------- views

function setMode(mode) {
  if (mode !== "edit" && state.inline) finishInlineEdit(true);
  state.mode = mode;
  $("createView").classList.toggle("hidden", mode !== "create");
  $("editView").classList.toggle("hidden", mode !== "edit");
  $("createTab").setAttribute("aria-selected", String(mode === "create"));
  $("editTab").setAttribute("aria-selected", String(mode === "edit"));
  updateTopbar();
}

function updateTopbar() {
  const editing = state.mode === "edit" && Boolean(state.deck);
  for (const id of ["deckTitleInput", "saveDeckBtn", "shareBtn", "undoBtn", "redoBtn", "presentBtn", "pdfBtn", "pptxBtn", "downloadBtn", "moreMenu"]) $(id).classList.toggle("hidden", !editing);
  $("saveDeckBtn").textContent = state.savedDeckId ? "上書き保存" : "資料を保存";
  $("deckReviseBtn").classList.toggle("hidden", !editing || !state.codexAuthorized);
  $("deckReviseBtn").disabled = state.aiBusy;
  $("undoBtn").disabled = !state.undo.length;
  $("redoBtn").disabled = !state.redo.length;
  $("downloadBtn").disabled = Boolean(state.exporting);
  $("pptxBtn").disabled = Boolean(state.exporting);
}

function renderAll() {
  updateTopbar();
  updateDesignButton();
  renderChatContext();
  renderFilmstrip();
  renderStage();
  renderInspector();
  renderIssueSummary();
}

function updateDesignButton() {
  if (!state.deck) return;
  const meta = themeMeta(state.deck.theme);
  const colors = [meta.swatch[0], meta.swatch[2], meta.swatch[3], meta.swatch[1]];
  $("designSwatch").replaceChildren(...colors.map((color) => h("i", { style: { background: color } })));
  $("designLabel").textContent = `${meta.name}・動き`;
}

let dragFrom = null;
function renderFilmstrip() {
  const strip_ = $("filmstrip");
  if (!state.deck) { strip_.replaceChildren(); return; }
  const last = state.deck.slides.length - 1;
  const story = storyOf();
  const items = state.deck.slides.map((slide, index) => {
    const issues = issuesFor(index).filter((issue) => issue.severity !== "info");
    const errors = issues.filter((issue) => issue.severity === "error").length;
    const parent = story.parent[index];
    const movable = index > 0 && index < last && parent == null;
    const flags = [];
    if (slide.media?.kind === "video" || E.youtubeId(slide.media?.src)) flags.push(h("span", { title: "動画あり" }, "▶"));
    if (slide.media?.kind === "lottie") flags.push(h("span", { title: "アニメーション（Lottie）あり" }, "✦"));
    if ((slide.elements || []).some((o) => o.kind === "audio")) flags.push(h("span", { title: "オーディオあり" }, "♪"));
    if (E.backdropOf(slide, slide.type, state.deck.motion) || (slide.kinetic && slide.kinetic !== "none")) flags.push(h("span", { title: "モーショングラフィックあり" }, "◎"));
    if (slide.details?.length) flags.push(h("span", { title: "クリックで開く詳細あり" }, "＋"));
    if (story.drills[index]?.length) flags.push(h("span", { title: `クリックで移る深掘りページ ${story.drills[index].length}枚` }, `↗${story.drills[index].length}`));
    const build = slide.animation || E.recommendedBuild(slide.type);
    if (build === "click") flags.push(h("span", { title: "クリックで順番に表示" }, "⋯"));
    if (slide.hidden) flags.push(h("span", { title: "非表示スライド（発表では飛ばします）" }, "非"));
    const open = (slide.comments || []).filter((c) => !c.done).length;
    if (open) flags.push(h("span", { class: "flag-comment", title: `コメント${open}件（校閲）` }, `💬${open}`));
    const item = h("div", {
      "data-index": String(index),
      class: `film-item${index === state.selected ? " selected" : ""}${parent != null ? " is-drill" : ""}${slide.hidden ? " is-hidden" : ""}`,
      draggable: movable ? "true" : null,
      title: `${index + 1}. ${parent != null ? `${parent + 1}枚目の深掘りページ：` : ""}${strip(slide.title) || typeLabel(slide.type)}${slide.hidden ? "（発表では非表示）" : ""}`,
      onclick: () => { select(index); $("filmstrip").focus({ preventScroll: true }); },
      oncontextmenu: (event) => { event.preventDefault(); select(index); slideMenu(index, event.clientX, event.clientY); },
      ondragstart: (event) => { dragFrom = index; event.dataTransfer.effectAllowed = "move"; item.classList.add("dragging"); },
      ondragend: () => { dragFrom = null; item.classList.remove("dragging"); strip_.querySelectorAll(".drop-before").forEach((el) => el.classList.remove("drop-before")); },
      ondragover: (event) => { if (dragFrom != null && index > 0 && index <= last) { event.preventDefault(); item.classList.add("drop-before"); } },
      ondragleave: () => item.classList.remove("drop-before"),
      ondrop: (event) => { event.preventDefault(); if (dragFrom != null) moveSlideBefore(dragFrom, index); },
    }, h("div", { class: "film-no" }, parent != null ? h("span", { title: `${index + 1}枚目（深掘りページ）` }, "↳") : index + 1), thumb(index, "film"),
    issues.length ? h("span", { class: `film-badge${errors ? " error" : ""}`, title: `${issues.length}件の注意` }, issues.length) : null,
    flags.length ? h("span", { class: "film-flags" }, flags) : null);
    return item;
  });
  // PowerPoint's sections: a header before the first slide of each, folding its slides away.
  const sections = sectionList();
  state.collapsedSections ||= new Set();
  const out = [];
  items.forEach((item, index) => {
    const head = sections.find((sec) => sec.start === index);
    if (head) out.push(sectionHeader(head));
    const within = sections.find((sec) => index >= sec.start && index < sec.end);
    if (!within || !state.collapsedSections.has(within.key)) out.push(item);
  });
  strip_.replaceChildren(...out, h("button", { class: "btn film-add", type: "button", onclick: () => openTypeDialog("insert") }, "＋ スライドを追加"));
  coedit?.decorateFilm(strip_.querySelectorAll(".film-item"));
}

// ---------------------------------------------------------------- sections (PowerPoint's, in the thumbnails)

/** The deck's sections, [{ start, end, name, key, implicit }]: none until a slide starts one. */
function sectionList(slides = state.deck?.slides || []) {
  const starts = slides.map((slide, i) => (slide.section ? i : -1)).filter((i) => i >= 0);
  if (!starts.length) return [];
  const list = [];
  if (starts[0] > 0) list.push({ start: 0, end: starts[0], name: "既定のセクション", implicit: true });
  starts.forEach((start, k) => list.push({ start, end: starts[k + 1] ?? slides.length, name: slides[start].section }));
  return list.map((sec) => ({ ...sec, key: `${sec.name}@${slides[sec.start].sid || sec.start}` }));
}
function sectionHeader(sec) {
  const collapsed = state.collapsedSections.has(sec.key);
  const count = sec.end - sec.start;
  return h("div", { class: `film-section${collapsed ? " collapsed" : ""}`, role: "button", "aria-expanded": String(!collapsed),
    title: `${sec.name}（${count}枚）：クリックで${collapsed ? "展開" : "折りたたむ"}・右クリックでメニュー`,
    onclick: () => { if (collapsed) state.collapsedSections.delete(sec.key); else state.collapsedSections.add(sec.key); renderFilmstrip(); },
    oncontextmenu: (event) => { event.preventDefault(); sectionMenu(sec, event.clientX, event.clientY); } },
  h("span", { class: "film-section-caret", "aria-hidden": "true" }, collapsed ? "▸" : "▾"), h("b", {}, sec.name), h("small", {}, String(count)));
}
function sectionMenu(sec, x, y) {
  editor.openMenu(x, y, [
    !sec.implicit && { label: "セクション名の変更…", run: () => renameSection(sec.start) },
    !sec.implicit && { label: "セクションの削除（スライドは残す）", run: () => removeSection(sec.start) },
    { label: "セクションを上へ移動", run: () => moveSection(sec, -1) },
    { label: "セクションを下へ移動", run: () => moveSection(sec, 1) },
    "-",
    { label: "すべて折りたたむ", run: () => { for (const s of sectionList()) state.collapsedSections.add(s.key); renderFilmstrip(); } },
    { label: "すべて展開", run: () => { state.collapsedSections.clear(); renderFilmstrip(); } },
  ].filter(Boolean));
}
/** A section starts at this slide (PowerPoint's セクションの追加). */
async function addSection(index = state.selected) {
  const slide = state.deck?.slides[index];
  if (!slide) return;
  if (slide.drillOf) return toast("深掘りページからはセクションを始められません（元のスライドで追加してください）");
  const name = await askDialog("セクションの追加", "セクション名", slide.section || `セクション${sectionList().filter((sec) => !sec.implicit).length + 1}`);
  if (name == null || !name.trim()) return;
  pushUndo();
  slide.section = name.trim().slice(0, 40);
  markChanged({ structural: true });
}
async function renameSection(start) {
  const slide = state.deck.slides[start];
  const name = await askDialog("セクション名の変更", "セクション名", slide.section || "");
  if (name == null || !name.trim()) return;
  pushUndo();
  slide.section = name.trim().slice(0, 40);
  markChanged({ structural: true });
}
function removeSection(start) {
  pushUndo();
  delete state.deck.slides[start].section;
  markChanged({ structural: true });
}
/**
 * Swap a section with the one before or after it. The cover stays first (the section holding it stays in front) and
 * the last slide stays last: a section ending with it moves without it.
 */
function moveSection(sec, dir) {
  const list = sectionList();
  const k = list.findIndex((x) => x.start === sec.start);
  const other = list[k + dir];
  if (!other) return toast("これ以上動かせません");
  const slides = state.deck.slides;
  const last = slides.length - 1;
  const [a, b] = dir < 0 ? [other, sec] : [sec, other];
  if (a.start === 0) return toast("表紙のあるセクションは先頭のままです");
  const end = Math.min(b.end, last);
  if (end <= b.start) return toast("これ以上動かせません");
  pushUndo();
  const first = slides.slice(a.start, a.end);
  const second = slides.slice(b.start, end);
  slides.splice(a.start, end - a.start, ...second, ...first);
  state.selected = dir < 0 ? a.start : a.start + second.length;
  markChanged({ structural: true });
}

// ---------------------------------------------------------------- 開発: the slide's code (JSON to edit, the HTML it becomes)

/** 開発 → スライドのJSON / HTMLを表示: the slide on the stage as data (edited and checked before it applies) or as HTML. */
function openCodeDialog(kind) {
  const slide = state.deck?.slides[state.selected];
  if (!slide) return;
  if (state.inline) finishInlineEdit(true);
  const json = kind === "json";
  $("codeTitle").textContent = json ? `スライド ${state.selected + 1} のJSON（編集できます）` : `スライド ${state.selected + 1} のHTML`;
  const text = $("codeText");
  text.readOnly = !json;
  text.value = json ? JSON.stringify(slide, null, 2) : prettyHtml(E.render(slide, renderOptions({ index: state.selected, mode: "present", fit: fitFor(state.selected) ?? undefined })).outerHTML);
  $("codeApply").hidden = !json;
  $("codeStatus").className = "hint";
  $("codeStatus").textContent = json
    ? "レイアウト（type）・文字・部品（elements）・アニメーション（timeline）・インタラクション（hover・action・loop）を直して「適用」。形式を確かめてから反映し、⌘Zで戻せます。"
    : "エンジンが描くこのスライドのHTMLです。書き出したHTMLファイルでも、これがエンジンのCSSとJavaScriptで動きます。";
  $("codeDialog").dataset.kind = kind;
  $("codeDialog").showModal();
}
function applyCodeDialog() {
  const index = state.selected;
  const status = $("codeStatus");
  let data;
  try { data = JSON.parse($("codeText").value); } catch (error) { status.className = "hint error"; status.textContent = `JSONとして読めません：${error.message}`; return; }
  if (!data || typeof data !== "object" || Array.isArray(data)) { status.className = "hint error"; status.textContent = "1枚のスライドのJSON（{ \"type\": … } のオブジェクト）にしてください"; return; }
  const before = state.deck.slides[index];
  if (!data.sid && before.sid) data.sid = before.sid;
  let next;
  try { next = normalizeSlide(data, index, state.deck.slides.length); } catch (error) { status.className = "hint error"; status.textContent = error.message; return; }
  pushUndo();
  state.deck.slides[index] = next;
  markChanged({ structural: true });
  $("codeDialog").close();
  toast("JSONをスライドに反映しました（⌘Zで戻せます）");
}
/** Line breaks and indents for reading (the HTML itself is not changed). */
function prettyHtml(html) {
  const VOID = /^<(br|img|input|meta|link|hr|source|path|circle|rect|line|polyline|polygon|ellipse|stop|use|wbr)\b/i;
  let depth = 0;
  return html.replace(/></g, ">\n<").split("\n").map((line) => {
    if (/^<\//.test(line)) depth = Math.max(0, depth - 1);
    const out = `${"  ".repeat(depth)}${line}`;
    if (/^<[a-zA-Z]/.test(line) && !VOID.test(line) && !/\/>$/.test(line) && !/<\/[^>]+>$/.test(line)) depth += 1;
    return out;
  }).join("\n");
}

// ---------------------------------------------------------------- the thumbnails' menu and keys (PowerPoint's)

/** Right-click on a thumbnail: what PowerPoint offers there. */
function slideMenu(index, x, y) {
  const slide = state.deck?.slides[index];
  if (!slide) return;
  const last = state.deck.slides.length - 1;
  editor.openMenu(x, y, [
    { label: "切り取り", keys: "⌘X", run: () => cutSlide(index) },
    { label: "コピー", keys: "⌘C", run: () => copySlide(index) },
    state.slideClipboard && { label: "貼り付け", keys: "⌘V", run: () => pasteSlide(index) },
    "-",
    { label: "新しいスライド…", keys: "⌘M", run: () => openTypeDialog("insert") },
    { label: "白紙のスライドを追加", run: () => editorApp.insertBlankSlide() },
    index > 0 && index < last && { label: "スライドの複製", keys: "⌘D", run: () => editorApp.duplicateSlide() },
    { label: "スライドの削除", keys: "Delete", run: () => deleteSlide(index) },
    { label: "上へ移動", keys: "⌘↑", run: () => moveSlideStep(index, -1) },
    { label: "下へ移動", keys: "⌘↓", run: () => moveSlideStep(index, 1) },
    "-",
    index > 0 && { label: slide.hidden ? "非表示スライドを解除" : "非表示スライドに設定", run: () => toggleHiddenSlide() },
    !slide.drillOf && { label: slide.section ? "セクション名の変更…" : "セクションの追加…", run: () => (slide.section ? renameSection(index) : addSection(index)) },
    { label: "レイアウトの変更…", run: () => openTypeDialog("change") },
    "-",
    { label: "このスライドから発表", keys: "⇧F5", run: () => openPresenter(index) },
    { label: "動きを確認", run: () => { setView("single"); previewMotion(); } },
  ].filter(Boolean));
}
function copySlide(index) {
  state.slideClipboard = clone(state.deck.slides[index]);
  toast(`${index + 1}枚目をコピーしました（サムネイルで ⌘V で貼り付け）`);
}
function cutSlide(index) {
  state.slideClipboard = clone(state.deck.slides[index]);
  deleteSlide(index);
}
/** Pastes the copied slide after `index` (a copy is a slide of its own: links to the original stay with it). */
function pasteSlide(index = state.selected) {
  if (!state.slideClipboard) return toast("貼り付けるスライドがありません（サムネイルで ⌘C でコピーします）");
  const at = Math.min(index + 1, state.deck.slides.length - 1);
  insertSlide(at, normalizeSlide(copyOf(state.slideClipboard), at, state.deck.slides.length + 1));
}
/** Keys on the thumbnails (when they have the focus), as in PowerPoint. Returns true when taken. */
function filmKey(event) {
  const meta = event.metaKey || event.ctrlKey;
  const key = event.key.toLowerCase();
  if (!meta && (event.key === "Delete" || event.key === "Backspace")) { deleteSlide(state.selected); return true; }
  if (!meta && event.key === "Enter") { openTypeDialog("insert"); return true; }
  if (meta && key === "c") { copySlide(state.selected); return true; }
  if (meta && key === "x") { cutSlide(state.selected); return true; }
  if (meta && key === "v") { pasteSlide(); return true; }
  if (meta && key === "d") { editorApp.duplicateSlide(); return true; }
  // ⌘↑ / ⌘↓: move the slide (with its deep-dive pages) one place up or down, as in PowerPoint.
  if (meta && (event.key === "ArrowUp" || event.key === "ArrowDown")) { moveSlideStep(state.selected, event.key === "ArrowUp" ? -1 : 1); return true; }
  return false;
}

/** One place up or down in the show (a deep-dive page moves with its slide; the cover and the last slide stay). */
function moveSlideStep(index, dir) {
  const story = storyOf();
  const slides = state.deck.slides;
  if (story.parent[index] != null) return;
  if (dir < 0) {
    let before = index - 1;
    while (before > 0 && story.parent[before] != null) before -= 1;
    moveSlideBefore(index, before);
  } else {
    const end = groupEnd(index, story);
    if (end >= slides.length - 1) return;
    moveSlideBefore(index, groupEnd(end, story));
  }
}

function renderStage() {
  const body = $("stageBody");
  if (state.inline && body.contains(state.inline.el)) { state.stageDirty = true; return; }
  if (editor.typing) { state.stageDirty = true; return; }
  editorUi.renderRibbon();
  if (state.motionPreview) return;
  E.stopLottie(body);
  E.stopModels?.(body);
  if (!state.deck) {
    body.replaceChildren(h("div", { class: "empty-stage" }, h("div", {}, h("strong", {}, "まだ資料がありません"), "「作成」から構成を作るか、雛形・サンプル・JSONから始めてください。")));
    return;
  }
  const deck = state.deck;
  const types = new Set(deck.slides.map((slide) => slide.type)).size;
  const videos = deck.slides.filter((slide) => slide.media?.kind === "video" || E.youtubeId(slide.media?.src)).length;
  const lotties = deck.slides.filter((slide) => slide.media?.kind === "lottie").length;
  const details = deck.slides.reduce((sum, slide) => sum + (slide.details?.length ?? 0), 0);
  const story = storyOf();
  const drillCount = Object.keys(story.parent).length;
  const hiddenCount = deck.slides.filter((slide) => slide.hidden).length;
  $("deckMeta").textContent = [deck.audience && `対象：${deck.audience}`, `${story.order.length}枚${drillCount ? `＋深掘り${drillCount}枚` : ""}${hiddenCount ? `＋非表示${hiddenCount}枚` : ""}`, `${types}種類のレイアウト`, videos ? `動画${videos}本` : "", lotties ? `アニメーション${lotties}個` : "", details ? `詳細${details}か所` : ""].filter(Boolean).join(" ・ ");
  // 最終版: the slide and the panes take no edits; the message bar says so and offers 編集する.
  const final = isFinal(state.deck) && state.mode === "edit";
  document.body.classList.toggle("is-final", final);
  for (const id of ["notesInput", "deckTitleInput", "chatInput", "memoInput"]) { const el = $(id); if (!el) continue; if (el.hasAttribute("contenteditable")) el.contentEditable = String(!final); else el.readOnly = final; }
  shell.renderMessage(masterCallout() || finalCallout() || importCallout());
  shell.renderStatus();
  if (state.view === "outline") return renderOutline(body);
  if (state.view === "grid") {
    // スライド一覧 (PowerPoint's slide sorter): the sections' names over their slides; drag a slide to move it.
    const sections = sectionList();
    const last = deck.slides.length - 1;
    const cells = [];
    deck.slides.forEach((slide, index) => {
      const head = sections.find((sec) => sec.start === index);
      if (head) cells.push(h("div", { class: "grid-section" }, h("b", {}, head.name), h("small", {}, `${head.end - head.start}枚`)));
      const movable = index > 0 && index < last && story.parent[index] == null;
      const cell = h("div", {
        class: `grid-item${index === state.selected ? " selected" : ""}${slide.hidden ? " is-hidden" : ""}`,
        draggable: movable ? "true" : null,
        onclick: () => { setView("single"); select(index); },
        oncontextmenu: (event) => { event.preventDefault(); select(index); slideMenu(index, event.clientX, event.clientY); },
        ondragstart: (event) => { dragFrom = index; event.dataTransfer.effectAllowed = "move"; cell.classList.add("dragging"); },
        ondragend: () => { dragFrom = null; cell.classList.remove("dragging"); },
        ondragover: (event) => { if (dragFrom != null && index > 0 && index <= last) { event.preventDefault(); cell.classList.add("drop-before"); } },
        ondragleave: () => cell.classList.remove("drop-before"),
        ondrop: (event) => { event.preventDefault(); cell.classList.remove("drop-before"); if (dragFrom != null) moveSlideBefore(dragFrom, index); },
      }, thumb(index, "grid"), h("div", { class: "grid-label" }, h("span", {}, story.parent[index] != null ? `↳ ${story.parent[index] + 1}枚目の深掘り・${typeLabel(slide.type)}` : `${index + 1}. ${typeLabel(slide.type)}${slide.hidden ? "（非表示）" : ""}`), issuesFor(index).length ? h("span", { style: { color: "var(--warn)" } }, `注意${issuesFor(index).length}`) : null));
      cells.push(cell);
    });
    body.replaceChildren(h("div", { class: "stage-grid" }, cells));
    return;
  }
  const index = state.selected;
  const master = state.masterEdit;
  const slide = master ? masterSlide() : deck.slides[index];
  const el = E.render(slide, renderOptions({ index, mode: "edit", fit: master ? undefined : fitFor(index) ?? undefined }));
  el.classList.add("hs-static");
  const wrap = h("div", { class: "slide-wrap", ondragover: onStageDragOver, ondrop: onStageDrop }, E.mount(el));
  // The slide alone in the middle of the stage; where it is, its notes and the zoom are around it (shell.mjs).
  body.replaceChildren(h("div", { class: "stage-single" }, wrap));
  // Mark what does not fit on the slide itself (red dashed outline) and keep the measurement in step.
  if (body.offsetParent && !master) {
    const result = E.fit(el);
    fitState.byKey.set(slideKey(index), { fs: result.fs, ts: result.ts, ...(result.objs ? { objs: result.objs } : {}), issues: result.issues.map((issue) => ({ ...issue, slide: index })) });
  }
  wirePlacedMedia(el, slide);
  // The editor stays still: a Lottie animation shows one frame from its middle; 3D models are drawn live.
  E.mountLottie(el, { play: false, frame: 0.5 });
  E.mountModels?.(el);
  editor.attach(wrap, el, index);
}

/** The status bar's 「スライド 3 / 15」 and what the slide is (a deep-dive page, hidden, its layout and build). */
function slideStatus(index = state.selected) {
  if (state.masterEdit && state.deck) return { pos: "スライド マスター", detail: `ここに置いた図形・文字・図は全スライド（${state.deck.slides.length}枚）に出ます` };
  const deck = state.deck;
  const slide = deck?.slides[index];
  if (!slide) return null;
  const story = storyOf();
  const build = slide.animation || E.recommendedBuild(slide.type);
  const detail = [
    story.parent[index] != null ? `${story.parent[index] + 1}枚目の深掘りページ` : null,
    slide.hidden ? "非表示スライド" : null,
    typeLabel(slide.type),
    `動き：${BUILD_INFO[build]?.[0] ?? build}`,
    slide.details?.length ? `詳細${slide.details.length}か所` : null,
    story.drills[index]?.length ? `深掘り${story.drills[index].length}枚` : null,
  ].filter(Boolean).join("・");
  return { pos: `スライド ${index + 1} / ${deck.slides.length}`, detail };
}

/** The notes pane under the slide (shell.mjs): typing is one undo step per burst, as in the 編集 tab. */
function setNotes(value) {
  const slide = state.deck?.slides[state.selected];
  if (!slide) return;
  beginEdit();
  setPath(slide, ["notes"], value);
  delete slide.notesRich;
  const field = document.querySelector('#inspector [data-path="notes"]');
  if (field && field !== document.activeElement) field.value = value;
  markChanged();
}

/** ノートの書式: the notes typed with bold words, bullets and links. `notes` keeps the plain words (everything that reads the notes keeps working). */
function setNotesRich(html) {
  const slide = state.deck?.slides[state.selected];
  if (!slide) return;
  const fields = E.noteFields(html);
  if (fields.notes === String(slide.notes ?? "") && (fields.notesRich ?? "") === (slide.notesRich ?? "")) return;
  beginEdit();
  slide.notes = fields.notes;
  if (fields.notesRich) slide.notesRich = fields.notesRich; else delete slide.notesRich;
  const field = document.querySelector('#inspector [data-path="notes"]');
  if (field && field !== document.activeElement) field.value = fields.notes;
  markChanged();
}

/** スライド ショー →「非表示スライドに設定」: the slide stays for editing and is skipped while presenting. */
function toggleHiddenSlide() {
  const slide = state.deck?.slides[state.selected];
  if (!slide) return;
  if (state.selected === 0) return toast("表紙は非表示にできません");
  setSlideFields({ hidden: slide.hidden ? null : true });
  toast(slide.hidden ? "このスライドを非表示にしました（発表では飛ばします。編集用に残ります）" : "このスライドを発表に戻しました");
}

/**
 * Others' changes in a shared deck (coedit.mjs), already merged with ours: no undo step of their own, and the undo
 * history learns them too, so ⌘Z takes back only what was done here.
 */
function applyRemote(next, { slides = new Map(), order = null, meta = null } = {}) {
  if (!state.deck) return;
  const selectedSid = state.deck.slides[state.selected]?.sid;
  const theirs = new Set(slides.keys());
  const n = next.slides.length;
  const deck = { ...next, slides: next.slides.map((slide, i) => (theirs.has(slide.sid) ? normalizeSlide(clone(slide), i, n) : slide)) };
  state.deck = adoptDeck(deck);
  const at = deck.slides.findIndex((slide) => slide.sid === selectedSid);
  state.selected = at >= 0 ? at : Math.min(state.selected, deck.slides.length - 1);
  const patch = (text) => {
    try {
      const snap = JSON.parse(text);
      let list = snap.deck.slides.map((slide) => (slide.sid && slides.has(slide.sid) ? slides.get(slide.sid) : slide));
      if (order) {
        const bySid = new Map(list.map((slide) => [slide.sid, slide]));
        for (const [sid, slide] of slides) if (!bySid.has(sid)) bySid.set(sid, slide);
        const ordered = order.map((sid) => bySid.get(sid)).filter(Boolean);
        // Slides only in this snapshot that the others never had (made here, then undone) keep their place.
        const known = new Set(order);
        list.forEach((slide, i) => { if (!slide.sid || (!known.has(slide.sid) && !bySid.has(slide.sid))) ordered.splice(Math.min(i, ordered.length), 0, slide); });
        list = ordered;
      }
      snap.deck = { ...(meta ? { ...snap.deck, ...meta } : snap.deck), slides: list.length ? list : snap.deck.slides };
      return JSON.stringify(snap);
    } catch { return text; }
  };
  state.undo = state.undo.map(patch);
  state.redo = state.redo.map(patch);
  $("deckTitleInput").value = deck.title || "";
  saveCurrent();
  thumbCache.clear?.();
  renderFilmstrip();
  renderStage();
  renderInspector();
  if (state.panel === "anim") editorUi.renderAnimPane();
  if (state.panel === "comment") comments.render();
  if (state.panel === "a11y") a11y.render();
  scheduleMeasure();
}

/** The link to a shared deck, shown (and copied) when sharing starts. */
function showShareLink(text, copied) {
  const dialog = $("shareDialog");
  $("shareLink").value = text;
  $("shareStatus").textContent = copied ? "リンクをコピーしました。チャットやメールで送ってください。" : "リンクをコピーして、チャットやメールで送ってください。";
  dialog.showModal();
  $("shareLink").select();
}

/** The comments on a slide changed (comments.mjs): one undo step. */
function setCommentsOf(index, list, { undo = true } = {}) {
  const slide = state.deck?.slides[index];
  if (!slide) return;
  if (undo) pushUndo();
  const clean = list.filter((c) => c && typeof c.text === "string" && c.text.trim()).slice(0, 200).map((c, i) => normalizeComment(c, `c${Date.now().toString(36)}${i}`, true));
  if (clean.length) slide.comments = clean; else delete slide.comments;
  markChanged({ structural: true });
  comments.render();
}

// スライド ショー →「発表者ツールを使用」: the presenter view opens with the slide show.
const presenterViewOn = () => { try { return localStorage.getItem("hpptx-presenter-view") === "1"; } catch { return false; } };
function setPresenterViewOn(on) {
  try { localStorage.setItem("hpptx-presenter-view", on ? "1" : "0"); } catch { /* private window */ }
  toast(on ? "発表を始めると発表者ビュー（ノート・次のスライド・経過時間）を別ウィンドウで開きます" : "発表者ビューは発表中に P キーで開けます");
  editorUi.renderRibbon();
}

// Second text line shown in the outline for each kind of slide.
const OUTLINE_SECOND = { title: ["subtitle", "サブタイトル", 50], section: ["takeaway", "補足", 50], closing: ["message", "次のアクション", 90], statement: ["text", "大きく見せる一文", 60] };

function renderOutline(body) {
  if (body.querySelector(".outline-table")?.contains(document.activeElement)) return;
  const deck = state.deck;
  const rows = deck.slides.map((slide, index) => {
    const [secondKey, secondLabel, secondMax] = OUTLINE_SECOND[slide.type] ?? ["takeaway", "キーメッセージ", 60];
    const issues = issuesFor(index);
    const field = (key, label, max, cls) => {
      const count = h("span", { class: "counter" });
      const update = (value) => { const n = Math.ceil(charCount(value)); count.textContent = `${n}/${max}字`; count.classList.toggle("over", n > max); };
      const input = h("input", { type: "text", class: cls, value: slide[key] ?? "", placeholder: label, "aria-label": `${index + 1}枚目の${label}`, oninput: (event) => {
        beginEdit();
        setPath(slide, [key], event.target.value);
        update(event.target.value);
        markChanged();
        if (index === state.selected) renderInspector();
      } });
      update(slide[key] ?? "");
      return h("div", {}, input, h("div", { class: "label-row" }, h("span", { class: "hint" }, label), count));
    };
    return h("tr", { class: index === state.selected ? "selected" : "" },
      h("td", {}, h("button", { class: "outline-no", type: "button", title: "このスライドを開く", onclick: () => { setView("single"); select(index); } }, index + 1)),
      h("td", { style: { width: "130px" } }, h("span", { class: "outline-type" }, typeLabel(slide.type)),
        slide.drillOf && index > 0 ? h("div", { class: "hint", style: { "margin-top": "6px" } }, "↳ 深掘りページ（本編に数えない）") : null,
        issues.length ? h("div", { class: "hint", style: { color: "var(--warn)", "margin-top": "6px" } }, `⚠ 注意${issues.length}件`) : null),
      h("td", {}, field("title", slide.type === "title" ? "資料タイトル" : "タイトル", 30, "title")),
      h("td", {}, field(secondKey, secondLabel, secondMax, "")));
  });
  body.replaceChildren(h("table", { class: "outline-table" },
    h("thead", {}, h("tr", {}, h("th", {}, "#"), h("th", {}, "レイアウト"), h("th", {}, "タイトル"), h("th", {}, "キーメッセージ（結論）"))),
    h("tbody", {}, rows)),
    h("p", { class: "hint", style: { "max-width": "1280px", margin: "10px auto 0" } }, "キーメッセージだけを上から読んで話が通じるかを確認してください。番号をクリックするとそのスライドを開きます。"));
}

function setView(view) {
  state.view = view;
  if (view !== "single") state.masterEdit = false;
  shell.viewChanged();
  const radio = document.querySelector(`input[name=view][value=${view}]`);
  if (radio) radio.checked = true;
  stopMotionPreview({ render: false });
  renderStage();
}

/** Clicking text on the slide (or an issue) jumps to the matching form field (an object: it is selected). */
function focusField(field) {
  if (!field) return;
  if (field.startsWith("obj:")) {
    const id = field.slice(4);
    if (state.deck?.slides[state.selected]?.elements?.some((o) => o.id === id)) { editor.select([id]); openFormatPanel(); }
    return;
  }
  if (state.panel !== "form") setPanel("form");
  const candidates = [field];
  let rest = field;
  while (/(\[\d+\]|\.[^.[\]]+)$/.test(rest)) {
    rest = rest.replace(/(\[\d+\]|\.[^.[\]]+)$/, "");
    if (rest) candidates.push(rest);
  }
  for (const candidate of candidates) {
    const input = document.querySelector(`#inspector [data-path="${CSS.escape(candidate)}"]`);
    if (!input) continue;
    input.scrollIntoView({ behavior: "smooth", block: "center" });
    input.focus({ preventScroll: true });
    const line = candidate !== field && input.tagName === "TEXTAREA" ? Number(field.slice(candidate.length).match(/^\[(\d+)\]/)?.[1]) : NaN;
    if (Number.isInteger(line)) {
      const lines = input.value.split("\n");
      const start = lines.slice(0, line).reduce((sum, text) => sum + text.length + 1, 0);
      input.setSelectionRange(start, start + (lines[line]?.length ?? 0));
    }
    input.classList.remove("flash");
    void input.offsetWidth;
    input.classList.add("flash");
    return;
  }
}

// ---------------------------------------------------------------- objects placed by hand (the PowerPoint-style editor)
// public/editor/: canvas.mjs (select, move, resize, rotate, type, clipboard, keys), ui.mjs (ribbon, 書式 pane),
// ops.mjs (plain operations). Objects are slide.elements; the engine draws them (engine/objects.js).

const editorApp = {
  E, h,
  state: () => state,
  deck: () => state.deck,
  slide: () => (state.masterEdit && state.deck ? masterSlide() : state.deck ? state.deck.slides[state.selected] ?? null : null),
  index: () => state.selected,
  // スライド マスター表示 (表示 → スライド マスター): the objects every slide shows are edited like a slide's own.
  masterView: () => Boolean(state.masterEdit),
  openMaster: () => openMaster(),
  closeMaster: () => closeMaster(),
  isFinal: () => isFinal(state.deck),
  setFinal: (on) => setFinal(on),
  undo: () => undoRedo("undo"),
  redo: () => undoRedo("redo"),
  canUndo: () => state.undo.length > 0,
  canRedo: () => state.redo.length > 0,
  setObjects,
  setTimeline,
  setSlideFields,
  resetAllActions,
  prepareActionPlanPrompt,
  previewMotion: (opts) => previewMotion(opts),
  previewEffect: (entries) => previewEffect(entries),
  previewTransition: () => previewTransition(),
  stopPreview: () => stopMotionPreview(),
  previewing: () => Boolean(state.motionPreview),
  openAnimationPane: () => setPanel("anim"),
  panel: () => state.panel,
  renderOptions: () => renderOptions(),
  rerender: () => { if (state.deck && state.mode === "edit" && state.view === "single") renderStage(); },
  toast,
  canEdit: () => state.mode === "edit" && state.view === "single" && Boolean(state.deck) && !state.player && !state.motionPreview && !document.querySelector("dialog[open]"),
  busy: () => Boolean(state.motionPreview || state.player),
  focusStage: () => { if (state.inline) finishInlineEdit(true); },
  insertFiles: (files, at) => insertObjectFiles(files, at),
  insertUrl: (url, kind) => insertObjectUrl(url, kind),
  insertRecording: (blob, opts) => insertRecording(blob, opts),
  replaceImage: (id, file) => replaceObjectImage(id, file),
  pickFiles,
  ask: askDialog,
  ensureSid,
  slideOfSid: (sid) => state.deck?.slides.findIndex((slide) => slide.sid === sid) ?? -1,
  insertBlankSlide: () => insertSlide(Math.min(state.selected + 1, state.deck.slides.length - 1), defaultSlide("blank")),
  // サマリー ズーム: a new slide after this one, made with its parts already on it.
  insertSlideWith: (partial) => insertSlide(Math.min(state.selected + 1, state.deck.slides.length - 1), { ...defaultSlide(partial?.type || "blank"), ...partial }),
  select: (index) => select(index),
  onSlide: () => state.mode === "edit" && state.view === "single" && Boolean(state.deck?.slides[state.selected]),
  editEquation: (id) => editorUi.editEquation(id),
  editPoll: (id) => editorUi.editPoll(id),
  duplicateSlide: () => { const i = state.selected; if (i > 0 && i < state.deck.slides.length - 1) insertSlide(i + 1, copyOf(state.deck.slides[i])); else toast("表紙と最後のスライドは複製できません"); },
  openTypeDialog: (mode) => openTypeDialog(mode),
  setView: (view) => setView(view),
  openPanel: (panel, focus) => openFormatPanel(focus),
  showPanel: (panel) => setPanel(panel),
  showTab: (tab, opts) => editorUi.showTab(tab, opts),
  editChart: (id) => editorUi.editChart(id),
  editSmartart: (id, item) => editorUi.editSmartart(id, item),
  smartartToShapes: (id) => editorUi.smartartToShapes(id),
  startCrop: (id) => editorUi.startCrop(id),
  editPoints: (id) => editorUi.editPoints(id),
  fitFor: (index) => fitFor(index),
  convertSlide: () => convertSlideToObjects(),
  enhance: (scope) => enhanceDeck(scope),
  present: () => openPresenter(state.selected),
  canConvert: () => !converter.refusal(state.deck?.slides[state.selected]),
  // デザイン: the deck's motion; スライド ショー; 校閲 (the ribbon's tabs around the slide).
  deckTheme: () => state.deck?.theme || E.DEFAULT_THEME,
  // The deck's cover in another theme, small (デザイン → テーマ's gallery).
  themePreview: (id) => {
    useFonts([id]);
    const cover = state.deck?.slides[0] || { type: "title", title: themeMeta(id).name };
    return E.mount(E.render(cover, { ...renderOptions(), deck: { ...(state.deck || {}), theme: id }, index: 0, mode: "thumb" }));
  },
  deckMotion: () => (state.deck ? { transition: state.deck.transition || "fade", ...normalizeMotion(state.deck.motion || DEFAULT_MOTION) } : null),
  setDeckDesign: (patch) => setDeckDesign(patch),
  // 表示 → 配布資料マスター・ノート マスター (deck.printMasters; public/editor/printmaster.mjs).
  printMasters: () => state.deck?.printMasters || {},
  setPrintMaster: (kind, value) => {
    if (!state.deck || isFinal(state.deck)) return;
    const next = printMastersOf({ ...(state.deck.printMasters || {}), [kind]: value });
    if (JSON.stringify(next || null) === JSON.stringify(state.deck.printMasters || null)) return;
    pushUndo();
    if (next) state.deck.printMasters = next; else delete state.deck.printMasters;
    markChanged({ structural: true });
  },
  renderCtx: () => renderOptions(),
  openPrintMaster: (kind) => printMasterUi.open(kind),
  // スライドのサイズ → スライド番号の開始番号 (deck.firstNumber).
  firstNumber: () => firstNumberOf(state.deck?.firstNumber),
  setFirstNumber: (value) => { if (!state.deck || isFinal(state.deck)) return; const n = firstNumberOf(value); if (n === firstNumberOf(state.deck.firstNumber)) return; pushUndo(); if (n === 1) delete state.deck.firstNumber; else state.deck.firstNumber = n; markChanged({ structural: true }); toast(`スライド番号は ${n} から始まります`); },
  openDesign: () => openDesignDialog(),
  presentFrom: (index) => openPresenter(index),
  presentCustom: (id) => openPresenter(0, { custom: id }),
  customShows: () => customShowsOf(state.deck?.customShows) || [],
  showSettings: () => showOf(state.deck?.show) || {},
  setShowSettings: (patch) => setDeckFields({ show: showOf({ ...(showOf(state.deck?.show) || {}), ...patch }) }),
  openShowSettings: () => showTools.openSettings(),
  openCustomShows: (id) => showTools.openCustomShows(id),
  setDeckFields: (patch) => setDeckFields(patch),
  // アクセシビリティ: fixes made from the checker (each one undo step).
  openAccessibility: () => setPanel("a11y"),
  editAlt: (id) => a11y.editAlt(state.selected, id),
  selectObjects: (ids) => editor.select(ids),
  setObjectFields: (index, id, patch) => setObjectFields(index, id, patch),
  setSlideTitle: (index, text) => { const slide = state.deck?.slides[index]; if (!slide) return; pushUndo(); slide.title = text; markChanged({ structural: true }); },
  setReadingOrder: (index, ids) => { const slide = state.deck?.slides[index]; if (!slide) return; pushUndo(); if (ids?.length > 1) slide.readingOrder = [...ids]; else delete slide.readingOrder; markChanged({ structural: true }); },
  toggleHiddenSlide: () => toggleHiddenSlide(),
  presenterView: () => presenterViewOn(),
  setPresenterView: (on) => setPresenterViewOn(on),
  openCheck: () => openCheckDialog(false),
  fixOverflow: () => fixAllOverflow(),
  canFixOverflow: () => !$("fixAllBtn").classList.contains("hidden") && !state.aiBusy,
  openReplace: () => openReplace(),
  openNotes: () => openNotesDialog(),
  reviseDeck: () => openRewriteDialog(),
  canAi: () => Boolean(state.codexAuthorized) && !state.aiBusy,
  exportReview: () => exportReviewFile(),
  importReview: () => $("reviewImportFile").click(),
  // The window (public/editor/window.mjs): the notes under the slide, the thumbnails and the task pane, the zoom.
  toggleNotes: () => shell.toggleNotes(),
  notesShown: () => shell.notesShown(),
  togglePane: (which) => shell.togglePane(which),
  toggleRulers: () => shell.toggleRulers(),
  rulersShown: () => shell.rulersShown(),
  paneOpen: (which) => shell.paneOpen(which),
  zoomChanged: () => shell?.renderZoom(),
  setGuides: (next) => { if (!state.deck) return; pushUndo(); const g = guidesOf(next); if (g) state.deck.guides = g; else delete state.deck.guides; saveCurrent(); scheduleVersion(); },
  // The motion IDE: the slide's automatic motion (one row each in the animation pane), HTML motion on import.
  autoMotion: () => autoMotionInfo(),
  setAutoMotion: (key, value) => setAutoMotion(key, value),
  autoHtml: () => autoHtmlOn(),
  setAutoHtml: (on) => setAutoHtml(on),
  // 校閲 → コメント; ホーム・サムネイル → セクション; スライド ショー → リハーサル; 開発 (the slide's code).
  newComment: () => { setPanel("comment"); comments.focusNew(); },
  commentGo: (dir) => comments.go(dir),
  clearComments: (scope) => comments.clear(scope),
  commentPins: () => comments.showPins(),
  setCommentPins: (on) => comments.setShowPins(on),
  userName: () => userName(),
  changeUserName: () => comments.changeName(),
  share: () => (coedit.active ? coedit.showMenu() : coedit.start()),
  addSection: () => addSection(),
  rehearse: () => rehearse(),
  recordShow: (from) => recordShow(from),
  clearRecording: (what, scope) => clearRecording(what, scope),
  designIdeas: () => { setPanel("chat"); requestVariants(state.selected); },
  // ファイル (the ribbon's first menu).
  goCreate: () => setMode("create"),
  openLibrary: () => openLibrary(),
  saveDeck: (opts) => saveToLibrary(opts),
  openHistory: () => openHistory(),
  openHelp: () => $("helpDialog").showModal(),
  openSlideJson: () => openCodeDialog("json"),
  openSlideHtml: () => openCodeDialog("html"),
  exportHtml: () => exportHtml(),
  exportPptx: () => exportPptx(),
  saveJson: () => saveJsonFile(),
  openJson: () => openJsonDialog(),
  printPdf: () => printPdf(),
  printSlidesNow: () => printSlidesNow(),
  openNotesView: () => printer.notesView(),
  setNotes: (value) => setNotes(value),
  setNotesRich: (html) => setNotesRich(html),
  version: APP_VERSION,
  // 校閲 → スペル チェック・表記ゆれ; ファイル → 情報・ドキュメント検査.
  spellcheck: () => spellcheckOn(),
  setSpellcheck: (on) => { try { localStorage.setItem("hpptx-spellcheck", on ? "1" : "0"); } catch { /* private window */ } editorUi.renderRibbon(); toast(on ? "スペル チェック：オン（入力中の文字に赤い波線が出ます。ブラウザの辞書を使います）" : "スペル チェック：オフ"); },
  openProofing: () => proofing.open(),
  openSynonyms: () => synonyms.open(),
  // One AI job (a path and its request) as a promise of the finished job.
  aiRun: (path, body) => new Promise((resolve, reject) => {
    jsonFetch(path, { method: "POST", body: JSON.stringify(body) })
      .then(({ jobId }) => watchJob(jobId, { onProgress: () => {}, onDone: resolve, onFail: (job) => reject(new Error(job.error || job.detail || "失敗しました")) }))
      .catch(reject);
  }),
  allTexts: () => { const out = []; if (state.deck) { eachText(state.deck.slides, (t) => { out.push(t); return t; }); eachObjectText(state.deck.slides, (t) => { out.push(t); return t; }); } return out; },
  transformAllText: (fn) => transformAllText(fn),
  openFileInfo: () => fileInfo.open(),
  openWordCount: () => fileInfo.openWordCount(),
  // スライドの再利用・アウトラインからスライド・フォト アルバム.
  translate: (to, scope) => translateDeck(to, scope),
  // 図の形式 → 背景の削除・透明色・図の圧縮 (public/editor/picture.mjs).
  openCompare: () => compareTool.open(),
  versions: () => listVersions(),
  replaceSlides: (slides) => { if (!state.deck || !slides?.length) return; pushUndo(); state.deck.slides = slides.map((slide, i) => normalizeSlide(slide, i, slides.length)); state.selected = Math.min(state.selected, state.deck.slides.length - 1); markChanged({ structural: true }); },
  imageUrl: (src) => E.resolveSrc(src, renderOptions()),
  storeBlob: (blob, name) => putMedia(blob, name),
  allImages: () => (state.deck?.slides || []).flatMap((slide, i) => (slide.elements || []).filter((o) => o.kind === "image").map((o) => ({ slide: i, o }))),
  patchObjects: (changes) => patchObjects(changes),
  // ビデオの作成: the show played once by itself while the tab is recorded.
  openVideoExport: () => videoExport.openDialog(),
  // 画像として保存 (public/editor/imagexport.mjs).
  openImageExport: () => imageExport.openDialog(),
  // 図として保存 (public/editor/picsave.mjs): the chosen objects as a PNG / JPEG / SVG; 図として貼り付け draws copied
  // objects into a picture the same way.
  saveAsPicture: (ids) => pictureSaver.open(ids),
  objectsPicture: (objects) => pictureSaver.objectsPicture(objects),
  editorObjects: () => editor.objects(),
  // 既定の図形・テキスト ボックス・線に設定 (deck.objectDefaults), and 背景の保存.
  objectDefaults: () => state.deck?.objectDefaults || null,
  setObjectDefault: (kind, o) => setObjectDefault(kind, o),
  clearObjectDefaults: () => { if (!state.deck?.objectDefaults) return; pushUndo(); delete state.deck.objectDefaults; saveCurrent(); scheduleVersion(); toast("既定の図形・テキスト ボックス・線を元に戻しました"); },
  saveBackground: () => saveBackground(),
  openGifExport: () => gifExport.openDialog(),
  exportLockedHtml: () => exportLockedHtml(),
  formatBackground: () => background.open(),
  openAutoCorrect: () => autoCorrectUi.openOptions(),
  setBackground: (target, bg, { undo = true } = {}) => setBackground(target, bg, { undo }),
  showOrder: () => showSlides(state.deck, showOf(state.deck?.show) || {}) || E.storyMap(state.deck.slides).order.filter((i) => !state.deck.slides[i].hidden),
  presentForVideo: (extra, { fullscreen = true } = {}) => openPresenter(0, { extra, fullscreen }),
  // オンライン プレゼンテーション (public/editor/online.mjs).
  openPresenter: (start, options) => openPresenter(start, options),
  openOnline: () => online.open(),
  // コーチによるリハーサル (public/editor/coach.mjs).
  rehearseWithCoach: () => coach.start(),
  presenterHost: () => $("presenter"),
  onlineActive: () => online.active,
  fitsForShow: async () => { await measureAll(); return state.deck.slides.map((_, i) => { const fit = fitFor(i); return fit ? { fs: fit.fs, ts: fit.ts, ...(fit.objs ? { objs: fit.objs } : {}) } : null; }); },
  setMediaUrl: (src, url) => { mediaUrls[src] = url; },
  openViewer: (deck, options) => openViewer(deck, options),
  openViewerMessage: (text) => openViewerMessage(text),
  ensureAllSids: () => { if (state.deck) uniqueSids(state.deck.slides); },
  mediaBlob: (src) => mediaBlob(src),
  refreshRibbon: () => editorUi.renderRibbon(),
  player: () => state.player,
  animBusy: () => { const el = [...document.querySelectorAll("#presenter .hs-player-slide .hs-slide")].pop(); return Boolean(el && E.animBusy?.(el)); },
  closeShow: () => state.player?.close?.(),
  onShowClosed: (fn) => { state.onShowClosed = fn; },
  download: (blob, name) => downloadBlob(blob, name),
  openReuse: () => slideTools.openReuse(),
  openOutlineSlides: () => slideTools.openOutline(),
  openPhotoAlbum: () => slideTools.openPhotoAlbum(),
  savedDecks: () => listSavedDecks(),
  savedDeck: (id) => getSavedDeck(id),
  deckFromFile: (file) => deckFromFile(file),
  insertSlides: (slides) => insertSlides(slides),
  storePicture: async (file) => { const blob = /gif|svg/.test(file.type) ? file : await downscaleImage(file, 2400, /png|webp/.test(file.type) ? "image/png" : "image/jpeg"); const src = await putMedia(blob, file.name); const [w, hh] = await naturalSize(mediaUrls[src]); return { src, w, h: hh }; },
  stripDeck: (what) => { if (!state.deck) return; pushUndo(); stripDeckData(state.deck, what, E.W, E.H); if (state.selected >= state.deck.slides.length) state.selected = state.deck.slides.length - 1; markChanged({ structural: true }); toast("削除しました（⌘Zで戻せます）"); },
  renderSlide: (index) => E.render(state.deck.slides[index], renderOptions({ index, mode: "print", fit: fitFor(index) ?? undefined })),
  beforePrint: () => measureAll(),
};
const converter = createConverter(editorApp);
// 配布資料マスター・ノート マスター.
const printMasterUi = createPrintMasters(editorApp);
// スライド ショーの設定・目的別スライド ショー (public/editor/show.mjs).
const showTools = createShowTools(editorApp);
// 校閲 → アクセシビリティ チェック・読み取り順序・代替テキスト (public/editor/a11y.mjs).
const a11y = createA11y(editorApp);
// ファイル → 印刷 (public/editor/print.mjs).
const printer = createPrinter(editorApp);
// 校閲 → 表記ゆれチェック; ファイル → 情報 (public/editor/proof.mjs, fileinfo.mjs).
const proofing = createProofing(editorApp);
const fileInfo = createFileInfo(editorApp);
// スライドの再利用・アウトラインからスライド・フォト アルバム (public/editor/slidetools.mjs).
const slideTools = createSlideTools(editorApp);
// ファイル → エクスポート → ビデオの作成 (public/editor/video.mjs).
const videoExport = createVideoExport(editorApp);
// ファイル → エクスポート → 画像として保存 (public/editor/imagexport.mjs).
const imageExport = createImageExport(editorApp);
const pictureSaver = createPictureSaver(editorApp);
// ファイル → エクスポート → アニメーション GIF の作成 (public/editor/gif.mjs).
const gifExport = createGifExport(editorApp);
// デザイン → 背景の書式設定 (public/editor/background.mjs).
const background = createBackground(editorApp);
// 校閲 → オートコレクトのオプション (public/editor/autocorrect.mjs).
const autoCorrectUi = createAutoCorrect(editorApp);
// 校閲 → 比較 (public/editor/compare.mjs).
const compareTool = createCompare(editorApp);
// スライド ショー → オンライン プレゼンテーション (public/editor/online.mjs).
const online = createOnline(editorApp);
// スライド ショー → コーチによるリハーサル (public/editor/coach.mjs).
const coach = createCoach(editorApp);
const editor = createCanvas(editorApp);
// 校閲 → 類義語 (public/editor/synonyms.mjs): looks up the words selected in the canvas and puts the pick in their place.
const synonyms = createSynonyms(editorApp, editor);
const editorUi = createEditorUi(editor, editorApp);
// The window around the slide: splitters, notes, message bar, status bar (public/editor/window.mjs).
const shell = createShell({
  E, h, editor, toast, ask: askDialog,
  slide: () => editorApp.slide(),
  view: () => state.view,
  setNotes,
  setNotesRich,
  slideStatus,
  openZoomMenu: (anchor) => editorUi.openPop(anchor, editorUi.zoomMenu()),
  ribbonChanged: () => editorUi.renderRibbon(),
});
// Shared editing (共有, public/editor/coedit.mjs): set once the deck is shared.
let coedit = null;
// Review comments (校閲, public/editor/comments.mjs).
const comments = createComments({
  h, E, editor, deck: () => state.deck, index: () => state.selected, select: (i) => select(i), toast, setCommentsOf, ask: askDialog,
  showPanel: (panel) => setPanel(panel, { reveal: true }), onSlide: () => state.view === "single",
  people: () => coedit?.names() || [], nameChanged: (name) => coedit?.rename(name),
});
// 共同編集 (public/editor/coedit.mjs): a room on the server, changes merged slide by slide, people shown.
coedit = createCoedit({
  h, E, toast, ask: askDialog, confirm: (message) => Promise.resolve(window.confirm(message)),
  deck: () => state.deck, index: () => state.selected, select: (i) => select(i), selection: () => editor.selection,
  ensureAllSids: () => { if (state.deck) uniqueSids(state.deck.slides); },
  loadShared: (deck) => loadDeck(deck, { source: "共同編集", selected: 0 }),
  applyRemote: (next, info) => applyRemote(next, info),
  busy: () => Boolean(editor.typing || state.inline || document.querySelector(".sa-pane input:focus, .cm-input:focus")),
  redraw: () => { editor.draw(); renderFilmstrip(); },
  mediaBlob: (src) => mediaBlob(src),
  hasMedia: async (src) => Boolean(mediaUrls[src] || (await mediaBlob(src).catch(() => null))),
  putMediaAs: (src, blob) => putMediaAs(src, blob),
  mediaArrived: () => { mediaEpoch += 1; thumbCache.clear(); renderFilmstrip(); renderStage(); },
  showLink: (text, copied) => showShareLink(text, copied),
  openMenu: (anchor, items) => editorUi.popMenu(anchor, items),
});
editor.overlay((list, k) => coedit.overlay(list, k));
editor.subscribe(() => coedit.sendPresence());
// Read-only access for the browser checks (qa/studio-objects.mjs): the deck and the slide on the stage.
window.__hpptx = { coedit: () => coedit?.info(), master: () => Boolean(state.masterEdit), deck: () => state.deck, slide: () => state.deck?.slides[state.selected] ?? null, selection: () => editor.selection, typing: () => editor.state.typing?.cell ?? (editor.typing ? "text" : null), cellRange: () => editor.cellRange };

/**
 * The objects of the slide on the stage changed (one undo step unless told otherwise). Their animations follow:
 * `timeline` replaces them in the same step, and animations of objects that are gone go with them.
 */
function setObjects(list, { undo = true, timeline = undefined } = {}) {
  // スライド マスター表示: what is drawn there is the deck's master objects.
  if (state.masterEdit && state.deck) {
    if (undo) pushUndo();
    if (list.length) state.deck.masterObjects = list; else delete state.deck.masterObjects;
    markChanged({ structural: true });
    return;
  }
  const slide = state.deck?.slides[state.selected];
  if (!slide) return;
  if (undo) pushUndo();
  const before = slide.elements || [];
  if (list.length) slide.elements = list; else delete slide.elements;
  const next = ops.reconcileTimeline(timeline !== undefined ? timeline : slide.timeline, before, list);
  if (next?.length) slide.timeline = next; else delete slide.timeline;
  markChanged({ structural: true });
}

/**
 * HTMLの動きをおまかせで付ける (editor/htmlfx.mjs): on the slide on the stage, or on every slide, in one undo step.
 * What a slide already has (its own animations, actions, hovers) stays.
 */
function enhanceDeck(scope = "slide", { quiet = false } = {}) {
  if (!state.deck) return null;
  const indices = scope === "all" ? state.deck.slides.map((_, i) => i).filter((i) => !state.deck.slides[i].hidden) : [state.selected];
  // The template's furniture repeated on slide after slide (a logo, 秘（B）, a slogan) is left alone.
  const chrome = chromeOf(state.deck.slides);
  const totals = {};
  let count = 0;
  for (const i of indices) {
    const slide = state.deck.slides[i];
    if (!slide?.elements?.length) continue;
    const out = enhanceSlide(slide, E, { chrome });
    if (!out.count) continue;
    if (!count) pushUndo();
    slide.elements = E.normalizeObjects(out.elements);
    const timeline = E.normalizeTimeline(out.timeline, slide);
    if (timeline.length) slide.timeline = timeline; else delete slide.timeline;
    count += out.count;
    for (const [key, n] of Object.entries(out.added)) totals[key] = (totals[key] || 0) + n;
  }
  if (!count) {
    if (!quiet) toast(scope === "all" ? "付けられる部品が見つかりませんでした（すでに動きがあるか、部品のないスライドです）" : "このスライドには付けられる部品がありません（すでに動きがあるか、部品がありません）");
    return null;
  }
  const summary = describeAdded(totals);
  // On a deck just brought over, the message bar says what was added and offers to take it back.
  if (scope === "all" && state.imported?.fidelity === "exact") state.imported.auto = { count, summary, undoDepth: state.undo.length };
  markChanged({ structural: true });
  if (!quiet) toast(`HTMLならではの動きを${count}か所に付けました（${summary}）。⌘Zで元に戻せます`);
  return { count, summary };
}

// HTMLの動きを自動で付ける: a deck brought over as it looks gets HTML's moves at once (one undo step takes them back).
const autoHtmlOn = () => { try { return localStorage.getItem("hpptx-auto-html") !== "0"; } catch { return true; } };
function setAutoHtml(on) {
  try { localStorage.setItem("hpptx-auto-html", on ? "1" : "0"); } catch { /* private window */ }
  toast(on ? "PowerPointを見た目どおりに取り込んだら、HTMLの動きを自動で付けます" : "取り込んだときに自動では付けません（「おまかせ」でいつでも付けられます）");
  editorUi.renderRibbon();
  renderStage();
}
/** The message bar's 「元に戻す」: the deck as it came from PowerPoint (the step おまかせ added, undone). */
function undoAutoHtml() {
  const auto = state.imported?.auto;
  if (!auto) return;
  if (state.undo.length === auto.undoDepth) undoRedo("undo");
  else toast("その後に編集があるため、⌘Zで順に戻してください");
  if (state.imported) state.imported.auto = null;
  renderStage();
}

/** The animations of the slide on the stage changed (the animation pane and ribbon). */
function setTimeline(list, { undo = true } = {}) {
  if (state.masterEdit) { toast("スライド マスターの部品には動きを付けられません"); return; }
  const slide = state.deck?.slides[state.selected];
  if (!slide) return;
  if (undo) pushUndo();
  const clean = E.normalizeTimeline(list, slide);
  if (clean.length) slide.timeline = clean; else delete slide.timeline;
  markChanged({ structural: true });
}

/** Slide settings set from the ribbon (画面切り替え): one undo step, then the stage and thumbnails redraw. */
function setSlideFields(patch, { all = false } = {}) {
  if (!state.deck) return;
  pushUndo();
  const targets = all ? state.deck.slides : [state.deck.slides[state.selected]];
  for (const slide of targets) {
    for (const [key, value] of Object.entries(patch)) {
      if (value == null || value === "") delete slide[key]; else slide[key] = value;
    }
  }
  markChanged({ structural: true });
}

/** One undo step removes every imported and studio-authored presentation action. */
function resetAllActions() {
  if (!state.deck || state.player || state.aiBusy) return;
  if (state.inline) finishInlineEdit(true);
  const { deck, removed, changed } = resetDeckActions(state.deck);
  if (!changed) return toast("この資料の動きとクリック動作はすでに削除されています");
  stopMotionPreview({ render: false });
  pushUndo();
  state.deck = deck;
  state.imported = null;
  thumbCache.clear();
  markChanged({ structural: true });
  toast(`全${removed.slides}枚の動きとクリック動作を削除しました（⌘Zで元に戻せます）`);
}

/** Put a whole-deck, HTML-first animation brief in chat for review before sending it to AI. */
function prepareActionPlanPrompt() {
  if (!state.deck) return;
  setPanel("chat");
  $("chatInput").value = "全ページの内容と前後の流れを読み、PowerPoint由来の動きに引きずられず、HTMLならではのアニメーションとクリック操作をゼロから考えてください。まずは各ページ番号ごとに、動かす対象、効果、開始条件、目的を具体的に提案してください。元の文字・数値・図形・写真は変えず、過剰な演出は避けてください。今回は提案だけを返し、資料は変更しないでください。";
  $("chatInput").dispatchEvent(new Event("input", { bubbles: true }));
  $("chatInput").focus();
}

/**
 * 図形に変換: the slide on the stage becomes 白紙, its layout drawn as objects people edit one by one (its build
 * as animations on them). One undo step, as in PowerPoint.
 */
async function convertSlideToObjects() {
  if (!state.deck || state.converting) return;
  const index = state.selected;
  const why = converter.refusal(state.deck.slides[index]);
  if (why) { toast(why); return; }
  if (state.inline) finishInlineEdit(true);
  state.converting = true;
  try {
    const before = state.deck.slides[index];
    const result = await converter.convert(index);
    if (!state.deck || state.deck.slides[index] !== before) return;
    const next = result.slide;
    const timeline = E.normalizeTimeline(next.timeline, next);
    if (timeline.length) next.timeline = timeline; else delete next.timeline;
    pushUndo();
    state.deck.slides[index] = next;
    editor.select?.([]);
    markChanged({ structural: true });
    toast(`図形に変換しました（${result.objects.length}個の部品）。1つずつ動かして編集できます。⌘Zで元に戻せます`);
  } catch (error) {
    console.error(error);
    toast(`図形に変換できませんでした：${error.message || error}`);
  } finally {
    state.converting = false;
  }
}

function openFormatPanel(focus = null) {
  setPanel("format");
  editorUi.renderPane(focus);
}

/** The slide id links point at (made when first needed). */
function ensureSid(index) {
  const slide = state.deck?.slides[index];
  if (!slide) return null;
  if (!slide.sid) slide.sid = newSid(new Set(state.deck.slides.map((s) => s.sid).filter(Boolean)));
  return slide.sid;
}

/** Ask for files with the browser's picker (resolves to an empty list when cancelled). */
function pickFiles(accept, multiple = false) {
  return new Promise((resolve) => {
    const input = h("input", { type: "file", accept, multiple: multiple || null, class: "hidden" });
    input.addEventListener("change", () => { resolve([...(input.files || [])]); input.remove(); });
    input.addEventListener("cancel", () => { resolve([]); input.remove(); });
    document.body.append(input);
    input.click();
  });
}

/** A one-line question in a dialog (resolves to the text, or null when cancelled). */
function askDialog(title, label = "", value = "") {
  return new Promise((resolve) => {
    const dialog = $("askDialog");
    $("askTitle").textContent = title;
    $("askLabel").textContent = label;
    $("askLabel").hidden = !label;
    const input = $("askInput");
    input.value = value ?? "";
    let answer = null;
    const ok = () => { answer = input.value; dialog.close(); };
    const onKey = (event) => { if (event.key === "Enter" && !event.isComposing) { event.preventDefault(); ok(); } };
    $("askOk").onclick = ok;
    input.addEventListener("keydown", onKey);
    dialog.addEventListener("close", () => { input.removeEventListener("keydown", onKey); resolve(answer); }, { once: true });
    dialog.showModal();
    input.focus();
    input.select();
  });
}

const naturalSize = (url) => new Promise((resolve) => {
  const img = new Image();
  img.onload = () => resolve([img.naturalWidth || 960, img.naturalHeight || 540]);
  img.onerror = () => resolve([960, 540]);
  img.src = url;
});
const videoSize = (url) => new Promise((resolve) => {
  const video = document.createElement("video");
  video.preload = "metadata";
  video.onloadedmetadata = () => resolve([video.videoWidth || 1280, video.videoHeight || 720]);
  video.onerror = () => resolve([1280, 720]);
  video.src = url;
});
/** A box of the picture's own proportions that fits in w × h (slide pixels). */
const fitBox = ([w, hh], maxW = 960, maxH = 640) => { const k = Math.min(maxW / w, maxH / hh, 1); return { x: 0, y: 0, w: Math.round(w * k), h: Math.round(hh * k) }; };

/** Pictures, videos and Lottie animations from files → objects on the slide (kept in this browser). */
async function insertObjectFiles(files, at = null) {
  if (!state.deck) return [];
  const made = [];
  for (const file of files) {
    try {
      if (isLottieFile(file)) {
        const text = await readLottie(file);
        const src = await putMedia(new Blob([text], { type: "application/json" }), file.name);
        made.push(ops.makeObject("lottie", { x: 0, y: 0, w: 480, h: 480 }, { src, fileName: file.name }));
      } else if (file.type.startsWith("video/")) {
        if (file.size > 400_000_000) { toast("動画が大きすぎます（400MBまで）"); continue; }
        const src = await putMedia(file, file.name);
        made.push(ops.makeObject("video", fitBox(await videoSize(mediaUrls[src])), { src, fileName: file.name, autoplay: true, loop: true, muted: true }));
      } else if (file.type.startsWith("audio/") || /\.(mp3|m4a|wav|ogg|oga|aac|flac|opus)$/i.test(file.name)) {
        if (file.size > 200_000_000) { toast("音声が大きすぎます（200MBまで）"); continue; }
        const src = await putMedia(file, file.name);
        made.push(audioObject(src, file.name));
      } else if (file.type.startsWith("image/")) {
        const blob = /gif|svg/.test(file.type) ? file : await downscaleImage(file, 2400, /png|webp/.test(file.type) ? "image/png" : "image/jpeg");
        const src = await putMedia(blob, file.name);
        made.push(ops.makeObject("image", fitBox(await naturalSize(mediaUrls[src])), { src, fileName: file.name, alt: file.name.replace(/\.[a-z0-9]+$/i, "") }));
      } else toast(`「${file.name}」は入れられません（画像・動画・音声・Lottie JSON）`);
    } catch (error) {
      toast(`読み込めませんでした：${error.message}`);
    }
  }
  // A sound goes in the middle of the slide, as PowerPoint puts it.
  if (made.length) editor.insert(made, { at: at ?? (made.every((o) => o.kind === "audio") ? [E.W / 2, E.H / 2] : null) });
  return made;
}

/** A sound on the slide: the speaker icon in the middle of the slide, as PowerPoint puts it (挿入 → オーディオ). */
function audioObject(src, name) {
  const size = 120;
  return ops.makeObject("audio", { x: (E.W - size) / 2, y: (E.H - size) / 2, w: size, h: size }, { src, fileName: name, name: String(name || "オーディオ").replace(/\.[a-z0-9]+$/i, "").slice(0, 60) });
}

/** A recording made in the studio (オーディオの録音・画面録画): kept in this browser, then placed on the slide. */
async function insertRecording(blob, { kind = "audio", name = "" } = {}) {
  if (!state.deck || !blob?.size) return;
  const src = await putMedia(blob, name);
  if (kind === "video") {
    editor.insert([ops.makeObject("video", fitBox(await videoSize(mediaUrls[src])), { src, fileName: name, autoplay: false, loop: false, muted: false })]);
    toast("画面録画をビデオとして入れました（再生タブでトリミングできます）");
  } else {
    editor.insert([audioObject(src, name)], { at: [E.W / 2, E.H / 2] });
    toast("録音を入れました（発表中にアイコンをクリックすると再生します）");
  }
}

/** An image, a video (YouTube, mp4), a sound or a Lottie animation from the web → an object. */
async function insertObjectUrl(url, kind) {
  if (!/^https:\/\//i.test(url)) return toast("https:// で始まるURLを入れてください");
  try {
    if (kind === "audio" || /\.(mp3|m4a|wav|ogg|oga|aac|flac|opus)(\?|#|$)/i.test(url)) {
      editor.insert([audioObject(url, decodeURIComponent(url.split("/").pop().split("?")[0]).slice(0, 60) || "オーディオ")], { at: [E.W / 2, E.H / 2] });
      toast("オーディオを入れました（発表の画面で再生されます。インターネットにつながっている必要があります）");
      return;
    }
    if (/\.json(\?|#|$)/i.test(url) || /^https:\/\/(lottie\.host|[a-z0-9-]+\.lottiefiles\.com)\//i.test(url)) {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const text = await readLottie(await response.blob());
      const src = await putMedia(new Blob([text], { type: "application/json" }), url.split("/").pop());
      editor.insert([ops.makeObject("lottie", { x: 0, y: 0, w: 480, h: 480 }, { src, fileName: decodeURIComponent(url.split("/").pop().split("?")[0]).slice(0, 60) })]);
      return;
    }
    const yt = E.youtubeId(url);
    if (yt || /\.(mp4|webm|mov|m4v)(\?|#|$)/i.test(url) || kind === "video") {
      editor.insert([ops.makeObject("video", { x: 0, y: 0, w: 960, h: 540 }, { src: url, fileName: yt ? "YouTube" : url.split("/").pop().split("?")[0].slice(0, 60), autoplay: !yt, loop: !yt, muted: !yt })]);
      toast(yt ? "YouTubeを入れました。発表の画面で再生できます" : "動画を入れました");
      return;
    }
    editor.insert([ops.makeObject("image", fitBox(await naturalSize(url)), { src: url, alt: "" })]);
  } catch (error) {
    toast(`読み込めませんでした（${error.message}）`);
  }
}

async function replaceObjectImage(id, file) {
  try {
    const blob = /gif|svg/.test(file.type) ? file : await downscaleImage(file, 2400, /png|webp/.test(file.type) ? "image/png" : "image/jpeg");
    const src = await putMedia(blob, file.name);
    const [w, hh] = await naturalSize(mediaUrls[src]);
    editor.apply((o) => {
      if (o.id !== id) return null;
      // The new picture keeps the box's width and takes its own proportions.
      const height = Math.round(o.w * (hh / w));
      return { src, fileName: file.name, crop: undefined, h: height, y: o.y + (o.h - height) / 2 };
    }, { ids: [id] });
  } catch (error) {
    toast(`差し替えられませんでした：${error.message}`);
  }
}

// ---------------------------------------------------------------- editing on the slide

const INLINE_AI = [["短く", "もっと短く"], ["言い換え", "伝わりやすく言い換えて"], ["具体的に", "具体例や数字を入れて具体的に"], ["数字を強調", "数字が目立つように"]];
const MULTILINE_FIELDS = new Set(["message", "text", "notes"]);

function fieldLabelFor(type, path) {
  const names = { title: "タイトル", takeaway: "キーメッセージ", subtitle: "サブタイトル", message: "次のアクション", desc: "説明", points: "箇条書き", items: "項目", steps: "工程", text: "本文", summary: "まとめ", label: "ラベル", value: "数値", q: "質問", a: "回答" };
  const last = [...path].reverse().find((part) => typeof part === "string");
  return names[last] || typeLabel(type);
}

/** Click a text on the slide to type straight into it; the form is only a click away. */
function beginInlineEdit(el) {
  if (state.inline) finishInlineEdit(true);
  const field = el.dataset.field;
  const index = state.selected;
  const slide = state.deck.slides[index];
  const path = parseField(field);
  const value = getPath(slide, path);
  if (typeof value !== "string" || squash(value) !== squash(el.textContent)) { focusField(field); return; }
  const multiline = MULTILINE_FIELDS.has(String(path.at(-1))) || value.includes("\n");
  const original = el.innerHTML;
  el.classList.add("hs-editing");
  el.textContent = value;
  el.setAttribute("contenteditable", "plaintext-only");
  if (el.contentEditable !== "plaintext-only") el.setAttribute("contenteditable", "true");
  el.spellcheck = false;
  el.focus();
  const range = document.createRange();
  range.selectNodeContents(el);
  const selection = window.getSelection();
  selection.removeAllRanges();
  selection.addRange(range);
  const label = fieldLabelFor(slide.type, path);
  const tools = h("div", { class: "inline-tools", role: "toolbar", "aria-label": "この文字をAIで直す" },
    ...INLINE_AI.map(([text, ask]) => h("button", { type: "button", onmousedown: (event) => event.preventDefault(), onclick: () => inlineAsk(ask) }, `✦ ${text}`)),
    h("button", { type: "button", onmousedown: (event) => event.preventDefault(), onclick: () => inlineAsk(null) }, "✦ 指示…"),
    h("span", { class: "sep" }),
    h("button", { type: "button", onmousedown: (event) => event.preventDefault(), onclick: () => { finishInlineEdit(true); focusField(field); } }, "編集欄で開く"));
  const wrap = el.closest(".slide-wrap");
  const wrapRect = wrap.getBoundingClientRect();
  const rect = el.getBoundingClientRect();
  tools.style.left = `${Math.max(0, Math.min(rect.left - wrapRect.left, wrapRect.width - 380))}px`;
  tools.style.top = `${Math.max(38, rect.top - wrapRect.top)}px`;
  wrap.append(tools);
  state.inline = { el, field, index, path, original, tools, label, before: value, multiline };
}

function finishInlineEdit(commit = true) {
  const inline = state.inline;
  if (!inline) return;
  state.inline = null;
  inline.tools.remove();
  inline.el.classList.remove("hs-editing");
  inline.el.removeAttribute("contenteditable");
  let next = inline.el.innerText.replace(/\r/g, "").replace(/\n{3,}/g, "\n\n").replace(/\n$/, "");
  if (!inline.multiline) next = next.replace(/\n/g, "").trim();
  const changed = commit && next !== inline.before && strip(next);
  const slide = state.deck?.slides[inline.index];
  if (!changed || !slide) {
    inline.el.innerHTML = inline.original;
    if (state.stageDirty) { state.stageDirty = false; renderStage(); }
    return;
  }
  pushUndo();
  setPath(slide, inline.path, next);
  state.stageDirty = false;
  markChanged({ structural: true });
}

function inlineAsk(ask) {
  const inline = state.inline;
  if (!inline) return;
  const current = inline.el.innerText.trim();
  finishInlineEdit(true);
  const target = `@${inline.index + 1} の${inline.label}「${String(current).slice(0, 60)}」`;
  if (ask) { sendChat(`${target}を${ask}`); return; }
  setPanel("chat");
  $("chatInput").value = `${target}を`;
  $("chatInput").focus();
  renderChatContext();
}

// ---- photos and videos placed by hand: drag to move, the corner to resize, ✕ to remove

function placementOwner(slide) {
  if (slide.media?.src) return { get: () => slide.media.placement, set: (p) => { slide.media.placement = p; }, remove: () => { delete slide.media; } };
  if (slide.customImage) return { get: () => slide.imagePlacement, set: (p) => { slide.imagePlacement = p; }, remove: () => { delete slide.customImage; delete slide.imagePlacement; } };
  return null;
}

function wirePlacedMedia(slideEl, slide) {
  const placed = slideEl.querySelector(".hs-placed[data-placed]");
  const owner = placementOwner(slide);
  if (!placed || !owner) return;
  placed.title = "ドラッグで移動・右下の丸でサイズ変更";
  const resize = h("span", { class: "ph-handle ph-resize", title: "サイズを変える" }, "↘");
  const remove = h("button", { class: "ph-handle ph-remove", type: "button", title: "スライドから外す" }, "×");
  placed.append(resize, remove);
  remove.addEventListener("pointerdown", (event) => event.stopPropagation());
  remove.addEventListener("click", (event) => {
    event.stopPropagation();
    pushUndo();
    owner.remove();
    markChanged({ structural: true });
    toast("写真・動画を外しました（⌘Zで元に戻せます）");
  });
  placed.addEventListener("pointerdown", (event) => {
    if (event.button !== 0) return;
    event.preventDefault();
    event.stopPropagation();
    const sizing = event.target === resize;
    const rect = slideEl.getBoundingClientRect();
    const start = { x: event.clientX, y: event.clientY, p: { ...(owner.get() || DEFAULT_PLACEMENT) } };
    let next = start.p;
    placed.setPointerCapture(event.pointerId);
    placed.classList.add("dragging");
    const move = (e) => {
      const dx = (e.clientX - start.x) / rect.width;
      const dy = (e.clientY - start.y) / rect.height;
      next = clampPlacement(sizing ? { ...start.p, w: Math.max(0.08, start.p.w + dx), h: Math.max(0.08, start.p.h + dy) } : { ...start.p, x: start.p.x + dx, y: start.p.y + dy });
      Object.assign(placed.style, { left: `${next.x * 100}%`, top: `${next.y * 100}%`, width: `${next.w * 100}%`, height: `${next.h * 100}%` });
    };
    const up = () => {
      placed.removeEventListener("pointermove", move);
      placed.removeEventListener("pointerup", up);
      placed.removeEventListener("pointercancel", up);
      placed.classList.remove("dragging");
      if (JSON.stringify(next) === JSON.stringify(start.p) && owner.get()) return;
      pushUndo();
      owner.set({ x: +next.x.toFixed(4), y: +next.y.toFixed(4), w: +next.w.toFixed(4), h: +next.h.toFixed(4) });
      markChanged({ structural: true });
    };
    placed.addEventListener("pointermove", move);
    placed.addEventListener("pointerup", up);
    placed.addEventListener("pointercancel", up);
  });
}

function onStageDragOver(event) {
  const types = event.dataTransfer?.types ?? [];
  if (types.includes("text/x-hs-generated-image") || types.includes("Files")) { event.preventDefault(); event.dataTransfer.dropEffect = "copy"; }
}

async function onStageDrop(event) {
  const id = event.dataTransfer.getData("text/x-hs-generated-image");
  const files = [...(event.dataTransfer.files || [])];
  if (!id && !files.length) return;
  event.preventDefault();
  const slideEl = event.currentTarget.querySelector(".hs-slide");
  const rect = slideEl.getBoundingClientRect();
  const at = [((event.clientX - rect.left) / rect.width) * E.W, ((event.clientY - rect.top) / rect.height) * E.H];
  if (id) {
    // A picture the AI made in the chat becomes an object where it was dropped.
    const message = state.chat.messages.find((entry) => entry.id === id);
    const meta = message?.generatedImage;
    if (!meta?.src) return toast("この画像は配置できません。再生成してください");
    // Kept in this browser like any other picture (the deck itself stays small).
    const src = meta.src.startsWith("data:") ? await putMedia(await dataUrlToBlob(meta.src), "AIで作った画像") : meta.src;
    const [w, hh] = await naturalSize(src.startsWith("idb:") ? mediaUrls[src] : src);
    editor.insert([ops.makeObject("image", fitBox([w, hh], 760, 560), { src, alt: "AIで作った画像" })], { at });
    meta.status = "applied";
    saveChat();
    renderChat();
    return;
  }
  if (files.some((file) => /\.lottie$/i.test(file.name || ""))) return toast("「.lottie」形式は使えません。LottieFilesでは「Lottie JSON」を選んでください");
  await insertObjectFiles(files, at);
}

// ---- motion preview on the stage (the editor itself stays still)

/**
 * Play the slide on the stage as in a presentation (「動きを確認」, the animation tab's 「プレビュー」).
 * `from`: start at that click step with the earlier ones already done (the animation pane's 「ここから再生」);
 * `auto: false` waits for clicks instead of playing the clicks by itself.
 */
function previewMotion({ from = 0, auto = true, brief = false, slide: shown = null } = {}) {
  if (!state.deck || state.view !== "single") { setView("single"); }
  const wrap = document.querySelector("#stageBody .slide-wrap");
  if (!wrap) return;
  if (state.inline) finishInlineEdit(true);
  stopMotionPreview({ render: false });
  const index = state.selected;
  // `shown`: the slide with other animations than its own (描画で再生 plays only the ink being written).
  const slide = shown || state.deck.slides[index];
  const el = E.render(slide, renderOptions({ index, mode: "present", fit: fitFor(index) ?? undefined }));
  const scaler = E.mount(el);
  const steps = () => E.stepsOf(el);
  let step = Math.max(0, Math.min(from, steps()));
  const count = h("span", {});
  const updateCount = () => { count.textContent = steps() ? `クリックで次へ（${step}/${steps()}）` : "マウスを乗せる・クリックで詳細"; };
  const banner = h("div", { class: "motion-banner" }, "▶ 動きを確認中", count,
    h("button", { class: "btn", type: "button", onclick: (event) => { event.stopPropagation(); previewMotion({ auto, brief, slide: shown }); } }, "もう一度"),
    h("button", { class: "btn", type: "button", onclick: (event) => { event.stopPropagation(); stopMotionPreview(); } }, "編集に戻る"));
  wrap.replaceChildren(scaler, banner);
  // Starting later, what comes before is already done; the step itself then plays.
  let first = 0;
  if (step > 0) { E.play(el, { step: step - 1, animate: false }); first = E.reveal(el, step) || 0; } else first = E.play(el, { step: 0 }) || 0;
  const act = E.activate(el, { details: slide.details || [] });
  E.playMedia(el);
  // A click while animations still play finishes them first, as in a presentation. Returns the next step's length.
  const advance = () => {
    if (act.detailOpen) return 0;
    if (E.animBusy?.(el)) { E.animFinish(el); return 0; }
    if (step < steps()) { step += 1; const ms = E.reveal(el, step) || 0; updateCount(); return ms; }
    return 0;
  };
  scaler.addEventListener("click", advance);
  updateCount();
  state.motionPreview = { el, act, advance, brief };
  // Click steps play themselves once (each after the one before has finished), so the whole sequence can be
  // seen without clicking.
  if (steps() && auto) {
    const next = () => {
      if (state.motionPreview?.el !== el) return;
      if (E.animBusy?.(el)) { state.motionPreview.timer = setTimeout(next, 150); return; }
      if (step < steps()) { advance(); state.motionPreview.timer = setTimeout(next, 700); }
    };
    state.motionPreview.timer = setTimeout(next, Math.max(1100, first + 500));
  }
}

/** Plays only these animations on the stage (the slide's own are left alone), then goes back to editing. */
function previewEffect(entries) {
  const slide = state.deck?.slides[state.selected];
  if (!slide || !entries?.length) return;
  const timeline = E.normalizeTimeline(entries, slide);
  if (!timeline.length) return;
  previewMotion({ slide: { ...slide, timeline }, brief: true });
  const el = state.motionPreview?.el;
  const total = Math.max(...timeline.map((e) => (e.delay || 0) + (e.dur || 600)));
  setTimeout(() => { if (state.motionPreview?.el === el) stopMotionPreview(); }, Math.min(10000, total + 1200));
}

/** The 画面切り替え tab's 「プレビュー」: the slide before gives way to this one, with this slide's transition. */
async function previewTransition() {
  if (!state.deck) return;
  if (state.view !== "single") setView("single");
  const wrap = document.querySelector("#stageBody .slide-wrap");
  if (!wrap) return;
  if (state.inline) finishInlineEdit(true);
  stopMotionPreview({ render: false });
  const index = state.selected;
  const slide = state.deck.slides[index];
  const type = TRANSITIONS.includes(slide.transition) ? slide.transition : state.deck.transition || "fade";
  const before = index > 0 ? index - 1 : index;
  const from = E.render(state.deck.slides[before], renderOptions({ index: before, mode: "thumb", fit: fitFor(before) ?? undefined }));
  const to = E.render(slide, renderOptions({ index, mode: "present", fit: fitFor(index) ?? undefined }));
  const host = h("div", { class: "hs-player-stage transition-preview" });
  wrap.replaceChildren(host);
  state.motionPreview = { el: to, act: null, advance: () => 0, transition: true };
  await E.transitionPreview(host, from, to, type, { dur: slide.transitionDur, dir: slide.transitionDir, pairs: type === "morph" && before !== index ? E.morphPairs(state.deck.slides[before].elements, slide.elements) : null });
  if (state.motionPreview?.el !== to) return;
  E.play(to, { step: 0 });
  state.motionPreview.timer = setTimeout(() => { if (state.motionPreview?.el === to) stopMotionPreview(); }, 1600);
}

function stopMotionPreview({ render = true } = {}) {
  const preview = state.motionPreview;
  if (!preview) return;
  clearTimeout(preview.timer);
  preview.act?.destroy();
  E.stopMedia(preview.el);
  state.motionPreview = null;
  if (render) renderStage();
}

// ---------------------------------------------------------------- speaker notes

const E_ROW = "えけげせぜてでねへべぺめれいきぎしじちぢにひびぴみり";

/** Turn a written-style fragment into a polite spoken sentence (です・ます). */
function politeSentence(text) {
  let value = strip(text).replace(/[。．.、]+$/, "");
  if (!value) return "";
  const rules = [
    [/だった$/, "でした"], [/である$/, "です"], [/だ$/, "です"], [/ていた$/, "ていました"], [/ている$/, "ています"], [/ていく$/, "ていきます"],
    [/できる$/, "できます"], [/できた$/, "できました"], [/した$/, "しました"], [/する$/, "します"], [/ある$/, "あります"], [/あった$/, "ありました"],
    [/なった$/, "なりました"], [/なる$/, "なります"], [/いる$/, "います"], [/(ない|たい|しい|高い|低い|多い|少ない|大きい|小さい|早い|遅い|良い|よい)$/, "$1です"],
    [/った$/, "りました"], [/んだ$/, "みました"], [/いた$/, "きました"],
  ];
  for (const [pattern, replacement] of rules) {
    if (pattern.test(value)) return `${value.replace(pattern, replacement)}。`;
  }
  if (/る$/.test(value)) return `${value.slice(0, -1)}${E_ROW.includes(value.at(-2)) ? "ます" : "ります"}。`;
  if (/[うくぐすつぬぶむ]$/.test(value)) {
    const stem = { う: "い", く: "き", ぐ: "ぎ", す: "し", つ: "ち", ぬ: "に", ぶ: "び", む: "み" }[value.at(-1)];
    return `${value.slice(0, -1)}${stem}ます。`;
  }
  if (/(共有|確認|実施|導入|展開|移行|習得|改善|削減|推進|検討|整備|提供|標準化|強化|支援)$/.test(value) && /[をに]/.test(value)) return `${value}します。`;
  if (/(ます|でした|ました)$/.test(value)) return `${value}。`;
  return `${value}です。`;
}

const topicOf = (title) => strip(title).replace(/^(結論|要点|まとめ|提案)[：:]\s*/, "");

/** Instant, offline speaker script built from the slide's own content. */
function quickNotes(index) {
  const deck = state.deck;
  const slide = deck.slides[index];
  const next = deck.slides[index + 1];
  const title = topicOf(slide.title);
  const bridge = !next ? "" : next.type === "closing" ? "最後に、お願いしたいことをお伝えします。" : `次に、${topicOf(next.title || next.text)}についてご説明します。`;
  if (slide.type === "title") {
    return `本日は「${strip(slide.title)}」についてご説明します。${slide.subtitle ? `テーマは、${strip(slide.subtitle).replace(/[。．]$/, "")}です。` : ""}${deck.purpose ? `${deck.purpose}のためのご報告です。` : ""}よろしくお願いいたします。`;
  }
  if (slide.type === "section") return `ここからは「${title}」についてです。${politeSentence(slide.takeaway)}${bridge}`;
  if (slide.type === "closing") return `以上を踏まえて、次のアクションです。${politeSentence(String(slide.message || "").replace(/\n/g, "、"))}ご検討をよろしくお願いいたします。`;
  if (slide.type === "statement") return `${politeSentence(slide.text)}${slide.takeaway ? politeSentence(slide.takeaway) : ""}${bridge}`;
  if (slide.type === "hero") return `${politeSentence(slide.title)}${slide.takeaway ? politeSentence(slide.takeaway) : ""}${bridge}`;
  const order = ["まず", "次に", "また", "さらに"];
  let body;
  if (["kpi", "dashboard"].includes(slide.type)) {
    body = (slide.items || []).slice(0, 4).map((item, i) => `${order[i]}、${strip(item.label)}は${strip(item.value)}です${item.change ? `（${strip(item.change)}）` : ""}。`).join("");
  } else {
    const units = extractUnits(slide).filter((unit) => unit.title && !PLACEHOLDERS.has(unit.title)).slice(0, 4);
    body = units.map((unit, i) => `${order[i]}、${unit.title}${unit.desc ? `については、${politeSentence(unit.desc)}` : "です。"}`).join("");
  }
  const lead = slide.takeaway ? `結論から申し上げますと、${politeSentence(slide.takeaway)}` : "";
  return `このスライドでは、${title}についてお伝えします。${lead}${body}${bridge}`;
}

function applyNotes(notes) {
  pushUndo();
  for (const { slide, text } of notes) if (state.deck.slides[slide]) { state.deck.slides[slide].notes = text.slice(0, 1200); delete state.deck.slides[slide].notesRich; }
  markChanged({ structural: true });
}

function aiNotes(indices, { seconds = 60, tone = "丁寧" } = {}, onStatus = () => {}) {
  return new Promise((resolve) => {
    if (state.aiBusy) return resolve(false);
    setAiBusy(true);
    jsonFetch("/api/decks/notes", { method: "POST", body: JSON.stringify({ deck: serverDeck(), indices, seconds, tone }) })
      .then(({ jobId }) => watchJob(jobId, {
        onProgress: (job) => onStatus(`${job.stage}：${job.detail}`),
        onDone: (job) => { setAiBusy(false); applyNotes(job.notes || []); onStatus(job.detail); resolve(true); },
        onFail: (job) => { setAiBusy(false); onStatus(`作成できませんでした：${job.error || job.detail}`); resolve(false); },
      }))
      .catch((error) => { setAiBusy(false); onStatus(`作成できませんでした：${error.message}`); resolve(false); });
  });
}

/** 校閲 → 翻訳: the words of this slide (or of the whole deck) in another language, put back in one undo step. */
function translateDeck(to, scope = "slide") {
  if (!state.deck) return;
  if (!state.codexAuthorized) { toast("翻訳にはCodexへの接続が必要です（作成画面の上部から接続）"); return; }
  if (state.aiBusy) return;
  const indices = scope === "deck" ? state.deck.slides.map((_, i) => i) : [state.selected];
  const wanted = (t) => typeof t === "string" && t.trim() && /\p{L}/u.test(t);
  const found = new Set();
  const collect = (t) => { if (wanted(t)) found.add(t); return t; };
  const targets = indices.map((i) => state.deck.slides[i]);
  eachText(targets, collect);
  eachObjectText(clone(targets), collect);
  if (scope === "deck") collect(state.deck.title || "");
  const texts = [...found].slice(0, 600);
  if (!texts.length) { toast("訳す文字がありません"); return; }
  const name = { en: "英語", zh: "中国語", ko: "韓国語", ja: "日本語" }[to];
  setAiBusy(true);
  toast(`${scope === "deck" ? "資料全体" : "このスライド"}を${name}に翻訳しています…`);
  jsonFetch("/api/decks/translate", { method: "POST", body: JSON.stringify({ texts, to }) })
    .then(({ jobId }) => watchJob(jobId, {
      onProgress: () => {},
      onDone: (job) => {
        setAiBusy(false);
        const map = new Map(texts.map((t, i) => [t, job.texts?.[i] ?? t]));
        const swap = (t) => map.get(t) ?? t;
        pushUndo();
        const next = eachText(indices.map((i) => state.deck.slides[i]), swap);
        indices.forEach((i, k) => { state.deck.slides[i] = next[k]; });
        eachObjectText(indices.map((i) => state.deck.slides[i]), swap);
        if (scope === "deck" && map.has(state.deck.title)) { state.deck.title = map.get(state.deck.title); $("deckTitleInput").value = state.deck.title; }
        markChanged({ structural: true });
        toast(`${name}に翻訳しました（${texts.length}か所。⌘Zで戻せます）`);
      },
      onFail: (job) => { setAiBusy(false); toast(`翻訳できませんでした：${job.error || job.detail}`); },
    }))
    .catch((error) => { setAiBusy(false); toast(`翻訳できませんでした：${error.message}`); });
}

function openNotesDialog() {
  const canAi = state.codexAuthorized;
  const ai = document.querySelector("input[name=notesMethod][value=ai]");
  ai.disabled = !canAi;
  if (!canAi) document.querySelector("input[name=notesMethod][value=quick]").checked = true;
  $("notesStatus").textContent = canAi ? "AIは結論→根拠→次のスライドへのつなぎの順で、スライドにある事実だけを使って書きます。" : "Codexに接続するとAIで自然な原稿を作れます。今は簡易作成を使えます。";
  $("notesRun").disabled = false;
  $("notesDialog").showModal();
}

async function runNotes() {
  const all = document.querySelector("input[name=notesTarget]:checked").value === "all";
  const method = document.querySelector("input[name=notesMethod]:checked").value;
  const seconds = Number(document.querySelector("input[name=notesSeconds]:checked").value);
  const tone = $("notesTone").value;
  const indices = state.deck.slides.map((slide, index) => (all || !strip(slide.notes) ? index : -1)).filter((index) => index >= 0);
  if (!indices.length) { $("notesStatus").textContent = "ノートが空のスライドはありません。「すべて作り直す」を選んでください。"; return; }
  if (method === "quick") {
    applyNotes(indices.map((index) => ({ slide: index, text: quickNotes(index) })));
    $("notesDialog").close();
    toast(`${indices.length}枚分のノートを作成しました（⌘Zで元に戻せます）`);
    return;
  }
  $("notesRun").disabled = true;
  const ok = await aiNotes(indices, { seconds, tone }, (text) => { $("notesStatus").textContent = text; });
  $("notesRun").disabled = false;
  if (ok) { $("notesDialog").close(); toast(`${indices.length}枚分のノートを作成しました（⌘Zで元に戻せます）`); }
}

// ---------------------------------------------------------------- find & replace

// 検索・置換: the options (大文字と小文字・単語・全角と半角) are kept in this browser; the matching itself is editor/find.mjs.
const FIND_KEY = "hpptx-find";
function findOptions() {
  return { needle: $("findInput").value, caseSensitive: $("findCase").checked, wholeWord: $("findWord").checked, width: !$("findWidth").checked };
}
function loadFindOptions() {
  try {
    const saved = JSON.parse(localStorage.getItem(FIND_KEY) || "{}");
    $("findCase").checked = saved.caseSensitive === true;
    $("findWord").checked = saved.wholeWord === true;
    $("findWidth").checked = saved.distinguishWidth === true;
  } catch { /* the usual options */ }
}
function saveFindOptions() {
  try { localStorage.setItem(FIND_KEY, JSON.stringify({ caseSensitive: $("findCase").checked, wholeWord: $("findWord").checked, distinguishWidth: $("findWidth").checked })); } catch { /* not kept */ }
}
/** The slides a search covers: all of them, or only the one on the stage. */
const findScope = () => ($("findScope").value === "slide" ? [state.selected] : (state.deck?.slides || []).map((_, i) => i));

/** Where the text is: for each slide in scope the number of hits and the words round the first. */
function findHits(m) {
  const hits = [];
  if (!m || !state.deck) return hits;
  for (const index of findScope()) {
    const slide = state.deck.slides[index];
    if (!slide) continue;
    let count = 0;
    let first = "";
    const visit = (text) => {
      const ranges = m.ranges(text);
      if (ranges.length) { count += ranges.length; first ||= findSnippet(text, ranges[0]); }
      return text;
    };
    eachText([slide], visit);
    eachObjectText([slide], visit);
    if (count) hits.push({ index, count, text: first });
  }
  // スライド マスター: the words of its objects are searched with the deck (index -1).
  if ($("findScope").value === "deck" && state.deck.masterObjects?.length) {
    let count = 0;
    let first = "";
    eachObjectText([{ elements: state.deck.masterObjects }], (text) => { const ranges = m.ranges(text); if (ranges.length) { count += ranges.length; first ||= findSnippet(text, ranges[0]); } return text; });
    if (count) hits.push({ index: -1, count, text: first });
  }
  return hits;
}
const findInTitle = (m) => ($("findScope").value === "deck" && m ? m.count(state.deck?.title || "") : 0);

function updateFindCount() {
  const m = findMatcher(findOptions());
  const hits = findHits(m);
  const total = hits.reduce((sum, hit) => sum + hit.count, 0) + findInTitle(m);
  $("findCount").textContent = !m ? "資料タイトル・本文・ノート・詳細を対象にします。" : total ? `${total}か所見つかりました（${hits.length}枚のスライド）` : "見つかりませんでした";
  $("replaceAllBtn").disabled = !total;
  $("findHits").replaceChildren(...hits.slice(0, 200).map((hit) => h("li", {}, h("button", { type: "button", class: "find-hit", title: `${hit.index + 1}枚目へ移動`, onclick: () => { $("replaceDialog").close(); if (hit.index < 0) openMaster(); else select(hit.index); } },
    h("b", {}, hit.index < 0 ? "マスター" : `${hit.index + 1}枚目`), ` ${hit.text}`, hit.count > 1 ? h("small", {}, `（${hit.count}か所）`) : null))));
}

/** Every word of the deck (slides, objects, notes, details) through `fn`: one undo step. */
function transformAllText(fn) {
  if (!state.deck) return;
  pushUndo();
  state.deck.slides = eachText(state.deck.slides, fn);
  eachObjectText(state.deck.slides, fn);
  markChanged({ structural: true });
}
const spellcheckOn = () => { try { return localStorage.getItem("hpptx-spellcheck") === "1"; } catch { return false; } };

function replaceAll() {
  const m = findMatcher(findOptions());
  if (!m) return;
  const replacement = $("replaceInput").value;
  const hits = findHits(m);
  const inTitle = findInTitle(m);
  const count = hits.reduce((sum, hit) => sum + hit.count, 0) + inTitle;
  if (!count) return;
  pushUndo();
  const swap = (text) => m.replace(text, replacement);
  for (const hit of hits) {
    if (hit.index < 0) { eachObjectText([{ elements: state.deck.masterObjects }], swap); continue; }
    state.deck.slides[hit.index] = eachText([state.deck.slides[hit.index]], swap)[0];
    eachObjectText([state.deck.slides[hit.index]], swap);
  }
  if (inTitle) { state.deck.title = m.replace(state.deck.title || "", replacement); $("deckTitleInput").value = state.deck.title; }
  $("replaceDialog").close();
  markChanged({ structural: true });
  toast(`${count}か所を置き換えました（⌘Zで元に戻せます）`);
}

function openReplace() {
  loadFindOptions();
  updateFindCount();
  $("replaceDialog").showModal();
  $("findInput").focus();
}

// ---------------------------------------------------------------- issue summary

function renderIssueSummary() {
  const chip = $("issueSummary");
  if (!state.deck) return;
  const pending = state.deck.slides.some((_, index) => !measured(index));
  const overflowSlides = state.deck.slides.map((_, i) => (overflowFor(i).length ? i : -1)).filter((i) => i >= 0);
  const lint = lintDeck().flat();
  const lintWarnings = lint.filter((issue) => issue.severity !== "info").length;
  const parts = [];
  if (overflowSlides.length) parts.push(`${overflowSlides.length}枚に文字あふれ`);
  if (lint.length) parts.push(`構成の指摘${lint.length}件`);
  chip.className = `issue-chip ${overflowSlides.length ? "error" : lintWarnings ? "warn" : "ok"}`;
  chip.textContent = parts.length ? `⚠ ${parts.join("・")}` : pending ? "チェック中…" : "✓ チェックOK";
  chip.title = parts.length ? "クリックで一覧を表示（移動・AIで直す）" : "文字あふれ・構成の問題は見つかりませんでした";
  $("fixAllBtn").classList.toggle("hidden", !overflowSlides.length || !state.codexAuthorized);
  $("fixAllBtn").disabled = state.aiBusy;
}

// ---------------------------------------------------------------- inspector

function refreshInspectorIssues() {
  const container = document.getElementById("inspectorIssues");
  if (!container) return;
  container.replaceWith(issueSection());
  const issues = issuesFor(state.selected);
  document.querySelectorAll("#inspector [data-path]").forEach((input) => {
    const path = input.dataset.path;
    const hit = issues.find((issue) => issue.severity !== "info" && (issue.field === path || issue.field.startsWith(`${path}[`) || issue.field.startsWith(`${path}.`) || path.startsWith(`${issue.field}.`)));
    input.style.borderColor = hit ? (hit.kind === "overflow" ? "var(--danger)" : "var(--warn)") : "";
  });
}

function issueSection() {
  const overflow = overflowFor(state.selected);
  const lint = lintDeck()[state.selected] ?? [];
  if (!overflow.length && !lint.length) return h("div", { id: "inspectorIssues" });
  const canAi = state.codexAuthorized;
  const item = (issue) => h("li", { class: issue.severity === "error" ? "error" : issue.severity === "info" ? "info" : "", onclick: () => focusField(issue.field), title: "クリックで該当欄へ" },
    h("span", {}, `${issue.plain ? "" : `${fieldLabel(issue.field)}：`}${issue.message}`),
    issue.fix ? h("button", { class: "btn issue-fix", type: "button", onclick: (event) => { event.stopPropagation(); issue.fix(); } }, "直す") : null);
  return h("div", { id: "inspectorIssues", class: "section" },
    overflow.length ? [
      h("div", { class: "section-title" }, h("span", {}, "文字あふれ")),
      h("ul", { class: "issue-list" }, overflow.map(item)),
      h("button", { class: "btn btn-ai", type: "button", disabled: !canAi || state.aiBusy, onclick: () => reviseSlide(state.selected, "スライドに収まらない文字を、意味を保って短く言い換える（補足は details に回してよい）", overflow), title: canAi ? "" : "Codexに接続すると使えます" }, "✦ AIで収まるように直す"),
      h("span", { class: "hint", style: { "margin-left": "8px" } }, "または赤枠の欄を短くしてください"),
    ] : null,
    lint.length ? [
      h("div", { class: "section-title", style: { "margin-top": overflow.length ? "12px" : "0" } }, h("span", {}, "構成チェック")),
      h("ul", { class: "issue-list" }, lint.map(item)),
    ] : null);
}

function fieldLabel(field) {
  const slide = state.deck.slides[state.selected];
  const root = String(field).split(/[.[]/)[0];
  const spec = specFor(slide.type).find((item) => item.key === root);
  const index = String(field).match(/\[(\d+)\]/);
  return `${spec?.label?.replace(/（.*$/, "") ?? field}${index ? ` ${Number(index[1]) + 1}` : ""}`;
}

function counter(value, max) {
  if (!max) return null;
  const count = Math.ceil(charCount(value));
  return h("span", { class: `counter${count > max ? " over" : ""}` }, `${count}/${max}字`);
}

function inputFor(field, path, slide) {
  const value = getPath(slide, path);
  const key = pathKey(path);
  const onText = (event) => {
    beginEdit();
    setPath(slide, path, event.target.value);
    const count = event.target.closest(".field")?.querySelector(".counter");
    if (count && field.max) { const n = Math.ceil(charCount(event.target.value)); count.textContent = `${n}/${field.max}字`; count.classList.toggle("over", n > field.max); }
    markChanged();
  };
  const labelRow = (extra) => h("div", { class: "label-row" }, h("label", { for: `f-${key}` }, field.label), extra);
  switch (field.kind) {
    case "text":
      return h("div", { class: "field" }, labelRow(counter(value, field.max)), h("input", { id: `f-${key}`, type: "text", value: value ?? "", placeholder: field.placeholder || "", "data-path": key, oninput: onText }));
    case "area":
      return h("div", { class: "field" }, labelRow(counter(value, field.max)), h("textarea", { id: `f-${key}`, rows: field.rows || 2, "data-path": key, oninput: onText }, value ?? ""));
    case "number":
      return h("div", { class: "field" }, labelRow(), h("input", { id: `f-${key}`, type: "number", min: field.min, max: field.max, step: field.free ? "any" : null, value: value ?? "", "data-path": key, oninput: (event) => { beginEdit(); setPath(slide, path, event.target.value === "" ? undefined : Number(event.target.value)); markChanged(); } }));
    case "icon":
      return iconField(field, path, slide, value);
    case "check":
      return h("label", { class: "inline-check field" }, h("input", { type: "checkbox", checked: Boolean(value), onchange: (event) => { pushUndo(); setPath(slide, path, event.target.checked || undefined); markChanged({ structural: true }); } }), field.label);
    case "select":
      return h("div", { class: "field" }, labelRow(), h("select", { id: `f-${key}`, "data-path": key, onchange: (event) => { pushUndo(); setPath(slide, path, event.target.value === "" ? undefined : field.number ? Number(event.target.value) : event.target.value); markChanged({ structural: true }); } },
        field.options.map(([optionValue, text]) => h("option", { value: optionValue, selected: String(value ?? "") === String(optionValue) || (value == null && optionValue === field.options[0][0]) }, text))));
    case "list": {
      const lines = Array.isArray(value) ? value : [];
      const over = field.maxChars ? lines.filter((line) => charCount(line) > field.maxChars).length : 0;
      return h("div", { class: "field" },
        labelRow(h("span", { class: `counter${lines.length > (field.max ?? 99) ? " over" : ""}` }, `${lines.length}${field.max ? `/${field.max}` : ""}項目${over ? `・長い行${over}` : ""}`)),
        h("textarea", { id: `f-${key}`, rows: Math.max(3, Math.min(10, lines.length + 1)), "data-path": key, oninput: (event) => { beginEdit(); setPath(slide, path, event.target.value.split("\n").map((line) => line.replace(/^[・•\-*]\s*/, ""))); markChanged(); } }, lines.join("\n")),
        field.hint ? h("div", { class: "hint" }, field.hint) : null);
    }
    case "group": {
      const list = Array.isArray(value) ? value : [];
      const groups = list.map((item, i) => h("div", { class: "form-group" },
        h("div", { class: "form-group-head" }, h("span", {}, `${field.label.replace(/（.*$/, "")} ${i + 1}`), h("span", { class: "btns" },
          h("button", { class: "btn", type: "button", title: "上へ", disabled: i === 0, onclick: () => { pushUndo(); [list[i - 1], list[i]] = [list[i], list[i - 1]]; markChanged({ structural: true }); } }, "↑"),
          h("button", { class: "btn", type: "button", title: "下へ", disabled: i === list.length - 1, onclick: () => { pushUndo(); [list[i + 1], list[i]] = [list[i], list[i + 1]]; markChanged({ structural: true }); } }, "↓"),
          h("button", { class: "btn btn-danger", type: "button", title: "削除", disabled: list.length <= (field.min ?? 0), onclick: () => { pushUndo(); list.splice(i, 1); markChanged({ structural: true }); } }, "✕"))),
        field.fields.map((sub) => inputFor(sub, [...path, i, sub.key], slide))));
      return h("div", { class: "field" }, h("div", { class: "field-label" }, `${field.label}（${list.length}${field.max ? `/${field.max}` : ""}）`), groups,
        list.length < (field.max ?? 99) ? h("button", { class: "btn add-item", type: "button", onclick: () => { pushUndo(); setPath(slide, path, [...list, field.newItem ? field.newItem() : {}]); markChanged({ structural: true }); } }, `＋ ${field.label.replace(/（.*$/, "")}を追加`) : null);
    }
    case "table": return tableEditor(slide);
    case "chart": return chartEditor(field, slide);
    default: return null;
  }
}

function tableEditor(slide) {
  slide.headers ??= ["項目", "内容"];
  slide.rows ??= [["", ""]];
  const cols = slide.headers.length;
  const rowStyle = { "grid-template-columns": `repeat(${cols}, minmax(70px, 1fr))` };
  const cell = (value, onInput, head = false) => h("input", { type: "text", value: value ?? "", "data-path": head ? "headers" : "rows", oninput: (event) => { beginEdit(); onInput(event.target.value); markChanged(); } });
  const structural = (fn) => () => { pushUndo(); fn(); markChanged({ structural: true }); };
  return h("div", { class: "field" },
    h("div", { class: "field-label" }, `表（${cols}列 × ${slide.rows.length}行）`),
    h("div", { class: "table-editor" },
      h("div", { class: "row head", style: rowStyle }, slide.headers.map((value, ci) => cell(value, (v) => { slide.headers[ci] = v; }, true))),
      slide.rows.map((row) => h("div", { class: "row", style: rowStyle }, slide.headers.map((_, ci) => cell(row[ci], (v) => { row[ci] = v; }))))),
    h("div", { class: "table-tools" },
      h("button", { class: "btn", type: "button", disabled: slide.rows.length >= 8, onclick: structural(() => slide.rows.push(slide.headers.map(() => ""))) }, "＋ 行"),
      h("button", { class: "btn", type: "button", disabled: slide.rows.length <= 1, onclick: structural(() => slide.rows.pop()) }, "− 行"),
      h("button", { class: "btn", type: "button", disabled: cols >= 6, onclick: structural(() => { slide.headers.push("列"); slide.rows.forEach((row) => row.push("")); }) }, "＋ 列"),
      h("button", { class: "btn", type: "button", disabled: cols <= 2, onclick: structural(() => { slide.headers.pop(); slide.rows.forEach((row) => row.pop()); }) }, "− 列")),
    h("div", { class: "hint" }, "最大6列×8行。数値の列は右ぞろえになります。行にマウスを乗せると発表中に強調されます。"));
}

function chartToCsv(image) {
  const data = image?.data || {};
  // A ranking by view: one column per view. Before → after: two columns.
  if (Array.isArray(data.views) && data.views.length) {
    const labels = [...new Set(data.views.flatMap((view) => (view.items || []).map((item) => item.label)))];
    return [["ラベル", ...data.views.map((view) => view.label || "切り口")].join(","), ...labels.map((label) => [label, ...data.views.map((view) => view.items?.find((item) => item.label === label)?.value ?? "")].join(","))].join("\n");
  }
  if (image?.chartType === "shift") {
    return [["ラベル", data.beforeLabel || "前", data.afterLabel || "後"].join(","), ...(data.items || []).map((item) => [item.label, item.before ?? "", item.value ?? ""].join(","))].join("\n");
  }
  if (Array.isArray(data.series) && data.series.length) {
    const labels = data.xAxisLabels || [];
    return [["ラベル", ...data.series.map((item) => item.label || "系列")].join(","), ...labels.map((label, i) => [label, ...data.series.map((item) => item.values?.[i] ?? "")].join(","))].join("\n");
  }
  if (Array.isArray(data.barData) && data.barData.length) {
    return [["ラベル", ...(data.legendLabels || [])].join(","), ...data.barData.map((row) => [row.label, ...(row.values || [])].join(","))].join("\n");
  }
  return (data.items || []).map((item) => `${item.label},${item.value ?? item.barValue ?? ""}`).join("\n");
}

function csvToChartData(csv, chartType, title, keep = {}) {
  const rows = csv.split("\n").map((line) => line.split(/[,\t，]/).map((cell) => cell.trim())).filter((row) => row.length && row[0] !== "");
  const numeric = (value) => Number(String(value).replace(/[^\d.\-]/g, ""));
  const multi = rows.length && rows[0].length > 2 && rows[0].slice(1).some((cell) => Number.isNaN(numeric(cell)) || cell === "");
  const data = { ...(title ? { title } : {}), ...(keep.unit ? { unit: keep.unit } : {}), ...(keep.highlight && chartType === "rank" ? { highlight: keep.highlight } : {}) };
  const hasHead = rows.length > 1 && rows[0].slice(1).some((cell) => cell !== "" && Number.isNaN(Number(cell.replace(/[,，]/g, ""))));
  const value = (cell) => (cell === "" || cell == null || Number.isNaN(numeric(cell)) ? undefined : numeric(cell));
  if (chartType === "rank") {
    const [head, ...body] = hasHead ? rows : [["ラベル", ...rows[0].slice(1).map((_, i) => `切り口${i + 1}`)], ...rows];
    data.views = head.slice(1, 5).map((label, i) => ({ label: label || `切り口${i + 1}`, items: body.slice(0, 12).map((row) => ({ label: row[0], value: value(row[i + 1]) })).filter((item) => item.value !== undefined) }));
    if (!data.views.length) data.views = [{ label: "全体", items: [] }];
    return data;
  }
  if (chartType === "shift") {
    const [head, ...body] = hasHead ? rows : [["ラベル", "前", "後"], ...rows];
    data.beforeLabel = head[1] || "前";
    data.afterLabel = head[2] || "後";
    data.items = body.slice(0, 8).map((row) => ({ label: row[0], before: value(row[1]) ?? 0, value: value(row[2]) ?? 0 }));
    return data;
  }
  if (multi) {
    const [head, ...body] = rows;
    const names = head.slice(1);
    if (chartType === "stacked-bar" || chartType === "100-stacked-bar") {
      data.legendLabels = names;
      data.barData = body.slice(0, 20).map((row) => ({ label: row[0], values: names.map((_, i) => numeric(row[i + 1]) || 0) }));
    } else {
      data.xAxisLabels = body.slice(0, 20).map((row) => row[0]);
      data.series = names.slice(0, 8).map((name, i) => ({ label: name, values: body.slice(0, 20).map((row) => numeric(row[i + 1]) || 0) }));
    }
  } else {
    data.items = rows.slice(0, 20).map((row) => ({ label: row[0], value: numeric(row[1]) || 0 }));
  }
  return data;
}

function chartEditor(field, slide) {
  const image = slide[field.key];
  const mode = image && typeof image === "object" ? "chart" : typeof image === "string" && image.startsWith("data:") ? "picture" : field.allowPhoto ? "photo" : "none";
  const modes = [["chart", "グラフ"], ...(field.allowPhoto ? [["photo", "写真・動画"]] : [["none", "なし"]])];
  const setImageMode = (next) => {
    pushUndo();
    if (next === "chart") slide[field.key] = { chartType: "bar", data: { title: "", items: [{ label: "A", value: 10 }, { label: "B", value: 14 }] } };
    else delete slide[field.key];
    if (next === "photo" && slide.media?.placement && slide.type === "imageText") delete slide.media.placement;
    markChanged({ structural: true });
    if (next === "photo" && !slide.media && !slide.visualAsset) setTimeout(() => document.querySelector("#inspector [data-key=media]")?.scrollIntoView({ behavior: "smooth", block: "start" }), 60);
  };
  const body = [];
  if (mode === "chart") {
    const chartTypes = [["bar", "棒"], ["line", "折れ線"], ["donut", "ドーナツ"], ["multi-line", "複数の折れ線"], ["stacked-bar", "積み上げ棒"], ["100-stacked-bar", "100%積み上げ"], ["combo", "棒＋折れ線"], ["rank", "順位（切り口で並び替わる）"], ["shift", "前後の差（Before→After）"]];
    const update = (patch) => { beginEdit(); const current = slide[field.key]; slide[field.key] = { ...current, ...patch }; markChanged(); };
    const keep = () => ({ unit: slide[field.key].data?.unit, highlight: slide[field.key].data?.highlight });
    const setData = (key, value) => update({ data: { ...slide[field.key].data, [key]: value || undefined } });
    const hints = {
      rank: "1行目に「ラベル,切り口A,切り口B」（例：業種,全体,一次請け,二次請け）。発表中は切り口のボタンで全部の棒が並び替わります。値のない項目は最後に並びます。",
      shift: "1行目に「ラベル,前の名前,後の名前」（例：内訳,2024年度,2025年度）、2行目から「ラベル,前の値,後の値」。発表中は棒が前から後へ縮み（伸び）、減った分が点線で残ります。",
    };
    body.push(
      h("div", { class: "grid-2" },
        h("div", { class: "field" }, h("label", {}, "種類"), h("select", { onchange: (event) => { pushUndo(); const current = slide[field.key]; slide[field.key] = { chartType: event.target.value, data: csvToChartData(chartToCsv(current), event.target.value, current.data?.title, { unit: current.data?.unit, highlight: current.data?.highlight }) }; markChanged({ structural: true }); } }, chartTypes.map(([value, text]) => h("option", { value, selected: image.chartType === value }, text)))),
        h("div", { class: "field" }, h("label", {}, "グラフタイトル"), h("input", { type: "text", value: image.data?.title ?? "", oninput: (event) => update({ data: { ...slide[field.key].data, title: event.target.value || undefined } }) }))),
      ["rank", "shift"].includes(image.chartType) ? h("div", { class: "grid-2" },
        h("div", { class: "field" }, h("label", {}, "単位"), h("input", { type: "text", maxlength: 10, placeholder: "例：%・時間", value: image.data?.unit ?? "", oninput: (event) => setData("unit", event.target.value) })),
        image.chartType === "rank" ? h("div", { class: "field" }, h("label", {}, "主役（強調する項目名）"), h("input", { type: "text", maxlength: 80, placeholder: "空欄なら1位を強調", value: image.data?.highlight ?? "", oninput: (event) => setData("highlight", event.target.value) })) : null) : null,
      h("div", { class: "field" }, h("label", {}, "データ（CSV）"),
        h("textarea", { rows: 5, "data-path": field.key, style: { "font-family": "ui-monospace, Menlo, monospace", "font-size": "12px" }, oninput: (event) => update({ data: csvToChartData(event.target.value, slide[field.key].chartType, slide[field.key].data?.title, keep()) }) }, chartToCsv(image)),
        h("div", { class: "hint" }, hints[image.chartType] ?? "1行に「ラベル,値」。複数系列は1行目を「ラベル,系列A,系列B」にします。発表中はグラフが伸び、マウスを乗せると値が出ます。")));
  } else if (mode === "picture") {
    body.push(h("img", { src: image, alt: "", style: { width: "100%", "border-radius": "8px", border: "1px solid var(--line)" } }),
      h("button", { class: "btn", type: "button", onclick: () => { pushUndo(); delete slide[field.key]; markChanged({ structural: true }); } }, "この画像を外す"));
  } else if (mode === "photo") {
    body.push(h("div", { class: "hint" }, "下の「写真・動画」で選んだ写真・動画が図の場所に入ります。"));
  }
  return h("div", { class: "field" },
    h("div", { class: "field-label" }, field.label),
    h("div", { class: "segmented", role: "radiogroup" }, modes.map(([value, text]) => h("label", {}, h("input", { type: "radio", name: `chartmode-${field.key}`, value, checked: mode === value || (mode === "picture" && value === "photo"), onchange: () => setImageMode(value) }), h("span", {}, text)))),
    ...body);
}

function iconField(field, path, slide, value) {
  if (!ICON_TYPES.has(slide.type)) return null;
  const set = (name, event) => { event.currentTarget.closest("details")?.removeAttribute("open"); pushUndo(); setPath(slide, path, name || undefined); markChanged({ structural: true }); };
  const entries = Object.entries(E.icons);
  const grid = h("div", { class: "icon-grid" },
    h("button", { type: "button", class: value ? "" : "selected", title: "アイコンなし（番号を表示）", onclick: (event) => set("", event) }, "なし"),
    entries.map(([name, entry]) => h("button", { type: "button", class: value === name ? "selected" : "", title: entry.label, onclick: (event) => set(name, event) }, uiIcon(name))));
  return h("details", { class: "field icon-field", "data-key": pathKey(path) },
    h("summary", {}, h("span", { class: "field-label" }, field.label),
      value && E.icons[value] ? h("span", { class: "icon-current" }, uiIcon(value), E.icons[value].label) : h("span", { class: "hint" }, "なし（クリックで選択）")),
    grid);
}

// ---- photo & video

function mediaInfo(slide) {
  const desc = E.mediaOf(slide, {});
  if (!desc) return null;
  const libraryKey = desc.src.startsWith("asset:") ? desc.src.slice(6) : null;
  const url = libraryKey ? `/assets/${E.PHOTOS[libraryKey]?.[0]}` : desc.src.startsWith("idb:") ? mediaUrls[desc.src] || "" : desc.src;
  const label = slide.media?.name || (libraryKey ? `写真ライブラリ：${E.PHOTOS[libraryKey]?.[1] ?? ""}` : desc.kind === "youtube" ? "YouTube" : desc.kind === "lottie" ? "アニメーション" : slide.customImage ? "取り込んだ写真" : "写真");
  return { ...desc, url, label, libraryKey };
}

/** Put media on a slide: its own photo slot where the layout has one, otherwise floating where the user puts it. */
function setSlideMedia(slide, media, placement = null) {
  delete slide.customImage;
  delete slide.imagePlacement;
  delete slide.visualAsset;
  const chartSlot = slide.type === "imageText" && slide.image && typeof slide.image === "object";
  const needsPlace = !SLOTTED.has(slide.type) || chartSlot;
  slide.media = { ...media, ...(placement || needsPlace ? { placement: placement || { ...DEFAULT_PLACEMENT } } : {}) };
}

/** A Lottie animation (JSON from After Effects/Bodymovin or LottieFiles) → the checked text, or an error message. */
async function readLottie(blob) {
  const text = await blob.text();
  let data = null;
  try { data = JSON.parse(text); } catch { throw new Error("JSONとして読めません"); }
  if (!data || !Array.isArray(data.layers) || !Number.isFinite(Number(data.fr)) || !Number.isFinite(Number(data.op))) throw new Error("Lottieアニメーション（JSON）ではないようです");
  return text;
}

const isLottieFile = (file) => /\.json$/i.test(file?.name || "") || file?.type === "application/json";

async function addLottieFile(file, { placement = null, index = state.selected } = {}) {
  if (/\.lottie$/i.test(file.name || "")) return toast("「.lottie」形式は使えません。LottieFilesでは「Lottie JSON」を選んでダウンロードしてください");
  if (file.size > 20_000_000) return toast("アニメーションが大きすぎます（20MBまで）");
  try {
    const text = await readLottie(file);
    const src = await putMedia(new Blob([text], { type: "application/json" }), file.name);
    pushUndo();
    setSlideMedia(state.deck.slides[index], { src, kind: "lottie", name: file.name, autoplay: true, loop: true }, placement);
    markChanged({ structural: true });
    toast("アニメーションを入れました。発表中に再生されます（編集画面では途中の1コマを表示）");
  } catch (error) {
    toast(`読み込めませんでした：${error.message}`);
  }
}

async function addMediaFile(file, { placement = null, index = state.selected } = {}) {
  if (!file || !state.deck?.slides[index]) return;
  if (isLottieFile(file) || /\.lottie$/i.test(file.name || "")) return addLottieFile(file, { placement, index });
  const video = file.type.startsWith("video/");
  if (!video && !file.type.startsWith("image/")) return toast("写真（JPEG・PNG・WebP・GIF）か動画（MP4・WebM・MOV）を選んでください");
  if (video && file.size > 400_000_000) return toast("動画が大きすぎます（400MBまで）。短く切り出すか、YouTubeのURLを使ってください");
  try {
    toast(video ? "動画を保存しています…" : "写真を保存しています…");
    const blob = video || /gif|svg/.test(file.type) ? file : await downscaleImage(file, 2400, file.type === "image/png" ? "image/png" : "image/jpeg");
    const src = await putMedia(blob, file.name);
    pushUndo();
    setSlideMedia(state.deck.slides[index], { src, kind: video ? "video" : "image", name: file.name, ...(video ? { autoplay: true, loop: true, muted: true } : {}) }, placement);
    markChanged({ structural: true });
    toast(video ? "動画を入れました。発表中は自動で再生されます（音は「編集」→「写真・動画」で）" : "写真を入れました");
  } catch (error) {
    toast(`読み込めませんでした：${error.message}`);
  }
}

async function applyMediaUrl() {
  const url = $("mediaUrlInput").value.trim();
  if (!/^https:\/\//i.test(url)) return showStatus("mediaUrlStatus", "https:// で始まるURLを入れてください", "error");
  if (/\.json(\?|#|$)/i.test(url) || /^https:\/\/(lottie\.host|[a-z0-9-]+\.lottiefiles\.com)\//i.test(url)) {
    // Lottie from the web is copied into this browser, so the presentation and exported files work offline.
    showStatus("mediaUrlStatus", "アニメーションを読み込んでいます…");
    try {
      const response = await fetch(url);
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const text = await readLottie(await response.blob());
      const name = decodeURIComponent(url.split("/").pop().split("?")[0]).slice(0, 60) || "アニメーション";
      const src = await putMedia(new Blob([text], { type: "application/json" }), name);
      pushUndo();
      setSlideMedia(state.deck.slides[state.selected], { src, kind: "lottie", name, autoplay: true, loop: true });
      markChanged({ structural: true });
      $("mediaUrlDialog").close();
      toast("アニメーションを入れました。発表中に再生されます");
    } catch (error) {
      showStatus("mediaUrlStatus", `読み込めませんでした（${error.message}）。LottieFilesの「Lottie JSON」のURL（lottie.host）を使うか、JSONファイルをダウンロードして「アニメーション」から入れてください`, "error");
    }
    return;
  }
  const yt = E.youtubeId(url);
  const video = yt || /\.(mp4|webm|mov|m4v)(\?|#|$)/i.test(url);
  pushUndo();
  setSlideMedia(state.deck.slides[state.selected], { src: url, kind: video ? "video" : "image", name: yt ? "YouTube" : url.split("/").pop().split("?")[0].slice(0, 60), ...(video ? { autoplay: !yt, loop: !yt, muted: !yt } : {}) });
  markChanged({ structural: true });
  $("mediaUrlDialog").close();
  toast(yt ? "YouTubeを入れました。発表の画面で再生できます" : "URLの写真・動画を入れました");
}

function chooseLibraryPhoto(slide, key) {
  pushUndo();
  delete slide.customImage;
  delete slide.imagePlacement;
  const chartSlot = slide.type === "imageText" && slide.image && typeof slide.image === "object";
  if (!key) { delete slide.visualAsset; delete slide.media; }
  else if (SLOTTED.has(slide.type) && !chartSlot && !slide.media?.placement) { slide.visualAsset = key; delete slide.media; }
  else { delete slide.visualAsset; slide.media = { src: `asset:${key}`, kind: "image", name: E.PHOTOS[key][1], placement: slide.media?.placement || { ...DEFAULT_PLACEMENT } }; }
  markChanged({ structural: true });
}

function mediaSection(slide) {
  const info = mediaInfo(slide);
  const media = slide.media;
  const slotted = SLOTTED.has(slide.type) && !(slide.type === "imageText" && slide.image && typeof slide.image === "object");
  const photoInput = h("input", { type: "file", class: "hidden", accept: "image/png,image/jpeg,image/webp,image/gif", onchange: (event) => { addMediaFile(event.target.files?.[0]); event.target.value = ""; } });
  const videoInput = h("input", { type: "file", class: "hidden", accept: "video/mp4,video/webm,video/quicktime,video/*", onchange: (event) => { addMediaFile(event.target.files?.[0]); event.target.value = ""; } });
  const lottieInput = h("input", { type: "file", class: "hidden", accept: ".json,application/json,.lottie", onchange: (event) => { addMediaFile(event.target.files?.[0]); event.target.value = ""; } });
  const set = (fn) => { pushUndo(); fn(); markChanged({ structural: true }); };
  const nodes = [];
  if (info) {
    const lottieBox = info.kind === "lottie" && info.url ? h("div", { class: "lottie-thumb" }, h("div", { class: "hs-lottie-host", "data-src": info.url, "data-loop": "", "data-autoplay": "" })) : null;
    const preview = info.kind === "youtube"
      ? h("img", { src: `https://i.ytimg.com/vi/${info.yt}/mqdefault.jpg`, alt: "" })
      : info.kind === "lottie" ? lottieBox || h("span", {}, "アニメーションがありません")
        : info.kind === "video" ? (info.url ? h("video", { src: info.url, muted: true, preload: "metadata" }) : h("span", {}, "動画がありません"))
          : info.url ? h("img", { src: info.url, alt: "" }) : h("span", {}, "写真がありません");
    if (lottieBox) requestAnimationFrame(() => E.mountLottie(lottieBox, { play: true }));
    nodes.push(h("div", { class: "media-now" },
      h("div", { class: "thumb" }, preview, h("span", { class: "kind" }, { youtube: "YouTube", video: "動画", lottie: "アニメーション" }[info.kind] ?? "写真")),
      h("div", { class: "info" },
        h("b", { title: info.label }, info.label),
        h("span", { class: "hint" }, info.placement ? "自由に配置（スライド上でドラッグ）" : "レイアウトの写真枠に表示"),
        h("div", { class: "row" },
          slotted && (media || slide.customImage) ? h("button", { class: "btn btn-sm", type: "button", onclick: () => set(() => {
            if (media) { if (media.placement) delete media.placement; else media.placement = { ...DEFAULT_PLACEMENT }; }
            else if (slide.imagePlacement) delete slide.imagePlacement; else slide.imagePlacement = { ...DEFAULT_PLACEMENT };
          }) }, info.placement ? "枠に戻す" : "自由に配置") : null,
          slotted && info.libraryKey && !media ? h("button", { class: "btn btn-sm", type: "button", onclick: () => set(() => { slide.media = { src: `asset:${info.libraryKey}`, kind: "image", name: E.PHOTOS[info.libraryKey][1], placement: { ...DEFAULT_PLACEMENT } }; delete slide.visualAsset; }) }, "自由に配置") : null,
          h("button", { class: "btn btn-sm btn-danger", type: "button", onclick: () => set(() => { delete slide.media; delete slide.customImage; delete slide.imagePlacement; delete slide.visualAsset; }) }, "外す")))));
    const options = [];
    if (media && info.kind === "lottie") {
      options.push(
        h("label", { class: "inline-check" }, h("input", { type: "checkbox", checked: media.autoplay !== false, onchange: (event) => set(() => { media.autoplay = event.target.checked; }) }), "スライドが出たら自動で再生（オフならクリックで再生）"),
        h("label", { class: "inline-check" }, h("input", { type: "checkbox", checked: media.loop !== false, onchange: (event) => set(() => { media.loop = event.target.checked; }) }), "くり返し再生"),
        h("label", { class: "inline-check" }, h("input", { type: "checkbox", checked: media.fit === "cover", onchange: (event) => set(() => { if (event.target.checked) media.fit = "cover"; else delete media.fit; }) }), "枠いっぱいに広げる（はみ出た部分は切れる）"));
    } else if (media) {
      options.push(h("label", { class: "inline-check" }, h("input", { type: "checkbox", checked: media.fit === "contain", onchange: (event) => set(() => { if (event.target.checked) media.fit = "contain"; else delete media.fit; }) }), "切り取らずに全体を見せる"));
    }
    if (media && (info.kind === "video" || info.kind === "youtube")) {
      options.push(
        h("label", { class: "inline-check" }, h("input", { type: "checkbox", checked: media.autoplay !== false, onchange: (event) => set(() => { media.autoplay = event.target.checked; }) }), "スライドが出たら自動で再生"),
        info.kind === "video" ? h("label", { class: "inline-check" }, h("input", { type: "checkbox", checked: media.loop !== false, onchange: (event) => set(() => { media.loop = event.target.checked; }) }), "くり返し再生") : null,
        h("label", { class: "inline-check" }, h("input", { type: "checkbox", checked: media.muted === false, onchange: (event) => set(() => { media.muted = !event.target.checked; }) }), "音を出す（Zoomでは「音声を共有」をオン）"));
    }
    if (info.kind !== "youtube" && info.kind !== "lottie") {
      options.push(h("label", { class: "field", style: { "margin-bottom": "0", "min-width": "180px" } }, h("span", { class: "field-label" }, "写真の動き（発表中）"),
        h("select", { onchange: (event) => set(() => { if (event.target.value === "none") delete slide.photoMotion; else slide.photoMotion = event.target.value; }) },
          PHOTO_MOTIONS.map(([value, text]) => h("option", { value, selected: (slide.photoMotion || "none") === value }, text)))));
    }
    if (options.length) nodes.push(h("div", { class: "opt-row" }, options), h("div", { style: { height: "10px" } }));
  }
  nodes.push(h("div", { class: "media-actions" },
    h("button", { class: "btn", type: "button", onclick: () => photoInput.click() }, uiIcon("camera"), "写真"),
    h("button", { class: "btn", type: "button", onclick: () => videoInput.click() }, uiIcon("video"), "動画"),
    h("button", { class: "btn", type: "button", title: "After EffectsやLottieFilesのアニメーション（Lottie JSON）", onclick: () => lottieInput.click() }, uiIcon("sparkles"), "アニメーション"),
    h("button", { class: "btn", type: "button", onclick: () => { $("mediaUrlInput").value = ""; clearStatus("mediaUrlStatus"); $("mediaUrlDialog").showModal(); $("mediaUrlInput").focus(); } }, uiIcon("link"), "URL・YouTube")),
  photoInput, videoInput, lottieInput);
  const selectedKey = slide.visualAsset || (slide.media?.src?.startsWith("asset:") ? slide.media.src.slice(6) : "");
  nodes.push(h("details", { class: "illust-picker", "data-key": "library", open: !info || Boolean(selectedKey) },
    h("summary", { class: "hint", style: { cursor: "pointer", "font-weight": "700", "margin-bottom": "6px" } }, `写真ライブラリ（${Object.keys(E.PHOTOS).length}枚）`),
    h("div", { class: "image-picker" },
      h("button", { type: "button", class: !info ? "selected" : "", title: "写真を置かない", onclick: () => chooseLibraryPhoto(slide, "") }, h("span", {}, "なし")),
      Object.entries(E.PHOTOS).map(([key, [file, label]]) => h("button", { type: "button", class: selectedKey === key ? "selected" : "", title: label, onclick: () => chooseLibraryPhoto(slide, key) },
        h("img", { src: `/assets/${file}`, alt: label, loading: "lazy" }), h("span", { class: "cap" }, label))))));
  const where = SLOTTED.has(slide.type) ? "このレイアウトは写真枠があります。" : "このレイアウトでは、写真・動画をスライド上の好きな位置に置けます。";
  return h("div", { class: "section", "data-key": "media" },
    h("div", { class: "section-title" }, h("span", {}, "写真・動画・アニメーション")),
    h("p", { class: "hint section-note" }, `${where}スライドへ直接ドラッグ＆ドロップもできます。`),
    nodes);
}

// ---- motion (builds) and "click to open" details

function motionSection(slide, index) {
  const current = slide.animation && slide.animation !== "auto" ? slide.animation : "auto";
  const recommended = E.recommendedBuild(slide.type);
  const effective = current === "auto" ? recommended : current;
  const set = (value) => { pushUndo(); if (value === "auto") delete slide.animation; else slide.animation = value; markChanged({ structural: true }); };
  const glyph = (kind) => h("span", { class: "glyph", "data-kind": kind === "auto" ? recommended : kind }, h("i"), h("i"), h("i"));
  return h("div", { class: "section" },
    h("div", { class: "section-title" }, h("span", {}, "動き（中身の出し方）"),
      h("span", { class: "btns" },
        h("button", { class: "btn btn-sm", type: "button", onclick: () => previewMotion() }, "▶ 確認"),
        h("button", { class: "btn btn-sm", type: "button", title: "この設定を全スライドに適用", onclick: () => {
          pushUndo();
          for (const other of state.deck.slides) { if (current === "auto") delete other.animation; else if (TITLED(other.type)) other.animation = current; }
          markChanged({ structural: true });
          toast(`全スライドを「${BUILD_INFO[current][0]}」にしました`);
        } }, "全体に適用"))),
    h("div", { class: "motion-grid", role: "radiogroup" }, Object.entries(BUILD_INFO).map(([value, [label]]) => h("button", { type: "button", role: "radio", "aria-checked": String(value === current), class: value === current ? "selected" : "", title: BUILD_INFO[value][1], onclick: () => set(value) }, glyph(value), label))),
    h("div", { class: "hint", style: { "margin-top": "6px" } }, current === "auto" ? `おまかせ：このレイアウトは「${BUILD_INFO[recommended][0]}」。${BUILD_INFO[recommended][1]}` : BUILD_INFO[effective][1]),
    h("button", { class: "btn btn-ghost btn-sm", type: "button", style: { "margin-top": "6px", padding: "0" }, onclick: () => openDesignDialog() }, "切り替え・登場のしかた・マウスを乗せたときの動き（資料全体）→"));
}

// ---- the motion IDE: what moves on its own on the slide on the stage, one row each in the animation pane

/**
 * The slide's automatic motion as editable rows ({ key, label, options: [[value, label]], value, now }): the
 * transition into it, how it comes in, how its content builds, its title / big words, its background, its
 * emphasis. "auto" follows the deck (the design tab) or the layout. A page brought over as it looks keeps its own
 * look: only its transition moves on its own (its objects move by their animations and interactions).
 */
function autoMotionInfo() {
  const slide = state.deck?.slides[state.selected];
  if (!slide) return null;
  const deckMotion = normalizeMotion(state.deck.motion || DEFAULT_MOTION);
  const source = slide.master === "source";
  const titled = TITLED(slide.type);
  const deckWord = (label) => `資料の設定（${String(label || "なし").replace(/（.*）/, "")}）`;
  const opts = (info, autoLabel) => Object.entries(info).map(([value, label]) => [value, value === "auto" ? autoLabel : String(Array.isArray(label) ? label[0] : label).replace(/（.*）/, "")]);
  const rows = [
    { key: "transition", label: "画面切り替え", options: opts(SLIDE_TRANSITION_INFO, deckWord(E.TRANSITIONS[state.deck.transition || "fade"])), value: slide.transition || "auto" },
  ];
  if (!source) {
    const kinetic = E.kineticOf(slide, slide.type, deckMotion);
    const backdrop = E.backdropOf(slide, slide.type, deckMotion);
    const recommended = E.recommendedBuild(slide.type);
    // A blank slide has no layout content to build or emphasise (its objects move by their animations).
    const blank = slide.type === "blank";
    const title = !(blank && slide.hideTitle);
    if (title) rows.push({ key: "entrance", label: "登場のしかた", options: opts(ENTRANCE_INFO, deckWord(deckMotion.entrance === "none" ? "動かさない" : E.ENTRANCES[deckMotion.entrance])), value: slide.entrance || "auto" });
    if (titled && !blank) rows.push({ key: "animation", label: "中身の出し方", options: opts(BUILD_INFO, `おまかせ（${BUILD_INFO[recommended][0]}）`), value: slide.animation || "auto" });
    if (title) rows.push({ key: "kinetic", label: titled ? "タイトルの動き" : "大きな文字の動き", options: opts(KINETIC_INFO, `おまかせ（${kinetic ? E.KINETIC[kinetic].replace(/（.*）/, "") : "動かさない"}）`), value: slide.kinetic || "auto" });
    rows.push({ key: "backdrop", label: "背景の動き", options: opts(BACKDROP_INFO, `おまかせ（${backdrop ? E.BACKDROPS[backdrop] : "なし"}）`), value: slide.backdrop || "auto" });
    if (!blank) rows.push({ key: "emphasis", label: "強調（**語句**）", options: opts(EMPHASIS_INFO, deckWord(E.EMPHASES[deckMotion.emphasis])), value: slide.emphasis || "auto" });
  }
  // Sounds and videos: how each starts (再生 →「開始」), so what plays by itself is in the same list.
  const media = (slide.elements || []).filter((o) => (o.kind === "audio" || (o.kind === "video" && !E.youtubeId(o.src))) && !o.hidden);
  for (const o of media) {
    const inSequence = (slide.timeline || []).some((e) => e.cls === "media" && e.fx === "play" && e.el === o.id && !e.trigger);
    const name = o.name || o.fileName || E.KIND_LABELS[o.kind];
    rows.push({ key: `media:${o.id}`, label: `${o.kind === "audio" ? "🔊" : "▶"} ${String(name).slice(0, 18)}${o.across ? "（切り替え後も）" : ""}`,
      options: [["auto", "自動で再生"], ["sequence", "一連のクリック動作"], ["click", "クリック時（アイコン）"]], value: o.autoplay ? "auto" : inSequence ? "sequence" : "click" });
  }
  return { source, rows };
}
function setAutoMotion(key, value) {
  if (key.startsWith("media:")) return setMediaStart(key.slice(6), value);
  if (!["transition", "entrance", "animation", "kinetic", "backdrop", "emphasis"].includes(key)) return;
  setSlideFields({ [key]: value === "auto" ? null : value });
}
/** How a sound or a video starts, in one undo step: with the slide, in the click order, or on a click on it. */
function setMediaStart(id, value) {
  const slide = state.deck?.slides[state.selected];
  if (!slide?.elements?.some((o) => o.id === id)) return;
  const elements = slide.elements.map((o) => (o.id === id ? { ...o, autoplay: value === "auto" } : o));
  let timeline = (slide.timeline || []).filter((e) => !(e.cls === "media" && e.fx === "play" && e.el === id && !e.trigger));
  if (value === "sequence") timeline = [...timeline, { id: `a${Math.random().toString(36).slice(2, 9)}`, el: id, cls: "media", fx: "play", start: "click", dur: 1, delay: 0 }];
  setObjects(elements, { timeline });
}

// ---- motion graphics for one slide (the deck's defaults live in the design dialog)

function motionGraphicsSection(slide) {
  const deckMotion = normalizeMotion(state.deck.motion || DEFAULT_MOTION);
  const kinetic = E.kineticOf(slide, slide.type, deckMotion);
  const backdrop = E.backdropOf(slide, slide.type, deckMotion);
  const set = (key, value) => { pushUndo(); if (value === "auto") delete slide[key]; else slide[key] = value; markChanged({ structural: true }); };
  const select = (key, info, label) => h("label", { class: "field", style: { "margin-bottom": "0" } }, h("span", { class: "field-label" }, label),
    h("select", { "data-mg": key, onchange: (event) => set(key, event.target.value) }, Object.entries(info).map(([value, text]) => h("option", { value, selected: (slide[key] || "auto") === value }, text))));
  const now = `いま：文字は「${kinetic ? E.KINETIC[kinetic].replace(/（.*）/, "") : "動かさない"}」、背景は「${backdrop ? E.BACKDROPS[backdrop] : "なし"}」`;
  return h("div", { class: "section", "data-key": "motion-graphics" },
    h("div", { class: "section-title" }, h("span", {}, "モーショングラフィック・この1枚の動き"),
      h("span", { class: "btns" }, h("button", { class: "btn btn-sm", type: "button", onclick: () => previewMotion() }, "▶ 確認"))),
    h("div", { class: "grid-2 mg-grid" },
      select("kinetic", KINETIC_INFO, TITLED(slide.type) && !["hero", "statement"].includes(slide.type) ? "タイトルの動き" : "大きな文字の動き"),
      select("backdrop", BACKDROP_INFO, "背景の動き"),
      select("entrance", ENTRANCE_INFO, "登場のしかた"),
      select("emphasis", EMPHASIS_INFO, "強調（**語句**）の見せ方"),
      select("transition", SLIDE_TRANSITION_INFO, "このスライドへの切り替え")),
    h("div", { class: "hint", style: { "margin-top": "6px" } }, `${now}。おまかせは「デザインと動き」の設定に従います（文字・背景は表紙・章扉・ひと言・最後のスライドに効きます）。`));
}

const ITEM_NAMES = { items: "項目", steps: "工程", points: "要点", rows: "行", milestones: "時期", lanes: "レーン", levels: "段", branches: "枝", stats: "指標", flows: "流れ", message: "アクション", leftItems: "左", rightItems: "右", measures: "打ち手", inputs: "条件" };

/** The words of one item ("items[1]") of a slide. */
function itemText(slide, key) {
  const path = parseField(key);
  let value = getPath(slide, path);
  if (path[0] === "message") value = String(slide.message || "").split(/\n|／|(?<=。)(?=.)/).map((part) => part.trim()).filter(Boolean)[path[1]];
  if (Array.isArray(value)) value = value.join(" / ");
  if (value && typeof value === "object") value = value.title ?? value.label ?? value.q ?? value.value ?? (Array.isArray(value.steps) ? value.steps.join("→") : "");
  return strip(value);
}

function itemLabel(slide, key) {
  const path = parseField(key);
  const name = ITEM_NAMES[path[0]] ?? path[0];
  const n = typeof path[1] === "number" ? path[1] + 1 : "";
  return `${name}${n}「${itemText(slide, key).slice(0, 22)}」`;
}

function itemKeys(slide, index) {
  const el = E.render(slide, renderOptions({ index, mode: "thumb" }));
  return [...new Set([...el.querySelectorAll("[data-item]")].map((node) => node.dataset.item))].map((key) => ({ key, label: itemLabel(slide, key) }));
}

/** "ラベル,値" lines ⇄ a detail's breakdown rows. */
const rowsToText = (rows) => (rows || []).map((row) => `${row.label},${row.value ?? ""}`).join("\n");
const textToRows = (text) => String(text).split("\n").map((line) => line.split(/[,\t，]/)).filter((cells) => strip(cells[0])).map(([label, ...rest]) => ({ label: label.trim().slice(0, 40), value: rest.join(",").trim().slice(0, 30) })).slice(0, 8);

function detailsSection(slide, index) {
  // The key message can carry the evidence for the slide's conclusion ("根拠"), besides any item.
  const keys = [...(TITLED(slide.type) && strip(slide.takeaway) ? [{ key: "takeaway", label: "キーメッセージの根拠" }] : []), ...itemKeys(slide, index)];
  const details = Array.isArray(slide.details) ? slide.details : [];
  const commit = (fn, structural = false) => { if (structural) pushUndo(); else beginEdit(); fn(); if (!slide.details?.length) delete slide.details; markChanged({ structural }); };
  const setOptional = (detail, key, value) => { if (strip(value)) detail[key] = value; else delete detail[key]; };
  const cards = details.map((detail, i) => h("div", { class: "detail-card" },
    h("div", { class: "detail-head" },
      h("select", { "aria-label": "詳細を開く項目", onchange: (event) => commit(() => { detail.target = event.target.value; }, true) },
        keys.some((entry) => entry.key === detail.target) ? null : h("option", { value: detail.target, selected: true }, `（見つからない項目：${detail.target}）`),
        keys.map((entry) => h("option", { value: entry.key, selected: entry.key === detail.target }, entry.label))),
      h("button", { class: "btn btn-ghost btn-sm btn-danger", type: "button", title: "この詳細を削除", onclick: () => commit(() => { details.splice(i, 1); }, true) }, "✕")),
    h("input", { type: "text", maxlength: 60, placeholder: "詳細の見出し（任意）", value: detail.title ?? "", oninput: (event) => commit(() => { if (event.target.value) detail.title = event.target.value; else delete detail.title; }) }),
    h("textarea", { rows: 3, maxlength: 400, placeholder: "クリックしたときに開く説明（根拠・具体例・数字の出典など）", oninput: (event) => commit(() => { detail.text = event.target.value; }) }, detail.text ?? ""),
    h("details", { class: "detail-evidence", "data-key": `evidence-${i}`, open: Boolean(detail.rows?.length || detail.source || detail.note) },
      h("summary", {}, "根拠パネルにする（内訳・出所・前提）"),
      h("textarea", { rows: 3, placeholder: "内訳（1行に「ラベル,値」。例：資料作成,620時間）", "aria-label": "内訳", oninput: (event) => commit(() => { const rows = textToRows(event.target.value); if (rows.length) detail.rows = rows; else delete detail.rows; }) }, rowsToText(detail.rows)),
      h("input", { type: "text", maxlength: 120, placeholder: "出所（例：社内アンケート 2025年9月、n=120）", "aria-label": "出所", value: detail.source ?? "", oninput: (event) => commit(() => setOptional(detail, "source", event.target.value)) }),
      h("textarea", { rows: 2, maxlength: 200, placeholder: "前提・注記（任意）", "aria-label": "前提", oninput: (event) => commit(() => setOptional(detail, "note", event.target.value)) }, detail.note ?? ""),
      h("div", { class: "hint" }, "どれかを入れると、発表中は右から根拠パネルが開きます（内訳は数字なら棒で表示）。"))));
  const free = keys.find((entry) => !details.some((detail) => detail.target === entry.key));
  return h("div", { class: "section" },
    h("div", { class: "section-title" }, h("span", {}, `クリックで開く詳細（${details.length}）`),
      h("span", { class: "btns" }, h("button", { class: "btn btn-sm", type: "button", disabled: !state.codexAuthorized || state.aiBusy || !keys.length, title: "各項目の補足をAIに書いてもらう", onclick: () => sendChat(`@${index + 1} の主な項目に、クリックで開く詳細（根拠・具体例・数字の内訳）を付けて`) }, "✦ AIで作る"))),
    keys.length
      ? h("p", { class: "hint section-note" }, "発表中に項目をクリックすると、説明のカードが開きます。スライドに書ききれない根拠や具体例に。内訳・出所を入れると、右から開く根拠パネルになります。")
      : h("p", { class: "hint section-note" }, "このレイアウトには、詳細を付けられる項目がありません。"),
    cards,
    keys.length && details.length < 12 ? h("button", { class: "btn add-item", type: "button", onclick: () => commit(() => { slide.details = [...details, { target: (free ?? keys[0]).key, text: "" }]; }, true) }, "＋ 詳細を追加") : null);
}

/** A blank deep-dive page for one item of a slide, right after the slide's other deep-dive pages. */
function addDrill(index, target) {
  const slide = state.deck.slides[index];
  const text = itemText(slide, target) || "深掘り";
  const at = insertSlide(groupEnd(index), { type: "content", title: text.slice(0, 40), takeaway: "", points: ["背景：なぜそうなるのか", "具体例：現場ではどうなっているか", "根拠：数字や出典"], drillOf: target });
  if (at >= 0) toast(`「${text.slice(0, 20)}」の深掘りページを作りました。発表中にその項目をクリックすると開きます`);
}

function addDrillWithAi(index, target) {
  const slide = state.deck.slides[index];
  const status = (text) => { const el = document.getElementById("aiStatus"); if (el) el.textContent = text; };
  const instruction = `${index + 1}枚目「${strip(slide.title)}」の${itemLabel(slide, target)}を深掘りするページを作る。本編の流れには入らず、発表中にこの項目をクリックしたときだけ開く補足として、この項目の背景・内訳・具体例・根拠を1枚で見せる。元のスライドに書いてあることを繰り返さない。type は title・section・closing 以外から内容に合うものを選ぶ`;
  return aiInsertSlide(groupEnd(index), instruction, status, { drillOf: target });
}

/** Make a deep-dive page part of the story: it moves after its slide's other deep-dive pages. */
function makeStorySlide(index) {
  const end = groupEnd(index);
  pushUndo();
  const [page] = state.deck.slides.splice(index, 1);
  delete page.drillOf;
  state.deck.slides.splice(end - 1, 0, page);
  state.selected = end - 1;
  markChanged({ structural: true });
  toast("本編のスライドにしました（⌘Zで元に戻せます）");
}

function drillSection(slide, index) {
  const story = storyOf();
  const keys = itemKeys(slide, index);
  const drills = story.drills[index] ?? [];
  const taken = new Set(drills.map((drill) => drill.target));
  const free = keys.filter((entry) => !taken.has(entry.key));
  const pick = h("select", { "aria-label": "深掘りページを開く項目" }, free.map((entry) => h("option", { value: entry.key }, entry.label)));
  const rows = drills.map((drill) => {
    const page = state.deck.slides[drill.index];
    return h("div", { class: "drill-row" },
      h("select", { "aria-label": "このページを開く項目", onchange: (event) => { pushUndo(); page.drillOf = event.target.value; markChanged({ structural: true }); } },
        keys.some((entry) => entry.key === drill.target) ? null : h("option", { value: drill.target, selected: true }, `（見つからない項目：${drill.target}）`),
        keys.map((entry) => h("option", { value: entry.key, selected: entry.key === drill.target, disabled: entry.key !== drill.target && taken.has(entry.key) }, entry.label))),
      h("button", { class: "btn btn-sm drill-open", type: "button", title: strip(page.title) || typeLabel(page.type), onclick: () => select(drill.index) }, `${drill.index + 1}枚目を開く`),
      h("button", { class: "btn btn-ghost btn-sm btn-danger", type: "button", title: "この深掘りページを削除", onclick: () => deleteSlide(drill.index) }, "✕"));
  });
  return h("div", { class: "section", id: "drillSection" },
    h("div", { class: "section-title" }, h("span", {}, `クリックで移る深掘りページ（${drills.length}）`)),
    h("p", { class: "hint section-note" }, keys.length
      ? "発表中に項目をクリックすると、本編の流れに入っていない深掘りページへ移ります。Esc・←、または最後まで進めると元のスライドに戻ります。"
      : "このレイアウトには、深掘りページを付けられる項目がありません。"),
    rows,
    free.length ? h("div", { class: "drill-add" }, pick,
      h("span", { class: "btns" },
        h("button", { class: "btn btn-sm", type: "button", onclick: () => addDrill(index, pick.value) }, "＋ 白紙で作る"),
        h("button", { class: "btn btn-sm", type: "button", disabled: !state.codexAuthorized || state.aiBusy, title: state.codexAuthorized ? "その項目を深掘りするページをAIに作ってもらう" : "Codexに接続すると使えます", onclick: () => addDrillWithAi(index, pick.value) }, "✦ AIで作る"))) : null);
}

function drillInfoSection(slide, index, parent) {
  const story = storyOf();
  const owner = state.deck.slides[parent];
  const keys = itemKeys(owner, parent);
  const taken = new Set((story.drills[parent] ?? []).filter((drill) => drill.index !== index).map((drill) => drill.target));
  return h("div", { class: "section drill-info" },
    h("div", { class: "section-title" }, h("span", {}, "深掘りページ")),
    h("p", { class: "hint section-note" }, `${parent + 1}枚目「${strip(owner.title || owner.message || "")}」の項目をクリックすると開くページです。本編の番号には数えず、発表中は Esc・← で元のスライドに戻ります。`),
    h("div", { class: "field" }, h("div", { class: "label-row" }, h("label", { for: "drillTarget" }, "開く項目")),
      h("select", { id: "drillTarget", onchange: (event) => { pushUndo(); slide.drillOf = event.target.value; markChanged({ structural: true }); } },
        keys.some((entry) => entry.key === slide.drillOf) ? null : h("option", { value: slide.drillOf, selected: true }, `（見つからない項目：${slide.drillOf}）`),
        keys.map((entry) => h("option", { value: entry.key, selected: entry.key === slide.drillOf, disabled: taken.has(entry.key) }, entry.label)))),
    h("div", { class: "btns", style: { display: "flex", gap: "6px", "flex-wrap": "wrap" } },
      h("button", { class: "btn btn-sm", type: "button", onclick: () => select(parent) }, "← 元のスライドへ"),
      h("button", { class: "btn btn-sm", type: "button", title: "深掘りページをやめて、本編の流れに入れます", onclick: () => makeStorySlide(index) }, "本編のスライドにする")));
}

function renderInspector() {
  const panel = $("inspector");
  if (!state.deck || state.mode !== "edit") { panel.replaceChildren(); return; }
  const index = state.selected;
  const slide = state.deck.slides[index];
  const total = state.deck.slides.length;
  const locked = index === 0 || index === total - 1;
  // Re-rendering the same slide (after a picker click, reorder…) keeps the scroll position and open pickers.
  const sameSlide = panel.dataset.slide === String(index);
  const keepScroll = sameSlide ? panel.querySelector(".inspector-body")?.scrollTop ?? 0 : 0;
  const keepOpen = sameSlide ? [...panel.querySelectorAll("details[data-key][open]")].map((el) => el.dataset.key) : [];
  const keepClosed = sameSlide ? [...panel.querySelectorAll("details[data-key]:not([open])")].map((el) => el.dataset.key) : [];
  panel.dataset.slide = String(index);
  E.stopLottie(panel);
  const canAi = state.codexAuthorized;
  const chips = ["もっと簡潔に", "結論を先に", "具体例を加えて", "数字を強調", "図解に変えて", "やさしい言葉で"];
  panel.replaceChildren(
    h("div", { class: "inspector-head" },
      h("b", {}, `${index + 1} / ${total}`),
      h("button", { class: "btn", type: "button", style: { flex: "1", "justify-content": "space-between" }, disabled: locked, title: locked ? "表紙と最後のスライドはレイアウトを変えられません" : "レイアウトを変更", onclick: () => openTypeDialog("change") }, typeLabel(slide.type), h("span", { style: { color: "var(--muted)" } }, locked ? "固定" : "変更 ▾"))),
    h("div", { class: "inspector-body" },
      issueSection(),
      h("div", { class: "inspector-ai" },
        h("b", {}, "✦ AIに頼む"),
        chips.map((chip) => h("button", { class: "chip", type: "button", disabled: !canAi, title: "AIとの会話で、この1枚への変更を提案してもらいます", onclick: () => sendChat(`この1枚を${chip}`) }, chip)),
        h("button", { class: "chip ai", type: "button", disabled: !canAi, title: "見せ方の違う3つの案を並べて選べます", onclick: () => requestVariants(index) }, "✦ 別案を3つ"),
        h("span", { id: "aiStatus", class: "hint", style: { "flex-basis": "100%" } }, canAi ? "変更は「AIと話す」に提案として届き、確認してから採用できます。" : "Codexに接続すると使えます。")),
      storyOf().parent[index] != null ? drillInfoSection(slide, index, storyOf().parent[index]) : null,
      h("div", { class: "section" }, h("div", { class: "section-title" }, h("span", {}, "内容")),
        slide.type === "hero" ? h("p", { class: "hint section-note" }, "写真いっぱいに見出しを重ねるスライドです。写真・動画は下で選びます。") : null,
        specFor(slide.type).map((field) => inputFor(field, [field.key], slide))),
      mediaSection(slide),
      TITLED(slide.type) ? motionSection(slide, index) : null,
      motionGraphicsSection(slide),
      TITLED(slide.type) || slide.type === "closing" ? detailsSection(slide, index) : null,
      TITLED(slide.type) && index > 0 && storyOf().parent[index] == null ? drillSection(slide, index) : null,
      h("div", { class: "section" }, h("div", { class: "section-title" }, h("span", {}, "スピーカーノート"),
          h("span", { class: "btns" },
            h("button", { class: "btn btn-sm", type: "button", title: "スライドの内容から、すぐに読み上げ原稿を作ります", onclick: () => { applyNotes([{ slide: index, text: quickNotes(index) }]); toast("ノートを作成しました（⌘Zで元に戻せます）"); } }, "簡易作成"),
            h("button", { class: "btn btn-sm", type: "button", disabled: !canAi || state.aiBusy, title: canAi ? "AIで自然な読み上げ原稿を作ります" : "Codexに接続すると使えます", onclick: () => aiNotes([index], {}, (text) => { const el = document.getElementById("aiStatus"); if (el) el.textContent = text; }) }, "✦ AIで作成"))),
        h("textarea", { rows: 4, placeholder: "発表時に話す内容（発表者ビュー・ノート欄に表示されます）", "data-path": "notes", oninput: (event) => { beginEdit(); setPath(slide, ["notes"], event.target.value); delete slide.notesRich; markChanged(); } }, slide.notes ?? ""))),
    h("div", { class: "inspector-foot" },
      h("button", { class: "btn", type: "button", onclick: () => openTypeDialog("insert") }, "＋ 後ろに追加"),
      h("button", { class: "btn", type: "button", disabled: locked, onclick: () => insertSlide(index + 1, copyOf(slide)) }, "複製"),
      h("span", { style: { flex: "1" } }),
      h("button", { class: "btn btn-danger", type: "button", disabled: locked, onclick: () => deleteSlide(index) }, "削除")));
  for (const key of keepOpen) panel.querySelector(`details[data-key="${CSS.escape(key)}"]`)?.setAttribute("open", "");
  for (const key of keepClosed) panel.querySelector(`details[data-key="${CSS.escape(key)}"]`)?.removeAttribute("open");
  if (keepScroll) panel.querySelector(".inspector-body").scrollTop = keepScroll;
  refreshInspectorIssues();
}

// ---------------------------------------------------------------- layout picker

function openTypeDialog(mode) {
  const dialog = $("typeDialog");
  const current = state.deck.slides[state.selected];
  $("typeDialogTitle").textContent = mode === "change" ? "レイアウトを変更（内容はできるだけ引き継ぎます）" : "スライドを追加";
  const types = SLIDE_TYPES.filter((type) => !["title", "closing"].includes(type));
  document.getElementById("aiInsertBox")?.remove();
  if (mode === "insert" && state.codexAuthorized) {
    const input = h("textarea", { rows: 2, placeholder: "例：部門別の効果を比べる表／導入スケジュール／想定される質問と回答" });
    const status = h("span", { class: "hint" }, "前後の流れを踏まえて、AIが内容とレイアウトを決めます。");
    const button = h("button", { class: "btn btn-ai", type: "button", disabled: state.aiBusy, onclick: async () => {
      if (!input.value.trim()) return input.focus();
      button.disabled = true;
      status.textContent = "作成中…";
      const ok = await aiInsertSlide(Math.min(state.selected + 1, state.deck.slides.length - 1), input.value.trim(), (text) => { status.textContent = text; });
      if (ok) dialog.close();
      else button.disabled = false;
    } }, "✦ AIで作成");
    $("typeGrid").before(h("div", { id: "aiInsertBox", class: "section ai-box" },
      h("div", { class: "section-title" }, h("span", {}, "✦ AIに作ってもらう")), input,
      h("div", { class: "ai-row", style: { "margin-top": "8px" } }, status, button),
      h("div", { class: "section-title", style: { margin: "14px 0 0" } }, h("span", {}, "または、レイアウトを選んで自分で作る"))));
  }
  const previewDeck = { ...state.deck, slides: [...state.deck.slides] };
  $("typeGrid").replaceChildren(...types.map((type) => {
    const sample = mode === "change" ? convertSlide(current, type) : defaultSlide(type);
    let picture = null;
    try { picture = slidePicture(sample, Math.max(1, state.selected), previewDeck); } catch { picture = null; }
    return h("button", {
      type: "button",
      class: mode === "change" && current.type === type ? "current" : "",
      onclick: () => {
        dialog.close();
        if (mode === "change") {
          if (current.type === type) return;
          replaceSlide(state.selected, convertSlide(current, type));
          toast(`${typeLabel(type)}に変更しました`);
        } else {
          insertSlide(Math.min(state.selected + 1, state.deck.slides.length - 1), defaultSlide(type));
        }
      },
    }, picture, h("b", {}, TYPE_INFO[type][0]), h("span", {}, TYPE_INFO[type][1]));
  }));
  dialog.showModal();
}

// ---------------------------------------------------------------- design & motion (deck-wide)

function syncDesignControls() {
  const deck = state.deck;
  const motion = normalizeMotion(deck.motion || DEFAULT_MOTION);
  $("transitionSelect").value = deck.transition || "fade";
  $("entranceSelect").value = motion.entrance || "rise";
  $("hoverSelect").value = motion.hover || "lift";
  $("numbersCheck").checked = motion.numbers !== false;
  $("ambientCheck").checked = motion.ambient !== false;
  $("drawCheck").checked = motion.draw !== false;
  for (const [grid, value] of [["backdropGrid", motion.backdrop], ["kineticGrid", motion.kinetic], ["emphasisGrid", motion.emphasis]]) {
    for (const card of document.querySelectorAll(`#${grid} .motion-card`)) card.setAttribute("aria-checked", String(card.dataset.value === value));
  }
}

/** The design dialog's selects list what the engine knows. */
function fillMotionSelects() {
  const fill = (id, entries) => $(id).replaceChildren(...entries.map(([value, text]) => h("option", { value }, text)));
  fill("transitionSelect", Object.entries(E.TRANSITIONS));
  fill("entranceSelect", [...Object.entries(E.ENTRANCES), ["none", "動かさない"]]);
  fill("hoverSelect", [...Object.entries(E.HOVERS), ["none", "変化なし"]]);
}

// Live previews of every motion graphic, drawn on the deck's own cover (and replayed while the dialog is open).
let motionCardsTimer = null;
function renderMotionGrids() {
  const deck = state.deck;
  const cover = deck.slides[0];
  const motion = normalizeMotion(deck.motion || DEFAULT_MOTION);
  // Emphasis is shown on a one-line statement in the deck's own design.
  const phrase = { type: "statement", title: "強調の見せ方", text: "伝えたいことは**ひと言**で" };
  const phraseDeck = { ...deck, slides: [cover, phrase, deck.slides.at(-1)] };
  const card = (current, value, label, patch, sample = null) => {
    const el = sample
      ? E.render(sample, { ...renderOptions(), deck: { ...phraseDeck, motion: { ...motion, backdrop: "none", ...patch } }, index: 1, mode: "preview" })
      : E.render(cover, { ...renderOptions(), deck: { ...deck, motion: { ...motion, ...patch } }, index: 0, mode: "preview", fit: fitFor(0) ?? undefined });
    return h("button", { type: "button", class: "motion-card", role: "radio", "data-value": value, "aria-checked": String(current === value), onclick: () => setDeckDesign({ motion: patch }) }, E.mount(el), h("b", {}, label));
  };
  $("backdropGrid").replaceChildren(...["none", ...Object.keys(E.BACKDROPS)].map((kind) => card(motion.backdrop, kind, kind === "none" ? "テンプレートの波紋だけ" : E.BACKDROPS[kind], { backdrop: kind })));
  $("kineticGrid").replaceChildren(...["none", ...Object.keys(E.KINETIC)].map((style) => card(motion.kinetic, style, style === "none" ? "動かさない" : E.KINETIC[style].replace(/（.*）/, ""), { kinetic: style })));
  $("emphasisGrid").replaceChildren(...Object.entries(E.EMPHASES).map(([style, label]) => card(motion.emphasis, style, label, { emphasis: style }, phrase)));
  const replay = () => { for (const slide of document.querySelectorAll("#kineticGrid .motion-card .hs-slide, #emphasisGrid .motion-card .hs-slide")) E.play(slide); };
  clearInterval(motionCardsTimer);
  requestAnimationFrame(replay);
  motionCardsTimer = setInterval(() => { if ($("designDialog").open) replay(); else clearInterval(motionCardsTimer); }, 3600);
}

function openDesignDialog() {
  if (!state.deck) return;
  useFonts([state.deck.theme]);
  renderMotionGrids();
  syncDesignControls();
  $("designDialog").showModal();
}

/** Fields of several objects on any slides ([{ slide, id, patch }], undefined removes a field): one undo step. */
function patchObjects(changes) {
  if (!state.deck || !changes?.length) return;
  pushUndo();
  for (const { slide: i, id, patch } of changes) {
    const slide = state.deck.slides[i];
    const at = slide?.elements?.findIndex((o) => o.id === id) ?? -1;
    if (at < 0) continue;
    const next = { ...slide.elements[at] };
    for (const [key, value] of Object.entries(patch)) { if (value === undefined) delete next[key]; else next[key] = value; }
    slide.elements[at] = E.normalizeObject(next) || slide.elements[at];
  }
  markChanged({ structural: true });
}

/** One object's fields on any slide (undefined removes a field): one undo step. */
function setObjectFields(index, id, patch) {
  const slide = state.deck?.slides[index];
  const at = slide?.elements?.findIndex((o) => o.id === id) ?? -1;
  if (at < 0) return;
  pushUndo();
  const next = { ...slide.elements[at] };
  for (const [key, value] of Object.entries(patch)) { if (value === undefined) delete next[key]; else next[key] = value; }
  slide.elements[at] = E.normalizeObject(next) || slide.elements[at];
  markChanged({ structural: true });
}

/** Deck settings that are not slides (the show's settings, custom shows): one undo step; null removes one. */
function setDeckFields(patch) {
  if (!state.deck) return;
  pushUndo();
  for (const [key, value] of Object.entries(patch)) {
    if (value == null) delete state.deck[key]; else state.deck[key] = value;
  }
  markChanged({ structural: true });
}

function setDeckDesign(patch) {
  if (!state.deck) return;
  pushUndo();
  const deck = state.deck;
  // テーマ: the deck's colours and fonts change; what was painted in the old theme's colours (テーマの色) takes the new
  // theme's, as PowerPoint recolours theme colours.
  if (patch.theme && THEME_IDS.has(patch.theme) && patch.theme !== deck.theme) {
    E.retheme(deck, deck.theme, patch.theme);
    deck.theme = patch.theme;
    E.useTheme(patch.theme);
    rememberTheme(patch.theme);
    // Another theme sets other fonts and sizes: every slide is measured again.
    fitState.epoch += 1;
    fitState.byKey.clear();
    thumbCache.clear();
  }
  if (patch.transition) deck.transition = patch.transition;
  if (patch.motion) deck.motion = normalizeMotion({ ...(deck.motion || DEFAULT_MOTION), ...patch.motion });
  useFonts([deck.theme]);
  markChanged({ structural: true });
  if (patch.theme) { scheduleMeasure(0); updateDesignButton(); editorUi.renderRibbon(); }
  if ($("designDialog").open) {
    // Motion previews follow each other (the kinetic cards show the chosen backdrop).
    if (patch.motion) renderMotionGrids();
    syncDesignControls();
  }
}


// ---------------------------------------------------------------- AI jobs

function showImageCoverage() {
  if (autoImageRun || !state.deck) return;
  const plan = picturePlan(state.deck);
  $("autoImageStatus").textContent = state.codexImageAuthorized ? `画像 ${plan.present}/${state.deck.slides.length}枚` : "";
  $("autoImageAction").textContent = "画像を追加";
  $("autoImageAction").classList.toggle("hidden", !state.codexImageAuthorized || !plan.needed);
  $("autoImageAction").disabled = false;
}

function showAutoImageProgress(run) {
  if (autoImageRun !== run) return;
  const present = picturePlan(run.deck).present;
  const active = !run.cancelled && (run.inFlight > 0 || (present < run.target && run.next < run.candidates.length));
  $("autoImageStatus").textContent = run.cancelled && run.inFlight ? "画像生成を停止中"
    : active
    ? `画像を生成中 ${present}/${run.target}枚${run.failures ? `（失敗 ${run.failures}件）` : ""}`
    : present >= run.target ? `画像 ${present}/${run.deck.slides.length}枚` : `画像 ${present}/${run.target}枚・未完了`;
  $("autoImageAction").textContent = active ? "停止" : "再試行";
  $("autoImageAction").classList.toggle("hidden", present >= run.target || (run.cancelled && run.inFlight > 0));
}

async function requestAutoImage(deck, index) {
  const { jobId } = await jsonFetch("/api/decks/image", {
    method: "POST",
    body: JSON.stringify({ deck: serverDeck(deck), index, message: "この1枚の主張と具体的な本文に直接合う画像を新規生成してください。文字・ロゴ・数値・グラフは描かず、内容にない人物や成果も加えないでください。" }),
  });
  const job = await new Promise((resolve, reject) => watchJob(jobId, {
    onDone: resolve,
    onFail: (failed) => reject(new Error(failed.error || failed.detail || "画像生成に失敗しました")),
  }));
  if (!job.image?.url) throw new Error("生成画像が返されませんでした");
  const response = await fetch(job.image.url, { credentials: "same-origin" });
  if (!response.ok) throw new Error(`画像の取得に失敗しました（${response.status}）`);
  return putMedia(await response.blob(), `自動生成_${index + 1}枚目`);
}

async function autoIllustrateDeck(deck) {
  const plan = picturePlan(deck);
  if (!plan.needed || state.deck !== deck) return;
  if (!state.codexImageAuthorized) {
    $("autoImageStatus").textContent = "画像の追加生成にはCodex接続が必要です";
    return;
  }
  if (autoImageRun) autoImageRun.cancelled = true;
  const run = { deck, target: plan.target, candidates: plan.candidates, next: 0, inFlight: 0, failures: 0, cancelled: false };
  autoImageRun = run;
  showAutoImageProgress(run);
  const worker = async () => {
    while (!run.cancelled && state.deck === deck) {
      if (picturePlan(deck).present + run.inFlight >= run.target) break;
      const candidate = run.candidates[run.next++];
      if (!candidate) break;
      const slide = deck.slides[candidate.index];
      if (hasSlidePicture(slide)) continue;
      const before = JSON.stringify(slide);
      run.inFlight += 1;
      showAutoImageProgress(run);
      try {
        const src = await requestAutoImage(deck, candidate.index);
        if (run.cancelled || state.deck !== deck || deck.slides[candidate.index] !== slide || JSON.stringify(slide) !== before) continue;
        pushUndo();
        setSlideMedia(slide, { src, kind: "image", name: "スライドに合わせた生成画像" }, candidate.slotted ? null : { x: 0.79, y: 0.08, w: 0.17, h: 0.19 });
        if (!slide.photoMotion) slide.photoMotion = "zoom";
        markChanged({ structural: true });
      } catch (error) {
        run.failures += 1;
        if (/401|認証|接続/.test(error.message)) run.cancelled = true;
      } finally {
        run.inFlight -= 1;
        showAutoImageProgress(run);
      }
    }
  };
  await Promise.all([worker(), worker()]);
  if (autoImageRun !== run) return;
  showAutoImageProgress(run);
  const present = picturePlan(deck).present;
  toast(present >= run.target ? `${deck.slides.length}枚中${present}枚に画像を用意しました` : run.cancelled ? "画像生成を停止しました。残りは「再試行」で追加できます" : `画像は${present}/${run.target}枚です。残りは「再試行」で追加できます`);
}

let progressTimer = null;
let progressStarted = 0;
const STAGE_STEP = [[/受付|起動|送信/, 0], [/検討|骨子/, 1], [/書き出し|生成|作成/, 2], [/品質|修正|確認/, 3], [/完了|できました/, 4]];

function showProgress(stage, detail = "", { failed = false } = {}) {
  const box = $("generationProgress");
  box.classList.remove("hidden");
  box.classList.toggle("failed", failed);
  $("progressStage").textContent = stage;
  $("progressDetail").textContent = detail;
  const step = STAGE_STEP.find(([pattern]) => pattern.test(stage))?.[1] ?? 1;
  box.querySelectorAll(".progress-steps li").forEach((li, i) => { li.className = i < step || step === 4 ? "done" : i === step && !failed ? "active" : ""; });
  if (!progressTimer && !failed && step < 4) {
    progressStarted = Date.now();
    progressTimer = setInterval(() => { $("progressElapsed").textContent = `${Math.round((Date.now() - progressStarted) / 1000)}秒`; }, 1000);
  }
}

function finishProgress(stage, detail, failed = false) {
  showProgress(stage, detail, { failed });
  clearInterval(progressTimer);
  progressTimer = null;
  setCreateBusy(false);
}

function watchJob(jobId, { onProgress, onDone, onFail }) {
  let finished = false;
  let stream = null;
  const handle = (job) => {
    if (finished) return;
    onProgress?.(job);
    if (job.status === "completed") {
      finished = true;
      stream?.close();
      onDone(job);
      if (job.provider === "local") toast(`Codexが使えなかったため、予備AI（${job.localModel}）で作りました。仕上がりを確認してください`);
    }
    if (job.status === "failed") { finished = true; stream?.close(); onFail(job); }
  };
  // A dropped connection (the server waking up, a proxy timeout) is not a failed job: try a few more times.
  let misses = 0;
  async function poll() {
    try {
      const job = await jsonFetch(`/api/decks/${encodeURIComponent(jobId)}`, { method: "GET" });
      misses = 0;
      handle(job);
      if (!finished) setTimeout(poll, 1500);
    } catch (error) {
      if (finished) return;
      const gone = /404|見つかりません|not found/i.test(error.message);
      if (!gone && ++misses <= 5) { setTimeout(poll, 1500 * misses); return; }
      finished = true;
      onFail({ error: gone ? "サーバーが再起動したため、処理が中断されました。もう一度お試しください。" : error.message });
    }
  }
  if (typeof EventSource === "function") {
    stream = new EventSource(`/api/decks/${encodeURIComponent(jobId)}/events`);
    stream.addEventListener("progress", (event) => { try { handle(JSON.parse(event.data)); } catch { /* ignore malformed event */ } });
    stream.onerror = () => {
      if (finished) return;
      stream.close();
      poll();
    };
  } else poll();
}

function toneText() {
  return $("toneInput").selectedOptions[0]?.textContent || "標準ビジネス";
}

function generationInput(extra = {}) {
  const form = formState();
  return {
    deckTitle: "",
    brief: form.brief,
    audience: form.audience,
    purpose: form.purpose,
    tone: toneText(),
    settings: { slideCount: Math.max(2, Math.min(50, form.slideCount)), textDensity: form.density },
    ...extra,
  };
}

function setCreateBusy(busy, label = "") {
  state.createBusy = busy;
  for (const id of ["askChatGptBtn", "quickDeckBtn", "outlineBuildBtn", "outlineReviseBtn"]) $(id).disabled = busy;
  $("askChatGptBtn").textContent = busy && label ? label : state.outline ? "骨子を作り直す →" : "骨子を作る →";
}

function requireBrief() {
  if ($("briefInput").value.trim()) return true;
  showStatus("generationStatus", "伝えたいことや素材を入力してください。", "error");
  $("briefInput").focus();
  return false;
}

function requireCodex() {
  if (state.codexAuthorized) return true;
  showStatus("generationStatus", "先に「Codexに接続」から認証してください（雛形・サンプル・JSONは接続なしで使えます）。", "error");
  $("standaloneConnection").scrollIntoView({ behavior: "smooth", block: "center" });
  return false;
}

async function generateDeck({ outline = null } = {}) {
  if (!requireBrief() || !requireCodex()) return;
  clearStatus("generationStatus");
  setCreateBusy(true, "作成中…");
  showProgress("送信中", outline ? "骨子に沿ってスライドを書いています。" : "設定と素材を送信しています。");
  renderLiveSlides([], outline || []);
  if (outline) $("outlinePanel").classList.add("hidden");
  $("generationProgress").scrollIntoView({ behavior: "smooth", block: "center" });
  saveCurrent();
  try {
    const { jobId } = await jsonFetch("/api/decks", {
      method: "POST",
      body: JSON.stringify(generationInput(outline ? { outline, deckTitle: state.outline?.deckTitle || "" } : {})),
    });
    setPill("Codexで生成中", "busy");
    watchJob(jobId, {
      onProgress: (job) => {
        showProgress(job.attempt > 1 && job.status !== "completed" ? `${job.stage}（${job.attempt}回目）` : job.stage, job.detail);
        if (job.partial?.length) renderLiveSlides(job.partial, outline || []);
      },
      onDone: (job) => {
        finishProgress("完了", job.detail);
        $("liveSlides").replaceChildren();
        restorePill();
        state.historyId = null;
        state.outline = null;
        renderOutlinePanel();
        const deck = normalizeDeck(job.deck);
        // The template dresses the cover, chapters and close with its own ripples: no moving backdrop on top.
        deck.motion = { ...deck.motion, backdrop: "none" };
        for (const slide of deck.slides) delete slide.backdrop;
        loadDeck(deck, { source: "Codex" });
        afterGeneration().then(() => autoIllustrateDeck(deck)).catch((error) => toast(`画像の自動生成を開始できませんでした：${error.message}`));
      },
      onFail: (job) => {
        finishProgress("生成できませんでした", job.error || job.detail || "不明なエラー", true);
        renderOutlinePanel();
        restorePill();
        if (/接続|認証|401/.test(job.error || "")) checkCodexStatus();
      },
    });
  } catch (error) {
    finishProgress("生成できませんでした", error.message, true);
    if (/接続|認証/.test(error.message)) checkCodexStatus();
  }
}

/** After generation: measure every slide in the real fonts, then let the AI shorten what still does not fit. */
async function afterGeneration() {
  await wait(300);
  await measureAll();
  renderFilmstrip();
  renderIssueSummary();
  refreshInspectorIssues();
  const over = state.deck.slides.map((_, i) => i).filter((i) => overflowFor(i).length);
  if (over.length && state.codexAuthorized) {
    toast(`${over.length}枚で文字が枠に収まりません。AIで順に短くしています`);
    await fixAllOverflow();
  }
}

/** First step: a quick skeleton the user can reshape before any slide is written. */
async function generateOutline({ instruction = "" } = {}) {
  if (!requireBrief() || !requireCodex()) return;
  clearStatus("generationStatus");
  setCreateBusy(true, instruction ? "骨子を直しています…" : "骨子を作成中…");
  showProgress("送信中", instruction ? "骨子への指示を送っています。" : "骨子を考えています（数十秒）。");
  $("liveSlides").replaceChildren();
  saveCurrent();
  try {
    const body = generationInput(instruction && state.outline ? { outline: state.outline.slides, instruction } : {});
    const { jobId } = await jsonFetch("/api/decks/outline", { method: "POST", body: JSON.stringify(body) });
    watchJob(jobId, {
      onProgress: (job) => showProgress(job.stage, job.detail),
      onDone: (job) => {
        finishProgress("骨子ができました", job.detail);
        state.outline = job.outline;
        $("outlineInstruction").value = "";
        renderOutlinePanel();
        saveCurrent();
        $("outlinePanel").scrollIntoView({ behavior: "smooth", block: "start" });
      },
      onFail: (job) => finishProgress("骨子を作れませんでした", job.error || job.detail, true),
    });
  } catch (error) {
    finishProgress("骨子を作れませんでした", error.message, true);
    if (/接続|認証/.test(error.message)) checkCodexStatus();
  }
}

const OUTLINE_TYPES = () => SLIDE_TYPES.filter((type) => !["title", "closing"].includes(type));

/** Layout choices grouped by how they look, so picking a different look is one glance away. */
function typeOptions(selected) {
  const groups = Object.entries(LOOKS).map(([, look]) => [look.label, look.types]);
  const grouped = new Set(groups.flatMap(([, types]) => types));
  const other = OUTLINE_TYPES().filter((type) => !grouped.has(type));
  return [...groups, ["区切り", other]].map(([label, types]) => h("optgroup", { label },
    types.filter((type) => TYPE_INFO[type]).map((type) => h("option", { value: type, selected: type === selected }, typeLabel(type)))));
}

function renderOutlineVariety() {
  const note = $("outlineVariety");
  const issues = state.outline ? varietyIssues(state.outline.slides) : [];
  note.classList.toggle("hidden", !issues.length);
  note.replaceChildren(...issues.map((issue) => h("div", {}, `⚠ ${issue.message}`)),
    issues.length ? h("div", { class: "hint" }, `内容に合わせて形を変えると伝わりやすくなります：${LOOK_ADVICE.join("／")}。「AIで骨子を直す」に「見た目を散らして」と頼むこともできます。`) : null);
}

function renderOutlinePanel() {
  const panel = $("outlinePanel");
  const outline = state.outline;
  panel.classList.toggle("hidden", !outline);
  setCreateBusy(Boolean(state.createBusy));
  if (!outline) return;
  $("outlineTitle").textContent = `骨子：${outline.deckTitle}`;
  $("outlineReply").textContent = outline.reply || "行を直接書き換えたり、並べ替えたり、下の欄でAIに指示したりできます。";
  const items = outline.slides;
  const last = items.length - 1;
  const edit = () => saveCurrent();
  $("outlineList").replaceChildren(...items.map((item, index) => {
    const fixed = index === 0 || index === last;
    return h("li", { class: "outline-item" },
      h("span", { class: "no" }, String(index + 1)),
      fixed
        ? h("span", { class: "hint", style: { "padding-top": "7px" } }, typeLabel(item.type))
        : h("select", { "aria-label": "レイアウト", onchange: (event) => { item.type = event.target.value; edit(); renderOutlineVariety(); } }, typeOptions(item.type)),
      h("div", { class: "texts" },
        h("input", { class: "title-input", type: "text", value: item.title, maxlength: 90, "aria-label": "タイトル", oninput: (event) => { item.title = event.target.value; edit(); } }),
        index === 0 ? null : h("input", { type: "text", value: item.takeaway || "", maxlength: 160, placeholder: "このスライドの結論（一文）", "aria-label": "結論", oninput: (event) => { item.takeaway = event.target.value; edit(); } }),
        index === 0 ? null : h("textarea", { rows: 1, maxlength: 400, placeholder: "載せる中身のメモ（動き：クリックで順番に／詳細を付ける、なども書けます）", "aria-label": "中身のメモ", oninput: (event) => { item.content = event.target.value; edit(); } }, item.content || "")),
      h("span", { class: "btns" },
        h("button", { class: "btn btn-ghost", type: "button", title: "上へ", disabled: fixed || index === 1, onclick: () => { [items[index - 1], items[index]] = [items[index], items[index - 1]]; renderOutlinePanel(); edit(); } }, "↑"),
        h("button", { class: "btn btn-ghost", type: "button", title: "下へ", disabled: fixed || index === last - 1, onclick: () => { [items[index + 1], items[index]] = [items[index], items[index + 1]]; renderOutlinePanel(); edit(); } }, "↓"),
        h("button", { class: "btn btn-ghost btn-danger", type: "button", title: "削除", disabled: fixed || items.length <= 2, onclick: () => { items.splice(index, 1); renderOutlinePanel(); edit(); } }, "✕")));
  }));
  $("outlineCount").textContent = `全${items.length}枚`;
  renderOutlineVariety();
}

/** Second step: write the slides from the agreed outline. */
function buildFromOutline() {
  const outline = state.outline;
  if (!outline) return;
  const slides = outline.slides.map((item) => ({ type: item.type, title: strip(item.title) || "（タイトル）", takeaway: strip(item.takeaway || ""), content: strip(item.content || "") }));
  generateDeck({ outline: slides });
}

/** Slides written so far, drawn as they arrive while the rest is still being generated. */
function renderLiveSlides(partial, planned = []) {
  const box = $("liveSlides");
  const total = Math.max(partial.length + 1, planned.length || 0);
  const liveDeck = { title: strip(partial[0]?.title) || "作成中", theme: preferredTheme(), transition: "fade", motion: DEFAULT_MOTION, slides: [] };
  const cards = partial.map((raw, index) => {
    let picture = box.querySelector(`.live-card[data-index="${index}"] .hs-scaler`);
    if (!picture) {
      try {
        const slide = normalizeSlide(raw, index, total);
        liveDeck.slides = partial.map((item, i) => (i === index ? slide : item));
        picture = E.mount(E.render(slide, { deck: liveDeck, index, mode: "thumb", assetBase: "/assets/" }));
      } catch { picture = null; }
    }
    return h("div", { class: "live-card", "data-index": index }, picture,
      h("b", {}, `${index + 1}. ${strip(raw.title || raw.message || typeLabel(raw.type))}`),
      h("span", {}, typeLabel(raw.type)));
  });
  const pending = planned.slice(partial.length).map((item, offset) => h("div", { class: "live-card pending" },
    h("b", {}, `${partial.length + offset + 1}. ${strip(item.title)}`), h("span", {}, typeLabel(item.type))));
  useFonts([preferredTheme()]);
  box.replaceChildren(...cards, ...pending);
}

function setAiBusy(busy, text = "") {
  state.aiBusy = busy;
  const status = document.getElementById("aiStatus");
  if (status && text) status.textContent = text;
  $("fixAllBtn").disabled = busy;
  $("deckReviseBtn").disabled = busy;
  setPill(busy ? "AIが作業中" : state.codexAuthorized ? (state.codexImageAuthorized ? "Codex接続済み" : "予備AIで動作中") : "Codex未接続", busy ? "busy" : state.codexAuthorized ? "ok" : "");
}

/** The AI sees "[画像あり]" / "[写真・動画あり]" instead of pictures; put the originals back. */
function restoreImages(slides, previous) {
  const isImage = (value) => typeof value === "string" && value.startsWith("data:");
  const byTitle = new Map(previous.map((slide) => [strip(slide.title), slide]));
  slides.forEach((slide, index) => {
    const source = previous.length === slides.length ? previous[index] : byTitle.get(strip(slide.title));
    for (const key of ["customImage", "image"]) {
      if (typeof slide[key] === "string" && slide[key].startsWith("[")) delete slide[key];
      if (source && isImage(source[key]) && !slide[key] && (key === "customImage" || slide.type === "imageText")) slide[key] = source[key];
    }
    if (typeof slide.media === "string" || (slide.media && !slide.media.src)) delete slide.media;
    if (source?.media && !slide.media) slide.media = clone(source.media);
    if (source?.drillOf && !slide.drillOf && TITLED(slide.type)) slide.drillOf = source.drillOf;
    if (source?.imagePlacement && slide.customImage === source.customImage && !slide.imagePlacement) slide.imagePlacement = source.imagePlacement;
    // What people placed by hand stays: a 白紙 page comes back as it was, objects stay with their slide.
    if (source?.type === "blank") { slides[index] = clone(source); return; }
    for (const key of ["elements", "timeline", "sid", "transitionDur", "transitionSound", "transitionDir", "advance", "background"]) if (source?.[key] != null && slide[key] == null) slide[key] = clone(source[key]);
  });
  // A slide with objects that found no place in the new deck is kept (before the close), so no work is lost.
  const placed = new Set(slides.flatMap((slide) => (slide.elements || []).map((o) => o.id)));
  for (const prev of previous) {
    if (prev.elements?.length && !prev.elements.some((o) => placed.has(o.id))) slides.splice(Math.max(1, slides.length - 1), 0, clone(prev));
  }
  // Photos and videos whose slide disappeared float on the first slides that have none.
  const kept = new Set(slides.map((slide) => slide.media?.src).filter(Boolean));
  const orphans = previous.filter((slide) => slide.media?.src && !kept.has(slide.media.src)).map((slide) => slide.media);
  for (const slide of slides) {
    if (!orphans.length) break;
    if (!slide.media && !slide.customImage && TITLED(slide.type)) setSlideMedia(slide, clone(orphans.shift()));
  }
  return slides;
}

function reviseSlide(index, instruction, issues = []) {
  return new Promise((resolve) => {
    if (state.aiBusy) return resolve(false);
    setAiBusy(true, "Codexに依頼しています…");
    const originalJson = JSON.stringify(state.deck.slides[index]);
    jsonFetch("/api/decks/revise", {
      method: "POST",
      body: JSON.stringify({ deck: serverDeck(), slideIndex: index, instruction, issues: issues.map((issue) => ({ field: String(issue.field).slice(0, 80), message: String(issue.message).slice(0, 200) })) }),
    }).then(({ jobId }) => watchJob(jobId, {
      onProgress: (job) => { const status = document.getElementById("aiStatus"); if (status && state.selected === index) status.textContent = `${job.stage}：${job.detail}`; },
      onDone: (job) => {
        setAiBusy(false, "作り直しました。⌘Zで元に戻せます。");
        const edited = JSON.stringify(state.deck.slides[index]) !== originalJson;
        if (edited) {
          toast(`${index + 1}枚目は作業中に編集されたため、結果を新しいスライドとして追加しました`);
          insertSlide(index + 1, normalizeSlide(restoreImages([job.slide], [JSON.parse(originalJson)])[0], index + 1, state.deck.slides.length + 1));
        } else {
          replaceSlide(index, restoreImages([job.slide], [JSON.parse(originalJson)])[0]);
          toast(`${index + 1}枚目を作り直しました`);
        }
        const invented = unsourcedNote(job.issues);
        if (invented) toast(`${index + 1}枚目：${invented}`);
        resolve(true);
      },
      onFail: (job) => { setAiBusy(false, `作り直せませんでした：${job.error || job.detail}`); toast("作り直しに失敗しました"); resolve(false); },
    })).catch((error) => { setAiBusy(false, `作り直せませんでした：${error.message}`); resolve(false); });
  });
}

/** "素材にない数値があります：「43%」「77」" for figures the material does not contain. */
function unsourcedNote(issues = []) {
  const values = [...new Set((issues || []).filter((issue) => issue.kind === "unsourced").flatMap((issue) => issue.values || []))];
  return values.length ? `素材にない数値があります：${values.slice(0, 4).map((value) => `「${value}」`).join("")}${values.length > 4 ? "ほか" : ""}。確認してください` : "";
}

function aiInsertSlide(at, instruction, onStatus = () => {}, { drillOf = null } = {}) {
  return new Promise((resolve) => {
    if (state.aiBusy) return resolve(false);
    setAiBusy(true);
    jsonFetch("/api/decks/revise", { method: "POST", body: JSON.stringify({ deck: serverDeck(), slideIndex: at, instruction, mode: "insert" }) })
      .then(({ jobId }) => watchJob(jobId, {
        onProgress: (job) => onStatus(`${job.stage}：${job.detail}`),
        onDone: (job) => {
          setAiBusy(false);
          const slide = normalizeSlide(job.slide, at, state.deck.slides.length + 1);
          if (drillOf) slide.drillOf = drillOf;
          const placed = insertSlide(at, slide);
          if (placed >= 0) toast(drillOf ? `${placed + 1}枚目に深掘りページを作りました。発表中にその項目をクリックすると開きます` : `${placed + 1}枚目にスライドを追加しました`);
          resolve(placed >= 0);
        },
        onFail: (job) => { setAiBusy(false); onStatus(`作成できませんでした：${job.error || job.detail}`); resolve(false); },
      }))
      .catch((error) => { setAiBusy(false); onStatus(`作成できませんでした：${error.message}`); resolve(false); });
  });
}

async function fixAllOverflow() {
  if (state.fixingAll) return;
  state.fixingAll = true;
  const targets = state.deck.slides.map((_, i) => ({ i, list: overflowFor(i) })).filter((item) => item.list.length);
  try {
    for (const [n, { i, list }] of targets.entries()) {
      select(i);
      $("fixAllBtn").textContent = `修正中 ${n + 1}/${targets.length}`;
      const ok = await reviseSlide(i, "スライドに収まらない文字を、意味を保って短く言い換える（補足は details に回してよい）", list);
      if (!ok) break;
      await wait(400);
    }
  } finally {
    state.fixingAll = false;
    $("fixAllBtn").textContent = "あふれをAIで直す";
  }
}

const REWRITE_CHIPS = ["取り込んだ資料を伝わる構成に磨き上げる", "各スライドをもっと簡潔に", "役員が3分で判断できるよう結論を強く", "新入社員にも分かる言葉で", "数値の根拠を目立たせる", "図解・表を増やして文字を減らす", "似た見た目のスライドを減らし、内容に合う形で変化をつける", "動きのある見せ方に（クリックで順番・詳細）", "ストーリーの流れを見直す"];
const BRUSHUP_INSTRUCTION = "取り込んだ元資料の事実・数値・写真を保ったまま、画面共有で伝わるプレゼンに磨き上げる。各本文スライドに結論（takeaway）を一文で入れ、内容に合う図解・表・グラフのレイアウトへ置き換え、補足は「クリックで開く詳細」に回し、発表用のノートを付ける";

function openRewriteDialog(prefill = "", flexible = false) {
  if (typeof prefill === "string" && prefill) $("rewriteInstruction").value = prefill;
  $("rewriteFlexible").checked = flexible;
  $("rewriteStatus").textContent = "";
  $("rewriteSubmit").disabled = state.aiBusy;
  $("rewriteChips").replaceChildren(...REWRITE_CHIPS.map((chip) => h("button", { class: "chip", type: "button", onclick: () => {
    const area = $("rewriteInstruction");
    area.value = area.value.trim() ? `${area.value.trim()}、${chip}` : chip;
  } }, chip)));
  $("rewriteDialog").showModal();
}

function rewriteDeck() {
  const instruction = $("rewriteInstruction").value.trim();
  if (!instruction) return $("rewriteInstruction").focus();
  if (state.aiBusy) return;
  setAiBusy(true);
  $("rewriteSubmit").disabled = true;
  $("rewriteStatus").textContent = "Codexに依頼しています…";
  const previous = clone(state.deck.slides);
  const baseDeck = clone(state.deck);
  jsonFetch("/api/decks/rewrite", { method: "POST", body: JSON.stringify({ deck: serverDeck(), instruction, settings: { textDensity: formState().density, flexibleCount: $("rewriteFlexible").checked } }) })
    .then(({ jobId }) => watchJob(jobId, {
      onProgress: (job) => { $("rewriteStatus").textContent = `${job.stage}：${job.detail}`; },
      onDone: (job) => {
        setAiBusy(false);
        $("rewriteSubmit").disabled = false;
        const next = normalizeDeck(job.deck, baseDeck);
        restoreImages(next.slides, previous);
        $("rewriteDialog").close();
        $("rewriteInstruction").value = "";
        pushChat({ role: "user", text: `資料全体を見直して：${instruction}` });
        proposeWholeDeck("資料全体を見直しました。変更点を確認して、よければ採用してください。", next.slides, { ...baseDeck, slides: previous });
      },
      onFail: (job) => { setAiBusy(false); $("rewriteSubmit").disabled = false; $("rewriteStatus").textContent = `見直せませんでした：${job.error || job.detail}`; },
    }))
    .catch((error) => { setAiBusy(false); $("rewriteSubmit").disabled = false; $("rewriteStatus").textContent = `見直せませんでした：${error.message}`; });
}

// ---------------------------------------------------------------- AI chat (conversational editing)

const proposals = new Map();
const variantSets = new Map();
const wantsVariants = (text) => /別案|3案|三案|案を[3３三]つ|パターンを[3３三]/.test(text);
const IMAGE_REQUEST = /(?:gpt[\s-]*image|(?:画像|イラスト|写真)[^。！？!?\n]{0,18}(?:生成|描いて|作って|作成)|(?:生成|作成)[^。！？!?\n]{0,10}(?:画像|イラスト))/i;

function chatKey() {
  return state.historyId || "current";
}

function loadChat() {
  let all = {};
  try { all = JSON.parse(localStorage.getItem(STORAGE.chat) || "{}"); } catch { all = {}; }
  // Proposals live in memory only; after a reload an unanswered one can no longer be applied.
  state.chat.messages = (all[chatKey()] || []).map((message) => ({ ...message,
    ...(message.proposal?.status === "pending" ? { proposal: { ...message.proposal, status: "expired" } } : {}),
    ...(message.generatedImage?.status === "pending" ? { generatedImage: { ...message.generatedImage, status: "expired" } } : {}),
  }));
  renderChat();
  $("memoToggle").classList.toggle("active", Boolean(state.deck?.memo?.trim()));
}

function saveChat() {
  try {
    const all = JSON.parse(localStorage.getItem(STORAGE.chat) || "{}");
    all[chatKey()] = state.chat.messages.slice(-40).map(({ id, role, text, at, suggestions, proposal, variants, review, error, generatedImage }) => ({ id, role, text, at, suggestions, error, proposal: proposal && { status: proposal.status, summary: proposal.summary }, variants, review, generatedImage: generatedImage && { status: generatedImage.status, index: generatedImage.index } }));
    const keys = Object.keys(all);
    for (const key of keys.slice(0, Math.max(0, keys.length - 20))) delete all[key];
    localStorage.setItem(STORAGE.chat, JSON.stringify(all));
  } catch { /* the conversation is a convenience; the deck itself is saved elsewhere */ }
}

/** Shows a task pane (reopening the closed task pane, unless `reveal` is false: the one remembered at startup). */
function setPanel(panel, { reveal = true } = {}) {
  state.panel = panel;
  if (reveal && !shell.paneOpen("side")) shell.setPane("side", true);
  if (panel === "chat") $("chatTab").textContent = "✦ AIと話す";
  $("chatPane").hidden = panel !== "chat";
  $("inspector").hidden = panel !== "form";
  $("formatPane").hidden = panel !== "format";
  $("animPane").hidden = panel !== "anim";
  $("commentPane").hidden = panel !== "comment";
  $("a11yPane").hidden = panel !== "a11y";
  // The アクセシビリティ tab shows once the checker has been opened (as PowerPoint's task pane).
  if (panel === "a11y") $("a11yTab").hidden = false;
  $("a11yTab").setAttribute("aria-selected", String(panel === "a11y"));
  $("animTab").setAttribute("aria-selected", String(panel === "anim"));
  $("commentTab").setAttribute("aria-selected", String(panel === "comment"));
  $("chatTab").setAttribute("aria-selected", String(panel === "chat"));
  $("formTab").setAttribute("aria-selected", String(panel === "form"));
  $("formatTab").setAttribute("aria-selected", String(panel === "format"));
  if (panel === "format") editorUi.renderPane();
  if (panel === "anim") editorUi.renderAnimPane();
  if (panel === "comment") comments.render();
  if (panel === "a11y") a11y.render();
  // The animation order marks on the stage show while the animation pane is open.
  editor.draw();
  try { localStorage.setItem(STORAGE.panel, panel); } catch { /* optional */ }
  if (panel === "chat") { renderChatContext(); scrollChat(); }
}

function scrollChat() {
  const log = $("chatLog");
  requestAnimationFrame(() => { log.scrollTop = log.scrollHeight; });
}

/** "@3", "＠3", "3枚目" in the message point the AI at those slides. */
function mentionedSlides(text) {
  const total = slideCount();
  const found = [...text.matchAll(/[@＠](\d{1,2})|(\d{1,2})\s*枚目/g)].map((match) => Number(match[1] ?? match[2]) - 1);
  return [...new Set(found.filter((index) => index >= 0 && index < total))];
}

function renderChatContext() {
  const box = $("chatContext");
  if (!state.deck) { box.replaceChildren(); return; }
  const mentioned = mentionedSlides($("chatInput").value);
  const slide = state.deck.slides[state.selected];
  box.replaceChildren(
    h("span", {}, "見ているスライド："),
    h("span", { class: "tag" }, `${state.selected + 1}枚目 ${strip(slide?.title || typeLabel(slide?.type)).slice(0, 14)}`),
    ...(mentioned.length ? [h("span", {}, "指定："), ...mentioned.map((index) => h("span", { class: "tag" }, `@${index + 1}`))] : []),
    ...(state.chat.attachment ? [h("span", { class: "tag attach" }, `📎 ${state.chat.attachment.name}`, h("button", { type: "button", "aria-label": "添付を外す", onclick: () => { state.chat.attachment = null; renderChatContext(); } }, "✕"))] : []));
}

function starterChips() {
  const slide = state.deck?.slides[state.selected];
  const chips = ["この1枚をもっと簡潔に", "全体の流れを確認して"];
  if (slide && ["process", "processList", "timeline", "roadmap", "stepUp", "cards", "headerCards"].includes(slide.type)) chips.push("この1枚をクリックで1つずつ出して");
  if (slide && TITLED(slide.type)) chips.push("この1枚の項目にクリックで開く詳細を付けて");
  if (slide && ["content", "agenda", "compare"].includes(slide.type)) chips.push("この1枚を図解にして");
  if (slide && /\d/.test(JSON.stringify(slide))) chips.push("この1枚の数字を目立たせて");
  chips.push("この1枚の別案を3つ見せて", "テーマを発表の雰囲気に合わせて", "足りないスライドを提案して");
  return chips.slice(0, 7);
}

function chatBubble(message) {
  const time = message.at ? new Date(message.at).toLocaleTimeString("ja-JP", { hour: "2-digit", minute: "2-digit" }) : "";
  return h("div", { class: `msg ${message.role}${message.error ? " error" : ""}`, "data-id": message.id },
    h("div", { class: "bubble" }, message.text),
    message.proposal ? proposalCard(message) : null,
    message.variants ? variantsCard(message) : null,
    message.review ? reviewCard(message) : null,
    message.generatedImage ? generatedImageCard(message) : null,
    message.error && message.retry ? h("button", { class: "btn", type: "button", style: { "align-self": "flex-start", height: "28px" }, onclick: () => sendChat(message.retry) }, "もう一度送る") : null,
    h("span", { class: "msg-meta" }, time));
}

function renderChat() {
  const log = $("chatLog");
  if (!log) return;
  const messages = state.chat.messages;
  const nodes = messages.map(chatBubble);
  if (!messages.length) {
    const canAi = state.codexAuthorized;
    nodes.push(h("div", { class: "chat-empty" },
      h("span", { class: "spark" }, "✦"),
      h("strong", {}, "AIと話しながら仕上げましょう"),
      canAi ? "直してほしいこと・付けたい動きをそのまま書いてください。変更は提案として表示され、確認してから採用できます。" : "Codexに接続すると、会話で資料を直せます（「作成」画面の上部から接続）。",
      canAi && state.deck ? h("div", { class: "ai-chips" }, starterChips().map((chip) => h("button", { class: "chip", type: "button", onclick: () => sendChat(chip) }, chip))) : null));
  }
  const last = messages.at(-1);
  if (last?.role === "assistant" && last.suggestions?.length && !state.chat.busy) {
    nodes.push(h("div", { class: "suggest-chips" }, last.suggestions.map((chip) => h("button", { class: "chip", type: "button", onclick: () => sendChat(chip) }, chip))));
  }
  if (state.chat.busy) {
    nodes.push(h("div", { class: "msg assistant typing" }, h("div", { class: "bubble" }, h("span", { class: "dots" }, h("i"), h("i"), h("i")), h("span", { id: "chatProgress" }, state.chat.partialReply || state.chat.progress || "考えています…"))));
  }
  log.replaceChildren(...nodes);
  $("chatSend").disabled = state.chat.busy;
  scrollChat();
}

function pushChat(message) {
  const entry = { id: `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`, at: Date.now(), ...message };
  state.chat.messages.push(entry);
  saveChat();
  renderChat();
  return entry;
}

async function sendChat(raw) {
  const text = String(raw ?? $("chatInput").value).trim();
  if (!text || state.chat.busy || !state.deck) return;
  if (isFinal(state.deck)) { toast("最終版です。編集するにはメッセージ バーの「編集する」を押してください"); return; }
  if (state.panel !== "chat") setPanel("chat");
  if (raw == null) $("chatInput").value = "";
  const imageRequest = IMAGE_REQUEST.test(text);
  if (!imageRequest && wantsVariants(text)) {
    const target = mentionedSlides(text)[0] ?? state.selected;
    return requestVariants(target, text.replace(/[@＠]\d{1,2}|\d{1,2}\s*枚目/g, "").replace(/(の)?(別案|3案|三案)(を)?(3つ|三つ)?(見せて|出して|作って)?/g, "").trim(), text);
  }
  const history = state.chat.messages.filter((message) => !message.error).slice(-10).map((message) => ({ role: message.role, text: String(message.text).slice(0, 4000) }));
  pushChat({ role: "user", text: state.chat.attachment ? `${text}\n📎 ${state.chat.attachment.name}` : text });
  if (imageRequest) return generateChatImage(text);
  if (!state.codexAuthorized) {
    pushChat({ role: "assistant", text: "Codexに接続すると、会話で資料を直せます。「作成」画面の上部から接続してください。", error: true });
    return;
  }
  const base = JSON.stringify(state.deck.slides);
  const baseDeck = clone(state.deck);
  state.chat.busy = true;
  state.chat.progress = "依頼を送っています…";
  setAiBusy(true);
  renderChat();
  const done = () => { state.chat.busy = false; state.chat.progress = ""; state.chat.partialReply = ""; setAiBusy(false); };
  const focus = mentionedSlides(text);
  const attachment = state.chat.attachment;
  state.chat.attachment = null;
  renderChatContext();
  try {
    const { jobId } = await jsonFetch("/api/decks/chat", { method: "POST", body: JSON.stringify({ deck: serverDeck(), message: text, history, current: state.selected, focus, ...(attachment ? { attachment } : {}) }) });
    watchJob(jobId, {
      onProgress: (job) => {
        state.chat.progress = job.status === "completed" ? "まとめています…" : `${job.stage}：${job.detail}`;
        if (job.partialReply) state.chat.partialReply = job.partialReply;
        const el = document.getElementById("chatProgress");
        if (el) el.textContent = state.chat.partialReply || state.chat.progress;
        scrollChat();
      },
      onDone: (job) => {
        done();
        const chat = job.chat;
        if (!chat) { pushChat({ role: "assistant", text: "回答を受け取れませんでした。", error: true, retry: text }); return; }
        const message = pushChat({ role: "assistant", text: chat.reply, suggestions: chat.suggestions, proposal: chat.changed ? { status: "pending", summary: chat.summary || "変更の提案" } : null });
        if (chat.changed) createProposal(message.id, chat, base, baseDeck);
        if (state.panel !== "chat") { $("chatTab").replaceChildren("✦ AIと話す", h("span", { class: "badge-dot", title: "AIから返事が届いています" })); toast(chat.changed ? "AIから変更の提案が届きました" : "AIから返事が届きました"); }
      },
      onFail: (job) => { done(); pushChat({ role: "assistant", text: `うまく処理できませんでした：${job.error || job.detail}`, error: true, retry: text }); },
    });
  } catch (error) {
    done();
    if (error.body?.code === "use_image_generation") return generateChatImage(text);
    pushChat({ role: "assistant", text: `送信できませんでした：${error.message}`, error: true, retry: text });
    if (/接続|認証|401/.test(error.message)) checkCodexStatus();
  }
}

async function generateChatImage(text) {
  if (!state.codexImageAuthorized) { pushChat({ role: "assistant", text: "画像生成にはCodexへの接続が必要です（予備AIは画像を作れません）。画面上部からCodexに接続してください。", error: true }); return; }
  const index = mentionedSlides(text)[0] ?? state.selected;
  const claim = state.deck.slides[index].takeaway || state.deck.slides[index].message || state.deck.slides[index].title || "";
  state.chat.busy = true;
  state.chat.progress = "画像生成を開始しています…";
  setAiBusy(true);
  renderChat();
  const done = () => { state.chat.busy = false; state.chat.progress = ""; setAiBusy(false); };
  try {
    const deck = serverDeck();
    // A photo kept in this browser goes along as the reference when the request asks to build on it.
    const media = state.deck.slides[index].media;
    if (media?.kind === "image" && /^(idb|asset):/.test(media.src)) {
      const blob = await mediaBlob(media.src).catch(() => null);
      if (blob && blob.size < 2_000_000) deck.slides[index].media = { ...media, src: await blobToDataUrl(blob) };
    }
    const { jobId } = await jsonFetch("/api/decks/image", { method: "POST", body: JSON.stringify({ deck, message: text, index }) });
    watchJob(jobId, {
      onProgress: (job) => { state.chat.progress = `${job.stage}：${job.detail}`; const el = document.getElementById("chatProgress"); if (el) el.textContent = state.chat.progress; },
      onDone: async (job) => {
        try {
          if (!job.image?.url) throw new Error("生成画像が返されませんでした");
          const response = await fetch(job.image.url, { credentials: "same-origin" });
          if (!response.ok) throw new Error(`画像の取得に失敗しました（${response.status}）`);
          const blob = await response.blob();
          const src = await putMedia(blob, `生成画像_${index + 1}枚目`);
          done();
          pushChat({ role: "assistant", text: `${index + 1}枚目の主張に合わせて画像を新しく生成しました。よければスライドに配置してください（スライドへドラッグして好きな位置にも置けます）。`, generatedImage: { status: "pending", index, claim, src, url: mediaUrls[src] } });
        } catch (error) { done(); pushChat({ role: "assistant", text: `画像を受け取れませんでした：${error.message}`, error: true, retry: text }); }
      },
      onFail: (job) => { done(); pushChat({ role: "assistant", text: `画像を生成できませんでした：${job.error || job.detail}`, error: true, retry: text }); },
    });
  } catch (error) { done(); pushChat({ role: "assistant", text: `画像生成を開始できませんでした：${error.message}`, error: true, retry: text }); }
}

function generatedImageCard(message) {
  const meta = message.generatedImage;
  const url = meta.url || mediaUrls[meta.src];
  return h("div", { class: "generated-image-card" },
    url ? h("img", { src: url, alt: `${meta.index + 1}枚目向けに生成した画像`, draggable: meta.status === "pending" ? "true" : "false", ondragstart: (event) => {
      event.dataTransfer.setData("text/x-hs-generated-image", message.id);
      event.dataTransfer.effectAllowed = "copy";
    } }) : null,
    h("span", { class: "hint" }, meta.status === "expired" ? "画像の確認期限が切れました。もう一度生成してください。" : meta.status === "applied" ? `${meta.index + 1}枚目に配置しました` : meta.status === "dismissed" ? "この画像は採用しませんでした" : "スライドへドラッグして好きな位置に置くか、ボタンで配置できます"),
    meta.status === "pending" ? h("div", { class: "generated-image-actions" },
      h("button", { class: "btn btn-primary", type: "button", onclick: () => { if (state.selected !== meta.index) select(meta.index); applyGeneratedImage(message); } }, "スライドに配置"),
      h("button", { class: "btn", type: "button", onclick: () => { meta.status = "dismissed"; saveChat(); renderChat(); } }, "使わない"),
      h("button", { class: "btn btn-ghost", type: "button", onclick: () => { $("chatInput").value = `${meta.index + 1}枚目の画像をもう一度生成して、`; $("chatInput").focus(); } }, "別の画像を生成")) : null);
}

function applyGeneratedImage(message, placement = null) {
  const meta = message.generatedImage;
  if (!meta || meta.status !== "pending" || !state.deck?.slides[meta.index]) return toast("この画像は配置できません。再生成してください");
  if (state.selected !== meta.index) return toast(`${meta.index + 1}枚目を表示してから画像を置いてください`);
  pushUndo();
  setSlideMedia(state.deck.slides[meta.index], { src: meta.src, kind: "image", name: "生成画像" }, placement);
  meta.status = "applied";
  markChanged({ structural: true });
  saveChat();
  renderChat();
  toast("画像を配置しました（⌘Zで戻せます）");
}

/** Keep the proposal in memory with its before/after pictures. */
function createProposal(id, chat, base, baseDeck, { whole = false } = {}) {
  const total = chat.slides ? chat.slides.length : baseDeck.slides.length;
  const slides = chat.slides ? chat.slides.map((slide, index) => normalizeSlide(slide, index, total)) : clone(baseDeck.slides);
  const selected = new Set();
  (chat.items || []).forEach((item, index) => { if (item.from == null) selected.add(`n${index}`); else if (item.changed) selected.add(`c${index}`); });
  for (const index of chat.deleted || []) selected.add(`d${index}`);
  if (chat.moved) selected.add("move");
  if (chat.deckTitle) selected.add("title");
  if (chat.transition) selected.add("transition");
  const motion = chat.motion && typeof chat.motion === "object" && Object.keys(chat.motion).length ? chat.motion : null;
  if (motion) selected.add("motion");
  const proposal = { id, base, baseDeck, baseSlides: clone(baseDeck.slides), baseTitle: baseDeck.title, slides, items: chat.items || [], deleted: chat.deleted || [], moved: Boolean(chat.moved), deckTitle: chat.deckTitle, transition: chat.transition, motion, issues: chat.issues || [], selected, whole };
  proposals.set(id, proposal);
  renderChat();
}

const TRANSITION_LABEL = Object.fromEntries(Object.entries(E.TRANSITIONS).map(([key, label]) => [key, label.replace(/（.*）/, "")]));

/** "文字：マスク → タイプライター／背景：なし → 軌道" for a deck-wide motion change. */
function motionChangeText(before = {}, patch = {}) {
  const was = normalizeMotion(before);
  const parts = [];
  if (patch.kinetic) parts.push(`大きな文字：${KINETIC_INFO[was.kinetic]} → ${KINETIC_INFO[patch.kinetic]}`);
  if (patch.backdrop) parts.push(`背景：${BACKDROP_INFO[was.backdrop]} → ${BACKDROP_INFO[patch.backdrop]}`);
  if (patch.entrance) parts.push(`登場：${ENTRANCE_INFO[was.entrance]} → ${ENTRANCE_INFO[patch.entrance]}`);
  if (patch.hover) parts.push(`マウスを乗せたとき：${E.HOVERS[was.hover] ?? "変化なし"} → ${E.HOVERS[patch.hover] ?? "変化なし"}`);
  if (patch.emphasis) parts.push(`強調：${E.EMPHASES[was.emphasis]} → ${E.EMPHASES[patch.emphasis]}`);
  if (typeof patch.draw === "boolean") parts.push(`線を描く：${patch.draw ? "オン" : "オフ"}`);
  return parts.join("／");
}

function proposalCard(message) {
  const meta = message.proposal;
  const proposal = proposals.get(message.id);
  if (!proposal || meta.status !== "pending") {
    const text = { applied: "✓ 採用しました", rejected: "却下しました", reverted: "採用を取り消しました", expired: "この提案は期限切れです（再読み込みしたため）" }[meta.status] || meta.status;
    return h("div", { class: "proposal" },
      h("div", { class: "proposal-head" }, h("span", {}, `変更の提案：${meta.summary}`)),
      h("div", { class: `proposal-status ${meta.status}` }, text,
        meta.status === "applied" && proposal?.before ? h("button", { class: "btn btn-ghost", type: "button", style: { height: "26px" }, onclick: () => revertProposal(message.id) }, "元に戻す") : null));
  }
  const rows = [];
  const toggle = (key) => h("input", { type: "checkbox", checked: proposal.selected.has(key), disabled: proposal.whole, onchange: (event) => { if (event.target.checked) proposal.selected.add(key); else proposal.selected.delete(key); } });
  const designDeck = proposal.baseDeck;
  proposal.items.forEach((item, index) => {
    if (!item.changed) return;
    const isNew = item.from == null;
    const title = strip(proposal.slides[index]?.title || typeLabel(proposal.slides[index]?.type));
    const afterDeck = { ...designDeck, slides: proposal.slides };
    const before = isNew ? null : () => slidePicture(proposal.baseSlides[item.from], item.from, proposal.baseDeck);
    const after = () => slidePicture(proposal.slides[index], index, afterDeck);
    rows.push(h("div", { class: "change-row" },
      h("label", {}, toggle(isNew ? `n${index}` : `c${index}`), isNew ? `新しいスライド（${index + 1}枚目）` : `${item.from + 1}枚目`, h("span", { class: `kind${isNew ? " new" : ""}` }, isNew ? "追加" : "修正")),
      h("span", { class: "hint", style: { "font-size": "11.5px", overflow: "hidden", "text-overflow": "ellipsis", "white-space": "nowrap" } }, title),
      h("div", { class: `change-thumbs${isNew ? " single" : ""}`, title: "クリックで拡大", onclick: () => openCompare(before, after, isNew ? "追加するスライド" : `${item.from + 1}枚目の変更`) },
        before ? before() : null,
        before ? h("span", { class: "arrow" }, "→") : null,
        after())));
  });
  for (const index of proposal.deleted) {
    const before = () => slidePicture(proposal.baseSlides[index], index, proposal.baseDeck);
    rows.push(h("div", { class: "change-row" },
      h("label", {}, toggle(`d${index}`), `${index + 1}枚目`, h("span", { class: "kind del" }, "削除")),
      h("span", { class: "hint", style: { "font-size": "11.5px" } }, strip(proposal.baseSlides[index]?.title || "")),
      h("div", { class: "change-thumbs single", onclick: () => openCompare(before, null, `${index + 1}枚目を削除`) }, before())));
  }
  if (proposal.moved) rows.push(h("div", { class: "change-row" }, h("label", {}, toggle("move"), "スライドの順番を入れ替え"), h("span")));
  if (proposal.deckTitle) rows.push(h("div", { class: "change-row" }, h("label", {}, toggle("title"), "資料名"), h("span", { class: "hint", style: { "font-size": "11.5px" } }, `${proposal.baseTitle} → ${proposal.deckTitle}`)));
  if (proposal.transition) rows.push(h("div", { class: "change-row" }, h("label", {}, toggle("transition"), "スライドの切り替え"), h("span", { class: "hint", style: { "font-size": "11.5px" } }, `${TRANSITION_LABEL[proposal.baseDeck.transition] ?? "フェード"} → ${TRANSITION_LABEL[proposal.transition]}`)));
  if (proposal.motion) {
    const motionDeck = { ...designDeck, motion: normalizeMotion({ ...(proposal.baseDeck.motion || DEFAULT_MOTION), ...proposal.motion }) };
    rows.push(h("div", { class: "change-row" }, h("label", {}, toggle("motion"), "モーショングラフィック（資料全体）"), h("span", { class: "hint", style: { "font-size": "11.5px" } }, motionChangeText(proposal.baseDeck.motion, proposal.motion)),
      h("div", { class: "change-thumbs", onclick: () => openCompare(() => slidePicture(proposal.baseSlides[0], 0, proposal.baseDeck), () => slidePicture(proposal.baseSlides[0], 0, motionDeck), "動きの変更（表紙）") },
        slidePicture(proposal.baseSlides[0], 0, proposal.baseDeck), h("span", { class: "arrow" }, "→"), slidePicture(proposal.baseSlides[0], 0, motionDeck))));
  }
  return h("div", { class: "proposal" },
    h("div", { class: "proposal-head" }, h("span", {}, `変更の提案：${meta.summary}`)),
    h("div", { class: "proposal-rows" }, rows),
    proposal.issues.some((issue) => issue.kind === "overflow") ? h("div", { class: "proposal-note" }, `⚠ ${proposal.issues.filter((issue) => issue.kind === "overflow").length}か所は長めの文があります（採用後に赤枠で確認できます）`) : null,
    ...proposal.issues.filter((issue) => issue.kind === "visual").map((issue) => h("div", { class: "proposal-note" }, `⚠ ${issue.message}`)),
    unsourcedNote(proposal.issues) ? h("div", { class: "proposal-note" }, `⚠ ${unsourcedNote(proposal.issues)}`) : null,
    h("div", { class: "proposal-actions" },
      h("button", { class: "btn btn-primary", type: "button", onclick: () => applyProposal(message.id) }, proposal.whole ? "採用する" : "選んだ変更を採用"),
      h("button", { class: "btn", type: "button", onclick: () => rejectProposal(message.id) }, "却下"),
      h("span", { style: { flex: "1" } }),
      h("button", { class: "btn btn-ghost", type: "button", title: "この提案を元に、さらに直してもらう", onclick: () => { $("chatInput").value = "この提案を、"; $("chatInput").focus(); } }, "さらに直す")));
}

function openCompare(before, after, title) {
  $("compareTitle").textContent = title;
  const empty = (text) => h("div", { class: "empty-stage", style: { "min-height": "120px" } }, text);
  $("compareBefore").replaceChildren(before ? before() : empty("（なし）"));
  $("compareAfter").replaceChildren(after ? after() : empty("（削除）"));
  $("compareDialog").showModal();
}

/** The accepted slides: the proposal where selected, the original elsewhere. */
function proposalSlides(proposal) {
  if (proposal.whole) return clone(proposal.slides);
  const sel = proposal.selected;
  const groups = new Map();
  const seen = [];
  let anchor = 0;
  proposal.items.forEach((item, index) => {
    if (item.from == null) { groups.get(anchor)?.inserted.push(index); return; }
    anchor = item.from;
    groups.set(item.from, { slide: item.changed && sel.has(`c${index}`) ? proposal.slides[index] : proposal.baseSlides[item.from], inserted: [] });
    seen.push(item.from);
  });
  for (const index of proposal.deleted) {
    if (sel.has(`d${index}`)) continue;
    groups.set(index, { slide: proposal.baseSlides[index], inserted: [] });
    const after = seen.reduce((best, from, position) => (from < index && (best < 0 || from > seen[best]) ? position : best), -1);
    seen.splice(after + 1, 0, index);
  }
  if (!seen.length) return clone(proposal.baseSlides);
  const order = sel.has("move") ? seen : [...seen].sort((a, b) => a - b);
  return clone(order.flatMap((from) => [groups.get(from).slide, ...groups.get(from).inserted.filter((index) => sel.has(`n${index}`)).map((index) => proposal.slides[index])]));
}

function applyProposal(id) {
  const proposal = proposals.get(id);
  const message = state.chat.messages.find((item) => item.id === id);
  if (!proposal || !message) return;
  if (JSON.stringify(state.deck.slides) !== proposal.base && !state.confirmOverwrite) {
    state.confirmOverwrite = true;
    toast("提案の作成中に資料が編集されています。もう一度押すと、その編集は提案の内容で上書きされます");
    setTimeout(() => { state.confirmOverwrite = false; }, 6000);
    return;
  }
  state.confirmOverwrite = false;
  const slides = proposalSlides(proposal);
  if (slides.length < 2) return toast("スライドは2枚以上必要です");
  saveVersion("AIの提案を採用する前");
  proposal.before = JSON.stringify(state.deck);
  pushUndo();
  state.deck.slides = slides.map((slide, index) => normalizeSlide(slide, index, slides.length));
  if (proposal.deckTitle && proposal.selected.has("title")) { state.deck.title = proposal.deckTitle; $("deckTitleInput").value = proposal.deckTitle; }
  if (proposal.transition && proposal.selected.has("transition")) state.deck.transition = proposal.transition;
  if (proposal.motion && proposal.selected.has("motion")) state.deck.motion = normalizeMotion({ ...(state.deck.motion || DEFAULT_MOTION), ...proposal.motion });
  const first = proposal.items.findIndex((item, index) => item.changed && (proposal.selected.has(`c${index}`) || proposal.selected.has(`n${index}`)));
  if (first >= 0) state.selected = Math.min(first, slides.length - 1);
  state.selected = Math.min(state.selected, slides.length - 1);
  state.imported = null;
  message.proposal.status = "applied";
  saveChat();
  markChanged({ structural: true });
  renderChat();
  toast("提案を採用しました（「元に戻す」で取り消せます）");
}

function rejectProposal(id) {
  const message = state.chat.messages.find((item) => item.id === id);
  if (!message?.proposal) return;
  message.proposal.status = "rejected";
  saveChat();
  renderChat();
}

function revertProposal(id) {
  const proposal = proposals.get(id);
  const message = state.chat.messages.find((item) => item.id === id);
  if (!proposal?.before || !message) return;
  saveVersion("提案の取り消し前");
  pushUndo();
  state.deck = adoptDeck(JSON.parse(proposal.before));
  state.selected = Math.min(state.selected, state.deck.slides.length - 1);
  $("deckTitleInput").value = state.deck.title;
  message.proposal.status = "reverted";
  saveChat();
  useFonts([state.deck.theme]);
  markChanged({ structural: true });
  renderChat();
  toast("提案を採用する前に戻しました");
}

/** Whole-deck rewrites (dialog) arrive as a proposal too. */
function proposeWholeDeck(reply, nextSlides, baseDeck) {
  const same = nextSlides.length === baseDeck.slides.length;
  const items = nextSlides.map((slide, index) => ({ from: same ? index : null, changed: !same || JSON.stringify(slide) !== JSON.stringify(baseDeck.slides[index]) }));
  const changed = items.filter((item) => item.changed).length;
  const message = pushChat({ role: "assistant", text: reply, proposal: { status: "pending", summary: same ? `${changed}枚を修正` : `全体を${nextSlides.length}枚に再構成` } });
  createProposal(message.id, { slides: nextSlides, items, deleted: [], moved: false }, JSON.stringify(baseDeck.slides), baseDeck, { whole: !same });
  setPanel("chat");
}

// ---- three alternatives for one slide

async function requestVariants(index, instruction = "", shown = null) {
  if (state.chat.busy || !state.deck) return;
  if (state.panel !== "chat") setPanel("chat");
  pushChat({ role: "user", text: shown || `@${index + 1} の別案を3つ見せて${instruction ? `（${instruction}）` : ""}` });
  if (!state.codexAuthorized) {
    pushChat({ role: "assistant", text: "別案づくりはCodexに接続すると使えます。", error: true });
    return;
  }
  const baseJson = JSON.stringify(state.deck.slides[index]);
  state.chat.busy = true;
  state.chat.progress = "3つの案を考えています…";
  setAiBusy(true);
  renderChat();
  const done = () => { state.chat.busy = false; state.chat.progress = ""; setAiBusy(false); };
  try {
    const { jobId } = await jsonFetch("/api/decks/variants", { method: "POST", body: JSON.stringify({ deck: serverDeck(), slideIndex: index, instruction }) });
    watchJob(jobId, {
      onProgress: (job) => { state.chat.progress = `${job.stage}：${job.detail}`; const el = document.getElementById("chatProgress"); if (el) el.textContent = state.chat.progress; },
      onDone: (job) => {
        done();
        const previous = JSON.parse(baseJson);
        const message = pushChat({ role: "assistant", text: `${index + 1}枚目の案を${job.variants.length}つ用意しました。よいものを選ぶと、そのスライドに反映します。`, variants: { status: "pending", index, labels: job.variants.map((variant) => variant.label) } });
        variantSets.set(message.id, { index, baseJson, variants: job.variants.map((variant) => ({ ...variant, slide: restoreImages([normalizeSlide(variant.slide, index, state.deck.slides.length)], [previous])[0] })) });
        renderChat();
      },
      onFail: (job) => { done(); pushChat({ role: "assistant", text: `案を作れませんでした：${job.error || job.detail}`, error: true }); },
    });
  } catch (error) {
    done();
    pushChat({ role: "assistant", text: `送信できませんでした：${error.message}`, error: true });
  }
}

function variantsCard(message) {
  const meta = message.variants;
  const set = variantSets.get(message.id);
  if (!set || meta.status !== "pending") {
    const text = meta.status === "chosen" ? `✓「${meta.chosen}」を採用しました` : meta.status === "pending" ? "この案は期限切れです（再読み込みしたため）" : "却下しました";
    return h("div", { class: "proposal" }, h("div", { class: "proposal-head" }, h("span", {}, `${meta.index + 1}枚目の別案：${meta.labels.join("／")}`)), h("div", { class: `proposal-status ${meta.status === "chosen" ? "applied" : ""}` }, text));
  }
  const before = () => slidePicture(JSON.parse(set.baseJson), set.index);
  return h("div", { class: "proposal" },
    h("div", { class: "proposal-head" }, h("span", {}, `${set.index + 1}枚目の別案`)),
    h("div", { class: "variant-grid" }, set.variants.map((variant) => h("div", { class: "variant" },
      h("div", { class: "variant-thumb", title: "クリックで拡大", onclick: () => openCompare(before, () => slidePicture(variant.slide, set.index), `${set.index + 1}枚目：${variant.label}`) }, slidePicture(variant.slide, set.index)),
      h("div", { class: "variant-label" }, h("b", {}, variant.label),
        variant.issues?.some((issue) => issue.kind === "overflow") ? h("span", { class: "hint", style: { color: "var(--warn)" } }, "長めの文あり") : null,
        variant.unsourced?.length ? h("span", { class: "hint", style: { color: "var(--warn)" }, title: unsourcedNote(variant.unsourced) }, `素材にない数値${variant.unsourced.length}`) : null),
      h("button", { class: "btn btn-primary", type: "button", onclick: () => chooseVariant(message.id, variant) }, "この案にする")))),
    h("div", { class: "proposal-actions" },
      h("button", { class: "btn", type: "button", onclick: () => { meta.status = "rejected"; saveChat(); renderChat(); } }, "どれも使わない"),
      h("span", { style: { flex: "1" } }),
      h("button", { class: "btn btn-ghost", type: "button", onclick: () => requestVariants(set.index, "さっきとは違う切り口で") }, "別の3案")));
}

function chooseVariant(id, variant) {
  const set = variantSets.get(id);
  const message = state.chat.messages.find((item) => item.id === id);
  if (!set || !message) return;
  saveVersion("別案を採用する前");
  const original = state.deck.slides[set.index];
  const chosen = clone(variant.slide);
  for (const key of ["elements", "timeline", "sid", "transitionDur", "transitionSound", "transitionDir", "advance", "background"]) if (original?.[key] != null && chosen[key] == null) chosen[key] = clone(original[key]);
  replaceSlide(set.index, chosen);
  message.variants.status = "chosen";
  message.variants.chosen = variant.label;
  saveChat();
  renderChat();
  toast(`「${variant.label}」を採用しました（⌘Zで元に戻せます）`);
}

// ---------------------------------------------------------------- voice input

const Recognition = window.SpeechRecognition || window.webkitSpeechRecognition;

function voiceButton(target) {
  if (!Recognition) return null;
  const button = h("button", { class: "btn btn-ghost voice-btn", type: "button", title: "音声で入力（もう一度押すと停止）", "aria-label": "音声で入力" }, "🎤");
  let recognition = null;
  button.addEventListener("click", () => {
    if (recognition) { recognition.stop(); return; }
    const input = $(target);
    const prefix = input.value ? `${input.value.replace(/\s+$/, "")}${/[。、\n]$/.test(input.value) ? "" : " "}` : "";
    recognition = new Recognition();
    recognition.lang = "ja-JP";
    recognition.interimResults = true;
    recognition.continuous = true;
    recognition.onresult = (event) => {
      const text = [...event.results].map((result) => result[0].transcript).join("");
      input.value = `${prefix}${text}`;
      input.dispatchEvent(new Event("input", { bubbles: true }));
    };
    recognition.onerror = (event) => { if (event.error !== "no-speech" && event.error !== "aborted") toast(event.error === "not-allowed" ? "マイクの使用が許可されていません" : "音声を認識できませんでした"); };
    recognition.onend = () => { recognition = null; button.classList.remove("recording"); button.textContent = "🎤"; };
    recognition.start();
    button.classList.add("recording");
    button.textContent = "● 停止";
  });
  return button;
}

// ---------------------------------------------------------------- review round trip (comments → AI)

function mapReviewSlides(comments, reviewedSlides) {
  const slides = state.deck?.slides ?? [];
  const byTitle = new Map(slides.map((slide, index) => [squash(slide.title || slide.message || ""), index]));
  return comments.map((comment) => {
    let index = comment.slide;
    if (reviewedSlides && reviewedSlides.length !== slides.length) {
      const reviewed = reviewedSlides[comment.slide];
      index = byTitle.get(squash(reviewed?.title || reviewed?.message || "")) ?? null;
    }
    return { ...comment, slide: Number.isInteger(index) && index < slides.length ? index : null };
  });
}

function postReview(comments, source) {
  if (!comments.length) { toast(`${source} にコメントは見つかりませんでした`); return; }
  setPanel("chat");
  pushChat({ role: "assistant", text: `${source} から${comments.length}件のレビューコメントを読み込みました。反映したいものを選んで、AIに頼めます。`, review: { status: "pending", comments } });
}

async function importReview(file) {
  if (!file || !state.deck) return;
  const name = file.name.toLowerCase();
  try {
    if (name.endsWith(".json")) {
      const data = JSON.parse(await file.text());
      if (data?.kind !== "html-slide-review") throw new Error("レビュー用ファイルから書き出したコメントではありません");
      const reviewer = strip(data.reviewer || "");
      const comments = (data.comments || []).filter((item) => strip(item.text)).map((item) => ({ slide: Number.isInteger(item.slide) ? item.slide : null, author: reviewer, text: strip(item.text).slice(0, 600) }));
      postReview(mapReviewSlides(comments, data.slideTitles?.map((title) => ({ title }))), file.name);
      return;
    }
    if (!name.endsWith(".pptx")) throw new Error("コメント付きのPowerPoint（.pptx）か、レビュー結果（.json）を選んでください");
    toast(`${file.name} のコメントを読み込んでいます…`);
    const response = await fetch(`/api/import?name=${encodeURIComponent(file.name)}`, { method: "POST", credentials: "same-origin", headers: { "content-type": "application/octet-stream" }, body: file });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || `HTTP ${response.status}`);
    postReview(mapReviewSlides(result.comments || [], result.slideData), file.name);
  } catch (error) {
    toast(`レビューを読み込めません：${error.message}`);
  }
}

function reviewCard(message) {
  const meta = message.review;
  const comments = meta.comments || [];
  meta.selected ??= comments.map((_, index) => index);
  const selected = new Set(meta.selected);
  const done = meta.status !== "pending";
  return h("div", { class: "proposal" },
    h("div", { class: "proposal-head" }, h("span", {}, `レビューコメント ${comments.length}件`)),
    h("div", { class: "proposal-rows" }, comments.map((comment, index) => h("div", { class: "change-row review-row" },
      h("label", {},
        h("input", { type: "checkbox", checked: selected.has(index), disabled: done, onchange: (event) => { meta.selected = event.target.checked ? [...selected, index] : [...selected].filter((i) => i !== index); saveChat(); } }),
        h("span", { class: "kind" }, comment.slide == null ? "全体" : `${comment.slide + 1}枚目`),
        comment.author ? h("span", { class: "hint" }, comment.author) : null),
      h("span", { class: "review-text", onclick: () => { if (comment.slide != null) select(comment.slide); } }, comment.text)))),
    done
      ? h("div", { class: "proposal-status applied" }, "✓ AIに反映を依頼しました")
      : h("div", { class: "proposal-actions" },
        h("button", { class: "btn btn-ai", type: "button", disabled: !state.codexAuthorized, onclick: () => {
          const picked = comments.filter((_, index) => selected.has(index));
          if (!picked.length) return toast("反映するコメントを選んでください");
          meta.status = "sent";
          saveChat();
          const lines = picked.map((comment) => `- ${comment.slide == null ? "全体" : `@${comment.slide + 1}`}${comment.author ? `（${comment.author}）` : ""}：${comment.text}`);
          sendChat(`次のレビューコメントを反映してください。対応しないほうがよいものは理由を教えてください。\n${lines.join("\n")}`);
        } }, "✦ 選んだコメントを反映して"),
        h("button", { class: "btn", type: "button", title: "選んだコメントを各スライドのコメント（校閲）として残す。書いた人の名前も付きます", onclick: () => {
          const picked = comments.filter((_, index) => selected.has(index));
          if (!picked.length) return toast("残すコメントを選んでください");
          reviewToComments(picked);
        } }, "スライドのコメントにする"),
        h("button", { class: "btn", type: "button", onclick: () => { meta.status = "dismissed"; saveChat(); renderChat(); } }, "閉じる")));
}

/** Review comments from a file become the slides' own comments (校閲), with the reviewer's name: one undo step. */
function reviewToComments(picked) {
  if (!state.deck) return;
  pushUndo();
  const at = new Date().toISOString();
  picked.forEach((comment, k) => {
    const index = comment.slide ?? 0;
    const slide = state.deck.slides[index];
    if (!slide) return;
    const text = `${comment.slide == null ? "（資料全体）" : ""}${comment.text}`;
    if ((slide.comments || []).some((c) => c.text === text && (c.by || "") === (comment.author || ""))) return;
    slide.comments = [...(slide.comments || []), normalizeComment({ id: `rv${Date.now().toString(36)}${k}`, text, at, by: comment.author || "レビュー" }, `rv${k}`, true)].slice(0, 200);
  });
  markChanged({ structural: true });
  setPanel("comment", { reveal: true });
  toast(`${picked.length}件をスライドのコメントにしました（校閲 → コメント）`);
}

// ---------------------------------------------------------------- attachments in the chat

async function extractFileText(file) {
  if (/\.(pptx|docx|pdf)$/i.test(file.name)) {
    const response = await fetch(`/api/extract?name=${encodeURIComponent(file.name)}`, { method: "POST", credentials: "same-origin", headers: { "content-type": "application/octet-stream" }, body: file });
    const payload = await response.json();
    if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
    return payload.text || "";
  }
  return file.text();
}

async function attachToChat(file) {
  if (!file) return;
  if (file.size > 15_000_000) return toast("ファイルが大きすぎます（15MBまで）");
  try {
    if (/\.(pptx|docx|pdf)$/i.test(file.name)) toast(`${file.name} から文章を取り出しています…`);
    const text = (await extractFileText(file)).trim();
    if (!text) throw new Error("文章が見つかりませんでした");
    state.chat.attachment = { name: file.name, text: text.slice(0, 60_000) };
    renderChatContext();
    $("chatInput").focus();
    if (text.length > 60_000) toast("長い資料のため、先頭の約6万字を使います");
  } catch (error) {
    toast(`添付できません：${error.message}`);
  }
}

// ---------------------------------------------------------------- command palette (⌘K)

function commandList() {
  const deck = state.deck;
  const editing = state.mode === "edit" && deck;
  const cmd = (group, icon, label, run, desc = "") => ({ group, icon, label, run, desc });
  const list = [];
  if (editing) {
    list.push(
      cmd("資料", "▶", "最初から発表する", () => openPresenter(0), "F5"),
      cmd("資料", "▶", "このスライドから発表する", () => openPresenter(state.selected), "⇧F5"),
      cmd("資料", "⤓", "PowerPoint に書き出す（.pptx・編集できる）", () => exportPptx()),
      cmd("資料", "⤓", "HTMLファイルに書き出す（動きごと）", () => exportHtml()),
      cmd("資料", "⎙", "PDFとして保存・印刷", () => printPdf()),
      cmd("資料", "✓", "チェック結果を見る", () => openCheckDialog(false)),
      cmd("資料", "⌕", "検索・置換", () => openReplace()),
      cmd("編集", "↻", "直前の操作を繰り返す", () => editor.repeat(), "F4 / ⌘Y（やり直す操作がないとき）"),
      cmd("資料", "▣", "資料を保存庫に保存", () => saveToLibrary()),
      cmd("資料", "⌕", "保存庫から資料・スライドを探す", () => openLibrary()),
      cmd("資料", "🗒", "スピーカーノートを作る", () => openNotesDialog()),
      cmd("資料", "🕘", "版の履歴・過去の資料", () => openHistory()),
      cmd("資料", "{}", "JSONで保存", () => saveJsonFile()),
      cmd("デザイン", "◐", "動き（切り替え・登場・背景）", () => openDesignDialog()),
      cmd("動き", "✕", "全ページの動き・クリック動作を削除", () => resetAllActions(), "1回の元に戻すで復元可能"),
      cmd("AI", "✦", "全ページのHTML演出案をAIに相談", () => prepareActionPlanPrompt(), "送信前に依頼文を編集できます"),
      ...Object.entries(E.BACKDROPS).map(([kind, label]) => cmd("動き", "◎", `背景の動き：${label}`, () => { setDeckDesign({ motion: { backdrop: kind } }); toast(`表紙・章扉などの背景を「${label}」にしました`); }, "モーショングラフィック（資料全体）")),
      ...Object.entries(E.KINETIC).map(([style, label]) => cmd("動き", "◎", `文字の動き：${label.replace(/（.*）/, "")}`, () => { setDeckDesign({ motion: { kinetic: style } }); toast(`大きな文字の動きを「${label.replace(/（.*）/, "")}」にしました`); }, "モーショングラフィック（資料全体）")),
      ...Object.entries(E.ENTRANCES).map(([style, label]) => cmd("動き", "◎", `登場のしかた：${label}`, () => { setDeckDesign({ motion: { entrance: style } }); toast(`登場のしかたを「${label}」にしました`); }, "資料全体")),
      ...Object.entries(E.EMPHASES).map(([style, label]) => cmd("動き", "◎", `強調の見せ方：${label}`, () => { setDeckDesign({ motion: { emphasis: style } }); toast(`強調の見せ方を「${label}」にしました`); }, "**語句** の目立たせ方（資料全体）")),
      ...Object.entries(E.TRANSITIONS).map(([kind, label]) => cmd("動き", "◎", `切り替え：${label.replace(/（.*）/, "")}`, () => { setDeckDesign({ transition: kind }); toast(`スライドの切り替えを「${label.replace(/（.*）/, "")}」にしました`); }, "資料全体")),
      cmd("スライド", "▶", "このスライドの動きを確認", () => previewMotion()),
      cmd("スライド", "＋", "スライドを追加", () => openTypeDialog("insert")),
      cmd("スライド", "⇄", "このスライドのレイアウトを変更", () => openTypeDialog("change")),
      cmd("スライド", "⧉", "このスライドを複製", () => insertSlide(state.selected + 1, copyOf(deck.slides[state.selected]))),
      cmd("スライド", "✎", "このスライドを編集欄で開く", () => setPanel("form")),
      cmd("レビュー", "✉", "レビュー用ファイルを書き出す（相手はブラウザでコメント）", () => exportReviewFile()),
      cmd("レビュー", "⇩", "レビューを取り込む（コメント付きPPTX・レビュー結果）", () => $("reviewImportFile").click()),
      cmd("AI", "✦", "AIと話す", () => { setPanel("chat"); $("chatInput").focus(); }),
      cmd("AI", "✦", "資料全体をAIで見直す", () => openRewriteDialog()),
      cmd("AI", "✦", "このスライドの別案を3つ見る", () => requestVariants(state.selected)),
      cmd("AI", "📌", "前提条件（AIへの共通の指示）を書く", () => { setPanel("chat"); $("memoBox").hidden = false; $("memoInput").value = deck.memo || ""; $("memoInput").focus(); }),
      ...deck.slides.map((slide, index) => cmd("移動", `${index + 1}`, strip(slide.title || slide.message || typeLabel(slide.type)), () => select(index), typeLabel(slide.type))),
      // Every command of the ribbon, whichever tab it is on (shown once something is typed).
      ...(state.ribbonCmds || []).map((c) => cmd("リボン", "▸", c.label, () => editorUi.runRibbonCommand(c.ref), `${[c.tabLabel, c.group].filter(Boolean).join(" › ")}${c.disabled ? "（今は使えません）" : ""}`)),
    );
  }
  list.push(
    cmd("はじめる", "⌕", "保存庫から資料・スライドを探す", () => openLibrary()),
    cmd("はじめる", "✚", "新しい資料を作る", () => { setMode("create"); $("briefInput").focus(); }),
    cmd("はじめる", "⇪", "既存の資料を取り込む（PowerPoint・PDF・Word）", () => $("importDeckFile").click()),
    cmd("はじめる", "▦", "雛形から作る", () => openTemplateDialog()),
    cmd("はじめる", "★", "サンプル資料を開く", () => openSample()),
    cmd("はじめる", "?", "使い方とショートカット", () => $("helpDialog").showModal()),
  );
  return list;
}

function openCommandPalette() {
  const dialog = $("commandDialog");
  if (dialog.open) return;
  if (state.inline) finishInlineEdit(true);
  $("commandInput").value = "";
  state.commandActive = 0;
  state.ribbonCmds = state.mode === "edit" && state.deck ? editorUi.ribbonCommands() : [];
  renderCommands();
  dialog.showModal();
  $("commandInput").focus();
}

function filteredCommands() {
  const query = $("commandInput").value.trim();
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  const matches = commandList().filter((item) => (words.length || item.group !== "リボン") && words.every((word) => `${item.label} ${item.desc} ${item.group}`.toLowerCase().includes(word)));
  const canAsk = query && state.deck && state.mode === "edit" && state.codexAuthorized;
  if (!canAsk) return matches;
  const ask = { group: "AIに頼む", icon: "✦", label: `AIに頼む：「${query.slice(0, 40)}」`, run: () => sendChat(query), desc: "AIへの依頼として送る" };
  return matches.length ? [...matches, ask] : [ask];
}

function renderCommands() {
  const items = filteredCommands().slice(0, 80);
  state.commandActive = Math.min(state.commandActive, Math.max(0, items.length - 1));
  const nodes = [];
  let group = "";
  items.forEach((item, index) => {
    if (item.group !== group) { group = item.group; nodes.push(h("div", { class: "command-group" }, group)); }
    nodes.push(h("button", { class: `command-item${index === state.commandActive ? " active" : ""}`, type: "button", role: "option", onmousemove: () => { if (state.commandActive !== index) { state.commandActive = index; renderCommands(); } }, onclick: () => runCommand(item) },
      h("span", { class: "icon" }, item.icon), h("span", {}, item.label), item.desc ? h("span", { class: "desc" }, item.desc) : null));
  });
  if (!items.length) nodes.push(h("div", { class: "command-group" }, "見つかりません"));
  $("commandList").replaceChildren(...nodes);
  $("commandList").querySelector(".command-item.active")?.scrollIntoView({ block: "nearest" });
}

function runCommand(item) {
  $("commandDialog").close();
  item.run();
}

// ---------------------------------------------------------------- checks before export

function validateForExport() {
  const slides = state.deck.slides;
  const problems = [];
  // A deck brought over from PowerPoint as it looked keeps its own first and last pages (白紙 pages).
  if (slides[0]?.type !== "title" && slides[0]?.type !== "blank") problems.push("1枚目を表紙にしてください");
  if (slides.at(-1)?.type !== "closing" && slides.at(-1)?.type !== "blank") problems.push("最後をクロージングにしてください");
  return problems;
}

function checkItems() {
  const items = [];
  state.deck.slides.forEach((_, index) => {
    for (const issue of overflowFor(index)) items.push({ index, severity: "error", kind: "overflow", field: issue.field, message: `文字あふれ：${issue.message}`, issue });
  });
  lintDeck().flat().forEach((issue) => items.push({ index: issue.slide, severity: issue.severity, kind: "lint", field: issue.field, message: issue.message, fix: issue.fix, ai: issue.ai }));
  const order = { error: 0, warning: 1, info: 2 };
  return items.sort((a, b) => order[a.severity] - order[b.severity] || a.index - b.index);
}

async function openCheckDialog(forExport) {
  if (!state.deck) return;
  await measureAll();
  const problems = validateForExport();
  const items = checkItems();
  const canAi = state.codexAuthorized;
  $("exportCheckTitle").textContent = forExport ? "出力前のチェック" : "チェック結果";
  $("exportCheckLead").textContent = problems.length
    ? `出力できません：${problems.join("／")}`
    : items.length ? `${items.length}件の注意点があります。各項目から該当スライドへ移動したり、AIに直してもらったりできます。` : "問題は見つかりませんでした。";
  $("exportCheckList").replaceChildren(...items.map((item) => h("div", { class: "check-item" },
    h("span", { class: `badge ${item.severity}` }, `${item.index + 1}枚目`),
    h("span", { class: "msg-text" }, item.message),
    h("button", { class: "btn btn-ghost", type: "button", onclick: () => { $("exportCheckDialog").close(); setView("single"); select(item.index); if (item.field) setTimeout(() => focusField(item.field), 50); } }, "移動"),
    item.fix ? h("button", { class: "btn", type: "button", onclick: () => { item.fix(); openCheckDialog(forExport); } }, "直す") : null,
    !item.fix && canAi && (item.severity !== "info" || item.ai) ? h("button", { class: "btn", type: "button", onclick: () => {
      $("exportCheckDialog").close();
      select(item.index);
      if (item.kind === "overflow") reviseSlide(item.index, "スライドに収まらない文字を、意味を保ったまま短くする（補足は details に回してよい）", [item.issue]);
      else if (item.ai) sendChat(item.ai);
      else sendChat(`@${item.index + 1} の指摘「${item.message}」を直して`);
    } }, "✦ AIで直す") : null)));
  $("exportCheckGoBtn").hidden = !forExport;
  $("exportCheckGoBtn").disabled = problems.length > 0;
  $("exportCheckFixBtn").textContent = forExport ? "戻って直す" : "閉じる";
  if (!$("exportCheckDialog").open) $("exportCheckDialog").showModal();
}

// ---------------------------------------------------------------- presenting

async function openPresenter(start = state.selected, { custom = null, extra = null, onClosed = null, fullscreen = true } = {}) {
  if (!state.deck || state.player) return;
  // While presenting online, every slide show goes to the viewers too.
  if (online?.active && !extra && !onClosed) return online.present(start, { custom });
  if (onClosed) state.onShowClosed = onClosed;
  // スライド ショーの設定: which slides (a range or a custom show) and how the show runs.
  const settings = showOf(state.deck.show) || {};
  const only = showSlides(state.deck, settings, custom);
  if (only && !only.includes(start)) start = only[0];
  // Ink written in the show comes back to the slides it was written on, found by their ids.
  const sidsAtStart = state.deck.slides.map((slide) => slide.sid || null);
  if (state.inline) finishInlineEdit(true);
  stopMotionPreview({ render: false });
  const host = $("presenter");
  host.classList.remove("hidden");
  host.replaceChildren();
  if (fullscreen) host.requestFullscreen?.().catch(() => {});
  useFonts([state.deck.theme]);
  await measureAll();
  state.player = E.createPlayer(host, {
    deck: clone(state.deck),
    start: Math.max(0, Math.min(start, state.deck.slides.length - 1)),
    fitFor: (i) => fitFor(i) ?? undefined,
    renderOptions: { assetBase: "/assets/", mediaUrls },
    fullscreenTarget: host,
    ...playerOptions(settings),
    only,
    ...(extra || {}),
    onClose: ({ index, ink }) => closePresenter(index, { ink, sidsAtStart }),
  });
  if (presenterViewOn()) state.player.openPresenterView?.();
}

/**
 * スライド ショー →「リハーサル」: present from the start timing each slide (a clock in the corner); at the end, the
 * times can be kept as each slide's 自動で切り替え (slide.advance), as PowerPoint's rehearse timings.
 */
async function rehearse() {
  if (!state.deck || state.player) return;
  await openPresenter(0);
  if (!state.player) return;
  const times = new Map();
  let at = state.player.index;
  let since = Date.now();
  const started = since;
  const clock = h("div", { class: "rehearse-clock", role: "status" });
  $("presenter").append(clock);
  const fmt = (ms) => { const sec = Math.floor(ms / 1000); return `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, "0")}`; };
  const tick = () => {
    const player = state.player;
    if (!player) return;
    const now = Date.now();
    if (player.index !== at) { times.set(at, (times.get(at) || 0) + now - since); at = player.index; since = now; }
    clock.textContent = `リハーサル　このスライド ${fmt((times.get(at) || 0) + now - since)}　合計 ${fmt(now - started)}`;
  };
  state.rehearsal = { times, timer: setInterval(tick, 200), finish: () => { times.set(at, (times.get(at) || 0) + Date.now() - since); } };
  tick();
}
/**
 * スライド ショー →「記録」: present while the microphone records; each slide gets its own narration (the last take
 * of it) and its time. At the end both can be kept: the narration as a hidden sound that plays with the slide
 * (slide.elements, narration: true) and the time as the slide's 自動で切り替え.
 */
async function recordShow(start = 0) {
  if (!state.deck || state.player) return;
  if (!navigator.mediaDevices?.getUserMedia || typeof MediaRecorder === "undefined") { toast("このブラウザでは録音できません"); return; }
  let stream;
  try { stream = await navigator.mediaDevices.getUserMedia({ audio: true }); } catch { toast("マイクを使えません（ブラウザでマイクを許可してください）"); return; }
  await openPresenter(start);
  if (!state.player) { stream.getTracks().forEach((t) => t.stop()); return; }
  const takes = new Map();
  const times = new Map();
  let at = state.player.index;
  let since = Date.now();
  const started = since;
  let current = null;
  const begin = (index) => {
    const chunks = [];
    const rec = new MediaRecorder(stream);
    const done = new Promise((resolve) => { rec.onstop = () => { if (chunks.length) takes.set(index, new Blob(chunks, { type: rec.mimeType || "audio/webm" })); resolve(); }; });
    rec.ondataavailable = (event) => { if (event.data?.size) chunks.push(event.data); };
    rec.start();
    current = { rec, done };
  };
  const end = () => { const c = current; current = null; if (c && c.rec.state !== "inactive") c.rec.stop(); return c?.done || Promise.resolve(); };
  begin(at);
  const clock = h("div", { class: "rehearse-clock recording", role: "status" });
  $("presenter").append(clock);
  const fmt = (ms) => { const sec = Math.floor(ms / 1000); return `${Math.floor(sec / 60)}:${String(sec % 60).padStart(2, "0")}`; };
  const tick = () => {
    const player = state.player;
    if (!player) return;
    const now = Date.now();
    if (player.index !== at) { times.set(at, (times.get(at) || 0) + now - since); end(); at = player.index; since = now; begin(at); }
    clock.textContent = `● 記録中　このスライド ${fmt((times.get(at) || 0) + now - since)}　合計 ${fmt(now - started)}`;
  };
  state.recording = { timer: setInterval(tick, 200), finish: async () => { times.set(at, (times.get(at) || 0) + Date.now() - since); await end(); stream.getTracks().forEach((t) => t.stop()); return { times, takes }; } };
  tick();
}
async function keepRecording({ times, takes }) {
  if (!state.deck || (!takes.size && !times.size)) return;
  const sec = Math.round([...times.values()].reduce((a, b) => a + b, 0) / 1000);
  if (!window.confirm(`記録したナレーション（${takes.size}枚）とタイミング（合計 ${Math.floor(sec / 60)}分${sec % 60}秒）を保存しますか？\n保存すると、発表でスライドごとにナレーションが流れ、その時間で自動的に切り替わります。`)) return;
  const sids = new Map([...new Set([...takes.keys(), ...times.keys()])].map((i) => [i, state.deck.slides[i]?.sid || null]));
  const media = new Map();
  for (const [i, blob] of takes) if (blob.size > 800) media.set(i, await putMedia(blob, `ナレーション ${i + 1}.webm`));
  pushUndo();
  let count = 0;
  for (const [i, sid] of sids) {
    const at = sid ? state.deck.slides.findIndex((slide) => slide.sid === sid) : i;
    const slide = state.deck.slides[at];
    if (!slide) continue;
    const ms = times.get(i) || 0;
    if (ms >= 500) slide.advance = Math.min(600, Math.round(ms / 100) / 10);
    const src = media.get(i);
    if (!src) continue;
    // A new take replaces the slide's narration.
    const size = 96;
    const sound = { ...audioObject(src, "ナレーション"), x: E.W - size - 40, y: E.H - size - 150, w: size, h: size, narration: true, autoplay: true, hideIcon: true };
    slide.elements = [...(slide.elements || []).filter((o) => !o.narration), E.normalizeObject(sound)].filter(Boolean);
    count += 1;
  }
  markChanged({ structural: true });
  toast(`${count}枚のナレーションとタイミングを保存しました（スライド ショー →「ナレーションの再生」でオン・オフ。⌘Zで戻せます）`);
}
/** 記録 → クリア: the timings or the narration of this slide or of all slides. */
function clearRecording(what, scope) {
  if (!state.deck) return;
  const targets = scope === "all" ? state.deck.slides : [state.deck.slides[state.selected]].filter(Boolean);
  const has = targets.some((slide) => (what === "timings" ? slide.advance != null : (slide.elements || []).some((o) => o.narration)));
  if (!has) { toast(what === "timings" ? "消すタイミングはありません" : "消すナレーションはありません"); return; }
  pushUndo();
  for (const slide of targets) {
    if (what === "timings") delete slide.advance;
    else { slide.elements = (slide.elements || []).filter((o) => !o.narration); if (!slide.elements.length) delete slide.elements; }
  }
  markChanged({ structural: true });
  toast(`${scope === "all" ? "すべてのスライド" : "このスライド"}の${what === "timings" ? "タイミング" : "ナレーション"}を消しました（⌘Zで戻せます）`);
}

function keepRehearsal(times) {
  const total = [...times.values()].reduce((a, b) => a + b, 0);
  if (total < 1000 || !state.deck) return;
  const sec = Math.round(total / 1000);
  if (!window.confirm(`リハーサルの時間（合計 ${Math.floor(sec / 60)}分${sec % 60}秒）を保存して、スライドを自動で切り替えるようにしますか？`)) return;
  pushUndo();
  let count = 0;
  for (const [i, ms] of times) {
    const slide = state.deck.slides[i];
    if (!slide || ms < 500) continue;
    slide.advance = Math.min(600, Math.round(ms / 100) / 10);
    count += 1;
  }
  markChanged({ structural: true });
  toast(`${count}枚のタイミングを保存しました（画面切り替えタブの「自動的に切り替え」で変えられます。⌘Zで戻せます）`);
}

function closePresenter(index, { ink = [], sidsAtStart = null } = {}) {
  const closed = state.onShowClosed;
  state.onShowClosed = null;
  closed?.(index);
  const rehearsal = state.rehearsal;
  if (rehearsal) { clearInterval(rehearsal.timer); rehearsal.finish(); state.rehearsal = null; }
  const recording = state.recording;
  if (recording) { clearInterval(recording.timer); state.recording = null; recording.finish().then((result) => setTimeout(() => keepRecording(result), 0)); }
  queueMicrotask(() => { if (rehearsal) keepRehearsal(rehearsal.times); if (ink?.length) keepShowInk(ink, sidsAtStart); });
  state.player = null;
  const host = $("presenter");
  host.classList.add("hidden");
  host.replaceChildren();
  if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {});
  if (Number.isInteger(index)) select(index);
}

/**
 * オンライン プレゼンテーション, the viewer's side (public/editor/online.mjs): someone else's deck fills the window
 * and follows its presenter. This browser's own deck is not touched.
 */
function openViewer(deck, { fits = null, start = 0, step = 0, label = null, onReact = null, onVote = null } = {}) {
  state.player?.destroy?.();
  document.body.classList.add("is-viewer");
  const host = $("presenter");
  host.classList.remove("hidden");
  host.replaceChildren();
  useFonts([deck.theme]);
  const first = Math.max(0, Math.min(Number(start) || 0, deck.slides.length - 1));
  state.player = E.createPlayer(host, {
    deck,
    start: first,
    fitFor: (i) => fits?.[i] || undefined,
    renderOptions: { assetBase: "/assets/", mediaUrls },
    fullscreenTarget: host,
    viewer: true,
    viewerLabel: label,
    closable: false,
    onReact,
    onVote,
  });
  if (step) state.player.follow(first, step);
  return state.player;
}
function openViewerMessage(text) {
  document.body.classList.add("is-viewer");
  const host = $("presenter");
  host.classList.remove("hidden");
  host.replaceChildren(h("div", { class: "op-message" }, h("p", {}, text), h("a", { class: "btn", href: location.pathname }, "スタジオを開く")));
}

/** PowerPoint asks, after a show with ink: keep the annotations? They become ink objects on their slides (one undo). */
function keepShowInk(ink, sidsAtStart) {
  if (!state.deck) return;
  const count = ink.reduce((sum, item) => sum + (item.strokes?.length || 0), 0);
  if (!count || !window.confirm(`発表中に書いたインク（${count}本）を保持しますか？\n「OK」でスライドに残し、「キャンセル」で破棄します。`)) return;
  const made = keptInk(ink, sidsAtStart, state.deck, (strokes) => inkObject(strokes, E.newObjectId()));
  if (!made.size) return;
  pushUndo();
  for (const [i, objects] of made) state.deck.slides[i].elements = [...(state.deck.slides[i].elements || []), ...objects];
  markChanged({ structural: true });
  toast(`${made.size}枚のスライドにインクを残しました（描画タブで消しゴム・図形に変換。⌘Zで戻せます）`);
}

/** ファイル → 印刷: the print dialog (layouts, range, preview); public/editor/print.mjs. */
async function printPdf() {
  if (!state.deck) return;
  await measureAll();
  printer.openDialog();
}

/** Every slide at full size straight to the browser's print dialog (kept for callers that want no dialog). */
async function printSlidesNow() {
  if (!state.deck) return;
  await measureAll();
  const root = $("printRoot");
  root.replaceChildren(...state.deck.slides.map((slide, i) => E.render(slide, renderOptions({ index: i, mode: "print", fit: fitFor(i) ?? undefined }))));
  await Promise.all([...root.querySelectorAll("img")].map((img) => img.decode?.().catch(() => {})));
  await E.mountLottie(root, { play: false, frame: 0.5 });
  await wait(120);
  await document.fonts?.ready;
  const cleanup = () => { E.stopLottie(root); root.replaceChildren(); window.removeEventListener("afterprint", cleanup); };
  window.addEventListener("afterprint", cleanup);
  toast("印刷画面で「PDFとして保存」を選び、余白を「なし」にしてください");
  setTimeout(() => window.print(), 80);
}

// ---------------------------------------------------------------- exporting (one self-contained HTML file)

async function downloadBlob(blob, fileName) {
  const link = h("a", { href: URL.createObjectURL(blob), download: fileName });
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 4000);
}

async function fetchText(url) {
  const response = await fetch(url, { credentials: "same-origin" });
  if (!response.ok) throw new Error(`${url} を読み込めません（${response.status}）`);
  return response.text();
}

async function engineBundle() {
  const [css, engine, objects, animate, motion, models] = await Promise.all([
    fetchText(`/engine/engine.css?v=${APP_VERSION}`),
    fetchText(`/engine/engine.js?v=${APP_VERSION}`),
    fetchText(`/engine/objects.js?v=${APP_VERSION}`),
    fetchText(`/engine/animate.js?v=${APP_VERSION}`),
    fetchText(`/engine/motion.js?v=${APP_VERSION}`),
    fetchText(`/engine/models.js?v=${APP_VERSION}`),
  ]);
  const js = (text) => text.replace(/<\/script/gi, "<\\/script");
  return { css: css.replace(/<\/style/gi, "<\\/style"), engine: js(engine), objects: js(objects), animate: js(animate), motion: js(motion), models: js(models) };
}

/**
 * three.js for an exported file with 3D models: its modules as text, turned into blob: modules when the file opens
 * (each import pointed at the blob made before it), so the file needs nothing from the network.
 */
async function threeBundle() {
  const names = { core: "three.core.js", module: "three.module.js", bgu: "BufferGeometryUtils.js", sku: "SkeletonUtils.js", gltf: "GLTFLoader.js" };
  const sources = Object.fromEntries(await Promise.all(Object.entries(names).map(async ([key, name]) => [key, await fetchText(`/vendor/three/${name}?v=${APP_VERSION}`)])));
  const boot = `(function () {
  var src = JSON.parse(document.getElementById("hs-three").textContent);
  var url = function (code) { return URL.createObjectURL(new Blob([code], { type: "text/javascript" })); };
  var swap = function (code, from, to) { return code.split("'" + from + "'").join("'" + to + "'").split('"' + from + '"').join('"' + to + '"'); };
  var core = url(src.core);
  var mod = url(swap(src.module, "./three.core.js", core));
  var bgu = url(swap(src.bgu, "./three.module.js", mod));
  var sku = url(swap(src.sku, "./three.module.js", mod));
  var gltf = url(swap(swap(swap(src.gltf, "./three.module.js", mod), "./BufferGeometryUtils.js", bgu), "./SkeletonUtils.js", sku));
  window.SlideEngine.threeUrls = { three: mod, gltf: gltf };
})();`;
  return { data: safeJson(sources), boot };
}

/** The deck with every photo and video it uses written into it, so the file works anywhere. */
async function portableDeck() {
  const deck = clone(state.deck);
  const assets = {};
  let bytes = 0;
  const addAsset = async (key) => {
    const file = E.PHOTOS[key]?.[0];
    if (!file || assets[file]) return;
    const blob = await (await fetch(`/assets/${file}`)).blob();
    assets[file] = await blobToDataUrl(blob);
    bytes += blob.size;
  };
  const missing = [];
  // スライド マスター: its pictures travel too (the master is the last "slide" of the loop).
  const master = deck.masterObjects?.length ? { elements: deck.masterObjects } : null;
  // 配布資料マスター・ノート マスター: the pictures of their pages travel too.
  const printed = Object.keys(deck.printMasters || {}).filter((kind) => deck.printMasters[kind]?.objects?.length).map((kind) => ({ kind, elements: deck.printMasters[kind].objects }));
  for (const [index, slide] of [...deck.slides, ...(master ? [master] : []), ...printed].entries()) {
    if (slide.visualAsset && E.PHOTOS[slide.visualAsset]) await addAsset(slide.visualAsset);
    const src = slide.media?.src;
    if (src?.startsWith("asset:")) await addAsset(src.slice(6));
    if (src?.startsWith("idb:")) {
      const blob = await mediaBlob(src).catch(() => null);
      if (!blob) { missing.push(index + 1); delete slide.media; }
      else { slide.media.src = await blobToDataUrl(blob); bytes += blob.size; }
    }
    // 背景の書式設定: the background picture travels inside the file too.
    const bgSrc = slide.background?.image;
    if (bgSrc?.startsWith("asset:")) await addAsset(bgSrc.slice(6));
    if (bgSrc?.startsWith("idb:")) {
      const blob = await mediaBlob(bgSrc).catch(() => null);
      if (!blob) { if (!missing.includes(index + 1)) missing.push(index + 1); delete slide.background.image; if (!slide.background.color) delete slide.background; }
      else { slide.background.image = await blobToDataUrl(blob); bytes += blob.size; }
    }
    // Pictures, videos and animations placed as objects travel inside the file too.
    if (slide.elements?.length) {
      const kept = [];
      for (const o of slide.elements) {
        if (o.src?.startsWith("asset:")) await addAsset(o.src.slice(6));
        if (o.src?.startsWith("idb:")) {
          const blob = await mediaBlob(o.src).catch(() => null);
          if (!blob) { if (!missing.includes(index + 1)) missing.push(index + 1); continue; }
          o.src = await blobToDataUrl(blob);
          bytes += blob.size;
        }
        // 図で塗りつぶし: the shape's picture too (without it the shape keeps its colour).
        if (o.fillImg?.startsWith("asset:")) await addAsset(o.fillImg.slice(6));
        if (o.fillImg?.startsWith("idb:")) {
          const blob = await mediaBlob(o.fillImg).catch(() => null);
          if (blob) { o.fillImg = await blobToDataUrl(blob); bytes += blob.size; } else { delete o.fillImg; delete o.fillTile; if (!missing.includes(index + 1)) missing.push(index + 1); }
        }
        kept.push(o);
      }
      slide.elements = kept;
    }
  }
  if (master) { if (master.elements?.length) deck.masterObjects = master.elements; else delete deck.masterObjects; }
  for (const { kind, elements } of printed) { if (elements.length) deck.printMasters[kind].objects = elements; else delete deck.printMasters[kind].objects; }
  return { deck, assets, bytes, missing };
}

async function standaloneHtml({ title, body, boot, data, background = "#07080c", extraCss = "", player = false }) {
  const bundle = await engineBundle();
  const fonts = E.fontHref([data.deck.theme]);
  // The Lottie player travels with the file only when a slide has an animation to play.
  const lottie = player && withMaster(data.deck).some((slide) => slide.media?.kind === "lottie" || (slide.elements || []).some((o) => o.kind === "lottie")) ? (await fetchText("/vendor/lottie.js")).replace(/<\/script/gi, "<\\/script") : "";
  // So does three.js, only when a slide has a 3D model.
  const three = player && withMaster(data.deck).some((slide) => (slide.elements || []).some((o) => o.kind === "model")) ? await threeBundle() : null;
  return `<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="generator" content="HTML PPTX Studio ${APP_VERSION}">
${data.deck.info?.author ? `<meta name="author" content="${esc(data.deck.info.author)}">` : ""}${data.deck.info?.subject ? `<meta name="description" content="${esc(data.deck.info.subject)}">` : ""}${data.deck.info?.keywords ? `<meta name="keywords" content="${esc(data.deck.info.keywords)}">` : ""}
<title>${esc(title)}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
${fonts ? `<link rel="stylesheet" href="${esc(fonts)}" data-hs-fonts>` : ""}
<style data-hs-engine>${bundle.css}</style>
<style>html,body{margin:0;height:100%;background:${background}}body{position:relative}.hs-boot{position:fixed;inset:0;display:grid;place-items:center;color:#8b95a8;font:14px "Noto Sans JP",sans-serif}${extraCss}</style>
</head>
<body>
${body}
<script>${bundle.engine}</script>
<script>${bundle.objects}</script>
<script>${bundle.animate}</script>
<script>${bundle.motion}</script>
<script>${bundle.models}</script>
${three ? `<script type="application/json" id="hs-three">${three.data}</script>\n<script>${three.boot}</script>` : ""}
${lottie ? `<script>${lottie}</script>` : ""}
<script type="application/json" id="hs-data">${safeJson(data)}</script>
<script>${boot}</script>
</body>
</html>`;
}

const EXPORT_BOOT = `(async function () {
  var data = JSON.parse(document.getElementById("hs-data").textContent);
  var E = window.SlideEngine;
  for (var i = 0; i < data.deck.slides.length; i += 1) {
    var media = data.deck.slides[i].media;
    if (media && /^data:/.test(media.src)) { try { media.src = URL.createObjectURL(await (await fetch(media.src)).blob()); } catch (e) {} }
    var objects = data.deck.slides[i].elements || [];
    for (var j = 0; j < objects.length; j += 1) {
      if (objects[j].src && /^data:/.test(objects[j].src)) { try { objects[j].src = URL.createObjectURL(await (await fetch(objects[j].src)).blob()); } catch (e) {} }
    }
  }
  var boot = document.querySelector(".hs-boot"); if (boot) boot.remove();
  // "#5" opens slide 5; "#static" (or "#5-static") shows every slide finished, without motion.
  var still = /static/.test(location.hash);
  var start = Math.max(0, (parseInt((location.hash.match(/\\d+/) || ["1"])[0], 10) || 1) - 1);
  // スライド ショーの設定 (a kiosk, looping, the slides it plays, timings, the pen colour) travel with the file.
  var opts = {
    deck: data.deck, start: start, closable: false,
    fitFor: function (i) { return data.fits[i] || undefined; },
    renderOptions: { assetMap: data.assets, assetBase: "" },
    onChange: function (s) { try { history.replaceState(null, "", "#" + (s.index + 1) + (still ? "-static" : "")); } catch (e) {} }
  };
  var show = data.show || {};
  for (var key in show) if (Object.prototype.hasOwnProperty.call(show, key) && show[key] != null) opts[key] = show[key];
  opts.static = still || Boolean(show.static);
  if (opts.only && opts.only.indexOf(start) < 0) opts.start = opts.only[0];
  E.createPlayer(document.body, opts);
})();`;

/**
 * ファイル → 情報 → プレゼンテーションの保護 → パスワードを使用して暗号化: the password asked twice, then the HTML file
 * written locked with it. The password is not kept (forgotten, the file cannot be opened).
 */
function exportLockedHtml() {
  if (!state.deck) return;
  const pw = h("input", { type: "password", class: "lock-pw", autocomplete: "new-password", "aria-label": "パスワード" });
  const again = h("input", { type: "password", class: "lock-again", autocomplete: "new-password", "aria-label": "パスワードの再入力" });
  const msg = h("p", { class: "hint lock-msg", role: "alert" });
  const dialog = h("dialog", { class: "lock-dialog", "aria-label": "パスワードを使用して暗号化" },
    h("div", { class: "dialog-head" }, h("h3", {}, "パスワードを使用して暗号化"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
    h("div", { class: "dialog-body sh-form-col" },
      h("p", {}, "書き出すHTMLファイルをパスワードで保護します。開くときにパスワードを聞き、合っていれば発表できます（中身は暗号化され、ファイルを見ても読めません）。"),
      h("label", { class: "field" }, h("span", {}, "パスワード"), pw),
      h("label", { class: "field" }, h("span", {}, "パスワードの再入力"), again),
      h("p", { class: "hint" }, "注意：パスワードを忘れると、ファイルを開けなくなります。パスワードはどこにも保存されません。大文字と小文字は区別されます。"),
      msg),
    h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
      h("button", { type: "button", class: "btn btn-primary lock-go", onclick: () => {
        if (pw.value.length < 4) { msg.textContent = "パスワードは4文字以上にしてください"; pw.focus(); return; }
        if (pw.value !== again.value) { msg.textContent = "2つのパスワードが違います"; again.select(); return; }
        const password = pw.value;
        dialog.close();
        exportHtml({ password });
      } }, "書き出す")));
  document.body.append(dialog);
  dialog.addEventListener("close", () => dialog.remove());
  dialog.showModal();
  pw.focus();
}

async function exportHtml({ checked = false, password = null } = {}) {
  if (!state.deck || state.exporting) return;
  if (!checked && !password) {
    await measureAll();
    const problems = validateForExport();
    const warnings = checkItems().filter((item) => item.severity !== "info");
    if (problems.length || warnings.length) { state.pendingExport = "html"; return openCheckDialog(true); }
  }
  state.exporting = true;
  const button = $("downloadBtn");
  button.textContent = "書き出し中…";
  updateTopbar();
  try {
    await measureAll();
    const { deck, assets, bytes, missing } = await portableDeck();
    if (deck.slides[0].type === "title" && !deck.slides[0].date) deck.slides[0].date = new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long", day: "numeric" });
    const fits = state.deck.slides.map((_, i) => { const fit = fitFor(i); return fit ? { fs: fit.fs, ts: fit.ts } : null; });
    const plain = await standaloneHtml({ title: deck.title, body: '<div class="hs-boot">読み込み中…</div>', boot: EXPORT_BOOT, data: { deck, fits, assets, show: { ...playerOptions(showOf(deck.show)), only: showSlides(deck, showOf(deck.show)) } }, player: true });
    // パスワードを使用して暗号化: the file opens only with its password (protect.mjs).
    // The locked page names nothing of the deck (its title may be confidential too).
    const html = password ? await lockHtml(plain, password) : plain;
    const fileName = `${fileSafe(deck.title, "presentation")}_${today()}${password ? "_保護" : ""}.html`;
    await downloadBlob(new Blob([html], { type: "text/html" }), fileName);
    state.historyId = saveHistory({ exported: fileName });
    const size = bytes > 20_000_000 ? `（約${Math.round(bytes / 1_000_000)}MB。動画が入っているため大きめです）` : "";
    toast(`HTMLに書き出しました${password ? "（パスワード付き。開くときにパスワードを聞きます）" : ""}${size}。ブラウザで開くとそのまま発表できます${missing.length ? `／${missing.join("・")}枚目の動画・写真はこのブラウザにないため外しました` : ""}`);
  } catch (error) {
    toast(`書き出せませんでした：${error.message}`);
  } finally {
    state.exporting = false;
    button.textContent = "HTML出力";
    updateTopbar();
  }
}

/**
 * PowerPoint出力: the deck (media inside, as the HTML export carries it) goes to the server, which draws it in a
 * headless browser and writes an editable .pptx (server/pptx.mjs, tools/pptx_export.py).
 */
async function exportPptx({ checked = false } = {}) {
  if (!state.deck || state.exporting) return;
  if (!checked) {
    await measureAll();
    if (validateForExport().length) { state.pendingExport = "pptx"; return openCheckDialog(true); }
  }
  state.exporting = true;
  const button = $("pptxBtn");
  button.textContent = "書き出し中…";
  updateTopbar();
  try {
    const { deck, missing } = await portableDeck();
    if (deck.slides[0].type === "title" && !deck.slides[0].date) deck.slides[0].date = new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long", day: "numeric" });
    const response = await fetch("/api/export/pptx", { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify({ deck }) });
    if (!response.ok) throw new Error((await response.json().catch(() => ({}))).error || `HTTP ${response.status}`);
    const fileName = `${fileSafe(deck.title, "presentation")}_${today()}.pptx`;
    await downloadBlob(await response.blob(), fileName);
    state.historyId = saveHistory({ exported: fileName });
    let stats = {};
    try { stats = JSON.parse(response.headers.get("x-export-stats") || "{}"); } catch { /* none */ }
    const pictures = stats.rasterized ? `（PowerPoint にない部品${stats.rasterized}か所は図として入れました）` : "";
    toast(`PowerPoint に書き出しました${pictures}${missing.length ? `。${missing.join("・")}枚目の写真・動画は見つからなかったので外しました` : ""}`);
  } catch (error) {
    toast(`PowerPoint に書き出せませんでした：${error.message}`);
  } finally {
    state.exporting = false;
    button.textContent = "PowerPoint出力";
    updateTopbar();
  }
}

const REVIEW_BOOT = `(function () {
  var data = JSON.parse(document.getElementById("hs-data").textContent);
  var E = window.SlideEngine;
  var list = document.getElementById("rvList");
  data.deck.slides.forEach(function (slide, i) {
    var el = E.render(slide, { deck: data.deck, index: i, mode: "thumb", fit: data.fits[i] || undefined, assetMap: data.assets, assetBase: "" });
    var section = document.createElement("section"); section.className = "rv-slide";
    var left = document.createElement("div"); left.appendChild(E.mount(el));
    var note = document.createElement("div"); note.className = "rv-note";
    var b = document.createElement("b"); b.textContent = (i + 1) + "枚目　" + E.strip(slide.title || slide.message || "");
    var area = document.createElement("textarea"); area.dataset.slide = String(i); area.placeholder = "例：数字の根拠を書いてほしい／結論を先に";
    note.appendChild(b); note.appendChild(area); section.appendChild(left); section.appendChild(note); list.appendChild(section);
  });
  var KEY = data.key, TITLES = data.titles, DECK = data.deck.title;
  var boxes = Array.prototype.slice.call(document.querySelectorAll("textarea[data-slide]"));
  var name = document.getElementById("rvName");
  try { var saved = JSON.parse(localStorage.getItem(KEY) || "{}"); name.value = saved.name || ""; boxes.forEach(function (b) { b.value = (saved.notes || {})[b.dataset.slide] || ""; }); } catch (e) {}
  function count() { var n = boxes.filter(function (b) { return b.value.trim(); }).length; document.getElementById("rvCount").textContent = n ? n + "件のコメント" : "まだコメントはありません"; }
  function save() { try { var notes = {}; boxes.forEach(function (b) { notes[b.dataset.slide] = b.value; }); localStorage.setItem(KEY, JSON.stringify({ name: name.value, notes: notes })); } catch (e) {} count(); }
  document.addEventListener("input", save); count();
  document.getElementById("rvExport").onclick = function () {
    var comments = boxes.filter(function (b) { return b.value.trim(); }).map(function (b) { return { slide: b.dataset.slide === "all" ? null : Number(b.dataset.slide), text: b.value.trim() }; });
    if (!comments.length) { document.getElementById("rvCount").textContent = "コメントがまだありません"; return; }
    var out = { kind: "html-slide-review", version: 1, deckTitle: DECK, reviewer: name.value.trim(), slideTitles: TITLES, comments: comments, exportedAt: new Date().toISOString() };
    var a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([JSON.stringify(out, null, 1)], { type: "application/json" }));
    a.download = "レビュー結果_" + DECK.replace(/[\\\\/:*?"<>|]/g, "_").slice(0, 40) + (name.value.trim() ? "_" + name.value.trim() : "") + ".json";
    document.body.appendChild(a); a.click(); a.remove();
    document.getElementById("rvCount").textContent = "書き出しました。できたファイルを作成者に送ってください。";
  };
})();`;

const REVIEW_CSS = `body{overflow:auto;font-family:"Noto Sans JP","Hiragino Sans",sans-serif;color:#14161c}.rv{max-width:1240px;margin:0 auto;padding:28px 16px 90px}.rv h1{font-size:22px;margin:0 0 6px}.rv-lead{color:#454b58;line-height:1.8;margin:0 0 14px;font-size:14px}.rv-name{display:flex;gap:8px;align-items:center;margin:0 0 20px;font-weight:700;font-size:13px}.rv-name input{width:220px;padding:7px 9px;border:1px solid #d5d9e2;border-radius:8px;font:inherit}.rv-slide{display:grid;grid-template-columns:minmax(0,1fr) 320px;gap:16px;margin-bottom:24px;align-items:start}.rv-slide .hs-scaler{border-radius:8px;box-shadow:0 2px 12px rgba(20,22,28,.12)}.rv-note b{display:block;font-size:13px;margin-bottom:6px}.rv-note textarea{width:100%;box-sizing:border-box;min-height:150px;font:13px/1.6 inherit;padding:8px 10px;border:1px solid #d5d9e2;border-radius:8px}.rv-bar{position:sticky;bottom:0;display:flex;gap:12px;align-items:center;justify-content:flex-end;padding:12px 16px;background:#fff;border-top:1px solid #e6e8ee}.rv-bar button{height:42px;padding:0 20px;border:0;border-radius:10px;background:#14161c;color:#fff;font:700 14px inherit;cursor:pointer}#rvCount{color:#7a8190;font-size:12px}@media (max-width:820px){.rv-slide{grid-template-columns:1fr}}`;

/** A self-contained page reviewers open in any browser; they type comments per slide and send back a small JSON. */
async function exportReviewFile() {
  if (!state.deck) return;
  toast("レビュー用ファイルを作っています…");
  try {
    await measureAll();
    const { deck, assets } = await portableDeck();
    const fits = state.deck.slides.map((_, i) => { const fit = fitFor(i); return fit ? { fs: fit.fs, ts: fit.ts } : null; });
    const titles = deck.slides.map((slide) => strip(slide.title || slide.message || typeLabel(slide.type)));
    const body = `<div class="rv"><h1>${esc(deck.title)}</h1>
<p class="rv-lead">各スライドの右の欄に、気になる点・直してほしい点を書いてください（このブラウザに自動で残ります）。最後に下の「コメントを書き出す」を押し、できたファイルを作成者に送ってください。</p>
<label class="rv-name">お名前 <input id="rvName" type="text" maxlength="40"></label>
<div id="rvList"></div>
<section class="rv-slide"><div></div><div class="rv-note"><b>資料全体へのコメント</b><textarea data-slide="all" placeholder="全体の流れ・トーンなど"></textarea></div></section>
</div>
<div class="rv-bar"><span id="rvCount"></span><button id="rvExport" type="button">コメントを書き出す</button></div>`;
    const html = await standaloneHtml({ title: `レビュー：${deck.title}`, body, boot: REVIEW_BOOT, data: { deck, fits, assets, titles, key: `hs-review-${state.historyId || Date.now()}` }, background: "#f2f3f6", extraCss: REVIEW_CSS });
    await downloadBlob(new Blob([html], { type: "text/html" }), `レビュー用_${fileSafe(deck.title)}.html`);
    toast("レビュー用ファイルを書き出しました。相手はブラウザで開いてコメントを書き、結果のファイルを返送します");
  } catch (error) {
    toast(`書き出せませんでした：${error.message}`);
  }
}

// ---------------------------------------------------------------- JSON, source files, imports

async function saveJsonFile() {
  if (!state.deck) return;
  try {
    const { deck, bytes } = await portableDeck();
    const payload = { app: "HTML PPTX Studio", version: APP_VERSION, ...deck };
    await downloadBlob(new Blob([JSON.stringify(payload, null, 1)], { type: "application/json" }), `${fileSafe(deck.title, "slides")}.json`);
    toast(bytes > 20_000_000 ? `JSONで保存しました（動画・写真を含め約${Math.round(bytes / 1_000_000)}MB）` : "JSONで保存しました（別のPCで「JSONを読み込む」から続きを編集できます）");
  } catch (error) {
    toast(`保存できませんでした：${error.message}`);
  }
}

function openJsonDialog() {
  $("slideDataInput").value = state.deck ? JSON.stringify({ deckTitle: state.deck.title, purpose: state.deck.purpose, audience: state.deck.audience, theme: state.deck.theme, transition: state.deck.transition, slideData: state.deck.slides }, null, 2) : "";
  clearStatus("jsonStatus");
  $("jsonDialog").showModal();
}

/** Photos pasted into JSON (data: URLs) and pictures from imports move into the browser store. */
async function adoptDeckMedia(deck) {
  for (const slide of deck.slides) {
    try {
      if (typeof slide.customImage === "string" && slide.customImage.startsWith("data:image/")) {
        const src = await putMedia(await dataUrlToBlob(slide.customImage), "取り込んだ写真");
        slide.media = { src, kind: "image", name: "取り込んだ写真", ...(slide.imagePlacement ? { placement: slide.imagePlacement } : {}) };
        delete slide.customImage;
        delete slide.imagePlacement;
      }
      if (slide.type === "imageText" && typeof slide.image === "string" && slide.image.startsWith("data:image/") && !slide.media) {
        const src = await putMedia(await dataUrlToBlob(slide.image), "取り込んだ画像");
        slide.media = { src, kind: "image", name: "取り込んだ画像", fit: "contain" };
        delete slide.image;
      }
    } catch { /* keep the data URL as it is */ }
  }
  return storeInlineMedia(deck);
}

async function applyJson() {
  try {
    const deck = await adoptDeckMedia(normalizeDeck(JSON.parse($("slideDataInput").value)));
    state.historyId = null;
    loadDeck(deck, { source: "JSON" });
    $("jsonDialog").close();
  } catch (error) {
    showStatus("jsonStatus", `読み込めません：${error.message}`, "error");
  }
}

function updateBriefCount() {
  $("briefCount").textContent = `${$("briefInput").value.length.toLocaleString("ja-JP")}字`;
}

async function readSourceFile(file) {
  if (!file) return;
  if (file.size > 15_000_000 && !/\.json$/i.test(file.name)) return showStatus("generationStatus", "ファイルが大きすぎます（15MBまで）。", "error");
  if (/\.json$/i.test(file.name)) {
    try {
      const deck = await adoptDeckMedia(normalizeDeck(JSON.parse(await file.text())));
      state.historyId = null;
      loadDeck(deck, { source: file.name });
      return;
    } catch { /* not a deck: treat as text */ }
  }
  let text = "";
  try {
    if (/\.(pptx|docx|pdf)$/i.test(file.name)) showStatus("generationStatus", `${file.name} から文章を取り出しています…`, "info");
    text = await extractFileText(file);
  } catch (error) {
    return showStatus("generationStatus", `${file.name} を読み込めません：${error.message}`, "error");
  }
  const brief = $("briefInput");
  brief.value = brief.value.trim() ? `${brief.value.trim()}\n\n--- ${file.name} ---\n${text.trim()}` : text.trim();
  updateBriefCount();
  saveCurrent();
  showStatus("generationStatus", `${file.name} を素材に追加しました（${text.length.toLocaleString("ja-JP")}字）。`, "success");
}

/** How to bring a PowerPoint over: as it looks (objects in place) or rebuilt into the studio's layouts. */
function chooseImportMode(name) {
  return new Promise((resolve) => {
    const dialog = $("importModeDialog");
    $("importModeName").textContent = `「${name}」をどのように取り込みますか？`;
    let answer = null;
    for (const button of dialog.querySelectorAll("[data-mode]")) button.onclick = () => { answer = button.dataset.mode; dialog.close(); };
    dialog.addEventListener("close", () => resolve(answer), { once: true });
    dialog.showModal();
    dialog.querySelector('[data-mode="exact"]').focus();
  });
}

async function importExistingDeck(file, mode = null) {
  if (!file) return;
  if (file.size > 30_000_000) return showStatus("generationStatus", "ファイルが大きすぎます（30MBまで）。", "error");
  const pptx = /\.pptx$/i.test(file.name);
  mode = pptx ? mode || await chooseImportMode(file.name) : "layout";
  if (!mode) return;
  showStatus("generationStatus", `${file.name} を${mode === "exact" ? "見た目どおりに" : ""}取り込んでいます…`, "info");
  try {
    const response = await fetch(`/api/import?name=${encodeURIComponent(file.name)}${mode === "exact" ? "&mode=exact" : ""}`, { method: "POST", credentials: "same-origin", headers: { "content-type": "application/octet-stream" }, body: file });
    const result = await response.json();
    if (!response.ok) throw new Error(result.error || `HTTP ${response.status}`);
    clearStatus("generationStatus");
    state.historyId = null;
    const deck = await adoptDeckMedia(normalizeDeck(result));
    loadDeck(deck, { imported: { name: file.name, fidelity: result.fidelity, stats: result.stats } });
    if (result.fidelity === "exact" && autoHtmlOn()) enhanceDeck("all", { quiet: true });
    const s = result.stats || {};
    const kept = result.fidelity === "exact"
      ? [s.pictures && `写真${s.pictures}`, s.tables && `表${s.tables}`, s.charts && `グラフ${s.charts}`, s.animations && `アニメーション${s.animations}`].filter(Boolean).join("・")
      : [s.tables && `表${s.tables}`, s.charts && `グラフ${s.charts}`, s.images && `写真${s.images}`, s.notes && `ノート${s.notes}`].filter(Boolean).join("・");
    toast(`${deck.slides.length}枚を${result.fidelity === "exact" ? "見た目どおりに" : ""}取り込みました${kept ? `（${kept}を保持）` : ""}`);
    if (result.comments?.length) postReview(result.comments, file.name);
  } catch (error) {
    showStatus("generationStatus", `取り込めませんでした：${error.message}`, "error");
  }
}

/** 背景の書式設定: one slide's background (or every slide's, "all"); null or {} takes it away. */
function setBackground(target, bg, { undo = true } = {}) {
  if (!state.deck) return;
  const value = E.normalizeBackground(bg);
  const slides = target === "all" ? state.deck.slides : [state.deck.slides[target]].filter(Boolean);
  if (!slides.length) return;
  if (undo) pushUndo();
  for (const slide of slides) { if (value) slide.background = clone(value); else delete slide.background; }
  markChanged({ structural: true });
}

/** 最終版にする / 編集する: the document's 状態 set to 最終版 (or no longer). */
function setFinal(on) {
  if (!state.deck) return;
  const info = { ...(infoOf(state.deck.info) || {}) };
  if (on) info.status = FINAL_STATUS; else delete info.status;
  editor.select([]);
  setDeckFields({ info: infoOf(info) });
  toast(on ? "最終版にしました（メッセージ バーの「編集する」で編集できます）" : "編集できるようにしました");
  renderStage();
  editorUi.renderRibbon(true);
}
/** The master view's own page: a blank slide whose objects are the deck's master objects (written back by setObjects). */
function masterSlide() {
  const slide = (state.masterSlide ||= { type: "blank", hideTitle: true, hideMaster: true, sid: "master" });
  slide.elements = state.deck?.masterObjects || [];
  return slide;
}
function openMaster() {
  if (!state.deck || state.mode !== "edit") return;
  if (state.view !== "single") setView("single");
  if (editor.typing) editor.stopTyping(true);
  editor.clear();
  state.masterEdit = true;
  renderAll();
}
function closeMaster() {
  if (!state.masterEdit) return;
  if (editor.typing) editor.stopTyping(true);
  editor.clear();
  state.masterEdit = false;
  renderAll();
}
function masterCallout() {
  if (!state.masterEdit || state.mode !== "edit") return null;
  return h("div", { class: "callout callout-master" },
    h("div", { class: "text" }, h("b", {}, "スライド マスター"), "ここに置いた図形・文字・図は、すべてのスライドに出ます（スライドごとにデザイン → 背景グラフィックを表示しないで外せます）。"),
    h("button", { class: "btn btn-master-close", type: "button", onclick: () => closeMaster() }, "マスター表示を閉じる"));
}

function finalCallout() {
  if (!isFinal(state.deck) || state.mode !== "edit") return null;
  return h("div", { class: "callout callout-final" },
    h("div", { class: "text" }, h("b", {}, "最終版"), "作成者がこの資料を最終版として設定し、編集しないようにしています。"),
    h("button", { class: "btn btn-final-edit", type: "button", onclick: () => setFinal(false) }, "編集する"));
}

function importCallout() {
  if (!state.imported) return null;
  const canAi = state.codexAuthorized;
  const { name, fidelity } = state.imported;
  const close = h("button", { class: "btn btn-ghost btn-icon", type: "button", title: "閉じる", onclick: () => { state.imported = null; renderStage(); } }, "✕");
  if (fidelity === "exact") {
    const s = state.imported.stats || {};
    const notes = [s.hidden && `非表示のスライド${s.hidden}枚も編集用に保持しています`, s.media && `ビデオ・オーディオ${s.media}個は再生できる形で取り込みました（トリミング・フェード・ブックマークも）`, s.unsupported && `再現できない要素が${s.unsupported}個あります。原本との照合が必要です`].filter(Boolean);
    const auto = state.imported.auto;
    const editAnimations = h("button", { class: "btn btn-ghost", type: "button", title: "動き（自動の動き・アニメーション・インタラクション）をアニメーション ウィンドウで見る・直す", onclick: () => { if (state.view !== "single") setView("single"); editorUi.showTab("animation", { open: true }); setPanel("anim"); } }, "アニメーションを編集");
    if (auto) {
      return h("div", { class: "callout" },
        h("div", { class: "text" },
          h("b", {}, `「${name}」を見た目どおりに取り込み、HTMLの動きを自動で付けました`),
          `${auto.summary}に${auto.count}か所。${s.animations ? `PowerPointのアニメーション${s.animations}件はそのまま残し、その前に入ります。` : ""}動きはアニメーション ウィンドウで1つずつ変えたり外したりできます。`,
          notes.length ? h("span", { class: "callout-note" }, `${notes.join("。")}。`) : null),
        h("button", { class: "btn btn-ai", type: "button", title: "この取り込みのスライドを最初から、動きごと確認する", onclick: () => { setView("single"); select(0); previewMotion(); } }, "▶ 動きを見る"),
        editAnimations,
        h("button", { class: "btn", type: "button", title: "自動で付けた動きを外して、PowerPointのままに戻す（⌘Zと同じ）", onclick: () => undoAutoHtml() }, "元に戻す"),
        h("button", { class: "btn btn-ghost", type: "button", title: "次からPowerPointを取り込んでも自動では付けない（インタラクション タブで戻せます）", onclick: () => setAutoHtml(false) }, "次から自動にしない"),
        close);
    }
    return h("div", { class: "callout" },
      h("div", { class: "text" },
        h("b", {}, `「${name}」を見た目どおりに取り込みました`),
        `元の配置・色・文字のまま、図形・表・グラフ・写真を1つずつ動かして直せます。${s.animations ? `PowerPointのアニメーション${s.animations}件も引き継いでいます。` : ""}HTMLにした今は、PowerPointにはない動き（マウスを乗せると浮く・説明が出る、クリックで詳細・拡大・裏返す・タブ切り替え、数字のカウントアップ、グラフが伸びて値をなぞれる）を付けられます。`,
        notes.length ? h("span", { class: "callout-note" }, `${notes.join("。")}。`) : null),
      h("button", { class: "btn btn-ai", type: "button", title: "すべてのスライドに、HTMLならではの動きをおまかせで付けます（⌘Zで戻せます）", onclick: () => enhanceDeck("all") }, "✦ HTMLの動きをおまかせで付ける"),
      h("button", { class: "btn", type: "button", title: "リボンの「インタラクション」で、マウスの反応・クリックの操作・ずっと動く動きを部品に付けます", onclick: () => { if (state.view !== "single") setView("single"); editorUi.showTab("interact", { open: true }); } }, "自分で付ける"),
      editAnimations,
      close);
  }
  return h("div", { class: "callout" },
    h("div", { class: "text" },
      h("b", {}, `「${name}」を取り込みました`),
      fidelity === "high" ? "文章・表・グラフ・写真・ノートを元の構成のまま読み込んでいます。" : "文章だけを読み込んでいます（PDF・Wordは表やグラフを文章として扱います）。",
      canAi ? "AIでブラッシュアップすると、結論・図解・クリックで開く詳細・ノートを備えたプレゼンに磨き上げます。" : "Codexに接続すると、AIで磨き上げられます。自分で直す場合は構成チェックの指摘を順に直してください。"),
    h("button", { class: "btn btn-ai", type: "button", disabled: !canAi || state.aiBusy, onclick: () => openRewriteDialog(BRUSHUP_INSTRUCTION, true) }, "✦ AIでブラッシュアップ"),
    close);
}

async function openSample() {
  try {
    const sample = await jsonFetch(`/samples/ai-rollout.json?v=${APP_VERSION}`, { method: "GET" });
    state.historyId = null;
    loadDeck(normalizeDeck(sample), { source: "サンプル" });
  } catch (error) {
    showStatus("generationStatus", `サンプルを開けません：${error.message}`, "error");
  }
}

// ---------------------------------------------------------------- templates

const T_CARD = (title, desc) => ({ title, desc });
// A worked example of a deck the audience can work with: each page carries one action chosen by what its
// headline says (結論と根拠・切り口・差分・原因と結果・不足と打ち手・流れ), and evidence behind its figures.
const INTERACTIVE_SOURCE = "社内試行（5部門・120人・3か月）";
const INTERACTIVE_TEMPLATE = {
  id: "interactive", name: "動く資料（クリックで確かめる）", desc: "6つの構造ごとに1つの操作。例：生成AIの全社展開", audience: "役員", purpose: "意思決定・報告",
  slides: [
    { type: "title", title: "生成AIの全社展開", subtitle: "クリックで根拠を確かめられる資料（例）" },
    { type: "headerCards", title: "試行の結果", takeaway: "3か月で**月1,440時間**の作業が減った", source: INTERACTIVE_SOURCE,
      items: [{ title: "月1,440時間の削減", desc: "5部門・120人の3か月の試行で", icon: "clock" }, { title: "満足度4.3点", desc: "5点満点。8割が使い続けたい", icon: "smile" }, { title: "年間効果 約1億円", desc: "初年度の予算は2,400万円", icon: "yen" }],
      details: [
        { target: "takeaway", title: "月1,440時間の内訳", text: "資料作成と議事録で全体の7割を占める。", rows: [{ label: "資料作成", value: "620時間" }, { label: "議事録", value: "380時間" }, { label: "データ集計", value: "290時間" }, { label: "その他", value: "150時間" }], source: INTERACTIVE_SOURCE, note: "各部門の作業記録（試行前1か月と試行3か月目）の差。" },
        { target: "items[1]", title: "満足度の分布", text: "「使い続けたい」が8割を超えた。", rows: [{ label: "5点", value: "46%" }, { label: "4点", value: "38%" }, { label: "3点以下", value: "16%" }], source: "試行後アンケート（n=120）" },
        { target: "items[2]", title: "年間効果の出し方", text: "削減時間に平均の時間単価を掛けた試算。", rows: [{ label: "削減時間（年）", value: "17,280時間" }, { label: "時間単価", value: "5,800円" }, { label: "効果", value: "約1億円" }], note: "全社展開後の見込み。時間単価は社内の平均人件費から。" },
      ] },
    { type: "imageText", title: "作業ごとの削減時間", takeaway: "見方を変えても**資料作成**の効果が大きい", source: INTERACTIVE_SOURCE, imagePosition: "left",
      image: { chartType: "rank", data: { title: "作業ごとの削減時間", unit: "時間", highlight: "資料作成", views: [
        { label: "月の合計", items: [{ label: "資料作成", value: 620 }, { label: "議事録", value: 380 }, { label: "データ集計", value: 290 }, { label: "その他", value: 150 }] },
        { label: "1人あたり", items: [{ label: "資料作成", value: 5.2 }, { label: "議事録", value: 3.2 }, { label: "データ集計", value: 2.4 }, { label: "その他", value: 1.3 }] },
      ] } },
      points: ["切り口を切り替えると並びが変わる", "どちらで見ても資料作成が1位"] },
    { type: "imageText", title: "1件あたりの作業時間の変化", takeaway: "減ったのは**下書きと清書**で、確認の時間は残った", source: INTERACTIVE_SOURCE, imagePosition: "left",
      image: { chartType: "shift", data: { title: "資料1件あたりの作業時間（分）", unit: "分", beforeLabel: "試行前", afterLabel: "試行後", items: [{ label: "下書き", before: 90, value: 35 }, { label: "清書", before: 60, value: 25 }, { label: "確認", before: 30, value: 28 }] } },
      points: ["点線が減った分", "確認は人が担うので残る"] },
    { type: "simulator", title: "全社に広げたときの試算", takeaway: "利用率を**8割**まで上げれば目標に届く", source: "試行の実績から試算",
      inputs: [{ label: "利用する人数", value: 600, min: 120, max: 1200, step: 20, unit: "人" }, { label: "1人あたり削減", value: 12, min: 4, max: 20, step: 1, unit: "時間/月" }, { label: "利用率", value: 50, min: 10, max: 100, step: 5, unit: "%" }],
      formula: "a × b × c ÷ 100", resultLabel: "月の削減時間", resultUnit: "時間", compareLabel: "目標", compareValue: 5000 },
    { type: "gap", title: "目標までの不足と打ち手", takeaway: "3つの打ち手を重ねると**目標に届く**", source: "試行の実績から試算", unit: "時間",
      targetLabel: "目標", target: 5000, currentLabel: "いまの見込み", current: 3600,
      measures: [{ title: "業務別テンプレート", value: 600, desc: "部門ごとの指示文30種" }, { title: "各部の推進役", value: 500, desc: "身近な相談先を置く" }, { title: "入力ルール", value: 400, desc: "使ってよい情報を明示" }] },
    { type: "gantt", title: "展開の計画", takeaway: "いまは**本社の準備**、12月から支社へ広げる", periods: ["10月", "11月", "12月", "1月", "2月", "3月"], now: 0.5,
      items: [{ title: "本社", desc: "試行を拡大", start: 0, span: 2 }, { title: "支社", desc: "順に展開", start: 2, span: 2 }, { title: "全拠点", desc: "本格運用", start: 4, span: 2 }] },
    { type: "closing", title: "お願いしたいこと", message: "10月の経営会議で、初年度予算2,400万円のご承認をお願いします" },
  ],
};
const TEMPLATES = [
  INTERACTIVE_TEMPLATE,
  {
    id: "exec", name: "役員報告（意思決定）", desc: "結論→根拠→計画→お願いの順で、判断をもらう", audience: "役員", purpose: "意思決定・報告",
    slides: [
      { type: "title", title: "【資料タイトル】", subtitle: "【何を決めてほしいか】" },
      { type: "executiveSummary", title: "結論：【一文で】", takeaway: "【判断してほしいことを一文で】", conclusion: "【結論】", items: [T_CARD("背景", "【なぜ今この判断が必要か】"), T_CARD("根拠", "【結論を支える事実・数字】"), T_CARD("リスク", "【想定されるリスクと対策】")], action: "【いつまでに何を決めてほしいか】" },
      { type: "kpi", title: "現状の数字", takeaway: "【数字から言えること】", items: [{ label: "【指標1】", value: "【数値】", change: "【前年比など】" }, { label: "【指標2】", value: "【数値】", change: "【前年比など】" }, { label: "【指標3】", value: "【数値】", change: "【前年比など】" }] },
      { type: "compare", title: "現状と目指す姿", takeaway: "【何をどう変えるか】", leftTitle: "現状", rightTitle: "目指す姿", leftItems: ["【現状の課題1】", "【現状の課題2】"], rightItems: ["【変えた後の状態1】", "【変えた後の状態2】"] },
      { type: "roadmap", title: "実行計画", takeaway: "【段階的に進める方針】", items: [T_CARD("【時期1】", "【やること】"), T_CARD("【時期2】", "【やること】"), T_CARD("【時期3】", "【やること】")] },
      { type: "table", title: "投資と効果", takeaway: "【投資に見合う効果がある根拠】", headers: ["項目", "金額", "時期", "効果"], rows: [["【項目】", "【金額】", "【時期】", "【効果】"], ["【項目】", "【金額】", "【時期】", "【効果】"]] },
      { type: "faq", title: "想定される質問", takeaway: "【主な懸念への回答方針】", items: [{ q: "【質問1】", a: "【回答】" }, { q: "【質問2】", a: "【回答】" }] },
      { type: "closing", title: "お願いしたいこと", message: "【誰が・いつまでに・何を決めるか】" },
    ],
  },
  {
    id: "training", name: "研修・勉強会", desc: "目的→ポイント→注意点→手順→明日からの行動", audience: "新入社員", purpose: "研修・勉強会",
    slides: [
      { type: "title", title: "【研修タイトル】", subtitle: "【この研修で身につくこと】" },
      { type: "agenda", title: "本日の流れ", takeaway: "【研修のゴール】", items: ["【テーマ1：概要】", "【テーマ2：概要】", "【テーマ3：概要】", "まとめと明日からの行動"] },
      { type: "statement", title: "なぜ学ぶのか", text: "【学ぶ理由を大きな一文で】", takeaway: "【補足の一文】" },
      { type: "headerCards", title: "押さえるポイント", takeaway: "【最も大事なこと】", items: [T_CARD("【ポイント1】", "【説明】"), T_CARD("【ポイント2】", "【説明】"), T_CARD("【ポイント3】", "【説明】")] },
      { type: "checklist", title: "やってはいけないこと", takeaway: "【守るべきルール】", items: [T_CARD("【注意点1】", "【理由】"), T_CARD("【注意点2】", "【理由】"), T_CARD("【注意点3】", "【理由】")] },
      { type: "process", title: "基本の手順", takeaway: "【手順の要点】", steps: ["【手順1：説明】", "【手順2：説明】", "【手順3：説明】"] },
      { type: "faq", title: "よくある質問", takeaway: "【迷ったときの考え方】", items: [{ q: "【質問1】", a: "【回答】" }, { q: "【質問2】", a: "【回答】" }] },
      { type: "closing", title: "明日から試すこと", message: "【行動1】\n【行動2】\n【行動3】" },
    ],
  },
  {
    id: "review", name: "振り返り・結果報告", desc: "結果→要因→学び→次の打ち手", audience: "部長・マネージャー", purpose: "振り返り",
    slides: [
      { type: "title", title: "【施策名】の振り返り", subtitle: "【期間・対象】" },
      { type: "kpi", title: "結果", takeaway: "【結果を一文で】", items: [{ label: "【指標1】", value: "【数値】", change: "【計画比・前年比】" }, { label: "【指標2】", value: "【数値】", change: "【計画比・前年比】" }, { label: "【指標3】", value: "【数値】", change: "【計画比・前年比】" }] },
      { type: "statsCompare", title: "計画との比較", takeaway: "【差が出た指標と理由】", leftTitle: "計画", rightTitle: "実績", stats: [{ label: "【指標1】", leftValue: "【計画】", rightValue: "【実績】" }, { label: "【指標2】", leftValue: "【計画】", rightValue: "【実績】" }] },
      { type: "logicTree", title: "要因の分析", takeaway: "【最大の要因】", root: "【結果】", branches: [{ title: "【要因1】", items: ["【事象】"], highlight: true }, { title: "【要因2】", items: ["【事象】"] }, { title: "【要因3】", items: ["【事象】"] }] },
      { type: "compare", title: "良かった点と課題", takeaway: "【次に生かすこと】", leftTitle: "良かった点", rightTitle: "課題", leftItems: ["【良かった点1】", "【良かった点2】"], rightItems: ["【課題1】", "【課題2】"] },
      { type: "roadmap", title: "次の打ち手", takeaway: "【次期の方針】", items: [T_CARD("【すぐやること】", "【内容】"), T_CARD("【来月までに】", "【内容】"), T_CARD("【次期】", "【内容】")] },
      { type: "closing", title: "次のアクション", message: "【誰が・いつまでに・何をするか】" },
    ],
  },
  {
    id: "proposal", name: "企画提案", desc: "課題→提案の3本柱→効果→スケジュール", audience: "部長・マネージャー", purpose: "提案・企画",
    slides: [
      { type: "title", title: "【企画名】のご提案", subtitle: "【誰のどんな課題を解決するか】" },
      { type: "hero", title: "【提案を象徴するひと言】", takeaway: "【実現したい姿を一文で】", visualAsset: "transformationRoadmap", photoMotion: "zoom" },
      { type: "executiveSummary", title: "提案の要旨", takeaway: "【提案を一文で】", conclusion: "【何をすると、何がどう良くなるか】", items: [T_CARD("課題", "【いま困っていること】"), T_CARD("提案", "【何をするか】"), T_CARD("効果", "【期待できる成果】")], action: "【承認してほしいこと】" },
      { type: "headerThreeSummary", title: "提案の3本柱", takeaway: "【3つをセットで進める理由】", items: [T_CARD("【柱1】", "【内容】"), T_CARD("【柱2】", "【内容】"), T_CARD("【柱3】", "【内容】")], summary: "【3本柱で実現する姿】" },
      { type: "beforeAfter", title: "導入前後の変化", takeaway: "【最も大きな変化】", leftTitle: "導入前", rightTitle: "導入後", leftItems: ["【今の状態】"], rightItems: ["【導入後の状態】"] },
      { type: "timeline", title: "スケジュール", takeaway: "【いつ何が決まれば間に合うか】", milestones: [{ date: "【時期】", label: "【準備】", state: "next" }, { date: "【時期】", label: "【試行】" }, { date: "【時期】", label: "【本格展開】" }] },
      { type: "table", title: "体制と費用", takeaway: "【必要な体制・費用の要点】", headers: ["項目", "内容", "費用"], rows: [["【体制】", "【担当・人数】", "【費用】"], ["【システム】", "【内容】", "【費用】"]] },
      { type: "closing", title: "お願いしたいこと", message: "【いつまでに何を承認してほしいか】" },
    ],
  },
  {
    id: "weekly", name: "週次報告", desc: "今週の数字→トピック→進捗→課題→来週", audience: "部長・マネージャー", purpose: "進捗共有",
    slides: [
      { type: "title", title: "週次報告（【月日】週）", subtitle: "【チーム名】" },
      { type: "kpi", title: "今週の数字", takeaway: "【今週のポイントを一文で】", items: [{ label: "【指標1】", value: "【数値】", change: "【前週比】" }, { label: "【指標2】", value: "【数値】", change: "【前週比】" }, { label: "【指標3】", value: "【数値】", change: "【前週比】" }] },
      { type: "cards", title: "主なトピック", takeaway: "【最も伝えたいトピック】", items: [T_CARD("【トピック1】", "【内容】"), T_CARD("【トピック2】", "【内容】"), T_CARD("【トピック3】", "【内容】")] },
      { type: "checklist", title: "案件の進捗", takeaway: "【遅れている案件と対応】", items: [{ title: "【案件1】", desc: "【状況】", done: true }, { title: "【案件2】", desc: "【状況】" }, { title: "【案件3】", desc: "【状況】" }] },
      { type: "compare", title: "課題と対応", takeaway: "【相談したいこと】", leftTitle: "課題", rightTitle: "対応", leftItems: ["【課題1】", "【課題2】"], rightItems: ["【対応1】", "【対応2】"] },
      { type: "closing", title: "来週の予定", message: "【来週やること・決めること】" },
    ],
  },
];

function openTemplateDialog() {
  const renderTemplate = (template) => {
    const deck = { title: template.slides[0].title, theme: preferredTheme(), transition: "fade", motion: DEFAULT_MOTION, slides: template.slides };
    try { return E.mount(E.render(template.slides[1], { deck, index: 1, mode: "thumb", assetBase: "/assets/" })); } catch { return null; }
  };
  useFonts([preferredTheme()]);
  $("templateGrid").replaceChildren(
    ...TEMPLATES.map((template) => h("button", { type: "button", onclick: () => { $("templateDialog").close(); fromTemplate(template); } },
      renderTemplate(template), h("b", {}, template.name), h("span", {}, `${template.desc}（${template.slides.length}枚）`))),
    h("button", { type: "button", onclick: () => { $("templateDialog").close(); blankDeck(); } }, h("b", {}, "白紙"), h("span", {}, "枚数だけ決めて自由に作る")));
  $("templateDialog").showModal();
}

function fromTemplate(template) {
  const form = formState();
  const slides = clone(template.slides);
  if (!form.audience) { $("audienceInput").value = template.audience; showChoice("audienceInput"); }
  slides[0].date = new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long" });
  state.historyId = null;
  loadDeck(normalizeDeck({ title: slides[0].title, purpose: form.purpose || template.purpose, audience: form.audience || template.audience, theme: preferredTheme(), slides }), { source: template.name });
  toast(JSON.stringify(template.slides).includes("【")
    ? "【】の部分を書き換えてください。構成チェックが残りを知らせます"
    : "記入例の資料です。発表（F5）でクリック・スライダーを試してから、数字と文言を自分の内容に置き換えてください");
}

function blankDeck() {
  const form = formState();
  const count = Math.max(3, Math.min(12, form.slideCount));
  const title = form.brief.trim().split(/[。\n！？]/)[0].trim().slice(0, 30) || "新しい資料";
  const middle = ["agenda", "executiveSummary", "statement", "content", "compare", "process", "kpi", "table", "roadmap", "headerThreeSummary", "faq"];
  const slides = [
    { type: "title", title, subtitle: "", date: new Date().toLocaleDateString("ja-JP", { year: "numeric", month: "long" }) },
    ...Array.from({ length: count - 2 }, (_, i) => defaultSlide(middle[i % middle.length])),
    { type: "closing", title: "次のアクション", message: "誰が・いつまでに・何をするか" },
  ];
  state.historyId = null;
  loadDeck(normalizeDeck({ title, purpose: form.purpose, audience: form.audience, theme: preferredTheme(), slides }), { source: "白紙" });
}

// ---------------------------------------------------------------- wiring

function bind() {
  $("autoImageAction").addEventListener("click", () => {
    const run = autoImageRun;
    if (run && !run.cancelled && (run.inFlight || run.next < run.candidates.length)) {
      run.cancelled = true;
      showAutoImageProgress(run);
    } else if (state.deck) autoIllustrateDeck(state.deck);
  });
  $("createTab").addEventListener("click", () => setMode("create"));
  $("editTab").addEventListener("click", () => { if (state.deck) { setMode("edit"); renderAll(); } });
  $("resumeEditBtn").addEventListener("click", () => { setMode("edit"); renderAll(); });
  $("askChatGptBtn").addEventListener("click", () => generateOutline());
  $("quickDeckBtn").addEventListener("click", () => generateDeck());
  $("outlineBuildBtn").addEventListener("click", buildFromOutline);
  $("outlineForm").addEventListener("submit", (event) => {
    event.preventDefault();
    const instruction = $("outlineInstruction").value.trim();
    if (!instruction) return $("outlineInstruction").focus();
    generateOutline({ instruction });
  });
  $("outlineAddBtn").addEventListener("click", () => {
    if (!state.outline) return;
    state.outline.slides.splice(state.outline.slides.length - 1, 0, { type: "content", title: "新しいスライド", takeaway: "", content: "" });
    renderOutlinePanel();
    saveCurrent();
  });
  $("outlineDiscardBtn").addEventListener("click", () => { state.outline = null; renderOutlinePanel(); saveCurrent(); });
  $("briefInput").addEventListener("keydown", (event) => {
    if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) { event.preventDefault(); generateOutline(); }
  });
  $("blankDeckBtn").addEventListener("click", openTemplateDialog);
  $("sampleDeckBtn").addEventListener("click", openSample);
  $("importDeckBtn").addEventListener("click", () => $("importDeckFile").click());
  $("importDeckFile").addEventListener("change", (event) => { importExistingDeck(event.target.files?.[0]); event.target.value = ""; });
  $("importJsonBtn").addEventListener("click", openJsonDialog);
  $("connectCodexBtn").addEventListener("click", startCodexLogin);
  $("saveDeckBtn").addEventListener("click", () => saveToLibrary());
  $("libraryBtn").addEventListener("click", openLibrary);
  $("librarySearch").addEventListener("input", renderLibrary);
  $("saveAsNewBtn").addEventListener("click", () => saveToLibrary({ asNew: true }));
  $("copyDeviceCodeBtn").addEventListener("click", copyDeviceCode);
  $("passcodePanel").addEventListener("submit", submitPasscode);
  $("historyBtn").addEventListener("click", openHistory);
  $("chatTab").addEventListener("click", () => setPanel("chat"));
  $("formTab").addEventListener("click", () => setPanel("form"));
  $("formatTab").addEventListener("click", () => setPanel("format"));
  $("animTab").addEventListener("click", () => setPanel("anim"));
  $("commentTab").addEventListener("click", () => setPanel("comment"));
  $("a11yTab").addEventListener("click", () => setPanel("a11y"));
  // The clipboard: objects, pictures and text go onto the slide; copying objects copies them.
  document.addEventListener("copy", (event) => editor.onCopy(event));
  document.addEventListener("cut", (event) => editor.onCopy(event, true));
  document.addEventListener("paste", (event) => { editor.onPaste(event); });
  window.addEventListener("resize", () => { if (state.mode === "edit" && state.view === "single") { editor.applyZoom(); editor.draw(); } });
  $("chatForm").addEventListener("submit", (event) => { event.preventDefault(); sendChat(); });
  $("chatInput").addEventListener("keydown", (event) => {
    if (event.key === "Enter" && !event.shiftKey && !event.isComposing) { event.preventDefault(); sendChat(); }
  });
  $("chatInput").addEventListener("input", () => {
    const input = $("chatInput");
    input.style.height = "auto";
    input.style.height = `${Math.min(180, input.scrollHeight + 2)}px`;
    renderChatContext();
  });
  $("memoToggle").addEventListener("click", () => { $("memoBox").hidden = !$("memoBox").hidden; if (!$("memoBox").hidden) { $("memoInput").value = state.deck?.memo || ""; $("memoInput").focus(); } });
  $("memoInput").addEventListener("input", (event) => { if (!state.deck) return; state.deck.memo = event.target.value.slice(0, 2000); saveCurrent(); $("memoToggle").classList.toggle("active", Boolean(state.deck.memo.trim())); });
  $("chatClearBtn").addEventListener("click", () => {
    if (!state.chat.messages.length) return;
    if (!state.confirmClear) { state.confirmClear = true; toast("もう一度押すと、この資料の会話を消去します（資料は変わりません）"); setTimeout(() => { state.confirmClear = false; }, 5000); return; }
    state.confirmClear = false;
    state.chat.messages = [];
    proposals.clear();
    saveChat();
    renderChat();
  });
  $("undoBtn").addEventListener("click", () => undoRedo("undo"));
  $("redoBtn").addEventListener("click", () => undoRedo("redo"));
  $("shareBtn").addEventListener("click", () => (coedit.active ? coedit.showMenu() : coedit.start()));
  $("shareCopy").addEventListener("click", async () => { try { await navigator.clipboard.writeText($("shareLink").value); $("shareStatus").textContent = "コピーしました"; } catch { $("shareLink").select(); document.execCommand("copy"); $("shareStatus").textContent = "コピーしました"; } });
  $("downloadBtn").addEventListener("click", () => exportHtml());
  $("pptxBtn").addEventListener("click", () => exportPptx());
  $("fixAllBtn").addEventListener("click", fixAllOverflow);
  $("deckReviseBtn").addEventListener("click", () => openRewriteDialog());
  $("rewriteSubmit").addEventListener("click", rewriteDeck);
  $("presentBtn").addEventListener("click", () => openPresenter(state.selected));
  $("pdfBtn").addEventListener("click", printPdf);
  $("issueSummary").addEventListener("click", () => openCheckDialog(false));
  $("deckTitleInput").addEventListener("input", (event) => { if (!state.deck) return; beginEdit(); state.deck.title = event.target.value; markChanged(); });
  $("designBtn").addEventListener("click", openDesignDialog);
  renderCreateThemes();
  $("motionPreviewBtn").addEventListener("click", () => previewMotion());
  $("codeApply").addEventListener("click", applyCodeDialog);
  $("codeCopy").addEventListener("click", async () => { try { await navigator.clipboard.writeText($("codeText").value); toast("コピーしました"); } catch { $("codeText").select(); document.execCommand("copy"); toast("コピーしました"); } });
  $("transitionSelect").addEventListener("change", (event) => setDeckDesign({ transition: event.target.value }));
  $("entranceSelect").addEventListener("change", (event) => setDeckDesign({ motion: { entrance: event.target.value } }));
  $("hoverSelect").addEventListener("change", (event) => setDeckDesign({ motion: { hover: event.target.value } }));
  $("numbersCheck").addEventListener("change", (event) => setDeckDesign({ motion: { numbers: event.target.checked } }));
  $("ambientCheck").addEventListener("change", (event) => setDeckDesign({ motion: { ambient: event.target.checked } }));
  $("drawCheck").addEventListener("change", (event) => setDeckDesign({ motion: { draw: event.target.checked } }));
  $("designDialog").addEventListener("close", () => clearInterval(motionCardsTimer));
  fillMotionSelects();
  $("designPreviewBtn").addEventListener("click", () => { $("designDialog").close(); setView("single"); previewMotion(); });
  $("mediaUrlApply").addEventListener("click", applyMediaUrl);
  $("mediaUrlInput").addEventListener("keydown", (event) => { if (event.key === "Enter" && !event.isComposing) { event.preventDefault(); applyMediaUrl(); } });
  document.querySelectorAll("input[name=view]").forEach((radio) => radio.addEventListener("change", (event) => setView(event.target.value)));
  $("stageBody").addEventListener("click", (event) => {
    if (state.motionPreview || event.target.closest(".inline-tools, .ph-handle")) return;
    if (editor.consumeClick() || editor.tool || event.target.closest(".hs-obj, .ed-layer")) return;
    const badge = event.target.closest(".slide-wrap .hs-drill-badge[data-drill-to]");
    if (badge) { if (state.inline) finishInlineEdit(true); select(Number(badge.dataset.drillTo)); return; }
    const target = event.target.closest(".slide-wrap .hs-slide.hs-editable [data-field]");
    if (target && target.classList.contains("hs-editing")) return;
    if (state.inline) finishInlineEdit(true);
    if (!target) return;
    if (target.classList.contains("hs-t")) beginInlineEdit(target);
    else focusField(target.dataset.field);
  });
  $("stageBody").addEventListener("keydown", (event) => {
    if (!state.inline || !event.target.classList?.contains("hs-editing")) return;
    if (event.key === "Escape") { event.preventDefault(); finishInlineEdit(false); }
    else if (event.key === "Enter" && !event.shiftKey && !state.inline.multiline && !event.isComposing) { event.preventDefault(); finishInlineEdit(true); }
  });
  $("stageBody").addEventListener("focusout", (event) => {
    if (state.inline && event.target === state.inline.el && !event.relatedTarget?.closest?.(".inline-tools")) finishInlineEdit(true);
  });
  $("commandBtn").addEventListener("click", openCommandPalette);
  $("chatAttachBtn").addEventListener("click", () => $("chatAttachFile").click());
  $("chatAttachFile").addEventListener("change", (event) => { attachToChat(event.target.files?.[0]); event.target.value = ""; });
  $("reviewImportBtn").addEventListener("click", () => $("reviewImportFile").click());
  $("reviewImportFile").addEventListener("change", (event) => { importReview(event.target.files?.[0]); event.target.value = ""; });
  $("reviewExportBtn").addEventListener("click", exportReviewFile);
  $("jsonExportBtn").addEventListener("click", saveJsonFile);
  const chatMic = voiceButton("chatInput");
  if (chatMic) $("chatSend").before(chatMic);
  const briefMic = voiceButton("briefInput");
  if (briefMic) $("briefCount").after(briefMic);
  $("commandInput").addEventListener("input", () => { state.commandActive = 0; renderCommands(); });
  $("commandInput").addEventListener("keydown", (event) => {
    const items = filteredCommands();
    if (event.key === "ArrowDown") { event.preventDefault(); state.commandActive = Math.min(items.length - 1, state.commandActive + 1); renderCommands(); }
    else if (event.key === "ArrowUp") { event.preventDefault(); state.commandActive = Math.max(0, state.commandActive - 1); renderCommands(); }
    else if (event.key === "Enter" && !event.isComposing) { event.preventDefault(); const item = items[state.commandActive]; if (item) runCommand(item); }
  });
  // このまま出力: the export the check was opened for (PowerPoint or HTML).
  $("exportCheckGoBtn").addEventListener("click", () => { $("exportCheckDialog").close(); if (state.pendingExport === "pptx") exportPptx({ checked: true }); else exportHtml({ checked: true }); });
  $("exportCheckFixBtn").addEventListener("click", () => $("exportCheckDialog").close());
  document.addEventListener("click", (event) => { for (const menu of document.querySelectorAll(".more-menu[open]")) if (!menu.contains(event.target) || event.target.closest(".menu .btn")) menu.open = false; });
  $("notesBtn").addEventListener("click", openNotesDialog);
  $("notesRun").addEventListener("click", runNotes);
  $("replaceBtn").addEventListener("click", openReplace);
  $("findInput").addEventListener("input", updateFindCount);
  for (const id of ["findCase", "findWord", "findWidth", "findScope"]) $(id).addEventListener("change", () => { saveFindOptions(); updateFindCount(); });
  $("replaceAllBtn").addEventListener("click", replaceAll);
  $("helpBtn").addEventListener("click", () => $("helpDialog").showModal());
  $("briefInput").addEventListener("input", () => { updateBriefCount(); saveCurrent(); });
  for (const id of ["audienceInput", "purposeInput", "toneInput", "slideCount"]) $(id).addEventListener("change", saveCurrent);
  document.querySelectorAll("input[name=density]").forEach((radio) => radio.addEventListener("change", saveCurrent));
  wireChoices();
  $("sourceFile").addEventListener("change", (event) => { readSourceFile(event.target.files?.[0]); event.target.value = ""; });
  const drop = $("briefDrop");
  drop.addEventListener("dragover", (event) => { event.preventDefault(); drop.classList.add("dragging"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("dragging"));
  drop.addEventListener("drop", (event) => { event.preventDefault(); drop.classList.remove("dragging"); readSourceFile(event.dataTransfer.files?.[0]); });
  $("applyJsonBtn").addEventListener("click", applyJson);
  $("copyJsonBtn").addEventListener("click", async () => { try { await navigator.clipboard.writeText($("slideDataInput").value); showStatus("jsonStatus", "コピーしました。", "success"); } catch { showStatus("jsonStatus", "コピーできませんでした。", "error"); } });
  $("exportJsonBtn").addEventListener("click", saveJsonFile);
  $("jsonFile").addEventListener("change", async (event) => { const file = event.target.files?.[0]; if (file) $("slideDataInput").value = await file.text(); event.target.value = ""; });
  document.querySelectorAll("dialog [data-close]").forEach((button) => button.addEventListener("click", () => button.closest("dialog").close()));
  document.querySelectorAll("dialog").forEach((dialog) => dialog.addEventListener("click", (event) => { if (event.target === dialog) dialog.close(); }));
  $("sampleChips").append(...SAMPLES.map((sample) => h("button", { class: "chip", type: "button", onclick: () => { applyForm({ brief: sample.brief, audience: sample.audience, purpose: sample.purpose, slideCount: sample.slideCount, density: sample.density, tone: sample.tone }); saveCurrent(); } }, sample.label)));

  // A brief preview (an effect just chosen) gives way to whatever the person does next, as in PowerPoint.
  // A click on the slide itself only ends it; a click elsewhere (the film strip, the ribbon) also does its own job,
  // so the page is not redrawn under the pointer before that click lands.
  document.addEventListener("pointerdown", (event) => {
    if (!state.motionPreview?.brief || event.target.closest?.(".motion-banner")) return;
    const onStage = event.target.closest?.("#stageBody .slide-wrap");
    stopMotionPreview({ render: false });
    if (onStage) { event.preventDefault(); event.stopPropagation(); }
    requestAnimationFrame(() => { if (!state.motionPreview) renderStage(); });
  }, true);
  // The thumbnails keep the keys (Delete, ⌘C…) only while they have the focus: a click anywhere else takes it.
  document.addEventListener("pointerdown", (event) => { if (document.activeElement === $("filmstrip") && !event.target.closest?.("#filmstrip")) $("filmstrip").blur(); }, true);
  // 最終版: no keys that edit (moving between slides, the show, find, print, save and the like still work).
  // The places that edit the deck: their fields take no typing (other fields — search, dialogs — still do).
  const FINAL_LOCKED = "#stageBody, #inspector, #formatPane, #animPane, #commentPane, #a11yPane, #notesPane, #chatForm, #memoBox, #deckTitleInput";
  const FINAL_KEYS = new Set(["ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight", "PageUp", "PageDown", "Home", "End", "Escape", "Tab", "F5", "F1", "Shift", "Control", "Meta", "Alt", "?"]);
  const finalNow = () => !state.player && state.mode === "edit" && isFinal(state.deck);
  document.addEventListener("keydown", (event) => {
    if (!finalNow()) return;
    const locked = event.target?.closest?.(FINAL_LOCKED);
    if (!locked && (/^(INPUT|TEXTAREA|SELECT)$/.test(event.target?.tagName) || event.target?.isContentEditable || document.querySelector("dialog[open]"))) return;
    const meta = event.metaKey || event.ctrlKey;
    const k = event.key.toLowerCase();
    if (FINAL_KEYS.has(event.key) || (meta && ["c", "a", "f", "p", "s", "0", "=", "+", "-", "f1"].includes(k)) || (event.altKey && event.code === "KeyQ")) return;
    event.preventDefault();
    event.stopImmediatePropagation();
    toast("最終版です。編集するにはメッセージ バーの「編集する」を押してください");
  }, true);
  for (const type of ["dragstart", "contextmenu", "paste", "drop", "cut"]) {
    document.addEventListener(type, (event) => {
      if (!finalNow() || (!event.target?.closest?.(FINAL_LOCKED) && event.target?.closest?.("dialog, input, textarea, .rb-pop, .rb-search"))) return;
      event.preventDefault();
      event.stopImmediatePropagation();
    }, true);
  }
  document.addEventListener("keydown", (event) => {
    if (state.player) return; // the player has its own keys
    if (state.motionPreview?.brief && !["Shift", "Control", "Meta", "Alt"].includes(event.key)) stopMotionPreview();
    if (state.motionPreview && event.key === "Escape") { stopMotionPreview(); return; }
    if (state.motionPreview && ["ArrowRight", " ", "Enter"].includes(event.key) && !document.querySelector("dialog[open]")) { event.preventDefault(); state.motionPreview.advance(); return; }
    if (editor.keydown(event)) return;
    if ((event.key === "F5" || ((event.metaKey || event.ctrlKey) && event.key === "Enter")) && state.mode === "edit" && state.deck) {
      event.preventDefault();
      openPresenter(event.shiftKey ? state.selected : 0);
      return;
    }
    const typing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement?.tagName) || document.activeElement?.isContentEditable || document.querySelector("dialog[open]");
    const meta = event.metaKey || event.ctrlKey;
    if (meta && event.key.toLowerCase() === "k") { event.preventDefault(); openCommandPalette(); return; }
    // ⌥Q / Alt+Q: the ribbon's 検索 box (PowerPoint's Microsoft Search).
    if (event.altKey && !meta && event.code === "KeyQ" && state.mode === "edit" && state.deck && !typing) { event.preventDefault(); editorUi.focusSearch(); return; }
    // PowerPoint's window keys: ⌘F1 folds the ribbon, ⌘F finds, ⌘0 / ⌘＋ / ⌘− zoom the slide.
    if (state.mode === "edit" && state.deck && !document.querySelector("dialog[open]")) {
      if (meta && event.key === "F1") { event.preventDefault(); editorUi.toggleRibbon(); return; }
      if (meta && !event.shiftKey && ["f", "h"].includes(event.key.toLowerCase()) && !document.activeElement?.isContentEditable) { event.preventDefault(); openReplace(); return; }
      // フォント (⇧⌘F, as PowerPoint's Ctrl+Shift+F), also while typing.
      if (meta && event.shiftKey && event.key.toLowerCase() === "f" && state.view === "single" && !isFinal(state.deck) && editor.textState()) { event.preventDefault(); editorUi.fontDialog(); return; }
      if (!typing && shell.zoomKey(event)) { event.preventDefault(); return; }
      // PowerPoint's function keys: ⇧F9 the grid, ⌥F9 the guides, F7 the check of the words.
      if (!typing && event.key === "F9" && (event.shiftKey || event.altKey) && state.view === "single") { event.preventDefault(); const key = event.shiftKey ? "grid" : "guides"; editor.setView(key, !editor.state[key]); return; }
      if (!typing && event.key === "F7" && !meta && !event.shiftKey && !event.altKey) { event.preventDefault(); proofing.open(); return; }
      if (meta && event.key.toLowerCase() === "m" && !typing) { event.preventDefault(); openTypeDialog("insert"); return; }
      if (document.activeElement === $("filmstrip") && filmKey(event)) { event.preventDefault(); return; }
    }
    if (meta && event.key.toLowerCase() === "z" && state.mode === "edit" && !typing) {
      event.preventDefault();
      undoRedo(event.shiftKey ? "redo" : "undo");
      return;
    }
    // ⌘Y / Ctrl+Y: やり直し, or 繰り返し when there is nothing to redo, as in PowerPoint.
    if (meta && !event.shiftKey && event.key.toLowerCase() === "y" && state.mode === "edit" && !typing) {
      event.preventDefault();
      if (state.redo.length) undoRedo("redo"); else if (state.view === "single") editor.repeat();
      return;
    }
    if (!typing && event.key === "?") { event.preventDefault(); $("helpDialog").showModal(); return; }
    // ⌘P: the print dialog (layouts, range, preview), as PowerPoint's File → Print.
    if (meta && !event.shiftKey && event.key.toLowerCase() === "p" && state.deck && !state.player) { event.preventDefault(); printPdf(); return; }
    if (meta && event.key.toLowerCase() === "s") {
      event.preventDefault();
      if (state.deck) saveToLibrary();
      else { saveCurrent(); toast("入力内容は自動保存されています"); }
      return;
    }
    if (typing || state.mode !== "edit" || !state.deck) return;
    if (["ArrowDown", "ArrowRight", "PageDown"].includes(event.key)) { event.preventDefault(); select(state.selected + 1); }
    if (["ArrowUp", "ArrowLeft", "PageUp"].includes(event.key)) { event.preventDefault(); select(state.selected - 1); }
  });
  let lastFocusCheck = 0;
  window.addEventListener("focus", () => {
    if (state.codexImageAuthorized || Date.now() - lastFocusCheck < 15000) return;
    lastFocusCheck = Date.now();
    checkCodexStatus();
  });
}

function restore() {
  try {
    const saved = JSON.parse(localStorage.getItem(STORAGE.current) || "null");
    if (!saved) return;
    applyForm(saved.form);
    if (saved.outline?.slides?.length) { state.outline = saved.outline; renderOutlinePanel(); }
    if (saved.deck?.slides?.length) {
      state.deck = adoptDeck(normalizeDeck(saved.deck));
      state.selected = Math.min(saved.selected || 0, state.deck.slides.length - 1);
      state.historyId = saved.historyId || null;
      state.savedDeckId = saved.savedDeckId || null;
      $("deckTitleInput").value = state.deck.title;
      $("editTab").disabled = false;
      $("resumeEditBtn").classList.remove("hidden");
      useFonts([state.deck.theme]);
      setMode("edit");
      renderAll();
      scheduleMeasure(0);
      loadChat();
      ensureMedia().then((added) => {
        state.mediaChecked = true;
        lintMemo = { key: "", result: [] };
        if (added) thumbCache.clear();
        renderFilmstrip();
        renderStage();
        renderIssueSummary();
      });
    }
  } catch { /* ignore a corrupt autosave */ }
}

function reportClientError(message, where) {
  try {
    navigator.sendBeacon?.("/api/client-error", new Blob([JSON.stringify({ message, where, version: APP_VERSION })], { type: "application/json" }));
  } catch { /* reporting must never break the app */ }
}
window.addEventListener("error", (event) => reportClientError(event.message, `${event.filename}:${event.lineno}:${event.colno}`));
window.addEventListener("unhandledrejection", (event) => reportClientError(String(event.reason?.message ?? event.reason), "promise"));

bind();
try { const saved = localStorage.getItem(STORAGE.panel); setPanel(["form", "format", "anim", "comment", "a11y"].includes(saved) ? saved : "chat", { reveal: false }); } catch { setPanel("chat", { reveal: false }); }
updateBriefCount();
restore();
updateTopbar();
checkCodexStatus();
// A link to a shared deck (…?room=ID) joins the room; a link to an online presentation (…?watch=ID) watches it.
try { const room = new URL(location.href).searchParams.get("room"); if (room && /^[A-Za-z0-9_-]{16,40}$/.test(room)) coedit.join(room); } catch { /* file: */ }
try { const watch = new URL(location.href).searchParams.get("watch"); if (watch && /^[A-Za-z0-9_-]{16,40}$/.test(watch)) window.__hpptxWatch = online.watch(watch); } catch { /* file: */ }
