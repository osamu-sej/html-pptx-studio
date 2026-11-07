import { spawn } from "node:child_process";
import { EventEmitter } from "node:events";
import { existsSync, mkdirSync, readFileSync, realpathSync, statSync, unlinkSync, writeFileSync } from "node:fs";
import { isAbsolute, join, relative, sep } from "node:path";
import readline from "node:readline";

import { LOOK_ADVICE, LOOKS, maxSameLook } from "../public/layout-looks.mjs";
import { drillParents } from "./chat.mjs";
import { slideMeaning } from "./visual-relevance.mjs";
import { objectsSummary } from "./objects.mjs";
import { schemaSketch } from "./schema-sketch.mjs";

const DEFAULT_TIMEOUT_MS = 30_000;
const ICONS = Object.fromEntries(Object.entries(JSON.parse(readFileSync(new URL("../public/engine/icons.json", import.meta.url), "utf8"))).map(([key, value]) => [key, value.label]));

const TRANSIENT_ERRORS = new Set(["serverOverloaded", "internalServerError", "httpConnectionFailed", "responseStreamConnectionFailed",
  "responseStreamDisconnected", "responseTooManyFailedAttempts"]);

/** The slides the next part of a deck written in parts covers: [from, to] (1-based). */
function partRange(job) {
  const from = job.parts.done.length + 1;
  return [from, Math.min(job.parts.total, from + job.parts.size - 1)];
}

/** What to ask for the next part of a deck written in parts (the first comes after the whole request). */
export function partPrompt(job, { first = false, note = "" } = {}) {
  const { total, size, done } = job.parts;
  const [from, to] = partRange(job);
  const last = to === total ? "最後のスライドは type を closing にする。" : "";
  if (first) {
    // A new thread (the first part, or after an answer was cut off): the whole request, then what is already written.
    const written = done.length ? ["", `1〜${done.length}枚目は書き終えている（下の一覧）。話の流れ・用語・数値をこれにそろえ、繰り返さない。`,
      ...done.map((slide, index) => `${index + 1}. [${slide.type}] ${slide.title ?? ""}${slide.takeaway ? ` — ${slide.takeaway}` : ""}`)] : [];
    return [job.basePrompt, "", "# 書き方（何回かに分けて返す）",
      `資料全体は${total}枚（枚数の指示は資料全体の枚数）。1回の回答に入りきらないので、${size}枚ずつ何回かに分けて書く。`,
      ...written,
      `今回は全${total}枚のうち${from}〜${to}枚目だけを slideData に入れて返す（slideData 以外の項目もいつもどおり入れる）。続きはこのあと頼む。${last}`, note].filter(Boolean).join("\n");
  }
  return [note, `続けて、全${total}枚のうち${from}〜${to}枚目だけを slideData に入れて返す（slideData 以外の項目は前と同じでよい）。`,
    `これまでに書いたスライドと話の流れ・用語・数値をそろえ、書いたスライドは繰り返さない。${last}`].filter(Boolean).join("\n");
}

/** Codex stopped because the answer did not fit (the model's output limit). */
export const tooLong = (error) => /max_output_tokens/i.test(`${error?.message ?? ""} ${JSON.stringify(error?.codexErrorInfo ?? "")}`);

/** What a turn without the output format asks for instead (the format written out short, see schema-sketch.mjs). */
export function formatLines(schema) {
  return ["", "# 回答の形式",
    "回答は次の形の JSON オブジェクト1つだけにする（説明文・Markdown・コードブロックは付けない）。`名前?:` の項目は省いてよく、値のない項目は書かない。文字の長さは上の指示に従う。",
    schemaSketch(schema)].join("\n");
}

function errorKind(error) {
  const info = error?.codexErrorInfo;
  return typeof info === "string" ? info : info && typeof info === "object" ? Object.keys(info)[0] : "";
}

/** Codex's error in words the user can act on (the original message kept for support). */
export function friendlyCodexError(error) {
  const original = error?.message ?? "";
  const kind = errorKind(error);
  // Reported as a dropped stream, but nothing is wrong with the connection: the answer was too long.
  if (tooLong(error)) return `回答が長すぎて、Codexの1回の出力の上限で途中で切れました。枚数か文字量を減らすか、骨子を作ってから生成してください。（${original}）`;
  const plain = {
    usageLimitExceeded: "Codexの利用上限に達しました（接続しているChatGPTアカウントのプランの上限）。上限が回復するまで待つか、ChatGPTの利用状況を確認してください。",
    unauthorized: "Codexとの接続が切れました。画面右上から接続し直してください。",
    contextWindowExceeded: "素材が長すぎて一度に処理できませんでした。素材を短くするか、分けて作ってください。",
    sessionBudgetExceeded: "この依頼の処理量が上限を超えました。素材を短くするか、枚数を減らしてください。",
  }[kind] ?? (TRANSIENT_ERRORS.has(kind) ? "Codex側が混み合っているか、通信が不安定です。少し待ってからもう一度お試しください。" : "");
  if (!plain) return original || "Codexでエラーが発生しました。";
  return original ? `${plain}（${original}）` : plain;
}

function extractItemText(item) {
  if (!item) return "";
  if (typeof item.text === "string") return item.text;
  if (typeof item.content === "string") return item.content;
  if (Array.isArray(item.content)) {
    return item.content
      .map((part) => part?.text ?? part?.content ?? "")
      .filter(Boolean)
      .join("\n");
  }
  return "";
}

export function parseDeckResult(text) {
  const trimmed = String(text ?? "")
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();
  try {
    return JSON.parse(trimmed);
  } catch (error) {
    // Without the output format the answer may come with a sentence around the JSON.
    const start = trimmed.indexOf("{");
    const end = trimmed.lastIndexOf("}");
    if (start > 0 || (start === 0 && end < trimmed.length - 1)) {
      if (end > start) return JSON.parse(trimmed.slice(start, end + 1));
    }
    throw error;
  }
}

export function omitNullObjectValues(value) {
  if (Array.isArray(value)) return value.map(omitNullObjectValues);
  if (!value || typeof value !== "object") return value;

  return Object.fromEntries(
    Object.entries(value)
      .filter(([, child]) => child !== null)
      .map(([key, child]) => [key, omitNullObjectValues(child)]),
  );
}

export function normalizeCodexOutputSchema(value) {
  if (Array.isArray(value)) return value.map(normalizeCodexOutputSchema);
  if (!value || typeof value !== "object") return value;

  const normalized = Object.fromEntries(
    Object.entries(value).map(([key, child]) => [key, normalizeCodexOutputSchema(child)]),
  );
  if (Array.isArray(normalized.items)) {
    // Codex accepts only an object for array items; it does not support tuple
    // syntax or prefixItems. Exact validation still runs with Zod afterwards.
    const tupleItems = normalized.items;
    const firstItem = tupleItems[0] ?? {};
    const allItemsMatch = tupleItems.every((item) => JSON.stringify(item) === JSON.stringify(firstItem));
    normalized.items = allItemsMatch ? firstItem : {};
    normalized.minItems = Math.max(normalized.minItems ?? 0, tupleItems.length);
    normalized.maxItems = Math.min(normalized.maxItems ?? Infinity, tupleItems.length);
    delete normalized.prefixItems;
  }
  return normalized;
}

const DENSITY_GUIDE = {
  light: { label: "少なめ", rule: "1枚2〜3要素、説明は各15〜30字。大きな結論と余白を優先する" },
  standard: { label: "標準", rule: "1枚3〜4要素、説明は各25〜40字。結論・根拠・示唆を揃える" },
  rich: { label: "多め", rule: "1枚4〜5要素、説明は各30〜50字。背景・根拠・具体例・実行条件・リスクまで素材の範囲で掘り下げる" },
};

const PERSONA = "あなたはビジネス資料に強いプレゼンテーション設計者です。資料は動きのあるHTMLのプレゼンテーションとして作られ、そのままPowerPointに書き出して発表することも、画面共有（Zoomなど）で発表することもあり、項目が順に現れる・写真がゆっくり動く・クリックで根拠が開く・切り口を切り替える・条件を動かして試算するといった動きを使えます。";

const LAYOUT_GUIDE = [
  "利用可能なtypeと主要フィールド（すべての本文typeに title と takeaway を入れる）:",
  "- content: points[]（「要素名：説明」形式の2〜5項目は番号付きの行で見やすく表示される）、twoColumn、columns[[],[]]",
  "- agenda: items[]（「章名：概要」形式）",
  "- compare / beforeAfter: leftTitle, rightTitle, leftItems[], rightItems[]（beforeAfterは左右を「観点：内容」の同じ観点・同じ数でそろえると対照表になる）",
  "- process（横並び2〜5工程） / processList（縦並び最大8工程）: steps[]（「工程名：説明」形式）",
  "- flowChart: flows[{steps[]}]（1〜2本）",
  "- timeline: milestones[{label, date, state: done|next|todo}]",
  "- diagram: lanes[{title, items[]}]（2〜4レーン、各1〜4項目）",
  "- cycle: items[{label, subLabel}]、centerText",
  "- cards / headerCards / bulletCards / grid2x2 / headerTwoColumn / headerThreeSummary(summary付き) / stepUp / triangle / venn / roadmap / orgChart(root付き) / checklist(items[].done) / matrix(xLabel, yLabel。右上が狙う領域) / swot(強み・弱み・機会・脅威の順): items[{title, desc, icon}]",
  "- gantt: periods[]（期間ラベル）、items[{title, desc, start(0始まりの期間番号), span(期間数)}]、now（任意。「いま」の位置。1.5＝2つ目の期間の真ん中。今の時点が分かるときだけ）",
  "- waterfall（合計の内訳・増減の要因）: unit（単位）、items[{label, value（数値。減少はマイナス）, total（合計・起点の棒はtrue）}]（2〜8本。素材にある数値だけ）",
  "- logicTree（原因分解・論点分解）: root（課題）、branches[{title（10字程度）, items[]（葉。各16字以内・最大3個）, highlight（手を打つ枝だけtrue）}]（2〜4本）",
  "- table: headers[], rows[][]（最大5列×6行、セルは14字以内）",
  "- kpi: items[{label, value, change, status: good|bad|neutral}]（最大4個。1個なら数字を大きく、%ならゲージになる）",
  "- dashboard: items[{label, value, change}] と image{chartType, data}",
  "- statsCompare: leftTitle, rightTitle, stats[{label, leftValue, rightValue, trend: up|down|neutral}]",
  "- pyramid / funnel: levels[{title, description}]（3〜5段）",
  "- quote: text, author ／ faq: items[{q, a}]（最大4問） ／ executiveSummary: conclusion, items[{title, desc}]（2〜3個）, action",
  "- imageText: points[]、image={chartType: bar|line|donut|multi-line|stacked-bar|100-stacked-bar, data:{title, items[{label, value}] または xAxisLabels + series[{label, values[]}]}}",
  "  - image.chartType=rank（切り口で並び替わる順位）: data:{title, unit, highlight（主役の項目名）, views[{label（切り口の名前。例：全体・一次請け、2024年度・2025年度）, items[{label, value}]}]（2〜4個）}。発表中は切り口のボタンで全部の棒が並び替わる。同じ項目名を各切り口で使う",
  "  - image.chartType=shift（前後の差）: data:{title, unit, beforeLabel, afterLabel, items[{label, before, value（後の値）}]}（2〜6本）。発表中は棒が前から後へ形を変え、減った（増えた）分が点線で残る",
  "- simulator（原因と結果・試算）: inputs[{label, value（いまの値）, min, max, step, unit}]（1〜3個。上から a・b・c）、formula（a・b・c と + - × ÷ ( ) と数字だけ。例: a × b × c ÷ 100）、resultLabel、resultUnit、compareLabel と compareValue（目標・現状など比べる値。任意）。発表中はスライダーを動かすと結果と差が計算し直され、式が画面に出る。素材の数値を初期値にし、試算であることを takeaway か notes に書く",
  "- gap（不足と打ち手）: unit、targetLabel・target（目標）、currentLabel・current（現状）、measures[{title, value（上乗せ量）, desc}]（1〜5個）。発表中はクリックのたびに打ち手が1つずつオンになり、目標までの不足が埋まっていく（クリックでオン・オフもできる）",
  "- hero（全面写真に大きな一文。章の始まりや印象づけたい場面に。資料全体で1〜2枚まで）: title（一文で言い切る）, takeaway（補足）, visualAsset（写真）",
  "- statement（大きな一文だけのスライド。問いかけ・決意・転換点に）: title（小見出し）, text（40字以内の一文。**語句** で1か所強調）",
  "- title: title, subtitle, date ／ section: title, takeaway ／ closing: title, message（次のアクション。2〜3個なら改行で区切ると番号付きの列になる）",
  "動き（HTML発表のときだけ効く。PowerPointの図形アニメーションとは別物）:",
  "- animation: 項目の出し方。auto（レイアウトに合わせて自動。通常はこれ・省略可）/ click（クリックのたびに1項目ずつ。工程・計画を話しながら見せる）/ spotlight（順に注目。全部見せたまま、クリックのたびに1項目を強調しほかを薄くする。比較・選択肢・論点を1つずつ話すときに）/ cascade（順番に自動）/ fade（まとめて）/ none",
  "- photoMotion: 写真の動き。none / zoom（ゆっくり寄る）/ pan（ゆっくり横に流れる）/ float（ふわっと漂う）/ parallax（マウスに合わせて奥行き）/ reveal（幕が開くように現れる）/ drift（斜めにゆっくり流れる）/ tilt（ゆっくり傾く）。写真のあるスライドだけ",
  "- kinetic（モーショングラフィック：大きな文字の動き）: auto（省略可。表紙・章扉・statement・hero・closing は資料の設定で自動的に動く）/ none / mask（下から立ち上がる。上品で汎用）/ words（ことばごとにぼかしから浮かぶ）/ chars（1文字ずつ弾む。研修・キックオフ向け）/ type（タイプライター。問いかけ・引用に）/ scramble（文字が入れ替わって決まる。テクノロジー・データの話題に）/ wave（波打って並ぶ。明るい話題に）/ zoom（手前から迫る。決意・宣言に）/ flip（1文字ずつめくれる。発表・お披露目に）/ slide（横から滑り込む。スピード感・変化に）。本文スライドに付けるとタイトルが動く。資料全体で1〜2種類にそろえ、本文スライドには多用しない",
  "- backdrop（モーショングラフィック：スライドの後ろで動く図形）: auto（省略可）/ none / particles（粒子が昇る）/ waves（波が流れる）/ grid（グリッドと走査線。データ・DX）/ orbits（軌道を回る。AI・テクノロジー・全体像）/ gradient（色がゆらぐ。ビジョン・未来）/ lines（線を光が流れる。つながり・流れ・データ連携）/ shapes（図形が漂う。研修・アイデア）/ confetti（紙吹雪。表彰・達成・お祝い）/ network（点と線がつながる。組織・連携・ネットワーク）/ ripple（波紋が広がる。影響・波及・浸透）/ stars（星がまたたく。夢・長期ビジョン）/ rays（光の筋が回る。発表・始動・幕開け）。表紙・章扉・statement・closing に向く。本文の多いスライドには付けない。資料全体で同じ種類にそろえる",
  "- entrance（このスライドだけの登場のしかた。通常は省略して資料の設定に任せる）: auto / none / rise（下から浮かぶ）/ fade / blur / pop / slide（左から滑り込む）/ zoom（手前から迫る）/ flip（めくれる）/ wipe（幕が開く）/ drop（弾んで落ちる）/ spin（部品が3Dで1回転半しながら手前に現れる。見せ場向け）/ swirl（渦を巻いて遠くから飛び込む。見せ場向け）。見せ場のスライドに使う（下の「見せ場」）",
  "- emphasis（このスライドの **語句** の強調の見せ方。通常は省略）: auto / marker（マーカーを引く）/ underline（下線）/ circle（手書きの丸で囲む）/ box（枠で囲む）/ glow（光らせる）/ none（色だけ）",
  "- transition（このスライドへ切り替わるときだけの動き。通常は省略して資料の設定に任せる）: auto / none / fade / slide / zoom / morph / wipe / circle / push（下から押し上げる）/ flip（カードのように裏返る）/ dive（奥へ飛び込む）/ blinds（ブラインドが開く）/ curtain（幕が中央から開く）。章の変わり目・見せ場への切り替わりに使う",
  "- drillOf（深掘りページ）: このスライドを本編の流れから外し、直前の本編スライドの項目（items[1]、steps[0] など details の target と同じ書き方）をクリックしたときだけ開くページにする。本編の番号には数えず、発表中は Esc・← で元のスライドに戻る。1つの項目に深掘りページは1枚まで、深掘りページからさらに深掘りはしない。頼まれたときだけ作り、既存の drillOf のあるスライドは消さない・動かさない・drillOf を外さない",
  "- details（その項目をクリックすると開く補足カード）: [{target（項目の配列名と0始まりの番号。例: items[0]、steps[2]、milestones[1]、levels[0]、points[1]、branches[0]、leftItems[0]、stats[0]、rows[0]、measures[0]。キーメッセージの根拠なら takeaway）, title（任意・20字以内）, text（判断を1〜2文・120字以内）, rows（任意。内訳 [{label, value}] 最大8行。value は単位つきの文字）, source（任意。出所）, note（任意。前提・注記）}]。rows・source・note のどれかがあると、右から根拠パネルとして開く。スライドの本文は短いまま、話しながら見せたい根拠・具体例・内訳・数値の出所を書く（素材にない数値は書かない）。既存の details は頼まれない限りそのまま残す",
  "- source（本文スライドの出所。任意・40字程度）: 素材に出所・調査名・時点があるスライドだけ。ページ下とグラフの吹き出しに出る",
  "動く資料（1枚1操作）: 見出しが言いたいことの構造で、そのページで見る人が自分の手で確かめる操作を1つだけ選ぶ（動きは思いつきで付けない。見出しと関係ない動きは入れない）",
  "- 結論と根拠 → 押すと根拠が出る（details に rows・source を付けた根拠パネル。数字のある本文スライドのほぼ全部に付けてよい。target=takeaway ならキーメッセージの根拠）",
  "- 切り口（分類・順位） → 切り替えると並び替わる（imageText の chartType=rank）",
  "- 差分（前後・比較） → 形が変わり、差が残る（imageText の chartType=shift）",
  "- 原因と結果（試算・感度） → 条件を動かすと計算し直す（simulator）",
  "- 不足と打ち手 → オン・オフで不足が埋まる（gap）",
  "- 流れと時間（工程・計画） → 順に組み上がる（process・timeline・gantt などの animation=click。gantt は now で「いま」を示す）",
  "- 1ページの操作は1つだけ（simulator・gap・rank・shift のスライドに click や spotlight を重ねない）。素材の数字がないのに simulator・gap・rank・shift を作らない",
];

const CAPACITY_GUIDE = [
  "文字数の上限（1920×1080のスライドに大きな文字で収めるための目安。超えると自動で修正を依頼します）:",
  "- title 24字以内（1行）、takeaway 45字以内、subtitle 40字以内",
  "- 箇条書き1項目 45字以内・1枚5項目まで",
  "- カードの desc: 2列なら60字、3列なら40字、4列なら25字以内",
  "- process の説明: 3工程なら35字、4工程なら25字、5工程なら18字以内",
  "- faq の a 55字以内、quote の text 90字以内、statement の text 40字以内、closing の message 60字以内",
  "- 1枚の本文（タイトルとtakeawayを除く）は240字以内、表は320字以内。超えるなら要素を絞るか2枚に分ける（書ききれない補足は details に回す）",
];

// How professional decks read: the renderer handles the look, the content has to leave it room.
const WRITING_GUIDE = [
  "書き方（プロの資料の作法）:",
  "- takeaway はタイトルの言い換えではなく、そのスライドの結論・示唆を一文で言い切る（例：×「売上の推移」 ○「新商品が伸び、前年比108%に回復した」）",
  "- title は論点（24字以内・1行）、takeaway は主張の完全な文。takeaway だけを縦に読んで話が通るようにする",
  "- takeaway の中で最も伝えたい語句（数値・結論）を1か所だけ **語句** で囲む（その語句がアクセント色で強調される。最大2か所）",
  "- 同じスライドの項目名は長さと文体をそろえる（体言止めならすべて体言止め）。項目は3〜4個を基本にし、6個を超えるなら分ける",
  "- 数値は単位までそろえ、比較する数値は同じ単位・同じ桁で書く。記号の飾り（★・！・絵文字）は使わない",
  "- 証拠の型を内容で選ぶ：数字1つが主役→kpi（1項目。%ならゲージ）／指標を複数→kpi・dashboard／項目の大小比較→imageText（bar）／推移→imageText（line）／構成比→imageText（donut）／合計の内訳・増減の要因→waterfall／原因分解→logicTree／現状と目指す姿→beforeAfter／時系列の計画→roadmap・timeline・gantt／繰り返す業務→cycle／声→quote／決意・問い→statement",
  "- 見た目のグループ（下の一覧）を散らす。アイコン付きカードの羅列で図解の代わりにしない",
  "- 写真（visualAsset）は表紙・hero・写真が内容を補強するスライドだけ。図解・表・グラフのスライドには付けない",
  "- デザイン（テーマ・配色・フォント）はユーザーが選ぶ。あなたは変えない。ページ番号・資料名はテーマが描くので、タイトルや本文に書かない。タイトルは1行に入るよう24字以内を守る。表紙の date には作成日（例：2026年9月）を入れる",
  "- 強調は **語句** の1〜2か所に絞る（テーマのアクセント色になる）。accent は使わない",
];

// Different layouts can look the same on screen (a timeline, a process and a roadmap are all a line of points),
// so variety is asked for by look. The same groups drive the automatic check in server.js and the studio.
const LOOK_GUIDE = [
  "見た目のグループ（type の名前が違っても、画面では同じ形に見えるレイアウトの集まり。単調かどうかは type ではなくこのグループで判断する）:",
  ...Object.values(LOOKS).map((look) => `- ${look.label}: ${look.types.join("・")}`),
  "- 同じグループは本文の4枚に1枚まで（本文6枚なら2枚、10枚なら3枚）。同じグループを隣り合わせない。本文が6枚以上なら4グループ以上を使う",
  "- 横に並ぶ流れは、順番や時期が本当にあるときだけ使う：timeline・roadmap・gantt は日付・期間を書けるとき、process は聞き手がその順にやる手順、diagram は複数の担当の間で仕事が受け渡されるとき。歴史・今後の方向性・備えでも、時期を書けないなら流れにしない",
  "- 内容に合う形の選び方:",
  ...LOOK_ADVICE.map((line) => `  - ${line}`),
];

/** What to tell the AI when its layouts look monotonous (see varietyIssues in public/layout-looks.mjs). */
export function varietyRepairLines(issues, { outline = false } = {}) {
  return [
    "スライドの見た目が単調です（type の名前が違っても、画面では同じ形に見えます）。",
    ...issues.map((issue) => `- ${issue.message}`),
    "直し方:",
    "- 「〜が○枚あります」は、内容に最も合う枚数だけその形で残し、残りのスライドの type を別の見た目のグループに変える",
    "- 「続いています」は、どちらか内容に合わないほうの type を別のグループに変える",
    "- 「○種類だけです」は、数字・図形・大きな一文・表などのグループを使って種類を増やす",
    `- 新しい type は内容で選ぶ（${LOOK_ADVICE.join("／")}）。見た目を変えるためだけに内容と合わない形にしない`,
    outline
      ? "- 各スライドの title・takeaway の主旨、枚数、順序は変えない（type と、それに合わせた content のメモだけ直す）"
      : "- type を変えたスライドは、新しい type のフィールドで中身を書き直す（主張・数値・notes・details は保つ）。枚数・順序・type を変えないスライドはそのままにする",
  ];
}

const VISUAL_GUIDE = [
  "- visualAsset（内蔵写真）: ai=テクノロジー / aiWorkflow=AIワークフロー / promptDesign=設計・思考 / businessWorkshop=研修・会議 / businessEtiquette=接客・対話 / executiveDecision=経営判断 / storeOperations=店舗・現場 / dataInsight=データ分析 / transformationRoadmap=変革・計画 / customerExperience=顧客体験",
  `- items[].icon（カード類のアイコン）に使える名前：${Object.entries(ICONS).map(([k, v]) => `${k}=${v}`).join("、")}`,
  "- 写真・アイコンを選ぶときは、そのスライドの結論と具体的な本文に写る対象が一致するものだけ。資料全体のテーマだけで選ばない。合う素材がなければ追加しない",
];

export function deckNarrativeLines(deck, selected = []) {
  const slides = deck.slides ?? deck.slideData ?? [];
  const targets = new Set(selected);
  return slides.map((slide, index) => {
    const claim = slide.takeaway || slide.conclusion || slide.message || slide.text || slide.title || "";
    const body = slideMeaning(slide).replace(claim, "").trim();
    const parent = drillParents(slides)[index];
    const role = parent != null ? `深掘り（${parent + 1}枚目の ${slide.drillOf} から開く。本編外）` : index === 0 ? "導入" : index === slides.length - 1 ? "結論・依頼" : "論点";
    const detail = targets.has(index) || targets.has(index - 1) || targets.has(index + 1) ? body.slice(0, 260) : body.slice(0, 100);
    return `${index + 1}枚目 [${role}/${slide.type}] ${claim}${detail ? `｜根拠・内容: ${detail}` : ""}`;
  });
}

/** For a deck shown outside the company, a reminder of what must not leak or confuse. */
function audienceLines(audience) {
  return /社外|取引先|お客様|顧客|パートナー|外部/.test(String(audience ?? ""))
    ? ["- 社外向けの資料なので、社内の略語・部署内の呼び名・社外秘の数値や計画は使わず、相手にとっての価値・メリットから書く"]
    : [];
}

/** The deck's standing instructions, repeated in every AI request. */
function memoLines(deck) {
  const memo = String(deck?.memo ?? "").trim();
  return memo ? ["", "この資料の前提条件（必ず守る）:", memo, ""] : [];
}

// Objects people placed by hand and 白紙 pages are theirs: the AI sees a summary and leaves them alone.
const OBJECTS_RULE = "- 「[自由配置: …]」はユーザーが手で置いた図形・文字・画像（アニメーションを含む）。あなたは変えられず、どんな変更でもそのまま残る。本文はそれと重ならない内容にする。type が blank（白紙）のスライドはユーザーが自由に作ったページなので、頼まれない限り replace・delete しない";

const THEME_LINE = "デザイン: ユーザーが選んだテーマのまま（テーマ・色は変えない）。切り替え: none / fade / slide / zoom / morph（見出しがつながって動く） / wipe（色の帯が横切る） / circle（クリックした所から円が広がる） / push（下から押し上げる） / flip（カードのように裏返る） / dive（奥へ飛び込む） / blinds（ブラインドが開く） / curtain（幕が中央から開く）";

export function buildChatPrompt({ deck, message, history = [], current = 0, focus = [], attachment = null }) {
  const slides = withoutImageData(deck.slides ?? []);
  const total = slides.length;
  const talk = history.slice(-10).map((turn) => `${turn.role === "user" ? "ユーザー" : "あなた"}: ${String(turn.text).slice(0, 1500)}`);
  const targets = [...new Set(focus)].filter((index) => index < total).map((index) => index + 1);
  return [
    `${PERSONA}ユーザーと会話しながら、資料を少しずつ良くしていきます。`,
    "最終回答は指定されたJSONスキーマに一致するJSONだけを返してください。Web検索、ファイル操作、外部ツールは不要です。",
    "",
    "回答のルール:",
    "- reply: 日本語で2〜4文。何をどう変えたかを具体的に書く（例：「3枚目のカードを3つに絞り、各項目にクリックで開く補足を付けました」）。質問・相談には答えだけを書く。変えなかった理由や確認したいことがあれば1つだけ聞いてよい",
    "- operations: 資料を変える必要があるときだけ入れる。質問・感想・相談なら空配列",
    "  - replace: slide=元の枚数（1始まり）、content=そのスライドの完全なJSON（一部だけでなく全フィールド）",
    "  - insert: slide=元の何枚目の後ろに入れるか（1始まり）、content=新しいスライドの完全なJSON",
    "  - delete: slide=元の枚数",
    "  - 枚数は常に「今の資料」の番号で書く（他の操作で番号がずれることは考えなくてよい）",
    "- order: スライドの順番を入れ替えるときだけ、元の枚数を新しい順に並べた配列（削除したものは除く）。入れ替えないなら省略",
    "- deckTitle: 資料名の変更を頼まれたときだけ",
    "- transition: スライドの切り替えの変更を頼まれたときだけ（下の一覧から選ぶ）。デザイン（テーマ・色）はユーザーが「デザイン」タブで選ぶもので、あなたは変えない。「テーマを変えて」「色を変えて」と頼まれたら、デザイン タブのテーマで変えられることを伝え、あなたは動き（motion・切り替え）やレイアウトで応える",
    "- motion: 資料全体の動き（モーショングラフィック）の変更を頼まれたときだけ。変えるものだけ入れる。entrance（全スライドの登場のしかた）/ hover（項目にマウスを乗せたときの反応: lift・focus・tilt・glow・zoom）/ kinetic（表紙・章扉などの大きな文字の動き）/ backdrop（表紙・章扉・statement・最後のスライドの後ろで動く図形）/ emphasis（**語句** の強調の見せ方）/ draw（アイコン・線・マーカーを描くように見せるか）。1枚だけなら、そのスライドの entrance・kinetic・backdrop・emphasis・transition を変える",
    "- suggestions: ユーザーが次に頼みそうな改善を最大3つ。20字以内の依頼文（例：「結論を1枚目に寄せて」「工程をクリックで1つずつ出して」）",
    "- 頼まれていないスライドは変えない。直すときは必要最小限にする",
    "- 変更前に、資料の目的→各スライドの役割→対象スライドの主張と根拠→前後とのつながりを確認する。本文を読まずに見出しだけで判断しない",
    "- 指示が曖昧でも、現在のスライドと直近の会話から対象を特定する。明確な対象がない場合は勝手に広げず、質問する",
    "- 「動きをつけて」「クリックで詳しく」「写真を動かして」などは animation・details・photoMotion で応える。「モーショングラフィック」「文字を動かして」「背景を動かして」「もっと派手に」は kinetic・backdrop・entrance（1枚なら各スライド、全体なら motion）で応える。「1つずつ注目させて」は animation の spotlight、「強調を丸で囲んで」などは emphasis、「このスライドだけ切り替えを変えて」はそのスライドの transition で応える。「クリックで確かめられるように」「動く資料にして」は、見出しの構造に合う操作を1つ（根拠パネル・rank・shift・simulator・gap・click）で応える。「根拠を出して」「内訳を見せて」は details の rows・source、「条件を変えて試算したい」は simulator、「打ち手で埋まるのを見せたい」は gap、「切り口を切り替えたい」は chartType=rank、「前後の差を見せたい」は chartType=shift。動画やLottieアニメーションはユーザーが画面の「写真・動画・アニメーション」から追加するもので、あなたは入れられない",
    "- 写真は、ユーザーが明示的に求めたときだけ新規選択・変更する。見た目を良くする依頼でも、無関係な素材を足さずレイアウト・余白・情報の強弱を優先する",
    "- 1枚目は title、最後は closing のままにする。表紙と最後は削除・移動しない",
    "- 「〇〇を深掘りするページを作って」「クリックで詳しいページに飛べるように」と頼まれたら、元のスライドの後ろに insert し、content に drillOf（元のスライドの項目。例: items[1]）を入れる。本文はその項目の背景・内訳・具体例・根拠で、元のスライドの繰り返しにしない。drillOf のあるスライド（深掘りページ）は本編の流れに入らないので、順番の入れ替えや枚数の話では数えない",
    "- 入力にない数値や事実は作らない。既存の数値は変えない。「[画像あり]」「[写真・動画あり]」の値はそのまま残す",
    OBJECTS_RULE,
    "- 新しい画像は生成できない。既存の写真を選ぶことと画像生成を混同せず、生成したと説明しない",
    "- 本文スライドには結論を一文で言い切る takeaway を入れる。文字数の上限を守る",
    ...memoLines(deck),
    `資料: ${deck.title ?? ""}（対象者: ${deck.audience || "未指定"}、目的: ${deck.purpose || "未指定"}、全${total}枚、デザイン: ${deck.theme || "clarity"}、切り替え: ${deck.transition || "fade"}、登場: ${deck.motion?.entrance || "rise"}、大きな文字の動き: ${deck.motion?.kinetic || "mask"}、背景の動き: ${deck.motion?.backdrop || "none"}、強調: ${deck.motion?.emphasis || "marker"}）`,
    `ユーザーが今見ているスライド: ${current + 1}枚目${targets.length ? `／ユーザーが指定したスライド: ${targets.map((n) => `${n}枚目`).join("、")}` : ""}`,
    "「この1枚」「このスライド」は今見ているスライドを指す。",
    "資料全体の骨子（各枚の役割・主張・根拠。対象の前後は詳しく示す）:",
    ...deckNarrativeLines(deck, targets.length ? targets.map((n) => n - 1) : [current]),
    "",
    talk.length ? "これまでの会話:" : "",
    ...talk,
    "",
    `ユーザーの新しい依頼: ${message}`,
    attachment?.text ? `\nユーザーが添付した資料「${attachment.name || "添付"}」の内容（依頼に必要な部分だけを使う。数値はこの資料のものを正とする）:\n${attachment.text}` : "",
    "",
    "今の資料（スライドごとのJSON）:",
    ...slides.map((slide, index) => `${index + 1}枚目: ${JSON.stringify(slide)}`),
    "",
    ...LAYOUT_GUIDE,
    ...VISUAL_GUIDE,
    THEME_LINE,
    "",
    ...CAPACITY_GUIDE,
    ...WRITING_GUIDE,
    ...LOOK_GUIDE,
  ].filter((line) => line !== "").join("\n");
}

/** Slides that have fully arrived in a streamed JSON answer ("slideData": [{…}, {…}, …). */
export function partialItems(text, key = "slideData") {
  const start = text.indexOf(`"${key}"`);
  if (start < 0) return [];
  let i = text.indexOf("[", start);
  if (i < 0) return [];
  const items = [];
  let depth = 0;
  let inString = false;
  let escaped = false;
  let objectStart = -1;
  for (i += 1; i < text.length; i += 1) {
    const c = text[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (c === "\\") escaped = true;
      else if (c === '"') inString = false;
      continue;
    }
    if (c === '"') inString = true;
    else if (c === "{") { if (depth === 0) objectStart = i; depth += 1; }
    else if (c === "}") {
      depth -= 1;
      if (depth === 0 && objectStart >= 0) {
        try { items.push(JSON.parse(text.slice(objectStart, i + 1))); } catch { /* not complete yet */ }
        objectStart = -1;
      }
    } else if (c === "]" && depth === 0) break;
  }
  return items;
}

/** The part of a JSON string value written so far ("reply": "…), for showing the answer while it streams. */
export function partialString(text, key) {
  const match = new RegExp(`"${key}"\\s*:\\s*"`).exec(text);
  if (!match) return "";
  let out = "";
  for (let i = match.index + match[0].length; i < text.length; i += 1) {
    const c = text[i];
    if (c === '"') break;
    if (c !== "\\") { out += c; continue; }
    const next = text[i + 1];
    if (next === undefined) break;
    if (next === "u") {
      const hex = text.slice(i + 2, i + 6);
      if (hex.length < 4) break;
      out += String.fromCharCode(parseInt(hex, 16));
      i += 5;
      continue;
    }
    out += { n: "\n", t: "\t", r: "", b: "", f: "" }[next] ?? next;
    i += 1;
  }
  return out;
}

function outlineLines(outline) {
  if (!outline?.length) return [];
  return [
    "",
    "確定した構成（枚数・順番・各スライドのtype・主旨を必ず守る。内容は素材から肉付けする）:",
    ...outline.map((item, index) => `${index + 1}. [${item.type}] ${item.title}${item.takeaway ? ` — ${item.takeaway}` : ""}${item.content ? `（内容: ${item.content}）` : ""}`),
  ];
}

export function buildOutlinePrompt({ brief, audience, purpose, tone, settings = {}, outline, instruction }) {
  const revising = Boolean(outline?.length);
  const bodyCount = Math.max(1, (revising ? outline.length : settings.slideCount ?? 8) - 2);
  return [
    PERSONA,
    revising
      ? "次の資料構成（骨子）を、ユーザーの指示に沿って直してください。指示に関係しない部分は変えないでください。"
      : "本番のスライドを作る前に、資料の骨子（各スライドのtype・タイトル・結論・中身のメモ）を提案してください。",
    "最終回答は指定されたJSONスキーマに一致するJSONだけを返してください。Web検索や外部ツールは不要です。入力素材にない数値や事実は作らないでください。",
    "",
    `対象者: ${audience || "未指定"} ／ 目的: ${purpose || "未指定"} ／ トーン: ${tone || "標準ビジネス"}`,
    `枚数の目安: ${settings.slideCount ?? 8}枚（内容に合わせて前後2枚まで調整してよい。表紙と最後を含む）`,
    "",
    "骨子のルール:",
    "- slides[0] は type=title、最後は type=closing",
    "- 2枚目以降で結論を先に示す。各スライドは1メッセージに絞る",
    "- title は24字以内、takeaway はそのスライドの結論を45字以内で言い切る、content はそのスライドに載せる中身のメモ（80字以内、素材の数値をそのまま使う）",
    "- type は内容に合う専門レイアウトを選ぶ（手順はprocess、対比はcompare・beforeAfter、定量はkpi・imageText（グラフ）・dashboard、日付のある計画はroadmap・gantt・timeline、層・段階はpyramid・stepUp、関係はvenn・triangle・cycle、やること・ポイントはcards・checklist・grid2x2、内訳・増減はwaterfall、原因分解はlogicTree、印象づけたい一文はstatement・hero など）",
    `- 見た目を散らす：本文${bodyCount}枚なら、同じ見た目のグループ（下の一覧）は${maxSameLook(bodyCount)}枚まで、隣り合わせない。とくに timeline・roadmap・process・diagram・flowChart・gantt はどれも「横に並ぶ流れ」に見えるので、合わせて${maxSameLook(bodyCount)}枚まで`,
    "- reply: 骨子の狙いを1〜2文で説明する（修正時は何を変えたか）",
    ...audienceLines(audience),
    "",
    ...LOOK_GUIDE,
    ...(revising ? ["", `ユーザーの指示: ${instruction || "より良くする"}`, "", "現在の骨子:", ...outline.map((item, index) => `${index + 1}. [${item.type}] ${item.title} — ${item.takeaway ?? ""}（${item.content ?? ""}）`)] : []),
    "",
    `素材・依頼内容:\n${brief}`,
  ].join("\n");
}

export function buildDeckPrompt(input) {
  const settings = input.settings ?? {};
  const density = DENSITY_GUIDE[settings.textDensity] || DENSITY_GUIDE.rich;
  return [
    `${PERSONA}入力内容からスライドの構成JSONを作成してください。`,
    "最終回答は指定されたJSONスキーマに一致するJSONだけを返してください。Markdownや説明文は不要です。",
    "Web検索、ファイル操作、外部ツールは不要です。入力素材にない数値や事実を作らないでください。",
    "",
    `資料タイトル候補: ${input.deckTitle || "入力内容から決定"}`,
    `対象者: ${input.audience || "未指定"}`,
    `目的: ${input.purpose || "未指定"}`,
    `トーン: ${input.tone || "標準ビジネス"}`,
    `指定枚数: ${settings.slideCount}枚（slideDataを必ずこの枚数にする）`,
    `文字量: ${density.label}（${density.rule}）`,
    "",
    "構成ルール:",
    "- 1枚目はtitle、最後はclosingにする",
    "- 2枚目以降で結論を先に示す（6枚以上ならexecutiveSummaryかkpiを序盤に置く）。各スライドは1メッセージに絞る",
    "- 本文スライドには、そのスライドの結論を一文で言い切る takeaway を必ず入れる。見出しだけで終わらせず、各要素に理由・意味・具体像を含める",
    "- 内容に合う専門レイアウトを使い、見た目のグループ（下の一覧）を散らす。手順はprocess、対比はcompare、定量はkpi/statsCompare/グラフ、内訳・増減はwaterfall、原因分解はlogicTree、日付のある計画はroadmap/gantt/timeline、層・段階はpyramid/stepUp、関係はvenn/triangle/cycle、やること・ポイントはcards/checklist/grid2x2",
    "- 素材に比較可能な数値が2点以上あれば、最低1枚をimageTextのimage（グラフ）で可視化する。数値がない場合は捏造せず図解で構造を見せる",
    "- 10枚以上の資料では、章の区切りに section、印象づけたい一文に statement か hero を1枚使ってよい（hero には内容に合う visualAsset を付け、photoMotion は zoom）",
    "- 長い内容は削らず、table、compare、twoColumn、section区切りなど適切な構造へ分ける",
    "- 研修・勉強会資料では、理解から実行へ進む物語（背景→要点→比較→手順→注意点→次アクション）を優先する",
    "- 素材にスライドへ載せきれない補足（根拠・具体例・内訳・数値の出所）があるときは、その項目の details に書く（資料全体で3〜6か所まで。本文はそのぶん短くする）",
    "- animation は基本的に省略（auto）。発表者が話しながら1つずつ見せたい工程・計画・原因分解だけ click を指定してよい",
    "- 動く資料: 素材に数字と出所がある本文スライドには、見出しの構造に合う操作を1つ付ける（上の「動く資料（1枚1操作）」）。数字のある本文スライドの details には rows・source を入れて根拠パネルにし、出所があれば slide の source も入れる。rank・shift・simulator・gap は素材の数字で作れるときだけ、資料全体で合わせて1〜3枚",
    "- モーショングラフィック: 表紙・章扉・statement・closing には backdrop（後ろで動く図形）を付けてよい（1〜2種類にそろえる）。kinetic は省略（自動で mask）してよく、研修・キックオフなら chars、問いかけの statement なら type を使ってよい",
    "- 見せ場（HTMLならではの動き。必ず入れる）: 資料のここぞという場面（結論の一文・決定的な数字・転換点・表紙・締め）を、5枚以下なら2か所、6〜9枚なら3か所、10枚以上なら3枚に1か所ほど選び、ダイナミックな動きを付ける。数字・グラフ・カードのスライドは entrance を spin（3Dで回転して現れる）か swirl（渦を巻いて飛び込む）か flip にし、transition を zoom か flip にする。statement・hero は kinetic を zoom か flip、backdrop か photoMotion（tilt・zoom）を付ける。章扉は transition を dive にして backdrop を付ける。表紙・closing は kinetic と backdrop を付ける。同じ動きを続けず、見せ場どうしは離す（見せ場以外は省略して資料の設定に任せる）",
    "- 最後のclosingには、誰が・いつまでに・何をするかが分かる具体的な次のアクションを入れる",
    "- 各スライドに、発表者が読み上げられる自然な notes（2〜4文）を付ける",
    ...audienceLines(input.audience),
    ...outlineLines(input.outline),
    "",
    ...LAYOUT_GUIDE,
    ...VISUAL_GUIDE,
    "",
    ...CAPACITY_GUIDE,
    "",
    ...WRITING_GUIDE,
    ...LOOK_GUIDE,
    "",
    `素材・依頼内容:\n${input.brief}`,
  ].join("\n");
}

/** Replace embedded images and videos with a marker: the AI never needs the pixels, and they are huge. */
export function withoutImageData(value, key = "") {
  if (key === "media" && value && typeof value === "object") return "[写真・動画あり]";
  // Objects placed by hand: a summary of what they are and say (the AI cannot change them; they always stay).
  if (key === "elements") return objectsSummary(value);
  if (key === "timeline") return Array.isArray(value) && value.length ? `[アニメーション${value.length}件]` : undefined;
  if (typeof value === "string") return value.startsWith("data:") ? "[画像あり]" : value;
  if (Array.isArray(value)) return value.map((item) => withoutImageData(item));
  if (value && typeof value === "object") return Object.fromEntries(Object.entries(value).filter(([childKey]) => childKey !== "imagePlacement" && childKey !== "sid").map(([childKey, child]) => [childKey, withoutImageData(child, childKey)]).filter(([, child]) => child !== undefined));
  return value;
}

export function buildRevisePrompt({ deck, slideIndex, instruction, issues = [], mode = "replace" }) {
  const slides = withoutImageData(deck.slides ?? deck.slideData ?? []);
  const inserting = mode === "insert";
  const target = slides[slideIndex];
  const context = deckNarrativeLines({ ...deck, slides }, [slideIndex]);
  return [
    PERSONA,
    inserting
      ? `資料「${deck.title ?? deck.deckTitle ?? ""}」（対象者: ${deck.audience || "未指定"}、目的: ${deck.purpose || "未指定"}）の${slideIndex}枚目と${slideIndex + 1}枚目の間に入れる、新しいスライドを1枚作ってください。typeはtitleとclosing以外から内容に合うものを選んでください。`
      : `資料「${deck.title ?? deck.deckTitle ?? ""}」（対象者: ${deck.audience || "未指定"}、目的: ${deck.purpose || "未指定"}）のうち、${slideIndex + 1}枚目のスライドだけを作り直してください。`,
    "最終回答は指定されたJSONスキーマに一致する1枚分のスライドJSONだけを返してください。入力にない数値や事実は作らないでください。",
    "資料全体の結論と対象スライドの役割・前後の根拠を踏まえ、指示に関係しない内容・画像を変えない。写真は明示的に頼まれたときだけ変更し、本文に合わない素材を足さない。",
    OBJECTS_RULE,
    ...memoLines(deck),
    !inserting && slideIndex === 0 ? "このスライドは表紙なので type は title のままにする。" : "",
    !inserting && slideIndex === slides.length - 1 ? "このスライドは最後なので type は closing のままにする。" : "",
    "",
    `${inserting ? "新しいスライドの内容" : "修正の指示"}: ${instruction || (inserting ? "前後の流れを補う内容" : "内容を保ったまま、より伝わりやすく改善する")}`,
    issues.length ? `画面で検出した問題（必ず解消する）:\n${issues.map((issue) => `- ${issue.field}: ${issue.message}`).join("\n")}` : "",
    "",
    "デッキ全体の流れ（前後との重複を避け、一貫性を保つ）:",
    ...context,
    "",
    inserting ? "" : "現在のスライドJSON:",
    inserting ? "" : JSON.stringify(target),
    "",
    ...LAYOUT_GUIDE,
    ...VISUAL_GUIDE,
    "",
    ...CAPACITY_GUIDE,
    ...WRITING_GUIDE,
    ...LOOK_GUIDE,
  ].filter((line) => line !== "").join("\n");
}

export function buildVariantsPrompt({ deck, slideIndex, instruction = "" }) {
  const slides = withoutImageData(deck.slides ?? []);
  const total = slides.length;
  const context = deckNarrativeLines({ ...deck, slides }, [slideIndex]);
  return [
    PERSONA,
    `資料「${deck.title ?? ""}」（対象者: ${deck.audience || "未指定"}、目的: ${deck.purpose || "未指定"}）の${slideIndex + 1}枚目について、見せ方の違う案を3つ作ってください。`,
    "最終回答は指定されたJSONスキーマに一致するJSONだけを返してください。入力にない数値や事実は作らないでください。",
    ...memoLines(deck),
    "- 3案は、レイアウト（type）か切り口がはっきり違うものにする（例：図解で見せる案／数字を大きく見せる案／一文で言い切る案）。3案を同じ見た目のグループにせず、少なくとも1案は前後のスライドと違うグループにする",
    "- label はその案の特徴を12字以内で（例：「数字を大きく」「3ステップの図解」）",
    slideIndex === 0 ? "- 表紙なので type は title のまま、言い回しで変化をつける" : "",
    slideIndex === total - 1 ? "- 最後のスライドなので type は closing のまま、言い回しで変化をつける" : "",
    instruction ? `- ユーザーの希望: ${instruction}` : "",
    "- 「[画像あり]」「[写真・動画あり]」の値はそのまま残す",
    OBJECTS_RULE,
    "- 見せ方の違いは主張の強弱・レイアウト・根拠の整理で作る。写真を新たに付けるのは明示的な依頼があり、本文と直接一致する場合だけ",
    "",
    "デッキ全体の流れ:",
    ...context,
    "",
    "現在のスライドJSON:",
    JSON.stringify(slides[slideIndex]),
    "",
    ...LAYOUT_GUIDE,
    ...VISUAL_GUIDE,
    "",
    ...CAPACITY_GUIDE,
    ...WRITING_GUIDE,
    ...LOOK_GUIDE,
  ].filter((line) => line !== "").join("\n");
}

export function buildRewritePrompt({ deck, instruction, settings = {} }) {
  const slides = withoutImageData(deck.slides ?? deck.slideData ?? []);
  const countRule = settings.minCount
    ? `枚数: ${settings.minCount}〜${settings.maxCount}枚の範囲で、内容に合う枚数にする（現在${slides.length}枚）。1枚目はtitle、最後はclosing。`
    : `枚数: ${settings.slideCount ?? slides.length}枚（slideDataを必ずこの枚数にする）。1枚目はtitle、最後はclosingのまま。`;
  const density = DENSITY_GUIDE[settings.textDensity] || DENSITY_GUIDE.standard;
  return [
    `${PERSONA}次の資料全体を、指示に沿って見直してください。`,
    "最終回答は指定されたJSONスキーマに一致するJSONだけを返してください。入力にない数値や事実は作らないでください（既存の数値は変えない）。",
    "",
    ...memoLines(deck),
    `見直しの指示: ${instruction}`,
    countRule,
    "「[画像あり]」「[写真・動画あり]」と書かれた値は写真・動画なので、そのまま残す（別の文字に書き換えない）。",
    OBJECTS_RULE,
    "type が blank（白紙）のスライドは、ユーザーが図形や文字を自由に置いて作ったページ。同じ位置に、同じ title の type=content（points は空）として1枚残す（元のページに戻します）。",
    `文字量: ${density.label}（${density.rule}）`,
    `対象者: ${deck.audience || "未指定"} ／ 目的: ${deck.purpose || "未指定"}`,
    "各本文スライドの takeaway と notes を保つか改善する。指示に関係しない良い部分は無理に変えない。details・animation・photoMotion・kinetic・backdrop・entrance・emphasis・transition は指示がなければ残す。",
    "資料の主張と前後の流れを先に読み、修正後も根拠から結論へのつながりを保つ。写真は見た目の穴埋めに追加せず、個々のスライドの本文と具体的に合うものだけにする。",
    "",
    ...LAYOUT_GUIDE,
    ...VISUAL_GUIDE,
    "",
    ...CAPACITY_GUIDE,
    "",
    ...WRITING_GUIDE,
    ...LOOK_GUIDE,
    "",
    "現在の資料JSON:",
    JSON.stringify({ deckTitle: deck.title ?? deck.deckTitle, purpose: deck.purpose, audience: deck.audience, slideData: slides }),
  ].join("\n");
}

export function buildNotesPrompt({ deck, indices, seconds = 60, tone = "丁寧" }) {
  const slides = withoutImageData(deck.slides ?? deck.slideData ?? []);
  const targets = indices?.length ? indices : slides.map((_, index) => index);
  const chars = Math.round(seconds * 5);
  return [
    "あなたは発表する人のためのスピーカーノート（読み上げ原稿）を書く担当です。発表は画面共有で行い、スライドの項目は順に現れ、クリックで補足が開きます。",
    "最終回答は指定されたJSONスキーマに一致するJSONだけを返してください。",
    `資料「${deck.title ?? deck.deckTitle ?? ""}」（対象者: ${deck.audience || "未指定"}、目的: ${deck.purpose || "未指定"}）`,
    ...memoLines(deck),
    "",
    "書き方:",
    `- 1枚あたり約${seconds}秒で話せる長さ（${Math.round(chars * 0.8)}〜${Math.round(chars * 1.2)}字）。${tone}な話し言葉（です・ます調）`,
    "- 最初にそのスライドの結論（takeaway）を言い、次に根拠・具体例を話し、最後に次のスライドへのつなぎを一言入れる",
    "- スライドに書いてある数値・事実だけを使う（details の補足も使ってよい）。新しい数値や事実を作らない。箇条書きをそのまま読み上げず、聞き手に語りかける文章にする",
    "- 表紙はあいさつと本日の目的、クロージングはお願いしたいこと・次のアクションで締める",
    "",
    `ノートを書くスライド: ${targets.map((index) => `${index + 1}枚目`).join("、")}（slide には枚数の番号を入れる）`,
    "",
    "資料JSON（番号は1から）:",
    JSON.stringify(slides.map((slide, index) => ({ no: index + 1, ...slide, notes: undefined }))),
  ].join("\n");
}

export function buildRepairPrompt({ issues, slides, error }) {
  if (error) {
    return [
      "前回の回答は検証に失敗しました。次の問題を直し、スキーマに一致するJSONだけをもう一度返してください。",
      `問題: ${error}`,
    ].join("\n");
  }
  const lines = issues.map((issue) => {
    const slide = slides[issue.slide] ?? {};
    return `- ${issue.slide + 1}枚目「${slide.title ?? ""}」の ${issue.field}: ${issue.message}`;
  });
  return [
    "確認したところ、次のテキストがスライドに収まらない長さです。",
    ...lines,
    "該当箇所だけを意味を保ったまま短く言い換えるか、要素を分割・削減してください（書ききれない補足は details に回してよい）。枚数・順序・他のスライドは変えずに、全体のJSONをもう一度返してください。",
  ].join("\n");
}

export class CodexSlideServer extends EventEmitter {
  /** `fallback` (a LocalModel) answers a job when Codex cannot; it may be null. */
  constructor({ rootDir, fallback = null }) {
    super();
    this.rootDir = rootDir;
    this.fallback = fallback;
    this.runtimeDir = process.env.CODEX_RUNTIME_DIR ?? join(rootDir, ".codex-runtime");
    this.codexHome = process.env.SLIDE_STUDIO_CODEX_HOME ?? join(this.runtimeDir, ".codex");
    this.workspaceDir = join(this.runtimeDir, "workspace");
    // `npm start` puts node_modules/.bin on PATH; running `node server.js` directly finds the local copy too.
    const local = join(rootDir, "node_modules", ".bin", "codex");
    this.binary = process.env.CODEX_BIN ?? (process.platform !== "win32" && existsSync(local) ? local : "codex");
    this.closing = false;
    this.proc = null;
    this.readyPromise = null;
    this.nextId = 1;
    this.pending = new Map();
    this.jobs = new Map();
    this.threadJobs = new Map();
    this.loginSessions = new Map();
    mkdirSync(this.codexHome, { recursive: true });
    mkdirSync(this.workspaceDir, { recursive: true });
  }

  async ensureStarted() {
    if (this.proc && this.readyPromise) return this.readyPromise;

    this.readyPromise = new Promise((resolve, reject) => {
      const proc = spawn(this.binary, ["app-server"], {
        cwd: this.workspaceDir,
        env: {
          ...process.env,
          CODEX_HOME: this.codexHome,
          HOME: this.runtimeDir,
        },
        stdio: ["pipe", "pipe", "pipe"],
      });
      this.proc = proc;

      proc.once("error", (error) => {
        if (this.proc !== proc) return;
        for (const pending of this.pending.values()) pending.reject(error);
        this.pending.clear();
        this.proc = null;
        this.readyPromise = null;
        reject(error);
      });
      proc.once("exit", (code, signal) => {
        if (this.proc !== proc) return;
        const error = new Error(`Codex App Server stopped (code=${code ?? "null"}, signal=${signal ?? "null"})`);
        for (const pending of this.pending.values()) pending.reject(error);
        this.pending.clear();
        for (const login of this.loginSessions.values()) this.emit("login", { sessionId: login.sessionId, success: false, error: "Codexとの接続が切れました。もう一度接続してください。" });
        this.loginSessions.clear();
        for (const job of this.jobs.values()) {
          if (!["completed", "failed"].includes(job.status) && job.provider !== "local") this.#fallBack(job, error.message).catch((fallbackError) => this.failJob(job.id, fallbackError));
        }
        this.proc = null;
        this.readyPromise = null;
        if (!this.closing) this.emit("exit", error);
      });

      const lines = readline.createInterface({ input: proc.stdout });
      lines.on("line", (line) => this.#handleLine(line));
      proc.stderr.on("data", (chunk) => {
        const message = String(chunk).trim();
        if (message) console.error("[codex]", message);
      });

      this.request("initialize", {
        clientInfo: { name: "html_pptx_studio", title: "HTML PPTX Studio", version: "1.0.0" },
        capabilities: { experimentalApi: true },
      })
        .then(() => {
          this.notify("initialized", {});
          resolve();
        })
        .catch((error) => {
          if (this.proc === proc && !proc.killed) proc.kill("SIGTERM");
          reject(error);
        });
    });

    return this.readyPromise;
  }

  send(message) {
    if (!this.proc?.stdin?.writable) throw new Error("Codex App Server is not running");
    this.proc.stdin.write(`${JSON.stringify(message)}\n`);
  }

  request(method, params = {}, timeoutMs = DEFAULT_TIMEOUT_MS) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`${method} timed out`));
      }, timeoutMs);
      this.pending.set(id, {
        resolve: (value) => { clearTimeout(timer); resolve(value); },
        reject: (error) => { clearTimeout(timer); reject(error); },
      });
      try {
        this.send({ method, id, params });
      } catch (error) {
        clearTimeout(timer);
        this.pending.delete(id);
        reject(error);
      }
    });
  }

  notify(method, params = {}) {
    this.send({ method, params });
  }

  async accountRead() {
    await this.ensureStarted();
    return this.request("account/read", { refreshToken: false });
  }

  async accountLogout() {
    await this.ensureStarted();
    return this.request("account/logout", {});
  }

  async startDeviceLogin(sessionId) {
    await this.ensureStarted();
    const previous = this.getPendingLogin(sessionId);
    if (previous) {
      this.loginSessions.delete(previous.loginId);
      try { await this.request("account/login/cancel", { loginId: previous.loginId }); }
      catch { /* The old code may have expired just before the replacement request. */ }
    }
    const result = await this.request("account/login/start", { type: "chatgptDeviceCode" });
    if (result?.loginId) this.loginSessions.set(result.loginId, {
      sessionId,
      loginId: result.loginId,
      userCode: result.userCode,
      verificationUrl: result.verificationUrl,
    });
    return result;
  }

  getPendingLogin(sessionId) {
    return [...this.loginSessions.values()].find((login) => login.sessionId === sessionId) ?? null;
  }

  close() {
    this.closing = true;
    if (this.proc && !this.proc.killed) this.proc.kill("SIGTERM");
  }

  /**
   * Run a structured-output job that may take several turns in one thread.
   * `finalize(parsed, attempt)` returns { result, detail } to finish, or
   * { retry: prompt, detail, fallback? } to ask Codex for a corrected answer.
   */
  async runJob(job, { prompt, outputSchema, finalize, maxAttempts = 2, effort = "medium", codex = job.useCodex !== false, parts = null, partsTotal = null, partCheck = null }) {
    job.plan = { outputSchema, finalize, maxAttempts, effort, partsTotal: parts?.total ?? partsTotal, partCheck };
    job.attempt = 1;
    job.provider = "codex";
    this.jobs.set(job.id, job);
    job.basePrompt = prompt;
    // A long deck is written a few slides at a time (one answer cannot hold it all): see #handleAnswer.
    if (parts && parts.total > parts.size) job.parts = { total: parts.total, size: parts.size, done: [], meta: null, retries: 0 };
    job.initialPrompt = job.parts ? partPrompt(job, { first: true }) : prompt;
    // Not connected to Codex: go straight to the stand-by AI (the caller checked it is there).
    if (!codex) return this.#fallBack(job, "Codexに接続していません。");
    try {
      await this.ensureStarted();
      this.#updateJob(job, { status: "starting", stage: "Codexを起動", detail: "専用のスレッドを準備しています。" });
      await this.#newThread(job);
      await this.#startTurn(job, job.initialPrompt);
    } catch (error) {
      await this.#fallBack(job, error instanceof Error ? error.message : String(error));
    }
    return job;
  }

  /** Image generation is a tool turn, not structured JSON. Never substitute a text-only local model. */
  async runImageJob(job, prompt, { referenceImage = null } = {}) {
    job.plan = { outputSchema: undefined, effort: "medium" };
    job.attempt = 1;
    job.provider = "codex";
    this.jobs.set(job.id, job);
    if (!job.useCodex) return this.failJob(job.id, new Error("画像生成にはCodexへの接続が必要です。"));
    try {
      if (referenceImage?.bytes?.length > 100 && referenceImage.bytes.length < 5_000_000) {
        job.referencePath = join(this.workspaceDir, `reference-${job.id}.${referenceImage.extension}`);
        writeFileSync(job.referencePath, referenceImage.bytes, { mode: 0o600 });
      }
      await this.ensureStarted();
      this.#updateJob(job, { status: "starting", stage: "画像生成を準備", detail: "Codexの画像生成ツールを起動しています。" });
      const thread = await this.request("thread/start", { cwd: this.workspaceDir });
      job.threadId = thread?.thread?.id;
      if (!job.threadId) throw new Error("Codex did not return a thread id");
      this.threadJobs.set(job.threadId, job.id);
      await this.#startTurn(job, prompt);
    } catch (error) { this.failJob(job.id, error); }
    return job;
  }

  getImage(jobId) {
    const job = this.jobs.get(jobId);
    return job?.kind === "image" && job.status === "completed" ? job.imageBytes ?? null : null;
  }

  /**
   * Codex could not do the job: hand it to the stand-by AI (Gemma) from the first request, or fail with
   * the reason when there is none (the web deployment, Ollama not running).
   */
  async #fallBack(job, reason) {
    if (job.provider === "local" || ["completed", "failed"].includes(job.status)) return;
    if (job.kind === "image") return this.failJob(job.id, new Error(`画像を生成できませんでした：${reason}`));
    const local = this.fallback ? await this.fallback.status() : null;
    if (!local?.available) {
      const note = local && !local.web ? `（予備AI：${local.reason}）` : "";
      this.#updateJob(job, { status: "failed", stage: "生成失敗", detail: `${reason}${note}`, error: `${reason}${note}` });
      return;
    }
    job.provider = "local";
    job.codexProblem = reason;
    job.attempt = 1;
    job.lastGood = null;
    job.localModel = local.model;
    if (job.parts) { job.parts.done = []; job.parts.meta = null; job.parts.complete = false; job.parts.retries = 0; }
    job.localMessages = [{ role: "user", content: job.parts ? partPrompt(job, { first: true }) : job.initialPrompt }];
    await this.#localTurn(job);
  }

  async #localTurn(job) {
    const label = `予備AI（${job.localModel}）`;
    this.#updateJob(job, {
      status: "running",
      stage: job.attempt > 1 ? "自動修正中" : `${label}で作成中`,
      detail: job.attempt > 1 ? `${label}に問題を直してもらっています。` : `Codexが使えないため、${label}で作っています（Codexより時間がかかります）。`,
    });
    let text;
    try {
      text = await this.fallback.chat(job.localMessages, job.plan.outputSchema, (streamed) => {
        if (Date.now() - (job.lastProgressAt ?? 0) < 700) return;
        job.lastProgressAt = Date.now();
        const partial = job.kind === "deck" && job.attempt === 1 ? partialItems(streamed) : null;
        const reply = job.kind === "chat" && job.attempt === 1 ? partialString(streamed, "reply") : "";
        const detail = partial?.length ? `${partial.length}枚目まで書き上がりました（${label}）。` : reply ? "返事を書いています。" : `${streamed.length.toLocaleString("ja-JP")}文字を受信しました（${label}）。`;
        this.#updateJob(job, { detail, progress: streamed.length, ...(partial ? { partial } : {}), ...(reply ? { partialReply: reply } : {}) });
      });
    } catch (error) {
      const message = `Codexが使えず（${job.codexProblem}）、予備AIでも作れませんでした：${error instanceof Error ? error.message : String(error)}`;
      this.#updateJob(job, { status: "failed", stage: "生成失敗", detail: message, error: message });
      return;
    }
    job.localMessages.push({ role: "assistant", content: text });
    await this.#handleAnswer(job, text, (next) => {
      job.localMessages.push({ role: "user", content: next });
      return this.#localTurn(job);
    });
  }

  /** A new thread for the job; events from the one it had are no longer the job's. */
  async #newThread(job) {
    const threadResult = await this.request("thread/start", { cwd: this.workspaceDir });
    const threadId = threadResult?.thread?.id;
    if (!threadId) throw new Error("Codex did not return a thread id");
    if (job.threadId) this.threadJobs.delete(job.threadId);
    job.threadId = threadId;
    job.formatSent = false;
    this.threadJobs.set(threadId, job.id);
  }

  /** Stop the job's turn and leave its thread (Codex would send the same request again and again). */
  #abandonTurn(job) {
    clearTimeout(job.errorTimer);
    if (job.threadId) {
      this.threadJobs.delete(job.threadId);
      if (job.turnId) this.request("turn/interrupt", { threadId: job.threadId, turnId: job.turnId }).catch(() => {});
    }
  }

  /** Ask again in a new thread after an answer was cut off. */
  #retryFresh(job, prompt) {
    this.#newThread(job)
      .then(() => this.#startTurn(job, prompt))
      .catch((retryError) => this.#fallBack(job, retryError instanceof Error ? retryError.message : String(retryError)).catch((error) => this.failJob(job.id, error)));
  }

  async #startTurn(job, text) {
    job.messageParts = new Map();
    job.finalText = null;
    job.lastPrompt = text;
    job.turnError = null;
    job.reconnects = 0;
    // Without the output format, the shape it asks for goes with the first request of the thread.
    const withFormat = job.noSchema && job.plan.outputSchema && !job.formatSent;
    if (withFormat) job.formatSent = true;
    const input = withFormat ? `${text}\n${formatLines(job.plan.outputSchema)}` : job.noSchema ? `${text}\n（回答は最初に示した形式の JSON だけ）` : text;
    const turnResult = await this.request("turn/start", {
      threadId: job.threadId,
      input: [{ type: "text", text: input }, ...(job.kind === "image" && job.referencePath ? [{ type: "localImage", path: job.referencePath }] : [])],
      cwd: this.workspaceDir,
      approvalPolicy: "never",
      sandboxPolicy: { type: "workspaceWrite", writableRoots: [this.workspaceDir], networkAccess: false },
      effort: job.plan.effort,
      summary: "concise",
      ...(job.plan.outputSchema && !job.noSchema ? { outputSchema: job.plan.outputSchema } : {}),
    });
    job.turnId = turnResult?.turn?.id ?? null;
    const first = { image: ["画像を生成中", "指定された絵柄で新しい画像を描いています。"], chat: ["考えています", "依頼を読んで、資料の該当箇所を確認しています。"], variants: ["別案を考えています", "見せ方の違う案を組み立てています。"], outline: ["骨子を考えています", "話の流れと各スライドの役割を決めています。"], notes: ["原稿を書いています", "スライドごとの話す内容を組み立てています。"], translate: ["翻訳しています", "スライドの文字を訳しています。"], synonyms: ["類義語を探しています", "言い換えられる語を挙げています。"] }[job.kind] ?? ["構成を生成", "Codexが流れとレイアウトを設計しています。"];
    this.#updateJob(job, {
      status: "running",
      stage: job.attempt > 1 ? "自動修正中" : first[0],
      detail: job.attempt > 1 ? "検出した問題をCodexに直してもらっています。" : first[1],
    });
  }

  getJob(jobId) {
    const job = this.jobs.get(jobId);
    return job ? this.#publicJob(job) : null;
  }

  failJob(jobId, error) {
    const job = this.jobs.get(jobId);
    if (!job) return;
    const message = error instanceof Error ? error.message : String(error);
    this.#updateJob(job, { status: "failed", stage: "エラー", detail: message, error: message });
  }

  #handleLine(line) {
    let message;
    try {
      message = JSON.parse(line);
    } catch {
      console.error("Invalid Codex message", line.slice(0, 500));
      return;
    }

    if (message.id != null && !message.method) {
      const pending = this.pending.get(message.id);
      if (!pending) return;
      this.pending.delete(message.id);
      if (message.error) pending.reject(new Error(message.error.message ?? JSON.stringify(message.error)));
      else pending.resolve(message.result);
      return;
    }

    if (message.method === "account/login/completed") {
      const loginId = message.params?.loginId;
      const sessionId = loginId ? this.loginSessions.get(loginId)?.sessionId : null;
      if (loginId) this.loginSessions.delete(loginId);
      if (sessionId) this.emit("login", { sessionId, ...message.params });
      return;
    }

    if (message.id != null && message.method) {
      this.#handleServerRequest(message);
      return;
    }

    this.#handleNotification(message);
  }

  #handleServerRequest(message) {
    const job = this.#jobFor(message.params?.threadId);
    if (job) {
      this.#updateJob(job, {
        status: "failed",
        stage: "確認が必要",
        detail: `Codexが追加確認を要求しました: ${message.method}`,
        error: "条件を具体化して再実行してください。",
      });
    }
    this.send({ id: message.id, result: { decision: "decline", action: "decline", content: null } });
  }

  #handleNotification(message) {
    const params = message.params ?? {};
    const job = this.#jobFor(params.threadId ?? params.turn?.threadId);
    if (!job || ["completed", "failed"].includes(job.status)) return;

    switch (message.method) {
      case "item/completed":
        if (job.kind === "image" && params.item?.type === "imageGeneration") {
          job.imageItem = params.item;
          this.#updateJob(job, { stage: "画像を確認中", detail: "生成結果を保存・検証しています。" });
        }
        if (params.item?.type === "agentMessage") {
          const text = extractItemText(params.item);
          if (text) job.finalText = text;
        }
        break;
      case "item/agentMessage/delta": {
        const itemId = params.itemId ?? "final";
        job.messageParts.set(itemId, `${job.messageParts.get(itemId) ?? ""}${params.delta ?? ""}`);
        const received = [...job.messageParts.values()].reduce((sum, part) => sum + part.length, 0);
        if (Date.now() - (job.lastProgressAt ?? 0) > 700) {
          job.lastProgressAt = Date.now();
          // First drafts stream slide by slide, so the page can show each one as it is written.
          const streamed = [...job.messageParts.values()].join("");
          const partial = job.kind === "deck" && job.attempt === 1 ? [...(job.parts?.done || []), ...partialItems(streamed)] : null;
          const reply = job.kind === "chat" && job.attempt === 1 ? partialString(streamed, "reply") : "";
          const detail = partial?.length ? `${partial.length}枚目まで書き上がりました${job.parts ? `（全${job.parts.total}枚）` : ""}。` : reply ? "返事を書いています。" : `${received.toLocaleString("ja-JP")}文字を受信しました。`;
          this.#updateJob(job, { stage: job.attempt > 1 ? "自動修正中" : "構成を書き出し中", detail, progress: received, ...(partial ? { partial } : {}), ...(reply ? { partialReply: reply } : {}) });
        }
        break;
      }
      case "item/reasoning/summaryTextDelta": {
        job.thought = `${job.thought ?? ""}${params.delta ?? ""}`.slice(-400);
        this.#updateJob(job, { stage: job.attempt > 1 ? "自動修正中" : "構成を検討", detail: job.thought.split("\n").filter(Boolean).at(-1)?.slice(-160) || "論点を整理しています。" });
        break;
      }
      case "error":
        console.warn(`[codex] job ${job.id} (${job.kind ?? "deck"}) error${params.willRetry ? " (Codex retries)" : ""}: ${params.error?.message ?? ""} ${JSON.stringify(params.error?.codexErrorInfo ?? "")}`.slice(0, 600));
        // Codex says only "Reconnecting... n/5" when an answer is cut off at the output limit (the reason comes after
        // all its tries). A dropped stream twice in one turn of a long answer is that: do not wait for the rest.
        if (params.willRetry && errorKind(params.error) === "responseStreamDisconnected") job.reconnects = (job.reconnects ?? 0) + 1;
        const cutOff = tooLong(params.error) || (job.reconnects >= 2 && ["deck", "revise", "chat", "variants"].includes(job.kind ?? "deck"));
        if (params.willRetry && cutOff && job.plan.outputSchema) {
          // Cut off at the output limit: Codex would send the same request again, which is cut off the same way.
          this.#abandonTurn(job);
          this.#failTurn(job, tooLong(params.error) ? params.error : { message: `回答が途中で切れました（max_output_tokens と判断: ${params.error?.message ?? ""}）`, codexErrorInfo: params.error?.codexErrorInfo });
          break;
        }
        if (params.willRetry) {
          // Codex retries this itself (a dropped stream, a brief overload): keep waiting for the turn.
          this.#updateJob(job, { detail: TRANSIENT_ERRORS.has(errorKind(params.error)) ? "Codexの応答が途切れたため、Codexがもう一度送っています…" : "Codexがもう一度試しています…" });
          break;
        }
        // The turn ends right after with status "failed"; #completeTurn decides whether to try again.
        job.turnError = params.error ?? { message: "Codexでエラーが発生しました。" };
        clearTimeout(job.errorTimer);
        job.errorTimer = setTimeout(() => {
          if (!["completed", "failed"].includes(job.status) && job.turnError) this.#failTurn(job, job.turnError);
        }, 20_000);
        job.errorTimer.unref?.();
        break;
      case "turn/completed":
        this.#completeTurn(job, params.turn).catch((error) => this.failJob(job.id, error));
        break;
      default:
        break;
    }
  }

  /** A turn that failed: once per job, a transient failure (overload, dropped stream) is simply asked again. */
  #failTurn(job, error, status = "failed") {
    clearTimeout(job.errorTimer);
    if (tooLong(error) && job.provider !== "local") {
      this.#abandonTurn(job);
      // Less thinking leaves more of the output limit for the answer.
      job.plan.effort = "low";
      // First the same request without the strict output format (a long format can make the answer run on),
      // in a new thread: the format is written out in the request instead and the answer checked here.
      if (job.plan.outputSchema && !job.noSchema) {
        job.noSchema = true;
        this.#updateJob(job, { stage: "書き直し中", detail: "回答が途中で切れたため、書き方を変えてやり直しています。" });
        this.#retryFresh(job, job.parts ? partPrompt(job, { first: true }) : job.initialPrompt);
        return;
      }
      // The deck did not fit in one answer: write it a few slides at a time (fewer still if a part did not fit).
      if (job.kind === "deck" && job.plan.partsTotal > 2 && (!job.parts || job.parts.size > 2) && !job.splitTried?.[job.parts?.size ?? 0]) {
        job.splitTried = { ...(job.splitTried || {}), [job.parts?.size ?? 0]: true };
        if (job.parts) job.parts.size = 2;
        else job.parts = { total: job.plan.partsTotal, size: Math.min(4, Math.ceil(job.plan.partsTotal / 2)), done: [], meta: null, retries: 0 };
        this.#updateJob(job, { stage: "分けて作成中", detail: `1回の回答に入りきらなかったため、${job.parts.size}枚ずつに分けて書いています。` });
        this.#retryFresh(job, partPrompt(job, { first: true, note: job.parts.done.length ? "（前の回答は長すぎて途中で切れました。）" : "" }));
        return;
      }
    }
    // (Not an answer that was cut off: that is the same however often it is sent.)
    if (TRANSIENT_ERRORS.has(errorKind(error)) && !tooLong(error) && !job.transientRetried && job.lastPrompt) {
      job.transientRetried = true;
      this.#updateJob(job, { stage: "再試行中", detail: "Codex側が混み合っているため、もう一度依頼しています。" });
      this.#startTurn(job, job.lastPrompt).catch((retryError) => this.failJob(job.id, retryError));
      return;
    }
    const message = error ? friendlyCodexError(error) : `Codexの処理が終わりませんでした（${status}）。もう一度お試しください。`;
    // Too much material fails on the smaller stand-by model as well; everything else is worth a try.
    if (errorKind(error) === "contextWindowExceeded") {
      this.#updateJob(job, { status: "failed", stage: "生成失敗", detail: message, error: message });
      return;
    }
    this.#fallBack(job, message).catch((fallbackError) => this.failJob(job.id, fallbackError));
  }

  async #completeTurn(job, turn) {
    clearTimeout(job.errorTimer);
    if (turn?.status !== "completed") {
      this.#failTurn(job, turn?.error ?? job.turnError, turn?.status ?? "unknown");
      return;
    }
    if (job.kind === "image") {
      try {
        const item = job.imageItem;
        if (!item || item.status !== "completed" || !item.savedPath) throw new Error("Codexから画像ファイルが返されませんでした。画像生成が利用可能か確認してください。");
        const path = realpathSync(item.savedPath);
        const allowed = [this.workspaceDir, this.codexHome].some((root) => {
          const rel = relative(realpathSync(root), path);
          return rel && rel !== ".." && !rel.startsWith(`..${sep}`) && !isAbsolute(rel);
        });
        if (!allowed) throw new Error("画像の保存先を検証できませんでした。");
        const size = statSync(path).size;
        if (size < 100 || size > 25_000_000) throw new Error("生成画像のサイズが不正です。");
        const bytes = readFileSync(path);
        const png = bytes.subarray(0, 8).equals(Buffer.from("89504e470d0a1a0a", "hex"));
        const jpeg = bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff;
        const webp = bytes.toString("ascii", 0, 4) === "RIFF" && bytes.toString("ascii", 8, 12) === "WEBP";
        if (!png && !jpeg && !webp) throw new Error("生成結果が画像ではありません。" );
        job.imageBytes = bytes;
        job.imageMime = png ? "image/png" : jpeg ? "image/jpeg" : "image/webp";
        this.#updateJob(job, { status: "completed", stage: "完了", detail: "新しい画像を生成しました。確認してからスライドに採用できます。", result: { image: { url: `/api/decks/${job.id}/image`, mime: job.imageMime, revisedPrompt: item.revisedPrompt ?? null } } });
      } catch (error) { this.failJob(job.id, error); }
      return;
    }
    const streamed = [...job.messageParts.values()].join("\n");
    await this.#handleAnswer(job, job.finalText || streamed, (next) => this.#startTurn(job, next));
  }

  /** Check an answer (from Codex or the stand-by AI); `next(prompt)` asks the same AI again. */
  async #handleAnswer(job, text, next) {
    if (job.parts && !job.parts.complete) {
      const [from, to] = partRange(job);
      let part = null;
      try { part = parseDeckResult(text); } catch { /* asked again below */ }
      let slides = Array.isArray(part?.slideData) ? part.slideData : [];
      // The whole deck (or more than asked): keep the slides of this part.
      if (slides.length > to - from + 1) slides = slides.length >= to ? slides.slice(from - 1, to) : slides.slice(0, to - from + 1);
      // Each part is checked as it comes (a repair of the joined deck would ask for all of it again).
      const problem = slides.length && job.plan.partCheck ? job.plan.partCheck(slides, { from, to, total: job.parts.total }) : "";
      if (problem && job.parts.retries < 2) {
        job.parts.retries += 1;
        await next(partPrompt(job, { note: `前の回答の${from}〜${to}枚目に問題がありました：${problem}\n直して、同じ範囲をもう一度返してください。` }));
        return;
      }
      if (!slides.length) {
        if (job.parts.retries < 2) {
          job.parts.retries += 1;
          await next(partPrompt(job, { note: "前の回答から slideData を読み取れませんでした。JSON だけを返してください。" }));
          return;
        }
        this.#updateJob(job, { status: "failed", stage: "生成失敗", detail: `${from}〜${to}枚目を作れませんでした。`, error: `${from}〜${to}枚目を作れませんでした。`, rawOutput: text });
        return;
      }
      job.parts.meta ??= part;
      job.parts.done.push(...slides);
      job.parts.retries = 0;
      if (job.parts.done.length < job.parts.total) {
        this.#updateJob(job, { stage: "構成を書き出し中", detail: `全${job.parts.total}枚のうち${job.parts.done.length}枚を書き上げました。`, partial: [...job.parts.done] });
        await next(partPrompt(job));
        return;
      }
      job.parts.complete = true;
      text = JSON.stringify({ ...job.parts.meta, slideData: job.parts.done.slice(0, job.parts.total) });
      // A repair asks for the whole deck again, which is what did not fit: what the check finds is kept as 要確認.
      job.plan.maxAttempts = job.attempt;
    }
    const canRetry = job.attempt < job.plan.maxAttempts;
    const who = job.provider === "local" ? `予備AI（${job.localModel}）` : "Codex";
    let outcome;
    try {
      this.#updateJob(job, { stage: "品質チェック", detail: "文字量とスキーマを確認しています。" });
      outcome = await job.plan.finalize(parseDeckResult(text), job.attempt);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (canRetry) {
        job.attempt += 1;
        await next(buildRepairPrompt({ error: message }));
        return;
      }
      if (job.lastGood) {
        this.#updateJob(job, { status: "completed", stage: "完了", detail: "自動修正に失敗したため、直前の有効な構成を採用しました。", result: job.lastGood });
        return;
      }
      this.#updateJob(job, {
        status: "failed",
        stage: "結果解析エラー",
        detail: `${who}の回答をスライド構成として検証できませんでした。`,
        error: message,
        rawOutput: text,
      });
      return;
    }
    if (outcome.retry && canRetry) {
      job.attempt += 1;
      job.lastGood = outcome.fallback ?? job.lastGood;
      this.#updateJob(job, { stage: "自動修正中", detail: outcome.detail ?? "問題を修正しています。" });
      await next(outcome.retry);
      return;
    }
    const result = outcome.result ?? outcome.fallback ?? job.lastGood;
    if (!result) {
      this.#updateJob(job, { status: "failed", stage: "生成失敗", detail: outcome.detail ?? "結果を確定できませんでした。", error: outcome.detail ?? "結果を確定できませんでした。" });
      return;
    }
    const note = job.provider === "local" ? `（Codexが使えなかったため、${who}で作りました）` : "";
    this.#updateJob(job, { status: "completed", stage: "完了", detail: `${outcome.detail ?? "完了しました。"}${note}`, result });
  }

  #jobFor(threadId) {
    const jobId = threadId ? this.threadJobs.get(threadId) : null;
    return jobId ? this.jobs.get(jobId) : null;
  }

  #publicJob(job) {
    return {
      id: job.id,
      kind: job.kind ?? "deck",
      provider: job.provider ?? "codex",
      localModel: job.provider === "local" ? job.localModel : null,
      status: job.status,
      stage: job.stage,
      detail: job.detail,
      attempt: job.attempt ?? 1,
      progress: job.progress ?? 0,
      error: job.error ?? null,
      deck: job.result?.deck ?? null,
      slide: job.result?.slide ?? null,
      notes: job.result?.notes ?? null,
      texts: job.result?.texts ?? null,
      senses: job.result?.senses ?? null,
      chat: job.result?.chat ?? null,
      outline: job.result?.outline ?? null,
      variants: job.result?.variants ?? null,
      original: job.result?.original ?? null,
      image: job.result?.image ?? null,
      partial: job.status === "completed" ? null : job.partial ?? null,
      partialReply: job.status === "completed" ? null : job.partialReply ?? null,
      issues: job.result?.issues ?? [],
      rawOutput: job.rawOutput ?? null,
      createdAt: job.createdAt,
      updatedAt: job.updatedAt,
    };
  }

  #updateJob(job, patch) {
    Object.assign(job, patch, { updatedAt: new Date().toISOString() });
    if (patch.status === "failed") console.warn(`[job] ${job.id} (${job.kind ?? "deck"}, ${job.provider ?? "codex"}) failed: ${String(job.error ?? job.detail ?? "").slice(0, 400)}`);
    this.emit("job", this.#publicJob(job));
    if (["completed", "failed"].includes(job.status) && !job.cleanupTimer) {
      if (job.referencePath) { try { unlinkSync(job.referencePath); } catch { /* already removed */ } job.referencePath = null; }
      job.cleanupTimer = setTimeout(() => {
        this.jobs.delete(job.id);
        if (job.threadId) this.threadJobs.delete(job.threadId);
      }, 60 * 60 * 1000);
      job.cleanupTimer.unref?.();
    }
  }
}
