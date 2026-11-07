import assert from "node:assert/strict";
import test from "node:test";
import { imageBrief } from "../server/image-brief.mjs";

test("ordinary slide image brief includes the claim and concrete body", () => {
  const brief = imageBrief({ type: "content", title: "店舗の発注", takeaway: "火曜の訪店で店長と改善策を決める", points: ["発注記録を一緒に確認", "欠品した商品を特定"] });
  assert.match(brief, /火曜の訪店で店長と改善策を決める/);
  assert.match(brief, /発注記録を一緒に確認/);
  assert.match(brief, /欠品した商品を特定/);
});

test("Gantt image brief preserves the thesis, chronological steps and visual exclusions", () => {
  const brief = imageBrief({ type: "gantt", title: "一週間の準備と訪店", takeaway: "月曜夕方から火曜午前が準備の集中時間となる", periods: ["月曜", "月曜夕方", "火曜午前", "火曜午後", "水曜〜金曜"], items: [
    { title: "方針確認", start: 0, span: 1, desc: "会議内容を確認" },
    { title: "集中準備", start: 1, span: 2, desc: "データ整理から店舗別提案まで" },
    { title: "訪店開始", start: 3, span: 1, desc: "準備を基に訪店" },
  ] });
  assert.match(brief, /月曜夕方から火曜午前/);
  assert.match(brief, /集中準備: 月曜夕方→火曜午前/);
  assert.match(brief, /訪店開始: 火曜午後/);
  assert.match(brief, /最重要の作業/);
  assert.match(brief, /無関係な店舗アイコン群/);
});
