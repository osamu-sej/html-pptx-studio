import assert from "node:assert/strict";
import test from "node:test";

import { LOOKS, lookOf, maxSameLook, varietyIssues } from "../public/layout-looks.mjs";
import { SLIDE_TYPES } from "../server/schemas.mjs";
import { buildChatPrompt, buildDeckPrompt, buildOutlinePrompt, varietyRepairLines } from "../server/codex-app-server.mjs";

const deckOf = (...types) => ["title", ...types, "closing"].map((type) => ({ type }));

test("every layout has a look, except the cover, agenda, chapter dividers and the close", () => {
  const grouped = Object.values(LOOKS).flatMap((look) => look.types);
  assert.equal(new Set(grouped).size, grouped.length, "a layout belongs to one look");
  for (const type of SLIDE_TYPES) {
    if (["title", "agenda", "section", "closing"].includes(type)) assert.equal(lookOf({ type }), null, type);
    else assert.ok(lookOf({ type }), type);
  }
  assert.equal(lookOf({ type: "imageText", image: { chartType: "bar", data: {} } }), "numbers");
  assert.equal(lookOf({ type: "imageText", visualAsset: "storeOperations" }), "impact", "a photo with text is a photo slide");
});

test("a timeline, a process, a roadmap and a lane diagram count as the same look", () => {
  // The deck that prompted this: history and future of generative AI, 8 slides, 4 of them lines across the slide.
  const issues = varietyIssues(deckOf("executiveSummary", "timeline", "diagram", "compare", "process", "roadmap"));
  const crowded = issues.find((issue) => issue.rule === "crowded");
  assert.ok(crowded);
  assert.equal(crowded.look, "flow");
  assert.deepEqual(crowded.slides, [2, 3, 5, 6]);
  assert.equal(crowded.keep, 2);
  assert.match(crowded.message, /「横に並ぶ流れ」.*4枚.*3・4・6・7枚目.*2枚まで/);
  assert.ok(issues.some((issue) => issue.rule === "few" && /3種類/.test(issue.message)));
});

test("a varied deck passes; alike looks side by side and too few looks are named", () => {
  assert.deepEqual(varietyIssues(deckOf("executiveSummary", "kpi", "logicTree", "process", "cards", "statement")), []);
  const adjacent = varietyIssues(deckOf("executiveSummary", "kpi", "process", "roadmap", "cards", "statement"));
  assert.deepEqual(adjacent.map((issue) => [issue.rule, issue.slides]), [["adjacent", [3, 4]]]);
  // Charts and shapes differ from one another, so two of them in a row are fine.
  assert.deepEqual(varietyIssues(deckOf("executiveSummary", "kpi", "waterfall", "cycle", "pyramid", "cards")), []);
  assert.equal(maxSameLook(6), 2);
  assert.equal(maxSameLook(10), 3);
  assert.deepEqual(varietyIssues(deckOf("content")), [], "a single body slide has nothing to vary");
});

test("the AI is told which layouts look alike, how many it may use, and how to repair a plain outline", () => {
  const outline = buildOutlinePrompt({ brief: "生成AIの歴史と将来", settings: { slideCount: 8 } });
  assert.match(outline, /横に並ぶ流れ: process・timeline・roadmap・flowChart・gantt・diagram/);
  assert.match(outline, /本文6枚なら、同じ見た目のグループ.*2枚まで/);
  assert.match(outline, /時期を書けないなら流れにしない/);
  for (const prompt of [buildDeckPrompt({ brief: "x", settings: { slideCount: 8 } }), buildChatPrompt({ deck: { title: "x", slides: deckOf("content") }, message: "直して" })]) {
    assert.match(prompt, /見た目のグループ/);
  }
  const lines = varietyRepairLines(varietyIssues(deckOf("executiveSummary", "timeline", "diagram", "compare", "process", "roadmap")), { outline: true }).join("\n");
  assert.match(lines, /3・4・6・7枚目/);
  assert.match(lines, /type と、それに合わせた content のメモだけ直す/);
  assert.match(lines, /pyramid・stepUp/);
});
