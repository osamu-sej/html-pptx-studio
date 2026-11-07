import assert from "node:assert/strict";
import test from "node:test";

import { SLIDE_TYPES, codexChatSchema, codexDeckSchema, codexNotesSchema, codexOutlineSchema, codexSlideSchema, codexVariantsSchema, slideSchema } from "../server/schemas.mjs";

// Codex structured output uses strict JSON Schema: every object must forbid extra
// keys and list every property as required, and array items must be one schema.
function auditStrict(schema, path = "$") {
  const problems = [];
  const visit = (node, at) => {
    if (!node || typeof node !== "object") return;
    if (Array.isArray(node)) return node.forEach((child, i) => visit(child, `${at}[${i}]`));
    if (node.type === "object" || node.properties) {
      const keys = Object.keys(node.properties ?? {});
      if (node.additionalProperties !== false) problems.push(`${at}: additionalProperties must be false`);
      const required = new Set(node.required ?? []);
      for (const key of keys) if (!required.has(key)) problems.push(`${at}.${key}: must be required (nullable if optional)`);
    }
    if (Array.isArray(node.items)) problems.push(`${at}: tuple items are not supported`);
    if ("prefixItems" in node) problems.push(`${at}: prefixItems are not supported`);
    for (const [key, child] of Object.entries(node)) if (key !== "required") visit(child, `${at}.${key}`);
  };
  visit(schema, path);
  return problems;
}

test("Codex output schemas satisfy strict structured-output rules", () => {
  assert.deepEqual(auditStrict(codexDeckSchema), []);
  assert.deepEqual(auditStrict(codexSlideSchema), []);
  assert.deepEqual(auditStrict(codexNotesSchema), []);
  assert.deepEqual(auditStrict(codexChatSchema), []);
  assert.deepEqual(auditStrict(codexOutlineSchema), []);
  assert.deepEqual(auditStrict(codexVariantsSchema), []);
  // Uploaded photos, videos and hand placement are the user's; the AI may set builds, photo motion and details.
  for (const schema of [codexDeckSchema, codexChatSchema, codexSlideSchema, codexVariantsSchema]) {
    assert.doesNotMatch(JSON.stringify(schema), /customImage|imagePlacement|"media"/);
  }
  const chat = JSON.stringify(codexChatSchema);
  assert.match(chat, /"animation"/);
  assert.match(chat, /"photoMotion"/);
  assert.match(chat, /"details"/);
  assert.doesNotMatch(chat, /"theme"/, "the AI cannot change the theme");
  // Motion graphics: the AI may choose kinetic type and backdrops per slide, and change the deck's defaults in chat.
  for (const schema of [codexDeckSchema, codexChatSchema, codexSlideSchema, codexVariantsSchema]) {
    const text = JSON.stringify(schema);
    assert.match(text, /"kinetic"/);
    assert.match(text, /"backdrop"/);
    assert.match(text, /"orbits"/);
    assert.match(text, /"scramble"/);
  }
  assert.match(chat, /"motion"/);
  assert.match(chat, /"wipe"/);
  assert.match(chat, /"circle"/);
  // More motion types: each slide may pick its own entrance, emphasis and transition; chat may change the deck's.
  for (const schema of [codexDeckSchema, codexChatSchema, codexSlideSchema, codexVariantsSchema]) {
    const text = JSON.stringify(schema);
    for (const key of ["entrance", "emphasis", "transition", "spotlight", "confetti", "curtain", "drift"]) assert.match(text, new RegExp(`"${key}"`), key);
  }
  const motion = codexChatSchema.properties.motion.anyOf.find((option) => option.type === "object");
  assert.deepEqual(Object.keys(motion.properties).sort(), ["backdrop", "draw", "emphasis", "entrance", "hover", "kinetic"]);
  // Pages the audience works with: evidence panels, rankings by view, before/after, simulators and gaps.
  for (const schema of [codexDeckSchema, codexChatSchema, codexSlideSchema, codexVariantsSchema]) {
    const text = JSON.stringify(schema);
    for (const key of ["simulator", "gap", "rank", "shift", "views", "measures", "formula", "compareValue", "\"rows\"", "\"source\"", "\"note\"", "\"now\""]) assert.ok(text.includes(key.startsWith("\"") ? key : `"${key}"`), key);
  }
  assert.equal(codexSlideSchema.type, "object");
  assert.equal(SLIDE_TYPES.length, 44);
  assert.ok(SLIDE_TYPES.includes("hero") && SLIDE_TYPES.includes("statement"));
});

test("interactive slides and evidence validate, and nonsense does not", () => {
  const ok = (slide) => slideSchema.safeParse(slide).success;
  assert.ok(ok({ type: "simulator", title: "試算", inputs: [{ label: "人数", value: 120, min: 50, max: 600 }], formula: "a × 12" }));
  assert.ok(ok({ type: "gap", title: "不足", target: 3000, current: 1440, measures: [{ title: "テンプレート", value: 620 }] }));
  assert.ok(ok({ type: "imageText", title: "順位", image: { chartType: "rank", data: { views: [{ label: "全体", items: [{ label: "A", value: 1 }] }] } } }));
  assert.ok(ok({ type: "imageText", title: "差", image: { chartType: "shift", data: { items: [{ label: "A", before: 9, value: 4 }] } } }));
  assert.ok(ok({ type: "kpi", title: "成果", source: "社内集計", items: [{ label: "時間", value: "1,440h" }], details: [{ target: "takeaway", text: "内訳", rows: [{ label: "資料", value: "620h" }], source: "記録", note: "3か月目" }] }));
  assert.ok(!ok({ type: "simulator", title: "試算", inputs: [], formula: "a" }), "a simulator needs a condition");
  assert.ok(!ok({ type: "gap", title: "不足", target: 3000, current: 1440, measures: [] }), "a gap needs a measure");
});
