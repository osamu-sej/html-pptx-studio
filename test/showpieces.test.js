import assert from "node:assert/strict";
import test from "node:test";

import { isShowpiece, showpieceCount, stageShowpieces } from "../server/showpieces.mjs";
import { BACKDROP_KINDS, ENTRANCES, KINETIC_STYLES, PHOTO_MOTIONS, TRANSITIONS } from "../server/schemas.mjs";

const deck = (types) => types.map((type, i) => ({ type, title: `スライド${i + 1}` }));

test("a deck the AI left still has its showpieces: the key slides, spread out, never the same twice in a row", () => {
  const slides = deck(["title", "executiveSummary", "content", "kpi", "process", "cards", "statement", "content", "closing"]);
  const { slides: staged, added } = stageShowpieces(slides);
  assert.equal(added.length, showpieceCount(9));
  assert.equal(staged.filter(isShowpiece).length, 3);
  // The strongest moments: the one-line message and the figures come first.
  const chosen = added.map(([index]) => staged[index].type);
  assert.ok(chosen.includes("statement"), chosen.join(","));
  assert.ok(chosen.includes("kpi"), chosen.join(","));
  // Not next to each other.
  const at = added.map(([index]) => index).sort((a, b) => a - b);
  for (let i = 1; i < at.length; i += 1) assert.ok(at[i] - at[i - 1] > 1, at.join(","));
  // Entrances do not repeat.
  const entrances = added.map(([, what]) => what.entrance).filter(Boolean);
  for (let i = 1; i < entrances.length; i += 1) assert.notEqual(entrances[i], entrances[i - 1]);
  // Only values a slide may have; the input is not changed.
  const allowed = { entrance: ENTRANCES, transition: TRANSITIONS, kinetic: KINETIC_STYLES, backdrop: BACKDROP_KINDS, photoMotion: PHOTO_MOTIONS };
  for (const [, what] of added) for (const [key, value] of Object.entries(what)) assert.ok(allowed[key]?.includes(value), `${key}=${value}`);
  assert.ok(!slides.some(isShowpiece));
});

test("what the AI chose stays; only the missing showpieces are added", () => {
  const slides = deck(["title", "kpi", "content", "content", "statement", "closing"]);
  slides[1].entrance = "spin";
  slides[4].kinetic = "type";
  slides[4].backdrop = "none";
  const { slides: staged, added } = stageShowpieces(slides);
  assert.equal(added.length, 1);
  assert.equal(staged[1].entrance, "spin");
  assert.equal(staged[4].kinetic, "type");
  assert.equal(staged[4].backdrop, "none");
  // Enough already: nothing is added.
  slides[0].backdrop = "rays";
  assert.equal(stageShowpieces(slides).added.length, 0);
});

test("drill pages are left alone, and a short deck gets two showpieces", () => {
  const slides = deck(["title", "kpi", "content", "closing"]);
  slides.splice(2, 0, { type: "kpi", title: "深掘り", drillOf: 1 });
  const { slides: staged, added } = stageShowpieces(slides);
  assert.equal(added.length, 2);
  assert.ok(!isShowpiece(staged[2]));
  assert.equal(showpieceCount(1), 0);
  assert.equal(showpieceCount(30), 6);
});
