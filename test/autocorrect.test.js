// オートコレクト: symbols as they are typed, lists from a marker and a space at a paragraph's start (in the marker's
// style), each kind turned off by its option, and the options kept.
import assert from "node:assert/strict";
import test from "node:test";

import { AUTOCORRECT_DEFAULTS, autoCorrect, autoCorrectOptions, saveAutoCorrectOptions } from "../public/editor/autocorrect.mjs";

test("symbols: (c) ©, --> →, ... …, only when the last character completes them", () => {
  assert.deepEqual(autoCorrect("著作権 (c)", ")"), { replace: 3, with: "©" });
  assert.deepEqual(autoCorrect("A-->", ">"), { replace: 3, with: "→" });
  assert.deepEqual(autoCorrect("続く...", "."), { replace: 3, with: "…" });
  assert.deepEqual(autoCorrect("<=>", ">"), { replace: 3, with: "⇔" });
  assert.deepEqual(autoCorrect("(TM)", ")"), { replace: 4, with: "™" });
  assert.equal(autoCorrect("(c) と書く", "く"), null, "only right after it is typed");
  assert.equal(autoCorrect("abc", "c"), null);
  assert.equal(autoCorrect("(c)", ")", { options: { ...AUTOCORRECT_DEFAULTS, replace: false } }), null, "turned off");
});

test("lists: ・ / - / ● and a space make bullets, 1. / (1) / ① / a. numbers, in their style", () => {
  assert.deepEqual(autoCorrect("・ ", " "), { list: "bullet", style: "disc", remove: 2 });
  assert.deepEqual(autoCorrect("- ", " "), { list: "bullet", style: "dash", remove: 2 });
  assert.deepEqual(autoCorrect("◆　", "　"), { list: "bullet", style: "diamond", remove: 2 }, "a full-width space too");
  assert.deepEqual(autoCorrect("1. ", " "), { list: "number", style: "decimal", remove: 3 });
  assert.deepEqual(autoCorrect("1） ", " "), { list: "number", style: "paren", remove: 3 });
  assert.deepEqual(autoCorrect("(1) ", " "), { list: "number", style: "paren", remove: 4 });
  assert.deepEqual(autoCorrect("① ", " "), { list: "number", style: "circled", remove: 2 });
  assert.deepEqual(autoCorrect("a. ", " "), { list: "number", style: "lower-alpha", remove: 3 });
  assert.deepEqual(autoCorrect("I. ", " "), { list: "number", style: "upper-roman", remove: 3 });
  assert.equal(autoCorrect("2. ", " "), null, "a list starts at its first number");
  assert.equal(autoCorrect("売上 - ", " "), null, "only at the start of a paragraph");
  assert.equal(autoCorrect("・ ", " ", { inList: true }), null, "not inside a list");
  assert.equal(autoCorrect("・ ", " ", { options: { ...AUTOCORRECT_DEFAULTS, bullets: false } }), null);
  assert.equal(autoCorrect("1. ", " ", { options: { ...AUTOCORRECT_DEFAULTS, numbers: false } }), null);
});

test("the options are kept, each on unless turned off", () => {
  const store = new Map();
  const storage = { getItem: (k) => store.get(k) ?? null, setItem: (k, v) => store.set(k, v) };
  assert.deepEqual(autoCorrectOptions(storage), { replace: true, bullets: true, numbers: true });
  saveAutoCorrectOptions({ replace: true, bullets: false, numbers: true, other: 1 }, storage);
  assert.deepEqual(JSON.parse(store.get("hpptx-autocorrect")), { replace: true, bullets: false, numbers: true });
  assert.deepEqual(autoCorrectOptions(storage), { replace: true, bullets: false, numbers: true });
  store.set("hpptx-autocorrect", "not json");
  assert.deepEqual(autoCorrectOptions(storage), { replace: true, bullets: true, numbers: true });
});
