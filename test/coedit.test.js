// 共同編集 merges (coedit.mjs): two people on one slide keep each other's objects; the order of slides merges too.
import assert from "node:assert/strict";
import test from "node:test";

import { mergeOrder, mergeSlide } from "../public/editor/coedit.mjs";
import { newSid, uniqueSids } from "../public/editor/coedit-merge.mjs";

const slide = (elements, extra = {}) => ({ type: "blank", sid: "s1", title: "本文", elements, ...extra });
const box = (id, x, text = "") => ({ id, kind: "shape", x, y: 0, w: 100, h: 50, text });

test("a slide changed only by them, or only here, is simply theirs, or ours", () => {
  const base = slide([box("a", 0)]);
  const theirs = slide([box("a", 50)]);
  assert.deepEqual(mergeSlide(base, base, theirs), theirs);
  const mine = slide([box("a", 10)]);
  assert.deepEqual(mergeSlide(base, mine, base), mine);
});

test("both on one slide: their moved object and my new one and my edited text all stay", () => {
  const base = slide([box("a", 0), box("b", 200, "旧")]);
  const theirs = slide([box("a", 80), box("b", 200, "旧")], { title: "相手の題" });
  const mine = slide([box("a", 0), box("b", 200, "新しい文字"), box("c", 400)], { notes: "私のノート" });
  const out = mergeSlide(base, mine, theirs);
  assert.equal(out.elements.find((o) => o.id === "a").x, 80, "their move");
  assert.equal(out.elements.find((o) => o.id === "b").text, "新しい文字", "my text");
  assert.ok(out.elements.some((o) => o.id === "c"), "my new object");
  assert.equal(out.title, "相手の題");
  assert.equal(out.notes, "私のノート");
  // Deleted here: gone, even though they still have it.
  const deleted = mergeSlide(base, slide([box("a", 0)]), theirs);
  assert.deepEqual(deleted.elements.map((o) => o.id), ["a"]);
});

test("slide order: their move with my new slide; my deletion wins over their old order", () => {
  const base = ["s1", "s2", "s3", "s4"];
  assert.deepEqual(mergeOrder(base, base, ["s1", "s3", "s2", "s4"]), ["s1", "s3", "s2", "s4"]);
  assert.deepEqual(mergeOrder(base, ["s1", "s2", "n1", "s3", "s4"], ["s1", "s3", "s2", "s4"]), ["s1", "s3", "s2", "n1", "s4"], "my new slide after s2");
  assert.deepEqual(mergeOrder(base, ["s1", "s3", "s4"], ["s1", "s3", "s2", "s4", "t1"]), ["s1", "s3", "s4", "t1"], "s2 deleted here, their new t1 kept");
});

test("slide ids: many made in the same millisecond never repeat, and a repeated or missing one is replaced", () => {
  const taken = new Set();
  for (let i = 0; i < 5000; i += 1) {
    const sid = newSid(taken);
    assert.match(sid, /^[A-Za-z0-9_-]{1,32}$/, "what the rooms accept");
    assert.ok(!taken.has(sid), `repeated ${sid}`);
    taken.add(sid);
  }
  const slides = [{ sid: "a" }, {}, { sid: "a" }, { sid: "b" }, {}];
  uniqueSids(slides);
  assert.equal(slides[0].sid, "a", "the first keeps its id");
  assert.equal(slides[3].sid, "b");
  assert.equal(new Set(slides.map((s) => s.sid)).size, 5, "all different");
  assert.ok(slides.every((s) => /^[A-Za-z0-9_-]{1,32}$/.test(s.sid)));
});
