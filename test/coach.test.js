// コーチによるリハーサル (public/editor/coach.mjs): fillers, words said twice, reading the slide out, pace, the report;
// and 画面切り替えの効果のオプション: every option the engine offers is one the deck format accepts.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { coachReport, countFillers, readingShare, repeats, slideWords } from "../public/editor/coach.mjs";
import { deckShape } from "../server/schemas.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));

test("filler words are counted once each, the longer word first; English fillers only as words", () => {
  const counts = countFillers("えーと、本日は、えー、あのー、売上について。Um, the umbrella, uh");
  assert.deepEqual(Object.fromEntries(counts), { えーと: 1, えー: 1, あのー: 1, um: 1, uh: 1 });
  assert.equal(countFillers("まとめます").size, 0);
});

test("words said twice with a pause count as a repetition; words that double by nature do not", () => {
  assert.deepEqual(Object.fromEntries(repeats("その その計画を、the the plan")), { その: 1, the: 1 });
  assert.equal(repeats("いろいろな案を、ますます広げる").size, 0);
});

test("reading the slide's own words out is noticed; talking around it is not", () => {
  const shown = "店舗の発注作業を30分短縮し、廃棄を12%減らしました";
  assert.ok(readingShare("店舗の発注作業を30分短縮し、廃棄を12%減らしました", shown) > 0.8);
  assert.ok(readingShare("ここで大事なのは、現場の負担が減ったことです。時間の余裕が接客に回りました", shown) < 0.2);
  assert.equal(readingShare("短い", shown), 0, "too little said to tell");
  const words = slideWords({ type: "statement", title: "結論", text: shown, sid: "s1", elements: [{ kind: "text", text: "<p>補足の文字</p>" }] });
  assert.match(words, /結論/);
  assert.match(words, /補足の文字/);
  assert.doesNotMatch(words, /statement|s1/);
});

test("the report: pace from characters a minute, fillers, repeats, slides read out, and tips", () => {
  const deck = { slides: [{ type: "title", title: "発注の改革" }, { type: "statement", text: "店舗の発注作業を30分短縮し、廃棄を12%減らしました" }] };
  const said = [
    { text: "えーと、本日は、えー、発注の改革について、あのー、その その成果をご報告します", slide: 0 },
    { text: "店舗の発注作業を30分短縮し、廃棄を12%減らしました", slide: 1 },
  ];
  const times = new Map([[0, 9000], [1, 6000]]);
  const r = coachReport({ said, times, deck });
  assert.equal(r.ms, 15000);
  assert.ok(r.cpm > 100 && r.cpm < 260, `characters a minute: ${r.cpm}`);
  assert.equal(r.pace, "slow");
  assert.equal(r.fillerCount, 3);
  assert.equal(r.repeats[0].word, "その");
  assert.deepEqual(r.slides.map((s) => s.reading), [false, true], "slide 2 was read out");
  assert.ok(r.tips.some((t) => /2枚目/.test(t) && /読み上げ/.test(t)));
  assert.ok(r.tips.some((t) => /つなぎ言葉/.test(t)));
  // Fast and clean.
  const quick = coachReport({ said: [{ text: "あ".repeat(500), slide: 0 }], times: new Map([[0, 60000]]) });
  assert.equal(quick.pace, "fast");
  const fine = coachReport({ said: [{ text: "今期の成果と来期の計画をご説明します。".repeat(16), slide: 0 }], times: new Map([[0, 60000]]) });
  assert.equal(fine.pace, "good");
  assert.ok(fine.tips.some((t) => /よいペース/.test(t)));
});

test("every transition option the engine offers is one the deck format accepts", async () => {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const code = await readFile(join(root, "public", "engine", "engine.js"), "utf8");
  vm.runInContext(code.replace("/*__ICONS__*/{}", () => icons), vm.createContext(window), { filename: "engine.js" });
  const E = window.SlideEngine;
  for (const [type, list] of Object.entries(E.TRANSITION_OPTIONS)) {
    assert.ok(E.TRANSITIONS[type], `${type} is a transition`);
    for (const [dir] of list) assert.equal(deckShape.parse({ title: "t", slides: [{ type: "title", title: "x", transition: type, transitionDir: dir }] }).slides[0].transitionDir, dir);
  }
  assert.throws(() => deckShape.parse({ title: "t", slides: [{ type: "title", title: "x", transitionDir: "diagonal" }] }));
});
