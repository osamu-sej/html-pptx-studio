// スライド ショーの設定 and 目的別スライド ショー (show.mjs): settings checked, which slides a show plays, what the
// player is told, and the ink written in a show kept on its slides.
import assert from "node:assert/strict";
import test from "node:test";

import { customShowsOf, keptInk, playerOptions, showOf, showSlides } from "../public/editor/show.mjs";
import { inkObject } from "../public/editor/ink.mjs";
import { deckShape } from "../server/schemas.mjs";

const deck = {
  slides: [{ sid: "a" }, { sid: "b" }, { sid: "c" }, { sid: "d" }, { sid: "e" }],
  customShows: [{ id: "cs1", name: "短縮版", sids: ["d", "b"] }],
};

test("show settings keep only what PowerPoint has, and nothing when all is as it starts", () => {
  assert.equal(showOf(null), null);
  assert.equal(showOf({ loop: false, useTimings: true, penColor: "#ff0000", captionLang: "ja-JP" }), null);
  assert.deepEqual(showOf({ kiosk: true, loop: true, noAnimation: true, useTimings: false, penColor: "#0070c0", kioskSeconds: "12", range: { from: 2, to: 4 }, custom: "cs1", evil: "<x>" }),
    { kiosk: true, loop: true, noAnimation: true, useTimings: false, penColor: "#0070c0", kioskSeconds: 12, range: { from: 2, to: 4 }, custom: "cs1" });
  // Colours outside the pens offered, impossible ranges and odd ids are dropped.
  assert.equal(showOf({ penColor: "#ff00ff", range: { from: 4, to: 2 }, custom: "a b", kioskSeconds: 1 }), null);
});

test("custom shows: ids unique, slides by sid without repeats, empty ones dropped", () => {
  assert.equal(customShowsOf("x"), null);
  assert.deepEqual(customShowsOf([
    { id: "cs1", name: " 役員向け ", sids: ["d", "b", "d", 3, "bad id"] },
    { id: "cs1", name: "dup", sids: ["a"] },
    { id: "cs2", name: "", sids: [] },
    { id: "cs3", name: "", sids: ["e"] },
  ]), [{ id: "cs1", name: "役員向け", sids: ["d", "b"] }, { id: "cs3", name: "目的別スライド ショー", sids: ["e"] }]);
});

test("which slides a show plays: a custom show in its own order, a range, or all", () => {
  assert.equal(showSlides(deck, null), null);
  assert.deepEqual(showSlides(deck, { range: { from: 2, to: 4 } }), [1, 2, 3]);
  assert.deepEqual(showSlides(deck, { range: { from: 4, to: 99 } }), [3, 4]);
  assert.deepEqual(showSlides(deck, { custom: "cs1" }), [3, 1]);
  // Starting a custom show from the menu wins over the range of the settings.
  assert.deepEqual(showSlides(deck, { range: { from: 1, to: 2 } }, "cs1"), [3, 1]);
  // A custom show whose slides are gone plays everything.
  assert.equal(showSlides({ slides: [{ sid: "z" }], customShows: deck.customShows }, { custom: "cs1" }), null);
});

test("player options from the settings", () => {
  assert.deepEqual(playerOptions(null), { loop: false, kiosk: false, kioskSeconds: 8, useTimings: true, static: false, noAnimation: false, narration: true, penColor: "#ff0000", captions: false, captionLang: "ja-JP", mediaControls: true });
  const o = playerOptions({ kiosk: true, noAnimation: true, noNarration: true, useTimings: false, kioskSeconds: 5, penColor: "#000000", captions: true, captionLang: "en-US" });
  assert.equal(o.kiosk, true);
  assert.equal(o.static, true);
  assert.equal(o.noAnimation, true);
  assert.equal(o.narration, false);
  assert.equal(o.useTimings, false);
  assert.equal(o.kioskSeconds, 5);
  assert.equal(o.penColor, "#000000");
  assert.equal(o.captions, true);
  assert.equal(o.captionLang, "en-US");
});

test("ink written in a show goes back to its slide by sid (even if slides moved meanwhile)", () => {
  const stroke = (x) => ({ pts: [[x, 100], [x + 200, 160]], color: "#c00000", width: 5 });
  const now = { slides: [{ sid: "x" }, { sid: "a" }, { sid: "b" }] };
  let n = 0;
  const made = keptInk([{ index: 0, strokes: [stroke(100), stroke(400)] }, { index: 1, strokes: [] }, { index: 9, strokes: [stroke(1)] }], ["a", "b"], now, (strokes) => inkObject(strokes, `ink${++n}`));
  assert.deepEqual([...made.keys()], [1]);
  const [o] = made.get(1);
  assert.equal(o.kind, "ink");
  assert.equal(o.strokes.length, 2);
  // The box is the strokes plus half the pen width and a pixel.
  assert.ok(Math.abs(o.x - 96.5) < 0.01 && Math.abs(o.x + o.w - 603.5) < 0.01, `box ${o.x} ${o.w}`);
});

test("the server keeps the show settings and custom shows through AI edits", () => {
  const parsed = deckShape.parse({ title: "t", slides: [{ type: "blank", title: "" }], show: { loop: true }, customShows: [{ id: "cs1", name: "n", sids: ["a"] }] });
  assert.deepEqual(parsed.show, { loop: true });
  assert.deepEqual(parsed.customShows, [{ id: "cs1", name: "n", sids: ["a"] }]);
});

test("メディア コントロールを表示: on unless turned off; the player is told", () => {
  assert.equal(showOf({ noControls: false }), null, "on is the default");
  assert.equal(showOf({ noControls: true }).noControls, true);
  assert.equal(playerOptions(null).mediaControls, true);
  assert.equal(playerOptions(showOf({ noControls: true })).mediaControls, false);
});
