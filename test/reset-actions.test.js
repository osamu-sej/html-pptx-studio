import assert from "node:assert/strict";
import test from "node:test";
import { resetDeckActions, STILL_MOTION } from "../public/reset-actions.mjs";
import { deckShape } from "../server/schemas.mjs";

test("whole-deck reset removes presentation actions while retaining imported content and page states", () => {
  const source = { title: "資料", theme: "clarity", transition: "wipe", motion: { entrance: "rise", ambient: true }, slides: [
    { type: "blank", title: "表紙", master: "source", hidden: true, notes: "残すノート", animation: "click", timeline: [{ el: "i1", id: "a1", cls: "in", fx: "fade" }],
      media: { src: "https://www.youtube.com/watch?v=example", autoplay: true },
      elements: [{ id: "i1", kind: "image", src: "data:image/png;base64,AA==", action: { type: "slide", to: "p2" }, item: "items[0]", hover: "lift", tip: "説明", loop: "float" }] },
    { type: "content", title: "本文", notes: "残す本文", drillOf: "items[0]", details: [{ target: "takeaway", text: "根拠" }], photoMotion: "zoom", transition: "push", advance: 5,
      media: { kind: "lottie", src: "https://example.com/animation.json", autoplay: true }, elements: [{ id: "m1", kind: "video", src: "https://example.com/video.mp4", autoplay: true }] },
  ] };
  const original = structuredClone(source);
  const { deck, removed, changed } = resetDeckActions(source);
  assert.equal(changed, true);
  assert.deepEqual(source, original, "reset never mutates the original undo snapshot");
  assert.deepEqual(removed, { slides: 2, timeline: 1, links: 2, interactions: 3, details: 1, deepDives: 1, mediaAutoplay: 3 });
  assert.equal(deck.transition, "none");
  assert.deepEqual({ ...deck.motion }, { ...STILL_MOTION });
  for (const slide of deck.slides) {
    assert.equal(slide.timeline, undefined);
    assert.equal(slide.animation, "none");
    assert.equal(slide.transition, "none");
    assert.equal(slide.advance, undefined);
    assert.equal(slide.drillOf, undefined);
    assert.equal(slide.details, undefined);
  }
  assert.equal(deck.slides[0].elements[0].src, source.slides[0].elements[0].src);
  assert.equal(deck.slides[0].elements[0].action, undefined);
  assert.equal(deck.slides[0].elements[0].item, undefined);
  assert.equal(deck.slides[0].elements[0].hover, undefined);
  assert.equal(deck.slides[0].elements[0].tip, undefined);
  assert.equal(deck.slides[0].elements[0].loop, undefined);
  assert.equal(deck.slides[0].hidden, true);
  assert.equal(deck.slides[0].master, "source");
  assert.equal(deck.slides[0].notes, "残すノート");
  assert.equal(deck.slides[1].media.autoplay, false);
  assert.equal(deck.slides[0].media.autoplay, false);
  assert.equal(deck.slides[1].elements[0].autoplay, false);
  assert.doesNotThrow(() => deckShape.parse(deck));
  assert.equal(resetDeckActions(deck).changed, false, "a second reset is a no-op");
});

test("whole-deck reset also takes off マウスの通過 actions", () => {
  const source = { title: "資料", theme: "clarity", transition: "fade", slides: [{ type: "blank", title: "表紙", elements: [{ id: "s1", kind: "shape", x: 0, y: 0, w: 10, h: 10, overAction: { type: "next", sound: "chime" } }] }] };
  const { deck, removed } = resetDeckActions(source);
  assert.equal(deck.slides[0].elements[0].overAction, undefined);
  assert.equal(removed.links, 1);
});
