import assert from "node:assert/strict";
import test from "node:test";

import { hasSlidePicture, picturePlan } from "../public/auto-images.mjs";

test("picture plan fills about a third of a deck and favors slides with photo slots", () => {
  const slides = [
    { type: "title", title: "表紙" },
    { type: "imageText", image: { chartType: "bar", data: {} } },
    { type: "content", title: "現場" },
    { type: "table", title: "数値" },
    { type: "content", visualAsset: "storeOperations" },
    { type: "closing", title: "次の行動" },
  ];
  const plan = picturePlan({ slides });
  assert.deepEqual([plan.target, plan.present, plan.needed], [2, 1, 1]);
  assert.deepEqual(plan.candidates.slice(0, 2).map((item) => item.index), [2, 0]);
  assert.equal(plan.candidates.find((item) => item.index === 1).slotted, false);
});

test("ten slides target three pictures", () => {
  assert.equal(picturePlan({ slides: Array.from({ length: 10 }, () => ({ type: "content" })) }).target, 3);
  assert.equal(picturePlan({ slides: [] }).target, 0);
});

test("charts and videos are not counted as pictures", () => {
  assert.equal(hasSlidePicture({ image: { chartType: "bar" } }), false);
  assert.equal(hasSlidePicture({ media: { src: "idb:video", kind: "video" } }), false);
  assert.equal(hasSlidePicture({ media: { src: "idb:photo", kind: "image" } }), true);
});
