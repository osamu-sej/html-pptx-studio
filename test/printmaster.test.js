// 配布資料マスター・ノート マスター (public/editor/printmaster.mjs): the page and its placeholders, kept sparse, resolved with defaults.
import assert from "node:assert/strict";
import test from "node:test";

import { handoutBoxes, PAGE } from "../public/editor/print.mjs";
import { HANDOUT_LAYOUTS, masterSlides, normalizeMaster, normalizeMasters, notesBoxes, pageOf, placeholderBox, printDefaults, resolveMaster } from "../public/editor/printmaster.mjs";

test("the default page is the one the printouts have always had", () => {
  assert.deepEqual(pageOf("portrait"), { w: PAGE.w, h: PAGE.h });
  assert.deepEqual(pageOf("landscape"), { w: PAGE.h, h: PAGE.w });
  assert.deepEqual(pageOf("x"), pageOf("portrait"));
  assert.deepEqual(placeholderBox("header", "portrait"), { x: 48, y: 22, w: 349, align: "left" });
  assert.deepEqual(placeholderBox("pageNo", "portrait"), { x: 397, y: 1123 - 20 - 15, w: 349, align: "right" });
  const portrait = notesBoxes("portrait");
  assert.deepEqual(portrait.slide, { x: 78, y: 70, w: 638 });
  assert.deepEqual(portrait.notes, { x: 78, y: 470, w: 638, h: 593, fs: 14 });
  const landscape = notesBoxes("landscape");
  assert.ok(landscape.slide.x + landscape.slide.w < landscape.notes.x, "the slide is left of the notes");
  assert.ok(landscape.notes.x + landscape.notes.w <= pageOf("landscape").w);
});

test("a master with nothing changed is nothing", () => {
  for (const kind of ["handout", "notes"]) {
    assert.equal(normalizeMaster(null, kind), null);
    assert.equal(normalizeMaster({}, kind), null);
    assert.equal(normalizeMaster({ orientation: "portrait", fs: 11, header: { on: true }, date: { on: true }, footer: { on: false }, pageNo: { on: true, x: 397, y: 1088, w: 349, align: "right" } }, kind), null, "the same as the default is not kept");
  }
  assert.equal(normalizeMasters({ handout: {}, notes: { header: { on: true } } }), null);
});

test("what differs is kept and checked", () => {
  const m = normalizeMaster({ orientation: "landscape", perPage: "h4", fs: 99, header: { on: false, text: "  社外秘  ", x: -5, y: 40.4, w: 5000, align: "middle" }, footer: { on: true, text: "ACME" }, junk: 1 }, "handout");
  assert.deepEqual(m, { orientation: "landscape", perPage: "h4", fs: 24, header: { on: false, text: "  社外秘  ", x: 0, y: 40, w: 1123 }, footer: { on: true, text: "ACME" } });
  assert.equal(normalizeMaster({ perPage: "h5" }, "handout"), null, "only the layouts there are");
  assert.equal(normalizeMaster({ perPage: "h4" }, "notes"), null, "notes pages have no slides-per-page");
  assert.equal(normalizeMaster({ fs: 3 }, "handout").fs, 8);
  const notes = normalizeMaster({ slide: { x: 10, y: 20, w: 100 }, notes: { x: 40, y: 500, w: 700, h: 9999, fs: 18 } }, "notes");
  assert.deepEqual(notes, { slide: { x: 10, y: 20, w: 120 }, notes: { x: 40, y: 500, w: 700, h: 1123, fs: 18 } });
  assert.equal(normalizeMaster({ slide: { x: 10 } }, "handout"), null, "a handout has no slide box of its own");
  // The page's own limits depend on its orientation.
  assert.equal(normalizeMaster({ orientation: "landscape", header: { x: 1000 } }, "handout").header.x, 1000);
  assert.equal(normalizeMaster({ header: { x: 1000 } }, "handout").header.x, 754, "inside a portrait page");
});

test("pictures and words on the page go through the studio's own check and are limited", () => {
  const seen = [];
  const m = normalizeMaster({ objects: Array.from({ length: 40 }, (_, i) => ({ id: `o${i}`, kind: "text" })) }, "notes", { objects: (list) => { seen.push(list.length); return list.filter((_, i) => i % 2 === 0); } });
  assert.equal(seen[0], 30, "at most 30 are looked at");
  assert.equal(m.objects.length, 15);
  assert.equal(normalizeMaster({ objects: [{ id: "x" }] }, "notes", { objects: () => [] }), null, "none left is nothing");
});

test("resolveMaster fills in every default, in the page's own size", () => {
  const r = resolveMaster(null, "handout");
  assert.deepEqual(r.page, pageOf("portrait"));
  assert.deepEqual(Object.fromEntries(Object.entries(r.ph).map(([k, v]) => [k, v.on])), { header: true, date: true, footer: false, pageNo: true });
  assert.equal(r.ph.header.x, 48);
  assert.equal(r.slide, undefined, "handouts have no notes layout");
  const n = resolveMaster({ orientation: "landscape", header: { text: "H", x: 100 }, slide: { w: 400 }, notes: { fs: 18 } }, "notes");
  assert.deepEqual(n.page, pageOf("landscape"));
  assert.deepEqual([n.ph.header.text, n.ph.header.x, n.ph.header.y], ["H", 100, 22]);
  assert.equal(n.ph.date.x, 1123 - 48 - n.ph.date.w, "the date follows the page's right edge");
  assert.deepEqual([n.slide.w, n.slide.x], [400, 48]);
  assert.deepEqual([n.notes.fs, n.notes.y], [18, 70]);
});

test("the print dialog starts from the master's header, date, footer and page number", () => {
  const r = resolveMaster({ header: { on: false, text: "社内限り" }, footer: { on: true, text: "ACME" } }, "handout");
  assert.deepEqual(printDefaults(r), { header: false, headerText: "社内限り", date: true, pageNo: true, footer: true, footerText: "ACME" });
});

test("a master's pictures are found like a slide's (so they travel with the deck)", () => {
  const masters = { handout: { objects: [{ id: "a", kind: "image", src: "idb:logo" }] }, notes: { fs: 12 } };
  assert.deepEqual(masterSlides(masters), [{ elements: masters.handout.objects }]);
  assert.deepEqual(masterSlides(undefined), []);
});

test("handout slides sit inside a landscape page too", () => {
  for (const layout of Object.keys(HANDOUT_LAYOUTS)) {
    const page = pageOf("landscape");
    for (const b of handoutBoxes(layout, HANDOUT_LAYOUTS[layout], { page })) {
      assert.ok(b.x >= 0 && b.y >= 0 && b.x + b.w <= page.w && b.y + b.h <= page.h, `${layout} inside the page: ${JSON.stringify(b)}`);
    }
  }
  const portrait = handoutBoxes("h1", 1);
  const landscape = handoutBoxes("h1", 1, { page: pageOf("landscape") });
  assert.ok(landscape[0].w > portrait[0].w && landscape[0].h > portrait[0].h, "a single slide is larger on the wide page");
  assert.deepEqual(handoutBoxes("h6", 6), handoutBoxes("h6", 6, { page: PAGE }), "the default is the portrait page");
  const across = (boxes) => new Set(boxes.map((b) => b.x)).size;
  assert.equal(across(handoutBoxes("h6", 6)), 2, "6 slides: 2 across on a tall page");
  assert.equal(across(handoutBoxes("h6", 6, { page: pageOf("landscape") })), 3, "and 3 across on a wide one");
  assert.equal(across(handoutBoxes("h2", 2, { page: pageOf("landscape") })), 2, "2 slides side by side");
  assert.equal(across(handoutBoxes("h9", 9, { page: pageOf("landscape") })), 3);
});
