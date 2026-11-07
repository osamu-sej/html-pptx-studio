// Sounds on slides (挿入 → オーディオ) and the 再生 settings of sounds and videos: kept, drawn, played.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { objectsSummary } from "../server/objects.mjs";
import { objectSchema } from "../server/schemas.mjs";
import { resetDeckActions } from "../public/reset-actions.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));

async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "animate.js"), "utf8"), context, { filename: "animate.js" });
  return { E: window.SlideEngine };
}

test("a sound keeps its playback settings; nonsense is dropped", async () => {
  const { E } = await loadEngine();
  const o = E.normalizeObject({ id: "s1", kind: "audio", x: 900, y: 480, w: 120, h: 120, src: "data:audio/webm;base64,AAAA", fileName: "録音.webm",
    autoplay: true, loop: true, across: "all", hideIcon: true, rewind: true, volume: 0.33, trimStart: 1.5, trimEnd: 9, fadeIn: 2, fadeOut: 99, color: "#1f3864", fullscreen: true });
  assert.equal(o.kind, "audio");
  assert.equal(o.across, 999);
  assert.equal(o.hideIcon, true);
  assert.equal(o.volume, 0.33);
  assert.deepEqual([o.trimStart, o.trimEnd, o.fadeIn, o.fadeOut], [1.5, 9, 2, 60]);
  assert.equal(o.fullscreen, undefined, "全画面 is for videos");
  assert.equal(E.normalizeObject({ id: "s2", kind: "audio", src: "javascript:alert(1)" }), null, "only real sources");
  const bad = E.normalizeObject({ id: "s3", kind: "audio", src: "https://example.com/a.mp3", trimStart: 5, trimEnd: 4, volume: 1, across: 1 });
  assert.equal(bad.trimEnd, undefined, "an end before the start is dropped");
  assert.equal(bad.volume, undefined, "full volume is the default");
  assert.equal(bad.across, undefined, "one slide is not across slides");
  const video = E.normalizeObject({ id: "v1", kind: "video", src: "https://example.com/a.mp4", fullscreen: true, hideIdle: true, hideIcon: true });
  assert.equal(video.fullscreen, true);
  assert.equal(video.hideIdle, true);
  assert.equal(video.hideIcon, undefined);
});

test("a sound is a speaker icon; presenting adds the sound, its settings and a play bar", async () => {
  const { E } = await loadEngine();
  const sound = { id: "s1", kind: "audio", x: 900, y: 480, w: 120, h: 120, src: "https://example.com/a.mp3", autoplay: true, loop: true, across: 999, trimStart: 2, fadeIn: 1, hideIcon: true };
  const slide = { type: "blank", title: "音", elements: [sound] };
  const edit = E.render(slide, { mode: "edit", index: 1, deck: { slides: [slide], theme: "clarity" } });
  assert.ok(edit.querySelector('.hs-obj[data-kind="audio"] .hs-audio-icon'), "the icon");
  assert.equal(edit.querySelector("audio"), null, "no sound while editing");
  const live = E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  const audio = live.querySelector("audio");
  assert.ok(audio, "the sound");
  assert.equal(audio.getAttribute("src"), "https://example.com/a.mp3");
  for (const attr of ["data-autoplay", "data-loop", "data-across", "data-trim-start", "data-fade-in"]) assert.ok(audio.hasAttribute(attr), attr);
  assert.equal(audio.getAttribute("data-across"), "999");
  assert.ok(live.querySelector(".hs-audio[data-hide-icon]"), "the icon hides while presenting");
  assert.ok(live.querySelector(".hs-audio-bar .hs-audio-play"), "the play bar");
});

test("再生・一時停止・停止 animations work on a sound; the server and the reset know sounds", async () => {
  const { E } = await loadEngine();
  const slide = { type: "blank", elements: [{ id: "s1", kind: "audio", x: 0, y: 0, w: 120, h: 120, src: "https://example.com/a.mp3" }, { id: "t1", kind: "text", x: 0, y: 0, w: 300, h: 100, text: "<p>a</p>" }] };
  const list = E.normalizeTimeline([{ el: "s1", cls: "media", fx: "play", start: "click" }, { el: "t1", cls: "media", fx: "play" }], slide);
  assert.deepEqual(Array.from(list, (e) => `${e.el}:${e.fx}`), ["s1:play"], "media effects only on media");
  assert.ok(objectSchema.safeParse({ id: "s1", kind: "audio" }).success, "the server accepts a sound");
  assert.match(objectsSummary(slide.elements), /オーディオ1/);
  const deck = { slides: [{ type: "title", title: "表紙" }, { type: "blank", elements: [{ id: "s1", kind: "audio", src: "x", autoplay: true, across: 999 }] }] };
  const { deck: reset } = resetDeckActions(deck);
  assert.equal(reset.slides[1].elements[0].autoplay, false);
  assert.equal(reset.slides[1].elements[0].across, undefined);
});

test("ブックマーク: kept in time order with ids and names; a jump list for the show", async () => {
  const { E } = await loadEngine();
  const o = E.normalizeObject({ id: "v1", kind: "video", x: 0, y: 0, w: 640, h: 360, src: "https://example.com/v.mp4",
    bookmarks: [{ t: 12.346, name: "  まとめ  " }, { id: "m1", t: 3, name: "" }, { id: "m1", t: 7 }, { t: "x" }, null, { id: "bad id!", t: 99999999 }] });
  assert.deepEqual(JSON.parse(JSON.stringify(o.bookmarks.map((b) => b.t))), [3, 7, 12.35, 86400], "sorted, rounded, in range, nonsense dropped");
  assert.equal(o.bookmarks[0].id, "m1");
  assert.notEqual(o.bookmarks[1].id, "m1", "ids stay unique");
  assert.equal(o.bookmarks[0].name, "ブックマーク 1", "an empty name gets one");
  assert.equal(o.bookmarks[2].name, "まとめ");
  assert.match(o.bookmarks[3].id, /^[A-Za-z0-9_-]{1,16}$/);
  assert.equal(JSON.stringify(E.normalizeObject({ ...o }).bookmarks), JSON.stringify(o.bookmarks), "normalizing again changes nothing");
  assert.equal(E.normalizeObject({ id: "v2", kind: "video", src: "https://example.com/v.mp4", bookmarks: [] }).bookmarks, undefined);
  const slide = { type: "blank", elements: [o] };
  const live = E.render(slide, { mode: "present", index: 0, deck: { slides: [slide], theme: "clarity" } });
  assert.equal(live.querySelectorAll(".hs-media-marks .hs-mark").length, 4, "dots over the video while presenting");
  assert.equal(E.mediaMarks(live.querySelector("video")).length, 4);
  const edit = E.render(slide, { mode: "edit", index: 0, deck: { slides: [slide], theme: "clarity" } });
  assert.equal(edit.querySelector(".hs-media-marks"), null, "not while editing");
});
