import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { applyChatOperations } from "../server/chat.mjs";
import { objectsSummary, plainText } from "../server/objects.mjs";
import { ALL_SLIDE_TYPES, SLIDE_TYPES, codexChatSchema, codexDeckSchema, codexSlideSchema, codexVariantsSchema, deckShape } from "../server/schemas.mjs";
import { slideMeaning } from "../server/visual-relevance.mjs";
import { buildChatPrompt, withoutImageData } from "../server/codex-app-server.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));

async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "animate.js"), "utf8"), context, { filename: "animate.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "motion.js"), "utf8"), context, { filename: "motion.js" });
  return { E: window.SlideEngine, window };
}

const deckWith = (slide) => ({ title: "検証", theme: "clarity", transition: "fade", motion: {}, slides: [{ type: "title", title: "表紙" }, slide, { type: "closing" }] });

test("every shape draws a clean outline at any size, with handles that read back what they set", async () => {
  const { E } = await loadEngine();
  const keys = Object.keys(E.SHAPES);
  assert.ok(keys.length >= 110, `PowerPoint-sized catalog (${keys.length})`);
  assert.deepEqual([...new Set(E.SHAPE_GROUPS.flatMap(([, list]) => list))].sort(), [...keys].sort(), "every shape is in the gallery once");
  for (const key of keys) {
    const shape = E.SHAPES[key];
    assert.ok(shape.label, key);
    for (const [w, h] of [[300, 200], [60, 420], [1, 1], [1600, 40]]) {
      const g = E.geometry(key, w, h, shape.adj);
      for (const d of [...g.paths, ...g.extras.map((x) => x.d)]) assert.doesNotMatch(d, /NaN|undefined|Infinity/, `${key} ${w}×${h}`);
      assert.ok(g.text.every(Number.isFinite), `${key}: text box`);
      for (const handle of shape.handles ?? []) {
        const [x, y] = handle.pos(w, h, g.adj);
        assert.ok(Number.isFinite(x) && Number.isFinite(y), `${key}: handle`);
        const back = handle.set(x, y, w, h, g.adj);
        assert.equal(back.length, g.adj.length, `${key}: a handle sets every adjustment`);
        assert.ok(back.every(Number.isFinite), `${key}: handle value`);
      }
    }
  }
  // Out-of-range adjustments are clamped, missing ones take the default.
  assert.deepEqual([...E.adjOf(E.SHAPES.roundRect, [9])], [0.5]);
  assert.deepEqual([...E.adjOf(E.SHAPES.roundRect, [])], [0.1667]);
});

test("rich text keeps paragraphs, lists and marks, and nothing that could run", async () => {
  const { E } = await loadEngine();
  const html = E.sanitizeRich('<div>A<b>太字</b><script>alert(1)</script><img src=x onerror=alert(1)><span style="color: rgb(31,56,100); font-size: 24pt; position: fixed" onclick="x()">紺</span></div><ul><li>一<ul><li>二</li></ul></li></ul><a href="javascript:alert(1)">危険</a><a href="https://example.com" style="color:red">安全</a><iframe src="https://evil"></iframe><font size="7" color="#808080">大</font><p style="text-align: center; margin: 99px">中央</p>');
  assert.doesNotMatch(html, /script|onerror|onclick|javascript:|iframe|position|margin|<img/i);
  assert.match(html, /<b>太字<\/b>/);
  assert.match(html, /<span style="color: #1f3864; font-size: 48px">紺<\/span>/, "points become pixels (1pt = 2px), colours become hex");
  assert.match(html, /<ul><li>一<ul><li>二<\/li><\/ul><\/li><\/ul>/);
  assert.match(html, /<a rel="noopener noreferrer" target="_blank" href="https:\/\/example.com">安全<\/a>/);
  assert.match(html, />危険</, "an unsafe link keeps its words");
  assert.match(html, /<span style="color: #808080; font-size: 96px">大<\/span>/, "Chrome's <font> from the editor becomes a span");
  assert.match(html, /<p style="text-align: center">中央<\/p>/);
  assert.equal(E.sanitizeRich(html), html, "sanitizing is stable");
  assert.equal(E.richToText("<p>一行目<br>改行</p><ul><li>A<ul><li>B</li></ul></li></ul>"), "一行目\n改行\nA\nB");
  assert.equal(E.textToRich("a<b>\n\nc"), "<p>a&lt;b&gt;</p><p><br></p><p>c</p>");
});

test("objects are normalized: known kinds only, numbers clamped, links and media checked", async () => {
  const { E } = await loadEngine();
  const list = E.normalizeObjects([
    { id: "a", kind: "shape", shape: "nope", x: "12.345", y: 1e9, w: -5, h: 200, rot: 370, fill: "red", stroke: "#1F3864", text: "<p>hi</p><script>x</script>", evil: 1, fs: 9999 },
    { id: "a", kind: "text", x: 0, y: 0, w: 100, h: 50, action: { type: "url", href: "javascript:alert(1)" } },
    { kind: "image", src: "javascript:alert(1)", x: 0, y: 0, w: 1, h: 1 },
    { kind: "image", src: "idb:abc", x: 0, y: 0, w: 100, h: 100, crop: { l: 0.2, t: 0, r: 0.9, b: 0 } },
    { kind: "line", x1: 0, y1: 0, x2: 100, y2: 0, from: { id: "zz", site: 1 }, tail: "triangle", head: "bogus" },
    { kind: "icon", icon: "no-such-icon", x: 0, y: 0, w: 10, h: 10 },
    { kind: "frobnicator" },
    { id: "b", kind: "shape", x: 0, y: 0, w: 10, h: 10, action: { type: "slide", to: "s1" } },
  ]);
  assert.equal(list.length, 5);
  const [shape, text, image, line, linked] = list;
  assert.equal(shape.shape, "rect");
  assert.equal(shape.x, 12.35);
  assert.equal(shape.w, 1);
  assert.equal(shape.rot, 10);
  assert.equal(shape.fill, "#ff0000");
  assert.equal(shape.stroke, "#1f3864");
  assert.equal(shape.fs, 400);
  assert.equal(shape.text, "<p>hi</p>");
  assert.equal(shape.evil, undefined);
  assert.notEqual(text.id, "a", "ids are unique");
  assert.equal(text.action, undefined, "only web and mail links");
  assert.equal(image.crop, undefined, "a crop that leaves nothing is dropped");
  assert.equal(line.from, undefined, "a connector forgets objects that are not there");
  assert.equal(line.tail, "triangle");
  assert.equal(line.head, undefined);
  assert.deepEqual({ ...linked.action }, { type: "slide", to: "s1" });
});

test("objects render on their own layer, with text, pictures and lines", async () => {
  const { E } = await loadEngine();
  const slide = {
    type: "blank", title: "",
    elements: E.normalizeObjects([
      { id: "box", kind: "shape", shape: "roundRect", x: 100, y: 200, w: 400, h: 200, rot: 30, flipH: true, text: "<p>テキスト</p>", fill: "#dce4f2" },
      { id: "t", kind: "text", x: 600, y: 200, w: 400, h: 100, text: "<p>文字</p>", vertical: true },
      { id: "pic", kind: "image", src: "asset:ai", x: 1000, y: 300, w: 300, h: 200, mask: "ellipse" },
      { id: "l", kind: "line", x1: 0, y1: 0, x2: 0, y2: 0, from: { id: "box", site: 1 }, to: { id: "pic", site: 3 }, route: "elbow", tail: "triangle" },
      { id: "gone", kind: "shape", x: 0, y: 0, w: 10, h: 10, hidden: true },
    ]),
  };
  const el = E.render(slide, { deck: deckWith(slide), index: 1, mode: "edit", assetBase: "/assets/" });
  assert.equal(el.dataset.type, "blank");
  assert.equal(el.dataset.build, "none");
  const layer = el.querySelector(".hs-objects");
  assert.ok(layer);
  assert.equal(layer.querySelectorAll(".hs-obj").length, 4, "hidden objects are not drawn");
  assert.deepEqual([...layer.children].map((node) => node.dataset.el), ["box", "t", "pic", "l"], "drawing order is the list order");
  const box = layer.querySelector('[data-el="box"]');
  assert.match(box.querySelector(".hs-obj-rot").getAttribute("style"), /rotate\(30deg\) scale\(-1, 1\)/);
  assert.match(box.querySelector(".hs-obj-text").getAttribute("style"), /scale\(-1, 1\)/, "flipped text stays readable");
  assert.equal(box.dataset.fill, "#dce4f2");
  assert.ok(layer.querySelector('[data-el="t"] .hs-obj-text.is-vertical'));
  assert.match(layer.querySelector('[data-el="pic"] img').getAttribute("src"), /\/assets\/ai-executive-hero\.jpg/);
  assert.match(layer.querySelector('[data-el="pic"] .hs-obj-img').getAttribute("style"), /clip-path:\s*path\(/);
  // The connector runs from the box's right side to the picture's left side.
  const [p1, p2] = E.lineEnds(slide.elements[3], slide.elements);
  const sites = E.sites(slide.elements[0]);
  assert.deepEqual([...p1], [...sites[1]]);
  assert.deepEqual([...p2], [1000, 400]);
  assert.equal(layer.querySelectorAll('[data-el="l"] path').length, 2, "line and arrowhead");
  // The title placeholder only shows in the editor.
  assert.equal(el.querySelector(".hs-title").dataset.placeholder, "タイトルを入力");
  assert.equal(E.render(slide, { deck: deckWith(slide), index: 1, mode: "present" }).querySelector(".hs-title").dataset.placeholder, undefined);
  // Every layout can carry objects.
  const content = { type: "content", title: "本文", points: ["A"], elements: [{ id: "x", kind: "shape", x: 0, y: 0, w: 10, h: 10 }] };
  assert.ok(E.render(content, { deck: deckWith(content), index: 1, mode: "thumb" }).querySelector('.hs-objects [data-el="x"]'));
});

test("names, words and bounds of objects", async () => {
  const { E } = await loadEngine();
  assert.equal(E.objectName({ kind: "shape", shape: "rect" }, 2), "正方形/長方形 3");
  assert.equal(E.objectName({ kind: "text" }, 0), "テキスト ボックス 1");
  assert.equal(E.objectName({ kind: "line", tail: "triangle" }, 0), "矢印 1");
  assert.equal(E.objectName({ kind: "shape", name: "見出し" }, 0), "見出し");
  assert.equal(E.objectText({ kind: "shape", text: "<p>一</p><p>二</p>" }), "一\n二");
  const b = E.bounds({ kind: "shape", x: 0, y: 0, w: 100, h: 100, rot: 45 });
  assert.ok(Math.abs(b.w - 141.42) < 0.1 && Math.abs(b.x + 20.71) < 0.1, "a rotated box's upright bounds");
});

test("the server keeps objects people placed: schemas, AI summaries and every AI change", () => {
  // People may make 白紙 pages; the AI writes the 44 layouts only and never objects.
  assert.ok(ALL_SLIDE_TYPES.includes("blank"));
  assert.ok(!SLIDE_TYPES.includes("blank"));
  for (const schema of [codexDeckSchema, codexChatSchema, codexSlideSchema, codexVariantsSchema]) {
    const text = JSON.stringify(schema);
    assert.doesNotMatch(text, /"(elements|timeline|sid)":|"const":"blank"/);
  }
  const elements = [{ id: "a", kind: "shape", x: 1, y: 2, w: 3, h: 4, text: "<p>売上&amp;利益</p>" }, { id: "b", kind: "image", src: "idb:1", x: 0, y: 0, w: 1, h: 1 }];
  const deck = deckShape.parse({ title: "t", slides: [{ type: "title", title: "表紙" }, { type: "blank", elements, timeline: [{ el: "a", effect: "fade" }], sid: "s1" }, { type: "closing" }] });
  assert.equal(deck.slides[1].elements[0].text, "<p>売上&amp;利益</p>", "object fields pass through");
  assert.equal(deck.slides[1].timeline.length, 1);
  assert.equal(plainText("<p>一</p><ul><li>二&amp;三</li></ul>"), "一\n二&三");
  assert.equal(objectsSummary(elements), "[自由配置: 図形1・画像1／文字「売上&利益」]");
  const forAi = withoutImageData(deck.slides);
  assert.equal(forAi[1].elements, "[自由配置: 図形1・画像1／文字「売上&利益」]");
  assert.equal(forAi[1].timeline, "[アニメーション1件]");
  assert.equal(forAi[1].sid, undefined);
  assert.match(slideMeaning(deck.slides[1]), /売上&利益/);
  assert.doesNotMatch(slideMeaning(deck.slides[1]), /idb:|shape/);
  const prompt = buildChatPrompt({ deck, message: "全体を簡潔に" });
  assert.match(prompt, /自由配置: 図形1・画像1/);
  assert.match(prompt, /blank（白紙）のスライドはユーザーが自由に作ったページ/);
  // An AI rewrite of a slide keeps its objects, animations and id.
  const applied = applyChatOperations(deck.slides, { operations: [{ op: "replace", slide: 2, content: { type: "content", title: "整理", points: ["A"] } }] });
  assert.deepEqual(applied.slides[1].elements, deck.slides[1].elements);
  assert.equal(applied.slides[1].sid, "s1");
  assert.equal(applied.slides[1].timeline.length, 1);
});

test("an AI rewrite keeps the slide's place in the show, its section and the review comments on it", () => {
  const slides = [
    { type: "title", title: "表紙" },
    { type: "content", title: "現状", points: ["A"], hidden: true, section: "背景", comments: [{ id: "c1", text: "数字を確認", at: "2026-10-03T09:00:00Z" }] },
    { type: "closing", title: "おわり" },
  ];
  const applied = applyChatOperations(slides, { operations: [{ op: "replace", slide: 2, content: { type: "content", title: "現状の課題", points: ["A", "B"] } }] });
  assert.equal(applied.slides[1].title, "現状の課題");
  assert.equal(applied.slides[1].hidden, true, "a hidden slide stays out of the show");
  assert.equal(applied.slides[1].section, "背景");
  assert.deepEqual(applied.slides[1].comments, slides[1].comments);
  // The AI is never asked for them (as properties: "section" is also the name of a slide type).
  for (const schema of [codexDeckSchema, codexChatSchema]) assert.doesNotMatch(JSON.stringify(schema), /"(section|comments|hidden)":\{/, "the AI never writes them");
  // The deck people save keeps them.
  const saved = deckShape.parse({ title: "t", slides });
  assert.equal(saved.slides[1].section, "背景");
  assert.equal(saved.slides[1].comments[0].text, "数字を確認");
});
