// Builds qa/objects.json: every shape, line and text option on slides, for `node qa/shoot.mjs --themes=clarity --deck=/qa/objects.json`.
import { writeFile } from "node:fs/promises";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { parseHTML } from "linkedom";

const root = new URL("..", import.meta.url).pathname;
const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
const context = vm.createContext(window);
vm.runInContext(readFileSync(`${root}public/engine/engine.js`, "utf8").replace("/*__ICONS__*/{}", () => readFileSync(`${root}public/engine/icons.json`, "utf8").trim()), context);
vm.runInContext(readFileSync(`${root}public/engine/objects.js`, "utf8"), context);
const E = window.SlideEngine;

const slides = [{ type: "title", title: "オブジェクトの見本" }];
const keys = E.SHAPE_GROUPS.flatMap(([, list]) => list);
const perSlide = 28;
for (let start = 0; start < keys.length; start += perSlide) {
  const elements = [];
  keys.slice(start, start + perSlide).forEach((key, i) => {
    const col = i % 7;
    const row = Math.floor(i / 7);
    const x = 90 + col * 245;
    const y = 175 + row * 200;
    elements.push({ id: `s${start + i}`, kind: "shape", shape: key, x, y, w: 150, h: 110, fill: "#dce4f2", stroke: "#1f3864", strokeW: 2 });
    elements.push({ id: `t${start + i}`, kind: "text", x: x - 40, y: y + 116, w: 230, h: 50, fs: 18, align: "center", text: `<p>${E.SHAPES[key].label}</p>`, autofit: "none" });
  });
  slides.push({ type: "blank", title: `図形 ${start + 1}〜${Math.min(keys.length, start + perSlide)}`, elements });
}
// Lines, arrowheads, dashes, routes
const lines = [];
Object.keys(E.ARROWHEADS).forEach((head, i) => {
  lines.push({ id: `l${i}`, kind: "line", x1: 120, y1: 200 + i * 90, x2: 620, y2: 200 + i * 90, tail: head, head: i % 2 ? head : "none", strokeW: 4 + i });
});
Object.keys(E.DASHES).forEach((dash, i) => lines.push({ id: `d${i}`, kind: "line", x1: 720, y1: 200 + i * 80, x2: 1180, y2: 200 + i * 80, dash, strokeW: 6, stroke: "#1a1a1a" }));
lines.push({ id: "r1", kind: "shape", shape: "rect", x: 1290, y: 200, w: 160, h: 100, text: "<p>A</p>" });
lines.push({ id: "r2", kind: "shape", shape: "ellipse", x: 1580, y: 520, w: 160, h: 120, text: "<p>B</p>" });
lines.push({ id: "c1", kind: "line", x1: 0, y1: 0, x2: 0, y2: 0, from: { id: "r1", site: 1 }, to: { id: "r2", site: 0 }, route: "elbow", tail: "triangle" });
lines.push({ id: "c2", kind: "line", x1: 0, y1: 0, x2: 0, y2: 0, from: { id: "r1", site: 2 }, to: { id: "r2", site: 3 }, route: "curve", tail: "stealth", stroke: "#808080" });
slides.push({ type: "blank", title: "線・矢印・コネクタ", elements: lines });
// Text
slides.push({ type: "blank", title: "テキストと書式", elements: [
  { id: "x1", kind: "text", x: 80, y: 180, w: 820, h: 300, fs: 40, text: "<p><b>太字</b>・<i>斜体</i>・<u>下線</u>・<s>取り消し線</s>・x<sup>2</sup>・H<sub>2</sub>O</p><p><span style=\"color: #1f3864; font-size: 56px\">大きな濃紺</span>と<span style=\"background-color: #dce4f2\">マーカー</span></p><p style=\"text-align: right\">右揃えの段落</p>" },
  { id: "x2", kind: "text", x: 980, y: 180, w: 800, h: 300, fs: 34, text: "<ul><li>箇条書き1<ul><li>下の階層</li></ul></li><li>箇条書き2</li></ul><ol><li>番号付き</li><li>番号付き</li></ol>" },
  { id: "x3", kind: "shape", shape: "wedgeRoundRectCallout", x: 120, y: 560, w: 520, h: 230, text: "<p>吹き出しの中の文字は上下中央</p>", fs: 34, adj: [-0.3, 0.8] },
  { id: "x4", kind: "text", x: 760, y: 520, w: 160, h: 420, vertical: true, fs: 40, text: "<p>縦書きのテキスト</p>", fill: "#f5f0ea" },
  { id: "x5", kind: "shape", shape: "rightArrow", x: 1000, y: 600, w: 420, h: 200, text: "<p>回転・反転</p>", rot: -12, flipH: false },
  { id: "x6", kind: "shape", shape: "homePlate", x: 1470, y: 600, w: 340, h: 160, flipH: true, text: "<p>左右反転</p>", fill: "#d6c9b8" },
  { id: "x7", kind: "shape", shape: "rect", x: 1500, y: 820, w: 300, h: 90, fs: 40, autofit: "shrink", text: "<p>はみ出す場合だけ縮小する長い文字の例です</p>" },
] });
// Pictures, icons, colours of their own
slides.push({ type: "blank", title: "図・アイコン・ブランド検査", elements: [
  { id: "p1", kind: "image", src: "asset:ai", x: 80, y: 180, w: 480, h: 300 },
  { id: "p2", kind: "image", src: "asset:dataInsight", x: 620, y: 180, w: 300, h: 300, mask: "ellipse", stroke: "#1f3864", strokeW: 6 },
  { id: "p3", kind: "image", src: "asset:storeOperations", x: 980, y: 180, w: 420, h: 300, crop: { l: 0.2, t: 0.1, r: 0.1, b: 0.2 }, gray: true },
  { id: "p4", kind: "image", src: "asset:executiveDecision", x: 1460, y: 180, w: 300, h: 300, mask: "star5" },
  ...Object.keys(E.icons).slice(0, 8).map((icon, i) => ({ id: `i${i}`, kind: "icon", icon, x: 100 + i * 140, y: 560, w: 100, h: 100, color: ["#1f3864", "#808080", "#1a1a1a"][i % 3] })),
  { id: "bad1", kind: "shape", shape: "rect", x: 1250, y: 560, w: 300, h: 140, fill: "#ff8800", text: "<p>ブランド外の色</p>" },
  { id: "bad2", kind: "shape", shape: "rect", x: 1250, y: 740, w: 300, h: 140, fill: "#1f3864", text: "<p>濃紺に文字</p>", color: "#1a1a1a" },
  { id: "bad3", kind: "shape", shape: "rect", x: 1700, y: 30, w: 200, h: 80, fill: "#dce4f2", stroke: "#1f3864" },
] });
slides.push({ type: "closing", message: "以上" });
await writeFile(`${root}qa/objects.json`, `${JSON.stringify({ deckTitle: "オブジェクトの見本", audience: "社内", purpose: "QA", slideData: slides }, null, 1)}\n`);
console.log(`qa/objects.json: ${slides.length} slides`);
