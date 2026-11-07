// PowerPoint's formatting odds and ends: トリミング ▾ (縦横比・塗りつぶし・枠に合わせる), 箇条書き・段落番号の種類 (kept on the
// list, nothing else let through), 文字列の方向 (90 / 270 度回転), ビデオの表紙画像, and 最終版にする (the document's 状態).
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { cropFill, cropFit, cropToAspect } from "../public/editor/crop.mjs";
import { FINAL_STATUS, isFinal, slideParagraphs, wordCount } from "../public/editor/fileinfo.mjs";
import * as ops from "../public/editor/ops.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  return { E: window.SlideEngine, window };
}
const near = (a, b, eps = 1e-6) => Math.abs(a - b) < eps;
const plainJson = (value) => JSON.parse(JSON.stringify(value));

test("トリミング: 縦横比 keeps the picture's scale, 塗りつぶし keeps the frame, 枠に合わせる shows it all", () => {
  // A 4:3 picture shown 400 × 300 at (100, 100).
  const o = { x: 100, y: 100, w: 400, h: 300 };
  const sq = cropToAspect(o, 1, 4 / 3);
  assert.ok(near(sq.w, 300) && near(sq.h, 300), "1:1 — the largest middle square");
  assert.ok(near(sq.x + sq.w / 2, 300) && near(sq.y + sq.h / 2, 250), "the frame's middle kept");
  assert.ok(near(sq.crop.l, 0.125) && near(sq.crop.r, 0.125) && sq.crop.t === 0 && sq.crop.b === 0, `crop: ${JSON.stringify(sq.crop)}`);
  const wide = cropToAspect(o, 16 / 9, 4 / 3);
  assert.ok(near(wide.w, 400) && near(wide.h, 225) && near(wide.crop.t, 0.125), "16:9 — the top and bottom go");
  // An already cropped picture: the full picture is worked out from its crop.
  const again = cropToAspect({ ...sq, ...{ x: sq.x, y: sq.y } }, 4 / 3, 4 / 3);
  assert.ok(near(again.w, 400) && near(again.h, 300) && again.crop === undefined, "back to 4:3 — nothing cropped");
  // 塗りつぶし: a square frame filled by the 4:3 picture.
  const fill = cropFill({ x: 0, y: 0, w: 300, h: 300 }, 4 / 3);
  assert.ok(near(fill.crop.l, 0.125) && near(fill.crop.r, 0.125) && fill.crop.t === 0, "the sides cropped");
  assert.equal(fill.w, undefined, "the frame stays");
  // 枠に合わせる: the whole picture inside the square frame.
  const fit = cropFit({ x: 0, y: 0, w: 300, h: 300 }, 4 / 3);
  assert.ok(near(fit.w, 300) && near(fit.h, 225) && near(fit.y, 37.5) && fit.crop === undefined);
});

test("箇条書き・段落番号の種類: the marker kept on the list, only known ones; the editor reads it back", async () => {
  const { E, window } = await loadEngine();
  assert.equal(E.sanitizeRich('<ul data-style="diamond"><li>あ</li></ul>'), '<ul data-style="diamond"><li>あ</li></ul>');
  assert.equal(E.sanitizeRich('<ol data-style="circled"><li>あ</li></ol>'), '<ol data-style="circled"><li>あ</li></ol>');
  assert.equal(E.sanitizeRich('<ul data-style="circled"><li>あ</li></ul>'), "<ul><li>あ</li></ul>", "a number style is not a bullet");
  assert.equal(E.sanitizeRich('<ul data-style="x&quot; onclick=&quot;"><li>あ</li></ul>'), "<ul><li>あ</li></ul>");
  assert.equal(E.sanitizeRich('<ul data-style="disc"><li>あ</li></ul>'), "<ul><li>あ</li></ul>", "the plain one needs no mark");
  globalThis.document = window.document;
  try {
    const listed = ops.setList(E, "<p>一</p><p>二</p>", "number", "paren");
    assert.equal(listed, '<ol data-style="paren"><li>一</li><li>二</li></ol>');
    assert.equal(ops.listOf(listed), "number");
    assert.equal(ops.listStyleOf(listed), "paren");
    assert.equal(ops.listStyleOf("<ul><li>a</li></ul>"), "disc");
    assert.equal(ops.listStyleOf("<p>a</p>"), null);
    assert.equal(ops.setList(E, listed, null), "<p>一</p><p>二</p>", "なし: paragraphs again");
  } finally { delete globalThis.document; }
});

test("文字列の方向 and ビデオの表紙画像 are kept (and drawn); nothing else gets through", async () => {
  const { E } = await loadEngine();
  const t = E.normalizeObject({ id: "t", kind: "text", x: 0, y: 0, w: 200, h: 400, text: "<p>縦</p>", textRot: 270 });
  assert.equal(t.textRot, 270);
  assert.equal(E.normalizeObject({ id: "u", kind: "text", x: 0, y: 0, w: 200, h: 400, text: "<p>縦</p>", textRot: 45 }).textRot, undefined);
  const slide = { type: "blank", elements: [t] };
  const el = E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  assert.ok(el.querySelector(".hs-obj-text.is-rot270"), "drawn turned");
  const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
  const v = E.normalizeObject({ id: "v", kind: "video", x: 0, y: 0, w: 640, h: 360, src: "https://example.com/a.mp4", poster: png });
  assert.equal(v.poster, png);
  assert.equal(E.normalizeObject({ id: "w", kind: "video", x: 0, y: 0, w: 640, h: 360, src: "https://example.com/a.mp4", poster: "javascript:alert(1)" }).poster, undefined);
});

test("最終版にする is the document's 状態 (最終版)", () => {
  assert.equal(FINAL_STATUS, "最終版");
  assert.equal(isFinal({ info: { status: "最終版" } }), true);
  assert.equal(isFinal({ info: { status: "下書き" } }), false);
  assert.equal(isFinal({}), false);
  assert.equal(isFinal(null), false);
});

const plain = (v) => JSON.parse(JSON.stringify(v));
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

test("背景の書式設定: any colour, a picture with its transparency, drawn under everything", async () => {
  const { E } = await loadEngine();
  assert.deepEqual(plain(E.normalizeBackground({ color: "#DCE4F2" })), { color: "#dce4f2" });
  assert.deepEqual(plain(E.normalizeBackground({ color: "#1f3864" })), { color: "#1f3864" }, "a dark background too");
  assert.equal(E.normalizeBackground({ color: "nonsense" }), null, "not a colour");
  assert.deepEqual(plain(E.normalizeBackground({ image: PNG, transparency: 0.333, tile: true, extra: 1 })), { image: PNG, transparency: 0.33, tile: true });
  assert.equal(E.normalizeBackground({ image: "javascript:alert(1)" }), null);
  assert.equal(E.normalizeBackground({ color: "#f5f0ea", transparency: 0.5 }).transparency, undefined, "transparency is the picture's");
  const slide = { type: "content", title: "背景", body: "本文", background: { color: "#f5f0ea" } };
  const el = E.render(slide, { mode: "present", index: 0, deck: { slides: [slide], theme: "clarity" } });
  const bg = el.querySelector(".hs-bg");
  assert.ok(bg && el.firstElementChild === bg, "first in the slide: under the layout, its objects and the master");
  assert.match(bg.getAttribute("style") || "", /f5f0ea|245, 240, 234/);
  const pic = { type: "blank", background: { image: PNG, transparency: 0.5, tile: true } };
  const img = E.render(pic, { mode: "present", index: 0, deck: { slides: [pic], theme: "clarity" } }).querySelector(".hs-bg-img");
  assert.ok(img?.classList.contains("is-tiled"), "tiled");
  assert.match(img.getAttribute("style"), /url\("data:image\/png/);
  assert.match(img.getAttribute("style"), /opacity: ?0\.5/);
  assert.equal(E.render({ type: "blank" }, { mode: "present", index: 0, deck: { slides: [{ type: "blank" }], theme: "clarity" } }).querySelector(".hs-bg"), null, "none set, none drawn");
});

test("均等割り付け on a box, a cell and a paragraph; a cell's text can be vertical", async () => {
  const { E } = await loadEngine();
  const t = E.normalizeObject({ id: "t", kind: "text", x: 0, y: 0, w: 400, h: 100, text: "<p>あいう</p>", align: "distributed" });
  assert.equal(t.align, "distributed");
  const slide = { type: "blank", elements: [t] };
  const tx = E.render(slide, { mode: "present", index: 0, deck: { slides: [slide], theme: "clarity" } }).querySelector(".hs-obj-tx");
  assert.match(tx.getAttribute("style"), /text-align: ?justify/);
  assert.match(tx.getAttribute("style"), /text-align-last: ?justify/);
  assert.equal(E.sanitizeRich('<p style="text-align: justify; text-align-last: justify">あ</p>'), '<p style="text-align: justify; text-align-last: justify">あ</p>');
  assert.equal(E.sanitizeRich('<p style="text-align: justify">あ</p>'), '<p style="text-align: justify">あ</p>');
  assert.equal(E.sanitizeRich('<p style="text-align-last: justify">あ</p>'), "<p>あ</p>", "only with justified lines");
  const table = E.normalizeObject({ id: "g", kind: "table", x: 0, y: 0, w: 600, h: 200, cells: [[{ text: "<p>見出し</p>", align: "distributed", vertical: true, pad: [0, 0, 0, 0] }, { text: "<p>b</p>", vertical: "yes" }]] });
  assert.equal(table.cells[0][0].align, "distributed");
  assert.equal(table.cells[0][0].vertical, true);
  assert.equal(table.cells[0][1].vertical, undefined, "only true");
  assert.deepEqual(plain(table.cells[0][0].pad), [0, 0, 0, 0], "セルの余白 なし");
  const tslide = { type: "blank", elements: [table] };
  const cell = E.render(tslide, { mode: "present", index: 0, deck: { slides: [tslide], theme: "clarity" } }).querySelector(".hs-cell-tx");
  assert.ok(cell.classList.contains("is-vertical"));
  assert.match(cell.getAttribute("style"), /text-align-last: ?justify/);
});

test("罫線: outside, inside, a side, none; the cell across an edge gets the same line; a merge only on its own edges", () => {
  const grid = (n, m) => ({ kind: "table", cells: Array.from({ length: n }, () => Array.from({ length: m }, () => ({}))) });
  const pen = { c: "#1f3864", w: 3 };
  const out = ops.tableBorders(grid(3, 3), 0, 0, 1, 1, "outside", pen);
  assert.deepEqual(out.cells[0][0].bt, pen);
  assert.deepEqual(out.cells[0][0].bl, pen);
  assert.equal(out.cells[0][0].br, undefined, "no inside line");
  assert.deepEqual(out.cells[1][1].bb, pen);
  assert.deepEqual(out.cells[2][1].bt, pen, "the cell under the range takes the same edge");
  assert.deepEqual(out.cells[0][2].bl, pen, "and the one to its right");
  assert.notEqual(out.cells[0][0].bt, out.cells[0][1].bt, "each cell its own copy");
  const inside = ops.tableBorders(grid(2, 2), 0, 0, 1, 1, "inside", pen);
  assert.deepEqual(inside.cells[0][0].br, pen);
  assert.deepEqual(inside.cells[0][0].bb, pen);
  assert.equal(inside.cells[0][0].bt, undefined);
  const bottom = ops.tableBorders(grid(2, 2), 0, 0, 0, 1, "bottom", pen);
  assert.deepEqual([bottom.cells[0][0].bb, bottom.cells[0][1].bb, bottom.cells[1][0].bt], [pen, pen, pen]);
  assert.equal(bottom.cells[0][0].bt, undefined);
  const none = ops.tableBorders(out, 0, 0, 1, 1, "none", pen);
  assert.equal(none.cells[0][0].bt, "none");
  assert.equal(none.cells[2][1].bt, "none", "the shared edge goes too");
  // A cell merged across two rows: an inside line across the merge is not drawn through it.
  const merged = grid(2, 2);
  merged.cells[0][0] = { rs: 2 };
  merged.cells[1][0] = { merged: true };
  const m = ops.tableBorders(merged, 0, 0, 1, 1, "insideH", pen);
  assert.equal(m.cells[0][0].bb, undefined, "no line through the merged cell");
  assert.deepEqual(m.cells[0][1].bb, pen, "but between the cells beside it");
  assert.equal(m.cells[0][0].rs, 2, "still merged");
});

/** The first list's tag and attributes, sorted (linkedom writes attributes in another order than browsers). */
const listTag = (html) => { const m = String(html).match(/^<(ul|ol)\b([^>]*)>/); return m ? `${m[1]} ${[...m[2].matchAll(/([\w-]+)="([^"]*)"/g)].map((a) => `${a[1]}=${a[2]}`).sort().join(" ")}`.trim() : null; };

test("箇条書きと段落番号: the first number, the marks' colour and size kept on the list (and through a change of style)", async () => {
  const { E, window } = await loadEngine();
  assert.equal(listTag(E.sanitizeRich('<ol start="3" data-mark="accent" data-msize="125"><li>あ</li></ol>')), "ol data-mark=accent data-msize=125 start=3");
  assert.equal(E.sanitizeRich('<ul start="3" data-mark="red" data-msize="300"><li>あ</li></ul>'), "<ul><li>あ</li></ul>", "no start on bullets; only the colours and the sizes offered");
  assert.equal(E.sanitizeRich('<ol start="1"><li>あ</li></ol>'), "<ol><li>あ</li></ol>", "1 is the start anyway");
  globalThis.document = window.document;
  try {
    const list = ops.setListProps(E, "<ol><li>一</li><li>二</li></ol>", { start: 5, mark: "muted", msize: "150" });
    assert.equal(listTag(list), "ol data-mark=muted data-msize=150 start=5");
    assert.match(list, /<li>一<\/li><li>二<\/li><\/ol>$/);
    assert.deepEqual(ops.listPropsOf(list), { type: "number", start: 5, mark: "muted", msize: "150" });
    assert.equal(ops.listStyleOf('<ol data-style="circled" start="5"><li>一</li></ol>'), "circled", "the style read with other marks on the list");
    const restyled = ops.setList(E, list, "number", "circled");
    assert.equal(listTag(restyled), "ol data-mark=muted data-msize=150 data-style=circled start=5", "kept through a change of style");
    const bullets = ops.setList(E, list, "bullet", "diamond");
    assert.equal(listTag(bullets), "ul data-mark=muted data-msize=150 data-style=diamond", "bullets keep the colour and size, not the number");
    assert.equal(ops.setListProps(E, list, { start: 1, mark: "", msize: "100" }), "<ol><li>一</li><li>二</li></ol>", "all back to plain");
  } finally { delete globalThis.document; }
});

test("下線の種類・二重取り消し線: kept on the words and on the box, nothing else let through", async () => {
  const { E } = await loadEngine();
  assert.equal(E.sanitizeRich('<p><u data-line="wavy">波</u><s data-line="double">消</s></p>'), '<p><u data-line="wavy">波</u><s data-line="double">消</s></p>');
  assert.equal(E.sanitizeRich('<p><u data-line="zigzag">a</u><s data-line="wavy">b</s></p>'), "<p><u>a</u><s>b</s></p>", "only the kinds offered");
  const o = E.normalizeObject({ id: "t", kind: "text", x: 0, y: 0, w: 400, h: 100, text: "<p>見出し</p>", underline: true, uline: "double", strike: true, sline: "double" });
  assert.equal(o.uline, "double");
  assert.equal(o.sline, "double");
  assert.equal(E.normalizeObject({ id: "u", kind: "text", x: 0, y: 0, w: 400, h: 100, text: "<p>a</p>", uline: "zigzag" }).uline, undefined);
  const slide = { type: "blank", elements: [o] };
  const tx = E.render(slide, { mode: "present", index: 0, deck: { slides: [slide], theme: "clarity" } }).querySelector(".hs-obj-tx");
  assert.match(tx.getAttribute("style"), /text-decoration-style: ?double/);
  const thick = E.normalizeObject({ id: "v", kind: "text", x: 0, y: 0, w: 400, h: 100, text: "<p>a</p>", underline: true, uline: "thick" });
  const s2 = { type: "blank", elements: [thick] };
  assert.match(E.render(s2, { mode: "present", index: 0, deck: { slides: [s2], theme: "clarity" } }).querySelector(".hs-obj-tx").getAttribute("style"), /text-decoration-thickness/);
});

test("セルのサイズ: rows (or columns) set in pixels, the table growing by the difference, the others kept", () => {
  const t = { kind: "table", x: 0, y: 0, w: 600, h: 300, rows: [0.5, 0.25, 0.25], cols: [0.5, 0.5], cells: [[{}, {}], [{}, {}], [{}, {}]] };
  const taller = ops.tableSetSize(t, "rows", 1, 1, 150);
  assert.equal(taller.h, 375);
  const heights = taller.rows.map((f) => Math.round(f * taller.h));
  assert.deepEqual(heights, [150, 150, 75]);
  const narrow = ops.tableSetSize(t, "cols", 0, 1, 200);
  assert.equal(narrow.w, 400);
  assert.deepEqual(narrow.cols, [0.5, 0.5]);
  assert.equal(ops.tableSetSize(t, "rows", 0, 0, 2).rows.length, 3, "never below the smallest size");
});

test("図で塗りつぶし: a picture fills the shape (cut to its outline), stretched or tiled; only real pictures", async () => {
  const { E } = await loadEngine();
  const o = E.normalizeObject({ id: "s1", kind: "shape", shape: "ellipse", x: 100, y: 100, w: 400, h: 200, fill: "#dce4f2", fillImg: "data:image/png;base64,AAAA", fillTile: true });
  assert.equal(o.fillImg, "data:image/png;base64,AAAA");
  assert.equal(o.fillTile, true);
  assert.equal(E.normalizeObject({ id: "s2", kind: "shape", shape: "rect", fillImg: "javascript:alert(1)", fillTile: true }).fillImg, undefined, "only pictures");
  assert.equal(E.normalizeObject({ id: "s3", kind: "shape", shape: "rect", fillImg: "data:text/html,x" }).fillImg, undefined);
  const draw = (obj, ctx = {}) => { const slide = { type: "blank", elements: [obj] }; return E.render(slide, { mode: "present", index: 0, deck: { slides: [slide], theme: "clarity" }, ...ctx }); };
  const tiled = draw(o);
  const pattern = tiled.querySelector(".hs-obj pattern");
  assert.ok(pattern, "a pattern of the picture");
  assert.equal(pattern.getAttribute("width"), "160", "tiles");
  assert.equal(pattern.querySelector("image").getAttribute("href"), "data:image/png;base64,AAAA");
  assert.equal(pattern.querySelector("image").getAttribute("preserveAspectRatio"), "xMidYMid slice");
  assert.equal(tiled.querySelector(".hs-obj path").getAttribute("fill"), `url(#${pattern.getAttribute("id")})`, "the outline is filled with it");
  const stretched = draw({ ...o, fillTile: undefined });
  assert.equal(stretched.querySelector(".hs-obj pattern").getAttribute("width"), "400", "one picture over the whole shape");
  // A picture kept in this browser is drawn once its address is known (and left out until then).
  const stored = E.normalizeObject({ ...o, fillImg: "idb:abc" });
  assert.equal(draw(stored).querySelector(".hs-obj pattern"), null);
  assert.equal(draw(stored, { mediaUrls: { "idb:abc": "blob:x" } }).querySelector(".hs-obj pattern image").getAttribute("href"), "blob:x");
});

test("段落 → インデント: before the text, a first line indented or hanging; kept by the sanitizer, read back", async () => {
  const { E, window } = await loadEngine();
  globalThis.document = window.document;
  try {
    const hanging = ops.setIndent(E, '<p>一つ目</p><p style="text-align: center">二つ目</p>', { left: 56.69, first: -28.35 });
    assert.match(hanging, /<p style="[^"]*margin-left: ?56\.69px[^"]*text-indent: ?-28\.35px[^"]*">一つ目<\/p>/);
    assert.match(hanging, /text-align: ?center[^"]*margin-left: ?56\.69px/, "the paragraph keeps its alignment");
    assert.deepEqual({ ...ops.indentOf(hanging) }, { left: 56.69, first: -28.35 });
    assert.equal(ops.setIndent(E, hanging, { left: 0, first: 0 }), '<p>一つ目</p><p style="text-align: center">二つ目</p>', "0 takes them away");
    assert.match(ops.setIndent(E, "<ul><li>a</li></ul>", { left: 20, first: 10 }), /<li style="margin-left: ?20px; text-indent: ?10px">a<\/li>/);
    assert.equal(E.sanitizeRich('<p style="margin-left: 99999px; text-indent: abc">x</p>'), "<p>x</p>", "only lengths in range");
    // フォント → すべて大文字・小型英大文字 on the box.
    const o = E.normalizeObject({ id: "t", kind: "text", x: 0, y: 0, w: 400, h: 100, text: "<p>Seven</p>", caps: "small" });
    assert.equal(o.caps, "small");
    assert.equal(E.normalizeObject({ id: "t2", kind: "text", text: "<p>x</p>", caps: "huge" }).caps, undefined);
    const slide = { type: "blank", elements: [o] };
    const el = E.render(slide, { mode: "edit", index: 0, deck: { slides: [slide], theme: "clarity" } });
    assert.match(el.querySelector(".hs-obj-tx").getAttribute("style"), /font-variant: ?small-caps/);
  } finally { delete globalThis.document; }
});

test("図のレイアウト: pictures set out with a caption each, filling their frames, grouped", () => {
  const pics = [
    { id: "p1", kind: "image", x: 100, y: 300, w: 300, h: 200, src: "asset:office", alt: "売場", crop: { l: 0.1, t: 0, r: 0, b: 0 } },
    { id: "p2", kind: "image", x: 500, y: 320, w: 300, h: 200, src: "asset:office", fileName: "倉庫.jpg" },
    { id: "p3", kind: "image", x: 900, y: 340, w: 300, h: 200, src: "asset:office" },
  ];
  const other = { id: "t1", kind: "text", x: 0, y: 0, w: 100, h: 50, text: "<p>見出し</p>" };
  const { list, ids } = ops.pictureLayout([other, ...pics], ["p1", "p2", "p3", "t1"], "row");
  assert.equal(ids.length, 6, "three pictures and three captions");
  const placed = list.filter((o) => ids.includes(o.id));
  const group = placed[0].group;
  assert.ok(group && placed.every((o) => o.group === group), "one group");
  const [a, b, c] = ["p1", "p2", "p3"].map((id) => list.find((o) => o.id === id));
  assert.ok(Math.abs(a.y - b.y) < 0.01 && Math.abs(b.y - c.y) < 0.01 && a.x < b.x && b.x < c.x, "in a row");
  assert.ok(Math.abs(a.w - b.w) < 0.01 && Math.abs(a.h - c.h) < 0.01, "the same size");
  assert.equal(a.fit, "cover");
  assert.equal(a.crop, undefined, "the frame is filled instead of cropped");
  const caps = list.filter((o) => o.kind === "text" && ids.includes(o.id));
  assert.deepEqual(caps.map((o) => o.text), ["<p>売場</p>", "<p>倉庫</p>", "<p>説明を入力</p>"], "alternative text, then the file's name");
  assert.ok(caps.every((cap, i) => cap.y > [a, b, c][i].y + [a, b, c][i].h - 1), "each caption under its picture");
  assert.equal(list.find((o) => o.id === "t1").group, undefined, "other objects are left alone");
  const grid = ops.pictureLayout(pics, ["p1", "p2", "p3"], "grid").list;
  assert.equal(grid.find((o) => o.id === "p3").x, grid.find((o) => o.id === "p1").x, "2 columns: the third starts a new row");
  const side = ops.pictureLayout(pics, ["p1", "p2"], "side");
  const cap = side.list.find((o) => o.id === side.ids[2]);
  assert.ok(cap.x > side.list.find((o) => o.id === "p1").x && cap.align === "left", "the words to the right of the picture");
  assert.equal(ops.pictureLayout([other], ["t1"], "row").ids.length, 0, "only pictures");
});

test("既定の図形に設定: only the look of a shape, text box or line is kept (never its words, place or size)", () => {
  const shape = { id: "s", kind: "shape", shape: "rect", x: 10, y: 20, w: 300, h: 100, text: "<p>見出し</p>", fill: "#e2efda", stroke: "#1f3864", strokeW: 3, fs: 32, color: "#1f3864", bold: true, rot: 15 };
  const style = ops.defaultStyleOf("shape", shape);
  assert.deepEqual(style, { fill: "#e2efda", stroke: "#1f3864", strokeW: 3, fs: 32, color: "#1f3864", bold: true });
  assert.deepEqual(ops.defaultStyleOf("line", { kind: "line", x1: 0, y1: 0, x2: 9, y2: 9, stroke: "#808080", strokeW: 6, dash: "dash", tail: "triangle", fill: "#fff" }), { stroke: "#808080", strokeW: 6, dash: "dash", tail: "triangle" });
  assert.equal(ops.defaultStyleOf("table", shape), null);
  assert.equal(ops.defaultStyleOf("shape", { text: "x" }), null);
  const kept = ops.objectDefaultsOf({ shape, text: { wrap: false, fs: 28, x: 5 }, line: "x", chart: { fill: "#fff" } });
  assert.deepEqual(Object.keys(kept), ["shape", "text"]);
  assert.deepEqual(kept.text, { fs: 28, wrap: false });
  assert.equal(ops.objectDefaultsOf({ shape: { x: 1 } }), null);
  assert.equal(ops.objectDefaultsOf(null), null);
  assert.equal(ops.objectDefaultsOf([shape]), null);
  // The check passes each look (the studio normalizes it and moves colours to the deck's palette).
  const checked = ops.objectDefaultsOf({ shape: { fill: "#ff0000", fs: 30 } }, (kind, s) => (kind === "shape" ? { ...s, fill: "#dce4f2" } : s));
  assert.deepEqual(checked, { shape: { fill: "#dce4f2", fs: 30 } });
  assert.equal(ops.objectDefaultsOf({ shape: { fill: "#ff0000" } }, () => null), null, "a look the check refuses is dropped");
});

test("線の書式設定: 線端・結合点・複合線 are kept only when known and drawn on shapes, lines and picture outlines", async () => {
  const { E } = await loadEngine();
  const [a, b, c] = E.normalizeObjects([
    { id: "a", kind: "shape", shape: "rect", x: 0, y: 0, w: 200, h: 100, stroke: "#1f3864", strokeW: 8, cap: "round", join: "bevel", cmpd: "dbl" },
    { id: "b", kind: "shape", shape: "rect", x: 0, y: 0, w: 200, h: 100, stroke: "#1f3864", strokeW: 8, cap: "flat", join: "weird", cmpd: "quad" },
    { id: "c", kind: "line", x1: 0, y1: 0, x2: 300, y2: 0, stroke: "#1f3864", strokeW: 12, cap: "square", cmpd: "tri" },
  ]);
  assert.deepEqual([a.cap, a.join, a.cmpd], ["round", "bevel", "dbl"]);
  assert.deepEqual([b.cap, b.join, b.cmpd], [undefined, undefined, undefined], "flat is the usual end; unknown names are dropped");
  assert.deepEqual([c.cap, c.cmpd], ["square", "tri"]);
  const slide = { type: "blank", title: "検証", elements: [a, b, c] };
  const el = E.render(slide, { mode: "present", index: 1, deck: { slides: [{ type: "title", title: "表紙" }, slide], theme: "clarity" } });
  const stroked = (id) => [...el.querySelectorAll(`.hs-obj[data-el="${id}"] svg path`)];
  const outline = stroked("a").find((p) => p.getAttribute("mask"));
  assert.ok(outline, "the outline of a compound line is cut by a mask");
  assert.equal(outline.getAttribute("stroke-linecap"), "round");
  assert.equal(outline.getAttribute("stroke-linejoin"), "bevel");
  assert.ok(stroked("a").some((p) => p.getAttribute("fill") !== "none" && p.getAttribute("stroke") === "none"), "the fill is drawn alone, not cut");
  const mask = el.querySelector(`.hs-obj[data-el="a"] mask`);
  assert.equal(mask.querySelectorAll("path").length, 1, "a double line: one band cut");
  const lineMask = el.querySelector(`.hs-obj[data-el="c"] mask`);
  assert.equal(lineMask.querySelectorAll("path").length, 2, "a triple line: two bands, the middle one drawn back");
  assert.ok(stroked("c").some((p) => p.getAttribute("stroke-linecap") === "square" && p.getAttribute("mask")));
  assert.equal(el.querySelectorAll(`.hs-obj[data-el="b"] mask`).length, 0, "a plain outline needs no mask");
  assert.ok(stroked("b").every((p) => p.getAttribute("stroke-linejoin") === "miter"), "a shape's corners are sharp unless told otherwise");
});

test("表の貼り付け: Excel's tab-separated cells and a web page's table become a table (merged spans kept)", async () => {
  const { E, window } = await loadEngine();
  assert.deepEqual(ops.parseTsv('a\tb\n"x\ny"\t"q""r"\n'), [["a", "b"], ["x\ny", 'q"r']], "Excel's quoting");
  assert.equal(ops.clipboardGrid({ text: "ただの文章です。\nもう1行" }), null, "plain words are not cells");
  assert.equal(ops.clipboardGrid({ text: "ひとつのセル" }), null);
  const sheet = ops.clipboardGrid({ text: "品目\t売上\t前年比\nお茶\t1,200\t98%\n水\t800\t105%\n" });
  assert.equal(sheet.rows.length, 3, "the trailing newline makes no empty row");
  assert.deepEqual(plainJson(sheet.rows[1]), [{ text: "<p>お茶</p>" }, { text: "<p>1,200</p>", align: "right" }, { text: "<p>98%</p>", align: "right" }], "numbers align right");
  const table = ops.tableFromGrid(sheet);
  assert.equal(table.kind, "table");
  assert.equal(table.cells.length, 3);
  assert.equal(table.cols.length, 3);
  assert.ok(table.x >= 0 && table.x + table.w <= 1920 && table.y + table.h <= 1000, "inside the slide");
  const kept = E.normalizeObject({ id: "t", ...table });
  assert.equal(kept.cells[1][1].text, "<p>1,200</p>", "the normalizer keeps it");
  // A web page's table: header cells, a span, line breaks.
  const parse = (source) => parseHTML(source).document;
  const html = '<table><tr><th colspan="2">計画</th></tr><tr><td>A<br>B</td><td style="text-align:center">2</td></tr><tr><td rowspan="2">合計</td><td>3</td></tr><tr><td>4</td></tr></table>';
  const web = ops.clipboardGrid({ text: "", html }, parse);
  assert.equal(web.rows.length, 4);
  assert.equal(web.rows[0][0].cs, 2, "a column span");
  assert.equal(web.rows[0][1].merged, true, "the covered cell");
  assert.equal(web.rows[1][0].text, "<p>A</p><p>B</p>", "a line break is a new paragraph");
  assert.equal(web.rows[1][1].align, "center");
  assert.equal(web.rows[2][0].rs, 2);
  assert.equal(web.rows[3][0].merged, true);
  assert.equal(ops.clipboardGrid({ text: "", html: "<table><tr><td>one</td></tr></table>" }, parse), null, "one cell is just words");
  assert.equal(ops.clipboardGrid({ text: "", html: "<p>no table</p>" }, parse), null);
  // Pasted into a table at a cell: the table grows as the cells need; a merged cell's covered position is skipped.
  const base = ops.makeTable(2, 2, { x: 100, y: 100, w: 400, h: 150 });
  const filled = ops.tableFill(base, 1, 1, { rows: [[{ text: "<p>a</p>" }, { text: "<p>b</p>" }], [{ text: "<p>c</p>" }, { text: "<p>d</p>", align: "right" }]] });
  assert.equal(filled.cells.length, 3, "a row added");
  assert.equal(filled.cells[0].length, 3, "a column added");
  assert.equal(filled.cells[1][1].text, "<p>a</p>");
  assert.equal(filled.cells[2][2].text, "<p>d</p>");
  assert.equal(filled.cells[2][2].align, "right");
  assert.equal(filled.cells[0][0].text, undefined, "cells before are left alone");
  assert.ok(filled.h > base.h && filled.w > base.w, "the table grew");
});

test("再グループ化 and the slide number's first number", async () => {
  const { E } = await loadEngine();
  const list = [{ id: "a", kind: "shape" }, { id: "b", kind: "shape" }, { id: "c", kind: "shape" }];
  const grouped = ops.group(list, ["a", "b"]);
  const apart = ops.ungroup(grouped, ["a"]);
  assert.ok(apart.every((o) => !o.group), "taken apart");
  const again = ops.group(apart, ["a", "b"]);
  assert.equal(again.find((o) => o.id === "a").group, again.find((o) => o.id === "b").group, "put together again");
  // スライド番号の開始番号: the master's page number and the number field count from it.
  const slides = [{ type: "title", title: "表紙" }, { type: "blank", title: "本文", elements: [{ id: "f", kind: "text", x: 100, y: 300, w: 400, h: 80, text: '<p><span data-field="slideno">1</span></p>' }] }, { type: "blank", title: "次" }];
  const page = (index, firstNumber) => {
    const deck = { slides, theme: "clarity", ...(firstNumber != null ? { firstNumber } : {}) };
    const el = E.render(slides[index], { deck, index, mode: "present" });
    return { master: el.querySelector(".hs-foot .hs-page")?.textContent || "", field: el.querySelector('span[data-field="slideno"]')?.textContent, label: el.getAttribute("aria-label") };
  };
  assert.equal(page(1, undefined).field, "2", "numbered from 1 as before");
  assert.equal(page(1, 0).field, "1", "the cover is 0, so this is 1");
  assert.match(page(1, 0).master, /^01 \/ /, "the foot's page number counts from it too");
  assert.equal(page(1, 5).field, "6");
  assert.match(page(1, 5).label, /^2枚目/, "the screen reader says where it really is");
  assert.equal(page(1, 99999).field, "2", "out of range: as before");
});

test("グラデーション塗りつぶし and 図のスタイル use the theme's colours; 文字カウント counts the words", async () => {
  const { E } = await loadEngine();
  const colors = E.themeColors(E.DEFAULT_THEME);
  for (const [key] of ops.GRADIENT_SETS) for (const [dir] of ops.GRADIENT_DIRECTIONS) {
    const g = ops.makeGradient(key, dir, colors);
    assert.ok(g.stops.every((stop) => E.themeColorRef(stop.color, E.DEFAULT_THEME)), `${key}: theme colours (they follow the theme)`);
    assert.deepEqual(ops.gradientChoice(g, colors), { set: key, direction: dir }, "which one it is");
  }
  assert.equal(ops.gradientChoice({ angle: 33, stops: [{ at: 0, color: "#ff0000" }, { at: 1, color: "#00ff00" }] }), null, "one brought over from PowerPoint is none of them");
  assert.deepEqual(ops.makeGradient("nope", "nope"), ops.makeGradient("accent1", "down"), "unknown names: the first");
  const shape = E.normalizeObject({ id: "g", kind: "shape", shape: "rect", x: 0, y: 0, w: 300, h: 200, fill: "#dce4f2", gradient: ops.makeGradient("accent1-soft", "right") });
  assert.equal(shape.gradient.stops.length, 2);
  const slide = { type: "blank", title: "検証", elements: [shape] };
  const el = E.render(slide, { mode: "present", index: 1, deck: { slides: [{ type: "title", title: "表紙" }, slide], theme: "clarity" } });
  const grad = el.querySelector(`.hs-obj[data-el="g"] linearGradient`);
  assert.ok(grad, "drawn as a gradient");
  assert.equal(grad.getAttribute("x1"), "0%", "running to the right");
  // The painter and the defaults keep a gradient; a plain fill takes it off.
  assert.ok(ops.DEFAULT_STYLE_KEYS.shape.includes("gradient") && ops.DEFAULT_STYLE_KEYS.text.includes("gradient"));
  assert.equal(ops.update([{ id: "g", gradient: ops.makeGradient("gray") }], ["g"], { fill: "#fff", gradient: undefined })[0].gradient, undefined);
  // 図のスタイル: a frame and an outline that normalize as they are, and a plain style puts the picture back to none.
  for (const [label, look] of ops.PICTURE_STYLES) {
    const o = E.normalizeObject({ id: "p", kind: "image", src: "asset:storeOperations", x: 0, y: 0, w: 400, h: 300, ...Object.fromEntries(Object.entries(look).filter(([, v]) => v !== undefined)) });
    if (look.mask) assert.equal(o.mask, look.mask, label);
    if (look.stroke) { assert.match(look.stroke, /^#[0-9a-f]{6}$/, label); assert.equal(o.stroke, look.stroke); }
    if (look.cmpd) assert.equal(o.cmpd, look.cmpd, label);
  }
  const framed = { id: "p", kind: "image", mask: "ellipse", stroke: "#1f3864", strokeW: 18, cmpd: "dbl", cap: "round", x: 0, y: 0, w: 10, h: 10 };
  const plain = ops.update([framed], ["p"], ops.PICTURE_STYLES[0][1])[0];
  assert.deepEqual(Object.keys(plain).sort(), ["h", "id", "kind", "w", "x", "y"], "the first style takes every frame and outline off");
  // 文字カウント: titles, lists, text boxes, table cells and (when asked) notes; hidden slides optional.
  const deck = { slides: [
    { type: "content", title: "売上の報告", points: ["A店は 120% です", { title: "B店" }], notes: "ここで説明" },
    { type: "blank", title: "表", elements: [{ id: "t", kind: "text", x: 0, y: 0, w: 100, h: 50, text: "<p>一行目</p><p>二行目 ok</p>" }, { id: "h", kind: "text", hidden: true, x: 0, y: 0, w: 1, h: 1, text: "<p>隠れた文字</p>" }] },
    { type: "blank", title: "非表示", hidden: true },
  ] };
  assert.deepEqual(slideParagraphs(deck.slides[0], E).slice(0, 2), ["売上の報告", "A店は 120% です"]);
  assert.ok(!slideParagraphs(deck.slides[1], E).includes("隠れた文字"), "a hidden object is not counted");
  const all = wordCount(deck, E);
  assert.equal(all.pages, 3);
  assert.equal(all.paragraphs, 7, "売上の報告・A店・B店・表・一行目・二行目 ok・非表示");
  assert.ok(all.chars > all.charsNoSpace, "the spaces are counted in one");
  assert.ok(all.words >= 4, "Latin and number words: A, 120, B, ok");
  assert.equal(wordCount(deck, E, { hidden: false }).pages, 2);
  assert.equal(wordCount(deck, E, { notes: true }).paragraphs, 8);
  assert.deepEqual(wordCount({ slides: [] }, E), { pages: 0, paragraphs: 0, chars: 0, charsNoSpace: 0, words: 0 });
});

test("図形の効果: 影・反射・光彩・ぼかし are kept in range and drawn on the object; a shadow is flagged for the brand", async () => {
  const { E } = await loadEngine();
  const slide = (o) => ({ type: "blank", elements: [E.normalizeObject({ id: "a", kind: "shape", shape: "rect", x: 100, y: 100, w: 400, h: 200, fill: "#dce4f2", stroke: "none", ...o })] });
  const draw = (o) => E.render(slide(o), { mode: "present", index: 1, deck: { slides: [slide(o)], theme: "clarity" } });
  const full = E.normalizeObject({ id: "a", kind: "shape", shape: "rect", x: 0, y: 0, w: 100, h: 100,
    shadow: { dx: 6, dy: 6, blur: 10, color: "#000000", opacity: 0.35 }, reflect: { size: 0.5, opacity: 0.4, gap: 4 }, glow: { r: 16, color: "#B7C3DA", opacity: 0.6 }, soft: 10 });
  assert.deepEqual(plainJson(full.reflect), { size: 0.5, opacity: 0.4, gap: 4 });
  assert.deepEqual(plainJson(full.glow), { r: 16, color: "#b7c3da", opacity: 0.6 });
  assert.equal(full.soft, 10);
  // Out of range or incomplete: dropped or tidied.
  const odd = E.normalizeObject({ id: "b", kind: "shape", shape: "rect", x: 0, y: 0, w: 10, h: 10, reflect: { size: 0 }, glow: { r: 5000, color: "nonsense" }, soft: 0.2 });
  assert.equal(odd.reflect, undefined, "a reflection of nothing");
  assert.equal(odd.glow, undefined, "no colour, no glow");
  assert.equal(odd.soft, undefined);
  assert.equal(E.normalizeObject({ id: "c", kind: "shape", shape: "rect", x: 0, y: 0, w: 10, h: 10, glow: { r: 5000, color: "#808080" } }).glow.r, 200);
  assert.equal(E.normalizeObject({ id: "d", kind: "table", x: 0, y: 0, w: 10, h: 10, soft: 10, rows: 1, cols: 1 })?.soft, undefined, "only shapes, text boxes and pictures take effects");

  const el = draw({ shadow: full.shadow, reflect: full.reflect, glow: full.glow, soft: full.soft });
  const obj = el.querySelector(".hs-obj");
  assert.equal(obj.getAttribute("data-shadow"), "1", "the brand check looks for it");
  const rot = el.querySelector(".hs-obj-rot");
  const filter = rot.style.filter || rot.getAttribute("style");
  assert.equal((filter.match(/drop-shadow/g) || []).length, 3, "the shadow, and the glow drawn twice");
  assert.ok(/-webkit-box-reflect:\s*below 4px linear-gradient\(to bottom, transparent 50%/.test(rot.getAttribute("style")), rot.getAttribute("style"));
  assert.ok(/mask-image:\s*linear-gradient\(to right, transparent, #000 10px/.test(el.querySelector("svg.hs-obj-geom").getAttribute("style")), "the soft edge fades the shape, not its words");
  const plain = draw({});
  assert.equal(plain.querySelector(".hs-obj").getAttribute("data-shadow"), null);
  assert.ok(!/drop-shadow|box-reflect/.test(plain.querySelector(".hs-obj-rot").getAttribute("style") || ""));
});

test("図形の効果: the presets are plain numbers (a direction, a polar shadow, a reflection) and read back", () => {
  const br = ops.shadowPreset("br");
  assert.deepEqual(plainJson(br), { dx: 8, dy: 8, blur: 10, color: "#000000", opacity: 0.35 });
  assert.equal(ops.shadowDirection(br), "br");
  assert.equal(ops.shadowDirection(ops.shadowPreset("t")), "t");
  assert.equal(ops.shadowDirection(null), "");
  assert.equal(ops.shadowPreset("nowhere"), null);
  const polar = ops.shadowPolar({ dx: 0, dy: 8, blur: 10, color: "#1f3864", opacity: 0.5 });
  assert.deepEqual(plainJson(polar), { distance: 4, angle: 90, blur: 5, opacity: 0.5, color: "#1f3864" });
  assert.deepEqual(plainJson(ops.shadowFromPolar(polar)), { dx: 0, dy: 8, blur: 10, color: "#1f3864", opacity: 0.5 });
  assert.deepEqual(plainJson(ops.shadowFromPolar({ distance: 5, angle: 0, blur: 0, opacity: 2, color: "" })), { dx: 10, dy: 0, blur: 0, color: "#000000", opacity: 1 }, "an opacity above 1 is 1");
  assert.deepEqual(plainJson(ops.reflectionPreset("medium")), { size: 0.5, opacity: 0.4, gap: 2 });
  assert.equal(ops.reflectionPreset("huge"), null);
  assert.equal(ops.effectsWords({ shadow: {}, soft: 4 }), "影・ぼかし");
  assert.equal(ops.effectsWords({}), "");
  assert.ok(ops.GLOW_COLORS.every(([c]) => /^#[0-9a-f]{6}$/.test(c)) && ops.SHADOW_COLORS.length === 3);
});

test("文字の効果: 文字の影・光彩 are kept in range and drawn on the letters (not the box); a text shadow is flagged for the brand", async () => {
  const { E } = await loadEngine();
  const text = (o, kind = "text") => E.normalizeObject({ id: "t", kind, x: 100, y: 100, w: 600, h: 120, text: "<p>文字</p>", fs: 40, ...(kind === "shape" ? { shape: "rect", fill: "#dce4f2", stroke: "none" } : {}), ...o });
  const full = text({ tshadow: { dx: 3, dy: 3, blur: 4, color: "#000000", opacity: 0.4 }, tglow: { r: 10, color: "#B7C3DA", opacity: 0.7 } });
  assert.deepEqual(plainJson(full.tshadow), { dx: 3, dy: 3, blur: 4, color: "#000000", opacity: 0.4 });
  assert.deepEqual(plainJson(full.tglow), { r: 10, color: "#b7c3da", opacity: 0.7 });
  // Out of range or without any effect: tidied or dropped.
  const odd = text({ tshadow: { dx: 0, dy: 0, blur: 0, color: "#000000" }, tglow: { r: 5000, color: "nonsense" } });
  assert.equal(odd.tshadow, undefined, "a shadow that moves nothing and blurs nothing is no shadow");
  assert.equal(odd.tglow, undefined, "no colour, no glow");
  assert.equal(text({ tshadow: { dx: 900, dy: -900, blur: 900, color: "#808080" } }).tshadow.dx, 100, "a shadow stays near its letters");
  assert.equal(text({ tglow: { r: 5000, color: "#808080" } }).tglow.r, 100);
  assert.equal(text({ tshadow: { dx: 3, dy: 3, color: "#000000", opacity: 0 } }).tshadow, undefined, "an invisible shadow is dropped");
  assert.equal(E.normalizeObject({ id: "i", kind: "icon", name: "check", x: 0, y: 0, w: 80, h: 80, tshadow: { dx: 3, dy: 3, blur: 4, color: "#000000" } })?.tshadow, undefined, "only text takes them");
  // Drawn on the words, with the halo twice; a box with a shadow on its letters is flagged, one with a glow is not.
  const draw = (o, kind) => { const slide = { type: "blank", elements: [text(o, kind)] }; return E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } }); };
  for (const kind of ["text", "shape"]) {
    const el = draw({ tshadow: full.tshadow, tglow: full.tglow }, kind);
    const style = el.querySelector(".hs-obj-tx").getAttribute("style");
    assert.ok(/text-shadow:\s*3px 3px 4px rgba\(0,0,0,0.4\), 0px 0px 10px rgba\(183,195,218,0.7\)|text-shadow:\s*3px 3px 4px rgba\(0,0,0,0.4\), 0 0 10px rgba\(183,195,218,0.7\)/.test(style), `${kind}: ${style}`);
    assert.equal((style.match(/rgba\(183,195,218/g) || []).length, 2, "the halo twice, wide then close");
    assert.equal(el.querySelector(".hs-obj").getAttribute("data-tshadow"), "1", "the brand check looks for it");
    assert.ok(!(el.querySelector(".hs-obj-rot").getAttribute("style") || "").includes("drop-shadow"), "the box itself has no effect");
  }
  const glowOnly = draw({ tglow: full.tglow });
  assert.equal(glowOnly.querySelector(".hs-obj").getAttribute("data-tshadow"), null, "a glow is not a shadow");
  assert.ok(/text-shadow/.test(glowOnly.querySelector(".hs-obj-tx").getAttribute("style")));
  const plain = draw({});
  assert.ok(!/text-shadow/.test(plain.querySelector(".hs-obj-tx").getAttribute("style") || ""));
  assert.equal(plain.querySelector(".hs-obj").getAttribute("data-tshadow"), null);
  // The choices on the ribbon.
  const br = ops.textShadowPreset("br");
  assert.deepEqual(plainJson(br), { dx: 3, dy: 3, blur: 4, color: "#000000", opacity: 0.4 });
  assert.equal(ops.shadowDirection(br), "br");
  assert.ok(ops.TEXT_GLOW_SIZES.every(([r]) => r >= 1 && r <= 100));
  assert.equal(ops.textEffectsWords({ tshadow: {}, tglow: {} }), "影・光彩");
  assert.equal(ops.textEffectsWords({}), "");
  // A look copied or kept as the default carries them too.
  for (const kind of ["shape", "text"]) assert.ok(ops.DEFAULT_STYLE_KEYS[kind].includes("tshadow") && ops.DEFAULT_STYLE_KEYS[kind].includes("tglow"), kind);
  assert.deepEqual(plainJson(ops.defaultStyleOf("text", full)).tshadow, plainJson(full.tshadow));
});

test("パターンの塗りつぶし: a kind and two colours are kept, drawn as a repeating SVG pattern over the shape; the theme's colours are checked", async () => {
  const { E } = await loadEngine();
  const shape = (o) => E.normalizeObject({ id: "a", kind: "shape", shape: "rect", x: 100, y: 100, w: 400, h: 200, fill: "#dce4f2", stroke: "none", ...o });
  assert.deepEqual(plainJson(shape({ pattern: { kind: "ltUpDiag", fg: "#B7C3DA", bg: "#FFFFFF" } }).pattern), { kind: "ltUpDiag", fg: "#b7c3da", bg: "#ffffff" });
  assert.equal(shape({ pattern: { kind: "nonsense", fg: "#b7c3da", bg: "#ffffff" } }).pattern, undefined, "a kind that is not there");
  assert.equal(shape({ pattern: { kind: "pct50", fg: "no", bg: "#ffffff" } }).pattern, undefined, "a colour that is not one");
  assert.equal(E.normalizeObject({ id: "i", kind: "image", src: "data:image/png;base64,AA==", x: 0, y: 0, w: 100, h: 100, pattern: { kind: "pct50", fg: "#808080", bg: "#ffffff" } })?.pattern, undefined, "shapes and text boxes only");
  // The kinds: a name and a rule over an 8 × 8 cell; the dots grow with the share; a line pattern is a line.
  const kinds = Object.keys(E.PATTERN_KINDS);
  assert.ok(kinds.length >= 34, `kinds: ${kinds.length}`);
  const cells = (kind) => { let n = 0; for (let y = 0; y < 8; y += 1) for (let x = 0; x < 8; x += 1) if (E.PATTERN_KINDS[kind][1](x, y)) n += 1; return n; };
  assert.ok(cells("pct10") < cells("pct25") && cells("pct25") < cells("pct50") && cells("pct50") < cells("pct90"), "denser means more cells");
  assert.equal(cells("pct50"), 32, "half of the 64 cells");
  assert.equal(cells("horz"), 16, "a line in every 4 rows");
  assert.equal(cells("lgGrid"), 15, "the top row and the left column");
  assert.ok(kinds.every((k) => typeof E.PATTERN_KINDS[k][0] === "string" && cells(k) > 0 && cells(k) < 64), "every kind shows both colours");
  // Drawn: a <pattern> in the shape's svg, the path filled with it; a gradient or a picture goes under the pattern or over it as before.
  const draw = (o) => { const slide = { type: "blank", elements: [shape(o)] }; return E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } }); };
  const el = draw({ fill: "none", pattern: { kind: "dnDiag", fg: "#1f3864", bg: "#dce4f2" } });
  const pattern = el.querySelector(".hs-obj-geom pattern");
  assert.ok(pattern, "a pattern is made");
  assert.equal(pattern.querySelector("rect").getAttribute("fill"), "#dce4f2", "the background colour");
  assert.equal(pattern.querySelector("path").getAttribute("fill"), "#1f3864", "the cells' colour");
  assert.equal(pattern.getAttribute("width"), "12", "8 cells of 1.5 px");
  assert.equal(el.querySelector(".hs-obj-geom > path").getAttribute("fill"), `url(#${pattern.getAttribute("id")})`);
  assert.equal(el.querySelector(".hs-obj").getAttribute("data-pattern"), "#1f3864,#dce4f2", "the brand check looks at the two colours");
  assert.equal(draw({}).querySelector(".hs-obj-geom pattern"), null);
  // Applying one: a plain fill, a gradient or a quick style take it off; the format painter and the defaults carry it; colours snap to the palette.
  for (const kind of ["shape", "text"]) assert.ok(ops.DEFAULT_STYLE_KEYS[kind].includes("pattern"), kind);
  const snapped = ops.snapToPalette({ id: "p", kind: "shape", fill: "#ff0000", pattern: { kind: "pct25", fg: "#101010", bg: "#fafafa" } }, E.PALETTE);
  assert.ok(E.PALETTE.fill.some(([c]) => c === snapped.pattern.fg) && E.PALETTE.fill.some(([c]) => c === snapped.pattern.bg), JSON.stringify(snapped.pattern));
});

test("3-D 回転: turns about x and y with a camera's field of view are kept in range and drawn as a CSS perspective turn", async () => {
  const { E } = await loadEngine();
  const shape = (o) => E.normalizeObject({ id: "a", kind: "shape", shape: "rect", x: 100, y: 100, w: 400, h: 200, fill: "#dce4f2", stroke: "none", ...o });
  assert.deepEqual(plainJson(shape({ rot3d: { x: 30.04, y: -20 } }).rot3d), { x: 30, y: -20 }, "45° is the usual view and is not kept");
  assert.deepEqual(plainJson(shape({ rot3d: { x: 0, y: 40, p: 0 } }).rot3d), { x: 0, y: 40, p: 0 }, "a flat view is kept");
  assert.deepEqual(plainJson(shape({ rot3d: { x: 500, y: -500, p: 500 } }).rot3d), { x: 89, y: -89, p: 120 }, "kept to what can be seen");
  assert.equal(shape({ rot3d: { x: 0.2, y: 0.1 } }).rot3d, undefined, "no turn, nothing kept");
  assert.equal(shape({ rot3d: "no" }).rot3d, undefined);
  assert.equal(E.normalizeObject({ id: "t", kind: "table", x: 0, y: 0, w: 400, h: 200, cols: [1], rows: [1], cells: [[{ text: "<p>a</p>" }]], rot3d: { x: 20, y: 20 } })?.rot3d, undefined, "shapes, text boxes and pictures only");
  const draw = (o) => { const slide = { type: "blank", elements: [shape(o)] }; return E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } }); };
  const turned = draw({ rot3d: { x: 20, y: -30 }, rot: 10, flipH: true }).querySelector(".hs-obj-rot").getAttribute("style");
  assert.ok(/transform:\s*perspective\(482\.\d+px\) rotateX\(20deg\) rotateY\(-30deg\) rotate\(10deg\) scale\(-1, 1\)/.test(turned), turned);
  const flat = draw({ rot3d: { x: 35, y: -45, p: 0 } }).querySelector(".hs-obj-rot").getAttribute("style");
  assert.ok(/transform:\s*rotateX\(35deg\) rotateY\(-45deg\)/.test(flat) && !/perspective/.test(flat), flat);
  assert.ok(!/rotateX/.test(draw({}).querySelector(".hs-obj-rot").getAttribute("style") || ""));
  // The gallery.
  assert.deepEqual(plainJson(ops.rot3dPreset("right")), { x: 0, y: 28 });
  assert.equal(ops.rot3dPreset("nowhere"), null);
  assert.equal(ops.rot3dKey({ x: 35, y: -45, p: 0 }), "isoTopLeft");
  assert.equal(ops.rot3dKey({ x: 12, y: 3 }), "");
  assert.equal(ops.rot3dKey(null), "");
  for (const [key, , turn] of ops.ROT3D_PRESETS) assert.equal(ops.rot3dKey(plainJson(shape({ rot3d: turn }).rot3d || {})), key, `${key} is found again after normalizing`);
  assert.equal(ops.effectsWords({ rot3d: {}, soft: 4 }), "ぼかし・3-D 回転");
});

test("文字の輪郭: a line round the letters behind their fill", async () => {
  const { E } = await loadEngine();
  const text = (o, kind = "text") => E.normalizeObject({ id: "t", kind, x: 100, y: 100, w: 600, h: 120, text: "<p>文字</p>", fs: 40, ...(kind === "shape" ? { shape: "rect", fill: "#dce4f2", stroke: "none" } : {}), ...o });
  assert.deepEqual(plainJson(text({ toutline: { w: 2, color: "#1F3864" } }).toutline), { w: 2, color: "#1f3864" });
  assert.equal(text({ toutline: { w: 0, color: "#1f3864" } }).toutline, undefined, "no width, no outline");
  assert.equal(text({ toutline: { w: 2, color: "nonsense" } }).toutline, undefined, "no colour, no outline");
  assert.equal(text({ toutline: { w: 900, color: "#808080" } }).toutline.w, 24, "kept to a sensible width");
  assert.equal(E.normalizeObject({ id: "i", kind: "icon", name: "check", x: 0, y: 0, w: 80, h: 80, toutline: { w: 2, color: "#1f3864" } })?.toutline, undefined, "only text takes it");
  const draw = (o, kind) => { const slide = { type: "blank", elements: [text(o, kind)] }; return E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } }); };
  for (const kind of ["text", "shape"]) {
    const el = draw({ toutline: { w: 2, color: "#1f3864" } }, kind);
    const style = el.querySelector(".hs-obj-tx").getAttribute("style");
    assert.ok(/-webkit-text-stroke:\s*2px #1f3864/.test(style) && /paint-order:\s*stroke fill/.test(style), `${kind}: ${style}`);
    assert.equal(el.querySelector(".hs-obj").getAttribute("data-toutline"), "#1f3864");
  }
  assert.ok(!/text-stroke/.test(draw({}).querySelector(".hs-obj-tx").getAttribute("style") || ""));
  // The choices.
  assert.deepEqual(plainJson(ops.textOutline("#808080", { w: 4, color: "#1f3864" })), { w: 4, color: "#808080" }, "a new colour keeps the width");
  assert.deepEqual(plainJson(ops.textOutline("#808080")), { w: 2, color: "#808080" });
  assert.ok(ops.TEXT_OUTLINE_COLORS.every(([c]) => /^#[0-9a-f]{6}$/.test(c)), "colours as #rrggbb");
  assert.equal(ops.textEffectsWords({ tshadow: {}, tglow: {}, toutline: {} }), "影・光彩・輪郭");
  for (const kind of ["shape", "text"]) assert.ok(ops.DEFAULT_STYLE_KEYS[kind].includes("toutline"), kind);
});

test("スライド マスター: the deck's master objects lie under every slide's own layer; a slide can leave them out", async () => {
  const { E } = await loadEngine();
  const master = E.normalizeObjects([{ id: "m1", kind: "text", x: 100, y: 900, w: 600, h: 60, text: "<p>プロジェクト名</p>" }, { id: "m2", kind: "shape", shape: "rect", x: 0, y: 0, w: 50, h: 50, fill: "#dce4f2" }, { id: "bad", kind: "nothing" }]);
  assert.equal(master.length, 2, "an unknown kind is dropped");
  const deck = { slides: [], theme: "clarity", masterObjects: master };
  const own = E.normalizeObjects([{ id: "s1", kind: "shape", shape: "rect", x: 300, y: 300, w: 100, h: 100 }]);
  const slide = { type: "blank", hideTitle: true, elements: own };
  const draw = (s) => E.render(s, { mode: "present", index: 1, deck: { ...deck, slides: [s] } });
  const el = draw(slide);
  const layer = el.querySelector(".hs-master-layer");
  assert.ok(layer, "the master layer");
  assert.equal(layer.getAttribute("aria-hidden"), "true", "a screen reader skips what repeats on every slide");
  assert.equal(layer.querySelectorAll(".hs-obj").length, 2);
  assert.ok(layer.textContent.includes("プロジェクト名"));
  // Under the layout (frame) and the slide's own objects in the page, so it is built before them.
  const kids = [...el.children].map((c) => c.getAttribute("class") || "");
  const at = (name) => kids.findIndex((c) => c.includes(name));
  assert.ok(at("hs-master-layer") > -1 && at("hs-master-layer") < at("hs-frame") && at("hs-frame") < at("hs-objects"), kids.join(" | "));
  assert.equal(el.querySelectorAll(".hs-objects > .hs-obj").length, 1, "the slide's own layer holds only its own objects");
  // 背景グラフィックを表示しない
  assert.equal(draw({ ...slide, hideMaster: true }).querySelector(".hs-master-layer"), null);
  // No master objects, no layer.
  assert.equal(E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } }).querySelector(".hs-master-layer"), null);
  // Hidden master objects are not drawn.
  const hidden = draw({ ...slide });
  assert.ok(hidden);
  const withHidden = E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity", masterObjects: E.normalizeObjects([{ id: "h", kind: "shape", shape: "rect", x: 0, y: 0, w: 10, h: 10, hidden: true }]) } });
  assert.equal(withHidden.querySelector(".hs-master-layer"), null, "only hidden objects: no layer");
});
