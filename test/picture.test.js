// 図の形式 (picture.mjs): 背景の削除 from the edges, 透明色を指定, and the size of a compressed picture.
import assert from "node:assert/strict";
import test from "node:test";

import { compressedSize, makeTransparent, removeBackground } from "../public/editor/picture.mjs";

/** A w × h picture: a white background with a navy square in the middle (and a white dot inside the square). */
function picture(w = 40, h = 30) {
  const data = new Uint8ClampedArray(w * h * 4);
  for (let y = 0; y < h; y += 1) for (let x = 0; x < w; x += 1) {
    const i = (y * w + x) * 4;
    const inside = x >= 10 && x < 30 && y >= 8 && y < 22;
    const dot = x === 20 && y === 15;
    const c = inside && !dot ? [31, 56, 100] : [255, 255, 255];
    data.set([...c, 255], i);
  }
  return data;
}
const alpha = (data, w, x, y) => data[(y * w + x) * 4 + 3];

test("背景の削除 clears what the edges reach, keeps the subject and what is enclosed by it", () => {
  const w = 40;
  const data = picture();
  const share = removeBackground(data, w, 30);
  assert.equal(alpha(data, w, 0, 0), 0, "the corner goes");
  assert.equal(alpha(data, w, 5, 15), 0, "the background beside the subject goes");
  assert.equal(alpha(data, w, 15, 15), 255, "the subject stays");
  assert.equal(alpha(data, w, 20, 15), 255, "a white dot inside the subject is not background");
  assert.ok(Math.abs(share - (1200 - 280) / 1200) < 0.01, `share ${share}`);
});

test("透明色を指定 clears every pixel of that colour, wherever it is", () => {
  const w = 40;
  const data = picture();
  const n = makeTransparent(data, [255, 255, 255]);
  assert.equal(n, 1200 - 280 + 1);
  assert.equal(alpha(data, w, 20, 15), 0, "the dot too");
  assert.equal(alpha(data, w, 15, 15), 255);
});

test("compressed size keeps the shape and never enlarges", () => {
  assert.deepEqual(compressedSize(4000, 3000, 1600), [1600, 1200]);
  assert.deepEqual(compressedSize(800, 600, 1600), [800, 600]);
  assert.deepEqual(compressedSize(1000, 4000, 1280), [320, 1280]);
});

test("アート効果: blur softens an edge, 線画 finds it, モザイク makes blocks, ポスター keeps four levels; 色の変更 tones in the theme's accents", async () => {
  const { artEffect, recolor, ART_EFFECTS, RECOLORS } = await import("../public/editor/picture.mjs");
  const w = 60;
  const h = 40;
  // Left half orange, right half blue.
  const make = () => { const d = new Uint8ClampedArray(w * h * 4); for (let i = 0; i < d.length; i += 4) { const x = (i / 4) % w; d.set(x < 30 ? [220, 120, 40, 255] : [30, 80, 200, 255], i); } return d; };
  const px = (d, x, y) => Array.from(d.slice((y * w + x) * 4, (y * w + x) * 4 + 3));
  assert.deepEqual(Object.keys(ART_EFFECTS), ["blur", "pencil", "lines", "mosaic", "sepia", "poster"]);
  const blur = artEffect(make(), w, h, "blur");
  assert.deepEqual(px(blur, 2, 20), [220, 120, 40], "far from the edge: as it was");
  assert.ok(px(blur, 30, 20)[0] < 200 && px(blur, 30, 20)[0] > 50, "at the edge: in between");
  const lines = artEffect(make(), w, h, "lines");
  assert.ok(px(lines, 5, 20)[0] > 240 && px(lines, 30, 20)[0] < 200, "white inside, dark on the edge");
  const sketch = artEffect(make(), w, h, "pencil");
  assert.equal(px(sketch, 5, 20)[0], px(sketch, 5, 20)[1], "grey");
  const mosaic = artEffect(make(), w, h, "mosaic");
  assert.deepEqual(px(mosaic, 0, 0), px(mosaic, 5, 5), "one block, one colour");
  const poster = artEffect(make(), w, h, "poster");
  assert.ok(px(poster, 5, 5).every((v) => [0, 85, 170, 255].includes(v)));
  assert.deepEqual(Object.keys(RECOLORS), ["gray", "sepia", "wash", "accent1Dark", "accent1Light", "accent2Dark", "accent2Light"]);
  const gray = recolor(make(), "gray");
  assert.equal(new Set(px(gray, 5, 5)).size, 1);
  const navy = recolor(make(), "accent1Dark", { accent1: "#1f3864", accent2: "#c00000" });
  const [r, g, b] = px(navy, 50, 20);
  assert.ok(b >= g && g >= r, "a tone of the accent (a navy one: blue strongest)");
  const wash = recolor(make(), "wash");
  assert.ok(px(wash, 50, 20).every((v) => v > 180), "washed out: light");
});

test("画像として保存: a ZIP of the pictures (stored, UTF-8 names) that unzip reads back", async () => {
  const { zipFiles, crc32, imageName } = await import("../public/editor/imagexport.mjs");
  assert.equal(crc32(new TextEncoder().encode("123456789")), 0xcbf43926, "the standard CRC-32 check value");
  assert.equal(imageName(3, 12, "png"), "スライド03.png");
  assert.equal(imageName(7, 120, "jpg"), "スライド007.jpg");
  const files = [{ name: "スライド01.png", data: new Uint8Array([137, 80, 78, 71, 1, 2, 3]) }, { name: "スライド02.png", data: new Uint8Array([9, 9, 9]) }];
  const zip = zipFiles(files, new Date(2026, 9, 3, 12, 0, 0));
  const { execFileSync } = await import("node:child_process");
  const { mkdtemp, writeFile } = await import("node:fs/promises");
  const { tmpdir } = await import("node:os");
  const { join } = await import("node:path");
  const dir = await mkdtemp(join(tmpdir(), "imgx-"));
  await writeFile(join(dir, "a.zip"), zip);
  const listing = execFileSync("python3", ["-c", `import zipfile,sys;z=zipfile.ZipFile(sys.argv[1]);print(z.testzip());print("|".join(i.filename for i in z.infolist()));print(z.read("スライド02.png").hex())`, join(dir, "a.zip")]).toString().trim().split("\n");
  assert.deepEqual(listing, ["None", "スライド01.png|スライド02.png", "090909"]);
});
