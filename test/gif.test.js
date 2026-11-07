// アニメーション GIF の作成: the colours (exact when few, median cut otherwise), GIF's LZW, only the changed part of a
// frame, and the file read back by a small decoder to the same pixels.
import assert from "node:assert/strict";
import test from "node:test";

import { GifWriter, changedRect, lzwEncode, quantize } from "../public/editor/gif.mjs";

/** GIF's LZW, read back (min code size, the sub-blocked bytes joined) → indices. */
function lzwDecode(bytes, min, count) {
  const clear = 1 << min;
  const eoi = clear + 1;
  let size = min + 1;
  let dict = [];
  const reset = () => { dict = []; for (let i = 0; i < clear; i += 1) dict[i] = [i]; dict[clear] = []; dict[eoi] = null; size = min + 1; };
  reset();
  const out = [];
  let pos = 0;
  const read = () => {
    let code = 0;
    for (let i = 0; i < size; i += 1, pos += 1) if (bytes[pos >> 3] & (1 << (pos & 7))) code |= 1 << i;
    return code;
  };
  let prev = null;
  while (pos < bytes.length * 8) {
    const code = read();
    if (code === clear) { reset(); prev = null; continue; }
    if (code === eoi) break;
    let entry;
    if (code < dict.length) entry = dict[code];
    else if (code === dict.length && prev) entry = [...prev, prev[0]];
    else throw new Error(`bad code ${code} at ${out.length}`);
    out.push(...entry);
    if (prev) dict.push([...prev, entry[0]]);
    prev = entry;
    if (dict.length === 1 << size && size < 12) size += 1;
  }
  assert.equal(out.length, count, "every pixel");
  return out;
}

/** A small GIF reader: the frames ({ x, y, w, h, delay, palette, indices }) and the screen size. */
function readGif(bytes) {
  assert.equal(String.fromCharCode(...bytes.slice(0, 6)), "GIF89a");
  const width = bytes[6] | (bytes[7] << 8);
  const height = bytes[8] | (bytes[9] << 8);
  let p = 13;
  const frames = [];
  let delay = 0;
  let loops = false;
  while (bytes[p] !== 0x3b) {
    if (bytes[p] === 0x21) {
      const label = bytes[p + 1];
      if (label === 0xf9) delay = bytes[p + 4] | (bytes[p + 5] << 8);
      if (label === 0xff) loops = true;
      p += 2;
      while (bytes[p]) p += bytes[p] + 1;
      p += 1;
    } else if (bytes[p] === 0x2c) {
      const x = bytes[p + 1] | (bytes[p + 2] << 8);
      const y = bytes[p + 3] | (bytes[p + 4] << 8);
      const w = bytes[p + 5] | (bytes[p + 6] << 8);
      const h = bytes[p + 7] | (bytes[p + 8] << 8);
      const packed = bytes[p + 9];
      p += 10;
      assert.ok(packed & 0x80, "a local colour table");
      const n = 2 << (packed & 7);
      const palette = [];
      for (let i = 0; i < n; i += 1) palette.push([bytes[p + i * 3], bytes[p + i * 3 + 1], bytes[p + i * 3 + 2]]);
      p += n * 3;
      const min = bytes[p];
      p += 1;
      const data = [];
      while (bytes[p]) { data.push(...bytes.slice(p + 1, p + 1 + bytes[p])); p += bytes[p] + 1; }
      p += 1;
      frames.push({ x, y, w, h, delay, palette, indices: lzwDecode(Uint8Array.from(data), min, w * h) });
    } else throw new Error(`unexpected block ${bytes[p]} at ${p}`);
  }
  return { width, height, frames, loops };
}

const picture = (w, h, at) => {
  const px = new Uint8Array(w * h * 4);
  for (let y = 0; y < h; y += 1) for (let x = 0; x < w; x += 1) { const [r, g, b] = at(x, y); px.set([r, g, b, 255], (y * w + x) * 4); }
  return px;
};

test("GIF colours: exact when there are few, at most 256 close ones otherwise", () => {
  const flat = picture(8, 4, (x) => (x < 4 ? [0x1f, 0x38, 0x64] : [0xff, 0xff, 0xff]));
  const q = quantize(flat);
  assert.deepEqual(q.palette, [[0x1f, 0x38, 0x64], [255, 255, 255]]);
  assert.deepEqual([...q.indices.slice(0, 8)], [0, 0, 0, 0, 1, 1, 1, 1]);
  const smooth = picture(64, 64, (x, y) => [x * 4, y * 4, (x * y) & 255]);
  const m = quantize(smooth);
  assert.ok(m.palette.length <= 256 && m.palette.length > 100, `${m.palette.length} colours`);
  // Each pixel near its colour.
  let worst = 0;
  for (let i = 0; i < 64 * 64; i += 1) {
    const [r, g, b] = m.palette[m.indices[i]];
    worst = Math.max(worst, Math.abs(r - smooth[i * 4]), Math.abs(g - smooth[i * 4 + 1]), Math.abs(b - smooth[i * 4 + 2]));
  }
  assert.ok(worst <= 48, `worst channel error ${worst}`);
});

test("GIF's LZW reads back, through the 12-bit table filling up and clearing", () => {
  const random = Uint8Array.from({ length: 20000 }, (_, i) => (i * 7919 + (i >> 3) * 31) % 256);
  assert.deepEqual(lzwDecode(lzwEncode(random, 8), 8, random.length), [...random]);
  const runs = Uint8Array.from({ length: 5000 }, (_, i) => (i >> 6) & 3);
  const packed = lzwEncode(runs, 2);
  assert.ok(packed.length < 600, `runs compress (${packed.length} bytes)`);
  assert.deepEqual(lzwDecode(packed, 2, runs.length), [...runs]);
  assert.deepEqual(lzwDecode(lzwEncode(Uint8Array.of(1), 2), 2, 1), [1]);
});

test("Only what changed is written; the file reads back to the same pixels, looping", () => {
  const w = 40;
  const h = 24;
  const a = picture(w, h, () => [255, 255, 255]);
  const b = picture(w, h, (x, y) => (x >= 10 && x < 20 && y >= 5 && y < 9 ? [0xdc, 0xe4, 0xf2] : [255, 255, 255]));
  assert.equal(changedRect(a, a.slice(), w), null);
  assert.deepEqual(changedRect(a, b, w), { x: 10, y: 5, w: 10, h: 4 });
  const gif = new GifWriter(w, h);
  gif.add(a, 120);
  gif.add(b, 50, changedRect(a, b, w));
  const file = readGif(gif.finish());
  assert.equal(file.width, w);
  assert.equal(file.height, h);
  assert.ok(file.loops, "loops for ever");
  assert.equal(file.frames.length, 2);
  assert.deepEqual(file.frames.map((f) => f.delay), [120, 50]);
  assert.deepEqual([file.frames[1].x, file.frames[1].y, file.frames[1].w, file.frames[1].h], [10, 5, 10, 4]);
  // Draw the frames over each other: the second picture.
  const screen = new Uint8Array(w * h * 3);
  for (const f of file.frames) f.indices.forEach((k, i) => screen.set(f.palette[k], (((f.y + Math.floor(i / f.w)) * w) + f.x + (i % f.w)) * 3));
  for (let i = 0; i < w * h; i += 1) assert.deepEqual([...screen.slice(i * 3, i * 3 + 3)], [b[i * 4], b[i * 4 + 1], b[i * 4 + 2]], `pixel ${i}`);
});
