// ファイル → エクスポート → アニメーション GIF の作成 (PowerPoint's Export → Create an Animated GIF): the show plays by
// itself (as ビデオの作成 does) while frames of this tab are taken a few times a second, cut to the slide, made smaller
// and written into a looping GIF. Only what changed since the last frame is written (slides stand still most of the
// time), each frame with its own 256 colours (exact when the slide has few, median cut otherwise).

import { videoPlan } from "./video.mjs?v=__APP_VERSION__";

/** GIF's sizes (the slide's width; the height keeps 16:9), like PowerPoint's 極小・小・中・大. */
export const GIF_SIZES = [["xs", "極小（320×180）", 320], ["s", "小（640×360）", 640], ["m", "中（960×540）", 960], ["l", "大（1280×720）", 1280]];

/**
 * A picture's colours: at most `max` of them (exact when there are no more than that), and each pixel's index.
 * rgba is RGBA bytes; the palette is [r, g, b] triples.
 */
export function quantize(rgba, max = 256) {
  const n = rgba.length >> 2;
  const indices = new Uint8Array(n);
  // Few colours (flat slides): each its own entry.
  const exact = new Map();
  let i = 0;
  for (; i < n; i += 1) {
    const c = (rgba[i * 4] << 16) | (rgba[i * 4 + 1] << 8) | rgba[i * 4 + 2];
    let k = exact.get(c);
    if (k === undefined) {
      if (exact.size >= max) break;
      k = exact.size;
      exact.set(c, k);
    }
    indices[i] = k;
  }
  if (i === n) return { palette: [...exact.keys()].map((c) => [c >> 16, (c >> 8) & 255, c & 255]), indices };
  // Many (photos, smooth edges): median cut over 5 bits a channel.
  const hist = new Uint32Array(32768);
  const key = (p) => ((rgba[p] >> 3) << 10) | ((rgba[p + 1] >> 3) << 5) | (rgba[p + 2] >> 3);
  for (let p = 0; p < rgba.length; p += 4) hist[key(p)] += 1;
  const colors = [];
  for (let c = 0; c < 32768; c += 1) if (hist[c]) colors.push(c);
  const ch = (c, s) => (c >> (10 - s * 5)) & 31;
  const boxOf = (list) => {
    const lo = [31, 31, 31];
    const hi = [0, 0, 0];
    let count = 0;
    for (const c of list) {
      for (let s = 0; s < 3; s += 1) { const v = ch(c, s); if (v < lo[s]) lo[s] = v; if (v > hi[s]) hi[s] = v; }
      count += hist[c];
    }
    return { list, lo, hi, count };
  };
  const boxes = [boxOf(colors)];
  while (boxes.length < max) {
    // The box to split: the most pixels across the widest spread.
    let best = -1;
    let score = 0;
    boxes.forEach((b, k) => {
      if (b.list.length < 2) return;
      const spread = Math.max(b.hi[0] - b.lo[0], b.hi[1] - b.lo[1], b.hi[2] - b.lo[2]);
      const s = spread * Math.sqrt(b.count);
      if (s > score) { score = s; best = k; }
    });
    if (best < 0) break;
    const b = boxes[best];
    const spreads = [0, 1, 2].map((s) => b.hi[s] - b.lo[s]);
    const s = spreads.indexOf(Math.max(...spreads));
    const sorted = [...b.list].sort((x, y) => ch(x, s) - ch(y, s));
    let half = 0;
    let cut = 1;
    for (let k = 0; k < sorted.length - 1; k += 1) {
      half += hist[sorted[k]];
      cut = k + 1;
      if (half * 2 >= b.count) break;
    }
    boxes.splice(best, 1, boxOf(sorted.slice(0, cut)), boxOf(sorted.slice(cut)));
  }
  const lut = new Uint8Array(32768);
  const palette = boxes.map((b, k) => {
    let r = 0;
    let g = 0;
    let bl = 0;
    for (const c of b.list) {
      lut[c] = k;
      r += ((ch(c, 0) << 3) | 4) * hist[c];
      g += ((ch(c, 1) << 3) | 4) * hist[c];
      bl += ((ch(c, 2) << 3) | 4) * hist[c];
    }
    return [Math.round(r / b.count), Math.round(g / b.count), Math.round(bl / b.count)];
  });
  for (let p = 0, k = 0; p < rgba.length; p += 4, k += 1) indices[k] = lut[key(p)];
  return { palette, indices };
}

/** GIF's LZW (variable codes up to 12 bits, a clear code when the table fills): the bytes, before sub-blocking. */
export function lzwEncode(indices, minCodeSize) {
  const clear = 1 << minCodeSize;
  const eoi = clear + 1;
  let size = minCodeSize + 1;
  let next = eoi + 1;
  let table = new Map();
  const out = [];
  let cur = 0;
  let bits = 0;
  const emit = (code) => {
    cur |= code << bits;
    bits += size;
    while (bits >= 8) { out.push(cur & 255); cur >>>= 8; bits -= 8; }
  };
  emit(clear);
  if (!indices.length) { emit(eoi); if (bits > 0) out.push(cur & 255); return Uint8Array.from(out); }
  let code = indices[0];
  for (let i = 1; i < indices.length; i += 1) {
    const k = indices[i];
    const keyed = (code << 8) | k;
    const found = table.get(keyed);
    if (found !== undefined) { code = found; continue; }
    emit(code);
    if (next === 4096) {
      emit(clear);
      next = eoi + 1;
      size = minCodeSize + 1;
      table = new Map();
    } else {
      if (next >= 1 << size) size += 1;
      table.set(keyed, next);
      next += 1;
    }
    code = k;
  }
  emit(code);
  emit(eoi);
  if (bits > 0) out.push(cur & 255);
  return Uint8Array.from(out);
}

/** The box where two same-sized pictures differ ({ x, y, w, h }), or null when they are the same. */
export function changedRect(a, b, width) {
  const A = new Uint32Array(a.buffer, a.byteOffset, a.length >> 2);
  const B = new Uint32Array(b.buffer, b.byteOffset, b.length >> 2);
  const height = A.length / width;
  let top = -1;
  for (let y = 0; y < height && top < 0; y += 1) for (let x = 0, p = y * width; x < width; x += 1, p += 1) if (A[p] !== B[p]) { top = y; break; }
  if (top < 0) return null;
  let bottom = top;
  for (let y = height - 1; y > top; y -= 1) {
    let diff = false;
    for (let x = 0, p = y * width; x < width; x += 1, p += 1) if (A[p] !== B[p]) { diff = true; break; }
    if (diff) { bottom = y; break; }
  }
  let left = width;
  let right = -1;
  for (let y = top; y <= bottom; y += 1) {
    for (let x = 0, p = y * width; x < left; x += 1, p += 1) if (A[p] !== B[p]) { left = x; break; }
    for (let x = width - 1, p = y * width + width - 1; x > right; x -= 1, p -= 1) if (A[p] !== B[p]) { right = x; break; }
  }
  return { x: left, y: top, w: right - left + 1, h: bottom - top + 1 };
}

/** Writes a looping GIF frame by frame: add(rgba of the whole picture, delay in 1/100 s, the part to write). */
export class GifWriter {
  constructor(width, height, { loop = 0 } = {}) {
    this.width = width;
    this.height = height;
    this.parts = [];
    this.frames = 0;
    const head = new Uint8Array(13);
    head.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]); // GIF89a
    head[6] = width & 255; head[7] = width >> 8; head[8] = height & 255; head[9] = height >> 8;
    head[10] = 0x70; // no global colour table; 8 bits a channel
    this.parts.push(head);
    if (loop != null) this.parts.push(Uint8Array.from([0x21, 0xff, 0x0b, ...[..."NETSCAPE2.0"].map((c) => c.charCodeAt(0)), 0x03, 0x01, loop & 255, loop >> 8, 0x00]));
  }

  add(rgba, delay, rect = null) {
    const { x, y, w, h } = rect || { x: 0, y: 0, w: this.width, h: this.height };
    let pixels = rgba;
    if (w !== this.width || h !== this.height) {
      pixels = new Uint8Array(w * h * 4);
      for (let row = 0; row < h; row += 1) pixels.set(rgba.subarray(((y + row) * this.width + x) * 4, ((y + row) * this.width + x + w) * 4), row * w * 4);
    }
    const { palette, indices } = quantize(pixels, 256);
    const depth = Math.max(1, Math.ceil(Math.log2(Math.max(2, palette.length))));
    const d = Math.max(2, Math.min(65535, Math.round(delay)));
    // Graphic control: leave the frame in place (the next one draws over the part that changed).
    this.parts.push(Uint8Array.from([0x21, 0xf9, 0x04, 0x04, d & 255, d >> 8, 0x00, 0x00]));
    this.parts.push(Uint8Array.from([0x2c, x & 255, x >> 8, y & 255, y >> 8, w & 255, w >> 8, h & 255, h >> 8, 0x80 | (depth - 1)]));
    const table = new Uint8Array(3 << depth);
    palette.forEach(([r, g, b], k) => { table[k * 3] = r; table[k * 3 + 1] = g; table[k * 3 + 2] = b; });
    this.parts.push(table);
    const min = Math.max(2, depth);
    const data = lzwEncode(indices, min);
    const blocks = new Uint8Array(1 + data.length + Math.ceil(data.length / 255) + 1);
    blocks[0] = min;
    let at = 1;
    for (let p = 0; p < data.length; p += 255) {
      const len = Math.min(255, data.length - p);
      blocks[at] = len;
      blocks.set(data.subarray(p, p + len), at + 1);
      at += len + 1;
    }
    blocks[at] = 0;
    this.parts.push(blocks.subarray(0, at + 1));
    this.frames += 1;
  }

  /** The finished file's bytes. */
  finish() {
    const all = [...this.parts, Uint8Array.from([0x3b])];
    const out = new Uint8Array(all.reduce((n, p) => n + p.length, 0));
    let at = 0;
    for (const p of all) { out.set(p, at); at += p.length; }
    return out;
  }
}

export function createGifExport(app) {
  const { h } = app;
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));

  function openDialog() {
    const deck = app.deck();
    if (!deck) return;
    const size = h("select", { "aria-label": "サイズ", class: "gif-size" }, ...GIF_SIZES.map(([v, label]) => h("option", { value: v, selected: v === "s" }, label)));
    const secs = h("input", { type: "number", min: "1", max: "60", step: "0.5", value: "3", class: "sh-num gif-secs" });
    const timings = h("input", { type: "checkbox", checked: true });
    const order = app.showOrder();
    const from = h("input", { type: "number", min: "1", max: String(order.length), value: "1", class: "sh-num gif-from" });
    const to = h("input", { type: "number", min: "1", max: String(order.length), value: String(order.length), class: "sh-num gif-to" });
    const estimate = h("p", { class: "hint gif-estimate" });
    const range = () => {
      const a = Math.max(1, Math.min(order.length, Number(from.value) || 1));
      const b = Math.max(a, Math.min(order.length, Number(to.value) || order.length));
      return order.slice(a - 1, b);
    };
    const update = () => {
      const { total } = videoPlan(range(), deck.slides, { seconds: Number(secs.value) || 3, useTimings: timings.checked });
      estimate.textContent = `GIF の長さ：約${Math.round(total / 1000)}秒（${range().length}枚）`;
    };
    for (const el of [secs, from, to]) el.addEventListener("input", update);
    timings.addEventListener("change", update);
    const dialog = h("dialog", { class: "gif-dialog", "aria-label": "アニメーション GIF の作成" },
      h("div", { class: "dialog-head" }, h("h3", {}, "アニメーション GIF の作成"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body sh-form-col" },
        h("p", {}, "スライド ショーを自動で流し、アニメーション・画面切り替えごと、繰り返し再生される GIF にします。チャットやメールにそのまま貼れます（音は入りません）。"),
        h("label", { class: "sh-inline" }, "サイズ ", size),
        h("label", { class: "sh-choice" }, timings, h("span", {}, "記録されたタイミングを使用する")),
        h("label", { class: "sh-inline" }, "各スライドの表示時間 ", secs, " 秒"),
        h("div", { class: "sh-inline" }, "スライド ", from, " から ", to, " まで"),
        estimate,
        h("p", { class: "hint" }, "「作成」を押すとブラウザが共有する画面を聞いてきます。「このタブ」を選んで共有してください。作成中は操作しないでください（Escで中止）。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary gif-go", onclick: () => { const slides = range(); dialog.close(); record({ size: size.value, seconds: Number(secs.value) || 3, useTimings: timings.checked, slides }); } }, "作成")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
    update();
  }

  async function record({ size = "s", seconds = 3, useTimings = true, slides = null } = {}) {
    const deck = app.deck();
    if (!deck || app.player()) return;
    const order = slides?.length ? slides : app.showOrder();
    const width = (GIF_SIZES.find(([v]) => v === size) || GIF_SIZES[1])[2];
    const height = Math.round(width * 9 / 16);
    let stream;
    try {
      stream = await navigator.mediaDevices.getDisplayMedia({ video: { frameRate: 30, displaySurface: "browser" }, audio: false, preferCurrentTab: true, selfBrowserSurface: "include", surfaceSwitching: "exclude" });
    } catch (error) {
      if (error?.name !== "NotAllowedError") app.toast(`画面を共有できませんでした（${error.message || error.name}）`);
      return;
    }
    const video = h("video", { muted: true, playsinline: true });
    video.srcObject = stream;
    await video.play().catch(() => {});
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const g = canvas.getContext("2d", { willReadFrequently: true });
    const gif = new GifWriter(width, height);
    let prev = null;
    let pending = null;
    let stop = false;
    let cancelled = false;
    const take = async () => {
      // The slide's place: 16:9 in the middle of the stage (slides move about inside it while they change).
      const stage = document.querySelector("#presenter .hs-player-stage");
      if (!stage || video.readyState < 2) return;
      const bitmap = await createImageBitmap(video);
      const s = stage.getBoundingClientRect();
      const fw = Math.min(s.width, s.height * 16 / 9);
      const fh = fw * 9 / 16;
      const r = { left: s.left + (s.width - fw) / 2, top: s.top + (s.height - fh) / 2, width: fw, height: fh };
      const kx = bitmap.width / window.innerWidth;
      const ky = bitmap.height / window.innerHeight;
      g.fillStyle = "#ffffff";
      g.fillRect(0, 0, width, height);
      g.drawImage(bitmap, r.left * kx, r.top * ky, r.width * kx, r.height * ky, 0, 0, width, height);
      bitmap.close?.();
      const now = performance.now();
      const data = new Uint8Array(g.getImageData(0, 0, width, height).data.buffer);
      const rect = prev ? changedRect(prev, data, width) : { x: 0, y: 0, w: width, h: height };
      if (!rect) return;
      if (pending) gif.add(pending.data, (now - pending.at) / 10, pending.rect);
      pending = { data, rect, at: now };
      prev = data;
    };
    const onKey = (event) => { if (event.key === "Escape") { cancelled = true; stop = true; } };
    window.addEventListener("keydown", onKey, true);
    stream.getVideoTracks()[0]?.addEventListener("ended", () => { stop = true; });
    try {
      await app.presentForVideo({ kiosk: true, kioskSeconds: seconds, useTimings: true, ignoreTimings: !useTimings, stopAtEnd: true, narration: false, captions: false, only: order }, { fullscreen: false });
      app.player()?.el.classList.add("hs-capture");
      await wait(500);
      const { plan } = videoPlan(order, deck.slides, { seconds, useTimings });
      const lastIndex = plan.at(-1)?.index;
      const lastMs = plan.at(-1)?.ms ?? seconds * 1000;
      let arrived = null;
      app.onShowClosed(() => { stop = true; });
      while (!stop) {
        const started = performance.now();
        await take();
        const player = app.player();
        if (!player) break;
        if (player.index === lastIndex) {
          arrived ??= Date.now();
          if (Date.now() - arrived > lastMs + 400 && !app.animBusy()) break;
        }
        app.progress?.(`GIF を作っています… ${gif.frames + 1}コマ`);
        await wait(Math.max(10, 100 - (performance.now() - started)));
      }
      if (pending && !cancelled) gif.add(pending.data, Math.max(100, (performance.now() - pending.at) / 10), pending.rect);
    } finally {
      window.removeEventListener("keydown", onKey, true);
      stream.getTracks().forEach((t) => t.stop());
      app.closeShow();
    }
    if (cancelled || !gif.frames) { app.toast(cancelled ? "GIF の作成を中止しました" : "GIF を作れませんでした"); return; }
    const blob = new Blob([gif.finish()], { type: "image/gif" });
    app.download(blob, `${(deck.title || "スライド").replace(/[\\/:*?"<>|]/g, "_")}.gif`);
    app.toast(`アニメーション GIF を作成しました（${gif.frames}コマ・${Math.round(blob.size / 1024 / 1024 * 10) / 10}MB）`);
    return blob;
  }

  return { openDialog, record };
}
