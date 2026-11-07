// ファイル → エクスポート → 画像として保存 (PowerPoint's Save As → PNG / JPEG): each slide, exactly as the show draws
// it (animations finished, 3D models and videos as they appear), saved as a picture — one file for the current slide,
// a ZIP of all of them otherwise. The browser shares this tab once (as ビデオの作成 does); each slide is shown in a
// still slide show and its frame is cut out of the shared picture.

const CRC = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();
/** CRC-32 of bytes (as ZIP wants it). */
export function crc32(bytes) {
  let c = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) c = CRC[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** A ZIP (stored, no compression — pictures are compressed already) of [{ name, data: Uint8Array }]: its bytes. */
export function zipFiles(files, date = new Date()) {
  const enc = new TextEncoder();
  const time = ((date.getHours() & 31) << 11) | ((date.getMinutes() & 63) << 5) | ((date.getSeconds() / 2) & 31);
  const day = (((date.getFullYear() - 1980) & 127) << 9) | (((date.getMonth() + 1) & 15) << 5) | (date.getDate() & 31);
  const parts = [];
  const central = [];
  let offset = 0;
  for (const file of files) {
    const name = enc.encode(file.name);
    const crc = crc32(file.data);
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true); // names in UTF-8
    local.setUint16(8, 0, true);
    local.setUint16(10, time, true);
    local.setUint16(12, day, true);
    local.setUint32(14, crc, true);
    local.setUint32(18, file.data.length, true);
    local.setUint32(22, file.data.length, true);
    local.setUint16(26, name.length, true);
    local.setUint16(28, 0, true);
    parts.push(new Uint8Array(local.buffer), name, file.data);
    const entry = new DataView(new ArrayBuffer(46));
    entry.setUint32(0, 0x02014b50, true);
    entry.setUint16(4, 20, true);
    entry.setUint16(6, 20, true);
    entry.setUint16(8, 0x0800, true);
    entry.setUint16(10, 0, true);
    entry.setUint16(12, time, true);
    entry.setUint16(14, day, true);
    entry.setUint32(16, crc, true);
    entry.setUint32(20, file.data.length, true);
    entry.setUint32(24, file.data.length, true);
    entry.setUint16(28, name.length, true);
    entry.setUint32(42, offset, true);
    central.push(new Uint8Array(entry.buffer), name);
    offset += 30 + name.length + file.data.length;
  }
  const size = central.reduce((n, p) => n + p.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, files.length, true);
  end.setUint16(10, files.length, true);
  end.setUint32(12, size, true);
  end.setUint32(16, offset, true);
  const all = [...parts, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((n, p) => n + p.length, 0));
  let at = 0;
  for (const p of all) { out.set(p, at); at += p.length; }
  return out;
}

/** The file names: スライド01.png … (numbered by place in the show). */
export const imageName = (n, total, ext) => `スライド${String(n).padStart(Math.max(2, String(total).length), "0")}.${ext}`;

export function createImageExport(app) {
  const { h } = app;
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));

  function openDialog() {
    const deck = app.deck();
    if (!deck) return;
    const format = h("select", { "aria-label": "形式" }, h("option", { value: "png" }, "PNG（くっきり・透過なし）"), h("option", { value: "jpeg" }, "JPEG（軽い）"));
    const all = h("input", { type: "radio", name: "ix-range", value: "all", checked: true });
    const one = h("input", { type: "radio", name: "ix-range", value: "one" });
    const dialog = h("dialog", { class: "imgx-dialog", "aria-label": "画像として保存" },
      h("div", { class: "dialog-head" }, h("h3", {}, "画像として保存"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body sh-form-col" },
        h("label", { class: "sh-inline" }, "ファイルの種類 ", format),
        h("label", { class: "sh-choice" }, all, h("span", {}, `すべてのスライド（${app.showOrder().length}枚・ZIP）`)),
        h("label", { class: "sh-choice" }, one, h("span", {}, "現在のスライドのみ")),
        h("p", { class: "hint" }, "スライド ショーで見えるとおり（アニメーションの終わった状態・3D モデル・動画も）を、画面の解像度で画像にします。「保存」を押すとブラウザが共有する画面を聞いてきます。「このタブ」を選んでください。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary ix-go", onclick: () => { dialog.close(); save({ format: format.value, current: one.checked }); } }, "保存")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
  }

  /**
   * One frame of the shared tab, taken whole (a shared tab starts small and sharpens, so the frame's own size, not
   * the video's, says how it maps to the window); it waits a little for the full resolution.
   */
  async function frame(video) {
    const want = Math.round(window.innerWidth * (window.devicePixelRatio || 1) * 0.98);
    let bitmap = null;
    for (let t = 0; t < 25; t += 1) {
      bitmap?.close?.();
      bitmap = await createImageBitmap(video);
      if (bitmap.width >= want) break;
      await wait(100);
    }
    return bitmap;
  }

  /** The frame cut to the slide on screen, as a PNG or JPEG blob. */
  async function grab(video, type) {
    const slide = document.querySelector("#presenter .hs-player-stage > .hs-player-slide:last-child > .hs-slide") || [...document.querySelectorAll("#presenter .hs-player-stage .hs-slide")].pop();
    const bitmap = await frame(video);
    const r = slide.getBoundingClientRect();
    const kx = bitmap.width / window.innerWidth;
    const ky = bitmap.height / window.innerHeight;
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(r.width * kx));
    canvas.height = Math.max(1, Math.round(r.height * ky));
    const g = canvas.getContext("2d");
    if (type === "image/jpeg") { g.fillStyle = "#ffffff"; g.fillRect(0, 0, canvas.width, canvas.height); }
    g.drawImage(bitmap, r.left * kx, r.top * ky, r.width * kx, r.height * ky, 0, 0, canvas.width, canvas.height);
    bitmap.close?.();
    return new Promise((resolve) => canvas.toBlob(resolve, type, 0.92));
  }

  async function save({ format = "png", current = false } = {}) {
    const deck = app.deck();
    if (!deck || app.player()) return;
    const type = format === "jpeg" ? "image/jpeg" : "image/png";
    const ext = format === "jpeg" ? "jpg" : "png";
    const order = current ? [app.index()] : app.showOrder();
    let stream;
    try {
      stream = await navigator.mediaDevices.getDisplayMedia({ video: { frameRate: 15, displaySurface: "browser" }, audio: false, preferCurrentTab: true, selfBrowserSurface: "include", surfaceSwitching: "exclude" });
    } catch (error) {
      if (error?.name !== "NotAllowedError") app.toast(`画面を共有できませんでした（${error.message || error.name}）`);
      return;
    }
    const video = h("video", { muted: true, playsinline: true });
    video.srcObject = stream;
    await video.play().catch(() => {});
    const files = [];
    try {
      // A still show (every build shown, nothing moving), without its bar, in the window (the picture shared).
      await app.openPresenter(order[0], { fullscreen: false, extra: { static: true, captions: false, narration: false, closable: false } });
      app.player()?.el.classList.add("hs-capture");
      for (const [k, index] of order.entries()) {
        if (app.player()?.index !== index) app.player()?.go(index);
        await wait(700);
        // 3D models draw themselves once three.js is ready.
        for (let t = 0; t < 30 && document.querySelector("#presenter .hs-model[data-src]:not(.is-live):not(.is-broken)"); t += 1) await wait(150);
        await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(r)));
        const blob = await grab(video, type);
        if (blob) files.push({ name: imageName(k + 1, order.length, ext), blob });
        app.progress?.(`画像にしています… ${k + 1} / ${order.length}`);
      }
    } finally {
      stream.getTracks().forEach((t) => t.stop());
      app.closeShow();
    }
    const base = (deck.title || "スライド").replace(/[\\/:*?"<>|]/g, "_");
    if (!files.length) { app.toast("画像にできませんでした"); return; }
    if (files.length === 1) {
      app.download(files[0].blob, `${base}_${files[0].name}`);
    } else {
      const zip = zipFiles(await Promise.all(files.map(async (f) => ({ name: f.name, data: new Uint8Array(await f.blob.arrayBuffer()) }))));
      app.download(new Blob([zip], { type: "application/zip" }), `${base}_画像.zip`);
    }
    app.toast(`${files.length}枚の画像を保存しました`);
  }

  return { openDialog, save };
}
