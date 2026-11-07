// 図の形式 → 背景の削除・色（透明色を指定）・図の圧縮 (PowerPoint's Remove Background, Set Transparent Color and
// Compress Pictures). The picture is worked on in a canvas and saved as a new picture kept in this browser; the old one
// stays in the undo history.

/** Distance between two colours (RGB). */
const dist = (d, i, c) => Math.hypot(d[i] - c[0], d[i + 1] - c[1], d[i + 2] - c[2]);

/**
 * 背景の削除: the plain background around the subject goes. Pixels reached from the picture's edges whose colour is
 * close to an edge colour become transparent (softly at the border of the subject). `data` is RGBA (changed in place).
 * Returns the share of pixels removed.
 */
export function removeBackground(data, w, h, { tolerance = 42 } = {}) {
  const seeds = [];
  // The edge colours, grouped (a background may be two-tone, a gradient or a vignette).
  const palette = [];
  const addColor = (i) => {
    if (data[i + 3] < 16) return;
    const c = [data[i], data[i + 1], data[i + 2]];
    if (!palette.some((p) => Math.hypot(p[0] - c[0], p[1] - c[1], p[2] - c[2]) < tolerance * 0.6)) palette.push(c);
  };
  for (let x = 0; x < w; x += 1) { seeds.push(x, (h - 1) * w + x); }
  for (let y = 0; y < h; y += 1) { seeds.push(y * w, y * w + w - 1); }
  for (let k = 0; k < seeds.length; k += Math.max(1, Math.floor(seeds.length / 400))) addColor(seeds[k] * 4);
  if (!palette.length) return 0;
  const near = (i) => Math.min(...palette.map((c) => dist(data, i, c)));
  const seen = new Uint8Array(w * h);
  const queue = new Int32Array(w * h);
  let head = 0;
  let tail = 0;
  for (const p of seeds) if (!seen[p] && near(p * 4) < tolerance) { seen[p] = 1; queue[tail++] = p; }
  let removed = 0;
  while (head < tail) {
    const p = queue[head++];
    const i = p * 4;
    const d = near(i);
    // Well inside the background: clear; near the subject: partly see-through.
    data[i + 3] = d < tolerance * 0.6 ? 0 : Math.round(data[i + 3] * Math.min(1, (d - tolerance * 0.6) / (tolerance * 0.4)));
    removed += 1;
    const x = p % w;
    const y = (p - x) / w;
    for (const q of [x > 0 ? p - 1 : -1, x < w - 1 ? p + 1 : -1, y > 0 ? p - w : -1, y < h - 1 ? p + w : -1]) {
      if (q < 0 || seen[q]) continue;
      seen[q] = 1;
      if (near(q * 4) < tolerance) queue[tail++] = q;
    }
  }
  return removed / (w * h);
}

/** 透明色を指定: every pixel of (about) this colour becomes transparent. Returns how many changed. */
export function makeTransparent(data, color, { tolerance = 28 } = {}) {
  let n = 0;
  for (let i = 0; i < data.length; i += 4) if (dist(data, i, color) <= tolerance && data[i + 3]) { data[i + 3] = 0; n += 1; }
  return n;
}

/** The size a picture is saved at when compressed: its longest side at most `max` pixels. */
export function compressedSize(w, h, max) {
  const k = Math.min(1, max / Math.max(w, h));
  return [Math.max(1, Math.round(w * k)), Math.max(1, Math.round(h * k))];
}

// ---------------------------------------------------------------- アート効果 and 色の変更 (on the pixels)

/** A box blur of one channel set, radius r, run along rows then columns (running sums: as fast for any radius). */
function boxBlur(src, w, h, r) {
  const out = new Float32Array(src.length);
  const tmp = new Float32Array(src.length);
  const pass = (from, to, len, count, step, stride) => {
    for (let line = 0; line < count; line += 1) {
      for (let ch = 0; ch < 4; ch += 1) {
        const at = (i) => (line * stride + Math.min(len - 1, Math.max(0, i)) * step) * 4 + ch;
        let sum = 0;
        for (let i = -r; i <= r; i += 1) sum += from[at(i)];
        for (let i = 0; i < len; i += 1) {
          to[(line * stride + i * step) * 4 + ch] = sum / (2 * r + 1);
          sum += from[at(i + r + 1)] - from[at(i - r)];
        }
      }
    }
  };
  pass(src, tmp, w, h, 1, w);
  pass(tmp, out, h, w, w, 1);
  return out;
}
const luma = (d, i) => 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
const clamp8 = (v) => Math.max(0, Math.min(255, Math.round(v)));

/** 図の形式 → アート効果: what each does to the picture. */
export const ART_EFFECTS = { blur: "ぼかし", pencil: "鉛筆：スケッチ", lines: "線画", mosaic: "モザイク", sepia: "セピア", poster: "ポスター" };

/** Apply an アート効果 to RGBA pixels in place. */
export function artEffect(data, w, h, kind) {
  if (kind === "blur") {
    const r = Math.max(2, Math.round(Math.min(w, h) / 90));
    const out = boxBlur(boxBlur(data, w, h, r), w, h, r);
    for (let i = 0; i < data.length; i += 4) { data[i] = out[i]; data[i + 1] = out[i + 1]; data[i + 2] = out[i + 2]; }
  } else if (kind === "pencil") {
    // A pencil sketch: the grey picture "colour-dodged" with a blurred negative of itself.
    const neg = new Float32Array(data.length);
    for (let i = 0; i < data.length; i += 4) { const g = 255 - luma(data, i); neg[i] = neg[i + 1] = neg[i + 2] = g; neg[i + 3] = 255; }
    const soft = boxBlur(neg, w, h, Math.max(2, Math.round(Math.min(w, h) / 120)));
    for (let i = 0; i < data.length; i += 4) {
      const g = luma(data, i);
      const v = clamp8(soft[i] >= 255 ? 255 : (g * 255) / (255 - soft[i]));
      data[i] = data[i + 1] = data[i + 2] = v;
    }
  } else if (kind === "lines") {
    // 線画: the edges (Sobel), dark on white.
    const g = new Float32Array(w * h);
    for (let p = 0; p < w * h; p += 1) g[p] = luma(data, p * 4);
    const at = (x, y) => g[Math.min(h - 1, Math.max(0, y)) * w + Math.min(w - 1, Math.max(0, x))];
    for (let y = 0; y < h; y += 1) {
      for (let x = 0; x < w; x += 1) {
        const gx = at(x + 1, y - 1) + 2 * at(x + 1, y) + at(x + 1, y + 1) - at(x - 1, y - 1) - 2 * at(x - 1, y) - at(x - 1, y + 1);
        const gy = at(x - 1, y + 1) + 2 * at(x, y + 1) + at(x + 1, y + 1) - at(x - 1, y - 1) - 2 * at(x, y - 1) - at(x + 1, y - 1);
        const v = clamp8(255 - Math.hypot(gx, gy) * 0.9);
        const i = (y * w + x) * 4;
        data[i] = data[i + 1] = data[i + 2] = v;
      }
    }
  } else if (kind === "mosaic") {
    const s = Math.max(6, Math.round(Math.min(w, h) / 40));
    for (let by = 0; by < h; by += s) {
      for (let bx = 0; bx < w; bx += s) {
        const sum = [0, 0, 0];
        let n = 0;
        for (let y = by; y < Math.min(h, by + s); y += 1) for (let x = bx; x < Math.min(w, bx + s); x += 1) { const i = (y * w + x) * 4; sum[0] += data[i]; sum[1] += data[i + 1]; sum[2] += data[i + 2]; n += 1; }
        for (let y = by; y < Math.min(h, by + s); y += 1) for (let x = bx; x < Math.min(w, bx + s); x += 1) { const i = (y * w + x) * 4; data[i] = sum[0] / n; data[i + 1] = sum[1] / n; data[i + 2] = sum[2] / n; }
      }
    }
  } else if (kind === "sepia") recolor(data, "sepia");
  else if (kind === "poster") {
    for (let i = 0; i < data.length; i += 4) for (let c = 0; c < 3; c += 1) data[i + c] = Math.round(Math.round((data[i + c] / 255) * 3) * 85);
  }
  return data;
}

/** 図の形式 → 色 → 色の変更 (PowerPoint's recolor): the plain ones and the theme's accents, dark and light. */
export const RECOLORS = { gray: "グレースケール", sepia: "セピア", wash: "ウォッシュアウト", accent1Dark: "アクセント 1（濃色）", accent1Light: "アクセント 1（淡色）", accent2Dark: "アクセント 2（濃色）", accent2Light: "アクセント 2（淡色）" };
const OFFICE = { accent1: "#4472c4", accent2: "#ed7d31" };
const rgb = (hex) => [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
/** The two ends a duotone runs between (dark pixels to the first, light ones to the second) in the theme's colours. */
function duotone(kind, colors = OFFICE) {
  const m = /^accent([12])(Dark|Light)$/.exec(kind);
  if (!m) return null;
  const accent = rgb(colors[`accent${m[1]}`] || OFFICE[`accent${m[1]}`]);
  const mix = (a, b, t) => a.map((v, i) => Math.round(v + (b[i] - v) * t));
  return m[2] === "Dark" ? [mix(accent, [0, 0, 0], 0.55), mix(accent, [255, 255, 255], 0.85)] : [accent, [255, 255, 255]];
}

/** Recolour RGBA pixels in place. */
export function recolor(data, kind, colors) {
  const tone = duotone(kind, colors);
  for (let i = 0; i < data.length; i += 4) {
    const r = data[i];
    const g = data[i + 1];
    const b = data[i + 2];
    const l = luma(data, i);
    if (kind === "gray") { data[i] = data[i + 1] = data[i + 2] = clamp8(l); }
    else if (kind === "sepia") { data[i] = clamp8(0.393 * r + 0.769 * g + 0.189 * b); data[i + 1] = clamp8(0.349 * r + 0.686 * g + 0.168 * b); data[i + 2] = clamp8(0.272 * r + 0.534 * g + 0.131 * b); }
    else if (kind === "wash") { for (let c = 0; c < 3; c += 1) data[i + c] = clamp8(255 - (255 - data[i + c]) * 0.3); }
    else if (tone) {
      const [dark, light] = tone;
      const t = l / 255;
      for (let c = 0; c < 3; c += 1) data[i + c] = clamp8(dark[c] + (light[c] - dark[c]) * t);
    }
  }
  return data;
}

export const RESOLUTIONS = [[2400, "高品質（2400 ピクセル）"], [1600, "印刷用（1600 ピクセル）"], [1280, "Web（1280 ピクセル）"], [960, "電子メール（960 ピクセル）"]];

export function createPictureTools(editor, app) {
  const { h, E } = app;
  const images = () => editor.selectedObjects().filter((o) => o.kind === "image");

  async function load(o) {
    const url = app.imageUrl(o.src);
    const img = new Image();
    img.crossOrigin = "anonymous";
    await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = () => reject(new Error("画像を読み込めません")); img.src = url; });
    return img;
  }
  function canvasOf(img, crop = null) {
    const W = img.naturalWidth;
    const H = img.naturalHeight;
    const c = crop || { l: 0, t: 0, r: 0, b: 0 };
    const sx = Math.round(W * c.l);
    const sy = Math.round(H * c.t);
    const sw = Math.max(1, Math.round(W * (1 - c.l - c.r)));
    const sh = Math.max(1, Math.round(H * (1 - c.t - c.b)));
    const canvas = document.createElement("canvas");
    canvas.width = sw;
    canvas.height = sh;
    const g = canvas.getContext("2d", { willReadFrequently: true });
    g.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh);
    return { canvas, g };
  }
  const toBlob = (canvas, type = "image/png", q = 0.9) => new Promise((resolve, reject) => canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("画像にできません"))), type, q));

  /** 背景の削除 on the selected picture (its own pixels only: a crop stays a crop). */
  async function background() {
    const [o] = images();
    if (!o) { app.toast("図を選んでください"); return; }
    try {
      const img = await load(o);
      const { canvas, g } = canvasOf(img);
      const px = g.getImageData(0, 0, canvas.width, canvas.height);
      const share = removeBackground(px.data, canvas.width, canvas.height);
      if (share < 0.01) { app.toast("背景を見つけられませんでした（背景が一色に近い写真・ロゴで使えます）"); return; }
      g.putImageData(px, 0, 0);
      const src = await app.storeBlob(await toBlob(canvas), `${o.fileName || "図"}（背景なし）.png`);
      editor.apply((x) => (x.id === o.id ? { src } : null));
      app.toast(`背景を削除しました（${Math.round(share * 100)}%。⌘Zで戻せます）`);
    } catch (error) {
      app.toast(`背景を削除できませんでした（${error.message}。Webの画像はこのデバイスに保存してから使ってください）`);
    }
  }

  /** Change the selected pictures' pixels (a new picture for each; the old ones stay in the undo history). */
  async function pixels(work, label) {
    const list = images();
    if (!list.length) { app.toast("図を選んでください"); return; }
    try {
      const done = [];
      for (const o of list) {
        const img = await load(o);
        const { canvas, g } = canvasOf(img);
        const px = g.getImageData(0, 0, canvas.width, canvas.height);
        work(px.data, canvas.width, canvas.height);
        g.putImageData(px, 0, 0);
        done.push({ id: o.id, src: await app.storeBlob(await toBlob(canvas), `${o.fileName || "図"}（${label}）.png`) });
      }
      editor.apply((x) => { const d = done.find((c) => c.id === x.id); return d ? { src: d.src, gray: undefined } : null; }, { ids: done.map((d) => d.id) });
      app.toast(`${label}を適用しました（⌘Zで戻せます）`);
    } catch (error) {
      app.toast(`${label}を適用できませんでした（${error.message}。Webの画像はこのデバイスに保存してから使ってください）`);
    }
  }
  const art = (kind) => pixels((d, w, h) => artEffect(d, w, h, kind), `アート効果「${ART_EFFECTS[kind]}」`);
  const colorize = (kind) => pixels((d) => recolor(d, kind, E.themeColors(E.PALETTE.sheet?.themeId)), `色の変更「${RECOLORS[kind]}」`);

  /** 透明色を指定: click the colour on the picture. */
  function transparentColor() {
    const [o] = images();
    if (!o) { app.toast("図を選んでください"); return; }
    const wrap = editor.state.wrap;
    const layer = h("div", { class: "ed-pathdraw ed-pick-color", title: "透明にする色をクリック（Escでやめる）" });
    const stop = () => { layer.remove(); document.removeEventListener("keydown", key, true); };
    const key = (e) => { if (e.key === "Escape") { e.preventDefault(); e.stopPropagation(); stop(); } };
    document.addEventListener("keydown", key, true);
    layer.addEventListener("pointerdown", async (event) => {
      event.preventDefault();
      event.stopPropagation();
      const [x, y] = editor.toSlide(event);
      stop();
      if (x < o.x || y < o.y || x > o.x + o.w || y > o.y + o.h) { app.toast("図の上をクリックしてください"); return; }
      try {
        const img = await load(o);
        const crop = o.crop || { l: 0, t: 0, r: 0, b: 0 };
        const W = img.naturalWidth;
        const H = img.naturalHeight;
        const fx = crop.l + ((x - o.x) / o.w) * (1 - crop.l - crop.r);
        const fy = crop.t + ((y - o.y) / o.h) * (1 - crop.t - crop.b);
        const { canvas, g } = canvasOf(img);
        const px = g.getImageData(0, 0, W, H);
        const i = (Math.min(H - 1, Math.max(0, Math.floor(fy * H))) * W + Math.min(W - 1, Math.max(0, Math.floor(fx * W)))) * 4;
        const color = [px.data[i], px.data[i + 1], px.data[i + 2]];
        const n = makeTransparent(px.data, color);
        g.putImageData(px, 0, 0);
        const src = await app.storeBlob(await toBlob(canvas), `${o.fileName || "図"}（透明色）.png`);
        editor.apply((obj) => (obj.id === o.id ? { src } : null));
        app.toast(`この色を透明にしました（${Math.round((n / (W * H)) * 100)}%。⌘Zで戻せます）`);
      } catch (error) {
        app.toast(`透明にできませんでした（${error.message}）`);
      }
    });
    wrap.append(layer);
    app.toast("透明にする色を、図の上でクリックしてください（Escでやめる）");
  }

  /** 図の圧縮: smaller pictures (and the cropped-away parts gone), for this picture or all of them. */
  function compressDialog() {
    const res = h("select", { "aria-label": "解像度" }, RESOLUTIONS.map(([v, l]) => h("option", { value: String(v), selected: v === 1600 || null }, l)));
    const onlySel = h("input", { type: "checkbox", checked: images().length ? true : null, disabled: images().length ? null : true });
    const dropCrop = h("input", { type: "checkbox", checked: true });
    const dialog = h("dialog", { class: "compress-dialog", "aria-label": "画像の圧縮" },
      h("div", { class: "dialog-head" }, h("h3", {}, "画像の圧縮"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body sh-form-col" },
        h("label", { class: "sh-choice" }, onlySel, h("span", {}, "この画像だけに適用する")),
        h("label", { class: "sh-choice" }, dropCrop, h("span", {}, "図のトリミング部分を削除する")),
        h("label", { class: "sh-inline" }, "解像度 ", res),
        h("p", { class: "hint" }, "大きすぎる写真を小さくして、資料・HTML出力のファイルを軽くします（元に戻すには ⌘Z）。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary cp-ok", onclick: () => { dialog.close(); compress({ max: Number(res.value), all: !onlySel.checked, dropCrop: dropCrop.checked }); } }, "OK")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
  }
  async function compress({ max, all, dropCrop }) {
    const targets = all ? app.allImages() : images().map((o) => ({ slide: app.index(), o }));
    if (!targets.length) { app.toast("画像がありません"); return; }
    let before = 0;
    let after = 0;
    const changes = [];
    for (const { slide, o } of targets) {
      try {
        const url = app.imageUrl(o.src);
        const old = await (await fetch(url)).blob();
        const img = await load(o);
        const crop = dropCrop && o.crop ? o.crop : null;
        const { canvas } = canvasOf(img, crop);
        const [w, hh] = compressedSize(canvas.width, canvas.height, max);
        if (!crop && w === canvas.width && old.size < 400_000) continue;
        const out = document.createElement("canvas");
        out.width = w;
        out.height = hh;
        out.getContext("2d").drawImage(canvas, 0, 0, w, hh);
        const png = /png|gif|webp|svg/.test(old.type);
        const blob = await toBlob(out, png ? "image/png" : "image/jpeg", 0.86);
        if (blob.size >= old.size && !crop) continue;
        before += old.size;
        after += blob.size;
        changes.push({ slide, id: o.id, patch: { src: await app.storeBlob(blob, o.fileName || "図"), ...(crop ? { crop: undefined } : {}) } });
      } catch { /* a picture from the web that cannot be read stays as it is */ }
    }
    if (!changes.length) { app.toast("小さくできる画像はありませんでした"); return; }
    app.patchObjects(changes);
    app.toast(`${changes.length}枚の画像を圧縮しました（${Math.round(before / 1024)}KB → ${Math.round(after / 1024)}KB。⌘Zで戻せます）`);
  }

  return { background, transparentColor, compressDialog, compress, art, colorize };
}
