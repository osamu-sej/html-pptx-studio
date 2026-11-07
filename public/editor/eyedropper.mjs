// スポイト (PowerPoint's Eyedropper, in the colour menus): a colour taken from anywhere on the screen. Chrome and Edge
// lend their own eyedropper (window.EyeDropper); elsewhere the slide itself is sampled: a click on a picture reads
// its pixel, on a shape its fill (or line), on words their colour, else the colour behind.

/** A CSS colour as #rrggbb, or null when it is transparent or none. */
export function cssToHex(value) {
  const v = String(value || "").trim().toLowerCase();
  if (!v || v === "none" || v === "transparent") return null;
  let m = v.match(/^#([0-9a-f]{3})$/);
  if (m) return `#${[...m[1]].map((c) => c + c).join("")}`;
  m = v.match(/^#([0-9a-f]{6})([0-9a-f]{2})?$/);
  if (m) return m[2] === "00" ? null : `#${m[1]}`;
  m = v.match(/^rgba?\(\s*([\d.]+)[,\s]+([\d.]+)[,\s]+([\d.]+)(?:\s*[,/]\s*([\d.]+%?))?\s*\)$/);
  if (!m) return null;
  if (m[4] != null && parseFloat(m[4]) === 0) return null;
  return `#${[m[1], m[2], m[3]].map((n) => Math.max(0, Math.min(255, Math.round(Number(n)))).toString(16).padStart(2, "0")).join("")}`;
}

const SHAPES = new Set(["path", "rect", "circle", "ellipse", "polygon", "polyline", "line"]);

/** The colour on the screen at (x, y) inside `root` (the slide), read from what is drawn there. */
export function colorAt(x, y, root = document.body) {
  for (const el of document.elementsFromPoint(x, y)) {
    if (!root.contains(el) || el.closest?.(".ed-overlay, .ed-handle, .ed-eyedrop-swatch")) continue;
    const tag = el.tagName.toLowerCase();
    if (tag === "img" && el.naturalWidth) {
      try {
        const r = el.getBoundingClientRect();
        const c = document.createElement("canvas");
        c.width = 1;
        c.height = 1;
        const g = c.getContext("2d", { willReadFrequently: true });
        g.drawImage(el, ((x - r.left) / r.width) * el.naturalWidth, ((y - r.top) / r.height) * el.naturalHeight, 1, 1, 0, 0, 1, 1);
        const [red, green, blue, alpha] = g.getImageData(0, 0, 1, 1).data;
        if (alpha > 8) return `#${[red, green, blue].map((n) => n.toString(16).padStart(2, "0")).join("")}`;
      } catch { /* a picture from elsewhere cannot be read */ }
      continue;
    }
    const style = getComputedStyle(el);
    if (SHAPES.has(tag)) {
      const hex = cssToHex(style.fill) || cssToHex(style.stroke);
      if (hex) return hex;
      continue;
    }
    // Words: their colour, when the point is on them.
    const words = [...el.childNodes].some((n) => n.nodeType === 3 && n.textContent.trim());
    if (words) {
      const range = document.createRange();
      for (const n of el.childNodes) {
        if (n.nodeType !== 3 || !n.textContent.trim()) continue;
        range.selectNodeContents(n);
        if ([...range.getClientRects()].some((r) => x >= r.left && x <= r.right && y >= r.top && y <= r.bottom)) return cssToHex(style.color);
      }
    }
    const back = cssToHex(style.backgroundColor);
    if (back) return back;
  }
  return null;
}

/**
 * Picks a colour: the browser's eyedropper when there is one, else a click on the slide (`root`). Resolves with
 * #rrggbb, or null when given up (Esc).
 */
export function pickColor({ root, toast = () => {} } = {}) {
  if (typeof window.EyeDropper === "function") {
    return new window.EyeDropper().open().then((r) => cssToHex(r.sRGBHex)).catch(() => null);
  }
  return new Promise((resolve) => {
    const swatch = document.createElement("div");
    swatch.className = "ed-eyedrop-swatch";
    swatch.hidden = true;
    document.body.append(swatch);
    document.body.classList.add("ed-eyedropping");
    toast("スポイト：色を取りたい場所をクリックしてください（Escでやめる）");
    let last = null;
    const STOP = ["pointerdown", "mousedown", "pointerup", "mouseup", "dblclick", "contextmenu"];
    const finish = (color) => {
      document.removeEventListener("pointermove", move, true);
      for (const type of STOP) document.removeEventListener(type, swallow, true);
      document.removeEventListener("click", click, true);
      document.removeEventListener("keydown", key, true);
      document.body.classList.remove("ed-eyedropping");
      swatch.remove();
      resolve(color);
    };
    const move = (event) => {
      last = colorAt(event.clientX, event.clientY, root);
      swatch.hidden = !last;
      if (!last) return;
      swatch.style.left = `${event.clientX + 16}px`;
      swatch.style.top = `${event.clientY + 16}px`;
      swatch.style.setProperty("--c", last);
      swatch.textContent = last.toUpperCase();
    };
    // The whole click is the eyedropper's (the editor does not select what is under it).
    const swallow = (event) => { event.preventDefault(); event.stopPropagation(); };
    // A click off the slide gives up, as Esc does.
    const click = (event) => { swallow(event); finish(colorAt(event.clientX, event.clientY, root)); };
    const key = (event) => { if (event.key === "Escape") { swallow(event); finish(null); } };
    document.addEventListener("pointermove", move, true);
    for (const type of STOP) document.addEventListener(type, swallow, true);
    document.addEventListener("click", click, true);
    document.addEventListener("keydown", key, true);
  });
}
