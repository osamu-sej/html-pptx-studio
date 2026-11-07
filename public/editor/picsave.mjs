// 図として保存 (PowerPoint's Save as Picture, on the right-click menu): the chosen objects as a file of their own. One
// picture is drawn as it shows on the slide (its trimming, adjustments, shape and outline) into a PNG or JPEG; anything
// else (a shape with its text, a chart, an icon, a line, SmartArt, or several objects together) becomes an SVG that
// holds it as drawn — its styles written into it — or a PNG / JPEG made from that SVG, in a box just big enough for
// what shows (turned objects, arrowheads). 図として貼り付け uses the same drawing for copied objects (objectsPicture).

/** The box a w × h object turned by deg degrees needs (its middle stays in the middle). */
export function rotatedBox(w, h, deg = 0) {
  const a = (((Number(deg) || 0) % 360) * Math.PI) / 180;
  const [c, s] = [Math.abs(Math.cos(a)), Math.abs(Math.sin(a))];
  return { w: Math.round((w * c + h * s) * 100) / 100, h: Math.round((w * s + h * c) * 100) / 100 };
}

/** A file name from the object's name ("図 3.png"), without the characters files cannot have. */
export function pictureFileName(name, ext) {
  const base = String(name || "図").replace(/[\\/:*?"<>|\u0000-\u001f]/g, "_").trim().slice(0, 80) || "図";
  return `${base}.${ext}`;
}

// What an element's look depends on, written into the copy (SVG parts and the HTML of text boxes).
const SVG_PROPS = ["fill", "fill-opacity", "fill-rule", "stroke", "stroke-width", "stroke-opacity", "stroke-dasharray", "stroke-linecap", "stroke-linejoin", "opacity", "font-family", "font-size", "font-weight", "font-style", "text-anchor", "dominant-baseline", "letter-spacing", "visibility", "display"];
const HTML_PROPS = ["display", "position", "left", "top", "right", "bottom", "width", "height", "box-sizing", "margin", "padding", "color", "background-color", "font-family", "font-size", "font-weight", "font-style", "font-variant", "line-height", "letter-spacing", "text-align", "text-align-last", "text-indent", "text-transform", "text-decoration-line", "text-decoration-style", "text-decoration-color", "text-decoration-thickness", "white-space", "word-break", "overflow-wrap", "vertical-align", "list-style-type", "list-style-position", "writing-mode", "flex-direction", "justify-content", "align-items", "gap", "border-collapse", "border-top", "border-right", "border-bottom", "border-left", "border-radius", "column-count", "column-gap", "opacity", "transform", "transform-origin", "clip-path", "object-fit", "overflow", "visibility", "z-index"];

export function createPictureSaver(app) {
  const { h } = app;

  /** Copy the computed look of every element of `src` onto the same element of `copy` (both trees alike). */
  function inlineStyles(src, copy) {
    const walk = (a, b) => {
      if (a.nodeType !== 1) return;
      const cs = getComputedStyle(a);
      const props = a instanceof SVGElement ? SVG_PROPS : HTML_PROPS;
      const text = props.map((p) => { const v = cs.getPropertyValue(p); return v ? `${p}:${v}` : ""; }).filter(Boolean).join(";");
      if (text) b.setAttribute("style", text);
      b.removeAttribute("class");
      const [ac, bc] = [a.children, b.children];
      for (let i = 0; i < ac.length && i < bc.length; i += 1) walk(ac[i], bc[i]);
    };
    walk(src, copy);
  }

  /** blob: pictures inside the copy as data URLs (a picture made from an SVG cannot reach blob: addresses). */
  async function embedBlobs(root) {
    const toData = async (url) => {
      const blob = await (await fetch(url)).blob();
      return new Promise((resolve) => { const r = new FileReader(); r.onload = () => resolve(r.result); r.onerror = () => resolve(url); r.readAsDataURL(blob); });
    };
    for (const el of root.querySelectorAll("img[src^='blob:'], image")) {
      const attr = el.tagName.toLowerCase() === "img" ? "src" : el.hasAttribute("href") ? "href" : "xlink:href";
      const url = el.getAttribute(attr) || "";
      if (url.startsWith("blob:")) { try { el.setAttribute(attr, await toData(url)); } catch { /* left as it is */ } }
    }
  }

  /** Objects drawn on a slide (their .hs-obj nodes, in stacking order) as a standalone SVG document, each where it is
   *  on the slide. The box is the objects' own (`list`, exact), widened only where what they show sticks out more than
   *  a pixel (outlines, arrowheads, words) — measured on the page, which a zoomed-out stage rounds. */
  async function svgOfNodes(nodes, list = []) {
    const root = nodes[0].closest(".hs-slide");
    const base = root.getBoundingClientRect();
    const k = base.width / (root.offsetWidth || base.width) || 1;
    let box = null;
    for (const node of nodes) {
      for (const el of [node, ...node.querySelectorAll("*")]) {
        const r = el.getBoundingClientRect();
        if (r.width < 0.5 && r.height < 0.5) continue;
        const b = [(r.left - base.left) / k, (r.top - base.top) / k, (r.right - base.left) / k, (r.bottom - base.top) / k];
        box = box ? [Math.min(box[0], b[0]), Math.min(box[1], b[1]), Math.max(box[2], b[2]), Math.max(box[3], b[3])] : b;
      }
    }
    if (!box) throw new Error("nothing drawn");
    const own = nodes.map((n) => list.find((o) => o.id === n.dataset.el)).filter(Boolean).map((o) => app.E.bounds(o, list));
    if (own.length) {
      const exact = [Math.min(...own.map((b) => b.x)), Math.min(...own.map((b) => b.y)), Math.max(...own.map((b) => b.x + b.w)), Math.max(...own.map((b) => b.y + b.h))];
      box = [box[0] < exact[0] - 1 ? box[0] : exact[0], box[1] < exact[1] - 1 ? box[1] : exact[1], box[2] > exact[2] + 1 ? box[2] : exact[2], box[3] > exact[3] + 1 ? box[3] : exact[3]];
    }
    const [x, y] = [Math.floor(box[0] + 0.01), Math.floor(box[1] + 0.01)];
    const [w, hh] = [Math.max(1, Math.ceil(box[2] - 0.01) - x), Math.max(1, Math.ceil(box[3] - 0.01) - y)];
    const holder = document.createElement("div");
    holder.setAttribute("xmlns", "http://www.w3.org/1999/xhtml");
    holder.setAttribute("style", `position:relative;width:${w}px;height:${hh}px;overflow:visible`);
    for (const node of nodes) {
      const copy = node.cloneNode(true);
      inlineStyles(node, copy);
      copy.style.position = "absolute";
      copy.style.left = `${node.offsetLeft - x}px`;
      copy.style.top = `${node.offsetTop - y}px`;
      copy.style.margin = "0";
      for (const el of copy.querySelectorAll("[contenteditable]")) el.removeAttribute("contenteditable");
      await embedBlobs(copy);
      holder.append(copy);
    }
    const xml = new XMLSerializer().serializeToString(holder);
    return { text: `<?xml version="1.0" encoding="UTF-8"?>\n<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${hh}" viewBox="0 0 ${w} ${hh}"><foreignObject x="0" y="0" width="${w}" height="${hh}">${xml}</foreignObject></svg>`, x, y, w, h: hh };
  }

  /** 図として貼り付け: objects (copied, maybe from another slide) drawn off the screen into a PNG at twice the size,
   *  with where they sit on the slide. */
  async function objectsPicture(objects) {
    const probe = { type: "blank", title: "", hideTitle: true, elements: objects.map((o) => ({ ...o, hidden: undefined })) };
    const el = app.E.render(probe, { ...app.renderOptions(), mode: "thumb" });
    const host = document.createElement("div");
    host.setAttribute("aria-hidden", "true");
    host.style.cssText = "position:fixed;left:-30000px;top:0;width:1920px;height:1080px;overflow:hidden;pointer-events:none;";
    host.append(el);
    document.body.append(host);
    try {
      await document.fonts?.ready;
      await Promise.all([...el.querySelectorAll("img")].map((img) => img.decode?.().catch(() => {})));
      await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
      const ids = new Set(objects.map((o) => o.id));
      const nodes = [...el.querySelectorAll(".hs-obj[data-el]")].filter((n) => ids.has(n.dataset.el));
      if (!nodes.length) throw new Error("nothing drawn");
      const svg = await svgOfNodes(nodes, objects);
      return { blob: await rasterize(svg, "image/png"), x: svg.x, y: svg.y, w: svg.w, h: svg.h };
    } finally {
      host.remove();
    }
  }

  /** A canvas as a PNG or JPEG blob (JPEG on white). */
  const blobOf = (canvas, type) => new Promise((resolve, reject) => { try { canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("empty"))), type, 0.92); } catch (error) { reject(error); } });
  const loadImage = (url) => new Promise((resolve, reject) => { const img = new Image(); img.onload = () => resolve(img); img.onerror = () => reject(new Error("load")); img.src = url; });

  /** An SVG drawn into a PNG/JPEG at twice its size (crisper than the slide's pixels). */
  async function rasterize({ text, w, h: hh }, type) {
    const img = await loadImage(`data:image/svg+xml;charset=utf-8,${encodeURIComponent(text)}`);
    const k = Math.min(2, 8000 / Math.max(w, hh));
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(w * k));
    canvas.height = Math.max(1, Math.round(hh * k));
    const g = canvas.getContext("2d");
    if (type === "image/jpeg") { g.fillStyle = "#ffffff"; g.fillRect(0, 0, canvas.width, canvas.height); }
    g.drawImage(img, 0, 0, canvas.width, canvas.height);
    return blobOf(canvas, type);
  }

  /** A picture as it shows: its own pixels, trimmed, with its adjustments, cut to its shape, outlined, turned. */
  async function pictureBlob(o, node, type) {
    const shown = node.querySelector(".hs-obj-img img");
    if (!shown) throw new Error("no picture");
    const img = shown.complete && shown.naturalWidth ? shown : await loadImage(shown.src);
    const crop = o.crop || { l: 0, t: 0, r: 0, b: 0 };
    const sw = img.naturalWidth * (1 - crop.l - crop.r);
    const sh = img.naturalHeight * (1 - crop.t - crop.b);
    // The outline (図の枠線) as drawn on the slide; half of it lies outside the frame, so the picture is that much bigger.
    const outline = [...node.querySelectorAll(".hs-obj-geom path")].map((p) => ({ d: p.getAttribute("d"), color: p.getAttribute("stroke"), width: Number(p.getAttribute("stroke-width")) || 0, dash: String(p.getAttribute("stroke-dasharray") || "").split(/[\s,]+/).map(Number).filter((n) => n > 0) })).filter((l) => l.d && l.color && l.color !== "none" && l.width > 0);
    const m = outline.reduce((a, l) => Math.max(a, l.width / 2), 0);
    // As many pixels as the picture has across its frame (at most 4000).
    const k = Math.min(4000 / Math.max(o.w + 2 * m, o.h + 2 * m), Math.max(1, sw / o.w));
    const out = rotatedBox(o.w + 2 * m, o.h + 2 * m, o.rot || 0);
    const canvas = document.createElement("canvas");
    canvas.width = Math.max(1, Math.round(out.w * k));
    canvas.height = Math.max(1, Math.round(out.h * k));
    const g = canvas.getContext("2d");
    if (type === "image/jpeg") { g.fillStyle = "#ffffff"; g.fillRect(0, 0, canvas.width, canvas.height); }
    g.scale(k, k);
    g.translate(out.w / 2, out.h / 2);
    g.rotate(((o.rot || 0) * Math.PI) / 180);
    g.scale(o.flipH ? -1 : 1, o.flipV ? -1 : 1);
    g.translate(-o.w / 2, -o.h / 2);
    if (o.opacity != null) g.globalAlpha = o.opacity;
    g.save();
    if (o.mask) {
      const geo = app.E.geometry(o.mask, o.w, o.h, o.adj);
      g.clip(new Path2D(geo.paths.join(" ")), geo.rule === "evenodd" ? "evenodd" : "nonzero");
    }
    const filter = getComputedStyle(shown).filter;
    if (filter && filter !== "none") g.filter = filter;
    if (o.crop) g.drawImage(img, img.naturalWidth * crop.l, img.naturalHeight * crop.t, sw, sh, 0, 0, o.w, o.h);
    else if (o.fit === "cover" || o.fit === "contain") {
      const scale = (o.fit === "cover" ? Math.max : Math.min)(o.w / img.naturalWidth, o.h / img.naturalHeight);
      const [dw, dh] = [img.naturalWidth * scale, img.naturalHeight * scale];
      g.drawImage(img, (o.w - dw) / 2, (o.h - dh) / 2, dw, dh);
    } else g.drawImage(img, 0, 0, o.w, o.h);
    g.restore();
    for (const l of outline) {
      g.lineWidth = l.width;
      g.strokeStyle = l.color;
      g.setLineDash(l.dash);
      g.stroke(new Path2D(l.d));
    }
    return blobOf(canvas, type);
  }

  const stageNode = (id) => document.querySelector(`#stageBody .hs-obj[data-el="${CSS.escape(id)}"]`);
  /** The chosen objects as a file: one picture as its own pixels, anything else through an SVG. */
  async function write(chosen, format) {
    const nodes = chosen.map((o) => stageNode(o.id)).filter(Boolean);
    if (!nodes.length) { app.toast("スライド上にこのオブジェクトが見つかりません"); return; }
    const all = app.editorObjects();
    const name = chosen.length === 1 ? app.E.objectName(chosen[0], all.indexOf(chosen[0])) : "図";
    const ext = format === "jpeg" ? "jpg" : format;
    const type = format === "jpeg" ? "image/jpeg" : "image/png";
    const picture = chosen.length === 1 && chosen[0].kind === "image";
    try {
      if (picture) app.download(await pictureBlob(chosen[0], nodes[0], type), pictureFileName(name, ext));
      else {
        // In the order they lie on the slide (the front one last).
        nodes.sort((a, b) => (a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING ? -1 : 1));
        const svg = await svgOfNodes(nodes, all);
        if (format === "svg") app.download(new Blob([svg.text], { type: "image/svg+xml" }), pictureFileName(name, "svg"));
        else app.download(await rasterize(svg, type), pictureFileName(name, ext));
      }
      app.toast(chosen.length === 1 ? `「${name}」を図として保存しました` : `${chosen.length}個のオブジェクトを1つの図として保存しました`);
    } catch (error) {
      app.toast(picture ? "この画像は図として保存できません（ほかのサイトの画像はブラウザが書き出しを許しません）" : `図として保存できませんでした（${error.message || error}）。SVG ならたいてい保存できます`);
    }
  }

  /** The 図として保存 dialog: the file's kind, then the browser's download. */
  function open(ids) {
    const all = app.editorObjects();
    const chosen = (Array.isArray(ids) ? ids : [ids]).map((id) => all.find((x) => x.id === id)).filter(Boolean);
    if (!chosen.length) return;
    const picture = chosen.length === 1 && chosen[0].kind === "image";
    const choices = picture ? [["png", "PNG（透過を保つ）"], ["jpeg", "JPEG（軽い）"]] : [["png", "PNG（画像）"], ["svg", "SVG（拡大しても粗くならない）"], ["jpeg", "JPEG（軽い）"]];
    const format = h("select", { "aria-label": "ファイルの種類", class: "ps-format" }, choices.map(([v, l]) => h("option", { value: v }, l)));
    const dialog = h("dialog", { class: "ps-dialog", "aria-label": "図として保存" },
      h("div", { class: "dialog-head" }, h("h3", {}, "図として保存"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body sh-form-col" },
        h("label", { class: "sh-inline" }, "ファイルの種類 ", format),
        h("p", { class: "hint" }, picture ? "トリミング・明るさなどの修整・図形での切り抜き・枠線・回転を、スライドに見えるとおりに入れます。" : chosen.length > 1 ? `選んだ${chosen.length}個のオブジェクトを、スライドでの並びのまま1つの図にします。` : "スライドに見えるとおりの形・文字・色で保存します（回転もそのまま）。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary ps-go", onclick: () => { const f = format.value; dialog.close(); write(chosen, f); } }, "保存")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
  }

  return { open, write, objectsPicture };
}
