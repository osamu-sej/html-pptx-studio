// PowerPoint への書き出し（ブラウザ側）: runs in the server's headless browser (public/export.html, server/pptx.mjs). Each
// slide becomes PowerPoint-shaped objects: a layout the AI made is read back as 図形に変換 does (convert.mjs, the deck is
// not changed), a 白紙 page keeps its own objects with their fitted text sizes, and the deck's master objects lie under
// every slide. What PowerPoint has no part for (an icon, an equation, ink, a camera, a 3D model, a Lottie animation) is
// drawn alone on a clear slide for the server to photograph, and so is the theme's own background art.

import { createConverter, convertedSlide } from "./convert.mjs";

const E = window.SlideEngine;
// Keep these effects aligned with the layout's CSS entrances in engine.css. They are also in animate.js so an
// imported PowerPoint restores the editable animation instead of discarding an export-only effect name.
const LAYOUT_FILMS = {
  spin: { fx: "layoutSpin", dur: 1350 },
  swirl: { fx: "layoutSwirl", dur: 1200 },
  flip: { fx: "layoutFlip", dur: 700 },
};
const h = E.h;
let deck = null;
// The slide before (its objects as written), for 変形: PowerPoint pairs objects by name, the studio by morphPairs.
let previous = null;
const app = {
  E,
  h,
  deck: () => deck,
  renderOptions: () => ({ deck, assetBase: "/assets/" }),
  fitFor: () => undefined,
};
const converter = createConverter(app);
const stage = document.getElementById("stage");

// The slides' objects and animations, kept for recording their HTML-only effects (by slide index).
const forClips = new Map();
let clip = null;

/** An effect only the browser can draw: HTML's own (html: true), a 3D model's (model: true), or one with company. */
function htmlOnly(e) {
  if (e.with?.length) return true;
  const def = e.cls === "in" || e.cls === "out" ? E.ANIM_IN?.[e.fx] : e.cls === "em" ? E.ANIM_EM?.[e.fx] : null;
  return Boolean(def && (def.html || def.model));
}

// タイトルの動き (the layout's kinetic title) as an effect on its text box. HTML effects are filmed; remaining
// effects use PowerPoint's word-by-word or letter-by-letter animations.
const KINETIC_FX = {
  mask: { fx: "maskRise", dur: 900 }, type: { fx: "typewriter", dur: 1700 }, scramble: { fx: "decode", dur: 1300 },
  words: { fx: "floatIn", dir: "up", by: "word", dur: 600 }, chars: { fx: "rise", by: "char", dur: 500 }, wave: { fx: "floatIn", dir: "up", by: "char", dur: 500 },
  zoom: { fx: "kineticZoom", dur: 1500 }, flip: { fx: "kineticFlip", dur: 1600 }, slide: { fx: "flyIn", dir: "right", by: "word", dur: 500 },
};
let made = 0;
const madeId = () => `ar${(made += 1).toString(36)}`;

/**
 * What the layout itself moves as the slide arrives, as animations of the converted parts: the title's motion
 * (unless the slide's own animations move the title) and the figures that count up — each counts as its item
 * comes in (filmed with it), or as the slide arrives when nothing brings it in.
 */
function arrivalMotion(slide, timeline, { title, counts }, objects) {
  const out = [...timeline];
  const membersOf = (key) => objects.filter((o) => o.group === key.slice(4)).map((o) => o.id);
  const lead = [];
  const kinetic = E.kineticOf?.(slide, slide.type, deck.motion || {});
  const kind = kinetic && KINETIC_FX[kinetic];
  if (kind && title && !out.some((e) => e.el === title)) {
    const { fx, dur, ...rest } = kind;
    const entry = { id: madeId(), el: title, cls: "in", fx, start: "with", dur, delay: 120, ...rest };
    // Word and letter effects are PowerPoint's own on a text box (a group of pieces moves as one).
    if (title.startsWith("grp:")) delete entry.by;
    lead.push(entry);
  }
  // The layout header also plays the selected entrance when it has no kinetic title animation.
  const entrance = LAYOUT_FILMS[slide.entrance || deck.motion?.entrance];
  if (!kinetic && entrance && title && !out.some((e) => e.el === title)) {
    lead.push({ id: madeId(), el: title, cls: "in", fx: entrance.fx, start: "with", dur: entrance.dur, delay: 0 });
  }
  for (const id of counts) {
    const brings = out.find((e) => e.cls === "in" && (e.el === id || (e.el.startsWith("grp:") && membersOf(e.el).includes(id))));
    const count = { el: id, cls: "in", fx: "countUp", dur: 1100 };
    if (brings) brings.with = [...(brings.with || []), count];
    else lead.push({ id: madeId(), ...count, start: "with", delay: 0 });
  }
  return [...lead, ...out];
}

// Kinds PowerPoint draws itself; the rest are pictures (taken by the server).
const NATIVE = new Set(["shape", "text", "image", "line", "table", "chart", "video", "audio"]);
const clone = (v) => JSON.parse(JSON.stringify(v));
const wait = () => new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));

/** A picture of the library (asset:name) or the server (a path) as a data URL; data: and http(s) URLs stay. */
async function dataUrl(src) {
  if (!src || /^data:/.test(src)) return src;
  const url = E.resolveSrc(src, { assetBase: "/assets/" });
  if (!url) return null;
  try {
    const blob = await (await fetch(url)).blob();
    return await new Promise((resolve) => { const r = new FileReader(); r.onload = () => resolve(r.result); r.onerror = () => resolve(null); r.readAsDataURL(blob); });
  } catch {
    return null;
  }
}

/** Render a slide alone on the stage (full size, at the page's top left) and wait for its fonts and charts. */
async function mount(slide, index, { clear = false, mode = "print" } = {}) {
  stage.replaceChildren();
  stage.classList.toggle("is-clear", clear);
  const el = E.render(slide, { deck, index, mode, assetBase: "/assets/" });
  stage.append(el);
  await document.fonts?.ready;
  await wait();
  E.repaintCharts?.();
  await wait();
  return el;
}

/** Text sizes as the slide fits them ("shrink on overflow" objects get their own scale): written into the objects. */
function bakeFit(objects, scales) {
  for (const o of objects) {
    const k = scales?.[o.id];
    if (!k || k >= 0.999) continue;
    if (o.fs) o.fs = Math.round(o.fs * k * 100) / 100;
    if (typeof o.text === "string") o.text = o.text.replace(/font-size:\s*([\d.]+)px/g, (_, v) => `font-size: ${Math.round(Number(v) * k * 100) / 100}px`);
  }
}

/** The theme as PowerPoint wants it: colours by slot, fonts, and which colour is which theme colour. */
function themeModel() {
  const meta = E.THEMES.find((t) => t.id === deck.theme) || E.THEMES[0];
  const sheet = E.paletteSheet(meta.id);
  const refs = {};
  for (const col of sheet.columns) {
    if (!refs[col.hex]) refs[col.hex] = { slot: col.slot };
    for (const v of col.variants) if (!refs[v.hex]) refs[v.hex] = { slot: col.slot, mod: v.mod, off: v.off };
  }
  const c = E.themeColors(meta.id);
  return {
    id: meta.id, name: meta.name, dark: meta.dark,
    colors: { bg1: c.bg1, tx1: c.tx1, bg2: c.bg2, tx2: c.tx2, accents: [1, 2, 3, 4, 5, 6].map((i) => c[`accent${i}`]), hlink: c.hlink, folHlink: c.folHlink },
    fonts: { heading: meta.officeFonts?.heading, body: meta.officeFonts?.body },
    refs,
    marks: { accent: c.accent1, muted: c.tx2 },
  };
}

/** Each table's style colours (the same as the slide draws them, objects.js tableStyleVars). */
function tableLook(o) {
  const v = E.tableStyleVars(o.style || "medium2", deck.theme);
  return { a: v["--tb-a"], t20: v["--tb-t20"], t40: v["--tb-t40"], lt: v["--tb-lt"], dk: v["--tb-dk"], mid: v["--tb-mid"] };
}

/** A chart's series colours as drawn (the theme's accents, or its colour set). */
function seriesColors(o) {
  const vars = o.chart?.colors ? E.chartColorVars(o.chart.colors, deck.theme) : null;
  const c = E.themeColors(deck.theme);
  return [1, 2, 3, 4, 5, 6].map((i) => vars?.[`--c${i}`] || c[`accent${i}`]);
}

window.__pptx = {
  /** The deck to write (as the studio's HTML export carries it: media inside). */
  async load(source) {
    deck = clone(source);
    if (!E.THEMES.some((t) => t.id === deck.theme)) deck.theme = E.DEFAULT_THEME;
    E.useTheme(deck.theme);
    const meta = E.THEMES.find((t) => t.id === deck.theme);
    if (meta?.fonts?.length && !document.querySelector("link[data-theme-fonts]")) {
      const link = document.createElement("link");
      link.rel = "stylesheet";
      link.href = E.fontHref([deck.theme]);
      link.dataset.themeFonts = "";
      document.head.append(link);
      await new Promise((resolve) => { link.onload = resolve; link.onerror = resolve; setTimeout(resolve, 8000); });
    }
    await document.fonts?.ready;
    return { slides: deck.slides.length, theme: themeModel(), title: deck.title || "" };
  },

  /**
   * One slide as objects: { objects, timeline, background, notes, hidden, transition, rasters: [id], whole: bool }.
   * Objects that must be pictures are listed in `rasters` (the server photographs them with shoot()).
   */
  async slide(index) {
    const slide = deck.slides[index];
    let objects;
    let timeline = slide.timeline || [];
    let whole = false;
    let arrival = null;
    let layoutEntries = new Set();
    if (slide.type === "blank") {
      objects = clone(slide.elements || []);
      const el = await mount(slide, index);
      const fit = E.fit(el);
      bakeFit(objects, fit.objs);
      // The title (unless the page hides it) and the foot the studio draws on a 白紙 page; an animation on the
      // title ("@title") moves its text box.
      const chrome = await converter.chromeObjects(index);
      const fields = {};
      for (const o of chrome) if (o.field && !fields[o.field]) fields[o.field] = o.id;
      timeline = timeline.map((e) => (fields[e?.el] ? { ...e, el: fields[e.el] } : e));
      objects = [...chrome.map(({ field, ...o }) => o), ...objects];
    } else {
      const result = await converter.convert(index, { chrome: true, any: true });
      const made = convertedSlide(slide, result.objects, result.timeline);
      objects = clone(made.elements || []);
      timeline = made.timeline || [];
      const userEntries = new Set((slide.timeline || []).map((e) => e.id));
      layoutEntries = new Set(timeline.filter((e) => !userEntries.has(e.id)).map((e) => e.id));
      arrival = { title: result.fields?.["@kinetic"] || result.fields?.["@title"] || null, counts: result.counts || [] };
    }
    if (!objects.length && slide.type !== "blank") whole = true;
    if (deck.masterObjects?.length && !slide.hideMaster) objects = [...clone(deck.masterObjects).map((o) => ({ ...o, master: true })), ...objects];
    // The animations as the show plays them: known effects on parts that are on this slide, numbers in range.
    timeline = E.normalizeTimeline?.(timeline, { elements: objects }) ?? [];
    const entrance = slide.entrance || deck.motion?.entrance;
    const layoutFilm = LAYOUT_FILMS[entrance];
    if (layoutFilm) timeline = timeline.map((e) => layoutEntries.has(e.id) && e.cls === "in"
      ? { ...e, fx: layoutFilm.fx, dur: layoutFilm.dur } : e);
    if (arrival) timeline = arrivalMotion(slide, timeline, arrival, objects);
    const transition = slide.transition || deck.transition || "fade";
    // 変形: the parts the studio carries from the slide before (written with one name on both, "!!…").
    const morph = transition === "morph" && previous?.index === index - 1 && E.morphPairs
      ? E.morphPairs(previous.objects, objects).map(([a, b]) => [a.id, b.id]) : [];
    previous = { index, objects: clone(objects) };
    // What only HTML can draw (a typewriter, a number counting, a shine…): each such effect is recorded as a video
    // of its part alone (recordClip), played in PowerPoint at the same moment.
    const clips = timeline.filter((e) => htmlOnly(e)).map((e) => e.id);
    if (clips.length) forClips.set(index, { objects: clone(objects), timeline: clone(timeline) });
    const rasters = [];
    for (const o of objects) {
      if (!NATIVE.has(o.kind)) rasters.push(o.id);
      if (o.kind === "image" && o.src) o.src = await dataUrl(o.src);
      if ((o.kind === "video" || o.kind === "audio") && o.src) o.src = await dataUrl(o.src);
      if (o.kind === "table") o.look = tableLook(o);
      if (o.kind === "chart") o.seriesColors = seriesColors(o);
      if (o.fillImg) o.fillImg = await dataUrl(o.fillImg);
    }
    const background = slide.background ? { ...slide.background, ...(slide.background.image ? { image: await dataUrl(slide.background.image) } : {}) } : null;
    return {
      objects, timeline, background, whole, rasters, morph, clips,
      // 背景の動き: filmed as a looping video behind the slide (recordLoop).
      backdrop: (slide.type === "blank" && !slide.backdrop) || slide.master === "source" ? null : E.backdropOf?.(slide, slide.type, deck.motion || {}) || null,
      notes: slide.notes || "", hidden: Boolean(slide.hidden),
      transition, transitionDir: slide.transitionDir || null, transitionDur: Number(slide.transitionDur) || null,
      transitionSound: slide.transitionSound || null, advance: slide.advance ?? null,
    };
  },

  /** The studio's sounds by name as WAV files (they are synthesised in the show; PowerPoint plays files). */
  async sounds(names) {
    const out = {};
    for (const name of new Set(names)) {
      const wav = await E.renderSound?.(name).catch(() => null);
      if (wav) out[name] = wav;
    }
    return out;
  },

  /** One object drawn alone on a clear slide: its box on the page (the server takes the picture). */
  async mountObject(index, o) {
    const alone = { type: "blank", hideTitle: true, hideMaster: true, elements: [{ ...o, master: undefined }] };
    const el = await mount(alone, index, { clear: true });
    E.mountModels?.(el);
    E.mountLottie?.(el, { play: false, frame: 0.5 });
    await wait();
    const node = el.querySelector(`[data-el="${CSS.escape(o.id)}"]`);
    const r = (node || el).getBoundingClientRect();
    return { x: r.left, y: r.top, w: r.width, h: r.height };
  },

  /**
   * What lies under a slide's content (its ground, the theme's art, a full-bleed photo), alone: whether there is anything
   * to take besides the theme's own 背景 1 (which the PowerPoint's master already has).
   */
  async mountBackdrop(index) {
    const slide = deck.slides[index];
    const el = await mount({ ...slide, elements: [], timeline: undefined }, index);
    el.classList.add("hs-only-decor");
    await wait();
    const CONTENT = new Set(["hs-frame", "hs-overlay", "hs-objects", "hs-foot", "hs-master-layer", "hs-vlabel"]);
    const decor = el.querySelector(".hs-decor");
    const decorArt = decor && (decor.children.length || ["::before", "::after"].some((p) => { const cs = getComputedStyle(decor, p); return cs.content !== "none" && cs.content !== "normal"; })
      || getComputedStyle(decor).backgroundImage !== "none");
    const layers = [...el.children].some((child) => ![...child.classList].some((c) => CONTENT.has(c) || c === "hs-decor" || c === "hs-bg") && getComputedStyle(child).display !== "none");
    const parse = (c) => (c.match(/\d+(\.\d+)?/g) || []).slice(0, 3).map(Number);
    const ground = parse(getComputedStyle(el).backgroundColor);
    const bg1 = E.themeColors(deck.theme).bg1;
    const own = [1, 3, 5].map((i) => parseInt(bg1.slice(i, i + 2), 16));
    const otherGround = ground.length === 3 && ground.some((v, i) => Math.abs(v - own[i]) > 2);
    return { art: Boolean(decorArt || layers || otherGround || el.querySelector(".hs-bg")) };
  },

  /**
   * One HTML-only effect, ready to be filmed: the slide's ground (`background`, as the server took it), the parts
   * under the one that moves and the part itself, drawn as the show draws them, with the effect started and paused
   * at its beginning. Returns { dur, box, lasting, loop } — how long it plays (repeats included) and the region of
   * the page it moves in (sampled over its length), or null when there is nothing to film.
   */
  async mountClip(index, id, background) {
    const saved = forClips.get(index);
    const entry = saved?.timeline.find((e) => e.id === id);
    if (!entry) return null;
    const members = (o) => (entry.el.startsWith("grp:") ? o.group === entry.el.slice(4) : o.id === entry.el);
    const last = saved.objects.reduce((at, o, i) => (members(o) ? i : at), -1);
    if (last < 0) return null;
    // Only what lies under the part (and the part): what lies above it stays above the video in PowerPoint.
    const elements = saved.objects.slice(0, last + 1).filter((o, i) => members(o) || i < saved.objects.findIndex(members));
    const indefinite = entry.repeat === "click" || entry.repeat === "slide";
    const one = { ...entry, start: "click", delay: 0, trigger: undefined, after: undefined, sound: undefined, with: undefined, ...(indefinite ? { repeat: undefined } : {}) };
    // What moves with it (a figure counting up as its item comes in) is filmed in the same video.
    const company = (entry.with || []).map((w, k) => ({ ...w, id: `${entry.id}w${k}`, start: "with", delay: 0 }));
    const slide = { type: "blank", title: "", hideTitle: true, hideMaster: true, elements, timeline: [one, ...company], ...(background ? { background } : {}) };
    stage.replaceChildren();
    stage.classList.remove("is-clear");
    const el = E.render(slide, { deck, index, mode: "present", assetBase: "/assets/" });
    stage.append(el);
    await document.fonts?.ready;
    await wait();
    E.mountModels?.(el);
    E.repaintCharts?.();
    E.animStart?.(el, 0);
    const length = E.animStep?.(el, 1) || Number(entry.dur) || 500;
    // Every animation the effect started, kept from the start: one that does not hold its end (fill "none") leaves
    // the page's list once it is seeked past its end, and must still be found when seeking back.
    const kept = new Set();
    const anims = () => { for (const a of el.getAnimations ? el.getAnimations({ subtree: true }) : document.getAnimations()) kept.add(a); return [...kept]; };
    for (const a of anims()) a.pause();
    const nodes = () => [...el.querySelectorAll(".hs-obj")].filter((n) => elements.some((o) => members(o) && o.id === n.dataset.el))
      .flatMap((n) => [n, ...n.querySelectorAll(".hs-u")]);
    clip = { el, anims, length };
    // The region it moves in: the part's box at a few moments, a little room around it, inside the slide.
    let box = null;
    for (let k = 0; k <= 8; k += 1) {
      await this.seekClip((length * k) / 8);
      for (const n of nodes()) {
        const r = n.getBoundingClientRect();
        if (!(r.width > 0 && r.height > 0)) continue;
        box = box ? { x1: Math.min(box.x1, r.left), y1: Math.min(box.y1, r.top), x2: Math.max(box.x2, r.right), y2: Math.max(box.y2, r.bottom) } : { x1: r.left, y1: r.top, x2: r.right, y2: r.bottom };
      }
    }
    await this.seekClip(0);
    if (!box) return null;
    const room = Math.max(12, Math.min(60, 0.08 * Math.max(box.x2 - box.x1, box.y2 - box.y1)));
    const x = Math.max(0, Math.floor(box.x1 - room));
    const y = Math.max(0, Math.floor(box.y1 - room));
    const w = Math.min(1920, Math.ceil(box.x2 + room)) - x;
    const h = Math.min(1080, Math.ceil(box.y2 + room)) - y;
    const def = entry.cls === "em" ? E.ANIM_EM?.[entry.fx] : E.ANIM_IN?.[entry.fx];
    return { dur: Math.max(1, Math.round(length)), box: { x, y, w: w - (w % 2), h: h - (h % 2) }, lasting: Boolean(entry.cls === "em" && def?.lasting), loop: indefinite };
  },

  /**
   * 背景の動き alone (the ground, the theme's art and the moving backdrop, without the content), ready to be filmed as
   * a loop: its CSS animations held where they are and moved on by seekLoop.
   */
  async mountLoop(index) {
    const slide = deck.slides[index];
    // Drawn as the show draws it (a static page has no moving backdrop), the content hidden.
    const el = await mount({ ...slide, elements: [], timeline: undefined }, index, { mode: "present" });
    el.classList.add("hs-only-decor");
    await wait();
    const list = (el.getAnimations ? el.getAnimations({ subtree: true }) : document.getAnimations()).map((a) => { a.pause(); return { a, at: Number(a.currentTime) || 0 }; });
    clip = { el, anims: () => [], loop: list };
    return { animations: list.length };
  },

  /** The filmed backdrop `t` ms on from where it was mounted. */
  async seekLoop(t) {
    for (const { a, at } of clip?.loop || []) { try { a.currentTime = at + t; } catch { /* gone */ } }
    await wait();
  },

  /** The filmed effect at `t` ms (its frames drawn for that moment). */
  async seekClip(t) {
    if (!clip) return;
    for (const a of clip.anims()) { try { if (a.playState !== "paused") a.pause(); a.currentTime = Math.min(t, Number(a.effect?.getComputedTiming?.().endTime) || t); } catch { /* gone */ } }
    await wait();
  },

  /** The whole slide as the show draws it (a page the reader could not take apart). */
  async mountWhole(index) {
    await mount(deck.slides[index], index);
    return true;
  },
};
window.__pptxReady = true;
