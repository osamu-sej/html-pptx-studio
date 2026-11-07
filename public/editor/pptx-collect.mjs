// PowerPoint への書き出し（ブラウザ側）: runs in the server's headless browser (public/export.html, server/pptx.mjs). Each
// slide becomes PowerPoint-shaped objects: a layout the AI made is read back as 図形に変換 does (convert.mjs, the deck is
// not changed), a 白紙 page keeps its own objects with their fitted text sizes, and the deck's master objects lie under
// every slide. What PowerPoint has no part for (an icon, an equation, ink, a camera, a 3D model, a Lottie animation) is
// drawn alone on a clear slide for the server to photograph, and so is the theme's own background art.

import { createConverter, convertedSlide } from "./convert.mjs";

const E = window.SlideEngine;
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
async function mount(slide, index, { clear = false } = {}) {
  stage.replaceChildren();
  stage.classList.toggle("is-clear", clear);
  const el = E.render(slide, { deck, index, mode: "print", assetBase: "/assets/" });
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
    }
    if (!objects.length && slide.type !== "blank") whole = true;
    if (deck.masterObjects?.length && !slide.hideMaster) objects = [...clone(deck.masterObjects).map((o) => ({ ...o, master: true })), ...objects];
    // The animations as the show plays them: known effects on parts that are on this slide, numbers in range.
    timeline = E.normalizeTimeline?.(timeline, { elements: objects }) ?? [];
    const transition = slide.transition || deck.transition || "fade";
    // 変形: the parts the studio carries from the slide before (written with one name on both, "!!…").
    const morph = transition === "morph" && previous?.index === index - 1 && E.morphPairs
      ? E.morphPairs(previous.objects, objects).map(([a, b]) => [a.id, b.id]) : [];
    previous = { index, objects: clone(objects) };
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
      objects, timeline, background, whole, rasters, morph,
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

  /** The whole slide as the show draws it (a page the reader could not take apart). */
  async mountWhole(index) {
    await mount(deck.slides[index], index);
    return true;
  },
};
window.__pptxReady = true;
