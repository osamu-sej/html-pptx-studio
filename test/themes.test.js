// Themes are PowerPoint themes (engine.js THEMES, objects.js): their colour sheet is Office's, objects are painted in
// theme colours, a change of theme recolours them, and a slide is always drawn in its own deck's theme.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  return window.SlideEngine;
}
const tones = (E, hex) => [...E.lumVariants(hex)].map(([mod, off]) => E.lumColor(hex, mod, off));

test("a deck without a theme (or with one the engine does not know) is drawn in the default theme", async () => {
  const E = await loadEngine();
  assert.equal(E.DEFAULT_THEME, "clarity");
  for (const theme of [undefined, "sej", "nope"]) {
    const slide = { type: "content", title: "本文", points: ["A"] };
    assert.equal(E.render(slide, { deck: { theme, slides: [slide] }, index: 0, mode: "thumb" }).dataset.theme, "clarity", String(theme));
  }
});

test("the colour sheet is PowerPoint's: each theme colour over five tones made the way Office makes them", async () => {
  const E = await loadEngine();
  // Office's own default theme, as PowerPoint shows it.
  assert.deepEqual(tones(E, "#4472c4"), ["#dae3f3", "#b4c7e7", "#8faadc", "#2f5597", "#203864"], "a middle colour: 80/60/40 % lighter, 25/50 % darker");
  assert.deepEqual(tones(E, "#ffffff"), ["#f2f2f2", "#d9d9d9", "#bfbfbf", "#a6a6a6", "#7f7f7f"], "white: 5–50 % darker");
  assert.deepEqual(tones(E, "#000000"), ["#7f7f7f", "#595959", "#404040", "#262626", "#0d0d0d"], "black: 50–5 % lighter");
  assert.deepEqual(tones(E, "#e7e6e6"), ["#d0cece", "#afabab", "#767171", "#3b3838", "#181717"], "a light colour: 10–90 % darker");
  assert.deepEqual(tones(E, "#ed7d31"), ["#fbe5d6", "#f8cbad", "#f4b183", "#c55a11", "#843c0b"]);
  for (const theme of E.THEMES) {
    const sheet = E.paletteSheet(theme.id);
    assert.deepEqual([...sheet.columns].map((c) => c.slot), ["bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"], theme.id);
    assert.ok(sheet.columns.every((c) => c.variants.length === 5 && c.variants.every((v) => /^#[0-9a-f]{6}$/.test(v.hex))), theme.id);
    assert.equal(sheet.standard.length, 10, "標準の色");
    assert.equal(sheet.columns[4].hex, theme.colors.accents[0].toLowerCase(), `${theme.id}: アクセント 1 is the theme's first accent`);
  }
});

test("a colour of the sheet is known as a theme colour (slot and lumMod/lumOff); any other colour is not", async () => {
  const E = await loadEngine();
  const c = E.themeColors("clarity");
  assert.deepEqual({ ...E.themeColorRef(c.accent1, "clarity") }, { slot: "accent1" });
  const light = E.paletteSheet("clarity").columns[4].variants[0];
  assert.deepEqual({ ...E.themeColorRef(light.hex, "clarity") }, { slot: "accent1", mod: 20000, off: 80000 });
  assert.deepEqual({ ...E.themeColorRef(E.paletteSheet("clarity").columns[4].variants[4].hex, "clarity") }, { slot: "accent1", mod: 50000, off: 0 });
  assert.equal(E.themeColorRef("#123456", "clarity"), null);
  assert.equal(E.themeColorRef("nonsense", "clarity"), null);
});

test("a change of theme recolours what was painted in theme colours, as PowerPoint does, and leaves other colours", async () => {
  const E = await loadEngine();
  const from = E.paletteSheet("clarity").columns;
  const to = E.paletteSheet("forest").columns;
  const deck = {
    theme: "clarity",
    masterObjects: [{ id: "m", kind: "shape", fill: from[4].hex }],
    slides: [{
      type: "blank",
      background: { color: from[2].hex },
      timeline: [{ el: "a", cls: "em", fx: "fillColor", color: from[5].hex }],
      elements: [
        { id: "a", kind: "shape", fill: from[4].hex, color: from[0].hex, stroke: "#123456", text: `<p><span style="color: ${from[4].variants[1].hex}">x</span></p>` },
      ],
    }],
  };
  E.retheme(deck, "clarity", "forest");
  const [o] = deck.slides[0].elements;
  assert.equal(o.fill, to[4].hex, "アクセント 1 → the new アクセント 1");
  assert.equal(o.color, to[0].hex, "背景 1 → the new 背景 1");
  assert.equal(o.stroke, "#123456", "a colour of its own stays");
  assert.match(o.text, new RegExp(to[4].variants[1].hex), "a tone inside the words → the same tone");
  assert.equal(deck.slides[0].background.color, to[2].hex);
  assert.equal(deck.slides[0].timeline[0].color, to[5].hex);
  assert.equal(deck.masterObjects[0].fill, to[4].hex);
});

test("objects take PowerPoint's defaults in the deck's theme, and drawing another deck does not change the editor's", async () => {
  const E = await loadEngine();
  E.useTheme("clarity");
  const before = E.PALETTE.sheet.themeId;
  const shape = { id: "s", kind: "shape", shape: "rect", x: 0, y: 0, w: 200, h: 100, text: "<p>文字</p>" };
  const slide = { type: "blank", elements: [shape] };
  const el = E.render(slide, { deck: { theme: "sunset", slides: [slide] }, index: 0, mode: "thumb" });
  assert.equal(el.dataset.theme, "sunset");
  assert.equal(el.querySelector(".hs-obj").dataset.fill, E.themeColors("sunset").accent1, "a new shape is filled with the deck's アクセント 1");
  assert.equal(E.PALETTE.sheet.themeId, before, "the editor's colours stay its own deck's");
  assert.equal(E.OBJECT_DEFAULTS.shape.fill, E.themeColors("clarity").accent1);
  assert.equal(E.OBJECT_DEFAULTS.text.color, E.themeColors("clarity").tx1, "a text box writes in テキスト 1");
});

test("table styles are PowerPoint's, in the theme's accent and its tints", async () => {
  const E = await loadEngine();
  assert.ok(E.TABLE_STYLES.medium2 && E.TABLE_STYLES.light1 && E.TABLE_STYLES.grid && E.TABLE_STYLES.plain);
  const vars = E.tableStyleVars("medium2", "clarity");
  const accent = E.themeColors("clarity").accent1;
  assert.equal(vars["--tb-a"], accent);
  assert.equal(vars["--tb-t20"], E.lumColor(accent, 20000, 80000), "the body: 80 % lighter");
  assert.equal(vars["--tb-t40"], E.lumColor(accent, 40000, 60000), "the bands: 60 % lighter");
  assert.equal(E.tableStyleVars("medium2a2", "clarity")["--tb-a"], E.themeColors("clarity").accent2);
  const table = E.normalizeObject({ id: "t", kind: "table", x: 0, y: 0, w: 600, h: 200, cols: [1, 1], rows: [1, 1], cells: [[{ text: "<p>a</p>" }, { text: "<p>b</p>" }], [{ text: "<p>c</p>" }, { text: "<p>d</p>" }]] });
  assert.equal(table.style, undefined, "the default style is not stored");
  const slide = { type: "blank", elements: [table] };
  const drawn = E.render(slide, { deck: { theme: "clarity", slides: [slide] }, index: 0, mode: "thumb" }).querySelector(".hs-otable");
  assert.ok(drawn.classList.contains("ts-medium2"));
  assert.match(drawn.getAttribute("style"), new RegExp(`--tb-a: ?${accent}`));
});
