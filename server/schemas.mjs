import { readFileSync } from "node:fs";

import { z } from "zod";
import { zodToJsonSchema } from "zod-to-json-schema";

import { normalizeCodexOutputSchema } from "./codex-app-server.mjs";

export const studioSettingsSchema = z.object({
  slideCount: z.number().int().min(2).max(50).default(8),
  textDensity: z.enum(["light", "standard", "rich"]).default("standard"),
});

export const textItemSchema = z.string().min(1).max(180);
export const notesSchema = z.string().max(200000).optional();
// Built-in photo library (public/assets) — see PHOTOS in public/engine/engine.js.
export const visualAssetSchema = z.enum([
  "ai", "aiWorkflow", "promptDesign", "businessWorkshop", "businessEtiquette",
  "executiveDecision", "storeOperations", "dataInsight", "transformationRoadmap", "customerExperience", "none",
]);
export const customImageSchema = z.string().max(2_000_000).regex(/^data:image\/(?:png|jpeg|webp);base64,/, "画像はPNG・JPEG・WebPのみです");
export const placementSchema = z.object({
  x: z.number().min(0).max(1), y: z.number().min(0).max(1),
  w: z.number().min(0.05).max(1), h: z.number().min(0.05).max(1),
});
/** A photo, video or Lottie animation the user added (never written by the AI): a data URL, a file kept in the browser (idb:…), or a web address. */
export const mediaSchema = z.object({
  src: z.string().min(1).max(3_000_000),
  kind: z.enum(["image", "video", "lottie"]).optional(),
  name: z.string().max(200).optional(),
  fit: z.enum(["cover", "contain"]).optional(),
  autoplay: z.boolean().optional(),
  loop: z.boolean().optional(),
  muted: z.boolean().optional(),
  placement: placementSchema.optional(),
});
export const ICONS = Object.fromEntries(Object.entries(JSON.parse(readFileSync(new URL("../public/engine/icons.json", import.meta.url), "utf8"))).map(([key, value]) => [key, value.label]));
export const iconSchema = z.enum(Object.keys(ICONS));
// The kinds of motion mirror the catalogs in public/engine/engine.js (test/engine.test.js checks they agree).
// How the slide's items appear when presented: auto (by layout), none, fade (all at once), cascade (one after another),
// click (one per click), spotlight (all shown; each click brings one forward).
export const BUILDS = ["auto", "none", "fade", "cascade", "click", "spotlight"];
export const animationSchema = z.enum(BUILDS);
// Motion on the slide's photo while it is shown.
export const PHOTO_MOTIONS = ["zoom", "pan", "float", "parallax", "reveal", "drift", "tilt"];
export const photoMotionSchema = z.enum(["none", ...PHOTO_MOTIONS]);
// Motion graphics: how the big text sets itself in motion, and the moving graphic behind the slide.
// "auto" follows the deck (cover, chapters, statements, the close).
export const KINETIC_STYLES = ["mask", "words", "chars", "type", "scramble", "wave", "zoom", "flip", "slide"];
export const BACKDROP_KINDS = ["particles", "waves", "grid", "orbits", "gradient", "lines", "shapes", "confetti", "network", "ripple", "stars", "rays"];
export const kineticSchema = z.enum(["auto", "none", ...KINETIC_STYLES]);
export const backdropSchema = z.enum(["auto", "none", ...BACKDROP_KINDS]);
// How a slide's parts arrive, how items answer the mouse, how the **phrase** is set off, and how slides change.
export const ENTRANCES = ["rise", "fade", "blur", "pop", "slide", "zoom", "flip", "wipe", "drop", "spin", "swirl"];
export const HOVERS = ["lift", "focus", "tilt", "glow", "zoom"];
export const EMPHASES = ["marker", "underline", "circle", "box", "glow", "none"];
export const TRANSITIONS = ["fade", "slide", "zoom", "morph", "wipe", "circle", "push", "flip", "dive", "blinds", "curtain", "none"];
export const transitionSchema = z.enum(TRANSITIONS);
// Presentations only: text that opens when an item ("items[0]", "steps[2]") — or the key message ("takeaway") —
// is clicked. With a breakdown, a source or assumptions it opens as an evidence panel from the right.
export const detailSchema = z.object({
  target: z.string().min(1).max(30),
  title: z.string().max(60).optional(),
  text: z.string().min(1).max(400),
  rows: z.array(z.object({ label: z.string().min(1).max(40), value: z.string().max(30) })).max(8).optional(),
  source: z.string().max(120).optional(),
  note: z.string().max(200).optional(),
});

// Objects people place by hand (shapes, text boxes, pictures, lines…), their animations, and a slide's id that
// links between slides point at. People make them in the editor; the AI never writes them (USER_ONLY_FIELDS) and
// every AI change keeps them (server/chat.mjs). public/engine/objects.js validates their contents.
export const OBJECT_KINDS = ["shape", "text", "image", "line", "icon", "video", "audio", "lottie", "table", "chart", "smartart", "ink", "zoom", "camera", "equation"];
export const objectSchema = z.object({ id: z.string().min(1).max(32), kind: z.enum(OBJECT_KINDS) }).passthrough();
export const timelineEntrySchema = z.object({ el: z.string().min(1).max(64) }).passthrough();

const shared = {
  visualAsset: visualAssetSchema.optional(),
  customImage: customImageSchema.optional(),
  imagePlacement: placementSchema.optional(),
  media: mediaSchema.optional(),
  photoMotion: photoMotionSchema.optional(),
  animation: animationSchema.optional(),
  kinetic: kineticSchema.optional(),
  backdrop: backdropSchema.optional(),
  // This slide's own entrance, emphasis and way in (the deck's apply when left out).
  entrance: z.enum(["auto", "none", ...ENTRANCES]).optional(),
  emphasis: z.enum(["auto", ...EMPHASES]).optional(),
  transition: z.enum(["auto", ...TRANSITIONS]).optional(),
  // A deep-dive page: not part of the story, opened by clicking this item ("items[1]") of the slide above.
  drillOf: z.string().max(30).optional(),
  notes: notesSchema,
  // ノートの書式: the notes with their bold words and lists (kept only while it says what `notes` says).
  notesRich: z.string().max(200000).optional(),
  elements: z.array(objectSchema).max(5000).optional(),
  hidden: z.boolean().optional(),
  // PowerPoint's sections (the first slide of each carries its name) and review comments on a slide.
  section: z.string().max(40).optional(),
  // A thread: who wrote it (by, and the browser's uid), the object it is pinned to (anchor), and replies.
  comments: z.array(z.object({
    id: z.string().max(32), text: z.string().max(2000), at: z.string().max(40).optional(), done: z.boolean().optional(),
    by: z.string().max(40).optional(), uid: z.string().max(24).optional(), anchor: z.string().max(40).optional(), doneBy: z.string().max(40).optional(), edited: z.string().max(40).optional(),
    replies: z.array(z.object({ id: z.string().max(40), text: z.string().max(2000), at: z.string().max(40).optional(), by: z.string().max(40).optional(), uid: z.string().max(24).optional(), edited: z.string().max(40).optional() }).passthrough()).max(100).optional(),
  }).passthrough()).max(200).optional(),
  timeline: z.array(timelineEntrySchema).max(5000).optional(),
  sid: z.string().regex(/^[A-Za-z0-9_-]{1,32}$/).optional(),
  // 画面切り替え set by hand: the way in's length (ms) and moving on by itself after some seconds.
  transitionDur: z.number().min(100).max(10000).optional(),
  transitionSound: z.enum(["chime", "click", "camera", "whoosh", "drum", "applause", "coin", "bell", "stop"]).optional(),
  transitionDir: z.enum(["right", "left", "down", "up", "horizontal", "vertical"]).optional(),
  advance: z.number().min(0).max(600).optional(),
  // 読み取り順序 (校閲 → アクセシビリティ): the objects' ids in the order a screen reader reads them.
  readingOrder: z.array(z.string().max(32)).max(5000).optional(),
  // 背景グラフィックを表示しない: the deck's master objects (deck.masterObjects) stay off this slide.
  hideMaster: z.boolean().optional(),
  // 背景の書式設定: the slide's own background colour and picture (its transparency, tiled or not).
  background: z.object({ color: z.string().regex(/^#[0-9a-fA-F]{6}$/).optional(), image: z.string().max(80_000_000).optional(), transparency: z.number().min(0).max(1).optional(), tile: z.boolean().optional() }).optional(),
};
export const titledShape = {
  title: z.string().min(1).max(90),
  subhead: z.string().max(60).optional(),
  takeaway: z.string().max(160).optional(),
  // Where the slide's figures come from (shown at the foot of the page and in chart tooltips).
  source: z.string().max(120).optional(),
  details: z.array(detailSchema).max(12).optional(),
  ...shared,
};
export const cardItemSchema = z.object({ title: z.string().min(1).max(80), desc: z.string().max(220).optional(), icon: iconSchema.optional() });
export const chartDatumSchema = z.object({
  label: z.string().max(80),
  value: z.number().optional(),
  barValue: z.number().optional(),
  // "shift" charts: the value before (value is after).
  before: z.number().optional(),
});
export const chartSeriesSchema = z.object({
  id: z.string().max(60).optional(),
  label: z.string().max(80).optional(),
  values: z.array(z.number()).max(24),
});
// "rank": a ranking the audience re-sorts by switching views (切り口); "shift": bars move from before to after (差分).
export const CHART_TYPES = ["combo", "bar", "line", "donut", "multi-line", "stacked-bar", "100-stacked-bar", "rank", "shift"];
export const chartImageSchema = z.object({
  chartType: z.enum(CHART_TYPES),
  data: z.object({
    title: z.string().max(120).optional(),
    unit: z.string().max(10).optional(),
    highlight: z.string().max(80).optional(),
    beforeLabel: z.string().max(20).optional(),
    afterLabel: z.string().max(20).optional(),
    views: z.array(z.object({ label: z.string().min(1).max(20), items: z.array(chartDatumSchema).max(12) })).max(4).optional(),
    centerLabel: z.string().max(80).optional(),
    items: z.array(chartDatumSchema).max(24).optional(),
    xAxisLabels: z.array(z.string().max(80)).max(24).optional(),
    series: z.array(chartSeriesSchema).max(8).optional(),
    legendLabels: z.array(z.string().max(80)).max(8).optional(),
    barData: z.array(z.object({
      label: z.string().max(80),
      values: z.array(z.number()).max(8),
    })).max(24).optional(),
  }),
});

// The layouts the AI writes. People also have 白紙 (blank): the master and an optional title, filled with objects.
const AI_SLIDES = [
  z.object({ type: z.literal("title"), title: z.string().min(1).max(100), subtitle: z.string().max(180).optional(), date: z.string().max(32).optional(), ...shared }),
  z.object({ type: z.literal("section"), title: z.string().min(1).max(90), takeaway: z.string().max(160).optional(), sectionNo: z.number().int().min(1).max(99).optional(), ...shared }),
  z.object({ type: z.literal("closing"), title: z.string().max(90).optional(), message: z.string().max(300).optional(), ...shared }),
  z.object({ type: z.literal("hero"), ...titledShape }),
  z.object({ type: z.literal("statement"), ...titledShape, text: z.string().min(1).max(160) }),
  z.object({ type: z.literal("content"), ...titledShape, points: z.array(textItemSchema).max(10).default([]), twoColumn: z.boolean().optional(), columns: z.array(z.array(textItemSchema).max(7)).length(2).optional() }),
  z.object({ type: z.literal("agenda"), ...titledShape, items: z.array(textItemSchema).min(1).max(10) }),
  z.object({ type: z.literal("compare"), ...titledShape, leftTitle: z.string().max(60), rightTitle: z.string().max(60), leftItems: z.array(textItemSchema).max(8), rightItems: z.array(textItemSchema).max(8) }),
  z.object({ type: z.literal("process"), ...titledShape, steps: z.array(textItemSchema).min(2).max(6) }),
  z.object({ type: z.literal("processList"), ...titledShape, steps: z.array(textItemSchema).min(2).max(8) }),
  z.object({ type: z.literal("timeline"), ...titledShape, milestones: z.array(z.object({ label: z.string().min(1).max(80), date: z.string().max(40), state: z.enum(["done", "next", "todo"]).optional() })).min(2).max(8) }),
  z.object({ type: z.literal("diagram"), ...titledShape, lanes: z.array(z.object({ title: z.string().min(1).max(60), items: z.array(textItemSchema).max(6) })).min(2).max(5) }),
  z.object({ type: z.literal("cycle"), ...titledShape, items: z.array(z.object({ label: z.string().min(1).max(60), subLabel: z.string().max(100).optional() })).min(3).max(6), centerText: z.string().max(80).optional() }),
  z.object({ type: z.literal("cards"), ...titledShape, columns: z.union([z.literal(2), z.literal(3)]).optional(), items: z.array(z.union([textItemSchema, cardItemSchema])).min(1).max(6) }),
  z.object({ type: z.literal("headerCards"), ...titledShape, columns: z.union([z.literal(2), z.literal(3)]).optional(), items: z.array(cardItemSchema).min(1).max(6) }),
  z.object({ type: z.literal("table"), ...titledShape, headers: z.array(z.string().max(60)).min(2).max(6), rows: z.array(z.array(z.string().max(120)).max(6)).min(1).max(8) }),
  z.object({ type: z.literal("quote"), ...titledShape, text: z.string().min(1).max(420), author: z.string().max(100).optional() }),
  z.object({ type: z.literal("kpi"), ...titledShape, columns: z.union([z.literal(2), z.literal(3), z.literal(4)]).optional(), items: z.array(z.object({ label: z.string().min(1).max(60), value: z.string().min(1).max(40), change: z.string().max(60).optional(), status: z.enum(["good", "bad", "neutral"]).optional() })).min(1).max(4) }),
  z.object({ type: z.literal("bulletCards"), ...titledShape, items: z.array(cardItemSchema.extend({ desc: z.string().min(1).max(220) })).min(1).max(4) }),
  z.object({ type: z.literal("faq"), ...titledShape, items: z.array(z.object({ q: z.string().min(1).max(100), a: z.string().min(1).max(220) })).min(1).max(4) }),
  z.object({ type: z.literal("statsCompare"), ...titledShape, leftTitle: z.string().max(60), rightTitle: z.string().max(60), stats: z.array(z.object({ label: z.string().min(1).max(60), leftValue: z.string().max(50), rightValue: z.string().max(50), trend: z.enum(["up", "down", "neutral"]).optional() })).min(1).max(6) }),
  z.object({ type: z.literal("triangle"), ...titledShape, items: z.array(cardItemSchema).length(3) }),
  z.object({ type: z.literal("pyramid"), ...titledShape, levels: z.array(z.object({ title: z.string().min(1).max(60), description: z.string().max(160) })).min(3).max(5) }),
  z.object({ type: z.literal("flowChart"), ...titledShape, flows: z.array(z.object({ steps: z.array(textItemSchema).min(2).max(5) })).min(1).max(2) }),
  z.object({ type: z.literal("stepUp"), ...titledShape, items: z.array(cardItemSchema.extend({ desc: z.string().min(1).max(180) })).min(2).max(5) }),
  z.object({ type: z.literal("imageText"), ...titledShape, image: z.union([z.string().max(2_000_000), chartImageSchema]).optional(), imageCaption: z.string().max(120).optional(), imagePosition: z.enum(["left", "right"]).optional(), points: z.array(textItemSchema).max(8).default([]) }),
  z.object({ type: z.literal("grid2x2"), ...titledShape, items: z.array(cardItemSchema).min(1).max(4) }),
  z.object({ type: z.literal("headerTwoColumn"), ...titledShape, items: z.array(cardItemSchema).min(2).max(2) }),
  z.object({ type: z.literal("headerThreeSummary"), ...titledShape, items: z.array(cardItemSchema).min(3).max(3), summary: z.string().max(240).optional() }),
  z.object({ type: z.literal("funnel"), ...titledShape, levels: z.array(z.object({ title: z.string().min(1).max(60), description: z.string().max(160).optional() })).min(3).max(5) }),
  z.object({ type: z.literal("venn"), ...titledShape, items: z.array(cardItemSchema).min(2).max(3) }),
  z.object({ type: z.literal("gantt"), ...titledShape, periods: z.array(z.string().min(1).max(20)).max(8).optional(), items: z.array(cardItemSchema.extend({ start: z.number().int().min(0).max(7).optional(), span: z.number().int().min(1).max(8).optional() })).min(2).max(7), now: z.number().min(0).max(8).optional() }),
  z.object({ type: z.literal("orgChart"), ...titledShape, root: z.string().max(80).optional(), items: z.array(cardItemSchema).min(2).max(5) }),
  z.object({ type: z.literal("checklist"), ...titledShape, items: z.array(cardItemSchema.extend({ done: z.boolean().optional() })).min(1).max(7) }),
  z.object({ type: z.literal("matrix"), ...titledShape, xLabel: z.string().max(30).optional(), yLabel: z.string().max(30).optional(), items: z.array(cardItemSchema).length(4) }),
  z.object({ type: z.literal("beforeAfter"), ...titledShape, leftTitle: z.string().max(60), rightTitle: z.string().max(60), leftItems: z.array(textItemSchema).max(7), rightItems: z.array(textItemSchema).max(7) }),
  z.object({ type: z.literal("roadmap"), ...titledShape, items: z.array(cardItemSchema).min(2).max(6) }),
  z.object({ type: z.literal("dashboard"), ...titledShape, items: z.array(z.object({ label: z.string().min(1).max(60), value: z.string().min(1).max(40), change: z.string().max(60).optional() })).min(2).max(4), image: chartImageSchema.optional() }),
  z.object({ type: z.literal("swot"), ...titledShape, items: z.array(cardItemSchema).length(4) }),
  z.object({ type: z.literal("waterfall"), ...titledShape, unit: z.string().max(10).optional(), items: z.array(z.object({ label: z.string().min(1).max(40), value: z.number(), total: z.boolean().optional() })).min(2).max(8) }),
  z.object({ type: z.literal("logicTree"), ...titledShape, root: z.string().min(1).max(40), branches: z.array(z.object({ title: z.string().min(1).max(30), items: z.array(z.string().min(1).max(40)).max(3).optional(), highlight: z.boolean().optional() })).min(2).max(4) }),
  z.object({ type: z.literal("executiveSummary"), ...titledShape, conclusion: z.string().max(180).optional(), items: z.array(cardItemSchema).min(2).max(3), action: z.string().max(180).optional() }),
  // 因果: conditions (sliders, a/b/c in order) and a formula; the result is recalculated as they move.
  z.object({
    type: z.literal("simulator"), ...titledShape,
    inputs: z.array(z.object({ label: z.string().min(1).max(30), value: z.number(), min: z.number(), max: z.number(), step: z.number().min(0).optional(), unit: z.string().max(10).optional() })).min(1).max(3),
    formula: z.string().min(1).max(80), resultLabel: z.string().max(30).optional(), resultUnit: z.string().max(10).optional(), digits: z.number().int().min(0).max(2).optional(),
    compareLabel: z.string().max(20).optional(), compareValue: z.number().optional(),
  }),
  // 不足と打ち手: a target, where things stand, and measures that fill the gap as they are switched on.
  z.object({
    type: z.literal("gap"), ...titledShape, unit: z.string().max(10).optional(),
    targetLabel: z.string().max(20).optional(), target: z.number(), currentLabel: z.string().max(20).optional(), current: z.number(),
    measures: z.array(z.object({ title: z.string().min(1).max(30), value: z.number(), desc: z.string().max(60).optional() })).min(1).max(5),
  }),
];
export const aiSlideSchema = z.discriminatedUnion("type", AI_SLIDES);
// A 白紙 page: objects placed by hand. One brought over from PowerPoint as it looked keeps its title as an object
// (hideTitle) and may use the cover's master.
export const blankSlideSchema = z.object({ type: z.literal("blank"), ...titledShape, title: z.string().max(90).optional(), hideTitle: z.boolean().optional(), master: z.enum(["title", "content", "source"]).optional(), sourceViewport: z.object({ x: z.number(), y: z.number(), w: z.number().positive(), h: z.number().positive() }).optional() });
export const slideSchema = z.discriminatedUnion("type", [...AI_SLIDES, blankSlideSchema]);

export const THEMES = ["clarity", "midnight", "editorial", "mono", "forest", "sunset", "aurora", "kinari"];
export const themeSchema = z.enum(THEMES);
const motionFields = {
  entrance: z.enum([...ENTRANCES, "none"]).optional(),
  hover: z.enum([...HOVERS, "none"]).optional(),
  kinetic: z.enum(["none", ...KINETIC_STYLES]).optional(),
  backdrop: z.enum(["none", ...BACKDROP_KINDS]).optional(),
  emphasis: z.enum(EMPHASES).optional(),
  draw: z.boolean().optional(),
};
export const deckMotionSchema = z.object({
  ...motionFields,
  numbers: z.boolean().optional(),
  ambient: z.boolean().optional(),
});

export const defaultSettings = studioSettingsSchema.parse({});
export const generatedDeckSchema = z.object({
  deckTitle: z.string().min(1).max(100),
  purpose: z.string().max(180),
  audience: z.string().max(80),
  slideData: z.array(aiSlideSchema).min(2).max(50),
});
export const deckShape = z.object({
  title: z.string().min(1).max(100),
  purpose: z.string().max(180).optional().default(""),
  audience: z.string().max(80).optional().default(""),
  schemaVersion: z.string().max(20).optional().default("3.0"),
  theme: themeSchema.optional().default("clarity"),
  accent: z.string().regex(/^#[0-9a-fA-F]{6}$/).optional(),
  transition: transitionSchema.optional().default("fade"),
  motion: deckMotionSchema.optional(),
  // Standing instructions for every AI request on this deck ("役員向け", "数値は9月時点"…).
  memo: z.string().max(2000).optional().default(""),
  // The guides people drag on the slides (表示 → ガイド), in slide px.
  guides: z.object({ x: z.array(z.number()).max(20).optional(), y: z.array(z.number()).max(20).optional() }).optional(),
  // 既定の図形・テキスト ボックス・線に設定 (checked by the studio: objectDefaultsOf in public/editor/ops.mjs).
  objectDefaults: z.record(z.string(), z.record(z.string(), z.unknown())).optional(),
  // スライド マスター: the objects every slide shows (checked by the studio: normalizeObjects in public/engine/objects.js).
  masterObjects: z.array(z.record(z.string(), z.unknown())).max(500).optional(),
  // 配布資料マスター・ノート マスター (表示 → マスター表示; checked again by the studio: public/editor/printmaster.mjs).
  printMasters: z.object({ handout: z.record(z.string(), z.unknown()).optional(), notes: z.record(z.string(), z.unknown()).optional() }).optional(),
  // スライド番号の開始番号 (デザイン → スライドのサイズ): the number the first page carries (1 when not set).
  firstNumber: z.number().int().min(0).max(9999).optional(),
  // スライド ショーの設定 and 目的別スライド ショー (checked again by the studio: public/editor/show.mjs).
  show: z.record(z.string(), z.unknown()).optional(),
  customShows: z.array(z.object({ id: z.string().max(32), name: z.string().max(60), sids: z.array(z.string().max(32)).max(500) })).max(30).optional(),
  // ファイル → 情報 (author, subject, keywords…; checked by the studio: public/editor/fileinfo.mjs).
  info: z.record(z.string(), z.string().max(1000)).optional(),
  slides: z.array(slideSchema).min(1).max(500),
});
export const reviseRequestSchema = z.object({
  deck: deckShape,
  slideIndex: z.number().int().min(0).max(499),
  instruction: z.string().max(1000).optional().default(""),
  issues: z.array(z.object({ field: z.string().max(80), message: z.string().max(200) })).max(30).optional().default([]),
  mode: z.enum(["replace", "insert"]).optional().default("replace"),
});
export const rewriteRequestSchema = z.object({
  deck: deckShape,
  instruction: z.string().min(1).max(1000),
  settings: z.object({ textDensity: z.enum(["light", "standard", "rich"]).default("standard"), flexibleCount: z.boolean().default(false) }).optional().default({}),
});
export const notesRequestSchema = z.object({
  deck: deckShape,
  indices: z.array(z.number().int().min(0).max(499)).max(500).optional(),
  seconds: z.union([z.literal(30), z.literal(60), z.literal(90)]).optional().default(60),
  tone: z.enum(["丁寧", "簡潔", "親しみやすい"]).optional().default("丁寧"),
});
export const notesResultSchema = z.object({ notes: z.array(z.object({ slide: z.number().int().min(1).max(500), text: z.string().min(1).max(1200) })).max(500) });

// Conversational editing: the AI answers and proposes operations on the original slide numbers.
export const chatRequestSchema = z.object({
  deck: deckShape,
  message: z.string().min(1).max(4000),
  history: z.array(z.object({ role: z.enum(["user", "assistant"]), text: z.string().max(4000) })).max(20).optional().default([]),
  current: z.number().int().min(0).max(499).optional().default(0),
  focus: z.array(z.number().int().min(0).max(499)).max(500).optional().default([]),
  attachment: z.object({ name: z.string().max(200), text: z.string().max(60_000) }).optional(),
});
export const chatOperationSchema = z.object({
  op: z.enum(["replace", "insert", "delete"]),
  slide: z.number().int().min(1).max(500),
  content: aiSlideSchema.optional(),
});
export const chatResultSchema = z.object({
  reply: z.string().min(1).max(1500),
  operations: z.array(chatOperationSchema).max(50),
  order: z.array(z.number().int().min(1).max(500)).max(500).optional(),
  deckTitle: z.string().max(100).optional(),
  // No theme: the design is the user's choice (デザイン タブ), so the AI cannot switch it.
  transition: transitionSchema.optional(),
  // Deck-wide motion graphics, only when asked ("全体をもっと動かして", "表紙と章扉に動く背景を").
  motion: z.object(motionFields).optional(),
  suggestions: z.array(z.string().min(1).max(60)).max(3),
});

// Fields only people set (uploaded photos and videos, hand placement) are hidden from the AI's output schema.
const USER_ONLY_FIELDS = new Set(["customImage", "imagePlacement", "media", "elements", "timeline", "sid", "transitionDur", "transitionSound", "transitionDir", "advance", "hideTitle", "master", "sourceViewport", "hidden", "section", "comments", "readingOrder", "background", "hideMaster", "notesRich"]);
function withoutUserFields(node) {
  if (Array.isArray(node)) return node.map(withoutUserFields);
  if (!node || typeof node !== "object") return node;
  const out = Object.fromEntries(Object.entries(node).map(([key, value]) => [key, withoutUserFields(value)]));
  if (out.properties) {
    for (const field of USER_ONLY_FIELDS) delete out.properties[field];
    if (Array.isArray(out.required)) out.required = out.required.filter((key) => !USER_ONLY_FIELDS.has(key));
  }
  return out;
}
export const toCodexSchema = (schema) => withoutUserFields(normalizeCodexOutputSchema(zodToJsonSchema(schema, { target: "openAi", $refStrategy: "none" })));
export const codexDeckSchema = toCodexSchema(generatedDeckSchema);
export const codexSlideSchema = toCodexSchema(z.object({ slide: aiSlideSchema }));
export const codexNotesSchema = toCodexSchema(notesResultSchema);
export { translateRequestSchema, translateResultSchema } from "./translate.mjs";
export const codexChatSchema = toCodexSchema(chatResultSchema);

// The layouts the AI chooses from (44); ALL_SLIDE_TYPES adds 白紙.
export const SLIDE_TYPES = aiSlideSchema.options.map((option) => option.shape.type.value);
export const ALL_SLIDE_TYPES = slideSchema.options.map((option) => option.shape.type.value);

// Outline-first creation: a quick skeleton the user reshapes before the full deck is written.
export const outlineItemSchema = z.object({
  type: z.enum(SLIDE_TYPES),
  title: z.string().min(1).max(90),
  takeaway: z.string().max(160),
  content: z.string().max(400),
});
export const outlineResultSchema = z.object({
  deckTitle: z.string().min(1).max(100),
  reply: z.string().max(600),
  slides: z.array(outlineItemSchema).min(2).max(50),
});
export const codexOutlineSchema = toCodexSchema(outlineResultSchema);

// Three alternatives for one slide, to pick from side by side.
export const variantsRequestSchema = z.object({
  deck: deckShape,
  slideIndex: z.number().int().min(0).max(499),
  instruction: z.string().max(1000).optional().default(""),
});
export const variantsResultSchema = z.object({
  variants: z.array(z.object({ label: z.string().min(1).max(40), slide: aiSlideSchema })).min(2).max(3),
});
export const codexVariantsSchema = toCodexSchema(variantsResultSchema);
