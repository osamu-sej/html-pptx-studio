// PowerPoint出力 in a real browser: the sample deck goes to the server, which draws it in its headless browser and writes
// an editable .pptx (server/pptx.mjs, tools/pptx_export.py). The download is opened with python-pptx and checked: every
// slide, the theme as a PowerPoint theme, titles and key messages as text, layouts as shapes, a native chart, a picture
// for what PowerPoint has no part for, the builds as PowerPoint animations, every slide part valid against ECMA-376; a 白紙 page keeps its own objects; a check finding asks first and continues the
// PowerPoint export (not the HTML one).
// Usage: node qa/studio-pptx.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { execFileSync } from "node:child_process";
import { mkdir } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
process.env.PLAYWRIGHT_DISABLE_FORCED_CHROMIUM_PROXIED_LOOPBACK = "1";
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "/opt/node22/lib/node_modules/playwright");
const root = fileURLToPath(new URL("..", import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).map((arg) => arg.replace(/^--/, "").split("=")));
const base = args.base || "http://127.0.0.1:8787";
const python = process.env.PYTHON_BIN || "python3";
const outDir = join(root, "qa", "out");
await mkdir(outDir, { recursive: true });

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`); });
const shot = async (name) => { const file = join(outDir, `pptx-${name}.png`); await page.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`).catch(() => {}); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };

/** What python-pptx reads in a .pptx: slides with their shapes (kind, words), the theme's accent and fonts. */
function inspect(file) {
  const code = `
import json, sys
from pptx import Presentation
from lxml import etree
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
schema = etree.XMLSchema(etree.parse(sys.argv[2]))
prs = Presentation(sys.argv[1])
theme = etree.fromstring(prs.slide_master.part.part_related_by("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme").blob)
out = {"size": [prs.slide_width, prs.slide_height], "accent1": theme.find(f".//{A}accent1/{A}srgbClr").get("val"),
       "major": theme.find(f".//{A}majorFont/{A}ea").get("typeface"), "slides": []}
def flat(shapes):
    # A group's members (the studio's groups are PowerPoint groups).
    for sh in shapes:
        if sh.shape_type == 6:
            yield from flat(sh.shapes)
        else:
            yield sh
for s in prs.slides:
    shapes = []
    for sh in flat(s.shapes):
        kind = "chart" if sh.has_chart else "table" if sh.has_table else "picture" if sh.shape_type == 13 else "text" if sh.has_text_frame else str(sh.shape_type)
        shapes.append({"kind": kind, "text": sh.text_frame.text if sh.has_text_frame else ""})
    # The slide's animations and transition as PowerPoint reads them, and whether the part is valid (ECMA-376).
    el = etree.fromstring(etree.tostring(s._element))
    effects = [c.get("presetClass") for c in el.iter(f"{P}cTn") if c.get("presetClass")]
    videos = len(el.findall(f".//{A}videoFile"))
    tops = {c.get("id") for c in el.findall(f"{P}cSld/{P}spTree/*/*/{P}cNvPr")}
    inner = sorted({t.get("spid") for t in el.iter(f"{P}spTgt")} - tops)
    has_tr = el.find(f"{P}transition") is not None or el.find(f"{{{MC}}}AlternateContent") is not None
    for alt in list(el.iter(f"{{{MC}}}AlternateContent")):
        fb, parent = alt.find(f"{{{MC}}}Fallback"), alt.getparent()
        i = list(parent).index(alt)
        parent.remove(alt)
        for k, kid in enumerate(list(fb) if fb is not None else []):
            parent.insert(i + k, kid)
    valid = schema.validate(el)
    out["slides"].append({"shapes": shapes, "notes": s.notes_slide.notes_text_frame.text if s.has_notes_slide else "", "effects": effects,
                          "transition": has_tr, "valid": valid, "inner": inner, "videos": videos, "errors": [e.message[:160] for e in list(schema.error_log)[:2]]})
print(json.dumps(out, ensure_ascii=False))`;
  return JSON.parse(execFileSync(python, ["-c", code, file, join(root, "test", "ooxml-xsd", "pml.xsd")]).toString());
}

async function exportFile(name, click) {
  const [download] = await Promise.all([page.waitForEvent("download", { timeout: 180_000 }), click()]);
  const file = join(outDir, name);
  await download.saveAs(file);
  return file;
}

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);

let sample = null;
let allSample = null;
await step("the sample deck goes out as an editable PowerPoint (the top bar's PowerPoint出力)", async () => {
  await page.click("#sampleDeckBtn");
  await page.waitForFunction(() => document.querySelectorAll(".film-item").length >= 40, null, { timeout: 20000 });
  sample = await page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.deck())));
  allSample = structuredClone(sample);
  // Ten slides are enough to see every kind of part (and keep the run short): the first nine and the closing.
  await page.evaluate(() => { const d = window.__hpptx.deck(); d.slides.splice(9, d.slides.length - 10); });
  sample.slides.splice(9, sample.slides.length - 10);
  await page.click(".film-item:nth-child(1)");
  await page.waitForSelector("#pptxBtn:not([hidden])");
  const file = await exportFile("export-sample.pptx", async () => {
    await page.click("#pptxBtn");
    // A finding asks first; この­まま出力 goes on with the PowerPoint (not the HTML) export.
    const asked = await page.waitForSelector("#exportCheckDialog[open]", { timeout: 4000 }).then(() => true).catch(() => false);
    if (asked) await page.click("#exportCheckGoBtn");
  });
  await page.waitForFunction(() => /PowerPoint に書き出しました/.test(document.body.textContent), null, { timeout: 30000 });
  const got = inspect(file);
  assert(got.size[0] === 12192000 && got.size[1] === 6858000, `16:9: ${got.size}`);
  assert(got.slides.length === 10, `every slide: ${got.slides.length}`);
  const accent = await page.evaluate(() => window.SlideEngine.themeColors(window.__hpptx.deck().theme).accent1.slice(1).toUpperCase());
  assert(got.accent1 === accent, `the theme's accent as PowerPoint's アクセント 1: ${got.accent1} / ${accent}`);
  const words = (i) => got.slides[i].shapes.map((s) => s.text).join(" ");
  const titles = sample.slides.slice(0, 10).map((s) => (s.title || "").replace(/\*\*/g, ""));
  for (const [i, title] of titles.entries()) if (title && sample.slides[i].type !== "section") assert(words(i).replace(/\s/g, "").includes(title.replace(/\s/g, "")), `slide ${i + 1} has its title as text: ${title} / ${words(i).slice(0, 120)}`);
  const takeaway = (sample.slides[1].takeaway || "").replace(/\*\*/g, "");
  assert(words(1).replace(/\s/g, "").includes(takeaway.replace(/\s/g, "").slice(0, 10)), "the key message as text");
  const kinds = got.slides.flatMap((s) => s.shapes.map((x) => x.kind));
  assert(kinds.filter((k) => k === "text").length > 60, `the layouts as shapes and text boxes: ${kinds.filter((k) => k === "text").length}`);
  assert(kinds.includes("chart"), "a chart PowerPoint can edit");
  assert(got.slides.some((s, i) => sample.slides[i].notes && s.notes === sample.slides[i].notes), "speaker notes");
  // Every slide part is valid (PowerPoint opens it without repairing), the layouts' builds are PowerPoint animations
  // (entrances), and every slide has its transition.
  const invalid = got.slides.map((s, i) => [i + 1, s.errors]).filter(([i]) => !got.slides[i - 1].valid);
  assert(!invalid.length, `valid slide parts: ${JSON.stringify(invalid)}`);
  const animated = got.slides.filter((s) => s.effects.includes("entr")).length;
  assert(animated >= 5, `the layouts' builds as PowerPoint entrances: ${animated} slides`);
  assert(got.slides.every((s) => s.transition), "every slide's transition");
  // PowerPoint animates whole shapes only (never a member of a group).
  const inner = got.slides.map((s, i) => [i + 1, s.inner]).filter(([, list]) => list.length);
  assert(!inner.length, `animations on whole shapes: ${JSON.stringify(inner)}`);
  // What the layout itself moves: the cover's title (its タイトルの動き) is filmed and plays as the slide arrives.
  assert(got.slides[0].videos >= 1, `the cover's moving title filmed: ${got.slides[0].videos} videos`);
});

await step("a 白紙 page keeps its own objects (a shape with words, a table) and an icon becomes a picture", async () => {
  await page.evaluate(() => {
    const d = window.__hpptx.deck();
    d.slides.splice(1, 9);
    d.slides.push({ type: "blank", title: "", hideTitle: true, elements: [
      { id: "qs", kind: "shape", shape: "roundRect", x: 200, y: 200, w: 700, h: 300, text: "<p>手で置いた図形</p>", fs: 48 },
      { id: "qt", kind: "table", x: 1000, y: 200, w: 700, h: 300, cols: [0.5, 0.5], rows: [0.5, 0.5], cells: [[{ text: "<p>項目</p>" }, { text: "<p>値</p>" }], [{ text: "<p>A</p>" }, { text: "<p>1</p>" }]] },
      { id: "qi", kind: "icon", icon: "check", x: 200, y: 700, w: 160, h: 160 },
      { id: "qw", kind: "text", x: 1000, y: 700, w: 700, h: 100, text: "<p>タイプライターで現れる</p>", fs: 48 },
    ], timeline: [{ id: "qa1", el: "qw", cls: "in", fx: "typewriter", start: "click", dur: 1200 }] });
  });
  await page.evaluate(() => window.__hpptx.deck().slides.length);
  const file = await exportFile("export-blank.pptx", async () => {
    await page.click("#pptxBtn");
    const asked = await page.waitForSelector("#exportCheckDialog[open]", { timeout: 4000 }).then(() => true).catch(() => false);
    if (asked) await page.click("#exportCheckGoBtn");
  });
  const got = inspect(file);
  const last = got.slides.at(-1).shapes;
  assert(last.some((s) => s.kind === "text" && /手で置いた図形/.test(s.text)), `the shape's words: ${JSON.stringify(last)}`);
  assert(last.some((s) => s.kind === "table"), "the table");
  assert(last.some((s) => s.kind === "picture"), "the icon as a picture");
  // What only HTML draws (the typewriter) is filmed as a video that plays at its moment; the words stay as text.
  assert(got.slides.at(-1).videos >= 1, `the typewriter filmed: ${got.slides.at(-1).videos} videos`);
  assert(last.some((s) => s.kind === "text" && /タイプライターで現れる/.test(s.text)), "the typewriter's words stay as text");
  await shot("after-export");
});

await step("layout spin, swirl and flip are filmed while their objects remain editable", async () => {
  await page.evaluate((original) => {
    const d = window.__hpptx.deck();
    const sample = structuredClone(original.slides.find((s) => s.type === "agenda"));
    sample.items = ["動きの検証"];
    sample.build = "click";
    sample.backdrop = "none";
    sample.kinetic = "none";
    d.motion = { ...d.motion, backdrop: "none", kinetic: "none" };
    const title = { ...structuredClone(original.slides[0]), kinetic: "flip", backdrop: "none" };
    const statement = { ...structuredClone(original.slides.find((s) => s.type === "statement")), kinetic: "zoom", backdrop: "none" };
    const closing = { ...structuredClone(original.slides.at(-1)), kinetic: "none", backdrop: "none" };
    d.slides = [title, ...["spin", "swirl", "flip"].map((entrance, i) => ({ ...structuredClone(sample), title: `動き ${i + 1}`, entrance, ...(i === 1 ? { animation: "fade" } : {}) })), statement, closing];
  }, allSample);
  const file = await exportFile("export-layout-motion.pptx", async () => {
    await page.click("#pptxBtn");
    const asked = await page.waitForSelector("#exportCheckDialog[open]", { timeout: 4000 }).then(() => true).catch(() => false);
    if (asked) await page.click("#exportCheckGoBtn");
  });
  const got = inspect(file);
  assert(got.slides.length === 6, `slide count ${got.slides.length}`);
  for (const [i, slide] of got.slides.slice(1, 4).entries()) {
    assert(slide.videos >= 1, `slide ${i + 1}: HTML entrance was not filmed`);
    assert(slide.shapes.some((s) => s.kind === "text" && /動きの検証/.test(s.text)), `slide ${i + 1}: editable text missing`);
    assert(slide.valid, `slide ${i + 1}: invalid XML ${slide.errors}`);
  }
  assert(got.slides[0].videos >= 1 && got.slides[4].videos >= 1, "kinetic title and statement must be filmed");
  const readBack = `import json, sys
sys.path.insert(0, sys.argv[2])
from pptx_exact import read_pptx_exact
with open(sys.argv[1], 'rb') as f: deck = read_pptx_exact(f.read())
print(json.dumps([[e.get('fx') for e in s.get('timeline', [])] for s in deck['slideData']]))`;
  const effects = JSON.parse(execFileSync(python, ["-c", readBack, file, join(root, "tools")]).toString());
  assert(effects[1].includes("layoutSpin") && effects[2].includes("layoutSwirl") && effects[3].includes("layoutFlip"), `import restores motion: ${JSON.stringify(effects)}`);
  assert(effects[0].includes("kineticFlip") && effects[4].includes("kineticZoom"), `import restores kinetic motion: ${JSON.stringify(effects)}`);
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
