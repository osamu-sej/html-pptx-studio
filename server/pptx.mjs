// PowerPoint への書き出し（サーバー側）: the deck is drawn in a headless browser (public/export.html with
// editor/pptx-collect.mjs) and read back as PowerPoint-shaped objects; what PowerPoint has no part for is photographed
// (an icon, an equation, ink, a camera, a 3D model, a Lottie animation, the theme's background art, a page that cannot
// be taken apart), and tools/pptx_export.py writes the .pptx. One browser serves every export, one export at a time.
import { spawn, spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { pythonBinary } from "./extract.mjs";

const require = createRequire(import.meta.url);

// PowerPoint's built-in table styles for the studio's (objects.js TABLE_STYLES); cells are also written out in full, so
// the slide looks the same whichever style PowerPoint knows.
export const TABLE_STYLE_IDS = {
  medium2: "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}",
  medium2a2: "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}",
  light1: "{3B4B98B0-60AC-42C2-AFA5-B58CD77FA1E5}",
  light2: "{69012ECD-51FC-41F1-AA8D-1B2483CD663E}",
  grid: "{5940675A-B579-460E-94D1-54222C63F5DA}",
  lines: "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
  plain: "{2D5ABB26-0587-4C30-8999-92F81FD0307C}",
};

/** Playwright (playwright-core, or the full package where a machine has it). */
function playwright() {
  for (const name of [process.env.PLAYWRIGHT_MODULE, "playwright-core", "playwright"].filter(Boolean)) {
    try { return require(name); } catch { /* next */ }
  }
  throw new Error("PowerPoint の書き出しに必要なブラウザ（Playwright）がありません。");
}

/** A Chromium to drive: CHROME_PATH, the system's, or Playwright's own. */
function chromePath() {
  for (const path of [process.env.CHROME_PATH, "/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome", "/opt/pw-browsers/chromium"].filter(Boolean)) {
    if (existsSync(path)) return path;
  }
  return undefined;
}

let browserPromise = null;
async function browser() {
  if (!browserPromise) {
    browserPromise = playwright().chromium.launch({ executablePath: chromePath(), args: ["--no-sandbox", "--disable-dev-shm-usage", "--font-render-hinting=none"] });
    browserPromise.then((b) => b.on("disconnected", () => { browserPromise = null; })).catch(() => { browserPromise = null; });
  }
  return browserPromise;
}

let queue = Promise.resolve();
/** One export at a time (each takes the browser for a while). */
function serially(task) {
  const run = queue.then(task, task);
  queue = run.catch(() => {});
  return run;
}

const dataUrl = (buffer) => `data:image/png;base64,${buffer.toString("base64")}`;

// ---------------------------------------------------------------- what only HTML can draw, as video

const FPS = 30;
// A deck's worth of filming at most (each clip is a few dozen pictures).
const MAX_CLIPS = 80;
const MAX_CLIP_MS = 20_000;

let ffmpegChecked = null;
/** ffmpeg with an H.264 encoder (PowerPoint plays H.264 MP4 everywhere), or null. */
export function ffmpegBinary() {
  if (ffmpegChecked !== null) return ffmpegChecked || null;
  const bin = process.env.FFMPEG_PATH || "ffmpeg";
  const probe = spawnSync(bin, ["-hide_banner", "-encoders"], { encoding: "utf8" });
  ffmpegChecked = probe.status === 0 && /libx264/.test(probe.stdout) ? bin : false;
  return ffmpegChecked || null;
}

/** Pictures (PNG buffers, one per frame at FPS) as an H.264 MP4 that PowerPoint plays. */
async function encodeMp4(frames) {
  const bin = ffmpegBinary();
  const dir = await mkdtemp(join(tmpdir(), "hpptx-clip-"));
  try {
    for (const [i, frame] of frames.entries()) await writeFile(join(dir, `f${String(i).padStart(5, "0")}.png`), frame);
    const out = join(dir, "clip.mp4");
    await new Promise((resolve, reject) => {
      const child = spawn(bin, ["-hide_banner", "-loglevel", "error", "-y", "-framerate", String(FPS), "-i", join(dir, "f%05d.png"),
        "-c:v", "libx264", "-preset", "veryfast", "-crf", "16", "-pix_fmt", "yuv420p", "-vf", "pad=ceil(iw/2)*2:ceil(ih/2)*2",
        "-movflags", "+faststart", "-an", out], { stdio: ["ignore", "ignore", "pipe"] });
      let err = "";
      child.stderr.on("data", (chunk) => { err += String(chunk); });
      child.once("error", reject);
      child.once("exit", (code) => (code === 0 ? resolve() : reject(new Error(err.trim() || `ffmpeg ${code}`))));
    });
    return await readFile(out);
  } finally {
    await rm(dir, { recursive: true, force: true }).catch(() => {});
  }
}

/**
 * Film one HTML-only effect of slide `i` (pptx-collect's mountClip / seekClip): its region, frame by frame, as an MP4.
 * Returns { src, poster, box, dur, lasting, loop } for the entry, or null.
 */
async function recordClip(page, i, id, background) {
  const info = await page.evaluate(([index, entryId, bg]) => window.__pptx.mountClip(index, entryId, bg), [i, id, background]);
  if (!info || !(info.box.w > 0 && info.box.h > 0)) return null;
  const dur = Math.min(info.dur, MAX_CLIP_MS);
  const count = Math.max(2, Math.ceil((dur / 1000) * FPS) + 1);
  const clip = { x: info.box.x, y: info.box.y, width: info.box.w, height: info.box.h };
  const frames = [];
  for (let k = 0; k < count; k += 1) {
    await page.evaluate((t) => window.__pptx.seekClip(t), Math.min(dur, (k * 1000) / FPS));
    frames.push(await page.screenshot({ type: "png", clip, scale: "css" }));
  }
  const mp4 = await encodeMp4(frames);
  return { src: `data:video/mp4;base64,${mp4.toString("base64")}`, poster: dataUrl(frames[0]), box: info.box, dur, lasting: info.lasting, loop: info.loop };
}

/**
 * The PowerPoint model of a deck: { title, theme, slides: [{ objects, background, notes, hidden, transition, … }] }.
 * `base` is this server's own address (the headless browser loads the export page from it).
 */
export async function collectModel(deck, { base, timeoutMs = 240_000 }) {
  const b = await browser();
  const context = await b.newContext({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: 2 });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  const started = Date.now();
  try {
    await page.goto(`${base}/export.html`, { waitUntil: "load" });
    await page.waitForFunction(() => window.__pptxReady === true, null, { timeout: 30_000 });
    const info = await page.evaluate((d) => window.__pptx.load(d), deck);
    const slides = [];
    let filmed = 0;
    for (let i = 0; i < info.slides; i += 1) {
      if (Date.now() - started > timeoutMs) throw new Error("書き出しに時間がかかりすぎました。スライドを分けて試してください。");
      const s = await page.evaluate((index) => window.__pptx.slide(index), i);
      // The theme's art (and a picture background) under the slide, as one picture.
      const backdrop = await page.evaluate((index) => window.__pptx.mountBackdrop(index), i);
      if (backdrop.art || s.background?.image) s.background = { image: dataUrl(await page.screenshot({ type: "png" })) };
      if (s.whole) {
        await page.evaluate((index) => window.__pptx.mountWhole(index), i);
        s.objects = [{ id: "whole", kind: "picture", name: "スライド全体", x: 0, y: 0, w: 1920, h: 1080, src: dataUrl(await page.screenshot({ type: "png" })) }];
      }
      // Objects PowerPoint has no part for: their picture, where they are drawn (rotation and effects included).
      const shots = new Set(s.rasters);
      for (const [k, o] of s.objects.entries()) {
        if (!shots.has(o.id) && o.kind !== "video" && o.kind !== "audio") continue;
        const r = await page.evaluate(([index, obj]) => window.__pptx.mountObject(index, obj), [i, o]);
        if (!(r.w > 0 && r.h > 0)) continue;
        const clip = { x: Math.max(0, r.x), y: Math.max(0, r.y), width: Math.min(1920, r.x + r.w) - Math.max(0, r.x), height: Math.min(1080, r.y + r.h) - Math.max(0, r.y) };
        if (clip.width <= 0 || clip.height <= 0) continue;
        const png = dataUrl(await page.screenshot({ type: "png", clip, omitBackground: true }));
        // A video's first picture, a sound's icon: PowerPoint shows them as the studio draws them.
        if (o.kind === "video" || o.kind === "audio") { o.poster = png; continue; }
        s.objects[k] = { id: o.id, kind: "picture", from: o.kind, name: o.name, alt: o.alt, x: clip.x, y: clip.y, w: clip.width, h: clip.height, src: png };
      }
      // What only HTML draws: each effect filmed on its own, for PowerPoint to play at the same moment.
      if (s.clips?.length && ffmpegBinary()) {
        for (const id of s.clips) {
          if (filmed >= MAX_CLIPS || Date.now() - started > timeoutMs) break;
          const entry = s.timeline.find((e) => e.id === id);
          const got = entry ? await recordClip(page, i, id, s.background).catch((error) => { errors.push(`clip: ${error.message}`); return null; }) : null;
          if (got) { entry.clip = got; filmed += 1; }
        }
      }
      delete s.clips;
      slides.push(s);
    }
    // The sounds the slides play (画面切り替えのサウンド, an effect's サウンド), made once for the whole deck.
    const names = slides.flatMap((s) => [s.transitionSound, ...(s.timeline || []).map((e) => e.sound)]).filter((n) => n && n !== "stop");
    const sounds = names.length ? await page.evaluate((list) => window.__pptx.sounds(list), names) : {};
    if (errors.length) console.warn(`pptx export: ${errors.slice(0, 3).join(" | ")}`);
    return { title: info.title, theme: info.theme, tableStyles: TABLE_STYLE_IDS, slides, sounds };
  } finally {
    await context.close().catch(() => {});
  }
}

/** tools/pptx_export.py: the model in, the .pptx bytes out. */
export function writePptx(model, { rootDir, timeoutMs = 180_000 }) {
  return new Promise((resolve, reject) => {
    const child = spawn(pythonBinary(rootDir), [join(rootDir, "tools", "pptx_export.py")], { cwd: rootDir, stdio: ["pipe", "pipe", "pipe"] });
    const out = [];
    let stderr = "";
    const timer = setTimeout(() => { child.kill("SIGKILL"); reject(new Error("PowerPoint の書き出しがタイムアウトしました。")); }, timeoutMs);
    child.stdout.on("data", (chunk) => out.push(chunk));
    child.stderr.on("data", (chunk) => { stderr += String(chunk); });
    child.once("error", (error) => { clearTimeout(timer); reject(error); });
    child.once("exit", (code) => {
      clearTimeout(timer);
      if (code !== 0) return reject(new Error((stderr.trim().split("\n").at(-1) || "書き出せませんでした").replace(/^\w*Error:\s*/, "")));
      let stats = {};
      try { stats = JSON.parse(stderr.trim().split("\n").at(-1) || "{}"); } catch { /* none */ }
      resolve({ bytes: Buffer.concat(out), stats });
    });
    child.stdin.end(Buffer.from(JSON.stringify(model)));
  });
}

/** A deck (as the HTML export carries it: media inside) → { bytes, stats } of its .pptx. */
export function exportPptx(deck, { base, rootDir }) {
  return serially(async () => writePptx(await collectModel(deck, { base }), { rootDir }));
}
