// PowerPoint への書き出し（サーバー側）: the deck is drawn in a headless browser (public/export.html with
// editor/pptx-collect.mjs) and read back as PowerPoint-shaped objects; what PowerPoint has no part for is photographed
// (an icon, an equation, ink, a camera, a 3D model, a Lottie animation, the theme's background art, a page that cannot
// be taken apart), and tools/pptx_export.py writes the .pptx. One browser serves every export, one export at a time.
import { spawn, spawnSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { createRequire } from "node:module";
import { availableParallelism, tmpdir, totalmem } from "node:os";
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

const readNumber = (path) => { try { const v = readFileSync(path, "utf8").trim().split(/\s+/); return v; } catch { return null; } };

/**
 * What this machine may spend: CPUs and memory from the container's limits (cgroup v2, then v1), else the OS's.
 * A small Render instance has 0.15 CPU and 512 MB, which the OS figures do not show.
 */
export function machine() {
  let cpus = availableParallelism();
  const max = readNumber("/sys/fs/cgroup/cpu.max");
  if (max && max[0] !== "max") cpus = Math.min(cpus, Number(max[0]) / Number(max[1] || 100000));
  const quota = readNumber("/sys/fs/cgroup/cpu/cpu.cfs_quota_us");
  const period = readNumber("/sys/fs/cgroup/cpu/cpu.cfs_period_us");
  if (quota && Number(quota[0]) > 0 && period) cpus = Math.min(cpus, Number(quota[0]) / Number(period[0]));
  let memory = totalmem();
  for (const path of ["/sys/fs/cgroup/memory.max", "/sys/fs/cgroup/memory/memory.limit_in_bytes"]) {
    const v = readNumber(path);
    if (v && /^\d+$/.test(v[0]) && Number(v[0]) < memory) memory = Number(v[0]);
  }
  return { cpus, memory };
}

/**
 * A small machine (under 1 CPU or 1.5 GB): pictures at 1x, no filming (each film is hundreds of screenshots), the
 * browser closed between exports. PPTX_LEAN=1/0 decides it by hand; PPTX_FILM=1 films even on a small machine.
 */
export function leanMachine() {
  if (process.env.PPTX_LEAN === "1") return true;
  if (process.env.PPTX_LEAN === "0") return false;
  const { cpus, memory } = machine();
  return cpus < 1 || memory < 1.5 * 1024 ** 3;
}

let browserPromise = null;
async function browser(lean = false) {
  if (!browserPromise) {
    // A small machine starts the browser slowly (minutes when memory is short): wait longer than Playwright's 3.
    // On a small machine: one renderer, no GPU process, less JavaScript memory.
    const small = lean ? ["--disable-gpu", "--renderer-process-limit=1", "--disable-features=site-per-process,Translate,MediaRouter,OptimizationHints", "--disable-extensions", "--mute-audio", "--js-flags=--max-old-space-size=384"] : [];
    browserPromise = playwright().chromium.launch({ executablePath: chromePath(), timeout: lean ? 300_000 : 180_000, args: ["--no-sandbox", "--disable-dev-shm-usage", "--font-render-hinting=none", ...small] });
    browserPromise.then((b) => b.on("disconnected", () => { browserPromise = null; })).catch(() => { browserPromise = null; });
  }
  return browserPromise;
}

/** The memory the container uses now (cgroup), and this process's, for the log. */
function memoryInUse() {
  const mb = (n) => `${Math.round(n / 1024 ** 2)}MB`;
  let container = null;
  for (const path of ["/sys/fs/cgroup/memory.current", "/sys/fs/cgroup/memory/memory.usage_in_bytes"]) {
    const v = readNumber(path);
    if (v && /^\d+$/.test(v[0])) { container = Number(v[0]); break; }
  }
  const { cpus, memory } = machine();
  return `${container != null ? `${mb(container)} of ${mb(memory)}` : "?"}, server ${mb(process.memoryUsage().rss)}, ${Math.round(cpus * 100) / 100} CPU`;
}

/** Close the browser (after an export on a small machine, or to stop one that hangs). */
async function closeBrowser() {
  const pending = browserPromise;
  browserPromise = null;
  if (pending) await pending.then((b) => b.close()).catch(() => {});
}

/** Thrown when an export is stopped (a newer export of the same person, or the time limit). */
class Stopped extends Error {}
/** Thrown when filming has used up its time: the rest of the motion becomes PowerPoint's nearest effects. */
class OverBudget extends Error {}
// Filming time per export (PPTX_FILM_BUDGET_MS), and the slowest a frame may be on average: a machine slower than it
// looked (a frame takes seconds there) stops filming at once and still ends the export.
const FILM_BUDGET_MS = Number(process.env.PPTX_FILM_BUDGET_MS) || 300_000;
const SLOW_FRAME_MS = 1000;

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

// 背景の動き: LOOP_MS of it, and FADE_MS more that are blended into its start so the video loops without a jump.
const LOOP_FPS = 20;
const LOOP_MS = 8000;
const FADE_MS = 1000;

/** Pictures as a seamless loop: the last `fade` frames are mixed into the first ones (ffmpeg's blend), then the rest. */
async function encodeLoop(frames, fade) {
  const bin = ffmpegBinary();
  const dir = await mkdtemp(join(tmpdir(), "hpptx-loop-"));
  try {
    for (const [i, frame] of frames.entries()) await writeFile(join(dir, `f${String(i).padStart(5, "0")}.png`), frame);
    const length = frames.length - fade;
    const out = join(dir, "loop.mp4");
    const seq = (start) => ["-framerate", String(LOOP_FPS), "-start_number", String(start), "-i", join(dir, "f%05d.png")];
    // [0]: frames length… (what follows the end), [1]: frames 0… (the start): mixed from the first to the second;
    // [2]: the frames after the mix. Together: the end runs on into the start.
    const graph = `[0:v]trim=end_frame=${fade},setpts=PTS-STARTPTS,format=gbrp[a];[1:v]trim=end_frame=${fade},setpts=PTS-STARTPTS,format=gbrp[b];`
      + `[a][b]blend=all_expr='A*(1-N/${fade})+B*(N/${fade})'[mix];[2:v]trim=start_frame=${fade}:end_frame=${length},setpts=PTS-STARTPTS,format=gbrp[rest];`
      + "[mix][rest]concat=n=2:v=1:a=0,format=yuv420p,pad=ceil(iw/2)*2:ceil(ih/2)*2[v]";
    await new Promise((resolve, reject) => {
      const child = spawn(bin, ["-hide_banner", "-loglevel", "error", "-y", ...seq(length), ...seq(0), ...seq(0),
        "-filter_complex", graph, "-map", "[v]", "-c:v", "libx264", "-preset", "veryfast", "-crf", "20", "-r", String(LOOP_FPS), "-movflags", "+faststart", "-an", out],
      { stdio: ["ignore", "ignore", "pipe"] });
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

/** Slide `i`'s moving backdrop as a looping MP4 behind the slide: { src, poster }, or null. */
async function recordLoop(page, i, check = () => {}) {
  const info = await page.evaluate((index) => window.__pptx.mountLoop(index), i);
  if (!info?.animations) return null;
  const fade = Math.round((FADE_MS / 1000) * LOOP_FPS);
  const count = Math.round((LOOP_MS / 1000) * LOOP_FPS) + fade;
  const frames = [];
  for (let k = 0; k < count; k += 1) {
    check();
    await page.evaluate((t) => window.__pptx.seekLoop(t), (k * 1000) / LOOP_FPS);
    frames.push(await page.screenshot({ type: "png", scale: "css" }));
  }
  const mp4 = await encodeLoop(frames, fade);
  return { src: `data:video/mp4;base64,${mp4.toString("base64")}`, poster: dataUrl(frames[0]) };
}

/**
 * Film one HTML-only effect of slide `i` (pptx-collect's mountClip / seekClip): its region, frame by frame, as an MP4.
 * Returns { src, poster, box, dur, lasting, loop } for the entry, or null.
 */
async function recordClip(page, i, id, background, check = () => {}) {
  const info = await page.evaluate(([index, entryId, bg]) => window.__pptx.mountClip(index, entryId, bg), [i, id, background]);
  if (!info || !(info.box.w > 0 && info.box.h > 0)) return null;
  const dur = Math.min(info.dur, MAX_CLIP_MS);
  const count = Math.max(2, Math.ceil((dur / 1000) * FPS) + 1);
  const clip = { x: info.box.x, y: info.box.y, width: info.box.w, height: info.box.h };
  const frames = [];
  for (let k = 0; k < count; k += 1) {
    check();
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
export async function collectModel(deck, { base, timeoutMs = 240_000, lean = leanMachine(), film = !lean || process.env.PPTX_FILM === "1", progress = () => {}, signal = null }) {
  progress({ stage: "ブラウザを起動しています" });
  const b = await browser(lean);
  // Pictures at 2x read sharper in PowerPoint; a small machine draws them at 1x (a quarter of the memory and time).
  const context = await b.newContext({ viewport: { width: 1920, height: 1080 }, deviceScaleFactor: lean ? 1 : 2 });
  const page = await context.newPage();
  const errors = [];
  page.on("pageerror", (error) => errors.push(error.message));
  const started = Date.now();
  const check = () => { if (signal?.aborted) throw new Stopped(String(signal.reason || "取り消しました。")); };
  let skipped = 0;
  let filmUntil = 0;
  const frames = { count: 0, since: 0 };
  const filmCheck = () => {
    check();
    if (Date.now() > filmUntil) throw new OverBudget("filming took too long");
    if (!frames.since) frames.since = Date.now();
    frames.count += 1;
    if (frames.count > 10 && (Date.now() - frames.since) / frames.count > SLOW_FRAME_MS) throw new OverBudget("frames too slow on this machine");
  };
  try {
    progress({ stage: "準備しています" });
    await page.goto(`${base}/export.html`, { waitUntil: "load" });
    await page.waitForFunction(() => window.__pptxReady === true, null, { timeout: 60_000 });
    check();
    const info = await page.evaluate((d) => window.__pptx.load(d), deck);
    const slides = [];
    let filmed = 0;
    for (let i = 0; i < info.slides; i += 1) {
      check();
      if (Date.now() - started > timeoutMs) throw new Error("書き出しに時間がかかりすぎました。スライドを分けて試してください。");
      progress({ stage: `${i + 1}枚目を読み取っています`, done: i, total: info.slides });
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
      // (Not filmed: the nearest PowerPoint effect, and the backdrop as the picture under the slide.)
      const rethrow = (error) => {
        if (error instanceof Stopped) throw error;
        if (error instanceof OverBudget) { film = false; console.warn("pptx export: filming stopped (time budget); the rest uses PowerPoint's effects"); }
        return null;
      };
      if (film && !filmUntil) filmUntil = Date.now() + FILM_BUDGET_MS;
      if (s.clips?.length && film && ffmpegBinary()) {
        for (const [n, id] of s.clips.entries()) {
          if (filmed >= MAX_CLIPS || Date.now() - started > timeoutMs) break;
          progress({ stage: `${i + 1}枚目の動きを動画にしています（${n + 1}/${s.clips.length}）`, done: i, total: info.slides });
          const entry = s.timeline.find((e) => e.id === id);
          const got = entry && film ? await recordClip(page, i, id, s.background, filmCheck).catch((error) => { errors.push(`clip: ${error.message}`); return rethrow(error); }) : null;
          if (got) { entry.clip = got; filmed += 1; } else skipped += 1;
        }
      } else skipped += s.clips?.length ?? 0;
      delete s.clips;
      if (s.backdrop && film && ffmpegBinary() && Date.now() - started < timeoutMs) {
        progress({ stage: `${i + 1}枚目の背景の動きを動画にしています`, done: i, total: info.slides });
        const got = await recordLoop(page, i, filmCheck).catch((error) => { errors.push(`backdrop: ${error.message}`); return rethrow(error); });
        if (got) s.backdropVideo = { ...got, kind: s.backdrop };
        else skipped += 1;
      } else if (s.backdrop) skipped += 1;
      slides.push(s);
    }
    // The sounds the slides play (画面切り替えのサウンド, an effect's サウンド), made once for the whole deck.
    const names = slides.flatMap((s) => [s.transitionSound, ...(s.timeline || []).map((e) => e.sound)]).filter((n) => n && n !== "stop");
    const sounds = names.length ? await page.evaluate((list) => window.__pptx.sounds(list), names) : {};
    if (errors.length) console.warn(`pptx export: ${errors.slice(0, 3).join(" | ")}`);
    return { title: info.title, theme: info.theme, tableStyles: TABLE_STYLE_IDS, slides, sounds, unfilmed: skipped };
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

// The longest an export may take before it is stopped (the browser is closed, the next export can start).
const DEADLINE_MS = Number(process.env.PPTX_DEADLINE_MS) || 15 * 60_000;

/**
 * A deck (as the HTML export carries it: media inside) → { bytes, stats } of its .pptx. `progress({ stage, done,
 * total })` follows it; `signal` stops it. Exports run one at a time; one that runs past DEADLINE_MS is stopped.
 */
export function exportPptx(deck, { base, rootDir, progress = () => {}, signal = null, prepare = null }) {
  return serially(async () => {
    if (signal?.aborted) throw new Stopped(String(signal.reason || "取り消しました。"));
    const lean = leanMachine();
    // A small machine first gives back what it can spare (the AI's memory when no AI job is running).
    if (lean && prepare) {
      progress({ stage: "メモリを空けています" });
      await Promise.resolve(prepare()).catch(() => {});
    }
    console.log(`pptx export: starting (${deck.slides.length} slides${lean ? ", small machine" : ""}; memory in use ${memoryInUse()})`);
    const local = new AbortController();
    const stop = (reason) => { if (!local.signal.aborted) local.abort(reason); };
    // Stopped from outside (a newer export): close the browser too, so a page that hangs lets go at once.
    signal?.addEventListener("abort", () => { stop(signal.reason); closeBrowser(); }, { once: true });
    let timedOut = false;
    const timer = setTimeout(() => { timedOut = true; stop("timeout"); closeBrowser(); }, DEADLINE_MS);
    const started = Date.now();
    try {
      const model = await collectModel(deck, { base, lean, progress, signal: local.signal, timeoutMs: DEADLINE_MS - 60_000 });
      if (local.signal.aborted) throw new Stopped(String(local.signal.reason));
      progress({ stage: "PowerPoint のファイルを書いています", done: model.slides.length, total: model.slides.length });
      const result = await writePptx(model, { rootDir });
      result.stats = { ...result.stats, unfilmed: model.unfilmed || 0, lean };
      console.log(`pptx export: ${model.slides.length} slides in ${Math.round((Date.now() - started) / 1000)}s${lean ? " (small machine)" : ""}`);
      return result;
    } catch (error) {
      console.warn(`pptx export failed after ${Math.round((Date.now() - started) / 1000)}s: ${error.message}`);
      // A page that did not answer is still busy in the browser: start the next export with a fresh one.
      if (local.signal.aborted || /Target|closed|crash/i.test(error.message)) await closeBrowser();
      if (timedOut) throw new Error(`書き出しに時間がかかりすぎたため止めました（${Math.round(DEADLINE_MS / 60_000)}分）。スライドを分けて試してください。`);
      throw local.signal.aborted ? new Stopped(String(local.signal.reason)) : error;
    } finally {
      clearTimeout(timer);
      // A small machine needs its memory back for everything else (the AI, the next request).
      if (lean) await closeBrowser();
    }
  });
}

// ---------------------------------------------------------------- exports as jobs (the studio follows their progress)

const exportJobs = new Map();

/** Start an export for `owner` (a session): its earlier export, still waiting or running, is stopped first. */
export function startExportJob(deck, { owner, base, rootDir, prepare = null }) {
  for (const job of exportJobs.values()) {
    if (job.owner === owner && !job.finished) job.controller.abort("新しい書き出しを始めたため、前の書き出しを止めました。");
  }
  const job = { id: randomUUID(), owner, status: "queued", stage: "順番を待っています", done: 0, total: deck.slides.length, finished: false, controller: new AbortController(), startedAt: Date.now(), title: deck.title };
  exportJobs.set(job.id, job);
  const update = (patch) => Object.assign(job, patch, { status: "running", updatedAt: Date.now() });
  exportPptx(deck, { base, rootDir, signal: job.controller.signal, progress: update, prepare })
    .then(({ bytes, stats }) => Object.assign(job, { status: "completed", stage: "完了", bytes, stats, finished: true }))
    .catch((error) => Object.assign(job, { status: error instanceof Stopped ? "stopped" : "failed", error: error.message, finished: true }))
    .finally(() => {
      // The file waits a while to be fetched, then goes.
      setTimeout(() => exportJobs.delete(job.id), 15 * 60_000).unref?.();
    });
  return job;
}

/** An export's progress (for its owner only). */
export function exportJob(id, owner) {
  const job = exportJobs.get(id);
  if (!job || job.owner !== owner) return null;
  return job;
}

/** What the studio is told about an export. */
export function publicExportJob(job) {
  return { id: job.id, status: job.status, stage: job.stage, done: job.done, total: job.total, error: job.error ?? null, stats: job.stats ?? null, seconds: Math.round((Date.now() - job.startedAt) / 1000) };
}
