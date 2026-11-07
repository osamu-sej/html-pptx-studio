import { createHash, createHmac, randomUUID, timingSafeEqual } from "node:crypto";
import { createServer } from "node:http";
import { existsSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";
import { z } from "zod";

import {
  CodexSlideServer,
  buildChatPrompt,
  buildDeckPrompt,
  buildNotesPrompt,
  buildOutlinePrompt,
  buildRepairPrompt,
  buildRevisePrompt,
  buildVariantsPrompt,
  buildRewritePrompt,
  omitNullObjectValues,
  varietyRepairLines,
  withoutImageData,
} from "./server/codex-app-server.mjs";
import { applyChatOperations, summarizeChange } from "./server/chat.mjs";
import { capacityIssues } from "./server/capacity.mjs";
import { describeUnsourced, factRepairLines, unsourcedNumbers } from "./server/facts.mjs";
import { LocalModel } from "./server/local-ai.mjs";
import { reconcileChatVisuals, slideMeaning } from "./server/visual-relevance.mjs";
import { imageBrief } from "./server/image-brief.mjs";
import { extractText, importDeck, importSheets } from "./server/extract.mjs";
import { exportJob, exportPptx, publicExportJob, startExportJob } from "./server/pptx.mjs";
import { varietyIssues } from "./public/layout-looks.mjs";
import { createRooms } from "./server/rooms.mjs";
import { buildTranslatePrompt, checkTranslation, translateRequestSchema, translateResultSchema } from "./server/translate.mjs";
import { buildSynonymsPrompt, checkSynonyms, synonymsRequestSchema, synonymsResultSchema } from "./server/synonyms.mjs";
import { stageShowpieces } from "./server/showpieces.mjs";
import {
  chatRequestSchema,
  chatResultSchema,
  codexChatSchema,
  codexDeckSchema,
  codexNotesSchema,
  toCodexSchema,
  codexOutlineSchema,
  codexVariantsSchema,
  variantsRequestSchema,
  variantsResultSchema,
  outlineItemSchema,
  outlineResultSchema,
  codexSlideSchema,
  notesRequestSchema,
  notesResultSchema,
  deckShape,
  generatedDeckSchema,
  reviseRequestSchema,
  rewriteRequestSchema,
  aiSlideSchema,
  studioSettingsSchema,
} from "./server/schemas.mjs";

const here = dirname(fileURLToPath(import.meta.url));
const APP_VERSION = JSON.parse(readFileSync(join(here, "package.json"), "utf8")).version;
const APP_COMMIT = gitCommit();
const STARTED_AT = new Date();
const IMAGE_GENERATION_REQUEST = /(?:gpt[\s-]*image|(?:画像|イラスト|写真)[^。！？!?\n]{0,18}(?:生成|描いて|作って|作成)|(?:生成|作成)[^。！？!?\n]{0,10}(?:画像|イラスト))/i;
const REFERENCE_ASSETS = { executiveDecision: "executive-decision.jpg", storeOperations: "store-operations.jpg", customerExperience: "customer-experience.jpg", dataInsight: "data-insight.jpg", transformationRoadmap: "transformation-roadmap.jpg", businessWorkshop: "business-workshop.jpg", businessEtiquette: "business-etiquette.jpg", promptDesign: "prompt-design.jpg", aiWorkflow: "ai-workflow-visual.jpg", ai: "ai-executive-hero.jpg" };
const asksForImageGeneration = (message) => IMAGE_GENERATION_REQUEST.test(message);
/**
 * Which commit this server runs, so people can tell an updated app from a stale one:
 * Render passes it in the environment; a local checkout has it in .git (read without the git command).
 */
function gitCommit() {
  const fromEnv = process.env.RENDER_GIT_COMMIT || process.env.GIT_COMMIT || "";
  if (/^[0-9a-f]{7,40}$/i.test(fromEnv)) return fromEnv.slice(0, 7);
  try {
    const git = join(here, ".git");
    const head = readFileSync(join(git, "HEAD"), "utf8").trim();
    if (!head.startsWith("ref: ")) return head.slice(0, 7);
    const ref = head.slice(5);
    const loose = join(git, ref);
    const sha = existsSync(loose) ? readFileSync(loose, "utf8") : readFileSync(join(git, "packed-refs"), "utf8").split("\n").find((line) => line.endsWith(` ${ref}`)) ?? "";
    return /^[0-9a-f]{40}/.test(sha.trim()) ? sha.trim().slice(0, 7) : "";
  } catch {
    return "";
  }
}

function revisionEffort(instruction, focus = []) {
  return focus.length > 1 || /全体|構成|骨子|流れ|前後|写真|画像|図解|デザイン|説得|意思決定|根拠|動き/.test(instruction) ? "medium" : "low";
}
function imagePrompt({ message, deck, index }, hasReference = false) {
  const slide = deck.slides[index];
  const previous = deck.slides[index - 1];
  const next = deck.slides[index + 1];
  return [
    "ユーザーの依頼に従い、必ず組み込みの画像生成ツール（$imagegen）で新規画像を1枚生成してください。既存の画像検索・ストック素材・SVGやコードでの代替は禁止です。生成できなければできないと報告してください。",
    "スライド用の16:9横長ビジュアル。画像内に文字・ロゴ・商標・透かしを描かない。スライドの主張に関係しない汎用的な人物・会議風景を足さない。",
    `資料全体: ${deck.title}／対象者: ${deck.audience || "未指定"}／目的: ${deck.purpose || "未指定"}`,
    deck.memo ? `資料の前提: ${deck.memo.slice(0, 500)}` : "",
    `対象: ${index + 1}枚目「${slide.title || slide.message || deck.title}」`,
    `画像化の骨子（必ず反映）:\n${imageBrief(slide)}`,
    slide.takeaway ? `伝えたい内容: ${slide.takeaway}` : "",
    `対象スライドの具体的な本文: ${slideMeaning(slide).slice(0, 900)}`,
    previous ? `前のスライド: ${previous.title || ""}／${previous.takeaway || previous.message || ""}` : "",
    next ? `次のスライド: ${next.title || ""}／${next.takeaway || next.message || ""}` : "",
    hasReference ? "参考画像を添付しました。ユーザーが元写真を生かすよう求めた場合は、構図や被写体を参照しつつ新しい画像に仕上げてください。" : "",
    `ユーザーの指示（絵柄・内容・雰囲気を最優先）: ${message}`,
    "生成前に『この画像のどの部分が対象スライドの主張・工程・時間帯に対応するか』を自分で確認する。対応を説明できない汎用的な絵は不採用として作り直す。元資料にない成果・数字・事実・実在人物を暗示しない。",
    "生成した画像を保存し、完成画像ファイルを結果として返してください。",
  ].filter(Boolean).join("\n");
}

// ---------------------------------------------------------------- static files

const PUBLIC = join(here, "public");
const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8", ".jpg": "image/jpeg", ".png": "image/png", ".svg": "image/svg+xml", ".txt": "text/plain; charset=utf-8" };
const icons = readFileSync(join(PUBLIC, "engine", "icons.json"), "utf8").trim();
function loadStatic() {
  const files = new Map();
  const add = (path, body, type) => files.set(path, { body: Buffer.isBuffer(body) ? body : Buffer.from(body), type, etag: `"${createHash("sha1").update(body).digest("base64url").slice(0, 16)}"` });
  const started = STARTED_AT.toLocaleString("ja-JP", { timeZone: "Asia/Tokyo", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
  const build = `バージョン ${APP_VERSION}${APP_COMMIT ? `（コミット ${APP_COMMIT}）` : ""}・サーバー起動 ${started}`;
  const inject = (text) => text.replaceAll("__APP_VERSION__", APP_VERSION).replaceAll("__APP_BUILD__", build);
  add("/", inject(readFileSync(join(PUBLIC, "index.html"), "utf8")), TYPES[".html"]);
  add("/app.js", inject(readFileSync(join(PUBLIC, "app.js"), "utf8")), TYPES[".js"]);
  // The page the server's headless browser draws slides on to write a PowerPoint (server/pptx.mjs).
  add("/export.html", inject(readFileSync(join(PUBLIC, "export.html"), "utf8")), TYPES[".html"]);
  add("/reset-actions.mjs", readFileSync(join(PUBLIC, "reset-actions.mjs"), "utf8"), TYPES[".js"]);
  add("/saved-library.js", readFileSync(join(PUBLIC, "saved-library.js"), "utf8"), TYPES[".js"]);
  add("/auto-images.mjs", readFileSync(join(PUBLIC, "auto-images.mjs"), "utf8"), TYPES[".js"]);
  add("/app.css", readFileSync(join(PUBLIC, "app.css"), "utf8"), TYPES[".css"]);
  add("/layout-looks.mjs", readFileSync(join(PUBLIC, "layout-looks.mjs"), "utf8"), TYPES[".js"]);
  // The PowerPoint-style editor (objects placed by hand).
  for (const name of readdirSync(join(PUBLIC, "editor")).filter((n) => /^[a-z0-9-]+\.mjs$/.test(n))) add(`/editor/${name}`, readFileSync(join(PUBLIC, "editor", name), "utf8"), TYPES[".js"]);
  add("/engine/engine.js", readFileSync(join(PUBLIC, "engine", "engine.js"), "utf8").replace("/*__ICONS__*/{}", () => icons), TYPES[".js"]);
  add("/engine/objects.js", readFileSync(join(PUBLIC, "engine", "objects.js"), "utf8"), TYPES[".js"]);
  add("/engine/animate.js", readFileSync(join(PUBLIC, "engine", "animate.js"), "utf8"), TYPES[".js"]);
  add("/engine/motion.js", readFileSync(join(PUBLIC, "engine", "motion.js"), "utf8"), TYPES[".js"]);
  add("/engine/models.js", readFileSync(join(PUBLIC, "engine", "models.js"), "utf8"), TYPES[".js"]);
  add("/engine/engine.css", readFileSync(join(PUBLIC, "engine", "engine.css"), "utf8"), TYPES[".css"]);
  for (const name of readdirSync(join(PUBLIC, "assets")).filter((n) => /^[a-z0-9-]+\.jpg$/.test(n))) add(`/assets/${name}`, readFileSync(join(PUBLIC, "assets", name)), TYPES[".jpg"]);
  for (const name of readdirSync(join(PUBLIC, "samples")).filter((n) => /^[a-z0-9-]+\.json$/.test(n))) add(`/samples/${name}`, readFileSync(join(PUBLIC, "samples", name)), TYPES[".json"]);
  // Lottie player (lottie-web, MIT) for motion-graphic animations: the light SVG build, no eval, no workers.
  const lottie = join(here, "node_modules", "lottie-web", "build", "player", "lottie_light.min.js");
  if (existsSync(lottie)) add("/vendor/lottie.js", readFileSync(lottie), TYPES[".js"]);
  // Polygon boolean operations (polygon-clipping, MIT) for 図形の結合, loaded by the editor on first use.
  const clipping = join(here, "node_modules", "polygon-clipping", "dist", "polygon-clipping.umd.min.js");
  if (existsSync(clipping)) add("/vendor/polygon-clipping.js", readFileSync(clipping), TYPES[".js"]);
  // QR codes (qrcode-generator, MIT) for the link to an online presentation.
  const qrcode = join(here, "node_modules", "qrcode-generator", "dist", "qrcode.mjs");
  if (existsSync(qrcode)) add("/vendor/qrcode.mjs", readFileSync(qrcode), TYPES[".js"]);
  // 3D models (three.js, MIT): the core and the glTF loader, with their imports pointed at each other (no import map).
  for (const [name, code] of threeFiles()) add(`/vendor/three/${name}`, code, TYPES[".js"]);
  return files;
}
/** three.js's files for 3D models: [name, code] with the loader's imports made relative. */
function threeFiles() {
  const dir = join(here, "node_modules", "three");
  if (!existsSync(join(dir, "build", "three.module.js"))) return [];
  const relative = (code) => code.replace(/from\s*(['"])three\1/g, "from './three.module.js'").replace(/from\s*(['"])\.\.\/utils\/([A-Za-z]+\.js)\1/g, "from './$2'");
  return [
    ["three.core.js", readFileSync(join(dir, "build", "three.core.js"), "utf8")],
    ["three.module.js", readFileSync(join(dir, "build", "three.module.js"), "utf8")],
    ...[["GLTFLoader.js", "loaders/GLTFLoader.js"], ["BufferGeometryUtils.js", "utils/BufferGeometryUtils.js"], ["SkeletonUtils.js", "utils/SkeletonUtils.js"]]
      .map(([name, path]) => [name, relative(readFileSync(join(dir, "examples", "jsm", path), "utf8"))]),
  ];
}
const staticFiles = loadStatic();
const PORT = Number(process.env.PORT ?? 8787);
let clientErrors = { count: 0, resetAt: 0 };
function clientErrorBudget() {
  const now = Date.now();
  if (clientErrors.resetAt < now) clientErrors = { count: 0, resetAt: now + 60_000 };
  clientErrors.count += 1;
  return clientErrors.count <= 30;
}

const APP_ORIGIN = process.env.APP_ORIGIN ?? process.env.RENDER_EXTERNAL_URL;
const STATELESS_API = new Set(["/api/extract", "/api/import", "/api/tables/xlsx"]);
const localAi = new LocalModel();
const codex = new CodexSlideServer({ rootDir: here, fallback: localAi });
// Team mode (opt-in): anyone with STUDIO_PASSCODE may use the shared Codex login,
// from any number of browsers. Without it, the first browser to log in owns the app.
const TEAM_PASSCODE = process.env.STUDIO_PASSCODE ?? "";
const TEAM_MODE = TEAM_PASSCODE.length > 0;
const memberKey = createHash("sha256").update(`html-slide-studio:${TEAM_PASSCODE}`).digest();
const SESSION_COOKIE = "hs_session";
const MEMBER_COOKIE = "hs_member";
let codexReady = false;
seedCodexAuth();
const sessions = new Map();
const jobOwners = new Map();
const streams = new Map();
const ownerSessionPath = join(codex.runtimeDir, "owner-session-id");
let ownerSessionId = null;
try { ownerSessionId = readFileSync(ownerSessionPath, "utf8").trim() || null; }
catch { ownerSessionId = null; }

/** Restore a Codex login from CODEX_AUTH_JSON so free-tier redeploys don't log the team out. */
function seedCodexAuth() {
  const raw = process.env.CODEX_AUTH_JSON;
  const target = join(codex.codexHome, "auth.json");
  if (!raw || existsSync(target)) return;
  try {
    const text = raw.trim().startsWith("{") ? raw : Buffer.from(raw, "base64").toString("utf8");
    JSON.parse(text);
    writeFileSync(target, text, { encoding: "utf8", mode: 0o600 });
  } catch (error) {
    console.error("CODEX_AUTH_JSON is not valid JSON; ignoring it.", error.message);
  }
}

const memberToken = (sessionId) => createHmac("sha256", memberKey).update(sessionId).digest("base64url");

function passcodeMatches(candidate) {
  const a = createHash("sha256").update(String(candidate)).digest();
  const b = createHash("sha256").update(TEAM_PASSCODE).digest();
  return timingSafeEqual(a, b);
}

const passcodeAttempts = new Map();
function passcodeRateLimited(req) {
  const ip = String(req.headers["x-forwarded-for"] ?? req.socket.remoteAddress ?? "").split(",")[0].trim();
  const now = Date.now();
  const entry = passcodeAttempts.get(ip);
  if (!entry || entry.resetAt < now) {
    passcodeAttempts.set(ip, { count: 1, resetAt: now + 10 * 60_000 });
    return false;
  }
  entry.count += 1;
  return entry.count > 10;
}

function persistOwnerSession() {
  try { writeFileSync(ownerSessionPath, ownerSessionId ?? "", { encoding: "utf8", mode: 0o600 }); }
  catch (error) { console.error("Could not persist owner session", error); }
}

function zodMessage(error) {
  return error?.issues?.map((issue) => `${issue.path.join(".")}: ${issue.message}`).join("／") || error?.message || String(error);
}

function parseCookies(req) {
  return Object.fromEntries(String(req.headers.cookie ?? "").split(";").map((part) => {
    const index = part.indexOf("=");
    return index < 0 ? ["", ""] : [part.slice(0, index).trim(), decodeURIComponent(part.slice(index + 1))];
  }).filter(([key]) => key));
}

function getSession(req, res) {
  const cookies = parseCookies(req);
  let id = cookies[SESSION_COOKIE];
  if (!id || !sessions.has(id)) {
    if (!id || !/^[0-9a-f-]{36}$/.test(id)) id = randomUUID();
    const member = TEAM_MODE && cookies[MEMBER_COOKIE] === memberToken(id);
    sessions.set(id, { id, authorized: id === ownerSessionId, member, createdAt: Date.now() });
    const secure = req.headers["x-forwarded-proto"] === "https" ? "; Secure" : "";
    res.setHeader("Set-Cookie", `${SESSION_COOKIE}=${encodeURIComponent(id)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=2592000${secure}`);
  }
  return sessions.get(id);
}

function json(res, status, value) {
  res.writeHead(status, {
    "content-type": "application/json; charset=utf-8",
    "cache-control": "no-store",
    "x-content-type-options": "nosniff",
  }).end(JSON.stringify(value));
}

async function readJson(req, maxBytes = 500_000) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > maxBytes) throw new Error("リクエストが大きすぎます。");
    chunks.push(chunk);
  }
  const body = Buffer.concat(chunks).toString("utf8");
  return body ? JSON.parse(body) : {};
}

async function readBody(req, maxBytes, message) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > maxBytes) throw new Error(message);
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

function sameOrigin(req) {
  const origin = req.headers.origin;
  if (!origin) return true;
  const expected = `${req.headers["x-forwarded-proto"] ?? "http"}://${req.headers["x-forwarded-host"] ?? req.headers.host}`;
  return origin === expected || (APP_ORIGIN && origin === APP_ORIGIN.replace(/\/$/, ""));
}

function authorized(session) {
  if (TEAM_MODE) return Boolean(session.member && codexReady);
  return Boolean(session.authorized && ownerSessionId === session.id);
}

/**
 * May this session use AI? With Codex connected, yes. Without it, only when the stand-by AI (Gemma on
 * this machine) is there to answer; never on the web deployment.
 */
async function aiAllowed(session) {
  if (authorized(session)) return true;
  if (TEAM_MODE && !session.member) return false;
  return (await localAi.status()).available;
}

function sendEvent(res, event, data) {
  res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
}

codex.on("login", ({ sessionId, success, error }) => {
  const session = sessionId ? sessions.get(sessionId) : null;
  if (success) codexReady = true;
  if (success && session && !TEAM_MODE) {
    session.authorized = true;
    ownerSessionId = sessionId;
    persistOwnerSession();
  }
  if (session) session.loginError = success ? null : (typeof error === "string" ? error : error?.message) || "ログインに失敗しました。";
});

codex.on("job", (job) => {
  const listeners = streams.get(job.id);
  if (!listeners) return;
  for (const res of listeners) sendEvent(res, "progress", job);
  if (["completed", "failed"].includes(job.status)) {
    for (const res of listeners) res.end();
    streams.delete(job.id);
  }
});

codex.on("exit", (error) => console.error(error));

const deckRequestSchema = z.object({
  deckTitle: z.string().max(100).optional().default(""),
  brief: z.string().min(1).max(100_000),
  audience: z.string().max(80).optional().default(""),
  purpose: z.string().max(180).optional().default(""),
  tone: z.string().max(80).optional().default("標準ビジネス"),
  settings: studioSettingsSchema,
  outline: z.array(outlineItemSchema).min(2).max(50).optional(),
});
const outlineRequestSchema = deckRequestSchema.extend({
  outline: z.array(outlineItemSchema).min(2).max(50).optional(),
  instruction: z.string().max(1000).optional().default(""),
});

function newJob(kind, session, input) {
  const now = new Date().toISOString();
  const job = { id: randomUUID(), kind, status: "queued", stage: "受付完了", detail: "Codexへ送信します。", createdAt: now, updatedAt: now, input, useCodex: authorized(session) };
  jobOwners.set(job.id, session.id);
  return job;
}

function startOutlineJob(session, input) {
  const job = newJob("outline", session, { brief: input.brief.slice(0, 200) });
  // Revising someone's outline may keep the layouts they chose: only ask again when the AI made it plainer.
  const baseline = input.outline ? varietyIssues(input.outline).length : 0;
  codex.runJob(job, {
    prompt: buildOutlinePrompt(input),
    outputSchema: codexOutlineSchema,
    maxAttempts: 3,
    effort: "low",
    async finalize(candidate, attempt) {
      const outline = outlineResultSchema.parse(omitNullObjectValues(candidate));
      if (outline.slides[0].type !== "title") throw new Error("slides[0] の type を title にしてください。");
      if (outline.slides.at(-1).type !== "closing") throw new Error("最後のスライドの type を closing にしてください。");
      if (outline.slides.slice(1, -1).some((item) => ["title", "closing"].includes(item.type))) throw new Error("表紙と最後以外に title・closing を使わないでください。");
      const result = { outline };
      const plain = varietyIssues(outline.slides);
      if (plain.length > baseline && attempt <= 2) {
        return { retry: [...varietyRepairLines(plain, { outline: true }), "骨子全体のJSONをもう一度返してください。"].join("\n"), fallback: result, detail: "同じ見た目のスライドが続かないよう、レイアウトを選び直しています。" };
      }
      return { result, detail: `${outline.slides.length}枚の骨子ができました。` };
    },
  }).catch((error) => codex.failJob(job.id, error));
  return job;
}

function startImageJob(session, request) {
  const job = newJob("image", session, { message: request.message, index: request.index });
  const slide = request.deck.slides[request.index];
  let referenceImage = null;
  const asksForReference = /(?:既存|元|この|今の|現在の|添付)[^。！？!?\n]{0,15}(?:写真|画像|イラスト)|(?:写真|画像|イラスト)[^。！？!?\n]{0,15}(?:元に|参考|ベース|活か|生か|使って)/.test(request.message);
  const embeddedImage = slide.media?.src?.startsWith("data:image/") ? slide.media.src : slide.customImage || (typeof slide.image === "string" ? slide.image : "");
  if (asksForReference && embeddedImage) {
    const [, mime, encoded] = embeddedImage.match(/^data:image\/(png|jpeg|webp);base64,(.+)$/) ?? [];
    if (mime && encoded) referenceImage = { bytes: Buffer.from(encoded, "base64"), extension: mime === "jpeg" ? "jpg" : mime };
  } else if (asksForReference && REFERENCE_ASSETS[slide.visualAsset]) {
    referenceImage = { bytes: staticFiles.get(`/assets/${REFERENCE_ASSETS[slide.visualAsset]}`)?.body, extension: "jpg" };
  }
  codex.runImageJob(job, imagePrompt(request, Boolean(referenceImage)), { referenceImage }).catch((error) => codex.failJob(job.id, error));
  return job;
}

/** What the AI may take figures from: the material, the settings and anything the user wrote into the outline. */
function deckSource(input) {
  return [input.deckTitle, input.purpose, input.audience, input.brief, JSON.stringify(input.outline ?? [])].join("\n");
}

/** The deck the AI was shown (images left out), as the source of figures for edits to it. */
function slidesSource(slides, ...extra) {
  return [JSON.stringify(withoutImageData(slides)), ...extra].join("\n");
}

/**
 * `variety` is how many look problems (varietyIssues) the answer may have before the AI is asked to pick
 * other layouts: 0 for a deck written from scratch, the current count when rewriting a deck, and null
 * (never ask) when the layouts come from an outline the user agreed to.
 */
// Slides written per answer when a deck is written in parts.
const DECK_PART = 4;

/** What is wrong with slides written in a part of a deck (each slide against the slide schema), or "". */
function slidesProblem(slides, { from }) {
  const problems = [];
  slides.forEach((slide, index) => {
    const checked = aiSlideSchema.safeParse(omitNullObjectValues(slide));
    if (checked.success) return;
    const issue = checked.error.issues[0];
    problems.push(`${from + index}枚目（${slide?.type ?? "type なし"}）の ${issue?.path?.join(".") || "全体"}：${issue?.message ?? "形式が違います"}`);
  });
  return problems.slice(0, 4).join("／");
}

function startDeckJob(session, input, prompt = buildDeckPrompt(input), source = deckSource(input), { variety = input.outline ? null : 0 } = {}) {
  const job = newJob("deck", session, input);
  let layoutsAsked = false;
  // How many slides the deck will have (an outline fixes it; a range is met in the middle when it must be split).
  const { minCount, maxCount, slideCount } = input.settings;
  const total = input.outline?.length || (minCount ? Math.round((minCount + maxCount) / 2) : slideCount) || 0;
  codex.runJob(job, {
    prompt,
    outputSchema: codexDeckSchema,
    maxAttempts: 3,
    // From an outline, a long deck is written DECK_PART slides at a time from the start (one answer cannot hold
    // it all); otherwise only when an answer did not fit.
    ...(input.outline && total > DECK_PART ? { parts: { total, size: DECK_PART } } : { partsTotal: total }),
    partCheck: slidesProblem,
    async finalize(candidate, attempt) {
      const deck = generatedDeckSchema.parse(omitNullObjectValues(candidate));
      const count = deck.slideData.length;
      const { minCount, maxCount, slideCount } = input.settings;
      if (minCount && (count < minCount || count > maxCount)) throw new Error(`枚数は${minCount}〜${maxCount}枚にしてください（回答は${count}枚）。`);
      if (!minCount && count !== slideCount) throw new Error(`指定は${slideCount}枚ですが、回答は${count}枚でした。slideDataをちょうど${slideCount}枚にしてください。`);
      if (deck.slideData[0]?.type !== "title") throw new Error("先頭スライドのtypeをtitleにしてください。");
      if (deck.slideData.at(-1)?.type !== "closing") throw new Error("最終スライドのtypeをclosingにしてください。");
      const overflow = capacityIssues(deck.slideData).filter((issue) => issue.severity === "error");
      const unsourced = unsourcedNumbers(deck.slideData, source);
      // Writing a whole deck again is slow: layouts get one second chance.
      const plainIssues = variety === null || layoutsAsked ? [] : varietyIssues(deck.slideData);
      const plain = plainIssues.length > variety && attempt <= 2 ? plainIssues : [];
      if (plain.length) layoutsAsked = true;
      // 見せ場: the moments the AI left still get motion only HTML can give (spin, swirl, big words, a moving backdrop).
      const staged = stageShowpieces(deck.slideData);
      const result = { deck: { ...deck, slideData: staged.slides, settings: input.settings }, issues: [...overflow, ...unsourced] };
      if ((overflow.length || unsourced.length || plain.length) && attempt <= 2) {
        const retry = [overflow.length ? buildRepairPrompt({ issues: overflow, slides: deck.slideData }) : "", ...factRepairLines(unsourced, deck.slideData),
          ...(plain.length ? varietyRepairLines(plain) : []),
          !overflow.length && (unsourced.length || plain.length) ? "枚数・順序・他のスライドは変えずに、全体のJSONをもう一度返してください。" : ""].filter(Boolean).join("\n");
        return { retry, fallback: result, detail: [overflow.length ? `${overflow.length}か所の長すぎる文` : "", unsourced.length ? `素材にない数値${unsourced.length}か所` : "", plain.length ? "同じ見た目が続くレイアウト" : ""].filter(Boolean).join("と") + "を自動で修正しています。" };
      }
      return { result, detail: `${count}枚の構成が完成しました。${overflow.length ? `（${overflow.length}か所は要確認）` : ""}${unsourced.length ? describeUnsourced(unsourced) : ""}` };
    },
  }).catch((error) => codex.failJob(job.id, error));
  return job;
}

function startReviseJob(session, request) {
  const job = newJob("revise", session, request);
  const { deck, slideIndex } = request;
  const total = deck.slides.length;
  codex.runJob(job, {
    prompt: buildRevisePrompt(request),
    outputSchema: codexSlideSchema,
    maxAttempts: 2,
    effort: revisionEffort(request.instruction || ""),
    async finalize(candidate, attempt) {
      const { slide: proposed } = z.object({ slide: aiSlideSchema }).parse(omitNullObjectValues(candidate));
      const inserting = request.mode === "insert";
      const { answer: checked, warnings } = reconcileChatVisuals(deck, { operations: [{ op: inserting ? "insert" : "replace", slide: slideIndex + 1, content: proposed }] }, request.instruction || "");
      let slide = checked.operations[0]?.content ?? deck.slides[slideIndex];
      if (!inserting) slide = applyChatOperations(deck.slides, { operations: [{ op: "replace", slide: slideIndex + 1, content: slide }] }).slides[slideIndex];
      if (inserting && ["title", "closing"].includes(slide.type)) throw new Error("追加するスライドのtypeはtitleとclosing以外にしてください。");
      if (!inserting && slideIndex === 0 && slide.type !== "title") throw new Error("表紙なのでtypeはtitleのままにしてください。");
      if (!inserting && slideIndex === total - 1 && slide.type !== "closing") throw new Error("最終スライドなのでtypeはclosingのままにしてください。");
      const slides = inserting
        ? [...deck.slides.slice(0, slideIndex), slide, ...deck.slides.slice(slideIndex)]
        : deck.slides.map((item, index) => (index === slideIndex ? slide : item));
      const issues = capacityIssues(slides, { only: new Set([slideIndex]) }).filter((issue) => issue.severity === "error");
      const unsourced = unsourcedNumbers(slides, slidesSource(deck.slides, request.instruction, deck.memo), { only: new Set([slideIndex]) });
      const result = { slide, issues: [...issues, ...unsourced, ...warnings.map((message) => ({ kind: "visual", severity: "warning", message }))] };
      if ((issues.length || unsourced.length || warnings.length) && attempt < 2) {
        const retry = [issues.length ? buildRepairPrompt({ issues, slides }).replace("全体のJSON", "このスライドのJSON") : "", ...factRepairLines(unsourced, slides),
          ...warnings.map((warning) => `${warning} 依頼と本文に直接合う素材がなければ写真を変更しないでください。`),
          unsourced.length && !issues.length ? "このスライドのJSONをもう一度返してください。" : ""].filter(Boolean).join("\n");
        return { retry, fallback: result, detail: issues.length ? "長すぎる文を自動で修正しています。" : unsourced.length ? "素材にない数値を直しています。" : "画像の関連性を確認しています。" };
      }
      return { result, detail: inserting ? `${slideIndex + 1}枚目にスライドを追加しました。` : `${slideIndex + 1}枚目を作り直しました。` };
    },
  }).catch((error) => codex.failJob(job.id, error));
  return job;
}

function startChatJob(session, request) {
  const job = newJob("chat", session, { message: request.message });
  const { deck } = request;
  codex.runJob(job, {
    prompt: buildChatPrompt(request),
    outputSchema: codexChatSchema,
    maxAttempts: 2,
    effort: revisionEffort(request.message, request.focus),
    async finalize(candidate, attempt) {
      const parsed = chatResultSchema.parse(omitNullObjectValues(candidate));
      const { answer, warnings } = reconcileChatVisuals(deck, parsed, request.message);
      const applied = applyChatOperations(deck.slides, answer);
      let issues = [];
      let unsourced = [];
      if (applied.changed) {
        const changed = new Set(applied.slides.map((_, index) => index).filter((index) => applied.items[index]?.changed));
        issues = capacityIssues(applied.slides, { only: changed }).filter((issue) => issue.severity === "error");
        const talk = (request.history ?? []).filter((turn) => turn.role === "user").map((turn) => turn.text);
        unsourced = unsourcedNumbers(applied.slides, slidesSource(deck.slides, request.message, deck.memo, request.attachment?.text, ...talk), { only: changed });
      }
      const title = answer.deckTitle?.trim() && answer.deckTitle.trim() !== deck.title ? answer.deckTitle.trim() : null;
      const transition = answer.transition && answer.transition !== deck.transition ? answer.transition : null;
      // Deck-wide motion graphics: keep only what actually changes.
      const motionPatch = Object.fromEntries(Object.entries(answer.motion ?? {}).filter(([key, value]) => value != null && value !== (deck.motion ?? {})[key]));
      const motion = Object.keys(motionPatch).length ? motionPatch : null;
      const result = {
        chat: {
          reply: warnings.length ? `${applied.changed ? "文章・構成の修正案は用意しました。" : "画像の変更は提案していません。"}\n${warnings.join("\n")}` : answer.reply,
          suggestions: answer.suggestions,
          deckTitle: title,
          transition,
          motion,
          changed: applied.changed || Boolean(title || transition || motion),
          summary: [summarizeChange(applied), title ? "資料名を変更" : "", transition ? "切り替えを変更" : "", motion ? "動きを変更" : ""].filter(Boolean).join("・"),
          slides: applied.changed ? applied.slides : null,
          items: applied.items,
          deleted: applied.deleted,
          moved: applied.moved,
          issues: [...issues, ...unsourced, ...warnings.map((message) => ({ kind: "visual", severity: "warning", message }))],
        },
      };
      if ((issues.length || unsourced.length || warnings.length) && attempt < 2) {
        const where = (index) => (applied.items[index]?.from == null ? "追加したスライド" : `${applied.items[index].from + 1}枚目`);
        const lines = issues.map((issue) => `- ${where(issue.slide)}「${applied.slides[issue.slide]?.title ?? ""}」の ${issue.field}: ${issue.message}`);
        return {
          retry: [
            ...(issues.length ? ["提案した内容を確認したところ、次のテキストがスライドに収まらない長さです。", ...lines, "該当箇所を意味を保ったまま短くするか、要素を減らしてください（補足は details に回してよい）。"] : []),
            ...factRepairLines(unsourced, applied.slides, where),
            ...warnings.map((warning) => `- ${warning} 写真が必要ならスライドの主張と本文に直接合う素材を選び、なければ追加しないでください。`),
            "番号は元の資料のままで、同じ形式（reply・operations・suggestions）でもう一度回答してください。"].join("\n"),
          fallback: result,
          detail: issues.length ? "長すぎる文を自動で修正しています。" : unsourced.length ? "素材にない数値を直しています。" : "画像の関連性を確認しています。",
        };
      }
      return { result, detail: result.chat.summary || "回答しました。" };
    },
  }).catch((error) => codex.failJob(job.id, error));
  return job;
}

function startVariantsJob(session, request) {
  const job = newJob("variants", session, { slideIndex: request.slideIndex });
  const { deck, slideIndex } = request;
  const total = deck.slides.length;
  codex.runJob(job, {
    prompt: buildVariantsPrompt(request),
    outputSchema: codexVariantsSchema,
    maxAttempts: 2,
    effort: "low",
    async finalize(candidate) {
      const { variants } = variantsResultSchema.parse(omitNullObjectValues(candidate));
      const fixed = slideIndex === 0 ? "title" : slideIndex === total - 1 ? "closing" : null;
      if (fixed && variants.some((variant) => variant.slide.type !== fixed)) throw new Error(`このスライドの type は ${fixed} のままにしてください。`);
      if (!fixed && variants.some((variant) => ["title", "closing"].includes(variant.slide.type))) throw new Error("type は title と closing 以外にしてください。");
      const out = variants.map((variant) => {
        const { answer, warnings } = reconcileChatVisuals(deck, { operations: [{ op: "replace", slide: slideIndex + 1, content: variant.slide }] }, request.instruction || "");
        const slide = applyChatOperations(deck.slides, answer).slides[slideIndex];
        const slides = deck.slides.map((item, index) => (index === slideIndex ? slide : item));
        const unsourced = unsourcedNumbers(slides, slidesSource(deck.slides, request.instruction, deck.memo), { only: new Set([slideIndex]) });
        return { label: variant.label, slide, issues: [...capacityIssues(slides, { only: new Set([slideIndex]) }), ...warnings.map((message) => ({ kind: "visual", severity: "warning", message }))], unsourced };
      });
      return { result: { variants: out }, detail: `${out.length}つの案ができました。` };
    },
  }).catch((error) => codex.failJob(job.id, error));
  return job;
}

function startTranslateJob(session, request) {
  const job = newJob("translate", session, { to: request.to, count: request.texts.length });
  codex.runJob(job, {
    prompt: buildTranslatePrompt(request),
    outputSchema: toCodexSchema(translateResultSchema),
    maxAttempts: 2,
    effort: "low",
    async finalize(candidate) {
      const texts = checkTranslation(request, omitNullObjectValues(candidate));
      return { result: { texts }, detail: `${texts.length}か所を訳しました。` };
    },
  }).catch((error) => codex.failJob(job.id, error));
  return job;
}

function startSynonymsJob(session, request) {
  const job = newJob("synonyms", session, { word: request.word });
  codex.runJob(job, {
    prompt: buildSynonymsPrompt(request),
    outputSchema: toCodexSchema(synonymsResultSchema),
    maxAttempts: 2,
    effort: "low",
    async finalize(candidate) {
      const senses = checkSynonyms(request, omitNullObjectValues(candidate));
      return { result: { senses }, detail: senses.length ? `${senses.reduce((n, s) => n + s.words.length, 0)}語見つかりました。` : "言い換えの語は見つかりませんでした。" };
    },
  }).catch((error) => codex.failJob(job.id, error));
  return job;
}

function startNotesJob(session, request) {
  const job = newJob("notes", session, request);
  const total = request.deck.slides.length;
  const wanted = new Set((request.indices?.length ? request.indices : request.deck.slides.map((_, index) => index)).filter((index) => index < total));
  codex.runJob(job, {
    prompt: buildNotesPrompt(request),
    outputSchema: codexNotesSchema,
    maxAttempts: 2,
    effort: "low",
    async finalize(candidate, attempt) {
      const { notes } = notesResultSchema.parse(omitNullObjectValues(candidate));
      const usable = notes.filter((note) => wanted.has(note.slide - 1)).map((note) => ({ slide: note.slide - 1, text: note.text.trim() }));
      if (!usable.length) throw new Error("指定したスライドのノートがありません。slide には1から始まる枚数の番号を入れてください。");
      const result = { notes: usable };
      // A spoken figure the slides do not show is an invented one.
      const source = slidesSource(request.deck.slides, request.deck.memo);
      const unsourced = unsourcedNumbers(usable.map((note) => ({ notes: note.text })), source);
      if (unsourced.length && attempt < 2) {
        return {
          retry: ["次のノートに、スライドにない数値があります。スライドに書いてある数値だけを使って、同じ形式でもう一度すべてのノートを返してください。",
            ...unsourced.map((issue) => `- ${usable[issue.slide].slide + 1}枚目: ${issue.values.join("、")}`)].join("\n"),
          fallback: result,
          detail: "ノートの数値を確認しています。",
        };
      }
      return { result, detail: `${usable.length}枚分のノートを作成しました。` };
    },
  }).catch((error) => codex.failJob(job.id, error));
  return job;
}

let extractSlots = 3;
const extractQueue = [];
async function withExtractSlot(task) {
  if (extractSlots <= 0) await new Promise((resolve) => extractQueue.push(resolve));
  extractSlots -= 1;
  try { return await task(); }
  finally {
    extractSlots += 1;
    extractQueue.shift()?.();
  }
}

// Fonts (Google Fonts), YouTube embeds and Lottie files from LottieFiles are the only outside resources a presentation uses.
const CSP = [
  "default-src 'self'",
  "script-src 'self' 'unsafe-inline'",
  "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com",
  "font-src 'self' https://fonts.gstatic.com data:",
  "connect-src 'self' blob: data: https://lottie.host https://*.lottiefiles.com",
  "img-src 'self' data: blob: https:",
  "media-src 'self' data: blob: https:",
  "frame-src https://www.youtube-nocookie.com https://www.youtube.com",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'self'",
  "frame-ancestors 'self'",
].join("; ");

const httpServer = createServer(async (req, res) => {
  if (!req.url) return json(res, 400, { error: "Missing URL" });
  const url = new URL(req.url, `http://${req.headers.host ?? "localhost"}`);

  const file = req.method === "GET" || req.method === "HEAD" ? staticFiles.get(url.pathname === "/index.html" ? "/" : url.pathname) : null;
  if (file) {
    if (url.pathname === "/") getSession(req, res);
    if (req.headers["if-none-match"] === file.etag) return res.writeHead(304).end();
    const html = file.type.startsWith("text/html");
    res.writeHead(200, {
      "content-type": file.type,
      etag: file.etag,
      "cache-control": url.pathname.startsWith("/assets/") ? "public, max-age=31536000, immutable" : html ? "no-store" : "no-cache",
      "x-content-type-options": "nosniff",
      ...(html ? { "content-security-policy": CSP, "referrer-policy": "same-origin", "permissions-policy": "camera=(self), microphone=(self), display-capture=(self), geolocation=()" } : {}),
      ...(req.headers["x-forwarded-proto"] === "https" ? { "strict-transport-security": "max-age=31536000" } : {}),
    }).end(req.method === "HEAD" ? undefined : file.body);
    return;
  }

  if ((url.pathname === "/healthz" || url.pathname === "/readyz") && req.method === "GET") {
    return json(res, 200, { name: "HTML PPTX Studio", version: APP_VERSION, commit: APP_COMMIT || null, startedAt: STARTED_AT.toISOString(), status: "ok", mode: "standalone-codex-app-server" });
  }

  const stateless = STATELESS_API.has(url.pathname);
  if (url.pathname.startsWith("/api/") && (req.method === "POST" || req.method === "PUT") && !stateless && !sameOrigin(req)) {
    return json(res, 403, { error: "Origin mismatch" });
  }

  // 共同編集 (server/rooms.mjs): create a room, listen to it, post changes and presence, share media.
  if (url.pathname === "/api/rooms" || url.pathname.startsWith("/api/rooms/")) return roomRoute(req, res, url);

  if (url.pathname === "/api/client-error" && req.method === "POST") {
    // Browser-side failures end up in the server log, where they can actually be seen.
    if (clientErrorBudget() && sameOrigin(req)) {
      try {
        const report = await readJson(req, 8_000);
        console.error("[client]", String(report.message ?? "").slice(0, 300), String(report.where ?? "").slice(0, 200), String(report.version ?? "").slice(0, 20));
      } catch { /* ignore malformed reports */ }
    }
    return res.writeHead(204).end();
  }

  // PowerPoint への書き出し as a job the studio follows: start it, ask how far it got, fetch the file.
  if (url.pathname === "/api/export/pptx/jobs" && req.method === "POST") {
    const session = getSession(req, res);
    try {
      const body = await readJson(req, 200_000_000);
      if (!body?.deck || !Array.isArray(body.deck.slides) || !body.deck.slides.length) return json(res, 400, { error: "書き出す資料がありません。" });
      const job = startExportJob(body.deck, { owner: session.id, base: `http://127.0.0.1:${PORT}`, rootDir: here });
      return json(res, 202, publicExportJob(job));
    } catch (error) {
      return json(res, 422, { error: error.message });
    }
  }
  const exportMatch = url.pathname.match(/^\/api\/export\/pptx\/jobs\/([0-9a-f-]{36})(\/file)?$/);
  if (exportMatch && (req.method === "GET" || req.method === "DELETE")) {
    const session = getSession(req, res);
    const job = exportJob(exportMatch[1], session.id);
    if (!job) return json(res, 404, { error: "書き出しが見つかりません。もう一度書き出してください。" });
    if (req.method === "DELETE") {
      if (!sameOrigin(req)) return json(res, 403, { error: "Origin mismatch" });
      if (!job.finished) job.controller.abort("取り消しました。");
      return json(res, 200, publicExportJob(job));
    }
    if (!exportMatch[2]) return json(res, 200, publicExportJob(job));
    if (job.status !== "completed") return json(res, 409, { error: "まだ書き出し中です。" });
    const name = encodeURIComponent(`${String(job.title || "資料").replace(/[\\/:*?"<>|]/g, "_").slice(0, 80)}.pptx`);
    res.writeHead(200, {
      "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      "content-disposition": `attachment; filename*=UTF-8''${name}`,
      "content-length": job.bytes.length,
      "x-export-stats": JSON.stringify(job.stats ?? {}),
      "cache-control": "no-store",
    });
    return res.end(job.bytes);
  }

  // PowerPoint への書き出し in one request (scripts): the deck (media inside) → an editable .pptx.
  if (url.pathname === "/api/export/pptx" && req.method === "POST") {
    try {
      const body = await readJson(req, 200_000_000);
      if (!body?.deck || !Array.isArray(body.deck.slides) || !body.deck.slides.length) return json(res, 400, { error: "書き出す資料がありません。" });
      const { bytes, stats } = await exportPptx(body.deck, { base: `http://127.0.0.1:${PORT}`, rootDir: here });
      const name = encodeURIComponent(`${String(body.deck.title || "資料").replace(/[\\/:*?"<>|]/g, "_").slice(0, 80)}.pptx`);
      res.writeHead(200, {
        "content-type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        "content-disposition": `attachment; filename*=UTF-8''${name}`,
        "content-length": bytes.length,
        "x-export-stats": JSON.stringify(stats),
        "cache-control": "no-store",
      });
      return res.end(bytes);
    } catch (error) {
      return json(res, 422, { error: error.message });
    }
  }

  if (url.pathname === "/api/import" && req.method === "POST") {
    const name = String(url.searchParams.get("name") ?? "");
    if (!/\.(pptx|docx|pdf)$/i.test(name)) return json(res, 400, { error: "取り込めるのは PowerPoint（.pptx）・PDF・Word（.docx）です。" });
    try {
      const buffer = await readBody(req, 30_000_000, "ファイルが大きすぎます（30MBまで）。");
      // "exact": a PowerPoint deck as it looks (objects on 白紙 pages); otherwise rebuilt into the studio's layouts.
      const mode = url.searchParams.get("mode") === "exact" && /\.pptx$/i.test(name) ? "exact" : "layout";
      return json(res, 200, await withExtractSlot(() => importDeck(buffer, name.toLowerCase(), { rootDir: here, mode })));
    } catch (error) {
      return json(res, 422, { error: error.message });
    }
  }

  // 挿入 → 表 → Excel／CSV から: a workbook's sheets as table cells (formulas kept, numbers shown as Excel shows them).
  if (url.pathname === "/api/tables/xlsx" && req.method === "POST") {
    const name = String(url.searchParams.get("name") ?? "");
    if (!/\.(xlsx|xlsm)$/i.test(name)) return json(res, 400, { error: "読み込めるのはExcel（.xlsx）です。古い形式（.xls）は「CSV UTF-8」か「Excelブック」で保存し直してください。" });
    try {
      const buffer = await readBody(req, 20_000_000, "ファイルが大きすぎます（20MBまで）。");
      return json(res, 200, await withExtractSlot(() => importSheets(buffer, name.toLowerCase(), { rootDir: here })));
    } catch (error) {
      return json(res, 422, { error: error.message });
    }
  }

  if (url.pathname === "/api/extract" && req.method === "POST") {
    const name = String(url.searchParams.get("name") ?? "");
    if (!/\.(pptx|docx|pdf)$/i.test(name)) return json(res, 400, { error: "pptx・docx・pdfのみ読み込めます。" });
    try {
      const buffer = await readBody(req, 15_000_000, "ファイルが大きすぎます（15MBまで）。");
      return json(res, 200, await withExtractSlot(() => extractText(buffer, name.toLowerCase(), { rootDir: here })));
    } catch (error) {
      return json(res, 422, { error: error.message });
    }
  }

  if (url.pathname === "/api/codex/status" && req.method === "GET") {
    const session = getSession(req, res);
    try {
      const result = await codex.accountRead();
      codexReady = Boolean(result?.account);
      if (!TEAM_MODE && !result?.account && authorized(session)) {
        session.authorized = false;
        ownerSessionId = null;
        persistOwnerSession();
      }
      const visible = TEAM_MODE ? session.member : authorized(session);
      const pendingLogin = !authorized(session) ? codex.getPendingLogin(session.id) : null;
      return json(res, 200, {
        available: true,
        fallback: await localAi.status(),
        team: TEAM_MODE,
        member: TEAM_MODE ? Boolean(session.member) : null,
        authenticated: authorized(session),
        authorized: authorized(session),
        accountType: result?.account?.type ?? null,
        planType: visible ? result?.account?.planType ?? null : null,
        loginError: session.loginError ?? null,
        pendingLogin: pendingLogin ? { userCode: pendingLogin.userCode, verificationUrl: pendingLogin.verificationUrl } : null,
      });
    } catch (error) {
      // Codex itself will not start; the page can still work through the stand-by AI when it is here.
      const fallback = await localAi.status();
      if (fallback.available) return json(res, 200, { available: false, authenticated: false, authorized: false, error: error.message, fallback });
      return json(res, 503, { available: false, authenticated: false, error: error.message, fallback });
    }
  }

  if (url.pathname === "/api/auth/passcode" && req.method === "POST") {
    const session = getSession(req, res);
    if (!TEAM_MODE) return json(res, 404, { error: "Not found" });
    if (passcodeRateLimited(req)) return json(res, 429, { error: "試行回数が多すぎます。10分ほど待ってからお試しください。" });
    try {
      const { passcode } = z.object({ passcode: z.string().max(200) }).parse(await readJson(req, 10_000));
      if (!passcodeMatches(passcode)) return json(res, 401, { error: "合言葉が違います。" });
      session.member = true;
      const secure = req.headers["x-forwarded-proto"] === "https" ? "; Secure" : "";
      res.setHeader("Set-Cookie", [
        `${SESSION_COOKIE}=${encodeURIComponent(session.id)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=2592000${secure}`,
        `${MEMBER_COOKIE}=${memberToken(session.id)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=2592000${secure}`,
      ]);
      return json(res, 200, { member: true });
    } catch (error) {
      return json(res, 400, { error: zodMessage(error) });
    }
  }

  if (url.pathname === "/api/codex/login" && req.method === "POST") {
    const session = getSession(req, res);
    if (TEAM_MODE && !session.member) return json(res, 401, { error: "先に合言葉を入力してください。" });
    if (!TEAM_MODE && ownerSessionId && ownerSessionId !== session.id) {
      return json(res, 423, { error: "このアプリは別のブラウザで使用中です。" });
    }
    try {
      const account = await codex.accountRead();
      if (TEAM_MODE && account?.account) {
        codexReady = true;
        return json(res, 200, { alreadyConnected: true });
      }
      if (account?.account && !authorized(session)) await codex.accountLogout();
      session.loginError = null;
      const login = await codex.startDeviceLogin(session.id);
      return json(res, 200, login);
    } catch (error) {
      session.loginError = error.message;
      return json(res, 503, { error: error.message });
    }
  }

  const aiRoute = async (handler, maxBytes = 15_000_000) => {
    const session = getSession(req, res);
    if (!(await aiAllowed(session))) return json(res, 401, { error: "先にCodexへ接続してください。" });
    try {
      const job = await handler(session, await readJson(req, maxBytes));
      return json(res, 202, { jobId: job.id });
    } catch (error) {
      return json(res, error.status || 400, { error: zodMessage(error), ...(error.code ? { code: error.code } : {}) });
    }
  };

  if (url.pathname === "/api/decks" && req.method === "POST") {
    return aiRoute((session, body) => {
      const input = deckRequestSchema.parse(body);
      // Writing from an agreed outline keeps its slide count.
      if (input.outline) input.settings = { ...input.settings, slideCount: input.outline.length };
      return startDeckJob(session, input);
    }, 500_000);
  }

  if (url.pathname === "/api/decks/outline" && req.method === "POST") {
    return aiRoute((session, body) => startOutlineJob(session, outlineRequestSchema.parse(body)), 500_000);
  }

  if (url.pathname === "/api/decks/rewrite" && req.method === "POST") {
    return aiRoute((session, body) => {
      const request = rewriteRequestSchema.parse(body);
      const length = request.deck.slides.length;
      const settings = {
        slideCount: length,
        textDensity: request.settings.textDensity,
        ...(request.settings.flexibleCount ? { minCount: Math.max(2, length - 3), maxCount: Math.min(50, length + 3) } : {}),
      };
      const input = { deckTitle: request.deck.title, brief: request.instruction, audience: request.deck.audience, purpose: request.deck.purpose, settings };
      return startDeckJob(session, input, buildRewritePrompt({ deck: request.deck, instruction: request.instruction, settings }),
        slidesSource(request.deck.slides, request.deck.title, request.instruction, request.deck.memo), { variety: varietyIssues(request.deck.slides).length });
    });
  }

  if (url.pathname === "/api/decks/translate" && req.method === "POST") {
    return aiRoute((session, body) => startTranslateJob(session, translateRequestSchema.parse(body)));
  }
  if (url.pathname === "/api/decks/synonyms" && req.method === "POST") {
    return aiRoute((session, body) => startSynonymsJob(session, synonymsRequestSchema.parse(body)));
  }
  if (url.pathname === "/api/decks/notes" && req.method === "POST") {
    return aiRoute((session, body) => startNotesJob(session, notesRequestSchema.parse(body)));
  }

  if (url.pathname === "/api/decks/chat" && req.method === "POST") {
    return aiRoute((session, body) => {
      const request = chatRequestSchema.parse(body);
      if (asksForImageGeneration(request.message)) throw Object.assign(new Error("画像生成の依頼は専用の画像生成機能で処理してください。"), { status: 422, code: "use_image_generation" });
      return startChatJob(session, request);
    });
  }

  if (url.pathname === "/api/decks/image" && req.method === "POST") {
    const session = getSession(req, res);
    if (!authorized(session)) return json(res, 401, { error: "画像生成にはCodexへの接続が必要です。" });
    try {
      const request = z.object({ deck: deckShape, message: z.string().trim().min(1).max(2000), index: z.number().int().min(0) }).parse(await readJson(req, 15_000_000));
      if (request.index >= request.deck.slides.length) return json(res, 422, { error: "指定したスライドが存在しません。" });
      const job = startImageJob(session, request);
      return json(res, 202, { jobId: job.id });
    } catch (error) { return json(res, 400, { error: zodMessage(error) }); }
  }

  if (url.pathname === "/api/decks/variants" && req.method === "POST") {
    return aiRoute((session, body) => {
      const request = variantsRequestSchema.parse(body);
      if (request.slideIndex >= request.deck.slides.length) throw new Error("slideIndexが範囲外です。");
      return startVariantsJob(session, request);
    });
  }

  if (url.pathname === "/api/decks/revise" && req.method === "POST") {
    return aiRoute((session, body) => {
      const request = reviseRequestSchema.parse(body);
      const total = request.deck.slides.length;
      const valid = request.mode === "insert" ? request.slideIndex >= 1 && request.slideIndex <= total - 1 : request.slideIndex < total;
      if (!valid) throw new Error("slideIndexが範囲外です。");
      if (request.mode === "insert" && total >= 500) throw new Error("スライドは500枚までです。");
      return startReviseJob(session, request);
    });
  }

  const imageJobMatch = url.pathname.match(/^\/api\/decks\/([0-9a-f-]{36})\/image$/);
  if (imageJobMatch && req.method === "GET") {
    const session = getSession(req, res);
    if (jobOwners.get(imageJobMatch[1]) !== session.id) return json(res, 404, { error: "Not found" });
    const bytes = codex.getImage(imageJobMatch[1]);
    const job = codex.getJob(imageJobMatch[1]);
    if (!bytes || !job?.image) return json(res, 404, { error: "画像が見つかりません。" });
    res.writeHead(200, { "content-type": job.image.mime, "content-length": bytes.length, "cache-control": "private, no-store", "x-content-type-options": "nosniff" }).end(bytes);
    return;
  }
  const deckJobMatch = url.pathname.match(/^\/api\/decks\/([0-9a-f-]{36})(?:\/events)?$/);
  if (deckJobMatch && req.method === "GET") {
    const session = getSession(req, res);
    const jobId = deckJobMatch[1];
    if (jobOwners.get(jobId) !== session.id) return json(res, 404, { error: "Not found" });
    const job = codex.getJob(jobId);
    if (!job) return json(res, 404, { error: "Not found" });
    if (!url.pathname.endsWith("/events")) return json(res, 200, job);
    res.writeHead(200, {
      "content-type": "text/event-stream; charset=utf-8",
      "cache-control": "no-store",
      connection: "keep-alive",
      "x-accel-buffering": "no",
    });
    sendEvent(res, "progress", job);
    if (["completed", "failed"].includes(job.status)) return res.end();
    if (!streams.has(jobId)) streams.set(jobId, new Set());
    streams.get(jobId).add(res);
    const keepAlive = setInterval(() => res.write(": keepalive\n\n"), 20_000);
    req.on("close", () => {
      clearInterval(keepAlive);
      streams.get(jobId)?.delete(res);
      if (streams.get(jobId)?.size === 0) streams.delete(jobId);
    });
    return;
  }

  return json(res, 404, { error: "Not found" });
});

// ---------------------------------------------------------------- 共同編集 (co-editing rooms)

const rooms = createRooms();
setInterval(() => rooms.heartbeat(), 20_000).unref();

/** With a team passcode, only members share decks; otherwise the unguessable link is the key. */
function roomAllowed(req, res) {
  if (!TEAM_MODE) return true;
  const session = getSession(req, res);
  return Boolean(session?.member);
}

async function roomRoute(req, res, url) {
  const parts = url.pathname.split("/").filter(Boolean); // api, rooms, id?, action?, mid?
  const action = parts[3] || "";
  // オンライン プレゼンテーション: anyone with the link may watch a show (its stream and its pictures), as in
  // PowerPoint; making, editing and presenting still need the team passcode in team mode.
  const watching = rooms.get(parts[2])?.kind === "show" && (((req.method === "GET" || req.method === "HEAD") && (!action || action === "media")) || ["leave", "react", "vote"].includes(action));
  if (!watching && !roomAllowed(req, res)) return json(res, 403, { error: "チームのパスコードを入れると共同編集できます" });
  try {
    if (parts.length === 2 && req.method === "POST") {
      const body = await readJson(req, 25_000_000);
      const room = rooms.create(body.deck, { id: typeof body.id === "string" ? body.id : null, show: body.show === true, key: typeof body.key === "string" ? body.key : null, fits: body.fits });
      return json(res, 200, { id: room.id, version: room.version, ...(room.kind === "show" ? { key: room.key, show: room.show } : {}) });
    }
    const room = rooms.get(parts[2]);
    if (!room) return json(res, 404, { error: "共同編集の部屋が見つかりません（時間が経って閉じられたか、サーバーが再起動しました）", gone: true });
    if (!action && req.method === "GET") {
      res.writeHead(200, { "content-type": "text/event-stream; charset=utf-8", "cache-control": "no-store", connection: "keep-alive", "x-accel-buffering": "no" });
      const client = String(url.searchParams.get("client") || "");
      rooms.join(room, res, { client, name: url.searchParams.get("name"), uid: url.searchParams.get("uid") });
      req.on("close", () => rooms.leave(room, client, res));
      return;
    }
    if (action === "ops" && req.method === "POST") {
      const body = await readJson(req, 25_000_000);
      return json(res, 200, { version: rooms.apply(room, String(body.client || ""), body.ops) });
    }
    if (action === "show" && req.method === "POST") {
      const body = await readJson(req, 4_000);
      return json(res, 200, { show: rooms.setShow(room, body.key, body) });
    }
    if (action === "react" && req.method === "POST") {
      const body = await readJson(req, 2_000);
      return json(res, 200, { ok: rooms.react(room, String(body.client || ""), String(body.emoji || "")) });
    }
    if (action === "vote" && req.method === "POST") {
      const body = await readJson(req, 2_000);
      return json(res, 200, { counts: rooms.vote(room, String(body.client || ""), { poll: String(body.poll || ""), option: body.option }) });
    }
    if (action === "caption" && req.method === "POST") {
      const body = await readJson(req, 4_000);
      rooms.caption(room, body.key, body.text);
      return res.writeHead(204).end();
    }
    if (action === "deck" && req.method === "PUT") {
      const body = await readJson(req, 25_000_000);
      return json(res, 200, { version: rooms.replaceDeck(room, body.key, body.deck, body.fits) });
    }
    if (action === "presence" && req.method === "POST") {
      const body = await readJson(req, 20_000);
      rooms.presence(room, String(body.client || ""), body);
      return res.writeHead(204).end();
    }
    if (action === "leave" && req.method === "POST") {
      const body = await readJson(req, 2_000);
      rooms.leave(room, String(body.client || ""));
      return res.writeHead(204).end();
    }
    if (action === "media" && parts[4]) {
      const mid = decodeURIComponent(parts[4]);
      if (req.method === "PUT") {
        const body = await readBody(req, 200_000_000, "ファイルが大きすぎます（200MBまで）。");
        rooms.putMedia(room, mid, req.headers["content-type"], body);
        return res.writeHead(204).end();
      }
      if (req.method === "GET" || req.method === "HEAD") {
        const media = rooms.getMedia(room, mid);
        if (!media) return json(res, 404, { error: "ありません" });
        res.writeHead(200, { "content-type": media.type, "content-length": media.body.length, "cache-control": "private, max-age=3600", "x-content-type-options": "nosniff" });
        return res.end(req.method === "HEAD" ? undefined : media.body);
      }
    }
    return json(res, 404, { error: "Not found" });
  } catch (error) {
    return json(res, 400, { error: error.message });
  }
}

httpServer.listen(PORT, process.env.HOST ?? "0.0.0.0", () => {
  console.log(`HTML PPTX Studio: http://localhost:${PORT}`);
});

for (const signal of ["SIGTERM", "SIGINT"]) {
  process.once(signal, () => {
    codex.close();
    // Open streams (co-editing rooms, job progress) would hold the server open: end them so it can stop now.
    rooms.closeAll();
    httpServer.close(() => process.exit(0));
    httpServer.closeAllConnections?.();
    setTimeout(() => process.exit(0), 3000).unref();
  });
}
