import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { createServer } from "node:http";
import { existsSync } from "node:fs";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

import { shortDeck } from "./fixtures.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
const PHOTO = join(root, "public", "assets", "data-insight.jpg");
let nextPort = 8800 + Math.floor(Math.random() * 400);

async function startServer(extraEnv = {}) {
  const port = nextPort++;
  const runtimeDir = await mkdtemp(join(tmpdir(), "html-studio-test-"));
  const child = spawn(process.execPath, ["server.js"], {
    cwd: root,
    env: { ...process.env, PORT: String(port), HOST: "127.0.0.1", CODEX_RUNTIME_DIR: runtimeDir, LOCAL_AI_URL: "http://127.0.0.1:9", RENDER: "", STUDIO_PASSCODE: "", ...extraEnv },
    stdio: ["ignore", "pipe", "pipe"],
  });
  const [chunk] = await once(child.stdout, "data");
  assert.match(String(chunk), /HTML PPTX Studio:/);
  const base = `http://127.0.0.1:${port}`;
  let cookie = "";
  const request = async (path, options = {}) => {
    const response = await fetch(`${base}${path}`, { ...options, headers: { ...(cookie ? { cookie } : {}), ...(options.headers || {}) } });
    const setCookie = response.headers.get("set-cookie");
    if (setCookie) cookie = setCookie.split(";")[0];
    return response;
  };
  const postJson = (path, body) => request(path, { method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(body) });
  const stop = async () => {
    child.kill();
    await once(child, "exit");
    await rm(runtimeDir, { recursive: true, force: true });
  };
  return { base, request, postJson, stop, runtimeDir };
}

async function waitForJob(server, jobId, timeoutMs = 20_000) {
  const started = Date.now();
  while (Date.now() - started < timeoutMs) {
    const job = await (await server.request(`/api/decks/${jobId}`)).json();
    if (["completed", "failed"].includes(job.status)) return job;
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error("job timed out");
}

/** A stand-in for Ollama serving gemma4:12b: answers each chat with the next scripted JSON, streamed. */
async function withFakeOllama(answers, fn) {
  const requests = [];
  const server = createServer(async (req, res) => {
    if (req.url === "/api/tags") return res.end(JSON.stringify({ models: [{ name: "gemma4:12b" }] }));
    if (req.url === "/api/chat" && req.method === "POST") {
      let body = "";
      for await (const chunk of req) body += chunk;
      requests.push(JSON.parse(body));
      const text = JSON.stringify(answers[Math.min(requests.length - 1, answers.length - 1)]);
      res.writeHead(200, { "content-type": "application/x-ndjson" });
      for (const piece of text.match(/[\s\S]{1,300}/g)) res.write(`${JSON.stringify({ message: { role: "assistant", content: piece }, done: false })}\n`);
      return res.end(`${JSON.stringify({ message: { role: "assistant", content: "" }, done: true })}\n`);
    }
    res.writeHead(404).end();
  });
  server.listen(0, "127.0.0.1");
  await once(server, "listening");
  try {
    await fn(`http://127.0.0.1:${server.address().port}`, requests);
  } finally {
    server.close();
  }
}

async function withFakeCodex(answers, fn, env = {}) {
  const dir = await mkdtemp(join(tmpdir(), "fake-codex-"));
  const scriptPath = join(dir, "script.json");
  const promptPath = join(dir, "prompts.json");
  await writeFile(scriptPath, JSON.stringify(answers));
  const server = await startServer({ CODEX_BIN: join(root, "test", "fake-codex.mjs"), FAKE_CODEX_SCRIPT: scriptPath, FAKE_CODEX_PROMPTS: promptPath, ...env });
  try {
    await server.request("/");
    const login = await server.postJson("/api/codex/login", {});
    assert.equal(login.status, 200);
    await new Promise((resolve) => setTimeout(resolve, 150));
    const status = await (await server.request("/api/codex/status")).json();
    assert.equal(status.authorized, true);
    await fn(server, async () => JSON.parse(await readFile(promptPath, "utf8")), async () => JSON.parse(await readFile(`${promptPath}.turns`, "utf8")));
  } finally {
    await server.stop();
    await rm(dir, { recursive: true, force: true });
  }
}

test("device login can be resumed after a reload and its code can be replaced", async () => {
  const dir = await mkdtemp(join(tmpdir(), "fake-codex-login-"));
  const scriptPath = join(dir, "script.json");
  await writeFile(scriptPath, "[]");
  const server = await startServer({ CODEX_BIN: join(root, "test", "fake-codex.mjs"), FAKE_CODEX_SCRIPT: scriptPath, FAKE_CODEX_LOGIN_DELAY_MS: "1000" });
  try {
    const first = await (await server.postJson("/api/codex/login", {})).json();
    assert.equal(first.userCode, "TEST-CODE-1");
    const pending = await (await server.request("/api/codex/status")).json();
    assert.equal(pending.pendingLogin.userCode, first.userCode);
    assert.equal(pending.authorized, false);

    const second = await (await server.postJson("/api/codex/login", {})).json();
    assert.equal(second.userCode, "TEST-CODE-2");
    assert.equal((await (await server.request("/api/codex/status")).json()).pendingLogin.userCode, second.userCode);
    await new Promise((resolve) => setTimeout(resolve, 1100));
    const connected = await (await server.request("/api/codex/status")).json();
    assert.equal(connected.authorized, true);
    assert.equal(connected.pendingLogin, null);
  } finally {
    await server.stop();
    await rm(dir, { recursive: true, force: true });
  }
});

test("a failed device login returns the Codex error and allows another attempt", async () => {
  const dir = await mkdtemp(join(tmpdir(), "fake-codex-login-"));
  const scriptPath = join(dir, "script.json");
  await writeFile(scriptPath, "[]");
  const server = await startServer({ CODEX_BIN: join(root, "test", "fake-codex.mjs"), FAKE_CODEX_SCRIPT: scriptPath, FAKE_CODEX_LOGIN_ERROR: "Device code expired" });
  try {
    assert.equal((await server.postJson("/api/codex/login", {})).status, 200);
    await new Promise((resolve) => setTimeout(resolve, 100));
    const failed = await (await server.request("/api/codex/status")).json();
    assert.equal(failed.authorized, false);
    assert.equal(failed.pendingLogin, null);
    assert.equal(failed.loginError, "Device code expired");
    const retry = await (await server.postJson("/api/codex/login", {})).json();
    assert.equal(retry.userCode, "TEST-CODE-2");
    assert.equal((await (await server.request("/api/codex/status")).json()).loginError, null);
  } finally {
    await server.stop();
    await rm(dir, { recursive: true, force: true });
  }
});

test("Codex reconnects after its app-server exits during startup", async () => {
  const dir = await mkdtemp(join(tmpdir(), "fake-codex-restart-"));
  const scriptPath = join(dir, "script.json");
  const marker = join(dir, "exited-once");
  await writeFile(scriptPath, "[]");
  const server = await startServer({ CODEX_BIN: join(root, "test", "fake-codex.mjs"), FAKE_CODEX_SCRIPT: scriptPath, FAKE_CODEX_INIT_EXIT_ONCE_FILE: marker });
  try {
    assert.equal((await server.request("/api/codex/status")).status, 503);
    const retried = await server.request("/api/codex/status");
    assert.equal(retried.status, 200);
    assert.equal((await retried.json()).available, true);
  } finally {
    await server.stop();
    await rm(dir, { recursive: true, force: true });
  }
});

test("app shell, engine, assets and security headers", async () => {
  const server = await startServer();
  try {
    const health = await (await server.request("/healthz")).json();
    assert.equal(health.status, "ok");
    assert.equal(health.name, "HTML PPTX Studio");
    const pkg = JSON.parse(await readFile(join(root, "package.json"), "utf8"));
    assert.equal(health.version, pkg.version);
    assert.match(health.commit ?? "", /^([0-9a-f]{7})?$/, "the running commit, when the checkout has one");
    assert.ok(!Number.isNaN(Date.parse(health.startedAt)));

    const page = await server.request("/");
    assert.equal(page.headers.get("referrer-policy"), "same-origin");
    const csp = page.headers.get("content-security-policy");
    assert.match(csp, /object-src 'none'/);
    assert.match(csp, /frame-src https:\/\/www\.youtube-nocookie\.com/);
    assert.match(csp, /font-src 'self' https:\/\/fonts\.gstatic\.com/);
    const html = await page.text();
    for (const id of ["briefInput", "askChatGptBtn", "createThemes", "filmstrip", "inspector", "chatPane", "designDialog", "mediaUrlDialog", "presenter", "measureRoot", "downloadBtn", "historyDialog"]) {
      assert.match(html, new RegExp(`id="${id}"`), `missing #${id}`);
    }
    assert.doesNotMatch(html, /__APP_VERSION__|__APP_BUILD__/);
    assert.match(html, new RegExp(`<small title="バージョン ${pkg.version.replace(/\./g, "\\.")}[^"]*サーバー起動 [^"]+">v${pkg.version.replace(/\./g, "\\.")}</small>`), "the header shows the version and, on hover, the running build");
    assert.equal((await server.postJson("/api/client-error", { message: "test", where: "unit" })).status, 204);

    const engine = await (await server.request("/engine/engine.js")).text();
    assert.doesNotMatch(engine, /\/\*__ICONS__\*\/\{\}/, "icons are injected into the engine");
    assert.match(engine, /"bulb":\{"label":/);
    assert.equal((await server.request("/engine/motion.js")).status, 200);
    // 3D models: models.js, and three.js with the glTF loader's imports made relative (no import map needed).
    assert.match(await (await server.request("/engine/models.js")).text(), /mountModels/);
    const loader = await (await server.request("/vendor/three/GLTFLoader.js")).text();
    assert.match(loader, /from '\.\/three\.module\.js'/);
    assert.match(loader, /from '\.\/BufferGeometryUtils\.js'/);
    assert.doesNotMatch(loader.replace(/\/\*[\s\S]*?\*\//g, ""), /from ['"]three['"]/);
    for (const name of ["three.module.js", "three.core.js", "BufferGeometryUtils.js", "SkeletonUtils.js"]) assert.equal((await server.request(`/vendor/three/${name}`)).status, 200, name);
    assert.match(await (await server.request("/")).text(), /\/engine\/models\.js/);
    assert.equal((await server.request("/saved-library.js")).status, 200);
    assert.match(await (await server.request("/engine/engine.css")).text(), /\.hs-slide/);
    const app = await (await server.request("/app.js")).text();
    assert.doesNotMatch(app, /__APP_VERSION__/);
    const looks = await server.request("/layout-looks.mjs");
    assert.equal(looks.status, 200, "the studio and the server share one list of layout looks");
    assert.match(looks.headers.get("content-type"), /javascript/);
    assert.match(await looks.text(), /export function varietyIssues/);
    const autoImages = await server.request("/auto-images.mjs");
    assert.equal(autoImages.status, 200);
    assert.match(await autoImages.text(), /export function picturePlan/);

    const sample = await (await server.request("/samples/ai-rollout.json")).json();
    assert.ok(sample.slideData.length >= 40);
    const image = await server.request("/assets/data-insight.jpg");
    assert.equal(image.status, 200);
    assert.equal(image.headers.get("content-type"), "image/jpeg");
    assert.equal((await server.request("/assets/../server.js")).status, 404);
    assert.equal((await server.request("/server.js")).status, 404);

    assert.equal((await server.postJson("/api/decks", {})).status, 401, "AI needs a Codex login");
    const foreign = await server.request("/api/decks/chat", { method: "POST", headers: { "content-type": "application/json", origin: "https://evil.example" }, body: "{}" });
    assert.equal(foreign.status, 403);
  } finally {
    await server.stop();
  }
});

test("Codex generation repairs a wrong slide count and text that cannot fit", async () => {
  const overflowing = shortDeck(5);
  overflowing.slideData[1].takeaway = "作業時間を大きく削減しながら売上と品質指標を維持し、全店展開に向けた体制整備と予算確保を同時に進める必要があるという結論を、根拠となる数値と現場の声、今後の課題とあわせて示す";
  await withFakeCodex([shortDeck(4), overflowing, shortDeck(5)], async (server, prompts) => {
    const response = await server.postJson("/api/decks", { brief: "テスト", settings: { slideCount: 5, textDensity: "standard" } });
    assert.equal(response.status, 202);
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.attempt, 3);
    assert.equal(job.deck.slideData.length, 5);
    const sent = await prompts();
    assert.match(sent[0], /指定枚数: 5枚/);
    assert.match(sent[0], /HTML/, "the prompt is written for moving HTML presentations");
    assert.match(sent[1], /5枚/);
    assert.match(sent[2], /収まらない/);
    assert.match(sent[2], /takeaway/);
  });
});

test("a figure the material does not contain is sent back once, and reported if it stays", async () => {
  const invented = shortDeck(4);
  invented.slideData[1].points = ["作業時間を月120時間削減", "発注時間を43%短縮"];
  const fixed = shortDeck(4);
  fixed.slideData[1].points = ["作業時間を月120時間削減", "発注時間を42%短縮"];
  await withFakeCodex([invented, fixed], async (server, prompts) => {
    const brief = "モデル30店で1店あたり月120時間の作業を削減。発注時間は42%短縮。";
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", { brief, settings: { slideCount: 4, textDensity: "standard" } })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.attempt, 2);
    assert.deepEqual(job.deck.slideData[1].points, fixed.slideData[1].points);
    const [, retry] = await prompts();
    assert.match(retry, /素材のどこにもありません/);
    assert.match(retry, /2枚目「本文1」の points\[1\]: 43%/);
  });
  await withFakeCodex([invented, invented, invented], async (server) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", { brief: "月120時間の削減", settings: { slideCount: 4, textDensity: "standard" } })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.match(job.detail, /素材にない数値が1か所あります（2枚目「43%」）/);
    assert.ok(job.issues.some((issue) => issue.kind === "unsourced" && issue.field === "points[1]"));
  });
});

test("Codex errors: a retried hiccup is waited out, a busy server is asked again, a usage limit is explained", async () => {
  const input = { brief: "テスト", settings: { slideCount: 4, textDensity: "standard" } };
  const busy = { __fail: { message: "server overloaded", codexErrorInfo: "serverOverloaded" } };
  await withFakeCodex([{ __hiccup: shortDeck(4) }], async (server) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", input)).json()).jobId);
    assert.equal(job.status, "completed", job.error);
  });
  await withFakeCodex([busy, shortDeck(4)], async (server, prompts) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", input)).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    const sent = await prompts();
    assert.equal(sent.length, 2);
    assert.equal(sent[1], sent[0], "the same request is sent again");
  });
  await withFakeCodex([busy, busy], async (server) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", input)).json()).jobId);
    assert.equal(job.status, "failed");
    assert.match(job.error, /混み合っている/);
  });
  await withFakeCodex([{ __fail: { message: "You've hit your usage limit.", codexErrorInfo: "usageLimitExceeded" } }], async (server) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks/outline", input)).json()).jobId);
    assert.equal(job.status, "failed");
    assert.match(job.error, /Codexの利用上限に達しました.*You've hit your usage limit/);
  });
});

test("stand-by AI: when Codex is cut off or not connected, gemma4:12b makes the deck; the web says it has none", async () => {
  const input = { brief: "テスト", settings: { slideCount: 4, textDensity: "standard" } };
  await withFakeOllama([shortDeck(4)], async (url, requests) => {
    await withFakeCodex([{ __fail: { message: "You've hit your usage limit.", codexErrorInfo: "usageLimitExceeded" } }], async (server, prompts) => {
      const status = await (await server.request("/api/codex/status")).json();
      assert.deepEqual(status.fallback, { available: true, model: "gemma4:12b", web: false, reason: "" });
      const job = await waitForJob(server, (await (await server.postJson("/api/decks", input)).json()).jobId);
      assert.equal(job.status, "completed", job.error);
      assert.equal(job.provider, "local");
      assert.equal(job.localModel, "gemma4:12b");
      assert.equal(job.deck.slideData.length, 4);
      const [codexPrompt] = await prompts();
      assert.equal(requests[0].messages[0].content, codexPrompt, "Gemma gets the same request");
      assert.equal(requests[0].format.type, "object", "structured output with the same schema");
    }, { LOCAL_AI_URL: url });
  });
  await withFakeOllama([shortDeck(5), shortDeck(4)], async (url, requests) => {
    const server = await startServer({ LOCAL_AI_URL: url, CODEX_BIN: "/nonexistent/codex" });
    try {
      const status = await (await server.request("/api/codex/status")).json();
      assert.equal(status.fallback.available, true);
      const response = await server.postJson("/api/decks", input);
      assert.equal(response.status, 202);
      const job = await waitForJob(server, (await response.json()).jobId);
      assert.equal(job.status, "completed", job.error);
      assert.equal(job.provider, "local");
      assert.equal(requests.length, 2, "the wrong slide count was sent back once");
      assert.match(requests[1].messages.at(-1).content, /4枚/);
    } finally {
      await server.stop();
    }
  });
  const web = await startServer({ RENDER: "true", CODEX_BIN: "/nonexistent/codex" });
  try {
    const status = await (await web.request("/api/codex/status")).json();
    assert.equal(status.fallback.available, false);
    assert.equal(status.fallback.web, true);
    assert.equal((await web.postJson("/api/decks", input)).status, 401);
  } finally {
    await web.stop();
  }
});

test("single-slide revision keeps the user's video and click-to-open details", async () => {
  const media = { src: "idb:abc123", kind: "video", name: "現場.mp4", autoplay: true, loop: true, muted: true };
  await withFakeCodex([{ slide: { type: "kpi", title: "改善後", takeaway: "短い結論", items: [{ label: "指標", value: "10%" }], details: [{ target: "items[0]", title: "算出方法", text: "前後の作業時間を比較" }] } }], async (server, prompts) => {
    const slides = shortDeck(4).slideData;
    slides[1].media = media;
    const response = await server.postJson("/api/decks/revise", { deck: { title: "テスト", slides }, slideIndex: 1, instruction: "数字を強調" });
    assert.equal(response.status, 202, await response.clone().text());
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.slide.type, "kpi");
    assert.deepEqual(job.slide.media, media, "videos are the user's and survive AI edits");
    assert.equal(job.slide.details[0].target, "items[0]");
    const [prompt] = await prompts();
    assert.match(prompt, /2枚目のスライドだけを作り直して/);
    assert.match(prompt, /数字を強調/);
    assert.doesNotMatch(prompt, /idb:abc123/, "media addresses are not shown to the AI");
    assert.match(prompt, /\[写真・動画あり\]/);
  });
});

test("AI can insert a new slide between existing ones", async () => {
  await withFakeCodex([{ slide: { type: "closing", message: "誤り" } }, { slide: { type: "table", title: "比較表", takeaway: "差を示す", headers: ["項目", "値"], rows: [["A", "1"]] } }], async (server, prompts) => {
    const deck = { title: "テスト", slides: shortDeck(4).slideData };
    assert.equal((await server.postJson("/api/decks/revise", { deck, slideIndex: 0, mode: "insert", instruction: "x" })).status, 400);
    const response = await server.postJson("/api/decks/revise", { deck, slideIndex: 2, mode: "insert", instruction: "部門別の比較表" });
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.slide.type, "table");
    const sent = await prompts();
    assert.match(sent[0], /2枚目と3枚目の間に入れる/);
    assert.match(sent[1], /titleとclosing以外/);
  });
});

test("AI can rewrite the whole deck while keeping the slide count, theme and transition", async () => {
  const rewritten = shortDeck(4);
  rewritten.slideData[1].title = "簡潔版";
  await withFakeCodex([rewritten], async (server, prompts) => {
    const deck = { title: "テスト", audience: "役員", theme: "aurora", transition: "morph", motion: { entrance: "blur", hover: "focus", numbers: true, ambient: true }, slides: shortDeck(4).slideData };
    const response = await server.postJson("/api/decks/rewrite", { deck, instruction: "もっと簡潔に" });
    assert.equal(response.status, 202, await response.clone().text());
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.deck.slideData[1].title, "簡潔版");
    const [prompt] = await prompts();
    assert.match(prompt, /見直しの指示: もっと簡潔に/);
    assert.match(prompt, /枚数: 4枚/);
    assert.match(prompt, /"title":"本文1"/);
  });
});

test("team passcode mode lets several browsers share one Codex login", async () => {
  const dir = await mkdtemp(join(tmpdir(), "fake-codex-"));
  const scriptPath = join(dir, "script.json");
  await writeFile(scriptPath, JSON.stringify([shortDeck(5), shortDeck(5)]));
  const server = await startServer({
    STUDIO_PASSCODE: "発表2026",
    CODEX_AUTH_JSON: JSON.stringify({ tokens: { test: true } }),
    CODEX_BIN: join(root, "test", "fake-codex.mjs"),
    FAKE_CODEX_SCRIPT: scriptPath,
  });
  const browser = () => {
    const cookies = {};
    return async (path, body) => {
      const response = await fetch(`${server.base}${path}`, {
        method: body ? "POST" : "GET",
        headers: { "content-type": "application/json", cookie: Object.entries(cookies).map(([k, v]) => `${k}=${v}`).join("; ") },
        body: body ? JSON.stringify(body) : undefined,
      });
      for (const header of response.headers.getSetCookie()) {
        const [pair] = header.split(";");
        const [key, value] = pair.split("=");
        cookies[key] = value;
      }
      return response;
    };
  };
  try {
    const seeded = JSON.parse(await readFile(join(server.runtimeDir, ".codex", "auth.json"), "utf8"));
    assert.equal(seeded.tokens.test, true, "CODEX_AUTH_JSON seeds the Codex login");
    const pc = browser();
    await pc("/");
    let status = await (await pc("/api/codex/status")).json();
    assert.equal(status.team, true);
    assert.equal(status.member, false);
    assert.equal((await pc("/api/decks", { brief: "x", settings: { slideCount: 5 } })).status, 401);
    assert.equal((await pc("/api/auth/passcode", { passcode: "違う" })).status, 401);
    assert.equal((await pc("/api/auth/passcode", { passcode: "発表2026" })).status, 200);
    status = await (await pc("/api/codex/status")).json();
    assert.equal(status.member, true);
    assert.equal(status.authorized, true);
    const phone = browser();
    await phone("/");
    await phone("/api/auth/passcode", { passcode: "発表2026" });
    for (const client of [pc, phone]) {
      const response = await client("/api/decks", { brief: "テスト", settings: { slideCount: 5 } });
      assert.equal(response.status, 202);
      const { jobId } = await response.json();
      let job;
      for (let i = 0; i < 100; i += 1) {
        job = await (await client(`/api/decks/${jobId}`)).json();
        if (job.status === "completed") break;
        await new Promise((resolve) => setTimeout(resolve, 100));
      }
      assert.equal(job.status, "completed");
    }
    const stranger = browser();
    await stranger("/");
    assert.equal((await stranger("/api/codex/login", {})).status, 401, "non-members cannot touch the Codex login");
  } finally {
    await server.stop();
    await rm(dir, { recursive: true, force: true });
  }
});

test("AI speaker notes are written only for the requested slides", async () => {
  await withFakeCodex([{ notes: [{ slide: 2, text: "本文1の結論から申し上げます。" }, { slide: 4, text: "対象外" }] }], async (server, prompts) => {
    const deck = { title: "テスト", audience: "役員", slides: shortDeck(4).slideData };
    const response = await server.postJson("/api/decks/notes", { deck, indices: [1], seconds: 30, tone: "簡潔" });
    assert.equal(response.status, 202);
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.deepEqual(job.notes, [{ slide: 1, text: "本文1の結論から申し上げます。" }]);
    const [prompt] = await prompts();
    assert.match(prompt, /約30秒/);
    assert.match(prompt, /ノートを書くスライド: 2枚目/);
  });
});

test("translation returns one string per string, keeps empty ones, and asks again when the count is wrong", async () => {
  await withFakeCodex([{ texts: ["Sales", "AI use"] }, { texts: ["Sales", "", "AI use"] }], async (server, prompts) => {
    const response = await server.postJson("/api/decks/translate", { texts: ["売上", "", "AIの活用"], to: "en" });
    assert.equal(response.status, 202);
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.deepEqual(job.texts, ["Sales", "", "AI use"]);
    const [first] = await prompts();
    assert.match(first, /英語に翻訳/);
    assert.match(first, /"AIの活用"/);
    assert.equal((await server.postJson("/api/decks/translate", { texts: ["a"], to: "fr" })).status, 400, "only the offered languages");
  });
});

test("類義語: the AI's senses come back tidied (the word itself, repeats and long phrases out), and the request is checked", async () => {
  await withFakeCodex([{ senses: [{ label: "問題・課題", words: ["課題", "論点", " 懸案 ", "論点", ""] }, { label: "空", words: ["課題"] }, { label: "", words: ["障壁", "とても長い言い回しで言い換えているところです"] }] }], async (server, prompts) => {
    const response = await server.postJson("/api/decks/synonyms", { word: "課題", context: "最大の課題は人手不足です" });
    assert.equal(response.status, 202);
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.deepEqual(job.senses, [{ label: "問題・課題", words: ["論点", "懸案"] }, { label: "類義語", words: ["障壁"] }]);
    const [prompt] = await prompts();
    assert.match(prompt, /語: 課題/);
    assert.match(prompt, /最大の課題は人手不足です/);
    assert.equal((await server.postJson("/api/decks/synonyms", { word: "" })).status, 400, "a word is needed");
    assert.equal((await server.postJson("/api/decks/synonyms", { word: "あ".repeat(41) })).status, 400, "and a short one");
  });
});

test("chat proposes operations on the original numbering, keeps photos and videos, repairs overflow, and never switches the theme", async () => {
  const photo = `data:image/jpeg;base64,${(await readFile(PHOTO)).toString("base64")}`;
  const long = "とても長い説明文".repeat(13);
  await withFakeCodex([
    { reply: "2枚目を簡潔にしました。", operations: [{ op: "replace", slide: 2, content: { type: "content", title: "本文1", takeaway: long, points: ["A"] } }], suggestions: ["結論を強く"] },
    {
      reply: "2枚目を簡潔にし、4枚目の後ろに1枚追加し、3枚目を削除しました。テーマも変えようとしました。",
      operations: [
        { op: "replace", slide: 2, content: { type: "content", title: "本文1", takeaway: "短い結論", points: ["A"], animation: "click" } },
        { op: "insert", slide: 4, content: { type: "cards", title: "追加", takeaway: "追加の結論", items: [{ title: "X", desc: "x" }, { title: "Y", desc: "y" }], details: [{ target: "items[0]", text: "Xの根拠" }] } },
        { op: "delete", slide: 3 },
      ],
      theme: "midnight",
      transition: "morph",
      suggestions: ["結論を強く", "図解にして"],
    },
  ], async (server, prompts) => {
    const slides = shortDeck(6).slideData;
    slides[1].customImage = photo;
    slides[4].media = { src: "https://www.youtube.com/watch?v=dQw4w9WgXcQ", kind: "video", name: "YouTube" };
    const deck = { title: "テスト", theme: "clarity", transition: "fade", memo: "役員向け。数値は9月時点。", slides };
    const response = await server.postJson("/api/decks/chat", { deck, message: "この1枚を簡潔に", current: 1, focus: [3], history: [{ role: "user", text: "前の依頼" }, { role: "assistant", text: "前の回答" }] });
    assert.equal(response.status, 202, await response.clone().text());
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "completed", job.error);
    const { chat } = job;
    assert.match(chat.reply, /4枚目の後ろに1枚追加/);
    assert.deepEqual(chat.items.map((item) => item.from), [0, 1, 3, null, 4, 5]);
    assert.deepEqual(chat.deleted, [2]);
    assert.equal(chat.slides[1].customImage, photo, "uploaded photos survive an AI rewrite");
    assert.equal(chat.slides[1].animation, "click");
    assert.equal(chat.slides[3].title, "追加");
    assert.equal(chat.slides[3].details[0].text, "Xの根拠");
    assert.equal(chat.slides[4].media.src, slides[4].media.src, "videos survive too");
    assert.equal(chat.theme, undefined, "the theme is the user's choice; a theme from the AI is dropped");
    assert.equal(chat.transition, "morph");
    assert.doesNotMatch(chat.summary, /テーマ/);
    const [first, retry] = await prompts();
    assert.match(first, /デザイン（テーマ・配色・フォント）はユーザーが選ぶ/);
    assert.doesNotMatch(first, /ミッドナイト|midnight/, "no theme list is offered to the AI");
    assert.match(first, /前提条件（必ず守る）:\n役員向け。数値は9月時点。/);
    assert.match(first, /今見ているスライド: 2枚目／ユーザーが指定したスライド: 4枚目/);
    assert.match(first, /ユーザー: 前の依頼/);
    assert.doesNotMatch(first, /base64/, "image data never goes to the AI");
    assert.match(retry, /2枚目「本文1」の takeaway/);
  });
});

test("chat can change the deck's motion graphics and keeps a slide's own motion when rewriting it", async () => {
  await withFakeCodex([
    {
      reply: "表紙と章扉の背景を波にし、2枚目の文言を短くしました。",
      operations: [{ op: "replace", slide: 2, content: { type: "content", title: "本文1", takeaway: "短い結論", points: ["A"] } }],
      motion: { backdrop: "waves", kinetic: "mask", entrance: "drop", emphasis: "marker" },
      suggestions: [],
    },
  ], async (server, prompts) => {
    const slides = shortDeck(4).slideData;
    slides[1].kinetic = "chars";
    slides[1].backdrop = "lines";
    Object.assign(slides[1], { entrance: "flip", emphasis: "circle", transition: "curtain" });
    const deck = { title: "テスト", theme: "clarity", transition: "fade", motion: { kinetic: "mask", backdrop: "none", emphasis: "marker" }, slides };
    const job = await waitForJob(server, (await (await server.postJson("/api/decks/chat", { deck, message: "表紙の背景をもっと動かして" })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    const { chat } = job;
    assert.deepEqual(chat.motion, { backdrop: "waves", entrance: "drop" }, "only what changes is proposed");
    assert.match(chat.summary, /動きを変更/);
    assert.equal(chat.slides[1].kinetic, "chars", "the slide's own kinetic type stays");
    assert.equal(chat.slides[1].backdrop, "lines", "and so does its backdrop");
    assert.deepEqual([chat.slides[1].entrance, chat.slides[1].emphasis, chat.slides[1].transition], ["flip", "circle", "curtain"], "and its own entrance, emphasis and transition");
    const [prompt] = await prompts();
    assert.match(prompt, /大きな文字の動き: mask、背景の動き: none、強調: marker/);
    assert.match(prompt, /kinetic・backdrop/);
    assert.match(prompt, /spotlight/);
    assert.match(prompt, /curtain（幕が中央から開く）/);
  });
});

test("chat can turn a slide into one the audience works with, and the prompt says how to choose", async () => {
  await withFakeCodex([
    {
      reply: "2枚目を、人数と利用率を動かせる試算にしました。",
      operations: [{ op: "replace", slide: 2, content: {
        type: "simulator", title: "本文1", takeaway: "利用率を**8割**にすれば届く", source: "試行の実績",
        inputs: [{ label: "人数", value: 120, min: 50, max: 600, step: 10, unit: "人" }, { label: "利用率", value: 60, min: 10, max: 100, unit: "%" }],
        formula: "a × 12 × b ÷ 100", resultLabel: "月の削減時間", resultUnit: "時間", compareLabel: "目標", compareValue: 1440,
        details: [{ target: "takeaway", text: "1人あたり12時間は試行の平均", rows: [{ label: "資料作成", value: "5時間" }], source: "試行の実績" }],
      } }],
      suggestions: [],
    },
  ], async (server, prompts) => {
    const slides = shortDeck(4).slideData;
    const deck = { title: "テスト", theme: "clarity", transition: "fade", slides };
    const job = await waitForJob(server, (await (await server.postJson("/api/decks/chat", { deck, message: "2枚目を条件を変えて試算できるようにして。1人あたり12時間、120人、目標1,440時間、利用率60%、資料作成5時間" })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    const slide = job.chat.slides[1];
    assert.equal(slide.type, "simulator");
    assert.equal(slide.formula, "a × 12 × b ÷ 100");
    assert.deepEqual(slide.details[0].rows, [{ label: "資料作成", value: "5時間" }]);
    const [prompt] = await prompts();
    assert.match(prompt, /動く資料（1枚1操作）/);
    assert.match(prompt, /原因と結果（試算・感度） → 条件を動かすと計算し直す（simulator）/);
    assert.match(prompt, /chartType=rank/);
  });
});

test("chat answers questions without changing the deck and rejects unusable operations", async () => {
  await withFakeCodex([
    { reply: "表紙は削除できないため、代わりに…", operations: [{ op: "delete", slide: 1 }], suggestions: [] },
    { reply: "流れは問題ありません。", operations: [], suggestions: ["結論を1枚目に"] },
  ], async (server, prompts) => {
    const deck = { title: "テスト", slides: shortDeck(4).slideData };
    const job = await waitForJob(server, (await (await server.postJson("/api/decks/chat", { deck, message: "流れはどう？", attachment: { name: "議事録.docx", text: "売上は前年比108%" } })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.chat.changed, false);
    assert.equal(job.chat.slides, null);
    const [first, retry] = await prompts();
    assert.match(first, /添付した資料「議事録.docx」の内容[^\n]*\n売上は前年比108%/);
    assert.match(retry, /表紙と最後のスライドは削除できません/);
  });
});

test("image generation produces a real image artifact and uses the slide's own photo as a reference when asked", async () => {
  await withFakeCodex([{ __image: true, __expectReference: true }, { reply: "写真ライブラリから選べます。", operations: [], suggestions: [] }], async (server, prompts) => {
    const deck = { title: "テスト", slides: shortDeck(4).slideData };
    deck.slides[1].media = { src: `data:image/jpeg;base64,${(await readFile(PHOTO)).toString("base64")}`, kind: "image" };
    const blocked = await server.postJson("/api/decks/chat", { deck, message: "2枚目に画像生成で水彩風のイラストを入れて" });
    assert.equal(blocked.status, 422);
    assert.equal((await blocked.json()).code, "use_image_generation");
    assert.equal((await server.postJson("/api/decks/image", { deck, message: "8枚目に画像生成", index: 7 })).status, 422);
    const generated = await server.postJson("/api/decks/image", { deck, message: "2枚目の今の写真を元に水彩風の画像を生成して", index: 1 });
    assert.equal(generated.status, 202, await generated.clone().text());
    const imageJob = await waitForJob(server, (await generated.json()).jobId);
    assert.equal(imageJob.status, "completed", imageJob.error);
    assert.equal(imageJob.kind, "image");
    const image = await server.request(imageJob.image.url);
    assert.equal(image.status, 200);
    assert.match(image.headers.get("content-type"), /^image\//);
    assert.ok((await image.arrayBuffer()).byteLength > 100);
    const [prompt] = await prompts();
    assert.match(prompt, /\$imagegen/);
    assert.match(prompt, /水彩風/);
    assert.match(prompt, /参考画像を添付しました/);
    const ordinary = await server.postJson("/api/decks/chat", { deck, message: "写真ライブラリから選んで" });
    assert.equal(ordinary.status, 202);
    assert.equal((await waitForJob(server, (await ordinary.json()).jobId)).status, "completed");
  }, { FAKE_CODEX_IMAGE_SOURCE: PHOTO });
});

test("image generation without a saved image fails honestly", async () => {
  await withFakeCodex([{ reply: "画像を作りました", operations: [], suggestions: [] }], async (server) => {
    const deck = { title: "テスト", slides: shortDeck(4).slideData };
    const response = await server.postJson("/api/decks/image", { deck, message: "画像を生成して", index: 1 });
    const job = await waitForJob(server, (await response.json()).jobId);
    assert.equal(job.status, "failed");
    assert.match(job.error, /画像ファイルが返されません/);
    assert.equal(job.image, null);
  });
});

test("outline first: a skeleton is proposed, revised, then written into slides", async () => {
  const outline = {
    deckTitle: "改革の中間報告",
    reply: "結論を2枚目に置きました。",
    slides: [
      { type: "title", title: "改革の中間報告", takeaway: "", content: "" },
      { type: "kpi", title: "成果", takeaway: "作業を月120時間削減", content: "120時間、42%" },
      { type: "closing", title: "お願い", takeaway: "予算承認", content: "2.4億円" },
    ],
  };
  await withFakeCodex([outline, { ...outline, reply: "リスクを追加しました。" }, shortDeck(3)], async (server, prompts) => {
    const base = { brief: "改革の中間報告。月120時間削減。", audience: "役員", settings: { slideCount: 8, textDensity: "standard" } };
    const first = await waitForJob(server, (await (await server.postJson("/api/decks/outline", base)).json()).jobId);
    assert.equal(first.status, "completed", first.error);
    assert.equal(first.outline.slides.length, 3);
    const revised = await waitForJob(server, (await (await server.postJson("/api/decks/outline", { ...base, outline: first.outline.slides, instruction: "リスクを足して" })).json()).jobId);
    assert.equal(revised.outline.reply, "リスクを追加しました。");
    const written = await waitForJob(server, (await (await server.postJson("/api/decks", { ...base, outline: first.outline.slides })).json()).jobId);
    assert.equal(written.status, "completed", written.error);
    assert.equal(written.deck.slideData.length, 3, "the outline decides the slide count");
    const [p1, p2, p3] = await prompts();
    assert.match(p1, /骨子/);
    assert.match(p2, /ユーザーの指示: リスクを足して/);
    assert.match(p3, /確定した構成/);
    assert.match(p3, /2\. \[kpi\] 成果 — 作業を月120時間削減/);
    assert.match(p3, /指定枚数: 3枚/);
  });
});

test("layouts that look alike are chosen again: a plain outline is sent back, the user's own outline is respected", async () => {
  const row = (type, title) => ({ type, title, takeaway: `${title}の結論`, content: "" });
  // Four of six body slides are lines across the slide (timeline, lanes, process, roadmap).
  const plain = { deckTitle: "生成AIの歴史と将来", reply: "流れで見せます。", slides: [
    row("title", "生成AIの歴史と将来"), row("executiveSummary", "結論"), row("timeline", "大きな流れ"), row("diagram", "三つの層"),
    row("compare", "ツール比較"), row("process", "試す手順"), row("roadmap", "将来への備え"), row("closing", "お願い"),
  ] };
  const varied = { ...plain, reply: "形を内容に合わせました。", slides: plain.slides.map((item) => ({ ...item, type: { diagram: "pyramid", roadmap: "checklist" }[item.type] ?? item.type })) };
  await withFakeCodex([plain, varied, plain, shortDeck(8)], async (server, prompts) => {
    const base = { brief: "生成AIの歴史と将来。社内向け。", settings: { slideCount: 8, textDensity: "standard" } };
    const first = await waitForJob(server, (await (await server.postJson("/api/decks/outline", base)).json()).jobId);
    assert.equal(first.status, "completed", first.error);
    assert.equal(first.attempt, 2);
    assert.deepEqual(first.outline.slides.map((item) => item.type), ["title", "executiveSummary", "timeline", "pyramid", "compare", "process", "checklist", "closing"]);
    // The user puts the lines back and asks for something else: their choice stands.
    const revised = await waitForJob(server, (await (await server.postJson("/api/decks/outline", { ...base, outline: plain.slides, instruction: "結論を強く" })).json()).jobId);
    assert.equal(revised.status, "completed", revised.error);
    assert.equal(revised.attempt, 1, "no worse than what the user sent, so it is not sent back");
    // Written from an agreed outline, the layouts are not second-guessed either.
    const written = await waitForJob(server, (await (await server.postJson("/api/decks", { ...base, outline: plain.slides })).json()).jobId);
    assert.equal(written.status, "completed", written.error);
    assert.equal(written.attempt, 1);
    const sent = await prompts();
    // Three for the outline, two for the deck (8 slides from an outline are written 4 at a time).
    assert.equal(sent.length, 5);
    assert.match(sent[4], /5〜8枚目だけ/);
    assert.match(sent[0], /横に並ぶ流れ: process・timeline・roadmap/);
    assert.match(sent[1], /「横に並ぶ流れ」（工程・年表・ロードマップ・レーン図）の見た目が4枚あります（3・4・6・7枚目）/);
    assert.match(sent[1], /骨子全体のJSONをもう一度返してください/);
  });
});

test("a deck written from scratch with alike layouts gets one second chance to vary them", async () => {
  const plain = shortDeck(8);
  plain.slideData = [plain.slideData[0], ...["timeline", "process", "roadmap", "process", "timeline", "roadmap"].map((type, i) => ({
    type, title: `本文${i + 1}`, takeaway: "短い結論",
    ...(type === "timeline" ? { milestones: [{ label: "開始", date: "4月" }, { label: "展開", date: "7月" }] } : type === "process" ? { steps: ["準備", "実行"] } : { items: [{ title: "段階1" }, { title: "段階2" }] }),
  })), plain.slideData.at(-1)];
  await withFakeCodex([plain, shortDeck(8)], async (server, prompts) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", { brief: "テスト", settings: { slideCount: 8, textDensity: "standard" } })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.attempt, 2);
    assert.equal(job.deck.slideData[2].type, "cards");
    const [, retry] = await prompts();
    assert.match(retry, /見た目が6枚あります/);
    assert.match(retry, /新しい type のフィールドで中身を書き直す/);
    assert.match(retry, /全体のJSONをもう一度返してください/);
  });
  // A first answer rejected for another reason does not use up the chance.
  await withFakeCodex([shortDeck(4), plain, shortDeck(8)], async (server) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", { brief: "テスト", settings: { slideCount: 8, textDensity: "standard" } })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.attempt, 3);
    assert.equal(job.deck.slideData[2].type, "cards");
  });
  // Only once: a second plain answer is kept rather than asking again.
  await withFakeCodex([plain, plain, shortDeck(8)], async (server) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", { brief: "テスト", settings: { slideCount: 8, textDensity: "standard" } })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.attempt, 2);
    assert.equal(job.deck.slideData[1].type, "timeline");
  });
});

test("three variants of one slide keep its photo and report overflow per variant", async () => {
  const photo = `data:image/jpeg;base64,${(await readFile(PHOTO)).toString("base64")}`;
  await withFakeCodex([{ variants: [
    { label: "数字を大きく", slide: { type: "kpi", title: "成果", takeaway: "削減", items: [{ label: "作業", value: "-120h" }] } },
    { label: "図解", slide: { type: "cards", title: "成果", takeaway: "とても長い結論の文章".repeat(10), items: [{ title: "A", desc: "a" }, { title: "B", desc: "b" }] } },
    { label: "ひと言で", slide: { type: "statement", title: "成果", text: "月120時間を**取り戻した**" } },
  ] }], async (server, prompts) => {
    const slides = shortDeck(4).slideData;
    slides[1].customImage = photo;
    const job = await waitForJob(server, (await (await server.postJson("/api/decks/variants", { deck: { title: "テスト", slides }, slideIndex: 1, instruction: "図解も見たい" })).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.deepEqual(job.variants.map((variant) => variant.label), ["数字を大きく", "図解", "ひと言で"]);
    assert.ok(job.variants.every((variant) => variant.slide.customImage === photo));
    assert.equal(job.variants[0].issues.length, 0);
    assert.ok(job.variants[1].issues.length > 0, "the long takeaway is reported");
    const [prompt] = await prompts();
    assert.match(prompt, /2枚目について、見せ方の違う案を3つ/);
    assert.match(prompt, /ユーザーの希望: 図解も見たい/);
  });
});

test("the deck schema accepts every layout, video, details, builds and deck-wide design", async () => {
  await withFakeCodex([{ reply: "OK", operations: [], suggestions: [] }], async (server) => {
    const sample = JSON.parse(await readFile(join(root, "public", "samples", "ai-rollout.json"), "utf8"));
    const deck = { title: sample.deckTitle, theme: "kinari", accent: "#2d4b78", transition: "wipe", motion: { entrance: "pop", hover: "none", numbers: false, ambient: false, kinetic: "type", backdrop: "orbits", draw: false }, slides: sample.slideData };
    deck.slides[2].media = { src: "idb:video1", kind: "video", autoplay: true, loop: false, muted: false, placement: { x: 0.5, y: 0.4, w: 0.3, h: 0.3 } };
    deck.slides[4].media = { src: "idb:anim1", kind: "lottie", name: "spin.json", fit: "cover", autoplay: false, loop: true };
    deck.slides[0].backdrop = "grid";
    deck.slides[1].kinetic = "chars";
    const response = await server.postJson("/api/decks/chat", { deck, message: "流れを確認して" });
    assert.equal(response.status, 202, await response.clone().text());
    const bad = await server.postJson("/api/decks/chat", { deck: { ...deck, theme: "neon" }, message: "x" });
    assert.equal(bad.status, 400);
    const badMotion = await server.postJson("/api/decks/chat", { deck: { ...deck, motion: { backdrop: "fireworks" } }, message: "x" });
    assert.equal(badMotion.status, 400);
  });
});

test("an Excel workbook is read as table cells (formulas and number formats kept)", async () => {
  const venv = join(root, ".venv", "bin", "python");
  const python = process.env.PYTHON_BIN || (existsSync(venv) ? venv : "python3");
  const script = [
    "import io, sys, zipfile",
    "buf = io.BytesIO()",
    "M = 'xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"'",
    "z = zipfile.ZipFile(buf, 'w')",
    "z.writestr('xl/workbook.xml', '<workbook ' + M + '><sheets><sheet name=\"売上\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>')",
    "z.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"ws\" Target=\"worksheets/sheet1.xml\"/></Relationships>')",
    "z.writestr('xl/styles.xml', '<styleSheet ' + M + '><cellXfs count=\"2\"><xf numFmtId=\"0\"/><xf numFmtId=\"3\"/></cellXfs></styleSheet>')",
    "z.writestr('xl/worksheets/sheet1.xml', '<worksheet ' + M + '><sheetData><row r=\"1\"><c r=\"A1\" s=\"1\"><v>1200</v></c><c r=\"B1\" s=\"1\"><f>A1*2</f><v>2400</v></c></row></sheetData></worksheet>')",
    "z.close(); sys.stdout.buffer.write(buf.getvalue())",
  ].join("\n");
  const child = spawn(python, ["-c", script], { stdio: ["ignore", "pipe", "pipe"] });
  const chunks = [];
  child.stdout.on("data", (chunk) => chunks.push(chunk));
  const [code] = await once(child, "exit");
  assert.equal(code, 0);
  const server = await startServer({ PYTHON_BIN: python });
  try {
    const done = await server.request("/api/tables/xlsx?name=%E5%A3%B2%E4%B8%8A.xlsx", { method: "POST", headers: { "content-type": "application/octet-stream" }, body: Buffer.concat(chunks) });
    assert.equal(done.status, 200, await done.clone().text());
    const { sheets } = await done.json();
    assert.equal(sheets[0].name, "売上");
    assert.deepEqual(sheets[0].rows[0].map((c) => [c.text, c.f || null]), [["1,200", null], ["2,400", "=A1*2"]]);
    const wrong = await server.request("/api/tables/xlsx?name=old.xls", { method: "POST", body: "x" });
    assert.equal(wrong.status, 400);
    const broken = await server.request("/api/tables/xlsx?name=broken.xlsx", { method: "POST", body: "not a zip" });
    assert.equal(broken.status, 422);
    assert.match((await broken.json()).error, /Excel/);
  } finally {
    await server.stop();
  }
});

test("an existing PowerPoint is imported into an editable deck and its text can be attached", async (t) => {
  const venv = join(root, ".venv", "bin", "python");
  const python = process.env.PYTHON_BIN || (existsSync(venv) ? venv : "python3");
  const script = [
    "import io, sys",
    "from pptx import Presentation",
    "from pptx.util import Inches, Pt",
    "prs = Presentation()",
    "prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)",
    "s = prs.slides.add_slide(prs.slide_layouts[5]); s.shapes.title.text = '改革の報告'",
    "s = prs.slides.add_slide(prs.slide_layouts[5]); s.shapes.title.text = '主要な数値'",
    "for i, (v, l) in enumerate([('-120h/月', '作業時間'), ('-0.3pt', '廃棄率')]):",
    "    for j, (text, size) in enumerate([(v, 32), (l, 20)]):",
    "        f = s.shapes.add_textbox(Inches(0.5 + i * 4), Inches(2 + j * 0.75), Inches(3.5), Inches(0.7)).text_frame",
    "        r = f.paragraphs[0].add_run(); r.text = text; r.font.size = Pt(size); r.font.bold = j == 0",
    "s.notes_slide.notes_text_frame.text = '数値の出典は業務ログ'",
    "s = prs.slides.add_slide(prs.slide_layouts[5]); s.shapes.title.text = '次のアクション'",
    "out = io.BytesIO(); prs.save(out); sys.stdout.buffer.write(out.getvalue())",
  ].join("\n");
  const child = spawn(python, ["-c", script], { stdio: ["ignore", "pipe", "pipe"] });
  const chunks = [];
  child.stdout.on("data", (chunk) => chunks.push(chunk));
  const [code] = await once(child, "exit");
  if (code !== 0) { t.skip("python-pptx is not installed (pip install -r tools/requirements.txt)"); return; }
  const pptx = Buffer.concat(chunks);
  const server = await startServer({ PYTHON_BIN: python });
  try {
    const imported = await server.request("/api/import?name=report.pptx", { method: "POST", headers: { "content-type": "application/octet-stream" }, body: pptx });
    assert.equal(imported.status, 200, await imported.clone().text());
    const deck = await imported.json();
    assert.equal(deck.slideData[0].type, "title");
    assert.equal(deck.slideData[0].title, "改革の報告");
    const kpi = deck.slideData.find((slide) => slide.type === "kpi");
    assert.deepEqual(kpi.items.map((item) => item.value), ["-120h/月", "-0.3pt"]);
    assert.match(kpi.notes, /業務ログ/);
    const extracted = await server.request("/api/extract?name=report.pptx", { method: "POST", headers: { "content-type": "application/octet-stream" }, body: pptx });
    assert.equal(extracted.status, 200);
    assert.match((await extracted.json()).text, /主要な数値/);
    const wrong = await server.request("/api/import?name=report.exe", { method: "POST", body: "x" });
    assert.equal(wrong.status, 400);
  } finally {
    await server.stop();
  }
});

test("a long deck is written a few slides at a time: from an outline, or when an answer does not fit", async () => {
  const deck = shortDeck(8);
  const part = (from, to) => ({ ...deck, slideData: deck.slideData.slice(from - 1, to) });
  const outline = deck.slideData.map((slide) => ({ type: slide.type, title: slide.title, takeaway: "", content: "" }));
  const input = { brief: "テスト", settings: { slideCount: 8, textDensity: "standard" }, outline };
  // From an outline: 4 slides per answer, the parts joined and checked as one deck.
  await withFakeCodex([part(1, 4), part(5, 8)], async (server, prompts) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", input)).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.deck.slideData.length, 8);
    assert.deepEqual(job.deck.slideData.map((s) => s.title), deck.slideData.map((s) => s.title));
    const sent = await prompts();
    assert.equal(sent.length, 2);
    assert.match(sent[0], /全8枚のうち1〜4枚目だけ/);
    assert.match(sent[1], /全8枚のうち5〜8枚目だけ/);
  });
  // Without an outline the deck is asked for at once. An answer cut off at the output limit (Codex would send the
  // same request again and again) is asked again at once in a new thread without the strict output format; cut off
  // again, it is asked in parts.
  const cutoff = { __cutoff: true };
  await withFakeCodex([cutoff, cutoff, part(1, 4), { __raw: `前置き ${JSON.stringify(part(5, 8))}` }], async (server, prompts, turns) => {
    const started = await (await server.postJson("/api/decks", { brief: "テスト", settings: { slideCount: 8, textDensity: "standard" } })).json();
    assert.ok(started.jobId, JSON.stringify(started));
    const job = await waitForJob(server, started.jobId);
    assert.equal(job.status, "completed", job.error);
    assert.equal(job.deck.slideData.length, 8);
    const sent = await prompts();
    const sentTurns = await turns();
    assert.equal(sent.length, 4);
    assert.deepEqual(sentTurns.map((t) => t.schema), [true, false, false, false]);
    assert.equal(new Set(sentTurns.slice(0, 3).map((t) => t.thread)).size, 3, "each cut-off answer is asked again in a new thread");
    assert.equal(sentTurns[3].thread, sentTurns[2].thread);
    assert.deepEqual(sentTurns.map((t) => t.effort), ["medium", "low", "low", "low"]);
    assert.doesNotMatch(sent[0], /何回かに分けて|# 回答の形式/);
    assert.match(sent[1], /# 回答の形式/);
    assert.match(sent[1], /SlideDataItem/);
    assert.doesNotMatch(sent[1], /何回かに分けて/);
    assert.match(sent[2], /4枚ずつ何回かに分けて/);
    assert.match(sent[2], /# 回答の形式/);
    assert.match(sent[3], /5〜8枚目だけ/);
    assert.doesNotMatch(sent[3], /# 回答の形式/);
  });
  // A part that still does not fit is asked again two slides at a time in a new thread, with the slides already
  // written; a whole deck in a part keeps its own slides; a part with a broken slide is asked again.
  const tooLong = { __fail: { message: "stream disconnected before completion: Incomplete response returned, reason: max_output_tokens" } };
  const broken = part(7, 8);
  broken.slideData = [{ ...broken.slideData[0], type: "agenda", items: undefined }, broken.slideData[1]];
  await withFakeCodex([part(1, 4), tooLong, tooLong, deck, broken, part(7, 8)], async (server, prompts, turns) => {
    const job = await waitForJob(server, (await (await server.postJson("/api/decks", input)).json()).jobId);
    assert.equal(job.status, "completed", job.error);
    assert.deepEqual(job.deck.slideData.map((s) => s.title), deck.slideData.map((s) => s.title));
    const sent = await prompts();
    const sentTurns = await turns();
    assert.equal(sent.length, 6);
    assert.match(sent[2], /5〜8枚目だけ/);
    assert.match(sent[2], /1〜4枚目は書き終えている/);
    assert.match(sent[3], /途中で切れました/);
    assert.match(sent[3], /2枚ずつ/);
    assert.match(sent[3], /5〜6枚目だけ/);
    assert.match(sent[4], /7〜8枚目だけ/);
    assert.match(sent[5], /問題がありました/);
    assert.match(sent[5], /7枚目（agenda）/);
    assert.equal(sentTurns[4].thread, sentTurns[3].thread);
  });
});
