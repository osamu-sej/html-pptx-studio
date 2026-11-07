#!/usr/bin/env node
// Minimal stand-in for `codex app-server`: answers JSON-RPC over stdio and
// replies to each turn with the next scripted answer from FAKE_CODEX_SCRIPT.
import { copyFileSync, existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import readline from "node:readline";

const script = JSON.parse(readFileSync(process.env.FAKE_CODEX_SCRIPT, "utf8"));
const promptLog = process.env.FAKE_CODEX_PROMPTS;
const prompts = [];
let turn = 0;
let login = 0;
const cancelledLogins = new Set();

const send = (message) => process.stdout.write(`${JSON.stringify(message)}\n`);

readline.createInterface({ input: process.stdin }).on("line", (line) => {
  const message = JSON.parse(line);
  const { id, method, params } = message;
  if (method === "initialize") {
    const marker = process.env.FAKE_CODEX_INIT_EXIT_ONCE_FILE;
    if (marker && !existsSync(marker)) {
      writeFileSync(marker, "1");
      process.exit(1);
    }
    return send({ id, result: {} });
  }
  if (method === "account/read") return send({ id, result: { account: { type: "chatgpt", planType: "test" } } });
  if (method === "account/logout") return send({ id, result: {} });
  if (method === "account/login/start") {
    const loginId = `login-${++login}`;
    send({ id, result: { type: "chatgptDeviceCode", loginId, userCode: `TEST-CODE-${login}`, verificationUrl: "https://example.com/device" } });
    setTimeout(() => {
      if (!cancelledLogins.has(loginId)) send({ method: "account/login/completed", params: { loginId, success: !process.env.FAKE_CODEX_LOGIN_ERROR, error: process.env.FAKE_CODEX_LOGIN_ERROR || null } });
    }, Number(process.env.FAKE_CODEX_LOGIN_DELAY_MS || 20));
    return;
  }
  if (method === "account/login/cancel") {
    cancelledLogins.add(params.loginId);
    return send({ id, result: {} });
  }
  if (method === "thread/start") return send({ id, result: { thread: { id: `thread-${Date.now()}` } } });
  if (method === "turn/start") {
    const threadId = params.threadId;
    prompts.push(params.input?.[0]?.text ?? "");
    if (promptLog) writeFileSync(promptLog, JSON.stringify(prompts));
    let answer = script[Math.min(turn, script.length - 1)];
    turn += 1;
    send({ id, result: { turn: { id: `turn-${turn}` } } });
    // { "__fail": error } ends the turn with that error; { "__hiccup": answer } reports an error Codex
    // retries by itself, then answers.
    if (answer?.__fail) {
      setTimeout(() => {
        send({ method: "error", params: { threadId, turnId: `turn-${turn}`, willRetry: false, error: answer.__fail } });
        send({ method: "turn/completed", params: { threadId, turn: { id: `turn-${turn}`, status: "failed", error: answer.__fail } } });
      }, 20);
      return;
    }
    if (answer?.__hiccup) {
      send({ method: "error", params: { threadId, turnId: `turn-${turn}`, willRetry: true, error: { message: "stream disconnected", codexErrorInfo: { responseStreamDisconnected: {} } } } });
      answer = answer.__hiccup;
    }
    if (answer?.__image) {
      if (answer.__expectReference && !params.input?.some((part) => part.type === "localImage" && readFileSync(part.path).length > 100)) {
        send({ method: "turn/completed", params: { threadId, turn: { id: `turn-${turn}`, status: "failed", error: { message: "Reference image missing" } } } });
        return;
      }
      if (answer.__expectNoReference && params.input?.some((part) => part.type === "localImage")) {
        send({ method: "turn/completed", params: { threadId, turn: { id: `turn-${turn}`, status: "failed", error: { message: "Unexpected reference image" } } } });
        return;
      }
      const folder = join(process.env.CODEX_HOME, "generated_images");
      mkdirSync(folder, { recursive: true });
      const savedPath = join(folder, `test-${turn}.png`);
      copyFileSync(process.env.FAKE_CODEX_IMAGE_SOURCE, savedPath);
      setTimeout(() => {
        send({ method: "item/completed", params: { threadId, item: { id: `image-${turn}`, type: "imageGeneration", status: "completed", result: "success", savedPath, revisedPrompt: "test illustration" } } });
        send({ method: "turn/completed", params: { threadId, turn: { id: `turn-${turn}`, status: "completed" } } });
      }, 30);
      return;
    }
    const text = JSON.stringify(answer);
    // FAKE_CODEX_STREAM=ms streams the answer in pieces, like the real app server.
    const every = Number(process.env.FAKE_CODEX_STREAM || 0);
    const pieces = every ? text.match(/[\s\S]{1,400}/g) : [];
    pieces.forEach((delta, index) => setTimeout(() => send({ method: "item/agentMessage/delta", params: { threadId, itemId: "m1", delta } }), every * (index + 1)));
    setTimeout(() => {
      send({ method: "item/completed", params: { threadId, item: { type: "agentMessage", text } } });
      send({ method: "turn/completed", params: { threadId, turn: { id: `turn-${turn}`, status: "completed" } } });
    }, 30 + every * (pieces.length + 1));
    return;
  }
  if (id != null) send({ id, result: {} });
});
