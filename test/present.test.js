// 画面切り替えのサウンド (motion.js's synthesiser, the slide field kept for AI edits) and the fingerprints co-editing
// uses to merge changes that crossed (public/editor/coedit-merge.mjs).
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { textHash } from "../public/editor/coedit-merge.mjs";
import { deckShape } from "../server/schemas.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));

/** A Web Audio stand-in that counts what the synthesiser makes, starts and stops. */
function fakeAudio() {
  const made = { oscillators: 0, noises: 0, started: 0, stopped: 0 };
  const param = () => ({ value: 0, setValueAtTime() {}, exponentialRampToValueAtTime() {} });
  const node = (extra = {}) => ({ connect: (next) => next, start: () => { made.started += 1; }, stop: () => { made.stopped += 1; }, ...extra });
  class AudioContext {
    constructor() { this.currentTime = 0; this.sampleRate = 8000; this.state = "running"; this.destination = node(); }
    createGain() { return node({ gain: param() }); }
    createOscillator() { made.oscillators += 1; return node({ type: "sine", frequency: param() }); }
    createBiquadFilter() { return node({ type: "lowpass", frequency: param(), Q: { value: 1 } }); }
    createBuffer(channels, length) { return { getChannelData: () => new Float32Array(length) }; }
    createBufferSource() { made.noises += 1; return node({ buffer: null, loop: false }); }
    resume() { return Promise.resolve(); }
  }
  return { AudioContext, made };
}

async function loadEngine(audio) {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  // (linkedom's window falls back to Node's globals: set both names either way.)
  window.AudioContext = audio || null;
  window.webkitAudioContext = null;
  window.window = window;
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  for (const name of ["engine.js", "objects.js", "animate.js", "motion.js"]) {
    const code = await readFile(join(root, "public", "engine", name), "utf8");
    vm.runInContext(name === "engine.js" ? code.replace("/*__ICONS__*/{}", () => icons) : code, context, { filename: name });
  }
  return window.SlideEngine;
}

test("every transition sound is synthesised (no files); 前のサウンドを停止 stops what still rings", async () => {
  const { AudioContext, made } = fakeAudio();
  const E = await loadEngine(AudioContext);
  const kinds = Object.keys(E.TRANSITION_SOUNDS);
  assert.ok(kinds.length >= 8 && kinds.includes("stop"));
  for (const kind of kinds.filter((k) => k !== "stop")) {
    const before = made.oscillators + made.noises;
    const seconds = E.playSound(kind);
    assert.ok(seconds > 0 && seconds < 3, `${kind} rings for a while (${seconds}s)`);
    assert.ok(made.oscillators + made.noises > before, `${kind} makes sound`);
  }
  assert.equal(E.playSound("nonsense"), 0, "an unknown sound is silent");
  const stoppedBefore = made.stopped;
  E.playSound("applause");
  E.playSound("stop");
  assert.ok(made.stopped > stoppedBefore, "stop ends the sounds still ringing");
});

test("without Web Audio the sounds are simply silent", async () => {
  const E = await loadEngine(null);
  assert.equal(E.playSound("chime"), 0);
});

test("a slide's transition sound is part of a deck (one of the known sounds)", () => {
  const parsed = deckShape.parse({ title: "t", slides: [{ type: "title", title: "表紙", transitionSound: "chime" }] });
  assert.equal(parsed.slides[0].transitionSound, "chime");
  assert.throws(() => deckShape.parse({ title: "t", slides: [{ type: "title", title: "表紙", transitionSound: "boom" }] }));
});

test("co-editing fingerprints name one version: the same JSON on both sides gives the same fingerprint", () => {
  const slide = { type: "blank", sid: "s1", elements: [{ id: "a", kind: "text", text: "<p>メモ😀</p>", x: 1.5 }] };
  const sent = JSON.stringify(slide);
  // What the server stores is the parsed request body; written out again, it is the very same text.
  assert.equal(textHash(JSON.stringify(JSON.parse(sent))), textHash(sent));
  assert.notEqual(textHash(sent), textHash(JSON.stringify({ ...slide, title: "x" })));
  assert.match(textHash(""), /^[0-9a-f]{8}$/);
});
