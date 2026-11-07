// PowerPoint Live in the deck: アンケート (polls) as kept and drawn, the choices typed one per line.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { parseOptions, POLL_DEFAULT } from "../public/editor/poll.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
const plain = (value) => JSON.parse(JSON.stringify(value));

async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  window.requestAnimationFrame = (fn) => setTimeout(fn, 0);
  for (const name of ["engine.js", "objects.js"]) {
    const code = await readFile(join(root, "public", "engine", name), "utf8");
    vm.runInContext(name === "engine.js" ? code.replace("/*__ICONS__*/{}", () => icons) : code, context, { filename: name });
  }
  return window.SlideEngine;
}

test("a poll keeps a plain question and two to eight plain choices", async () => {
  const E = await loadEngine();
  const base = { id: "p1", kind: "poll", x: 100, y: 100, w: 1000, h: 500 };
  const o = E.normalizeObject({ ...base, question: "<b>どれ</b>がよい？", options: ["A案", " <i>B案</i> ", "", "C案"], fill: "#dce4f2" });
  assert.equal(o.question, "どれがよい？");
  assert.deepEqual(plain(o.options), ["A案", "B案", "C案"]);
  assert.equal(o.fill, "#dce4f2");
  assert.equal(E.normalizeObject({ ...base, question: "x", options: ["only one"] }), null, "two choices at least");
  assert.equal(E.normalizeObject({ ...base, question: "x", options: Array.from({ length: 12 }, (_, i) => `案${i}`) }).options.length, 8);
  assert.ok(E.normalizeObject({ ...base, ...POLL_DEFAULT }), "the default poll is valid");
});

test("a poll is drawn with its question and a bar a choice (clickable in a show)", async () => {
  const E = await loadEngine();
  const slide = { type: "blank", title: "", elements: [{ id: "p1", kind: "poll", x: 100, y: 100, w: 1000, h: 500, question: "どれがよい？", options: ["A案", "B案"] }] };
  const edit = E.render(slide, { mode: "edit", index: 0, deck: { theme: "clarity", slides: [slide] } });
  const box = edit.querySelector(".hs-poll");
  assert.equal(box.dataset.poll, "p1");
  assert.equal(box.querySelector(".hs-poll-q").textContent, "どれがよい？");
  assert.equal(box.querySelectorAll(".hs-poll-opt").length, 2);
  assert.equal(box.querySelector(".hs-poll-opt").getAttribute("role"), null, "not clickable while editing");
  const show = E.render(slide, { mode: "present", index: 0, live: true, deck: { theme: "clarity", slides: [slide] } });
  assert.ok(show.querySelector(".hs-poll-opt .hs-poll-bar i"));
});

test("choices typed one per line: bullets and numbers go, blank lines go, eight at most", () => {
  assert.deepEqual(parseOptions("・A案\n2) B案\n\n- C案 \n"), ["A案", "B案", "C案"]);
  assert.equal(parseOptions(Array.from({ length: 10 }, (_, i) => `案${i}`).join("\n")).length, 8);
});
