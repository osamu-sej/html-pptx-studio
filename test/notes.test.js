// ノートの書式: the speaker notes kept as formatted words (`notesRich`) beside the plain words (`notes`), so everything that
// reads the notes (print, compare, word count, the AI, the presenter view) keeps working, and the formatting is never trusted
// past what the plain words say.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import { slideSchema } from "../server/schemas.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  return window.SlideEngine;
}

test("ノートの書式: bold, italic, lists and links are kept; colours, sizes and scripts are not", async () => {
  const E = await loadEngine();
  const html = E.noteHtml('<p>これは<b>大事</b>で<i>斜め</i>の<u>下線</u><s>消す</s></p><p><span style="color:#ff0000;font-size:40px">色と大きさ</span>は外れる</p><p>リンクは<a href="https://example.com/a" onclick="x()">ここ</a>と<a href="javascript:alert(1)">これ</a></p><script>alert(1)</script><img src=x onerror=alert(1)>');
  assert.ok(html.includes("<b>大事</b>") && html.includes("<i>斜め</i>") && html.includes("<u>下線</u>") && html.includes("<s>消す</s>"), html);
  assert.ok(!/color|font-size|style=|<span|onclick|onerror|<script|<img|javascript:/i.test(html), `nothing unsafe or decorative is left: ${html}`);
  assert.ok(html.includes('href="https://example.com/a"'), "a safe link stays");
  assert.ok(html.includes("色と大きさは外れる") || html.includes("色と大きさ</p>") || html.includes("色と大きさ"), "the words stay");
});

test("ノートの書式: words made bold or underlined with a style (an editor's way) keep it as tags", async () => {
  const E = await loadEngine();
  const html = E.noteHtml('<p><span style="font-weight: bold;">太</span><span style="font-style: italic; font-weight: 700">斜太</span><span style="text-decoration: underline line-through">二重</span><span style="color:red">色だけ</span></p>');
  assert.ok(html.includes("<b>太</b>") && /<(b|i)><(b|i)>斜太<\/(b|i)><\/(b|i)>/.test(html) && /<s><u>二重<\/u><\/s>|<u><s>二重<\/s><\/u>/.test(html) && html.includes("色だけ") && !html.includes("color"), html);
});

test("ノートの書式: the plain words keep the lines, bullets and numbers", async () => {
  const E = await loadEngine();
  const html = "<p>はじめに</p><ul><li>一つ目<ul><li>入れ子</li></ul></li><li>二つ目</li></ul><ol><li>手順<b>A</b></li><li>手順B</li></ol><p>おわり</p>";
  assert.equal(E.noteToText(html), "はじめに\n・一つ目\n  ・入れ子\n・二つ目\n1. 手順A\n2. 手順B\nおわり");
  assert.equal(E.noteToText('<ol start="3"><li>三</li><li>四</li></ol>'), "3. 三\n4. 四");
  assert.equal(E.noteToText("<p>a<br>b</p><p><br></p><p>c</p>"), "a\nb\n\nc");
});

test("ノートの書式: noteFields keeps the formatting only when something is formatted", async () => {
  const E = await loadEngine();
  const plain = E.noteFields("<div>ただの文字</div><div>二行目</div>");
  assert.equal(plain.notes, "ただの文字\n二行目");
  assert.equal(plain.notesRich, undefined, "plain paragraphs need no formatting kept");
  const bold = E.noteFields("<p>ここは<b>太字</b></p>");
  assert.equal(bold.notes, "ここは太字");
  assert.equal(bold.notesRich, "<p>ここは<b>太字</b></p>");
  const list = E.noteFields("<ul><li>a</li><li>b</li></ul>");
  assert.equal(list.notes, "・a\n・b");
  assert.ok(list.notesRich.startsWith("<ul>"));
  const empty = E.noteFields("<p><b></b></p>");
  assert.equal(empty.notes, "");
  assert.equal(empty.notesRich, undefined, "nothing written, nothing formatted");
  assert.equal(E.noteFields("").notes, "");
});

test("ノートの書式: noteOf shows the formatting only while it still says what the plain notes say", async () => {
  const E = await loadEngine();
  const fields = E.noteFields("<p>前置き</p><ul><li>要点<b>1</b></li></ul>");
  const slide = { type: "blank", ...fields };
  const same = E.noteOf(slide);
  assert.equal(same.formatted, true);
  assert.equal(same.rich, fields.notesRich);
  assert.equal(same.text, fields.notes);
  // Someone (the AI, a plain editor) changed the plain words: the old formatting is not shown over them.
  const rewritten = E.noteOf({ ...slide, notes: "AI が書き直したノート\n二行目" });
  assert.equal(rewritten.formatted, false);
  assert.equal(rewritten.rich, "<p>AI が書き直したノート</p><p>二行目</p>");
  // Trailing spaces and blank ends do not break the match.
  assert.equal(E.noteOf({ ...slide, notes: `${fields.notes}  \n` }).formatted, true);
  // No notes at all.
  const none = E.noteOf({ type: "blank" });
  assert.deepEqual({ ...none }, { rich: "", text: "", formatted: false });
  // Plain notes with markup characters are shown as words, never as markup.
  assert.equal(E.noteOf({ notes: "1 < 2 & <b>x</b>" }).rich, "<p>1 &lt; 2 &amp; &lt;b&gt;x&lt;/b&gt;</p>");
  // A stored notesRich that is not what it claims (markup that is only plain paragraphs) is not "formatted".
  assert.equal(E.noteOf({ notes: "x", notesRich: "<p>x</p>" }).formatted, false);
  assert.equal(E.noteOf({ notes: "x", notesRich: "<p><script>alert(1)</script>x</p>" }).formatted, false);
});

test("ノートの書式: the slide schema keeps notesRich, and it is the user's own field", async () => {
  const slide = slideSchema.parse({ type: "blank", title: "t", notes: "・a", notesRich: "<ul><li>a</li></ul>" });
  assert.equal(slide.notesRich, "<ul><li>a</li></ul>");
  assert.equal(slide.notes, "・a");
  const schemas = await readFile(join(root, "server", "schemas.mjs"), "utf8");
  assert.ok(/USER_ONLY_FIELDS = new Set\([^)]*"notesRich"/.test(schemas), "hidden from the AI's output schema");
  const chat = await readFile(join(root, "server", "chat.mjs"), "utf8");
  assert.ok(/"hideMaster", "notesRich"\]\) if \(original/.test(chat), "put back after the AI changes a slide");
});
