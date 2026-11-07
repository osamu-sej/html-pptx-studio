// パスワードを使用して暗号化: the written HTML opens only with its password (the locked page's own script reads it back),
// the presentation is not readable in the file, and nothing of the password is kept.
import assert from "node:assert/strict";
import test from "node:test";
import vm from "node:vm";

import { encryptText, lockHtml, UNLOCK_SCRIPT } from "../public/editor/protect.mjs";

/** The locked page's script, run without a form: its __hpptxUnlock. */
function unlocker() {
  const window = {};
  const context = vm.createContext({ window, document: { getElementById: () => null }, crypto, atob, TextEncoder, TextDecoder, Blob, Response, DecompressionStream, Uint8Array });
  vm.runInContext(UNLOCK_SCRIPT, context);
  return window.__hpptxUnlock;
}

test("the locked HTML opens with its password and not without", async () => {
  const html = "<!doctype html><html><body><h1>売上の見通し</h1><script>var secret = '機密';</script></body></html>".repeat(20);
  const page = await lockHtml(html, "pass-2026", { title: "見通し", iterations: 1000 });
  assert.ok(!page.includes("売上の見通し") && !page.includes("機密"), "the slides are not in the file as text");
  assert.ok(!page.includes("pass-2026"), "the password is not kept");
  assert.match(page, /<title>見通し（パスワードで保護）<\/title>/);
  const payload = JSON.parse(page.match(/<script type="application\/json" id="hs-locked">([^<]*)<\/script>/)[1]);
  assert.equal(payload.iter, 1000);
  const unlock = unlocker();
  assert.equal(await unlock(payload, "pass-2026"), html);
  await assert.rejects(() => unlock(payload, "pass-2025"), "a wrong password does not open it");
});

test("each lock is different (salt and IV), the text packed before it is encrypted", async () => {
  const text = "あ".repeat(5000);
  const a = await encryptText(text, "pw", { iterations: 1000 });
  const b = await encryptText(text, "pw", { iterations: 1000 });
  assert.notEqual(a.salt, b.salt);
  assert.notEqual(a.iv, b.iv);
  assert.notEqual(a.data, b.data);
  assert.ok(Buffer.from(a.data, "base64").length < 1000, "packed (gzip) — repeated text is small");
  await assert.rejects(() => lockHtml("<p>x</p>", ""), /パスワード/);
});
