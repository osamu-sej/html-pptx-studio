// ファイル → 情報 → プレゼンテーションの保護 → パスワードを使用して暗号化: the HTML file written out is locked with a
// password. The page is packed (gzip), encrypted (AES-GCM, the key from the password by PBKDF2-SHA-256) and put in a
// small page that asks for the password and opens the presentation in place. The password is never kept anywhere.

const ITERATIONS = 250000;
const b64 = (bytes) => {
  let s = "";
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
};

async function pack(text) {
  const stream = new Blob([new TextEncoder().encode(text)]).stream().pipeThrough(new CompressionStream("gzip"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** The text encrypted with the password: { v, iter, salt, iv, data } (base64). */
export async function encryptText(text, password, { iterations = ITERATIONS } = {}) {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const base = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]);
  const key = await crypto.subtle.deriveKey({ name: "PBKDF2", salt, iterations, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, false, ["encrypt"]);
  const data = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, await pack(text)));
  return { v: 1, iter: iterations, salt: b64(salt), iv: b64(iv), data: b64(data) };
}

/**
 * The locked page's script (plain ES5-ish, no modules): __hpptxUnlock(payload, password) gives back the text, or throws
 * when the password is wrong; the form opens the presentation.
 */
export const UNLOCK_SCRIPT = `(function () {
  function bytes(s) { var b = atob(s), out = new Uint8Array(b.length); for (var i = 0; i < b.length; i += 1) out[i] = b.charCodeAt(i); return out; }
  async function unlock(p, password) {
    var base = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]);
    var key = await crypto.subtle.deriveKey({ name: "PBKDF2", salt: bytes(p.salt), iterations: p.iter, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, false, ["decrypt"]);
    var packed = await crypto.subtle.decrypt({ name: "AES-GCM", iv: bytes(p.iv) }, key, bytes(p.data));
    var stream = new Blob([packed]).stream().pipeThrough(new DecompressionStream("gzip"));
    return new TextDecoder().decode(await new Response(stream).arrayBuffer());
  }
  window.__hpptxUnlock = unlock;
  var form = document.getElementById("lock");
  if (!form) return;
  var input = document.getElementById("pw"), msg = document.getElementById("msg"), go = document.getElementById("go");
  if (!window.crypto || !crypto.subtle || typeof DecompressionStream === "undefined") { msg.textContent = "このブラウザでは開けません。最新の Chrome・Edge・Safari・Firefox で開いてください。"; go.disabled = true; return; }
  form.addEventListener("submit", async function (event) {
    event.preventDefault();
    go.disabled = true; msg.textContent = "開いています…";
    try {
      var html = await unlock(JSON.parse(document.getElementById("hs-locked").textContent), input.value);
      document.open(); document.write(html); document.close();
    } catch (e) {
      go.disabled = false; msg.textContent = "パスワードが違います。"; input.select();
    }
  });
  input.focus();
})();`;

const esc = (s) => String(s ?? "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" })[c]);

/** The locked page for a whole HTML file. */
export async function lockHtml(html, password, { title = "プレゼンテーション", iterations } = {}) {
  if (!password) throw new Error("パスワードを入れてください");
  const payload = await encryptText(html, password, { iterations });
  return `<!doctype html>
<html lang="ja"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>${esc(title)}（パスワードで保護）</title>
<style>:root{color-scheme:light}body{margin:0;min-height:100vh;display:grid;place-items:center;background:#f4f5f7;font-family:"Noto Sans JP","Hiragino Sans","Yu Gothic",sans-serif;color:#1a1a1a}
form{width:min(380px,calc(100vw - 32px));padding:28px 24px;background:#fff;border:1px solid #d9dce3;border-radius:12px;box-sizing:border-box}
h1{font-size:17px;margin:0 0 6px}p{margin:0 0 16px;font-size:13px;line-height:1.7;color:#4a4f5a}
input{width:100%;box-sizing:border-box;height:40px;padding:0 10px;border:1px solid #b9bfcc;border-radius:8px;font:inherit;font-size:15px}
button{margin-top:12px;width:100%;height:40px;border:0;border-radius:8px;background:#dce4f2;color:#1f3864;font:inherit;font-weight:700;cursor:pointer}button:disabled{opacity:.6}
#msg{margin:10px 0 0;min-height:1.4em;color:#9b2c2c}</style></head>
<body><form id="lock" autocomplete="off"><h1>${esc(title)}</h1><p>このプレゼンテーションはパスワードで保護されています。パスワードを入れて開いてください。</p>
<label for="pw" style="font-size:12px;font-weight:700">パスワード</label><input id="pw" type="password" required autocomplete="current-password"><button id="go" type="submit">開く</button><p id="msg" role="alert"></p></form>
<script type="application/json" id="hs-locked">${JSON.stringify(payload)}</script>
<script>${UNLOCK_SCRIPT}</script></body></html>`;
}
