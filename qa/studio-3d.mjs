// 3D モデル in a real browser (WebGL): a stock model and a GLB file from this device inserted (each with its poster and
// drawn live), the 3D モデル tab (a view from the gallery, the turning handle, パンとズーム, the rotation fields, reset),
// the 3D animations (ターンテーブル turns the model while presenting) and the exported file drawing the model offline.
// Usage: node qa/studio-3d.mjs [--base=http://127.0.0.1:8787]   (with `npm start` running)
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
process.env.PLAYWRIGHT_DISABLE_FORCED_CHROMIUM_PROXIED_LOOPBACK = "1";
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || "/opt/node22/lib/node_modules/playwright");
const root = fileURLToPath(new URL("..", import.meta.url));
const args = Object.fromEntries(process.argv.slice(2).map((arg) => arg.replace(/^--/, "").split("=")));
const base = args.base || "http://127.0.0.1:8787";
const outDir = join(root, "qa", "out");
await mkdir(outDir, { recursive: true });

/** A small GLB: a cube with one colour (positions, normals, indices; glTF 2.0). */
function cubeGlb() {
  const faces = [[0, 0, 1], [0, 0, -1], [1, 0, 0], [-1, 0, 0], [0, 1, 0], [0, -1, 0]];
  const pos = [];
  const nrm = [];
  const idx = [];
  for (const [nx, ny, nz] of faces) {
    const [ux, uy, uz] = Math.abs(ny) === 1 ? [1, 0, 0] : [0, 1, 0];
    const [vx, vy, vz] = [ny * uz - nz * uy, nz * ux - nx * uz, nx * uy - ny * ux];
    const start = pos.length / 3;
    for (const [a, b] of [[-1, -1], [1, -1], [1, 1], [-1, 1]]) {
      pos.push((nx + a * ux + b * vx) / 2, (ny + a * uy + b * vy) / 2, (nz + a * uz + b * vz) / 2);
      nrm.push(nx, ny, nz);
    }
    idx.push(start, start + 1, start + 2, start, start + 2, start + 3);
  }
  const posBuf = Buffer.from(new Float32Array(pos).buffer);
  const nrmBuf = Buffer.from(new Float32Array(nrm).buffer);
  const idxBuf = Buffer.from(new Uint16Array(idx).buffer);
  const bin = Buffer.concat([posBuf, nrmBuf, idxBuf, Buffer.alloc((4 - ((posBuf.length + nrmBuf.length + idxBuf.length) % 4)) % 4)]);
  const json = {
    asset: { version: "2.0", generator: "studio-3d qa" },
    scene: 0, scenes: [{ nodes: [0] }], nodes: [{ mesh: 0 }],
    meshes: [{ primitives: [{ attributes: { POSITION: 0, NORMAL: 1 }, indices: 2, material: 0 }] }],
    materials: [{ pbrMetallicRoughness: { baseColorFactor: [0.85, 0.75, 0.6, 1], metallicFactor: 0, roughnessFactor: 0.6 } }],
    buffers: [{ byteLength: bin.length }],
    bufferViews: [{ buffer: 0, byteOffset: 0, byteLength: posBuf.length }, { buffer: 0, byteOffset: posBuf.length, byteLength: nrmBuf.length }, { buffer: 0, byteOffset: posBuf.length + nrmBuf.length, byteLength: idxBuf.length }],
    accessors: [
      { bufferView: 0, componentType: 5126, count: pos.length / 3, type: "VEC3", min: [-0.5, -0.5, -0.5], max: [0.5, 0.5, 0.5] },
      { bufferView: 1, componentType: 5126, count: nrm.length / 3, type: "VEC3" },
      { bufferView: 2, componentType: 5123, count: idx.length, type: "SCALAR" },
    ],
  };
  let jsonBuf = Buffer.from(JSON.stringify(json));
  jsonBuf = Buffer.concat([jsonBuf, Buffer.alloc((4 - (jsonBuf.length % 4)) % 4, 0x20)]);
  const header = Buffer.alloc(12);
  header.write("glTF", 0);
  header.writeUInt32LE(2, 4);
  header.writeUInt32LE(12 + 8 + jsonBuf.length + 8 + bin.length, 8);
  const chunk = (buf, type) => { const head = Buffer.alloc(8); head.writeUInt32LE(buf.length, 0); head.write(type, 4); return Buffer.concat([head, buf]); };
  return Buffer.concat([header, chunk(jsonBuf, "JSON"), chunk(bin, "BIN\0")]);
}
const glbFile = join(outDir, "3d-cube.glb");
await writeFile(glbFile, cubeGlb());

const browserArgs = process.env.PROXY_CA_SPKI ? [`--ignore-certificate-errors-spki-list=${process.env.PROXY_CA_SPKI}`] : [];
const browser = await chromium.launch({ executablePath: process.env.CHROME_PATH || "/opt/pw-browsers/chromium-1194/chrome-linux/chrome", proxy: process.env.HTTPS_PROXY ? { server: process.env.HTTPS_PROXY, bypass: "127.0.0.1,localhost" } : undefined, args: browserArgs });
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: true });
const page = await context.newPage();
const errors = [];
page.on("pageerror", (error) => errors.push(`pageerror: ${error.message}`));
page.on("console", (message) => { if (message.type() === "error" && !/ERR_TUNNEL|ytimg|ERR_CERT|fonts\.g|GPU stall|WebGL/.test(`${message.text()} ${message.location().url}`)) errors.push(`console: ${message.text()} ${message.location().url}`.trim()); });
page.on("dialog", (dialog) => dialog.accept());
const shot = async (name, p = page) => { const file = join(outDir, `3d-${name}.png`); await p.screenshot({ path: file }); console.log("saved", file); };
const step = async (label, fn) => {
  try { await page.waitForTimeout(150); await fn(); console.log("ok  ", label); } catch (error) { errors.push(`${label}: ${error.message}`); console.log("FAIL", label, error.message); await shot(`fail-${errors.length}`); }
  if (await page.isVisible("#presenter .hs-player")) { await page.keyboard.press("Escape"); await page.waitForTimeout(300); }
  if (await page.isVisible("dialog[open]")) await page.keyboard.press("Escape");
  if (await page.isVisible(".rb-pop")) await page.keyboard.press("Escape");
};
const assert = (ok, message) => { if (!ok) throw new Error(message); };
const objects = () => page.evaluate(() => JSON.parse(JSON.stringify(window.__hpptx.slide()?.elements ?? [])));
const tab = (label) => page.click(`.rb-tabs [role=tab]:text-is("${label}")`);
const ribbonBtn = async (label) => {
  const sel = `.rb-btn:not(.rb-folded):has-text("${label}")`;
  const direct = page.locator(`.rb-body ${sel}`).first();
  if ((await direct.count()) && (await direct.isVisible())) return direct.click();
  for (const folded of await page.locator(".rb-body .rb-folded").all()) {
    await folded.click();
    const inside = page.locator(`.rb-pop .rb-fold ${sel}`).first();
    if (await inside.count()) return inside.click();
    await page.keyboard.press("Escape");
  }
  throw new Error(`no ribbon button "${label}"`);
};
const menuItem = (label) => page.locator(`.rb-pop button:has-text("${label}")`).first().click();
/** How many pixels of a canvas or picture (in `scope`'s page) are drawn (not transparent). */
const drawn = (p, selector) => p.evaluate(async (sel) => {
  const el = document.querySelector(sel);
  if (!el) return -1;
  const src = el.tagName === "CANVAS" ? el.toDataURL("image/png") : el.getAttribute("src");
  const img = new Image();
  await new Promise((resolve, reject) => { img.onload = resolve; img.onerror = reject; img.src = src; });
  const c = document.createElement("canvas");
  c.width = img.naturalWidth;
  c.height = img.naturalHeight;
  const g = c.getContext("2d");
  g.drawImage(img, 0, 0);
  const d = g.getImageData(0, 0, c.width, c.height).data;
  let n = 0;
  for (let i = 3; i < d.length; i += 4) if (d[i] > 20) n += 1;
  return n / (c.width * c.height);
}, selector);
const liveCanvas = (id) => `#stageBody .hs-obj[data-el="${id}"] .hs-model.is-live canvas`;

await page.goto(base);
await page.evaluate(() => localStorage.clear());
await page.goto(base);
await page.click("#sampleDeckBtn");
await page.waitForSelector(".film-item");
await page.waitForSelector("#ribbon:not([hidden]) .rb-tabs");
await tab("挿入");
await ribbonBtn("新しいスライド");
await menuItem("白紙");
await page.waitForSelector(".slide-wrap .hs-slide");
let store = null;

await step("挿入 → 3D モデル → ストック: 店舗, with its poster, drawn live on the stage", async () => {
  await tab("挿入");
  await ribbonBtn("3D");
  await menuItem("店舗");
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).some((o) => o.kind === "model"), null, { timeout: 30000 });
  store = (await objects()).find((o) => o.kind === "model");
  assert(store.src === "builtin:store" && store.name === "店舗", `a stock model: ${store.src}`);
  assert(/^data:image\/png;base64,/.test(store.poster || ""), "a poster");
  await page.waitForSelector(liveCanvas(store.id), { timeout: 20000 });
  assert((await drawn(page, liveCanvas(store.id))) > 0.05, "the live drawing shows the model");
  assert(await page.isVisible('.rb-tabs [role=tab]:text-is("3D モデル")'), "the 3D モデル tab comes");
  await shot("stock");
});

await step("3D モデル ビュー: 背面 from the gallery (a new poster); the rotation field shows it", async () => {
  await tab("3D モデル");
  await ribbonBtn("ビュー");
  await page.click('.m3d-view[data-view="back"]');
  await page.waitForFunction((id) => window.__hpptx.slide().elements.find((o) => o.id === id)?.view?.yaw === 180, store.id, { timeout: 10000 });
  const now = (await objects()).find((o) => o.id === store.id);
  assert(now.poster !== store.poster, "the poster is drawn again");
  assert((await page.inputValue('input[aria-label="横の回転（度）"]')) === "180", "the field follows");
  store = now;
});

await step("the turning handle in the middle turns it (one ⌘Z back); 縦の回転 tilts it", async () => {
  const handle = page.locator(".m3d-rotate");
  const box = await handle.boundingBox();
  await page.mouse.move(box.x + box.width / 2, box.y + box.height / 2);
  await page.mouse.down();
  await page.mouse.move(box.x + box.width / 2 + 120, box.y + box.height / 2 + 40, { steps: 8 });
  await page.mouse.up();
  await page.waitForFunction((id) => { const v = window.__hpptx.slide().elements.find((o) => o.id === id)?.view; return v && v.yaw !== 180; }, store.id, { timeout: 10000 });
  const turned = (await objects()).find((o) => o.id === store.id).view;
  assert(turned.pitch > 5, `tilted too: ${JSON.stringify(turned)}`);
  await page.keyboard.press("Control+z");
  await page.waitForTimeout(400);
  assert((await objects()).find((o) => o.id === store.id).view.yaw === 180, "⌘Z brings the view back");
  await page.fill('input[aria-label="縦の回転（度）"]', "30");
  await page.press('input[aria-label="縦の回転（度）"]', "Enter");
  await page.waitForFunction((id) => window.__hpptx.slide().elements.find((o) => o.id === id)?.view?.pitch === 30, store.id, { timeout: 10000 });
});

await step("パンとズーム: the magnifier zooms in, dragging inside moves it in its box; reset brings the first view", async () => {
  await tab("3D モデル");
  await ribbonBtn("パンと");
  await page.waitForSelector(".m3d-zoom");
  const z = await page.locator(".m3d-zoom").boundingBox();
  await page.mouse.move(z.x + z.width / 2, z.y + z.height / 2);
  await page.mouse.down();
  await page.mouse.move(z.x + z.width / 2, z.y - 90, { steps: 6 });
  await page.mouse.up();
  await page.waitForFunction((id) => (window.__hpptx.slide().elements.find((o) => o.id === id)?.view?.zoom || 1) > 1.3, store.id, { timeout: 10000 });
  const pan = await page.locator(".m3d-pan").boundingBox();
  await page.mouse.move(pan.x + pan.width / 2, pan.y + pan.height / 2);
  await page.mouse.down();
  await page.mouse.move(pan.x + pan.width / 2 + 60, pan.y + pan.height / 2, { steps: 6 });
  await page.mouse.up();
  await page.waitForFunction((id) => (window.__hpptx.slide().elements.find((o) => o.id === id)?.view?.panX || 0) < -0.1, store.id, { timeout: 10000 });
  await shot("pan-zoom");
  await ribbonBtn("パンと");
  await ribbonBtn("のリセット");
  await menuItem("3D モデルのリセット");
  await page.waitForFunction((id) => { const v = window.__hpptx.slide().elements.find((o) => o.id === id)?.view; return v && v.yaw === -30 && v.pitch === 18 && !v.zoom && !v.panX; }, store.id, { timeout: 10000 });
});

let cube = null;
await step("挿入 → 3D モデル → このデバイス: a GLB file, kept in this browser, with its poster and live drawing", async () => {
  await tab("挿入");
  const [chooser] = await Promise.all([page.waitForEvent("filechooser"), (async () => { await ribbonBtn("3D"); await menuItem("このデバイス"); })()]);
  await chooser.setFiles(glbFile);
  await page.waitForFunction(() => (window.__hpptx.slide().elements || []).filter((o) => o.kind === "model").length === 2, null, { timeout: 30000 });
  cube = (await objects()).filter((o) => o.kind === "model").at(-1);
  assert(/^idb:/.test(cube.src) && cube.fileName === "3d-cube.glb", `from the file: ${cube.src}`);
  assert((await drawn(page, `#stageBody .hs-obj[data-el="${cube.id}"] .hs-model-poster`)) > 0.05, "its poster shows the cube");
  await page.waitForSelector(liveCanvas(cube.id), { timeout: 20000 });
  assert((await drawn(page, liveCanvas(cube.id))) > 0.05, "drawn live");
  await shot("glb");
});

await step("3D のアニメーション: ターンテーブル (only offered for 3D models) turns the model while presenting", async () => {
  await page.click(`#stageBody .hs-obj[data-el="${cube.id}"]`, { force: true });
  await tab("アニメーション");
  await page.locator('.rb-body .rb-btn[title*="すべての効果"]').first().click();
  await page.waitForSelector(".an-model-block");
  const names = await page.locator(".an-model-block .an-fx small").allTextContents();
  assert(["到着", "ターンテーブル", "スイング", "ジャンプしてターン", "退出"].every((n) => names.includes(n)), `3D effects: ${names}`);
  await page.click('.an-model-block .an-fx[data-fx="em:turntable3d"]');
  await page.waitForFunction(() => (window.__hpptx.slide().timeline || []).some((e) => e.fx === "turntable3d"), null, { timeout: 5000 });
  await page.keyboard.press("Shift+F5");
  await page.waitForSelector("#presenter .hs-player .hs-model");
  await page.waitForTimeout(800);
  await page.keyboard.press("ArrowRight");
  await page.waitForTimeout(700);
  const yaw = await page.evaluate((id) => Number(getComputedStyle(document.querySelector(`#presenter .hs-obj[data-el="${id}"] .hs-model`)).getPropertyValue("--m3d-yaw")), cube.id);
  assert(yaw > 20 && yaw < 340, `turning while presenting: ${yaw}`);
  await page.waitForSelector(`#presenter .hs-obj[data-el="${cube.id}"] .hs-model.is-live canvas`, { timeout: 10000 });
  await shot("turntable");
  await page.keyboard.press("Escape");
  await page.waitForTimeout(400);
});

await step("the exported file carries three.js and draws the models offline", async () => {
  const [download] = await Promise.all([page.waitForEvent("download"), page.click("#downloadBtn").then(async () => { if (await page.isVisible("#exportCheckDialog[open]")) await page.click("#exportCheckGoBtn"); })]);
  const file = join(outDir, "3d-export.html");
  await download.saveAs(file);
  const html = await readFile(file, "utf8");
  assert(/id="hs-three"/.test(html) && /GLTFLoader/.test(html), "three.js inside");
  const index = await page.evaluate(() => window.__hpptx.deck().slides.findIndex((s) => (s.elements || []).some((o) => o.kind === "model")));
  const viewer = await context.newPage();
  const viewerErrors = [];
  viewer.on("pageerror", (error) => viewerErrors.push(error.message));
  await viewer.goto(`file://${file}#${index + 1}`);
  await viewer.waitForSelector(".hs-player .hs-model.is-live canvas", { timeout: 20000 });
  await viewer.waitForTimeout(500);
  assert((await drawn(viewer, ".hs-player .hs-model.is-live canvas")) > 0.02, "drawn in the exported file");
  await shot("export", viewer);
  await viewer.close();
  assert(!viewerErrors.length, `exported file errors: ${viewerErrors.join(" / ")}`);
});

console.log(errors.length ? `errors:\n${errors.join("\n")}` : "no errors");
await browser.close();
process.exit(errors.length ? 1 : 0);
