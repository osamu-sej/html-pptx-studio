/*
 * HTML PPTX Studio — motion runtime and presentation player.
 * Plays a slide's entrance and builds, counts numbers up, grows charts, lifts items on hover, opens
 * "click for details" cards, plays video, and drives a full-screen presentation with a presenter view.
 * Shared by the studio's presenter and exported HTML files (no dependencies).
 */
(function (root) {
  "use strict";
  const E = root.SlideEngine;
  if (!E) throw new Error("engine.js must load before motion.js");
  const { h, s } = E;

  const reduced = () => typeof matchMedia === "function" && matchMedia("(prefers-reduced-motion: reduce)").matches;

  // ---------------------------------------------------------------- numbers

  function formatCount(value, decimals, comma) {
    const fixed = Math.abs(value).toFixed(decimals);
    if (!comma) return fixed;
    const [int, frac] = fixed.split(".");
    return `${Number(int).toLocaleString("en-US")}${frac ? `.${frac}` : ""}`;
  }

  /** "1,234.5" counts up from 0 in the same format; the final frame is the original text. */
  function countUp(el, duration = 1100) {
    const target = el.dataset.count;
    if (!target || el.dataset.counting) return;
    const clean = target.replace(/[,\s]/g, "").replace("−", "-").replace(/^[▲+]/, "").replace(/^▼/, "-");
    const value = parseFloat(clean);
    if (!Number.isFinite(value) || value === 0) { el.textContent = target; return; }
    const decimals = (clean.split(".")[1] || "").length;
    const comma = target.includes(",");
    const neg = /^[-−▼]/.test(target) ? target[0] : "";
    const pos = /^[+▲]/.test(target) ? target[0] : "";
    const start = performance.now();
    el.dataset.counting = "1";
    const tick = (now) => {
      const p = Math.min(1, (now - start) / duration);
      const eased = 1 - (1 - p) ** 3;
      el.textContent = p >= 1 ? target : `${value < 0 ? neg || "-" : pos}${formatCount(value * eased, decimals, comma)}`;
      if (p < 1 && el.isConnected) requestAnimationFrame(tick);
      else delete el.dataset.counting;
    };
    el.textContent = `${pos}${formatCount(0, decimals, comma)}`;
    requestAnimationFrame(tick);
  }

  function countWithin(scope, delay = 380) {
    const slide = scope.closest?.(".hs-slide") || scope;
    if (!slide.classList.contains("hs-numbers") || reduced()) return;
    const els = [...scope.querySelectorAll(".hs-count")];
    if (!els.length) return;
    for (const el of els) el.textContent = el.dataset.count.replace(/\d/g, "0").replace(/0+(?=[,.]|$)/, "0");
    setTimeout(() => els.forEach((el) => countUp(el)), delay);
  }

  // ---------------------------------------------------------------- kinetic type

  // The big lines of a slide: cover, chapter, hero, statement and closing text, and a titled slide's title
  const KINETIC_TARGETS = ".hs-cover-title, .hs-section-title, .hs-hero-title, .hs-statement-text, .hs-closing-message, .hs-closing-title, .hs-frame > .hs-head > .hs-title, .hs-frame > .hs-head > .hs-sej-titlebar > .hs-title, .hs-frame > .hs-sej-titlebar > .hs-title";
  // Characters that may not start a line ride with the unit before; opening brackets ride with the unit after.
  const NO_START = /^[、。，．,.)）」』】〕〉》！？!?:：;；ー〜…‥・ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ々ゝゞ%％]+$/u;
  const NO_END = /^[(（「『【〔〈《]+$/u;
  // [the whole line's time budget, the longest gap between two units] in ms, per style.
  const BUDGET = {
    mask: [900, 70], words: [1000, 90], chars: [1100, 45], type: [1700, 60], scramble: [1300, 55],
    wave: [1100, 45], zoom: [1000, 110], flip: [1100, 50], slide: [900, 90],
  };
  // Styles that move word by word (the rest move character by character).
  const BY_WORD = new Set(["mask", "words", "zoom", "slide"]);

  function segments(text, granularity) {
    try { return [...new Intl.Segmenter("ja", { granularity }).segment(text)].map((part) => part.segment); } catch { return Array.from(text); }
  }

  /** Text → units to animate ({ space } keeps spaces and line breaks as they are). */
  function unitsOf(text, mode) {
    const units = [];
    let carry = "";
    for (const part of segments(text, BY_WORD.has(mode) ? "word" : "grapheme")) {
      if (/^\s+$/.test(part)) {
        if (carry) { units.push(carry); carry = ""; }
        units.push({ space: part });
        continue;
      }
      if (NO_END.test(part)) { carry += part; continue; }
      const piece = carry + part;
      carry = "";
      if (NO_START.test(part) && typeof units[units.length - 1] === "string") units[units.length - 1] += piece;
      else units.push(piece);
    }
    if (carry) units.push(carry);
    return units;
  }

  /** Wrap every unit of `el`'s text in a span (inside <em> too). Returns the number of units. */
  function splitKinetic(el, mode) {
    if (el.classList.contains("hs-kin")) return el.querySelectorAll(".hs-k").length;
    const doc = el.ownerDocument;
    const perChar = !BY_WORD.has(mode);
    let n = 0;
    const wrap = (text) => {
      const frag = doc.createDocumentFragment();
      let word = null; // Latin words stay unbroken when split per character
      for (const unit of unitsOf(text, mode)) {
        if (typeof unit !== "string") { word = null; frag.append(unit.space); continue; }
        const k = doc.createElement("span");
        k.className = "hs-k";
        k.textContent = unit;
        k.style.setProperty("--k", String(n));
        n += 1;
        let node = k;
        if (mode === "mask") { node = doc.createElement("span"); node.className = "hs-km"; node.append(k); }
        if (perChar && /^[A-Za-z0-9.,:%+\-]+$/.test(unit)) {
          if (!word) { word = doc.createElement("span"); word.className = "hs-kw"; frag.append(word); }
          word.append(node);
        } else {
          word = null;
          frag.append(node);
        }
      }
      return frag;
    };
    const walk = (node) => {
      for (const child of [...node.childNodes]) {
        if (child.nodeType === 3) { if (child.data) child.replaceWith(wrap(child.data)); }
        else if (child.nodeType === 1 && child.namespaceURI === "http://www.w3.org/1999/xhtml" && !child.classList.contains("hs-detail-badge") && !child.classList.contains("hs-drill-badge")) walk(child);
      }
    };
    walk(el);
    el.classList.add("hs-kin");
    return n;
  }

  /** Split the slide's big lines and time them (called by play()). */
  function kinetic(slide) {
    const mode = slide.dataset.kinetic;
    if (!BUDGET[mode]) return;
    for (const el of slide.querySelectorAll(KINETIC_TARGETS)) {
      const n = splitKinetic(el, mode);
      if (!n) continue;
      const [budget, cap] = BUDGET[mode];
      const d = parseFloat(el.style.getPropertyValue("--d")) || 0;
      el.style.setProperty("--kst", `${Math.round(Math.min(cap, budget / n))}ms`);
      el.style.setProperty("--kd", `${Math.round(el.classList.contains("hs-enter") ? d * 90 + 120 : 60)}ms`);
      el.style.setProperty("--kn", String(n));
      const units = el.querySelectorAll(".hs-k");
      units[units.length - 1]?.classList.add("hs-k-last");
    }
    if (mode === "scramble") scramble(slide);
  }

  const GLYPHS = {
    wide: "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン",
    upper: "ABCDEFGHIJKLMNOPQRSTUVWXYZ",
    lower: "abcdefghijklmnopqrstuvwxyz",
    digit: "0123456789",
  };
  const poolOf = (ch) => (/[぀-ヿ㐀-鿿！-～]/u.test(ch) ? GLYPHS.wide : /[A-Z]/.test(ch) ? GLYPHS.upper : /[a-z]/.test(ch) ? GLYPHS.lower : /\d/.test(ch) ? GLYPHS.digit : null);

  /** "Decode": each unit flickers through random glyphs of its own kind, then settles, left to right. */
  function scramble(slide) {
    const units = [...slide.querySelectorAll(".hs-kin .hs-k")].map((el) => {
      el.classList.remove("on", "done");
      const host = el.closest(".hs-kin");
      const at = (parseFloat(host.style.getPropertyValue("--kd")) || 0) + Number(el.style.getPropertyValue("--k") || 0) * (parseFloat(host.style.getPropertyValue("--kst")) || 50);
      if (el.dataset.t == null) el.dataset.t = el.textContent;
      const text = el.dataset.t;
      return { el, text, chars: Array.from(text), at, until: at + 420, swapped: 0 };
    });
    if (!units.length) return;
    const start = performance.now();
    const finish = () => { for (const u of units) { u.el.textContent = u.text; u.el.classList.remove("on"); u.el.classList.add("done"); } };
    const last = Math.max(...units.map((u) => u.until));
    const safety = setTimeout(finish, last + 1500);
    const tick = (now) => {
      if (!slide.isConnected) { clearTimeout(safety); return; }
      const t = now - start;
      for (const u of units) {
        if (u.el.classList.contains("done")) continue;
        if (t >= u.until) { u.el.textContent = u.text; u.el.classList.remove("on"); u.el.classList.add("done"); continue; }
        if (t < u.at) continue;
        u.el.classList.add("on");
        if (now - u.swapped > 55) {
          u.swapped = now;
          u.el.textContent = u.chars.map((ch) => { const pool = poolOf(ch); return pool ? pool[Math.floor(Math.random() * pool.length)] : ch; }).join("");
        }
      }
      if (t < last) requestAnimationFrame(tick);
      else { clearTimeout(safety); finish(); }
    };
    requestAnimationFrame(tick);
  }

  // ---------------------------------------------------------------- entrance & builds

  // Click steps: the layout's own build first (data-lsteps of them), then one per group of the slide's
  // animations (animate.js). data-steps counts both.
  const stepsOf = (slide) => Number(slide.dataset.steps || 0);
  const layoutSteps = (slide) => Number(slide.dataset.lsteps ?? slide.dataset.steps ?? 0);

  /**
   * Start a slide's entrance. With a click build, `step` groups are already shown (going back shows all).
   * Returns how long the slide's first animations take (ms), for a preview that waits for them.
   */
  function play(slide, { step = 0, animate = true, delay = 0 } = {}) {
    slide.classList.remove("hs-play");
    const lsteps = layoutSteps(slide);
    const lstep = Math.min(step, lsteps);
    const toggling = measuresBuild(slide, lstep);
    const click = slide.dataset.build === "click" && !toggling;
    for (const el of slide.querySelectorAll("[data-g]")) {
      el.classList.remove("hs-in");
      el.classList.toggle("hs-hidden", click && Number(el.dataset.g) >= lstep);
    }
    spotlight(slide, lstep);
    const timed = E.animStart?.(slide, Math.max(0, step - lsteps), { animate: animate && !reduced(), delay }) || 0;
    if (!animate || reduced()) return timed;
    kinetic(slide);
    void slide.offsetWidth;
    slide.classList.add("hs-play");
    if (click) {
      countWithin(slide.querySelector(".hs-head") || slide, 300);
      for (const el of slide.querySelectorAll("[data-g]")) if (Number(el.dataset.g) < lstep) countWithin(el, 0);
    } else countWithin(slide);
    return timed;
  }

  /** A spotlight build keeps everything on screen and puts group `step - 1` in focus (0: nothing in focus). */
  function spotlight(slide, step) {
    if (slide.dataset.build !== "spotlight") return false;
    slide.classList.toggle("hs-spotting", step > 0);
    for (const el of slide.querySelectorAll("[data-g]")) el.classList.toggle("hs-spot", Number(el.dataset.g) === step - 1);
    return true;
  }

  /**
   * On a "gap" slide a click build turns the measures on one per click (everything stays on screen): step 0
   * shows the gap, the last step shows it filled. Starting a slide sets every switch; a click only turns the
   * next measure on, so measures the presenter switched by hand stay as they are. False for any other slide.
   */
  function measuresBuild(slide, step, { exact = true } = {}) {
    if (slide.dataset.build !== "click" || !slide.querySelector("[data-measure]")) return false;
    for (const el of slide.querySelectorAll("[data-measure]")) {
      const n = Number(el.dataset.measure);
      if (exact) el.classList.toggle("is-on", n < step);
      else if (n === step - 1) el.classList.add("is-on");
    }
    E.gapUpdate(slide);
    return true;
  }

  /** Show the next click step (1-based: step 1 shows group 0). Returns how long its animations take (ms). */
  function reveal(slide, step) {
    const lsteps = layoutSteps(slide);
    if (step > lsteps) return E.animStep?.(slide, step - lsteps, { animate: !reduced() }) || 0;
    if (measuresBuild(slide, step, { exact: false })) return 0;
    // In a spotlight build the item in focus counts its figure up again.
    if (spotlight(slide, step)) { for (const el of slide.querySelectorAll(".hs-spot")) countWithin(el, 120); return 0; }
    for (const el of slide.querySelectorAll(`[data-g="${step - 1}"]`)) {
      el.classList.remove("hs-hidden", "hs-in");
      void el.getBoundingClientRect();
      if (!reduced()) el.classList.add("hs-in");
      countWithin(el, 120);
    }
    return 0;
  }

  // ---------------------------------------------------------------- Lottie (motion graphics made in After Effects, LottieFiles…)

  // The studio points this at its copy of lottie-web; exported files carry the library inline instead.
  let lottieLoading = null;
  function lottieLib(doc) {
    const win = doc.defaultView || root;
    if (win.lottie) return Promise.resolve(win.lottie);
    if (!E.lottieUrl) return Promise.reject(new Error("Lottie のプレーヤーがありません"));
    if (!lottieLoading) {
      lottieLoading = new Promise((resolve, reject) => {
        const script = doc.createElement("script");
        script.src = E.lottieUrl;
        script.async = true;
        script.onload = () => (win.lottie ? resolve(win.lottie) : reject(new Error("Lottie を読み込めません")));
        script.onerror = () => { lottieLoading = null; reject(new Error("Lottie を読み込めません")); };
        doc.head.append(script);
      });
    }
    return lottieLoading;
  }

  const lottieFiles = new Map();
  function lottieJson(url) {
    if (!lottieFiles.has(url)) {
      lottieFiles.set(url, fetch(url).then((response) => {
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        return response.json();
      }).catch((error) => { lottieFiles.delete(url); throw error; }));
    }
    return lottieFiles.get(url);
  }

  /**
   * Load every Lottie animation in `scope`. play: start the ones set to autoplay (presentations);
   * otherwise hold one frame (`frame` 0–1 of the animation) — the editor, prints and reduced motion.
   */
  function mountLottie(scope, { play = true, frame = 0.5 } = {}) {
    const hosts = [...scope.querySelectorAll(".hs-lottie-host")];
    return Promise.all(hosts.map(async (host) => {
      try {
        if (!host.hsAnim) {
          const [lib, data] = await Promise.all([lottieLib(host.ownerDocument), lottieJson(host.dataset.src)]);
          if (!host.isConnected || host.hsAnim) return;
          host.hsAnim = lib.loadAnimation({
            container: host, renderer: "svg", loop: host.hasAttribute("data-loop"), autoplay: false,
            animationData: typeof structuredClone === "function" ? structuredClone(data) : JSON.parse(JSON.stringify(data)),
            rendererSettings: { preserveAspectRatio: host.dataset.fit === "cover" ? "xMidYMid slice" : "xMidYMid meet" },
          });
        }
        const anim = host.hsAnim;
        if (play && !reduced() && host.hasAttribute("data-autoplay")) { anim.goToAndPlay(0, true); return; }
        const hold = () => anim.goToAndStop(Math.max(0, Math.round(frame * (anim.totalFrames - 1))), true);
        if (anim.isLoaded) hold();
        else anim.addEventListener("DOMLoaded", hold);
      } catch {
        host.classList.add("hs-lottie-failed");
      }
    }));
  }

  function stopLottie(scope) {
    for (const host of scope.querySelectorAll(".hs-lottie-host")) {
      try { host.hsAnim?.destroy(); } catch { /* already gone */ }
      host.hsAnim = null;
    }
  }

  // ---------------------------------------------------------------- media

  function playMedia(slide, { sound = false, skip = null, noNarration = false } = {}) {
    for (const video of slide.querySelectorAll("video[data-autoplay]")) {
      video.muted = !sound || video.hasAttribute("data-muted");
      if (E.mediaPlay) E.mediaPlay(video, { fromStart: true });
      else video.play().catch(() => { video.muted = true; video.play().catch(() => {}); });
    }
    // Sounds that start with the slide (再生 →「自動」); one still playing from an earlier slide is not doubled.
    for (const audio of slide.querySelectorAll("audio[data-autoplay]")) {
      if (skip?.has(audio.closest("[data-el]")?.dataset.el)) continue;
      if (noNarration && audio.hasAttribute("data-narration")) continue;
      E.mediaPlay?.(audio, { fromStart: true });
    }
    for (const frame of slide.querySelectorAll("iframe[data-autoplay]")) {
      const send = () => frame.contentWindow?.postMessage(JSON.stringify({ event: "command", func: "playVideo", args: [] }), "*");
      frame.addEventListener("load", send, { once: true });
      send();
    }
    mountLottie(slide, { play: true, frame: 0 });
    // 3D models are drawn live while their slide is shown (models.js).
    E.mountModels?.(slide);
  }

  function stopMedia(slide) {
    E.animStop?.(slide);
    for (const media of slide.querySelectorAll("video, audio")) {
      if (E.mediaPause) E.mediaPause(media);
      else { try { media.pause(); } catch { /* detached */ } }
    }
    for (const frame of slide.querySelectorAll("iframe")) frame.contentWindow?.postMessage(JSON.stringify({ event: "command", func: "pauseVideo", args: [] }), "*");
    stopLottie(slide);
    E.stopModels?.(slide);
  }

  // ---------------------------------------------------------------- interaction (hover, tooltips, details, parallax)

  function slideScale(slide) {
    const rect = slide.getBoundingClientRect();
    return rect.width ? rect.width / E.W : 1;
  }

  function toSlide(slide, clientX, clientY) {
    const rect = slide.getBoundingClientRect();
    const k = slide.getBoundingClientRect().width / E.W || 1;
    return [(clientX - rect.left) / k, (clientY - rect.top) / k];
  }

  function rectIn(slide, el) {
    const base = slide.getBoundingClientRect();
    const r = el.getBoundingClientRect();
    const k = base.width / E.W || 1;
    return { x: (r.left - base.left) / k, y: (r.top - base.top) / k, w: r.width / k, h: r.height / k };
  }

  /** A detail's breakdown: bars when every value is a figure, otherwise a two-column list. */
  function evidenceRows(rows) {
    const nums = rows.map((row) => { const parts = E.numParts(row.value); return parts.num == null ? null : Number(parts.num.replace(/[^\d.]/g, "")); });
    const bars = nums.every((n) => n != null && Number.isFinite(n));
    const max = bars ? Math.max(...nums, 0) || 1 : 1;
    return h("ul", { class: ["hs-ev-rows", bars ? "bars" : ""] }, rows.map((row, i) => h("li", { style: { "--i": i } },
      h("span", { class: "hs-ev-label" }, E.strip(row.label)),
      bars ? h("span", { class: "hs-ev-track" }, h("i", { style: { width: `${(nums[i] / max) * 100}%` } })) : null,
      h("span", { class: "hs-ev-value" }, E.strip(row.value)))));
  }

  /** The words that name an item (its title), without its badges. */
  function nameOf(el) {
    if (!el) return "";
    const copy = el.cloneNode(true);
    copy.querySelectorAll(".hs-detail-badge, .hs-drill-badge").forEach((badge) => badge.remove());
    const title = copy.querySelector(".hs-card-title, .hs-row-title, .hs-mlabel, .hs-label, .hs-tree-branch, .lab, .hs-rank-label");
    return E.strip((title || copy).textContent).replace(/\s+/g, " ").slice(0, 28);
  }

  /** Wire a presented slide. Returns a function that removes everything again. */
  function activate(slide, { details = [], elements = [], onOpen, onClose, onDrill, onControl } = {}) {
    const overlay = slide.querySelector(".hs-overlay");
    let hotKey = null;
    let tip = null;
    // インタラクション (what only HTML does with an object): the object under the mouse, the one zoomed into, the
    // one in the spotlight.
    let hotObj = null;
    let zoomed = null;
    let spotted = null;
    const objNode = (id) => slide.querySelector(`.hs-obj[data-el="${E.cssEscape(id)}"]`);
    const objData = (node) => elements.find((o) => o.id === node?.dataset.el);
    // An object and the members of its group that answer the mouse the same way move together.
    const mates = (node) => {
      const group = node.dataset.group;
      return group ? [...slide.querySelectorAll(`.hs-obj[data-group="${E.cssEscape(group)}"]`)].filter((n) => n.dataset.hover === node.dataset.hover) : [node];
    };
    const setHotObj = (node) => {
      if (node === hotObj) return;
      if (hotObj) for (const n of slide.querySelectorAll(".hs-ix-hot")) { n.classList.remove("hs-ix-hot"); n.style.removeProperty("--rx"); n.style.removeProperty("--ry"); }
      slide.querySelector(".hs-objects")?.classList.remove("hs-ix-focusing");
      hotObj = node;
      if (!node) return;
      for (const n of mates(node)) n.classList.add("hs-ix-hot");
      if (node.dataset.hover === "focus") slide.querySelector(".hs-objects")?.classList.add("hs-ix-focusing");
    };
    const showObjTip = (node) => {
      if (!tip) { tip = h("div", { class: "hs-tip hs-tip-obj" }); overlay.append(tip); }
      tip.replaceChildren(...node.dataset.tip.split("\n").map((line) => h("span", {}, line)));
      const box = rectIn(slide, node);
      const above = box.y > 160;
      tip.classList.toggle("below", !above);
      tip.style.left = `${Math.max(160, Math.min(E.W - 160, box.x + box.w / 2))}px`;
      tip.style.top = `${above ? box.y : box.y + box.h}px`;
    };
    const setHot = (key) => {
      if (key === hotKey) return;
      for (const el of slide.querySelectorAll(".hs-hot")) el.classList.remove("hs-hot");
      slide.querySelectorAll(".hs-has-hot").forEach((el) => el.classList.remove("hs-has-hot"));
      hotKey = key;
      if (!key) return;
      const els = slide.querySelectorAll(`[data-item="${E.cssEscape(key)}"]`);
      for (const el of els) el.classList.add("hs-hot");
      els[0]?.closest(".hs-body")?.classList.add("hs-has-hot");
    };
    const showTip = (mark, event) => {
      if (!tip) { tip = h("div", { class: "hs-tip" }); overlay.append(tip); }
      // Value and name, and where the figure comes from when the slide says so.
      tip.replaceChildren(...[h("span", {}, mark.dataset.tip), slide.dataset.source ? h("small", {}, `出所：${slide.dataset.source}`) : null].filter(Boolean));
      const [x, y] = toSlide(slide, event.clientX, event.clientY);
      tip.style.left = `${Math.max(120, Math.min(E.W - 120, x))}px`;
      tip.style.top = `${Math.max(90, y)}px`;
      mark.closest(".hs-chart, .hs-wf, .hs-ochart")?.classList.add("hs-has-hot");
      for (const other of slide.querySelectorAll(".hs-mark.hs-hot")) other.classList.remove("hs-hot");
      mark.classList.add("hs-hot");
    };
    const hideTip = () => {
      tip?.remove(); tip = null;
      slide.querySelectorAll(".hs-mark.hs-hot").forEach((el) => el.classList.remove("hs-hot"));
      slide.querySelectorAll(".hs-chart.hs-has-hot, .hs-wf.hs-has-hot, .hs-ochart.hs-has-hot").forEach((el) => el.classList.remove("hs-has-hot"));
    };
    const onOver = (event) => {
      if (event.target.closest?.(".hs-popover")) return;
      if (zoomed) return;
      const mark = event.target.closest?.(".hs-mark");
      const obj = event.target.closest?.(".hs-obj");
      if (mark) showTip(mark, event);
      else if (obj?.dataset.tip && !obj.classList.contains("hs-ix-wait")) showObjTip(obj);
      else hideTip();
      setHotObj(obj?.dataset.hover && !obj.classList.contains("hs-ix-wait") ? obj : null);
      const itemEl = event.target.closest?.("[data-item]");
      setHot(itemEl && !itemEl.closest(".hs-popover") ? itemEl.dataset.item : null);
    };
    const onMove = (event) => {
      const mark = event.target.closest?.(".hs-mark");
      if (mark && tip) showTip(mark, event);
      // An object set to "tilt" leans towards the mouse (with the rest of its group).
      if (hotObj?.dataset.hover === "tilt") {
        const r = hotObj.getBoundingClientRect();
        if (r.width && r.height) {
          const ry = `${(((event.clientX - r.left) / r.width - 0.5) * 16).toFixed(2)}deg`;
          const rx = `${((0.5 - (event.clientY - r.top) / r.height) * 12).toFixed(2)}deg`;
          for (const n of slide.querySelectorAll(".hs-ix-hot")) { n.style.setProperty("--ry", ry); n.style.setProperty("--rx", rx); }
        }
      }
      // "Tilt" leans the item under the mouse towards it.
      if (slide.dataset.hover === "tilt") {
        const card = event.target.closest?.("[data-item]");
        if (card && !(card instanceof SVGElement) && !card.closest(".hs-popover")) {
          const r = card.getBoundingClientRect();
          if (r.width && r.height) {
            card.style.setProperty("--ry", `${(((event.clientX - r.left) / r.width - 0.5) * 12).toFixed(2)}deg`);
            card.style.setProperty("--rx", `${((0.5 - (event.clientY - r.top) / r.height) * 10).toFixed(2)}deg`);
          }
        }
      }
      const parallax = slide.querySelectorAll('.hs-media[data-motion="parallax"]');
      if (parallax.length) {
        const [x, y] = toSlide(slide, event.clientX, event.clientY);
        for (const media of parallax) {
          media.style.setProperty("--px", `${((x / E.W) - 0.5) * -3}%`);
          media.style.setProperty("--py", `${((y / E.H) - 0.5) * -3}%`);
        }
      }
    };
    const onLeave = () => { setHot(null); hideTip(); setHotObj(null); };
    const anchorOf = (target) => (target === "takeaway" ? slide.querySelector('[data-detail="takeaway"]')
      : target.startsWith("obj:") ? objNode(target.slice(4))
        : [...slide.querySelectorAll(`[data-item="${E.cssEscape(target)}"]`)].find((el) => !(el instanceof SVGElement)) || slide.querySelector(`[data-item="${E.cssEscape(target)}"]`));
    // Evidence (a breakdown, a source, assumptions) slides in from the right over a dimmed slide.
    const openPanel = (detail) => {
      const rows = (detail.rows || []).filter((row) => row && E.strip(row.label));
      const source = E.strip(detail.source || "");
      const note = E.strip(detail.note || "");
      const anchorName = detail.target.startsWith("obj:") ? E.strip(anchorOf(detail.target)?.textContent || "").replace(/\s+/g, " ").slice(0, 28) : nameOf(anchorOf(detail.target));
      const where = [E.strip(slide.querySelector(".hs-title")?.textContent || "") || E.strip(slide.dataset.title || ""), detail.target === "takeaway" ? "キーメッセージの根拠" : anchorName].filter(Boolean).join(" › ");
      const tabs = [rows.length ? ["rows", "内訳"] : null, source || note ? ["source", "出所と前提"] : null].filter(Boolean);
      const close = (event) => { event.stopPropagation(); closeDetail(); };
      const panel = h("aside", { class: "hs-evidence", role: "dialog", "aria-label": E.strip(detail.title || "根拠") },
        h("button", { class: "hs-popover-close", type: "button", "aria-label": "閉じる", onclick: close }, "×"),
        h("div", { class: "hs-ev-where" }, where),
        detail.title ? h("div", { class: "hs-ev-title" }, E.strip(detail.title)) : null,
        h("div", { class: "hs-ev-text" }, detail.html ? E.richNodes(detail.text) : E.rich(detail.text)),
        tabs.length > 1 ? h("div", { class: "hs-ev-tabs", role: "tablist" }, tabs.map(([key, label], i) => h("button", { type: "button", role: "tab", "data-tab": key, "aria-selected": String(i === 0) }, label))) : null,
        rows.length ? h("div", { class: "hs-ev-pane", "data-pane": "rows" }, evidenceRows(rows)) : null,
        source || note ? h("div", { class: "hs-ev-pane", "data-pane": "source", hidden: rows.length > 0 },
          note ? h("p", { class: "hs-ev-note" }, E.rich(detail.note)) : null,
          source ? h("p", { class: "hs-ev-source" }, `出所：${source}`) : null) : null,
        source && rows.length ? h("div", { class: "hs-ev-foot" }, `出所：${source}`) : null);
      panel.addEventListener("click", (event) => {
        event.stopPropagation();
        const tab = event.target.closest?.("[data-tab]");
        if (!tab) return;
        for (const other of panel.querySelectorAll("[data-tab]")) other.setAttribute("aria-selected", String(other === tab));
        for (const pane of panel.querySelectorAll("[data-pane]")) pane.hidden = pane.dataset.pane !== tab.dataset.tab;
      });
      overlay.append(h("div", { class: "hs-scrim dim", onclick: close }), panel);
      slide.dataset.detailOpen = detail.target;
      onOpen?.(detail);
    };
    const openDetail = (target) => {
      const detail = details.find((entry) => entry.target === target);
      if (!detail) return;
      closeDetail();
      if ((detail.rows || []).length || detail.source || detail.note) { openPanel(detail); return; }
      const anchor = anchorOf(target);
      const box = anchor ? rectIn(slide, anchor) : { x: 700, y: 300, w: 200, h: 100 };
      const width = 760;
      const scrim = h("div", { class: "hs-scrim", onclick: (event) => { event.stopPropagation(); closeDetail(); } });
      const pop = h("div", { class: "hs-popover", role: "dialog", "aria-label": E.strip(detail.title || "詳細") },
        h("button", { class: "hs-popover-close", type: "button", "aria-label": "閉じる", onclick: (event) => { event.stopPropagation(); closeDetail(); } }, "×"),
        detail.title ? h("div", { class: "hs-popover-title" }, E.strip(detail.title)) : null,
        h("div", { class: "hs-popover-text" }, detail.html ? E.richNodes(detail.text) : E.rich(detail.text)));
      pop.addEventListener("click", (event) => event.stopPropagation());
      const right = box.x + box.w + 32;
      const x = right + width < E.W - 48 ? right : Math.max(48, box.x - width - 32);
      pop.style.left = `${x}px`;
      pop.style.top = `${Math.max(72, Math.min(E.H - 420, box.y))}px`;
      overlay.append(scrim, pop);
      slide.dataset.detailOpen = target;
      onOpen?.(detail);
    };
    const closeDetail = () => {
      overlay.querySelectorAll(".hs-popover, .hs-scrim, .hs-evidence").forEach((el) => el.remove());
      if (slide.dataset.detailOpen) { delete slide.dataset.detailOpen; onClose?.(); }
      unzoom();
      unspot();
    };
    // 「拡大して見せる」: the slide moves in on the object (a camera), any click or Esc goes back.
    const zoomHint = h("div", { class: "hs-ix-zoomhint" }, "クリックで戻る");
    const zoomInto = (node) => {
      hideTip(); setHotObj(null);
      const b = rectIn(slide, node);
      const k = Math.min(3.4, (E.W * 0.84) / Math.max(1, b.w), (E.H * 0.84) / Math.max(1, b.h));
      if (k <= 1.04) return;
      const s = parseFloat(getComputedStyle(slide).getPropertyValue("--hs-s")) || slideScale(slide);
      const tx = (E.W / 2 - (b.x + b.w / 2) * k) * s;
      const ty = (E.H / 2 - (b.y + b.h / 2) * k) * s;
      zoomed = node;
      slide.classList.add("hs-ix-zoomed");
      slide.style.transition = "transform .75s cubic-bezier(.22,.61,.36,1)";
      slide.style.transform = `translate(${tx.toFixed(2)}px, ${ty.toFixed(2)}px) scale(${(s * k).toFixed(5)})`;
      slide.parentElement?.append(zoomHint);
    };
    function unzoom() {
      if (!zoomed) return;
      zoomed = null;
      slide.style.transform = "";
      zoomHint.remove();
      setTimeout(() => { if (!zoomed) { slide.style.transition = ""; slide.classList.remove("hs-ix-zoomed"); } }, 800);
    }
    // 「スポットライトを当てる」: everything else dims until the next click.
    const spot = (node) => {
      unspot();
      spotted = node;
      for (const n of mates(node).concat(node.dataset.group ? [...slide.querySelectorAll(`.hs-obj[data-group="${E.cssEscape(node.dataset.group)}"]`)] : [])) n.classList.add("hs-ix-spot");
      slide.classList.add("hs-ix-spotting");
    };
    function unspot() {
      if (!spotted) return;
      spotted = null;
      slide.classList.remove("hs-ix-spotting");
      for (const n of slide.querySelectorAll(".hs-ix-spot")) n.classList.remove("hs-ix-spot");
    }
    // 「ほかの部品を表示・非表示」: the objects a button names come and go; with "only", it works as a tab.
    const reveal = (node) => {
      const action = objData(node)?.action;
      if (!action?.targets) return;
      const targets = action.targets.map(objNode).filter(Boolean);
      const showing = targets.length && targets.every((n) => n.classList.contains("hs-ix-shown"));
      if (action.only) {
        for (const other of slide.querySelectorAll('.hs-obj[data-action="reveal"]')) {
          if (other === node) continue;
          other.classList.remove("is-on");
          for (const id of objData(other)?.action?.targets || []) if (!action.targets.includes(id)) objNode(id)?.classList.remove("hs-ix-shown");
        }
      }
      for (const n of targets) n.classList.toggle("hs-ix-shown", !showing);
      node.classList.toggle("is-on", !showing);
    };
    /** What a click on an object does that only HTML can (a link or a jump is the player's). True when handled. */
    const objectClick = (node) => {
      const type = node?.dataset.action;
      if (type === "zoom") { if (zoomed) unzoom(); else zoomInto(node); return true; }
      if (type === "flip") { node.classList.toggle("is-flipped"); return true; }
      if (type === "reveal") { reveal(node); return true; }
      if (type === "spot") { if (spotted === node) unspot(); else spot(node); return true; }
      return false;
    };
    const onClick = (event) => {
      // While zoomed in (or in a spotlight), a click comes back first.
      if (zoomed || (spotted && !event.target.closest?.('.hs-obj[data-action="spot"]'))) {
        event.stopPropagation();
        event.preventDefault();
        unzoom();
        unspot();
        return;
      }
      // Controls change the slide in place (and never advance it): a view of a ranking, a measure switched on or off.
      const control = event.target.closest?.(".hs-control");
      if (control && !event.target.closest(".hs-popover, .hs-evidence")) {
        event.stopPropagation();
        // A chart's legend key shows or hides its series.
        const key = event.target.closest("[data-series]");
        if (key) {
          const svg = key.closest("svg");
          const off = !key.classList.contains("hs-off");
          key.classList.toggle("hs-off", off);
          for (const el of svg.querySelectorAll(`[data-s="${key.dataset.series}"]`)) el.classList.toggle("hs-off", off);
          onControl?.("series");
        }
        const view = event.target.closest("[data-view]");
        if (view) { E.rankShow(view.closest(".hs-rank"), Number(view.dataset.view)); onControl?.("view"); }
        const measure = event.target.closest("[data-measure]");
        if (measure) { measure.classList.toggle("is-on"); E.gapUpdate(slide); onControl?.("measure"); }
        return;
      }
      const drill = event.target.closest?.("[data-drill]");
      if (drill && onDrill && !event.target.closest(".hs-popover")) {
        event.stopPropagation();
        event.preventDefault();
        onDrill(Number(drill.dataset.drill), drill);
        return;
      }
      const actor = event.target.closest?.(".hs-obj[data-action]");
      if (actor && !actor.classList.contains("hs-ix-wait") && !event.target.closest(".hs-popover") && objectClick(actor)) {
        event.stopPropagation();
        event.preventDefault();
        return;
      }
      const host = event.target.closest?.("[data-detail]");
      if (host && !event.target.closest(".hs-popover")) {
        event.stopPropagation();
        event.preventDefault();
        openDetail(host.dataset.detail);
        return;
      }
      // A bookmark's mark jumps there (and plays from it).
      const mark = event.target.closest?.(".hs-mark[data-t]");
      if (mark) {
        event.stopPropagation();
        event.preventDefault();
        const el = mark.closest(".hs-obj")?.querySelector("video, audio");
        if (el) { el.currentTime = Number(mark.dataset.t); if (el.paused) { el.muted = el.tagName === "VIDEO" && el.hasAttribute("data-muted"); E.mediaPlay?.(el); } }
        return;
      }
      const video = event.target.closest?.("video");
      if (video) {
        event.stopPropagation();
        if (video.paused) { video.muted = video.hasAttribute("data-muted"); if (E.mediaPlay) E.mediaPlay(video); else video.play().catch(() => {}); } else if (E.mediaPause) E.mediaPause(video); else video.pause();
      }
      // A sound's icon (or its bar's button) plays and pauses it; a click on the bar's track moves through it.
      const sound = event.target.closest?.(".hs-audio");
      if (sound && !sound.hasAttribute("data-hide-icon")) {
        event.stopPropagation();
        event.preventDefault();
        const audio = sound.querySelector("audio");
        const track = event.target.closest(".hs-audio-track");
        if (audio && track && E.mediaSpan) {
          const r = track.getBoundingClientRect();
          const [start, end] = E.mediaSpan(audio);
          if (Number.isFinite(end)) audio.currentTime = start + ((event.clientX - r.left) / r.width) * (end - start);
          if (audio.paused) E.mediaPlay(audio);
        } else if (audio) E.mediaToggle?.(audio);
        return;
      }
      // A Lottie animation that does not play by itself starts (or pauses) on click.
      const anim = event.target.closest?.(".hs-lottie-host")?.hsAnim;
      if (anim && !event.target.closest(".hs-lottie-host").hasAttribute("data-autoplay")) {
        event.stopPropagation();
        if (anim.isPaused) anim.play(); else anim.pause();
      }
      // Clicks inside the card reach its buttons (×); the card itself keeps them from advancing the slide.
      if (event.target.closest?.("iframe")) event.stopPropagation();
    };
    // Sliders recalculate as they move; letting go hands the keys back to the presentation.
    const onInput = (event) => { if (event.target.matches?.("input[data-sim]")) { E.simUpdate(slide); onControl?.("slider"); } };
    const onChange = (event) => { if (event.target.matches?.("input[data-sim]")) event.target.blur(); };
    slide.addEventListener("pointerover", onOver);
    slide.addEventListener("pointermove", onMove);
    slide.addEventListener("pointerleave", onLeave);
    slide.addEventListener("click", onClick, true);
    slide.addEventListener("input", onInput);
    slide.addEventListener("change", onChange);
    return {
      openDetail,
      closeDetail,
      get detailOpen() { return Boolean(slide.dataset.detailOpen) || Boolean(zoomed) || Boolean(spotted); },
      destroy() {
        closeDetail(); hideTip(); setHot(null); setHotObj(null);
        slide.removeEventListener("pointerover", onOver);
        slide.removeEventListener("pointermove", onMove);
        slide.removeEventListener("pointerleave", onLeave);
        slide.removeEventListener("click", onClick, true);
        slide.removeEventListener("input", onInput);
        slide.removeEventListener("change", onChange);
      },
    };
  }

  // ---------------------------------------------------------------- engine stylesheet (for the presenter view window)

  function engineCss() {
    const out = [];
    for (const sheet of document.styleSheets) {
      const node = sheet.ownerNode;
      if (!node?.hasAttribute?.("data-hs-engine")) continue;
      try { for (const rule of sheet.cssRules) out.push(rule.cssText); } catch { /* cross-origin sheet */ }
    }
    return out.join("\n");
  }

  // ---------------------------------------------------------------- player

  const TRANSITIONS = new Set(Object.keys(E.TRANSITIONS));
  const TRANSITION_MS = { wipe: 920, circle: 920, drill: 560, push: 680, flip: 940, dive: 860, blinds: 920, curtain: 920, morph: 900 };
  // These clip the incoming slide itself, so edges and click points share the slide's own coordinates.
  const CLIPPED = new Set(["wipe", "circle", "blinds", "curtain"]);

  // ---------------------------------------------------------------- transition options

  // 画面切り替え → 効果のオプション. Directions are where the new slide comes from; going back reverses them.
  const MOVE = { right: [1, 0], left: [-1, 0], down: [0, 1], up: [0, -1] };
  const OPPOSITE = { right: "left", left: "right", down: "up", up: "down" };
  const CLIP_FROM = { right: "inset(0 0 0 100%)", left: "inset(0 100% 0 0)", down: "inset(100% 0 0 0)", up: "inset(0 0 100% 0)" };
  /**
   * What a transition's option adds to the elements coming in and going out (classes and variables). Returns
   * { rev, band }: whether the reversed animation plays, and whether the wipe's colour band shows.
   */
  function transitionOption(type, option, back, entering, leaving) {
    const list = E.TRANSITION_OPTIONS?.[type];
    const valid = list && list.some(([key]) => key === option);
    if (!valid || option === list[0][0]) return { rev: back, band: true };
    if (type === "slide" || type === "push" || type === "wipe") {
      const way = back ? OPPOSITE[option] : option;
      const first = list[0][0];
      if (way === first) return { rev: false, band: true };
      if (way === OPPOSITE[first]) return { rev: true, band: true };
      const [x, y] = MOVE[way];
      for (const el of [entering, leaving]) {
        el.classList.add("tr-dir");
        el.style.setProperty("--tr-dx", `${x * 100}%`);
        el.style.setProperty("--tr-dy", `${y * 100}%`);
      }
      entering.style.setProperty("--tr-clip", CLIP_FROM[way]);
      return { rev: false, band: false };
    }
    // 右へ (flip) turns the other way; 縦 (blinds) always opens across; 横 (curtain) opens up and down.
    if (type === "flip") return { rev: !back, band: false };
    if (type === "blinds") return { rev: true, band: false };
    if (type === "curtain") { entering.classList.add("tr-h"); return { rev: back, band: false }; }
    return { rev: back, band: true };
  }

  // ---------------------------------------------------------------- transition sounds

  // 画面切り替え → サウンド: each sound is made with Web Audio (oscillators and filtered noise), so a deck and its
  // exported HTML need no sound files. "stop" stops the sounds still ringing from earlier slides.
  let soundCtx = null;
  const ringing = new Set();
  function audio() {
    const AC = typeof window !== "undefined" && (window.AudioContext || window.webkitAudioContext);
    if (!AC) return null;
    if (!soundCtx || soundCtx.state === "closed") soundCtx = new AC();
    if (soundCtx.state === "suspended") soundCtx.resume().catch(() => {});
    return soundCtx;
  }
  function stopSounds() {
    for (const node of ringing) { try { node.stop(); } catch { /* already stopped */ } }
    ringing.clear();
  }
  /** Play a transition sound (a key of E.TRANSITION_SOUNDS). Returns how long it rings (s), or 0. */
  // `context`: an OfflineAudioContext to draw the sound into instead of playing it (renderSound).
  function playSound(kind, { volume = 0.6, context = null } = {}) {
    if (kind === "stop") { if (!context) stopSounds(); return 0; }
    if (!E.TRANSITION_SOUNDS?.[kind]) return 0;
    const ac = context || audio();
    if (!ac) return 0;
    const t0 = ac.currentTime + 0.02;
    const out = ac.createGain();
    out.gain.value = Math.max(0, Math.min(1, volume));
    out.connect(ac.destination);
    const keep = (node, start, stop) => {
      if (!context) { ringing.add(node); node.onended = () => ringing.delete(node); }
      node.start(start);
      node.stop(stop);
    };
    // A tone: frequency (Hz), wave, start, length, peak level, optional slide to another pitch.
    const tone = (freq, { type = "sine", at = 0, len = 0.6, peak = 0.4, to = null } = {}) => {
      const osc = ac.createOscillator();
      const g = ac.createGain();
      osc.type = type;
      osc.frequency.setValueAtTime(freq, t0 + at);
      if (to) osc.frequency.exponentialRampToValueAtTime(to, t0 + at + len);
      g.gain.setValueAtTime(0.0001, t0 + at);
      g.gain.exponentialRampToValueAtTime(peak, t0 + at + 0.008);
      g.gain.exponentialRampToValueAtTime(0.0001, t0 + at + len);
      osc.connect(g).connect(out);
      keep(osc, t0 + at, t0 + at + len + 0.05);
    };
    let noiseBuf = null;
    // A burst of noise through a filter: start, length, peak, filter type and frequency (or a sweep to `to`).
    const noise = ({ at = 0, len = 0.1, peak = 0.5, filter = "bandpass", freq = 2000, q = 1, to = null, attack = 0.004 } = {}) => {
      if (!noiseBuf) {
        noiseBuf = ac.createBuffer(1, Math.round(ac.sampleRate * 2), ac.sampleRate);
        const d = noiseBuf.getChannelData(0);
        for (let i = 0; i < d.length; i += 1) d[i] = Math.random() * 2 - 1;
      }
      const src = ac.createBufferSource();
      src.buffer = noiseBuf;
      src.loop = true;
      const f = ac.createBiquadFilter();
      f.type = filter;
      f.frequency.setValueAtTime(freq, t0 + at);
      if (to) f.frequency.exponentialRampToValueAtTime(to, t0 + at + len);
      f.Q.value = q;
      const g = ac.createGain();
      g.gain.setValueAtTime(0.0001, t0 + at);
      g.gain.exponentialRampToValueAtTime(peak, t0 + at + Math.max(0.002, attack));
      g.gain.exponentialRampToValueAtTime(0.0001, t0 + at + len);
      src.connect(f).connect(g).connect(out);
      keep(src, t0 + at, t0 + at + len + 0.05);
    };
    const SOUNDS = {
      chime: () => { tone(1046.5, { len: 1.2, peak: 0.32 }); tone(1568, { at: 0.12, len: 1.3, peak: 0.24 }); tone(2093, { at: 0.24, len: 1.4, peak: 0.16 }); return 1.7; },
      click: () => { noise({ len: 0.035, peak: 0.7, filter: "highpass", freq: 2500 }); tone(1800, { len: 0.03, peak: 0.2, type: "square" }); return 0.1; },
      camera: () => { noise({ len: 0.05, peak: 0.8, freq: 3200, q: 2 }); noise({ at: 0.09, len: 0.07, peak: 0.6, freq: 1800, q: 2 }); return 0.2; },
      whoosh: () => { noise({ len: 0.7, peak: 0.45, freq: 280, to: 3600, q: 0.9, attack: 0.3 }); return 0.8; },
      drum: () => {
        for (let k = 0; k < 22; k += 1) noise({ at: k * 0.055, len: 0.06, peak: 0.12 + k * 0.012, filter: "lowpass", freq: 900 });
        noise({ at: 1.25, len: 0.5, peak: 0.8, filter: "lowpass", freq: 600 });
        tone(110, { at: 1.25, len: 0.5, peak: 0.6, to: 55 });
        return 1.8;
      },
      applause: () => {
        for (let k = 0; k < 90; k += 1) noise({ at: Math.random() * 1.9, len: 0.03 + Math.random() * 0.03, peak: 0.1 + Math.random() * 0.25, freq: 1400 + Math.random() * 2200, q: 1.5 });
        noise({ len: 2.2, peak: 0.12, freq: 2000, q: 0.6, attack: 0.25 });
        return 2.3;
      },
      coin: () => { tone(988, { type: "square", len: 0.09, peak: 0.16 }); tone(1319, { type: "square", at: 0.08, len: 0.5, peak: 0.16 }); return 0.65; },
      bell: () => { for (const [m, p] of [[1, 0.35], [2.76, 0.18], [5.4, 0.1], [8.93, 0.05]]) tone(523.25 * m, { len: 2 / Math.sqrt(m), peak: p }); return 2.1; },
    };
    try { return SOUNDS[kind](); } catch { return 0; }
  }

  /**
   * One of the transition sounds as a WAV file (a data: URL, 16-bit mono), drawn offline with the same oscillators and
   * noise the show plays — for a PowerPoint, which plays sound files (tools/pptx_timing.py).
   */
  async function renderSound(kind, { volume = 0.6, rate = 44100 } = {}) {
    const OAC = typeof window !== "undefined" && (window.OfflineAudioContext || window.webkitOfflineAudioContext);
    if (!OAC || !E.TRANSITION_SOUNDS?.[kind] || kind === "stop") return null;
    const ac = new OAC(1, Math.round(rate * 3), rate);
    const len = playSound(kind, { volume, context: ac });
    if (!len) return null;
    const buffer = await ac.startRendering();
    const data = buffer.getChannelData(0).subarray(0, Math.min(buffer.length, Math.round(rate * (len + 0.12))));
    const bytes = new Uint8Array(44 + data.length * 2);
    const view = new DataView(bytes.buffer);
    const text = (at, s) => { for (let i = 0; i < s.length; i += 1) view.setUint8(at + i, s.charCodeAt(i)); };
    text(0, "RIFF"); view.setUint32(4, 36 + data.length * 2, true); text(8, "WAVE"); text(12, "fmt ");
    view.setUint32(16, 16, true); view.setUint16(20, 1, true); view.setUint16(22, 1, true); view.setUint32(24, rate, true);
    view.setUint32(28, rate * 2, true); view.setUint16(32, 2, true); view.setUint16(34, 16, true); text(36, "data"); view.setUint32(40, data.length * 2, true);
    for (let i = 0; i < data.length; i += 1) view.setInt16(44 + i * 2, Math.max(-1, Math.min(1, data[i])) * 0x7fff, true);
    let binary = "";
    for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
    return `data:audio/wav;base64,${btoa(binary)}`;
  }

  /**
   * A presentation in `host` (the studio's presenter overlay or an exported file's body).
   * opts: { deck, start, step, fitFor(i), renderOptions, onChange({index, step}), onClose(), closable, keyboard, static }
   * The story is the slides without `drillOf`; a deep-dive page opens from its item and returns to where it left.
   * `static` shows every slide finished (no motion, every step shown), like the "reduce motion" setting.
   */
  function createPlayer(host, opts) {
    const deck = opts.deck;
    const slides = deck.slides || [];
    const total = slides.length;
    const story = E.storyMap(slides);
    // スライド ショーの設定: a custom show (in its own order) or a range plays only those slides.
    const shown = story.order.length ? story.order : slides.map((_, i) => i);
    const only = Array.isArray(opts.only) && opts.only.length ? new Set(opts.only.filter((i) => shown.includes(i))) : null;
    const order = only?.size ? [...only] : [...shown];
    if (!order.length) order.push(0);
    const kiosk = Boolean(opts.kiosk);
    // ビデオの作成 plays the show once to its end (stopAtEnd), even as a kiosk.
    const looping = !opts.stopAtEnd && Boolean(opts.loop || kiosk);
    const place = (i) => Math.max(0, order.indexOf(story.parent[i] ?? i));
    const transition = TRANSITIONS.has(deck.transition) ? deck.transition : "fade";
    const doc = host.ownerDocument;
    const win = doc.defaultView;
    let index = Math.max(0, Math.min(total - 1, opts.start || 0));
    if (only?.size && !only.has(story.parent[index] ?? index)) index = order[0];
    let step = 0;
    let current = null;
    let interaction = null;
    let busy = null;
    let pv = null;
    let gesture = false;
    let origin = null;
    // While a deep-dive page is shown: the slide (and build step) to go back to, and where it was opened from.
    let back = null;
    let hinted = false;
    let demoRun = null;
    let autoTimer = null;
    const still = Boolean(opts.static);
    // オンライン プレゼンテーション: a viewer only watches; the presenter's moves come in through follow().
    const viewer = Boolean(opts.viewer);
    const started = Date.now();

    const stage = h("div", { class: "hs-player-stage" });
    const progress = h("div", { class: "hs-player-progress" }, h("i"));
    const counter = h("span", { class: "hs-player-count" });
    const notes = h("div", { class: "hs-player-notes", hidden: true });
    const btn = (label, title, fn, cls = "") => h("button", { class: `hs-player-btn ${cls}`, type: "button", title, "aria-label": title, onclick: (event) => { event.stopPropagation(); fn(); } }, label);
    // PowerPoint Live: a viewer can send reactions (they float up on every screen).
    const REACTIONS = ["👍", "❤️", "👏", "😮", "💡", "😂"];
    const reactBtns = viewer && opts.onReact ? h("span", { class: "hs-react-bar" }, REACTIONS.map((emoji) => h("button", { type: "button", class: "hs-react-btn", title: "リアクションを送る", "aria-label": `リアクション ${emoji}`, onclick: (event) => { event.stopPropagation(); react(emoji); opts.onReact(emoji); } }, emoji))) : null;
    const bar = viewer ? h("div", { class: "hs-player-bar" }, counter, h("span", { class: "hs-player-spacer" }),
      h("span", { class: "hs-player-live" }, opts.viewerLabel || "発表者に合わせて表示しています"), reactBtns,
      btn("全画面", "全画面（F）", () => toggleFullscreen())) : h("div", { class: "hs-player-bar" },
      btn("‹", "前へ（←）", () => prev()),
      counter,
      btn("›", "次へ（→・クリック）", () => next()),
      h("span", { class: "hs-player-spacer" }),
      btn("一覧", "スライド一覧（O・G）", () => toggleGrid()),
      btn("自動デモ", "自動デモ：矢印が各ページを操作して見せます（D）", () => demo(), "demo"),
      btn("ペン", "ペンとレーザー ポインター（Ctrl+P ペン・Ctrl+I 蛍光ペン・Ctrl+L レーザー・Ctrl+E 消しゴム・E すべて消去・Ctrl+A 矢印）", () => togglePenMenu(), "pen"),
      btn("拡大", "スライドを拡大（＋ / −）：クリックした所を大きく表示し、ドラッグで見回す。Esc・右クリックで全体に戻る", () => toggleMagnifier(), "zoom"),
      btn("字幕", "話した言葉を字幕で表示（J）：マイクを使います（Chrome・Edge）", () => toggleCaptions(), "captions"),
      btn("ノート", "ノートを表示（N）", () => toggleNotes()),
      btn("発表者ビュー", "別ウィンドウにノート・次のスライド・経過時間（P）", () => openPresenterView()),
      btn("全画面", "全画面（F）", () => toggleFullscreen()),
      opts.joinCard ? btn("参加", "参加用の QR コードとリンクを表示（Q）", () => toggleJoin(), "join") : null,
      opts.closable === false ? null : btn("終了", "発表を終了（Esc）", () => close(), "end"));
    const grid = h("div", { class: "hs-player-grid", hidden: true });
    const black = h("div", { class: "hs-player-black", hidden: true, onclick: (event) => { event.stopPropagation(); black.hidden = true; } });
    const backBtn = h("button", { class: "hs-player-back", type: "button", hidden: true, title: "元のスライドへ戻る（Esc・←）", onclick: (event) => { event.stopPropagation(); closeDrill(); } }, "← 元のスライドへ");
    // Sounds set to play on across slides (再生 →「スライド切り替え後も再生」) move here when their slide is left.
    // Writing on the slides while presenting (ペン・蛍光ペン・レーザー ポインター), kept per slide.
    const inkSvg = s("svg", { class: "hs-show-ink", viewBox: "0 0 1920 1080", preserveAspectRatio: "none", "aria-hidden": "true" });
    const laser = h("div", { class: "hs-laser", hidden: true });
    const penMenu = h("div", { class: "hs-pen-menu", hidden: true });
    const annotations = new Map();
    let penMode = null;
    let penColor = /^#[0-9a-f]{6}$/i.test(opts.penColor || "") ? opts.penColor : "#ff0000";
    // 字幕 (live subtitles): what the speaker says, written under the slide (the browser's speech recognition).
    const captionBox = h("div", { class: "hs-captions", hidden: true, "aria-live": "polite" });
    let captions = null;
    const carryHost = h("div", { class: "hs-carry", hidden: true });
    const carried = [];
    // オンライン プレゼンテーション: the card with the QR code and the link to join (Q).
    const joinCard = h("div", { class: "hs-join", hidden: true });
    if (opts.joinCard) {
      const qr = h("div", { class: "hs-join-qr" });
      if (opts.joinCard.svg) qr.innerHTML = opts.joinCard.svg;
      joinCard.append(h("b", {}, "スマートフォン・PCで参加"), qr, h("span", { class: "hs-join-url" }, opts.joinCard.url || ""));
    }
    const reactLayer = h("div", { class: "hs-react-layer", "aria-hidden": "true" });
    const player = h("div", { class: "hs-player", tabindex: "-1" }, stage, progress, notes, backBtn, bar, grid, black, carryHost, inkSvg, laser, penMenu, captionBox, joinCard, reactLayer);
    function toggleJoin(force = null) { if (opts.joinCard) joinCard.hidden = force == null ? !joinCard.hidden : !force; }

    /** A reaction floats up from the bottom of the screen. */
    function react(emoji) {
      if (reactLayer.childElementCount > 40) reactLayer.firstElementChild?.remove();
      const el = h("span", { class: "hs-react", style: { left: `${8 + Math.random() * 84}%`, "--drift": `${Math.round((Math.random() - 0.5) * 80)}px` } }, emoji);
      reactLayer.append(el);
      setTimeout(() => el.remove(), 3200);
    }
    /** The presenter's captions on a viewer's screen (they fade after a pause). */
    let captionTimer = 0;
    function caption(text) {
      captionBox.textContent = String(text || "").slice(-120);
      captionBox.hidden = !text;
      player.classList.toggle("with-captions", Boolean(text));
      clearTimeout(captionTimer);
      captionTimer = setTimeout(() => { captionBox.hidden = true; player.classList.remove("with-captions"); }, 6000);
    }
    // アンケート: the answers so far (poll id → counts) and this viewer's own (poll id → choice).
    const votes = new Map();
    const mine = new Map();
    function paintPolls(scope = current?.firstElementChild) {
      for (const box of scope?.querySelectorAll?.(".hs-poll[data-poll]") || []) {
        const items = [...box.querySelectorAll(".hs-poll-opt")];
        const counts = votes.get(box.dataset.poll) || items.map(() => 0);
        const total = counts.reduce((a, b) => a + b, 0);
        const top = Math.max(...counts);
        items.forEach((li, k) => {
          const n = counts[k] || 0;
          li.querySelector(".hs-poll-bar i").style.width = `${total ? Math.round((n / total) * 100) : 0}%`;
          li.querySelector(".hs-poll-count").textContent = String(n);
          li.classList.toggle("is-top", total > 0 && n === top);
          li.classList.toggle("is-mine", mine.get(box.dataset.poll) === k);
        });
        const foot = box.querySelector(".hs-poll-total");
        if (foot) foot.textContent = total ? `回答 ${total}人` : viewer ? "選んで回答してください" : "回答をクリック・タップ";
      }
    }
    function setVotes(poll, counts) { votes.set(poll, Array.isArray(counts) ? counts.map((n) => Math.max(0, Number(n) || 0)) : []); paintPolls(); }
    /** A click on a choice: an answer (sent to the presenter online, or counted here). */
    function answer(choice) {
      const poll = choice.closest(".hs-poll")?.dataset.poll;
      const k = Number(choice.dataset.option);
      if (!poll || !Number.isInteger(k)) return;
      if (opts.onVote) { if (opts.onVote(poll, k) !== false) { mine.set(poll, k); paintPolls(); } return; }
      const counts = [...(votes.get(poll) || [...choice.parentElement.children].map(() => 0))];
      counts[k] = (counts[k] || 0) + 1;
      setVotes(poll, counts);
    }
    host.append(player);
    if (kiosk) player.classList.add("kiosk");

    /** The ink layer lies exactly over the slide shown (its own 1920 × 1080 coordinates). */
    function placeInk() {
      const slideEl = current?.firstElementChild;
      if (!slideEl) return;
      const r = slideEl.getBoundingClientRect();
      const base = player.getBoundingClientRect();
      Object.assign(inkSvg.style, { left: `${r.left - base.left}px`, top: `${r.top - base.top}px`, width: `${r.width}px`, height: `${r.height}px` });
    }
    const inkPathD = (pts) => (pts.length === 1 ? `M${pts[0][0]} ${pts[0][1]} l0.01 0` : `M${pts.map(([x, y]) => `${Math.round(x * 10) / 10} ${Math.round(y * 10) / 10}`).join(" L")}`);
    function drawInk() {
      const list = annotations.get(index) || [];
      inkSvg.replaceChildren(...list.map((st) => s("path", { d: inkPathD(st.pts), fill: "none", stroke: st.color, "stroke-width": st.width, "stroke-linecap": st.highlighter ? "square" : "round", "stroke-linejoin": "round", "stroke-opacity": st.highlighter ? 0.9 : 1, class: st.highlighter ? "hs-ink-hl" : null })));
      placeInk();
    }
    function setPen(mode) {
      penMode = penMode === mode ? null : mode;
      player.classList.toggle("pen-on", Boolean(penMode));
      player.dataset.pen = penMode || "";
      laser.hidden = penMode !== "laser";
      penMenu.hidden = true;
      if (penMode) flash({ pen: "ペン：ドラッグで書けます（Escで終わる）", highlighter: "蛍光ペン", laser: "レーザー ポインター", eraser: "消しゴム：消したい線をなぞる" }[penMode]);
    }
    function eraseAll() { annotations.delete(index); drawInk(); }
    // カメオ: one camera for the whole show, shown in every camera object; stopped when the show ends.
    let camera = null;
    async function startCameras(slideEl) {
      const feeds = [...slideEl.querySelectorAll(".hs-camera-feed")];
      if (!feeds.length) return;
      try {
        camera ||= navigator.mediaDevices.getUserMedia({ video: { width: { ideal: 1280 }, height: { ideal: 720 } }, audio: false });
        const stream = await camera;
        for (const video of feeds) { video.srcObject = stream; video.play().catch(() => {}); }
      } catch {
        camera = null;
        flash("カメラを使えません（ブラウザでカメラを許可してください）");
      }
    }
    function toggleCaptions() {
      if (captions) { const rec = captions; captions = null; try { rec.stop(); } catch { /* stopped */ } captionBox.hidden = true; player.classList.remove("with-captions"); flash("字幕をオフにしました"); return; }
      const Recognition = win.SpeechRecognition || win.webkitSpeechRecognition;
      if (!Recognition) { flash("このブラウザは字幕に対応していません（Chrome・Edgeで使えます）"); return; }
      const rec = new Recognition();
      rec.lang = opts.captionLang || "ja-JP";
      rec.continuous = true;
      rec.interimResults = true;
      let done = "";
      rec.onresult = (event) => {
        let interim = "";
        for (let i = event.resultIndex; i < event.results.length; i += 1) {
          const text = event.results[i][0].transcript;
          if (event.results[i].isFinal) done = `${done} ${text}`.trim().slice(-160); else interim += text;
        }
        captionBox.textContent = `${done} ${interim}`.trim().slice(-120);
        // オンライン プレゼンテーション: the words go to the viewers too.
        opts.onCaption?.(captionBox.textContent);
      };
      rec.onerror = (event) => { if (event.error === "not-allowed" || event.error === "service-not-allowed") { flash("マイクを使えないため字幕を出せません"); captions = null; captionBox.hidden = true; } };
      // Recognition stops by itself after a pause: it starts again while 字幕 is on.
      rec.onend = () => { if (captions === rec) { try { rec.start(); } catch { /* starting */ } } };
      try { rec.start(); } catch { flash("字幕を始められませんでした"); return; }
      captions = rec;
      captionBox.textContent = "（話すと字幕が出ます）";
      captionBox.hidden = false;
      player.classList.add("with-captions");
    }
    function togglePenMenu() {
      if (!penMenu.hidden) { penMenu.hidden = true; return; }
      const item = (label, mode, keys) => h("button", { type: "button", class: penMode === mode ? "on" : "", onclick: (event) => { event.stopPropagation(); setPen(mode); } }, h("span", {}, label), h("kbd", {}, keys));
      penMenu.replaceChildren(
        item("レーザー ポインター", "laser", "Ctrl+L"), item("ペン", "pen", "Ctrl+P"), item("蛍光ペン", "highlighter", "Ctrl+I"), item("消しゴム", "eraser", "Ctrl+E"),
        h("div", { class: "hs-pen-colors" }, [["#ff0000", "赤"], ["#0070c0", "青"], ["#000000", "黒"]].map(([c, label]) => h("button", { type: "button", title: `インクの色：${label}`, class: penColor === c ? "on" : "", style: { background: c }, onclick: (event) => { event.stopPropagation(); penColor = c; if (penMode !== "pen") setPen("pen"); else penMenu.hidden = true; } }))),
        h("button", { type: "button", onclick: (event) => { event.stopPropagation(); eraseAll(); penMenu.hidden = true; } }, h("span", {}, "スライド上のインクをすべて消去"), h("kbd", {}, "E")),
        h("button", { type: "button", onclick: (event) => { event.stopPropagation(); if (penMode) setPen(penMode); penMenu.hidden = true; } }, h("span", {}, "矢印（ペンを使わない）"), h("kbd", {}, "Ctrl+A")));
      penMenu.hidden = false;
    }
    const toSlide = (event) => { const r = inkSvg.getBoundingClientRect(); return [((event.clientX - r.left) / r.width) * 1920, ((event.clientY - r.top) / r.height) * 1080]; };
    inkSvg.addEventListener("pointerdown", (event) => {
      if (!penMode || penMode === "laser") return;
      event.preventDefault();
      event.stopPropagation();
      inkSvg.setPointerCapture(event.pointerId);
      const pts = [toSlide(event)];
      // The highlighter is yellow, laid over the slide like a marker (multiplied, the words stay black).
      const stroke = penMode === "eraser" ? null : { pts, color: penMode === "highlighter" ? "#ffff00" : penColor, width: penMode === "highlighter" ? 26 : 5, highlighter: penMode === "highlighter" };
      if (stroke) { if (!annotations.has(index)) annotations.set(index, []); annotations.get(index).push(stroke); }
      const move = (ev) => {
        const p = toSlide(ev);
        if (stroke) { pts.push(p); drawInk(); return; }
        const list = annotations.get(index) || [];
        const kept = list.filter((st) => !st.pts.some((q) => Math.hypot(q[0] - p[0], q[1] - p[1]) < 22 + st.width / 2));
        if (kept.length !== list.length) { annotations.set(index, kept); drawInk(); }
      };
      const up = () => { inkSvg.removeEventListener("pointermove", move); inkSvg.removeEventListener("pointerup", up); };
      inkSvg.addEventListener("pointermove", move);
      inkSvg.addEventListener("pointerup", up);
      drawInk();
    });
    inkSvg.addEventListener("click", (event) => { if (penMode) event.stopPropagation(); });
    // The slide is rescaled when the window changes: the ink layer follows it.
    const inkFollow = typeof ResizeObserver === "function" ? new ResizeObserver(() => requestAnimationFrame(placeInk)) : null;
    inkFollow?.observe(stage);
    player.addEventListener("pointermove", (event) => {
      if (penMode !== "laser") return;
      const base = player.getBoundingClientRect();
      laser.style.transform = `translate(${event.clientX - base.left}px, ${event.clientY - base.top}px)`;
    });

    /** Leaving a slide: its playing sounds that go on across slides keep playing (moved out of the slide). */
    function carrySounds(slideEl, from) {
      for (const audio of slideEl.querySelectorAll("audio[data-across]")) {
        if (audio.paused) continue;
        const n = Number(audio.dataset.across) || 0;
        const start = place(from);
        carried.push({ audio, id: audio.closest("[data-el]")?.dataset.el, start, until: n >= 999 ? Infinity : start + n - 1 });
        carryHost.append(audio); // moved in the same task, so it does not pause
      }
    }
    /** A carried sound stops once the talk goes beyond its slides (or back before the one it started on). */
    function pruneSounds(i) {
      const pos = place(i);
      for (let k = carried.length - 1; k >= 0; k -= 1) {
        const c = carried[k];
        if (pos > c.until || pos < c.start || c.audio.paused) {
          E.mediaPause?.(c.audio, { stop: true });
          c.audio.remove();
          carried.splice(k, 1);
        }
      }
      return new Set(carried.map((c) => c.id).filter(Boolean));
    }

    const renderAt = (i) => {
      const el = E.render(slides[i], { ...(opts.renderOptions || {}), deck, index: i, mode: "present", fit: opts.fitFor?.(i) });
      // アニメーションのサウンド: an effect's sound rings as it starts (unless the show plays without sounds).
      if (opts.sounds !== false && !opts.static) el.hsAnimSound = (kind) => playSound(kind);
      // メディア コントロールの表示 turned off: no video controls and no sound bar (a click on either still plays it).
      if (opts.mediaControls === false) { el.classList.add("hs-no-controls"); for (const v of el.querySelectorAll("video[controls]")) v.controls = false; }
      return E.mount(el, { contain: true, className: "hs-player-slide" });
    };

    /** ノート on the notes bar and in the presenter view: the formatted words (bold, lists, links) or the plain lines. */
    function showNotes(box, slide, empty) {
      const note = E.noteOf ? E.noteOf(slide) : null;
      if (note?.rich) box.innerHTML = note.rich;
      else box.textContent = E.strip(slide?.notes || "") || empty;
    }

    function update() {
      const pos = place(index);
      counter.textContent = `${pos + 1} / ${order.length}${back ? " ・ 深掘り" : ""}`;
      progress.firstChild.style.width = `${((pos + 1) / order.length) * 100}%`;
      backBtn.hidden = !back;
      player.classList.toggle("in-drill", Boolean(back));
      showNotes(notes, slides[index], "（ノートはありません）");
      opts.onChange?.({ index, step });
      syncPresenterView();
    }

    function nameShared(scaler, on) {
      const slide = scaler?.firstElementChild;
      if (!slide) return;
      const pick = (sel, name) => { const el = slide.querySelector(sel); if (el) el.style.viewTransitionName = on ? name : ""; };
      pick(".hs-title, .hs-cover-title, .hs-section-title, .hs-hero-title", "hs-title");
      pick(".hs-eyebrow", "hs-eyebrow");
      pick(".hs-takeaway", "hs-takeaway");
      pick(".hs-foot", "hs-foot");
    }

    /**
     * 「自動的に切り替え」: after the slide's seconds, the clicks it still has play one after another (each
     * waiting for its animations), then the talk moves on.
     */
    function scheduleAdvance(slideEl, i) {
      clearTimeout(autoTimer);
      if (opts.useTimings === false || viewer) return;
      // 自動プレゼンテーション (kiosk): a slide without a time of its own stays for a while, then the show moves on.
      // (ビデオの作成 may ignore the slides' own times and give each the same seconds.)
      const own = opts.ignoreTimings ? null : slides[i]?.advance;
      const secs = Number(own ?? (kiosk ? opts.kioskSeconds || 8 : NaN));
      // The finished view (#static) stays put; 「アニメーションを表示しない」 still keeps the slides' timings.
      if ((own == null && !kiosk) || !Number.isFinite(secs) || secs < 0 || (still && !opts.noAnimation)) return;
      const tick = () => {
        if (current?.firstElementChild !== slideEl || demoRun) return;
        if (E.animBusy?.(slideEl)) { autoTimer = setTimeout(tick, 200); return; }
        const more = step < stepsOf(slideEl);
        next();
        if (more) autoTimer = setTimeout(tick, 400);
      };
      autoTimer = setTimeout(tick, secs * 1000);
    }

    async function show(i, { dir = 1, fullStep = false, atStep = null, via = null, at = null } = {}) {
      if (busy) await busy;
      clearTimeout(autoTimer);
      resetZoom();
      const prevScaler = current;
      const prevSlide = prevScaler?.firstElementChild;
      if (prevSlide) { interaction?.destroy(); carrySounds(prevSlide, index); stopMedia(prevSlide); }
      const from = index;
      index = i;
      const playing = pruneSounds(i);
      const next = renderAt(i);
      const slide = next.firstElementChild;
      step = Math.min(still ? Infinity : atStep ?? (fullStep ? Infinity : 0), stepsOf(slide));
      // A slide may have its own way in; going back plays the way in of the slide being left, in reverse.
      const own = slides[dir < 0 ? from : i]?.transition;
      const type = still || reduced() || !prevScaler ? "none" : via || (TRANSITIONS.has(own) ? own : transition);
      // A slide may set how long its transition takes (画面切り替えの「期間」).
      const ownDur = Number(slides[dir < 0 ? from : i]?.transitionDur);
      const trMs = !via && ownDur >= 100 && ownDur <= 10000 ? ownDur : null;
      // A deep-dive page grows out of the item that opened it, and shrinks back into it.
      if (type === "drill" && at) for (const el of [next, prevScaler]) { el.style.setProperty("--ox", `${at.x}px`); el.style.setProperty("--oy", `${at.y}px`); }
      const enter = () => {
        // Coming back from a deep-dive page, the slide is shown as it was left, without its entrance again.
        // Animations that start with the slide wait for its way in (as in PowerPoint).
        const wayIn = type === "none" || type === "morph" ? 0 : Math.round((trMs ?? TRANSITION_MS[type] ?? 620) * 0.85);
        play(slide, { step, animate: atStep == null && !still, delay: wayIn });
        paintPolls(slide);
        // 画面切り替えのサウンド rings as the slide arrives going forward (not going back or out of a deep-dive page).
        if (!still && opts.sounds !== false && slides[i]?.transitionSound && atStep == null && dir >= 0) playSound(slides[i].transitionSound);
        interaction = activate(slide, { details: [...(slides[i]?.details || []), ...(E.objectDetails?.(slides[i]) || [])], elements: slides[i]?.elements || [], onDrill: back || viewer ? null : (to, el) => openDrill(to, el) });
        playMedia(slide, { sound: gesture, skip: playing, noNarration: opts.narration === false });
        scheduleAdvance(slide, i);
        drawInk();
        startCameras(slide);
        if (still) return;
        // The first slide with clickable items says how to use them (once per presentation), and every page
        // rings what can be clicked once, right after it has arrived.
        const kinds = [slide.querySelector(".hs-detail-badge") ? "「＋ 詳しく」" : "", slide.querySelector(".hs-drill-badge") ? "「↗ 深掘り」" : "", slide.querySelector(".hs-control") ? "切り替え・スライダー" : ""].filter(Boolean).join("・");
        if (kinds && !hinted) { hinted = true; setTimeout(() => flash(`${kinds}の付いた項目はクリックできます`), 900); }
        else if (!hinted && slide.querySelector(".hs-obj[data-action]:not([data-action=\"url\"])")) { hinted = true; setTimeout(() => flash("光った部品はクリックできます"), 900); }
        setTimeout(() => ring(slide), atStep == null ? 1900 : 300);
      };
      // 変形: the parts the two slides share travel from one slide to the other (the titles of layout slides use the browser's own).
      const morphed = type === "morph" ? E.morphPairs?.(slides[from]?.elements, slides[i]?.elements) || [] : [];
      if (type === "morph" && morphed.length) {
        stage.append(next);
        E.scale(next);
        current = next;
        const finished = morphIn(prevSlide, slide, morphed, { dur: trMs || 900 });
        prevScaler.remove();
        enter();
        busy = finished.then(() => { busy = null; });
      } else if (type === "morph" && doc.startViewTransition) {
        nameShared(prevScaler, true);
        const vt = doc.startViewTransition(() => {
          nameShared(prevScaler, false);
          prevScaler.remove();
          stage.append(next);
          E.scale(next);
          nameShared(next, true);
        });
        current = next;
        busy = vt.finished.catch(() => {}).then(() => { nameShared(next, false); busy = null; });
        vt.updateCallbackDone.then(enter, enter);
      } else if (type === "none" || type === "morph") {
        prevScaler?.remove();
        stage.append(next);
        current = next;
        E.scale(next);
        enter();
        if (type === "morph" && prevScaler) next.classList.add("hs-tr-in-fade");
      } else {
        stage.append(next);
        E.scale(next);
        current = next;
        // Wipe and circle clip the incoming slide itself, so the edge, its colour band and the click point
        // share the slide's own coordinates (letterboxing never shows them).
        const entering = CLIPPED.has(type) ? slide : next;
        const option = transitionOption(type, via ? null : slides[dir < 0 ? from : i]?.transitionDir, dir < 0, entering, prevScaler);
        const suffix = option.rev ? " rev" : "";
        let band = null;
        if (type === "circle") {
          const box = slide.getBoundingClientRect();
          const at = origin && box.width ? { x: ((origin.x - box.left) / box.width) * 100, y: ((origin.y - box.top) / box.height) * 100 } : { x: 50, y: 50 };
          slide.style.setProperty("--cx", `${Math.round(at.x)}%`);
          slide.style.setProperty("--cy", `${Math.round(at.y)}%`);
        }
        if (type === "wipe" && option.band) {
          const tone = win.getComputedStyle(slide);
          band = h("div", { class: `hs-tr-band${suffix}` });
          band.style.setProperty("--band", tone.getPropertyValue("--accent").trim() || "#2451e6");
          band.style.setProperty("--band2", tone.getPropertyValue("--accent2").trim() || "#13a89e");
          slide.append(band);
        }
        origin = null;
        entering.className += ` hs-tr-in-${type}${suffix}`;
        prevScaler.className += ` hs-tr-out-${type}${suffix}`;
        if (trMs) for (const el of [entering, prevScaler, band].filter(Boolean)) el.style.animationDuration = `${trMs}ms`;
        enter();
        busy = new Promise((resolve) => setTimeout(resolve, trMs ? trMs + 40 : TRANSITION_MS[type] ?? 620)).then(() => {
          prevScaler.remove();
          band?.remove();
          entering.classList.remove(`hs-tr-in-${type}`, "rev", "tr-dir", "tr-h");
          busy = null;
        });
      }
      update();
    }

    function next() {
      if (!black.hidden) { black.hidden = true; return; }
      if (interaction?.detailOpen) { interaction.closeDetail(); return; }
      const slide = current?.firstElementChild;
      // A click while animations still play finishes them first (as PowerPoint does).
      if (slide && E.animBusy?.(slide)) { E.animFinish(slide); return; }
      if (slide && step < stepsOf(slide)) {
        step += 1;
        reveal(slide, step);
        update();
        return;
      }
      // At the end of a deep-dive page (or a zoomed-into slide), the talk goes back to the slide it came from.
      if (zoomReturn()) return;
      if (back) { closeDrill(); return; }
      const pos = place(index);
      if (pos < order.length - 1) show(order[pos + 1], { dir: 1 });
      // 「Esc キーが押されるまで繰り返す」: after the last slide, the first again.
      else if (looping) show(order[0], { dir: 1 });
    }

    function prev() {
      if (interaction?.detailOpen) { interaction.closeDetail(); return; }
      if (zoomReturn()) return;
      if (back) { closeDrill(); return; }
      const pos = place(index);
      if (pos > 0) show(order[pos - 1], { dir: -1, fullStep: true });
    }

    /** スライド ズーム: the slide grows out of its picture; at its end (ズームに戻る) the talk shrinks back into it. */
    let zoomBack = null;
    function zoomTo(o, el) {
      const to = slides.findIndex((slide) => slide.sid === o.target);
      if (to < 0) { flash("ズーム先のスライドが見つかりません"); return; }
      const base = current?.getBoundingClientRect();
      const r = el.getBoundingClientRect();
      const at = base ? { x: Math.round(r.left + r.width / 2 - base.left), y: Math.round(r.top + r.height / 2 - base.top) } : null;
      zoomBack = o.back !== false ? { index, step, at } : null;
      show(to, { dir: 1, via: "drill", at });
    }
    function zoomReturn() {
      if (!zoomBack) return false;
      const z = zoomBack;
      zoomBack = null;
      show(z.index, { dir: -1, atStep: z.step, via: "drill", at: z.at });
      return true;
    }

    /** Open a deep-dive page (one level: not from another deep-dive page). */
    function openDrill(to, el = null) {
      if (back || !slides[to] || story.parent[to] !== index) return;
      let at = null;
      if (el && current) {
        const base = current.getBoundingClientRect();
        const r = el.getBoundingClientRect();
        at = { x: Math.round(r.left + r.width / 2 - base.left), y: Math.round(r.top + r.height / 2 - base.top) };
      }
      back = { index, step, at };
      show(to, { dir: 1, via: "drill", at });
    }

    function closeDrill() {
      if (!back) return false;
      const { index: to, step: at, at: point } = back;
      back = null;
      show(to, { dir: -1, atStep: at, via: "drill", at: point });
      return true;
    }

    /** Go to a slide by its place in the deck; a deep-dive page opens over its slide. */
    function go(i) {
      zoomBack = null;
      const target = Math.max(0, Math.min(total - 1, i));
      if (target === index) return;
      const parent = story.parent[target];
      back = parent != null ? { index: parent, step: Infinity, at: null } : null;
      show(target, { dir: place(target) >= place(index) ? 1 : -1 });
    }

    /**
     * オンライン プレゼンテーション: a viewer's show follows the presenter to slide `i`, build step `s`. The next step
     * on the same slide plays as a click would; anything else shows the slide as it is at that step.
     */
    async function follow(i, s = 0) {
      if (!slides[i]) return;
      if (busy) await busy;
      zoomBack = null;
      const want = Math.max(0, Number(s) || 0);
      if (i === index) {
        const slide = current?.firstElementChild;
        if (!slide || want === step) return;
        if (want === step + 1 && want <= stepsOf(slide)) { if (E.animBusy?.(slide)) E.animFinish(slide); step = want; reveal(slide, step); update(); return; }
        show(i, { dir: 1, atStep: want });
        return;
      }
      const parent = story.parent[i];
      back = parent != null ? { index: parent, step: Infinity, at: null } : null;
      show(i, { dir: place(i) >= place(index) ? 1 : -1, atStep: want ? want : null });
    }

    function toggleNotes() { notes.hidden = !notes.hidden; player.classList.toggle("with-notes", !notes.hidden); requestAnimationFrame(() => current && E.scale(current)); }

    function toggleGrid() {
      if (!grid.hidden) { grid.hidden = true; grid.replaceChildren(); return; }
      grid.replaceChildren(...order.map((i, pos) => {
        const slide = slides[i];
        const el = E.render(slide, { ...(opts.renderOptions || {}), deck, index: i, mode: "thumb", fit: opts.fitFor?.(i) });
        return h("button", { class: `hs-grid-cell${pos === place(index) ? " is-current" : ""}`, type: "button", onclick: (event) => { event.stopPropagation(); grid.hidden = true; grid.replaceChildren(); go(i); } },
          E.mount(el), h("span", {}, `${pos + 1}. ${E.strip(slide.title || slide.message || "")}`));
      }));
      grid.hidden = false;
    }

    async function toggleFullscreen() {
      try {
        if (doc.fullscreenElement) await doc.exitFullscreen();
        else await (opts.fullscreenTarget || player).requestFullscreen();
      } catch { /* not allowed here */ }
    }

    function close() {
      // Ink written during the show goes back to the editor (which asks whether to keep it, as PowerPoint does).
      const ink = [...annotations.entries()].filter(([, list]) => list.length).map(([i, list]) => ({ index: i, strokes: list }));
      destroy();
      opts.onClose?.({ index, ink });
    }

    // ---- presenter view: notes, the next slide and a clock in a second window (keep the audience window clean)
    function openPresenterView() {
      if (pv && !pv.win.closed) { pv.win.focus(); return; }
      const w = win.open("", "hs-presenter-view", "width=1180,height=760");
      if (!w) { flash("ポップアップがブロックされました。ブラウザで許可してください"); return; }
      const fonts = doc.querySelector("link[data-hs-fonts]")?.href || "";
      w.document.open();
      w.document.write(`<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>発表者ビュー</title>${fonts ? `<link rel="stylesheet" href="${fonts}">` : ""}<style>${engineCss()}\n${PV_CSS}</style></head><body><div class="pv"><section class="pv-now"></section><aside class="pv-side"><div class="pv-label">次のスライド</div><div class="pv-next"></div><div class="pv-clock"><b class="pv-timer">00:00</b><span class="pv-count"></span></div><div class="pv-btns"><button type="button" data-a="prev">‹ 前へ</button><button type="button" data-a="next">次へ ›</button><button type="button" data-a="reset">時間をリセット</button></div></aside><section class="pv-notes"></section></div></body></html>`);
      w.document.close();
      pv = { win: w, start: Date.now() };
      w.document.addEventListener("keydown", onKey);
      w.document.addEventListener("click", (event) => {
        const action = event.target.closest?.("button")?.dataset.a;
        if (action === "prev") prev();
        if (action === "next") next();
        if (action === "reset") pv.start = Date.now();
      });
      w.addEventListener("resize", () => layoutPresenterView());
      pv.timer = setInterval(() => {
        if (!pv || pv.win.closed) { clearInterval(pv?.timer); pv = null; return; }
        const sec = Math.floor((Date.now() - pv.start) / 1000);
        const el = pv.win.document.querySelector(".pv-timer");
        if (el) el.textContent = `${String(Math.floor(sec / 60)).padStart(2, "0")}:${String(sec % 60).padStart(2, "0")}`;
      }, 500);
      syncPresenterView();
    }

    function layoutPresenterView() {
      if (!pv || pv.win.closed) return;
      for (const box of pv.win.document.querySelectorAll(".pv-now, .pv-next")) {
        const slide = box.querySelector(".hs-slide");
        if (!slide) continue;
        const k = Math.min(box.clientWidth / E.W, box.clientHeight / E.H);
        slide.style.transform = `scale(${k})`;
        slide.style.left = `${(box.clientWidth - E.W * k) / 2}px`;
        slide.style.top = `${(box.clientHeight - E.H * k) / 2}px`;
      }
    }

    function syncPresenterView() {
      if (!pv || pv.win.closed) return;
      const d = pv.win.document;
      const put = (sel, i, stepShown) => {
        const box = d.querySelector(sel);
        if (!box) return;
        if (i == null || i >= total) { box.replaceChildren(d.createTextNode("（最後のスライドです）")); return; }
        const el = E.render(slides[i], { ...(opts.renderOptions || {}), deck, index: i, mode: "thumb", fit: opts.fitFor?.(i) });
        // Click steps not yet reached are hidden (a gap slide's measures are switched off instead).
        if (stepShown != null && el.dataset.build === "click") {
          if (el.querySelector("[data-measure]")) { for (const m of el.querySelectorAll("[data-measure]")) m.classList.toggle("is-on", Number(m.dataset.measure) < stepShown); E.gapUpdate(el); }
          else for (const g of el.querySelectorAll("[data-g]")) g.style.visibility = Number(g.dataset.g) < stepShown ? "" : "hidden";
        }
        el.style.position = "absolute";
        el.style.transformOrigin = "0 0";
        box.replaceChildren(d.adoptNode(el));
      };
      put(".pv-now", index, step);
      put(".pv-next", back ? back.index : order[place(index) + 1], back ? Infinity : null);
      d.querySelector(".pv-label").textContent = back ? "戻る先のスライド" : "次のスライド";
      showNotes(d.querySelector(".pv-notes"), slides[index], "（このスライドにノートはありません）");
      d.querySelector(".pv-count").textContent = `${place(index) + 1} / ${order.length}${back ? " ・ 深掘り" : ""}${stepsOf(current?.firstElementChild || d.body) ? `　（${step}/${stepsOf(current.firstElementChild)}）` : ""}`;
      layoutPresenterView();
    }

    /** One ring around each thing on the slide that can be clicked or moved (once, when the page has arrived). */
    function ring(slide) {
      if (!slide.isConnected || current?.firstElementChild !== slide || reduced()) return;
      const overlay = slide.querySelector(".hs-overlay");
      const targets = [...slide.querySelectorAll('[data-detail], [data-drill], .hs-rank-views, .hs-sim-input input, .hs-gap-measure, .hs-obj:is([data-action="zoom"], [data-action="flip"], [data-action="reveal"], [data-action="spot"])')]
        .filter((el) => !(el instanceof SVGElement) && !el.closest(".hs-hidden, .hs-ix-wait, .hs-anim-hide") && !el.parentElement?.closest("[data-detail], [data-drill]")).slice(0, 10);
      for (const el of targets) {
        const box = rectIn(slide, el);
        if (!box.w || !box.h) continue;
        const mark = h("span", { class: "hs-ring", "aria-hidden": "true", style: { left: `${box.x - 10}px`, top: `${box.y - 10}px`, width: `${box.w + 20}px`, height: `${box.h + 20}px` } });
        overlay.append(mark);
        setTimeout(() => mark.remove(), 1600);
      }
    }

    // ---- automatic demo (D): an arrow walks through every page and does what the page invites, then ends on the overview
    function stopDemo(message = "自動デモを止めました") {
      if (!demoRun) return;
      demoRun.stopped = true;
      demoRun.cursor.remove();
      demoRun = null;
      player.classList.remove("in-demo");
      if (message) flash(message);
    }

    async function demo() {
      if (demoRun) { stopDemo(); return; }
      if (!grid.hidden) toggleGrid();
      const run = { stopped: false, cursor: h("div", { class: "hs-demo-cursor", "aria-hidden": "true" }) };
      run.cursor.innerHTML = '<svg viewBox="0 0 32 32"><path d="M5 3 L5 25 L11 19.5 L15 29 L19.5 27 L15.5 18 L23 18 Z"/></svg>';
      demoRun = run;
      player.append(run.cursor);
      player.classList.add("in-demo");
      const wait = (ms) => new Promise((resolve, reject) => setTimeout(() => (run.stopped ? reject(new Error("stopped")) : resolve()), ms));
      const moveTo = (el, { x = 0.5, y = 0.5 } = {}) => {
        const base = player.getBoundingClientRect();
        const r = el.getBoundingClientRect();
        run.cursor.style.transform = `translate(${r.left - base.left + r.width * x}px, ${r.top - base.top + r.height * y}px)`;
      };
      const pointAt = async (el, where) => { moveTo(el, where); await wait(700); };
      const tap = async (el, fn, where) => {
        await pointAt(el, where);
        run.cursor.classList.add("tap");
        fn();
        await wait(220);
        run.cursor.classList.remove("tap");
      };
      const slideNow = () => current?.firstElementChild;
      flash("自動デモ：D・Esc・クリックで止まります");
      try {
        for (;;) {
          await wait(2400);
          let slide = slideNow();
          // Steps: items that appear (or measures that switch on) one per click.
          while (slide && step < stepsOf(slide)) {
            // Animations still playing are left to finish.
            for (let n = 0; n < 40 && E.animBusy?.(slide); n += 1) await wait(150);
            const target = slide.querySelector(`[data-measure="${step}"]`) || [...slide.querySelectorAll(`[data-g="${step}"]`)].find((el) => el.getBoundingClientRect().width);
            if (target) await tap(target, () => next()); else next();
            await wait(900);
            slide = slideNow();
          }
          // Views of a ranking.
          for (const tab of [...(slide?.querySelectorAll("[data-view]") || [])].slice(1)) { await tap(tab, () => tab.click()); await wait(1500); }
          // Sliders: to the far end and back.
          for (const range of slide?.querySelectorAll("input[data-sim]") || []) {
            const from = Number(range.value);
            const min = Number(range.min);
            const max = Number(range.max);
            const to = from + (max - min) * 0.35 <= max ? from + (max - min) * 0.35 : min + (max - min) * 0.25;
            const at = (v) => ({ x: (v - min) / (max - min || 1), y: 0.5 });
            await pointAt(range, at(from));
            run.cursor.classList.add("drag");
            for (let k = 1; k <= 30; k += 1) {
              range.value = String(from + ((to - from) * k) / 30);
              range.dispatchEvent(new Event("input", { bubbles: true }));
              moveTo(range, at(Number(range.value)));
              await wait(45);
            }
            run.cursor.classList.remove("drag");
            await wait(1400);
            range.value = String(from);
            range.dispatchEvent(new Event("input", { bubbles: true }));
            await wait(500);
          }
          // One piece of evidence, opened and closed again.
          const host = slide?.querySelector("[data-detail]");
          if (host && interaction) {
            await tap(host, () => interaction.openDetail(host.dataset.detail), { x: 0.9, y: 0.2 });
            await wait(3200);
            interaction.closeDetail();
            await wait(500);
          }
          if (place(index) >= order.length - 1 && !back) break;
          next();
        }
        await wait(1600);
        stopDemo(null);
        toggleGrid();
      } catch { /* stopped */ }
    }

    function flash(text) {
      const note = h("div", { class: "hs-player-flash" }, text);
      player.append(note);
      setTimeout(() => note.remove(), 2600);
    }

    function jumpMark(dir) {
      const all = [...(current?.firstElementChild?.querySelectorAll("video[data-bookmarks], audio[data-bookmarks]") || [])];
      const el = all.find((x) => !x.paused) || (all.length === 1 ? all[0] : null);
      if (!el || !E.mediaMarks) return;
      const times = E.mediaMarks(el).map((b) => b.t).sort((a, b) => a - b);
      const now = el.currentTime;
      const t = dir > 0 ? times.find((x) => x > now + 0.05) : [...times].reverse().find((x) => x < now - 0.5);
      if (t == null) return;
      el.currentTime = t;
      if (el.paused) E.mediaPlay?.(el);
    }

    // ---- input
    let digits = "";
    function onKey(event) {
      if (event.defaultPrevented) return;
      // アンケート: Enter or Space on a choice (reached with Tab) answers it.
      const choice = event.target?.closest?.(".hs-poll-opt");
      if (choice && (event.key === "Enter" || event.key === " ")) { event.preventDefault(); answer(choice); return; }
      if (viewer) { if ((event.key === "f" || event.key === "F") && !event.ctrlKey && !event.metaKey) toggleFullscreen(); return; }
      // PowerPoint's pen keys: Ctrl+P ペン, Ctrl+I 蛍光ペン, Ctrl+L レーザー, Ctrl+E 消しゴム, Ctrl+A 矢印.
      if ((event.ctrlKey || event.metaKey) && !event.altKey) {
        const k = event.key.toLowerCase();
        const mode = { p: "pen", i: "highlighter", l: "laser", e: "eraser" }[k];
        if (mode) { event.preventDefault(); if (penMode !== mode) setPen(mode); return; }
        if (k === "a") { event.preventDefault(); if (penMode) setPen(penMode); return; }
        return;
      }
      // Alt+End / Alt+Home: the next / previous bookmark of the media playing (or the slide's only one).
      if (event.altKey) { if (event.key === "End" || event.key === "Home") { event.preventDefault(); jumpMark(event.key === "End" ? 1 : -1); } return; }
      const key = event.key;
      if (kiosk && key !== "Escape") return;
      if (penMode && key === "Escape") { event.preventDefault(); setPen(penMode); return; }
      // 拡大: ＋ / − zoom where the pointer is; Esc shows the whole slide first.
      if (key === "+" || key === "=") { event.preventDefault(); zoomStep(1); return; }
      if (key === "-") { event.preventDefault(); zoomStep(-1); return; }
      if (key === "Escape" && resetZoom()) { event.preventDefault(); return; }
      if (key === "e" || key === "E") { eraseAll(); return; }
      if (key === "j" || key === "J") { toggleCaptions(); return; }
      gesture = true;
      if (key === "d" || key === "D") { demo(); return; }
      if (demoRun) { stopDemo(); if (key === "Escape") return; }
      // A slider in use keeps its arrow keys.
      if (event.target?.matches?.("input[data-sim]") && /^Arrow/.test(key)) return;
      if (["ArrowRight", "ArrowDown", "PageDown", " ", "Enter"].includes(key)) {
        event.preventDefault();
        if (key === "Enter" && digits) { go(order[Math.min(order.length, Math.max(1, Number(digits))) - 1]); digits = ""; return; }
        next();
      } else if (["ArrowLeft", "ArrowUp", "PageUp", "Backspace"].includes(key)) { event.preventDefault(); prev(); }
      else if (key === "Home") { event.preventDefault(); go(order[0]); }
      else if (key === "End") { event.preventDefault(); go(order[order.length - 1]); }
      else if (/^[0-9]$/.test(key)) { digits = `${digits}${key}`.slice(-3); }
      else if (key === "Escape") {
        if (interaction?.detailOpen) interaction.closeDetail();
        else if (!grid.hidden) toggleGrid();
        else if (!black.hidden) black.hidden = true;
        else if (back) closeDrill();
        else if (opts.closable !== false) close();
      } else if (key === "f" || key === "F") toggleFullscreen();
      else if (key === "n" || key === "N") toggleNotes();
      else if ((key === "q" || key === "Q") && opts.joinCard) toggleJoin();
      else if (key === "g" || key === "G" || key === "o" || key === "O") toggleGrid();
      else if (key === "p" || key === "P") openPresenterView();
      else if (key === "b" || key === "B" || key === ".") { black.style.background = "#000"; black.hidden = !black.hidden; }
      else if (key === "w" || key === "W") { black.style.background = "#fff"; black.hidden = !black.hidden; }
    }
    /** A shape, picture or text box with a link or an action (the editor's 「リンク・動作」) does that instead of advancing. */
    function act(el) {
      const own = (slides[index]?.elements || []).find((o) => o.id === el.dataset.el);
      if (own?.kind === "zoom") { zoomTo(own, el); return true; }
      const action = own?.action;
      if (!action) return false;
      runAction(action);
      return true;
    }
    /** What a click or a mouse-over action does: its sound (サウンドの再生), then the jump or the link. */
    function runAction(action) {
      if (action.sound && opts.sounds !== false && !opts.static) playSound(action.sound);
      if (action.type === "next") next();
      else if (action.type === "prev") prev();
      else if (action.type === "first") go(order[0]);
      else if (action.type === "last") go(order[order.length - 1]);
      else if (action.type === "end") { if (opts.closable !== false) close(); }
      else if (action.type === "slide") {
        const to = slides.findIndex((slide) => slide.sid === action.to);
        if (to >= 0) go(to);
        else flash("リンク先のスライドが見つかりません");
      } else if (action.type === "url") win.open(action.href, "_blank", "noopener");
    }
    // マウスの通過: the pointer moving onto an object with a mouse-over action does it once (again only after it
    // has left the object). A viewer only watches; a pen in hand writes over the slide instead.
    let overNow = null;
    const onStageOver = (event) => {
      const el = event.target.closest?.(".hs-obj[data-over]");
      if (el === overNow) return;
      overNow = el;
      if (!el || viewer || (penMode && penMode !== "laser") || zoom.level > 1 || event.pointerType === "touch") return;
      const own = (slides[index]?.elements || []).find((o) => o.id === el.dataset.el);
      if (own?.overAction) runAction(own.overAction);
    };
    const onStageOut = (event) => { if (overNow && !overNow.contains(event.relatedTarget)) overNow = null; };
    // 拡大 (PowerPoint's zoom in a slide show): the magnifier on the bar (or ＋) and a click where to look closer;
    // a drag looks around, − / Esc / a right-click shows the whole slide again; another slide starts whole.
    const zoom = { level: 1, tx: 0, ty: 0, picking: false, drag: null, moved: false };
    const ZOOM_STEPS = [1, 1.5, 2, 3, 4];
    const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
    function applyZoom() {
      stage.style.transformOrigin = "0 0";
      stage.style.transform = zoom.level > 1 ? `translate(${zoom.tx}px, ${zoom.ty}px) scale(${zoom.level})` : "";
      player.classList.toggle("zoomed", zoom.level > 1);
      player.classList.toggle("zoom-pick", zoom.picking);
      requestAnimationFrame(placeInk);
    }
    /** Zooms to `level` keeping the point (cx, cy) on the screen where it is (the middle when none). */
    function magnify(level, cx = null, cy = null) {
      const w = stage.offsetWidth;
      const ht = stage.offsetHeight;
      const base = player.getBoundingClientRect();
      const lx = cx == null ? w / 2 : cx - base.left - stage.offsetLeft;
      const ly = cy == null ? ht / 2 : cy - base.top - stage.offsetTop;
      const px = (lx - zoom.tx) / zoom.level;
      const py = (ly - zoom.ty) / zoom.level;
      zoom.level = clamp(level, 1, 4);
      zoom.tx = zoom.level > 1 ? clamp(lx - px * zoom.level, w - w * zoom.level, 0) : 0;
      zoom.ty = zoom.level > 1 ? clamp(ly - py * zoom.level, ht - ht * zoom.level, 0) : 0;
      zoom.picking = false;
      applyZoom();
    }
    function zoomStep(d) {
      const k = ZOOM_STEPS.findIndex((z) => z >= zoom.level - 1e-6);
      magnify(ZOOM_STEPS[clamp((k < 0 ? ZOOM_STEPS.length - 1 : k) + d, 0, ZOOM_STEPS.length - 1)], zoom.pointer?.x ?? null, zoom.pointer?.y ?? null);
    }
    function resetZoom() {
      if (zoom.level === 1 && !zoom.picking) return false;
      Object.assign(zoom, { level: 1, tx: 0, ty: 0, picking: false, drag: null });
      applyZoom();
      return true;
    }
    function toggleMagnifier() {
      if (resetZoom()) return;
      zoom.picking = true;
      applyZoom();
      flash("拡大する場所をクリックしてください（Esc・右クリックで戻る）");
    }
    stage.addEventListener("pointerdown", (event) => {
      if (zoom.level <= 1 || event.button !== 0 || (penMode && penMode !== "laser")) return;
      zoom.drag = { x: event.clientX, y: event.clientY, tx: zoom.tx, ty: zoom.ty };
      zoom.moved = false;
      player.classList.add("zoom-drag");
    });
    player.addEventListener("pointermove", (event) => {
      zoom.pointer = { x: event.clientX, y: event.clientY };
      if (!zoom.drag) return;
      const dx = event.clientX - zoom.drag.x;
      const dy = event.clientY - zoom.drag.y;
      if (Math.abs(dx) + Math.abs(dy) > 3) zoom.moved = true;
      zoom.tx = clamp(zoom.drag.tx + dx, stage.offsetWidth * (1 - zoom.level), 0);
      zoom.ty = clamp(zoom.drag.ty + dy, stage.offsetHeight * (1 - zoom.level), 0);
      applyZoom();
    });
    const endDrag = () => { if (zoom.drag) { zoom.drag = null; player.classList.remove("zoom-drag"); } };
    player.addEventListener("pointerup", endDrag);
    player.addEventListener("pointercancel", endDrag);
    stage.addEventListener("contextmenu", (event) => { if (resetZoom()) event.preventDefault(); });
    stage.addEventListener("transitionend", () => placeInk());

    const onStageClick = (event) => {
      if (event.target.closest(".hs-player-bar, .hs-player-grid, .hs-player-notes, .hs-control, a[href], .hs-pen-menu, .hs-join")) return;
      // アンケート: a click on a choice answers (it does not move the show on).
      const choice = event.target.closest(".hs-poll-opt");
      if (choice) { answer(choice); return; }
      // 拡大: the click says where to look closer; zoomed in, a click (or a drag) looks around and does not move on.
      if (zoom.picking && !viewer) { magnify(2, event.clientX, event.clientY); return; }
      if (zoom.level > 1) return;
      // With a pen in hand, a click writes instead of moving on (the laser pointer still clicks through).
      if (penMode && penMode !== "laser") return;
      if (viewer) return;
      // A kiosk show only answers its buttons and links.
      if (kiosk) { const actor = event.target.closest('.hs-obj[data-action], .hs-obj[data-kind="zoom"]'); if (actor) act(actor); return; }
      gesture = true;
      origin = { x: event.clientX, y: event.clientY };
      const actor = event.target.closest('.hs-obj[data-action], .hs-obj[data-kind="zoom"]');
      if (actor && act(actor)) return;
      next();
    };
    let touchX = null;
    const onTouchStart = (event) => { touchX = viewer || event.target.closest?.(".hs-control") ? null : event.touches[0]?.clientX ?? null; };
    const onTouchEnd = (event) => {
      if (touchX == null) return;
      const dx = (event.changedTouches[0]?.clientX ?? touchX) - touchX;
      if (Math.abs(dx) > 60) { event.preventDefault(); dx < 0 ? next() : prev(); }
      touchX = null;
    };
    let idle = null;
    const onMove = () => {
      player.classList.remove("idle");
      clearTimeout(idle);
      idle = setTimeout(() => player.classList.add("idle"), 2600);
    };
    if (opts.keyboard !== false) doc.addEventListener("keydown", onKey);
    stage.addEventListener("click", onStageClick);
    stage.addEventListener("pointerover", onStageOver);
    stage.addEventListener("pointerout", onStageOut);
    stage.addEventListener("touchstart", onTouchStart, { passive: true });
    stage.addEventListener("touchend", onTouchEnd);
    player.addEventListener("pointermove", onMove);
    // A tap (a phone has no pointer to move) brings the bar back too: a viewer's reactions are on it.
    player.addEventListener("pointerdown", onMove);
    // The audience's own click stops the automatic demo (the demo's clicks are not "trusted").
    player.addEventListener("pointerdown", (event) => { if (event.isTrusted && demoRun && !event.target.closest?.(".hs-player-btn.demo")) stopDemo(); }, true);
    onMove();

    function destroy() {
      if (demoRun) demoRun.stopped = true;
      clearTimeout(autoTimer);
      doc.removeEventListener("keydown", onKey);
      interaction?.destroy();
      if (current?.firstElementChild) stopMedia(current.firstElementChild);
      for (const c of carried.splice(0)) E.mediaPause?.(c.audio, { stop: true });
      stopSounds();
      inkFollow?.disconnect();
      camera?.then((stream) => stream.getTracks().forEach((t) => t.stop())).catch(() => {});
      if (captions) { const rec = captions; captions = null; try { rec.stop(); } catch { /* stopped */ } }
      clearTimeout(idle);
      if (pv?.win && !pv.win.closed) pv.win.close();
      clearInterval(pv?.timer);
      player.remove();
    }

    if (story.parent[index] != null) back = { index: story.parent[index], step: Infinity, at: null };
    show(index, { dir: 1, fullStep: Boolean(opts.fullStep) });
    player.focus({ preventScroll: true });
    // スライド ショー →「常に字幕を使用」.
    if (opts.captions) toggleCaptions();

    return {
      el: player, next, prev, go, follow, destroy, close, react, caption, setVotes, toggleJoin, magnify, resetZoom, toggleMagnifier, get zoom() { return zoom.level; }, toggleNotes, toggleGrid, toggleFullscreen, openPresenterView, openDrill, closeDrill, demo, stopDemo, setPen, toggleCaptions,
      get pen() { return penMode; }, get captioning() { return Boolean(captions); }, get order() { return [...order]; },
      get index() { return index; }, get step() { return step; }, get startedAt() { return started; }, get inDrill() { return Boolean(back); },
    };
  }

  // ---------------------------------------------------------------- 変形 (Morph)

  /**
   * 画面切り替え「変形」: the objects of the slide being left (`prevSlide`) that are the same as objects of the slide
   * coming in (`slide`; `pairs` from E.morphPairs) move, resize and turn into them; what has no partner fades out where it
   * was and fades in where it is. The old objects are taken over into a layer above the new slide, so call this before
   * the old slide goes. Resolves when it is over.
   */
  function morphIn(prevSlide, slide, pairs, { dur = 900 } = {}) {
    const layerA = prevSlide?.querySelector(":scope > .hs-objects");
    const layerB = slide?.querySelector(":scope > .hs-objects");
    if (!layerA || !layerB) return Promise.resolve();
    const nodeOf = (layer, o) => [...layer.children].find((n) => n.dataset.el === o.id);
    const box = (o) => (o.kind === "line" ? E.bounds(o) : { x: o.x, y: o.y, w: o.w, h: o.h });
    const centre = (b) => [b.x + b.w / 2, b.y + b.h / 2];
    const ease = "cubic-bezier(.65, 0, .35, 1)";
    const ghosts = h("div", { class: "hs-morph-layer", "aria-hidden": "true" });
    slide.append(ghosts);
    const take = (node) => { node.removeAttribute("data-el"); node.style.pointerEvents = "none"; ghosts.append(node); return node; };
    const move = (from, to) => {
      const [ax, ay] = centre(box(from));
      const [bx, by] = centre(box(to));
      const a = box(from);
      const b = box(to);
      return { dx: bx - ax, dy: by - ay, sx: Math.min(30, Math.max(0.03, (b.w || 1) / (a.w || 1))), sy: Math.min(30, Math.max(0.03, (b.h || 1) / (a.h || 1))), rot: (to.rot || 0) - (from.rot || 0) };
    };
    const matchedFrom = new Set();
    const matchedTo = new Set();
    const animations = [];
    const play = (node, keyframes, options) => { try { animations.push(node.animate(keyframes, { duration: dur, easing: ease, ...options })); } catch { /* detached */ } };
    for (const [from, to] of pairs) {
      const nodeA = nodeOf(layerA, from);
      const nodeB = nodeOf(layerB, to);
      if (!nodeA || !nodeB) continue;
      matchedFrom.add(from);
      matchedTo.add(to);
      const m = move(from, to);
      // The old part travels to where the new one is and gives way to it; the new one travels from where the old was.
      take(nodeA);
      play(nodeA, [
        { transform: "none", opacity: 1, offset: 0 },
        { opacity: 1, offset: 0.55 },
        { transform: `translate(${m.dx}px, ${m.dy}px) rotate(${m.rot}deg) scale(${m.sx}, ${m.sy})`, opacity: 0, offset: 1 },
      ], { fill: "forwards" });
      play(nodeB, [
        { transform: `translate(${-m.dx}px, ${-m.dy}px) rotate(${-m.rot}deg) scale(${1 / m.sx}, ${1 / m.sy})`, opacity: 0, offset: 0 },
        { opacity: 1, offset: 0.45 },
        { transform: "none", opacity: 1, offset: 1 },
      ], { fill: "backwards" });
    }
    // Parts with no partner: the old ones fade out where they were, the new ones fade in.
    for (const node of [...layerA.children]) {
      if (!node.classList?.contains("hs-obj")) continue;
      if (pairs.some(([from]) => matchedFrom.has(from) && node.dataset.el === from.id)) continue;
      take(node);
      play(node, [{ opacity: 1 }, { opacity: 0 }], { duration: dur * 0.5, easing: "ease-out", fill: "forwards" });
    }
    for (const node of layerB.children) {
      if (!node.classList.contains("hs-obj") || pairs.some(([, to]) => matchedTo.has(to) && node.dataset.el === to.id)) continue;
      play(node, [{ opacity: 0 }, { opacity: 1 }], { duration: dur * 0.6, delay: dur * 0.3, easing: "ease-in", fill: "backwards" });
    }
    return new Promise((resolve) => setTimeout(() => { for (const a of animations) { try { a.cancel(); } catch { /* done */ } } ghosts.remove(); resolve(); }, dur + 60));
  }

  /**
   * Play one slide transition in a box (the studio's 画面切り替え「プレビュー」): `from` gives way to `to`, as the
   * player would show it. Resolves when it is over.
   */
  function transitionPreview(host, fromEl, toEl, type, { dur = null, dir = null, pairs = null } = {}) {
    const kind = TRANSITIONS.has(type) && type !== "morph" ? type : type === "morph" ? "fade" : "none";
    host.replaceChildren();
    const prev = E.mount(fromEl, { contain: true, className: "hs-player-slide" });
    const next = E.mount(toEl, { contain: true, className: "hs-player-slide" });
    host.append(prev);
    E.scale(prev);
    // 変形 with parts the two slides share: they travel from one to the other.
    if (type === "morph" && pairs?.length) {
      host.append(next);
      E.scale(next);
      const finished = morphIn(fromEl, toEl, pairs, { dur: dur >= 100 ? dur : 900 });
      prev.remove();
      return finished;
    }
    if (kind === "none") { prev.remove(); host.append(next); E.scale(next); return Promise.resolve(); }
    host.append(next);
    E.scale(next);
    const entering = CLIPPED.has(kind) ? toEl : next;
    const option = transitionOption(kind, dir, false, entering, prev);
    let band = null;
    if (kind === "wipe" && option.band) {
      const tone = getComputedStyle(toEl);
      band = h("div", { class: `hs-tr-band${option.rev ? " rev" : ""}` });
      band.style.setProperty("--band", tone.getPropertyValue("--accent").trim() || "#2451e6");
      band.style.setProperty("--band2", tone.getPropertyValue("--accent2").trim() || "#13a89e");
      toEl.append(band);
    }
    entering.className += ` hs-tr-in-${kind}${option.rev ? " rev" : ""}`;
    prev.className += ` hs-tr-out-${kind}${option.rev ? " rev" : ""}`;
    const ms = dur >= 100 ? dur : TRANSITION_MS[kind] ?? 620;
    if (dur >= 100) for (const el of [entering, prev, band].filter(Boolean)) el.style.animationDuration = `${dur}ms`;
    return new Promise((resolve) => setTimeout(() => { prev.remove(); band?.remove(); entering.classList.remove(`hs-tr-in-${kind}`, "rev", "tr-dir", "tr-h"); resolve(); }, ms + 40));
  }

  const PV_CSS = `
html,body{margin:0;height:100%;background:#0d1017;color:#e8ecf4;font-family:"Noto Sans JP","Hiragino Sans",sans-serif}
.pv{display:grid;grid-template-columns:minmax(0,1.6fr) minmax(280px,1fr);grid-template-rows:minmax(0,1fr) minmax(140px,34%);gap:14px;height:100%;padding:14px;box-sizing:border-box}
.pv-now,.pv-next{position:relative;overflow:hidden;background:#000;border-radius:10px}
.pv-now{grid-row:1;grid-column:1}.pv-side{grid-row:1;grid-column:2;display:flex;flex-direction:column;gap:10px;min-height:0}
.pv-next{flex:0 0 auto;aspect-ratio:16/9;opacity:.9}.pv-label{font-size:12px;letter-spacing:.1em;color:#8b95a8;font-weight:700}
.pv-clock{display:flex;align-items:baseline;gap:14px;margin-top:6px}.pv-timer{font-size:44px;font-variant-numeric:tabular-nums;font-weight:800}.pv-count{color:#9aa5b8;font-size:15px}
.pv-btns{display:flex;flex-wrap:wrap;gap:8px;margin-top:auto}.pv-btns button{flex:1;min-width:90px;height:40px;border:1px solid #2c3446;border-radius:10px;background:#161b26;color:#e8ecf4;font:inherit;font-weight:700;cursor:pointer}
.pv-notes{grid-row:2;grid-column:1/-1;overflow:auto;padding:16px 20px;border-radius:10px;background:#161b26;font-size:22px;line-height:1.75;white-space:pre-wrap}
.pv-notes p{margin:0 0 .35em}.pv-notes ul,.pv-notes ol{margin:.1em 0 .35em;padding-left:1.5em}.pv-notes a{color:#9bb8ff}
.pv .hs-slide{position:absolute;top:0;left:0}`;

  /** How long a transition takes when the slide does not say (ms). */
  const transitionMs = (type) => (type === "none" ? 0 : TRANSITION_MS[type] ?? 620);

  Object.assign(E, { play, reveal, stepsOf, countUp, activate, playMedia, stopMedia, mountLottie, stopLottie, splitKinetic, createPlayer, engineCss, transitionPreview, transitionMs, morphIn, playSound, stopSounds, renderSound });
})(typeof window !== "undefined" ? window : globalThis);
