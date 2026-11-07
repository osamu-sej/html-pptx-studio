// Sound and video, PowerPoint style: 挿入 → オーディオ (a file, a recording made here, a URL) and 画面録画 (the
// screen, with the microphone if wanted), and the 再生 tab of a selected video or sound: preview, trim, fades,
// volume, how it starts (automatically, on a click on it, or in the click sequence), play across slides, loop,
// hide the icon, rewind, 全画面, 再生中のみ表示 and バックグラウンドで再生. The engine plays them
// (objects.js mediaPlay; motion.js carries a sound across slides).

const STARTS = { sequence: "一連のクリック動作", auto: "自動", click: "クリック時" };
const fmt = (sec) => {
  if (!Number.isFinite(sec)) return "--:--";
  const m = Math.floor(sec / 60);
  const s = sec - m * 60;
  return `${m}:${s < 10 ? "0" : ""}${s.toFixed(1)}`;
};

/** A recorded WebM reports no length until it has been read to the end: ask for a far point once to learn it. */
/** SubRip (.srt) subtitles as WebVTT (the timestamps' commas become dots); WebVTT is kept as it is. */
export function toVtt(text) {
  const t = String(text ?? "").replace(/^\uFEFF/, "").replace(/\r/g, "").trim();
  if (/^WEBVTT/.test(t)) return t;
  return `WEBVTT\n\n${t.replace(/(\d{2}:\d{2}:\d{2}),(\d{3})/g, "$1.$2")}`;
}

function knowDuration(el) {
  return new Promise((resolve) => {
    const done = () => resolve(Number.isFinite(el.duration) ? el.duration : NaN);
    const ready = () => {
      if (Number.isFinite(el.duration)) return done();
      const back = () => { el.removeEventListener("durationchange", back); el.currentTime = 0; done(); };
      el.addEventListener("durationchange", back);
      try { el.currentTime = 1e101; } catch { done(); }
      setTimeout(done, 3000);
    };
    if (el.readyState >= 1) ready(); else el.addEventListener("loadedmetadata", ready, { once: true });
    el.addEventListener("error", () => resolve(NaN), { once: true });
  });
}

export function createMedia(editor, app, kit) {
  const { E, h } = app;
  const { btn, drop, group, col, menu, updater } = kit;
  const selected = () => editor.selectedObjects();
  const playable = () => selected().filter((o) => o.kind === "audio" || (o.kind === "video" && !E.youtubeId(o.src)));
  const current = () => (playable().length === 1 ? E.withDefaults(playable()[0]) : null);
  const has = () => selected().some((o) => o.kind === "audio" || o.kind === "video");
  const urlOf = (o) => E.resolveSrc(o.src, { ...app.renderOptions(), mode: "present" });
  const setMedia = (patch) => editor.apply((o) => (o.kind === "audio" || o.kind === "video" ? (typeof patch === "function" ? patch(o) : patch) : null));

  // ---------------------------------------------------------------- inserting

  async function insertAudioFile() {
    const files = await app.pickFiles("audio/*,.mp3,.m4a,.wav,.ogg,.aac,.webm", false);
    if (files.length) await app.insertFiles(files);
  }
  async function insertAudioUrl() {
    const url = await app.ask("URLからオーディオを入れる", "音声ファイルのURL（https://…/sound.mp3 など）", "");
    if (url) await app.insertUrl(url.trim(), "audio");
  }

  /**
   * オーディオの録音 / 画面録画: a small recorder (MediaRecorder). A recording is kept in this browser like an
   * uploaded file and travels inside the exported HTML.
   */
  function recorder(kind) {
    const screen = kind === "screen";
    let stream = null;
    let rec = null;
    let chunks = [];
    let blob = null;
    let started = 0;
    let timer = 0;
    let meterRaf = 0;
    let audioCtx = null;
    const status = h("p", { class: "rec-status", role: "status" }, screen ? "「録画を開始」で録画する画面・ウィンドウ・タブを選びます。" : "「録音」でマイクから録音します。");
    const clock = h("b", { class: "rec-clock" }, "0:00.0");
    const meter = h("span", { class: "rec-meter", "aria-hidden": "true" }, h("i"));
    const preview = screen ? h("video", { class: "rec-preview", playsinline: true, muted: true, controls: false }) : h("audio", { class: "rec-preview", controls: true, hidden: true });
    const mic = h("input", { type: "checkbox", checked: true });
    const name = h("input", { type: "text", class: "rec-name", value: screen ? "画面録画" : "録音", "aria-label": "名前" });
    const startBtn = h("button", { type: "button", class: "btn btn-primary rec-start", onclick: () => start() }, screen ? "録画を開始" : "● 録音");
    const stopBtn = h("button", { type: "button", class: "btn rec-stop", disabled: true, onclick: () => stop() }, "■ 停止");
    const insertBtn = h("button", { type: "button", class: "btn btn-primary rec-insert", disabled: true, onclick: () => insert() }, "挿入");
    const dialog = h("dialog", { class: "rec-dialog", "aria-label": screen ? "画面録画" : "オーディオの録音" },
      h("div", { class: "dialog-head" }, h("h3", {}, screen ? "画面録画" : "オーディオの録音"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body" },
        status,
        h("div", { class: "rec-row" }, clock, meter),
        screen ? h("label", { class: "fp-check" }, mic, "マイクの音声も録音する") : null,
        preview,
        h("label", { class: "field" }, h("span", { class: "field-label" }, "名前"), name)),
      h("div", { class: "dialog-foot" }, startBtn, stopBtn, h("span", { style: { flex: "1" } }), h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "やめる"), insertBtn));
    document.body.append(dialog);
    const tickClock = () => { clock.textContent = fmt((Date.now() - started) / 1000); };
    function watchLevel(source) {
      try {
        audioCtx = new AudioContext();
        const analyser = audioCtx.createAnalyser();
        analyser.fftSize = 512;
        audioCtx.createMediaStreamSource(source).connect(analyser);
        const data = new Uint8Array(analyser.fftSize);
        const loop = () => {
          analyser.getByteTimeDomainData(data);
          let peak = 0;
          for (const v of data) peak = Math.max(peak, Math.abs(v - 128));
          meter.firstChild.style.width = `${Math.min(100, (peak / 128) * 160)}%`;
          meterRaf = requestAnimationFrame(loop);
        };
        loop();
      } catch { /* no level meter */ }
    }
    async function start() {
      blob = null;
      chunks = [];
      insertBtn.disabled = true;
      try {
        if (screen) {
          const display = await navigator.mediaDevices.getDisplayMedia({ video: { frameRate: 30 }, audio: true });
          const tracks = [...display.getTracks()];
          if (mic.checked) {
            try { const voice = await navigator.mediaDevices.getUserMedia({ audio: true }); tracks.push(...voice.getAudioTracks()); } catch { status.textContent = "マイクは使えませんでした（画面だけを録画します）"; }
          }
          // Several sound sources (the tab and the microphone) are mixed into one track.
          const sounds = tracks.filter((t) => t.kind === "audio");
          if (sounds.length > 1) {
            audioCtx = new AudioContext();
            const out = audioCtx.createMediaStreamDestination();
            for (const t of sounds) audioCtx.createMediaStreamSource(new MediaStream([t])).connect(out);
            stream = new MediaStream([...tracks.filter((t) => t.kind === "video"), ...out.stream.getAudioTracks()]);
            stream.hsSources = tracks;
          } else stream = new MediaStream(tracks);
          stream.hsSources ||= tracks;
          preview.srcObject = stream;
          preview.muted = true;
          preview.controls = false;
          preview.play().catch(() => {});
          // Ending the share from the browser's own bar stops the recording too.
          display.getVideoTracks()[0]?.addEventListener("ended", () => stop());
        } else {
          stream = await navigator.mediaDevices.getUserMedia({ audio: { echoCancellation: true, noiseSuppression: true } });
          stream.hsSources = stream.getTracks();
          watchLevel(stream);
        }
      } catch (error) {
        status.textContent = screen ? `画面を録画できません（${error.message || error.name}）` : `マイクを使えません（${error.message || error.name}）。ブラウザでマイクを許可してください`;
        return;
      }
      const types = screen ? ["video/webm;codecs=vp9,opus", "video/webm;codecs=vp8,opus", "video/webm", "video/mp4"] : ["audio/webm;codecs=opus", "audio/webm", "audio/mp4", "audio/ogg"];
      const mimeType = types.find((t) => window.MediaRecorder?.isTypeSupported?.(t)) || "";
      try { rec = new MediaRecorder(stream, mimeType ? { mimeType } : undefined); } catch (error) { status.textContent = `録音できません（${error.message}）`; release(); return; }
      rec.ondataavailable = (event) => { if (event.data?.size) chunks.push(event.data); };
      rec.onstop = () => {
        blob = new Blob(chunks, { type: rec.mimeType || mimeType || (screen ? "video/webm" : "audio/webm") });
        release();
        const url = URL.createObjectURL(blob);
        preview.srcObject = null;
        preview.src = url;
        preview.hidden = false;
        preview.muted = false;
        preview.controls = true;
        insertBtn.disabled = !blob.size;
        status.textContent = blob.size ? `${screen ? "録画" : "録音"}しました（${fmt((Date.now() - started) / 1000)}）。再生して確かめてから「挿入」。` : "何も録れていません。もう一度どうぞ";
      };
      rec.start(250);
      started = Date.now();
      timer = setInterval(tickClock, 100);
      startBtn.disabled = true;
      stopBtn.disabled = false;
      status.textContent = screen ? "録画中…（「停止」か、ブラウザの「共有を停止」で終わります）" : "録音中…";
      dialog.classList.add("recording");
    }
    function release() {
      clearInterval(timer);
      cancelAnimationFrame(meterRaf);
      for (const t of stream?.hsSources || stream?.getTracks() || []) t.stop();
      stream = null;
      audioCtx?.close().catch(() => {});
      audioCtx = null;
      startBtn.disabled = false;
      stopBtn.disabled = true;
      startBtn.textContent = screen ? "録画し直す" : "● 録音し直す";
      dialog.classList.remove("recording");
    }
    function stop() {
      if (rec && rec.state !== "inactive") rec.stop();
      else release();
    }
    async function insert() {
      if (!blob?.size) return;
      const ext = blob.type.includes("mp4") ? "mp4" : blob.type.includes("ogg") ? "ogg" : "webm";
      await app.insertRecording(blob, { kind: screen ? "video" : "audio", name: `${(name.value || (screen ? "画面録画" : "録音")).trim().slice(0, 60)}.${ext}` });
      dialog.close();
    }
    dialog.addEventListener("close", () => { if (rec && rec.state !== "inactive") { rec.onstop = null; rec.stop(); } release(); dialog.remove(); });
    dialog.showModal();
  }

  // ---------------------------------------------------------------- previewing in the editor

  let player = null; // { id, el }
  function stopPreview() {
    if (!player) return;
    E.mediaPause(player.el, { stop: true });
    if (player.own) player.el.remove();
    player = null;
    kit.refreshRibbon?.();
  }
  /** 再生 (プレビュー): the selected sound or video plays here, trimmed and faded as it will in the show. */
  function preview() {
    const o = current();
    if (!o) return;
    if (player?.id === o.id) { stopPreview(); return; }
    stopPreview();
    let el = o.kind === "video" ? document.querySelector(`.slide-wrap .hs-obj[data-el="${CSS.escape(o.id)}"] video`) : null;
    let own = false;
    if (!el) {
      const url = urlOf(o);
      if (!url) { app.toast("このブラウザにメディアのデータがありません"); return; }
      el = h("audio", { src: url, preload: "auto", hidden: true });
      document.body.append(el);
      own = true;
    }
    for (const [key, value] of Object.entries({ "trim-start": o.trimStart, "trim-end": o.trimEnd, "fade-in": o.fadeIn, "fade-out": o.fadeOut, volume: o.volume, rewind: o.rewind ? "" : null, loop: o.loop ? "" : null })) {
      if (value == null) el.removeAttribute(`data-${key}`); else el.setAttribute(`data-${key}`, String(value));
    }
    el.muted = false;
    player = { id: o.id, el, own };
    el.addEventListener("pause", () => { if (player?.el === el && (el.ended || el.paused)) { if (own) el.remove(); player = null; kit.refreshRibbon?.(); } }, { once: true });
    E.mediaPlay(el, { fromStart: true });
    kit.refreshRibbon?.();
  }

  // ---------------------------------------------------------------- トリミング

  /** 表紙画像 → ビデオの1コマから: a small player to find the moment, and that frame becomes the poster. */
  function posterFromFrame() {
    const o = current();
    if (o?.kind !== "video") return;
    const url = urlOf(o);
    if (!url) { app.toast("このブラウザにビデオのデータがありません"); return; }
    const el = h("video", { src: url, preload: "auto", class: "trim-media", controls: true, playsinline: true, muted: true });
    const take = () => {
      if (!el.videoWidth) { app.toast("ビデオを読み込んでから選んでください"); return; }
      const k = Math.min(1, 1280 / el.videoWidth);
      const c = document.createElement("canvas");
      c.width = Math.round(el.videoWidth * k);
      c.height = Math.round(el.videoHeight * k);
      try {
        c.getContext("2d").drawImage(el, 0, 0, c.width, c.height);
        const poster = c.toDataURL("image/jpeg", 0.86);
        setMedia((x) => (x.kind === "video" ? { poster } : null));
        dialog.close();
        app.toast("表紙画像にしました（⌘Zで戻せます）");
      } catch { app.toast("このビデオの画面は画像にできません"); }
    };
    const dialog = h("dialog", { class: "poster-dialog", "aria-label": "表紙画像" },
      h("div", { class: "dialog-head" }, h("h3", {}, "表紙画像：ビデオの1コマから"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body" }, el, h("p", { class: "hint" }, "再生・シークで表紙にしたい場面にして「この場面を表紙画像にする」を押してください。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary poster-take", onclick: take }, "この場面を表紙画像にする")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => { el.pause(); dialog.remove(); });
    dialog.showModal();
  }
  /** 表紙画像 → ファイルから画像 (made no larger than 1280 px across). */
  async function posterFromFile() {
    const [file] = await app.pickFiles("image/*", false);
    if (!file) return;
    const img = new Image();
    img.src = URL.createObjectURL(file);
    try { await img.decode(); } catch { app.toast("画像を読めませんでした"); return; }
    const k = Math.min(1, 1280 / img.naturalWidth);
    const c = document.createElement("canvas");
    c.width = Math.round(img.naturalWidth * k);
    c.height = Math.round(img.naturalHeight * k);
    c.getContext("2d").drawImage(img, 0, 0, c.width, c.height);
    URL.revokeObjectURL(img.src);
    const poster = c.toDataURL("image/jpeg", 0.86);
    setMedia((x) => (x.kind === "video" ? { poster } : null));
    app.toast("表紙画像にしました（⌘Zで戻せます）");
  }

  function trimDialog() {
    const o = current();
    if (!o) return;
    const url = urlOf(o);
    if (!url) { app.toast("このブラウザにメディアのデータがありません"); return; }
    const isVideo = o.kind === "video";
    const el = h(isVideo ? "video" : "audio", { src: url, preload: "auto", class: "trim-media", controls: true, playsinline: true });
    let duration = NaN;
    let start = o.trimStart || 0;
    let end = o.trimEnd || null;
    const startIn = h("input", { type: "number", step: "0.1", min: "0", class: "trim-num", "aria-label": "開始時間（秒）" });
    const endIn = h("input", { type: "number", step: "0.1", min: "0", class: "trim-num", "aria-label": "終了時間（秒）" });
    const range = h("div", { class: "trim-range" }, h("i", { class: "trim-sel" }), h("b", { class: "trim-head" }));
    const info = h("p", { class: "hint trim-info" });
    const sync = () => {
      const len = Number.isFinite(duration) ? duration : Math.max(end || 0, start + 1);
      const e = end ?? len;
      startIn.value = start.toFixed(1);
      endIn.value = e.toFixed(1);
      range.querySelector(".trim-sel").style.left = `${(start / len) * 100}%`;
      range.querySelector(".trim-sel").style.width = `${Math.max(0, ((e - start) / len) * 100)}%`;
      info.textContent = `長さ ${fmt(e - start)}（元の長さ ${fmt(duration)}）`;
    };
    el.addEventListener("timeupdate", () => {
      const len = Number.isFinite(duration) ? duration : 1;
      range.querySelector(".trim-head").style.left = `${(el.currentTime / len) * 100}%`;
      if (end != null && el.currentTime >= end) el.pause();
    });
    startIn.addEventListener("change", () => { start = Math.max(0, Math.min(Number(startIn.value) || 0, (end ?? duration) - 0.1)); sync(); el.currentTime = start; });
    endIn.addEventListener("change", () => { const v = Number(endIn.value); end = Number.isFinite(v) && v > start + 0.1 ? Math.min(v, Number.isFinite(duration) ? duration : v) : null; if (end != null && Number.isFinite(duration) && end >= duration - 0.05) end = null; sync(); });
    // Dragging on the bar moves the nearer end of the kept part.
    range.addEventListener("pointerdown", (event) => {
      const len = Number.isFinite(duration) ? duration : null;
      if (!len) return;
      const r = range.getBoundingClientRect();
      const at = (x) => Math.max(0, Math.min(len, ((x - r.left) / r.width) * len));
      const t0 = at(event.clientX);
      const which = Math.abs(t0 - start) <= Math.abs(t0 - (end ?? len)) ? "start" : "end";
      const move = (e) => {
        const t = at(e.clientX);
        if (which === "start") start = Math.min(t, (end ?? len) - 0.1);
        else { end = Math.max(t, start + 0.1); if (end >= len - 0.05) end = null; }
        sync();
        el.currentTime = which === "start" ? start : (end ?? len);
      };
      range.setPointerCapture(event.pointerId);
      range.addEventListener("pointermove", move);
      range.addEventListener("pointerup", () => range.removeEventListener("pointermove", move), { once: true });
      move(event);
    });
    const dialog = h("dialog", { class: "trim-dialog", "aria-label": isVideo ? "ビデオのトリミング" : "オーディオのトリミング" },
      h("div", { class: "dialog-head" }, h("h3", {}, isVideo ? "ビデオのトリミング" : "オーディオのトリミング"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body" },
        el, range,
        h("div", { class: "trim-fields" },
          h("label", { class: "field" }, h("span", { class: "field-label" }, "開始時間（秒）"), startIn, h("button", { type: "button", class: "btn btn-sm", onclick: () => { start = Math.min(el.currentTime, (end ?? duration) - 0.1); sync(); } }, "今の位置")),
          h("label", { class: "field" }, h("span", { class: "field-label" }, "終了時間（秒）"), endIn, h("button", { type: "button", class: "btn btn-sm", onclick: () => { end = Math.max(el.currentTime, start + 0.1); if (Number.isFinite(duration) && end >= duration - 0.05) end = null; sync(); } }, "今の位置"))),
        info),
      h("div", { class: "dialog-foot" },
        h("button", { type: "button", class: "btn", onclick: () => { el.currentTime = start; el.play().catch(() => {}); } }, "▶ 範囲を再生"),
        h("span", { style: { flex: "1" } }),
        h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary trim-ok", onclick: () => { setMedia({ trimStart: start > 0.05 ? Math.round(start * 10) / 10 : undefined, trimEnd: end != null ? Math.round(end * 10) / 10 : undefined }); dialog.close(); } }, "OK")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => { el.pause(); dialog.remove(); });
    dialog.showModal();
    knowDuration(el).then((d) => { duration = d; sync(); el.currentTime = start; });
    sync();
  }

  // ---------------------------------------------------------------- ブックマーク

  const markId = () => `b${Date.now().toString(36).slice(-5)}${Math.random().toString(36).slice(2, 5)}`;
  /** Keep the media's bookmarks; an animation started by a bookmark that is gone goes back to the click order. */
  function saveMarks(id, marks) {
    const clean = E.normalizeBookmarks(marks);
    const keep = new Set(clean.map((b) => `${id}@${b.id}`));
    const list = editor.objects().map((x) => (x.id === id ? { ...x, bookmarks: clean.length ? clean : undefined } : x));
    const timeline = (app.slide()?.timeline || []).map((e) => {
      if (!String(e.trigger || "").startsWith(`${id}@`) || keep.has(e.trigger)) return e;
      const { trigger: _, ...rest } = e;
      return rest;
    });
    editor.commit(list, { timeline });
  }
  /** ブックマークの追加: while the preview plays, a bookmark at that moment; otherwise the bookmark editor. */
  function addBookmark() {
    const o = current();
    if (!o) return;
    if (player?.id === o.id && Number.isFinite(player.el.currentTime)) {
      const t = Math.round(player.el.currentTime * 100) / 100;
      const marks = [...(o.bookmarks || []), { id: markId(), t, name: "" }];
      saveMarks(o.id, marks);
      app.toast(`${fmt(t)} にブックマークを付けました（アニメーションの「トリガー → ブックマーク時」で使えます）`);
      return;
    }
    bookmarkDialog();
  }
  function removeMenu() {
    const o = current();
    const marks = o?.bookmarks || [];
    if (!marks.length) return menu([{ label: "（ブックマークがありません）", disabled: true }]);
    return menu([
      { head: "ブックマークの削除" },
      ...marks.map((b) => ({ label: `${b.name}（${fmt(b.t)}）`, icon: "trash", run: () => saveMarks(o.id, marks.filter((x) => x.id !== b.id)) })),
      "-",
      { label: "すべて削除", icon: "trash", run: () => saveMarks(o.id, []) },
    ]);
  }
  /** The bookmarks of the selected media: play or seek to a moment, add one there, name, jump to and remove each. */
  function bookmarkDialog() {
    const o = current();
    if (!o) return;
    const url = urlOf(o);
    if (!url) { app.toast("このブラウザにメディアのデータがありません"); return; }
    const el = h(o.kind === "video" ? "video" : "audio", { src: url, preload: "auto", class: "trim-media", controls: true, playsinline: true });
    let marks = (o.bookmarks || []).map((b) => ({ ...b }));
    const listEl = h("ol", { class: "bm-list" });
    const draw = () => {
      marks.sort((a, b) => a.t - b.t);
      listEl.replaceChildren(...(marks.length ? marks.map((b) => h("li", { class: "bm-row", "data-mark": b.id },
        h("button", { type: "button", class: "btn btn-sm bm-go", title: "この位置へ移動して再生", onclick: () => { el.currentTime = b.t; el.play().catch(() => {}); } }, `▶ ${fmt(b.t)}`),
        h("input", { type: "text", class: "bm-name", value: b.name || "", placeholder: "名前", "aria-label": "ブックマークの名前", maxlength: 40, oninput: (event) => { b.name = event.target.value; } }),
        h("button", { type: "button", class: "btn btn-sm btn-ghost bm-del", title: "このブックマークを削除", onclick: () => { marks = marks.filter((x) => x !== b); draw(); } }, "削除")))
        : [h("li", { class: "hint bm-empty" }, "まだブックマークがありません。再生かシークで場面を選んで「現在の位置に追加」を押します。")]));
    };
    const add = () => {
      const t = Math.round((el.currentTime || 0) * 100) / 100;
      if (marks.some((b) => Math.abs(b.t - t) < 0.05)) { app.toast("その位置にはもうブックマークがあります"); return; }
      marks.push({ id: markId(), t, name: `ブックマーク ${marks.length + 1}` });
      draw();
    };
    const dialog = h("dialog", { class: "bm-dialog", "aria-label": "ブックマーク" },
      h("div", { class: "dialog-head" }, h("h3", {}, "ブックマーク"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body" }, el,
        h("p", { class: "hint" }, "発表では、オーディオのバーとビデオの上に印が出て、クリックでその位置へ移ります（Alt+End／Alt+Home で次・前）。アニメーションを「トリガー → ブックマーク時」にすると、その場面で動きます。"),
        listEl),
      h("div", { class: "dialog-foot" },
        h("button", { type: "button", class: "btn bm-add", onclick: add }, "＋ 現在の位置に追加"),
        h("span", { style: { flex: "1" } }),
        h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary bm-ok", onclick: () => { saveMarks(o.id, marks); dialog.close(); } }, "OK")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => { el.pause(); dialog.remove(); });
    draw();
    dialog.showModal();
  }

  // ---------------------------------------------------------------- how it starts (開始)

  const mediaPlays = (o) => (app.slide()?.timeline || []).filter((e) => e.cls === "media" && e.fx === "play" && e.el === o.id && !e.trigger);
  function startOf(o) {
    if (!o) return null;
    if (o.autoplay) return "auto";
    return mediaPlays(o).length ? "sequence" : "click";
  }
  /** 自動 plays with the slide, 一連のクリック動作 adds a 再生 to the click order, クリック時 waits for a click on it. */
  function setStart(value) {
    const o = current();
    if (!o) return;
    const list = editor.objects().map((x) => (x.id === o.id ? { ...x, autoplay: value === "auto" } : x));
    let timeline = (app.slide()?.timeline || []).filter((e) => !(e.cls === "media" && e.fx === "play" && e.el === o.id && !e.trigger));
    if (value === "sequence") timeline = [...timeline, { id: `a${Math.random().toString(36).slice(2, 9)}`, el: o.id, cls: "media", fx: "play", start: "click", dur: 1, delay: 0 }];
    editor.commit(list, { timeline });
  }

  // ---------------------------------------------------------------- the 再生 tab

  function check(label, title, get, set, { enabled = null } = {}) {
    const input = h("input", { type: "checkbox", onchange: (event) => set(event.target.checked) });
    const el = h("label", { class: "rb-field an-field an-check", title }, input, h("span", {}, label));
    updater(() => { const v = get(); input.checked = Boolean(v); input.disabled = v == null || (enabled ? !enabled() : false); el.classList.toggle("off", input.disabled); });
    return el;
  }
  function seconds(label, key, title) {
    const input = h("input", { type: "number", class: "rb-num an-sec", min: 0, max: 60, step: 0.25, title, "aria-label": title,
      onchange: (event) => { const v = parseFloat(event.target.value); if (Number.isFinite(v)) setMedia({ [key]: v > 0 ? Math.min(60, v) : undefined }); },
      onkeydown: (event) => { if (event.key === "Enter") event.target.blur(); } });
    updater(() => { const o = current(); input.disabled = !o; if (document.activeElement !== input) input.value = o ? String(o[key] || 0) : ""; });
    return h("label", { class: "rb-field an-field" }, h("span", {}, label), input);
  }
  function volumeMenu() {
    const o = current();
    const v = o?.volume ?? 1;
    return menu(Object.entries(E.VOLUMES).map(([value, label]) => ({ label, on: Math.abs(Number(value) - v) < 0.01, run: () => setMedia({ volume: Number(value) === 1 ? undefined : Number(value) }) })));
  }
  function playbackTab() {
    const isAudio = () => current()?.kind === "audio";
    const start = h("select", { class: "rb-font an-select", title: "開始", "aria-label": "開始", onchange: (event) => setStart(event.target.value) }, Object.entries(STARTS).map(([v, l]) => h("option", { value: v }, l)));
    updater(() => { const o = current(); start.disabled = !o; if (o) start.value = startOf(o); });
    const preBtn = btn("play", "再生", "この画面で再生して確かめる（トリミング・フェード・音量のとおり）。もう一度で停止", () => preview(), { big: true, enabled: () => Boolean(current()), pressed: () => Boolean(player && player.id === current()?.id) });
    return [
      group("プレビュー", preBtn),
      group("ブックマーク",
        btn("pin", "ブックマーク|の追加", "再生中ならその位置に、止まっていればブックマークの画面で場面を選んで付ける（発表で移動でき、アニメーションのトリガーにも使える）", () => addBookmark(), { big: true, enabled: () => Boolean(current()) }),
        col(btn("bullet", "ブックマーク…", "ブックマークの一覧：名前・移動・削除", () => bookmarkDialog(), { enabled: () => Boolean(current()) }),
          drop("trash", "ブックマークの削除", "ブックマークを外す（そのブックマークで始まるアニメーションはクリックの順に戻る）", () => removeMenu(), { enabled: () => Boolean(current()?.bookmarks?.length) }))),
      group("編集",
        btn("crop", "トリミング", "再生する範囲（開始と終了）を決める", () => trimDialog(), { big: true, enabled: () => Boolean(current()) }),
        col(seconds("フェードイン", "fadeIn", "フェードインの時間（秒）"), seconds("フェードアウト", "fadeOut", "フェードアウトの時間（秒）"))),
      group("オプション",
        drop("audio", "音量", "再生する音量", () => volumeMenu(), { big: true, enabled: () => Boolean(current()) }),
        col(h("label", { class: "rb-field an-field" }, h("span", {}, "開始"), start),
          check("スライド切り替え後も再生", "次のスライドへ進んでも鳴らし続ける（資料の最後まで）", () => (current() ? (isAudio() ? Boolean(current().across) : null) : null), (on) => setMedia((o) => (o.kind === "audio" ? { across: on ? 999 : undefined } : null))),
          check("停止するまで繰り返す", "最後まで再生したら最初に戻って繰り返す", () => (current() ? current().loop : null), (on) => setMedia({ loop: on }))),
        col(isAudioCheck("再生中のアイコンを隠す", "発表中はスピーカーのアイコンを表示しない（自動で再生するとき向け）", "hideIcon"),
          videoCheck("全画面再生", "再生を始めると全画面に広げる", "fullscreen"),
          videoCheck("再生中のみ表示", "再生していないときは見せない", "hideIdle"),
          check("再生が終了したら巻き戻す", "再生し終わったら最初の位置に戻す", () => (current() ? Boolean(current().rewind) : null), (on) => setMedia({ rewind: on || undefined })))),
      group("表紙画像",
        drop("image", "表紙画像", "再生する前に見せる画像（ビデオの1コマ・画像ファイル）", () => menu([
          { label: "ビデオの1コマから…", icon: "video", run: () => posterFromFrame() },
          { label: "ファイルから画像…", icon: "image", run: () => posterFromFile() },
          { label: "リセット（表紙画像なし）", icon: "reset", disabled: !current()?.poster, run: () => setMedia((o) => (o.kind === "video" ? { poster: undefined } : null)) },
        ]), { big: true, enabled: () => current()?.kind === "video" && !E.youtubeId(current().src) })),
      group("キャプション",
        btn("caption", "キャプション|の挿入", "ビデオに字幕ファイル（WebVTT .vtt・SubRip .srt）を付ける：発表で動画の下に字幕が出ます", () => insertCaptions(), { big: true, enabled: () => Boolean(current()) && !isAudio() }),
        btn("trash", "キャプション|の削除", "このビデオの字幕を外す", () => setMedia((o) => (o.kind === "video" ? { captions: undefined, captionLang: undefined } : null)), { big: true, enabled: () => Boolean(current()?.captions) })),
      group("オーディオ スタイル",
        btn("clear", "スタイル|なし", "再生の設定を既定に戻す（クリック時に再生・繰り返さない）", () => resetStyle(), { big: true, enabled: () => Boolean(current()) }),
        btn("audio", "バックグラウンド|で再生", "スライドショーの間ずっと流す：自動で開始・スライド切り替え後も再生・繰り返す・アイコンを隠す", () => background(), { big: true, enabled: () => isAudio() })),
    ];
    function isAudioCheck(label, title, key) {
      return check(label, title, () => (current() && isAudio() ? Boolean(current()[key]) : null), (on) => setMedia((o) => (o.kind === "audio" ? { [key]: on || undefined } : null)));
    }
    function videoCheck(label, title, key) {
      return check(label, title, () => (current() && !isAudio() ? Boolean(current()[key]) : null), (on) => setMedia((o) => (o.kind === "video" ? { [key]: on || undefined } : null)));
    }
  }
  async function insertCaptions() {
    const o = current();
    if (!o || o.kind !== "video") return;
    const [file] = await app.pickFiles(".vtt,.srt,text/vtt", false);
    if (!file) return;
    const vtt = toVtt(await file.text());
    if (vtt.length > 300_000 || !/-->/.test(vtt)) { app.toast("字幕ファイルを読めませんでした（WebVTT か SubRip の形式で、300KBまで）"); return; }
    const lang = /[ぁ-んァ-ヶ一-龠]/.test(vtt) ? "ja" : "en";
    setMedia((x) => (x.id === o.id ? { captions: vtt, captionLang: lang } : null));
    app.toast("キャプションを付けました（発表とHTML出力で動画の下に出ます）");
  }
  function resetStyle() {
    const o = current();
    if (!o) return;
    const list = editor.objects().map((x) => (x.id === o.id ? { ...x, autoplay: false, loop: false, across: undefined, hideIcon: undefined, rewind: undefined, fullscreen: undefined, hideIdle: undefined } : x));
    editor.commit(list, { timeline: (app.slide()?.timeline || []).filter((e) => !(e.cls === "media" && e.fx === "play" && e.el === o.id && !e.trigger)) });
  }
  function background() {
    const o = current();
    if (!o || o.kind !== "audio") return;
    const list = editor.objects().map((x) => (x.id === o.id ? { ...x, autoplay: true, loop: true, across: 999, hideIcon: true } : x));
    editor.commit(list, { timeline: (app.slide()?.timeline || []).filter((e) => !(e.cls === "media" && e.fx === "play" && e.el === o.id && !e.trigger)) });
    app.toast("バックグラウンドで再生：発表を始めると自動で流れ、最後のスライドまで繰り返します");
  }

  /** The 挿入 tab's オーディオ menu. */
  const audioMenu = () => menu([
    { label: "このデバイスのオーディオ…", icon: "audio", run: insertAudioFile },
    { label: "オーディオの録音…", icon: "mic", run: () => recorder("audio") },
    { label: "URLから…", icon: "link", run: insertAudioUrl },
  ]);

  return { playbackTab, audioMenu, recorder, has, preview, stopPreview, trimDialog, insertAudioFile };
}
