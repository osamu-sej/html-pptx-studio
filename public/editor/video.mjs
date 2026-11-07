// ファイル → エクスポート → ビデオの作成 (PowerPoint's Export → Create a Video): the show plays by itself from the
// first slide to the last — each slide for its saved timing (リハーサル・記録) or a number of seconds, with its
// animations and narration — while the browser records this tab; the recording is saved as a video file (WebM).

/** How long the show plays: [{ index, ms }] for the slides shown, and the total. */
export function videoPlan(order, slides, { seconds = 5, useTimings = true } = {}) {
  const plan = order.map((index) => ({ index, ms: Math.round(1000 * (useTimings && Number.isFinite(Number(slides[index]?.advance)) && slides[index]?.advance != null ? Number(slides[index].advance) : seconds)) }));
  return { plan, total: plan.reduce((sum, p) => sum + p.ms, 0) };
}

export function createVideoExport(app) {
  const { h } = app;
  function openDialog() {
    const deck = app.deck();
    if (!deck) return;
    const secs = h("input", { type: "number", min: "1", max: "60", step: "0.5", value: "5", class: "sh-num" });
    const timings = h("input", { type: "checkbox", checked: true });
    const narration = h("input", { type: "checkbox", checked: true });
    const estimate = h("p", { class: "hint vx-estimate" });
    const update = () => {
      const { total } = videoPlan(app.showOrder(), deck.slides, { seconds: Number(secs.value) || 5, useTimings: timings.checked });
      const s = Math.round(total / 1000);
      estimate.textContent = `動画の長さ：約${Math.floor(s / 60)}分${s % 60}秒（アニメーションの時間で少し延びます）`;
    };
    secs.addEventListener("input", update);
    timings.addEventListener("change", update);
    const dialog = h("dialog", { class: "video-dialog", "aria-label": "ビデオの作成" },
      h("div", { class: "dialog-head" }, h("h3", {}, "ビデオの作成"), h("button", { class: "btn btn-ghost btn-icon", type: "button", "aria-label": "閉じる", onclick: () => dialog.close() }, "✕")),
      h("div", { class: "dialog-body" },
        h("p", {}, "スライド ショーを最初から最後まで自動で流し、このタブを録画して動画ファイル（WebM）にします。アニメーション・画面切り替え・動画・ナレーションも入ります。"),
        h("label", { class: "sh-choice" }, timings, h("span", {}, "記録されたタイミングとナレーションを使用する")),
        h("label", { class: "sh-inline" }, "タイミングのないスライドの表示時間 ", secs, " 秒"),
        h("label", { class: "sh-choice" }, narration, h("span", {}, "ナレーションを入れる（音声も録る）")),
        estimate,
        h("p", { class: "hint" }, "「作成」を押すとブラウザが共有する画面を聞いてきます。「このタブ」を選んで共有してください。録画中は操作しないでください（Escで中止）。")),
      h("div", { class: "dialog-foot" }, h("button", { type: "button", class: "btn btn-ghost", onclick: () => dialog.close() }, "キャンセル"),
        h("button", { type: "button", class: "btn btn-primary vx-go", onclick: () => { dialog.close(); record({ seconds: Number(secs.value) || 5, useTimings: timings.checked, narration: narration.checked }); } }, "作成")));
    document.body.append(dialog);
    dialog.addEventListener("close", () => dialog.remove());
    dialog.showModal();
    update();
  }

  async function record({ seconds, useTimings, narration }) {
    let stream;
    try {
      stream = await navigator.mediaDevices.getDisplayMedia({ video: { frameRate: 30, displaySurface: "browser" }, audio: narration, preferCurrentTab: true, selfBrowserSurface: "include", surfaceSwitching: "exclude" });
    } catch (error) {
      if (error?.name !== "NotAllowedError") app.toast(`録画を始められませんでした（${error.message || error.name}）`);
      return;
    }
    const type = ["video/webm;codecs=vp9,opus", "video/webm;codecs=vp8,opus", "video/webm"].find((t) => MediaRecorder.isTypeSupported?.(t)) || "";
    const chunks = [];
    const rec = new MediaRecorder(stream, type ? { mimeType: type, videoBitsPerSecond: 6_000_000 } : undefined);
    rec.ondataavailable = (event) => { if (event.data?.size) chunks.push(event.data); };
    const stopped = new Promise((resolve) => { rec.onstop = resolve; });
    let done = false;
    const finish = async (save) => {
      if (done) return;
      done = true;
      clearInterval(watch);
      if (rec.state !== "inactive") rec.stop();
      await stopped;
      stream.getTracks().forEach((t) => t.stop());
      app.closeShow();
      if (!save || !chunks.length) { app.toast("ビデオの作成を中止しました"); return; }
      const blob = new Blob(chunks, { type: rec.mimeType || "video/webm" });
      app.download(blob, `${(app.deck().title || "スライド").replace(/[\\/:*?"<>|]/g, "_")}.webm`);
      app.toast(`ビデオを作成しました（${Math.round(blob.size / 1024 / 1024 * 10) / 10}MB）`);
    };
    stream.getVideoTracks()[0]?.addEventListener("ended", () => finish(true));
    // Play once from the first slide: by itself (a kiosk that stops at the end), with or without the timings.
    await app.presentForVideo({ kiosk: true, kioskSeconds: seconds, useTimings: true, ignoreTimings: !useTimings, stopAtEnd: true, narration, captions: false });
    await new Promise((r) => setTimeout(r, 600));
    rec.start(1000);
    const { plan } = videoPlan(app.showOrder(), app.deck().slides, { seconds, useTimings });
    const lastIndex = plan.at(-1)?.index;
    const lastMs = plan.at(-1)?.ms ?? seconds * 1000;
    let arrived = null;
    const watch = setInterval(() => {
      const player = app.player();
      if (!player) { finish(true); return; }
      if (player.index === lastIndex) {
        arrived ??= Date.now();
        if (Date.now() - arrived > lastMs + 800 && !app.animBusy()) finish(true);
      }
    }, 200);
    app.onShowClosed(() => finish(true));
  }
  return { openDialog, record };
}
