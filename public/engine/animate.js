/*
 * HTML PPTX Studio — animations people put on objects and on a layout's parts, the PowerPoint way
 * (アニメーション): entrance, emphasis, exit, motion paths and media, each started on a click, with the previous
 * one or after it, or by clicking a trigger object. Loaded after objects.js and before motion.js; exported
 * presentations carry it too.
 *
 * A slide's animations are `slide.timeline`, in order: [{ id, el, cls, fx, start, dur, delay, … }]. `el` names
 * what moves: an object's id, "grp:<group>" for a group of objects, or a part of the layout ("@title",
 * "@takeaway", "@g0" for the layout's first item). The player (motion.js) plays the layout's own click builds
 * first, then one group of animations per click. Effects run on the Web Animations API, on the wrappers
 * objects.js draws — .hs-obj-move for motion paths, .hs-obj-fx for everything else — or on a layout part itself.
 * Entrances and exits use `transform`, `opacity` and `clip-path`; emphasis uses the separate `scale` and
 * `rotate` properties (and colours), and paths `translate`, so effects on one object add up instead of
 * replacing each other, as they do in PowerPoint.
 */
(function (root) {
  "use strict";
  const E = root.SlideEngine;
  if (!E) throw new Error("engine.js must load before animate.js");
  const { W, H } = E;

  // ---------------------------------------------------------------- the catalogue

  const CLASSES = { in: "開始", em: "強調", out: "終了", path: "アニメーションの軌跡", media: "メディア" };
  const STARTS = { click: "クリック時", with: "直前の動作と同時", after: "直前の動作の後" };
  const EASES = { auto: "効果の既定", smooth: "滑らかに開始と終了", out: "滑らかに終了", in: "滑らかに開始", linear: "一定の速さ", bounce: "バウンド終了" };
  const BY = { all: "すべて同時", para: "段落ごと", word: "単語ごと", char: "文字ごと" };
  // A SmartArt's items one after another (PowerPoint's 効果のオプション → 個別).
  const BY_SMARTART = { all: "1つのオブジェクトとして", item: "1つずつ" };
  // アニメーションの後 (PowerPoint's After animation): fade back (暗くする), gone at once, or gone at the next click.
  const AFTERS = { none: "暗くしない", dim: "薄くする（暗くする）", hide: "アニメーションの後で非表示にする", hideNext: "次のアニメーションのクリック時に非表示にする" };
  const REPEATS = { 1: "なし", 2: "2回", 3: "3回", 4: "4回", 5: "5回", 10: "10回", click: "次のクリックまで", slide: "スライドの最後まで" };
  const SPEEDS = [[5000, "さらに遅く（5秒）"], [3000, "遅く（3秒）"], [2000, "普通（2秒）"], [1000, "速く（1秒）"], [500, "さらに速く（0.5秒）"]];

  const DIR8 = [["bottom", "下から"], ["left", "左から"], ["right", "右から"], ["top", "上から"], ["bottomLeft", "左下から"], ["bottomRight", "右下から"], ["topLeft", "左上から"], ["topRight", "右上から"]];
  const DIR4 = DIR8.slice(0, 4);
  const toward = (dirs) => dirs.map(([key, label]) => [key, label.replace("から", "へ")]);

  // Entrances. Each draws its way in from hidden to shown; the exit of the same name plays it backwards.
  const IN = {
    appear: { label: "アピール", out: "クリア", dur: 1, ease: "linear", kf: () => [{ opacity: 1 }, { opacity: 1 }] },
    fade: { label: "フェード", out: "フェード", dur: 500, ease: "smooth", kf: () => [{ opacity: 0 }, { opacity: 1 }] },
    flyIn: { label: "スライドイン", out: "スライドアウト", dur: 500, ease: "out", dirs: DIR8, outDirs: toward(DIR8), dir: "bottom",
      kf: (c) => { const [dx, dy] = offstage(c); return [{ transform: `translate(${px(dx)}, ${px(dy)})` }, { transform: "translate(0px, 0px)" }]; } },
    floatIn: { label: "フロートイン", out: "フロートアウト", dur: 1000, ease: "out", dirs: [["up", "上へ"], ["down", "下へ"]], outDirs: [["up", "下へ"], ["down", "上へ"]], dir: "up",
      kf: (c) => [{ opacity: 0, transform: `translate(0px, ${c.dir === "down" ? -120 : 120}px)` }, { opacity: 1, transform: "translate(0px, 0px)" }] },
    wipe: { label: "ワイプ", out: "ワイプ", dur: 500, ease: "linear", dirs: DIR4, outDirs: [["bottom", "上から"], ["left", "右から"], ["right", "左から"], ["top", "下から"]], dir: "bottom",
      kf: (c) => sampled(c, (p) => ({ clipPath: insetFor(c, wipeRegion(c.unit, c.dir, p)) })) },
    split: { label: "スプリット", out: "スプリット", dur: 500, ease: "linear", dirs: [["hOut", "横/外へ"], ["hIn", "横/内へ"], ["vOut", "縦/外へ"], ["vIn", "縦/内へ"]], outDirs: [["hOut", "横/内へ"], ["hIn", "横/外へ"], ["vOut", "縦/内へ"], ["vIn", "縦/外へ"]], dir: "vOut",
      kf: (c) => splitFrames(c) },
    shape: { label: "図形", out: "図形", dur: 1000, ease: "out", dirs: [["circleOut", "円・外へ"], ["circleIn", "円・内へ"], ["boxOut", "ボックス・外へ"], ["boxIn", "ボックス・内へ"], ["diamondOut", "ひし形・外へ"], ["diamondIn", "ひし形・内へ"], ["plusOut", "プラス・外へ"], ["plusIn", "プラス・内へ"]],
      outDirs: [["circleOut", "円・内へ"], ["circleIn", "円・外へ"], ["boxOut", "ボックス・内へ"], ["boxIn", "ボックス・外へ"], ["diamondOut", "ひし形・内へ"], ["diamondIn", "ひし形・外へ"], ["plusOut", "プラス・内へ"], ["plusIn", "プラス・外へ"]], dir: "circleOut",
      kf: (c) => shapeFrames(c) },
    wheel: { label: "ホイール", out: "ホイール", dur: 1000, ease: "linear", dirs: [["1", "スポーク 1"], ["2", "スポーク 2"], ["3", "スポーク 3"], ["4", "スポーク 4"], ["8", "スポーク 8"]], dir: "1",
      kf: (c) => wheelFrames(c) },
    randomBars: { label: "ランダムストライプ", out: "ランダムストライプ", dur: 600, ease: "linear", dirs: [["h", "横"], ["v", "縦"]], dir: "h",
      kf: (c) => barsFrames(c) },
    zoom: { label: "ズーム", out: "ズーム", dur: 500, ease: "out", dirs: [["in", "イン"], ["out", "アウト"], ["center", "スライドの中心から"]], outDirs: [["in", "イン"], ["out", "アウト"], ["center", "スライドの中心へ"]], dir: "in", origin: "center",
      kf: (c) => zoomFrames(c) },
    swivel: { label: "ターン", out: "ターン", dur: 1000, ease: "out", dirs: [["h", "横"], ["v", "縦"]], dir: "h", origin: "center",
      kf: (c) => [{ opacity: 0, transform: `perspective(1600px) rotate${c.dir === "v" ? "X" : "Y"}(-720deg)` }, { opacity: 1, offset: 0.4 }, { opacity: 1, transform: `perspective(1600px) rotate${c.dir === "v" ? "X" : "Y"}(0deg)` }] },
    bounce: { label: "バウンド", out: "バウンド", dur: 1200, ease: "linear", kf: (c) => bounceFrames(c) },
    growTurn: { label: "グローとターン", out: "シュリンクとターン", dur: 800, ease: "out", origin: "center",
      kf: () => [{ opacity: 0, transform: "scale(0) rotate(-90deg)" }, { opacity: 1, transform: "scale(1) rotate(0deg)" }] },
    stretch: { label: "ストレッチ", out: "折りたたみ", dur: 500, ease: "out", dirs: [["across", "左右"], ["left", "左から"], ["right", "右から"], ["top", "上から"], ["bottom", "下から"]], outDirs: [["across", "左右"], ["left", "左へ"], ["right", "右へ"], ["top", "上へ"], ["bottom", "下へ"]], dir: "across",
      kf: (c) => [{ transform: ["top", "bottom"].includes(c.dir) ? "scale(1, 0)" : "scale(0, 1)" }, { transform: "scale(1, 1)" }], origin: (c) => ({ across: "center", left: "left", right: "right", top: "top", bottom: "bottom" })[c.dir] },
    rise: { label: "ライズアップ", out: "シンクダウン", dur: 1000, ease: "out",
      kf: () => [{ opacity: 0, transform: "translate(0px, 140px)" }, { opacity: 1, transform: "translate(0px, -12px)", offset: 0.8 }, { opacity: 1, transform: "translate(0px, 0px)" }] },
    spinner: { label: "スピナー", out: "スピナー", dur: 800, ease: "out", origin: "center",
      kf: () => [{ opacity: 0, transform: "rotate(-180deg) scale(0.6)" }, { opacity: 1, transform: "rotate(0deg) scale(1)" }] },
    pinwheel: { label: "ピンウィール", out: "ピンウィール", dur: 1200, ease: "out", origin: "center",
      kf: () => [{ opacity: 0, transform: "rotate(-720deg) scale(0)" }, { opacity: 1, transform: "rotate(0deg) scale(1)" }] },
    expand: { label: "エクスパンド", out: "コントラクト", dur: 800, ease: "out", origin: "center",
      kf: () => [{ opacity: 0, transform: "scale(0.6, 1)" }, { opacity: 1, transform: "scale(1, 1)" }] },
    // What only HTML can do (html: true; the gallery's 「HTMLならでは」). They draw their own frames (custom).
    layoutSpin: { label: "3D回転で登場", out: "3D回転で退出", html: true, dur: 1350, ease: "out", origin: "center",
      kf: () => [{ opacity: 0, transform: "perspective(1600px) rotateY(-540deg) scale(.55)" }, { opacity: 1, offset: .4 }, { opacity: 1, transform: "perspective(1600px) rotateY(0deg) scale(1)" }] },
    layoutSwirl: { label: "渦巻きで登場", out: "渦巻きで退出", html: true, dur: 1200, ease: "out", origin: "center",
      kf: () => [{ opacity: 0, transform: "rotate(-200deg) scale(.15)", filter: "blur(10px)" }, { opacity: 1, filter: "blur(0px)", offset: .55 }, { transform: "rotate(8deg) scale(1.04)", offset: .8 }, { opacity: 1, transform: "rotate(0deg) scale(1)", filter: "blur(0px)" }] },
    layoutFlip: { label: "縦にめくれて登場", out: "縦にめくれて退出", html: true, dur: 700, ease: "out", origin: "top",
      kf: () => [{ opacity: 0, transform: "perspective(1400px) rotateX(-75deg)" }, { opacity: 1, transform: "perspective(1400px) rotateX(0deg)" }] },
    kineticZoom: { label: "ことばが迫る", html: true, noExit: true, dur: 1500, ease: "linear", custom: (c) => kineticWords(c, "zoom") },
    kineticFlip: { label: "文字がめくれる", html: true, noExit: true, dur: 1600, ease: "linear", custom: (c) => kineticWords(c, "flip") },
    typewriter: { label: "タイプライター", out: "タイプライター（消す）", html: true, dur: 1600, ease: "linear", custom: (c, exit) => typewriter(c, exit) },
    decode: { label: "デコード", html: true, noExit: true, dur: 1400, ease: "linear", custom: (c) => decode(c) },
    maskRise: { label: "マスクから立ち上がる", out: "マスクへ沈む", html: true, dur: 900, ease: "out", custom: (c, exit) => maskRise(c, exit) },
    blurIn: { label: "ぼかしから", out: "ぼかして消える", html: true, dur: 800, ease: "out", origin: "center",
      kf: () => [{ opacity: 0, filter: "blur(28px)", transform: "scale(1.05)" }, { opacity: 1, filter: "blur(0px)", transform: "scale(1)" }] },
    draw: { label: "線を描く", out: "線を消す", html: true, dur: 1500, ease: "smooth", custom: (c, exit) => drawIn(c, exit) },
    countUp: { label: "カウントアップ", html: true, noExit: true, dur: 1400, ease: "out", custom: (c) => countUp(c) },
    chartGrow: { label: "グラフが伸びる", html: true, noExit: true, dur: 1300, ease: "out", custom: (c) => chartGrow(c) },
    // 3D モデルの動き (model: true, for 3D models only): models.js turns the model by --m3d-yaw as the box flies in.
    arrive3d: { label: "到着", out: "退出", model: true, dur: 1600, ease: "out", origin: "center",
      kf: () => [{ opacity: 0, transform: "translate(0px, 160px) scale(0.7)", "--m3d-yaw": "-240" }, { opacity: 1, offset: 0.35 }, { opacity: 1, transform: "translate(0px, 0px) scale(1)", "--m3d-yaw": "0" }] },
  };

  // Emphasis: what changes for a moment (or stays changed) on something already on the slide.
  const AMOUNT_SPIN = [[90, "90°"], [180, "180°"], [360, "360°（1回転）"], [720, "720°（2回転）"]];
  const AMOUNT_GROW = [[110, "110%（少し大きく）"], [125, "125%"], [150, "150%（大きく）"], [200, "200%（特大）"], [75, "75%（小さく）"], [50, "50%"]];
  const AMOUNT_FADE = [[25, "25%"], [50, "50%"], [75, "75%"], [100, "100%（見えなくする）"]];
  const EM = {
    pulse: { label: "パルス", dur: 500, ease: "smooth", run: (c) => box(c, [{ scale: "1" }, { scale: "1.08" }, { scale: "1" }], { composite: "add" }) },
    colorPulse: { label: "カラーパルス", dur: 1000, ease: "smooth", color: "#ffc000", run: (c) => colorPulse(c) },
    teeter: { label: "シーソー", dur: 1000, ease: "smooth", run: (c) => box(c, [{ rotate: "0deg" }, { rotate: "7deg" }, { rotate: "-7deg" }, { rotate: "5deg" }, { rotate: "-5deg" }, { rotate: "0deg" }], { composite: "add" }) },
    spin: { label: "スピン", dur: 1000, ease: "smooth", dirs: [["cw", "時計回り"], ["ccw", "反時計回り"]], dir: "cw", amounts: AMOUNT_SPIN, amount: 360, lasting: true,
      run: (c) => box(c, [{ rotate: "0deg" }, { rotate: `${(c.dir === "ccw" ? -1 : 1) * c.amount}deg` }], { composite: "add" }) },
    growShrink: { label: "拡大/収縮", dur: 1000, ease: "smooth", dirs: [["both", "両方"], ["h", "横"], ["v", "縦"]], dir: "both", amounts: AMOUNT_GROW, amount: 150, lasting: true,
      run: (c) => { const k = c.amount / 100; const to = c.dir === "h" ? `${k} 1` : c.dir === "v" ? `1 ${k}` : `${k}`; return box(c, [{ scale: "1" }, { scale: to }], { composite: "add" }); } },
    transparency: { label: "透過性", dur: 500, ease: "smooth", amounts: AMOUNT_FADE, amount: 50, lasting: true, run: (c) => box(c, (el) => [{ opacity: getComputedStyle(el).opacity }, { opacity: String(1 - c.amount / 100) }]) },
    fillColor: { label: "塗りつぶしの色", dur: 1000, ease: "smooth", color: "#ffc000", lasting: true, run: (c) => paint(c, "fill", c.color) },
    lineColor: { label: "線の色", dur: 1000, ease: "smooth", color: "#c00000", lasting: true, run: (c) => paint(c, "stroke", c.color) },
    fontColor: { label: "フォントの色", dur: 1000, ease: "smooth", color: "#c00000", lasting: true, run: (c) => text(c, (el) => [{ color: getComputedStyle(el).color }, { color: c.color }]) },
    brighten: { label: "明るく", dur: 500, ease: "smooth", lasting: true, run: (c) => box(c, [{ filter: "brightness(1)" }, { filter: "brightness(1.3)" }]) },
    darken: { label: "暗く", dur: 500, ease: "smooth", lasting: true, run: (c) => box(c, [{ filter: "brightness(1)" }, { filter: "brightness(0.7)" }]) },
    desaturate: { label: "不飽和", dur: 500, ease: "smooth", lasting: true, run: (c) => box(c, [{ filter: "grayscale(0)" }, { filter: "grayscale(1)" }]) },
    boldFlash: { label: "太字フラッシュ", dur: 1000, ease: "linear", run: (c) => text(c, (el) => { const w = getComputedStyle(el).fontWeight; return [{ fontWeight: w }, { fontWeight: "700", offset: 0.1 }, { fontWeight: "700", offset: 0.9 }, { fontWeight: w }]; }) },
    blink: { label: "点滅", dur: 600, ease: "linear", run: (c) => box(c, (el) => { const o = getComputedStyle(el).opacity; return [{ opacity: o }, { opacity: "0", offset: 0.25 }, { opacity: o, offset: 0.5 }, { opacity: "0", offset: 0.75 }, { opacity: o }]; }) },
    wave: { label: "ウェーブ", dur: 1000, ease: "smooth", by: "char", run: (c) => box(c, [{ translate: "0px 0px" }, { translate: "0px -28px" }, { translate: "0px 0px" }], { composite: "add" }) },
    underline: { label: "下線", dur: 300, ease: "linear", lasting: true, run: (c) => text(c, [{ textDecorationLine: "none" }, { textDecorationLine: "underline" }]) },
    shine: { label: "光が走る", html: true, dur: 1000, ease: "smooth", run: (c) => shine(c) },
    ripple: { label: "波紋", html: true, dur: 1100, ease: "out", run: (c) => ripple(c) },
    marker: { label: "マーカーを引く", html: true, dur: 800, ease: "smooth", lasting: true, run: (c) => marker(c) },
    spotlight: { label: "スポットライト", html: true, dur: 500, ease: "smooth", lasting: true, run: (c, state) => spotlight(c, state) },
    // 3D モデルの強調: turning in place (its view stays; a whole turn ends where it began).
    turntable3d: { label: "ターンテーブル", model: true, dur: 2000, ease: "smooth", dirs: [["cw", "右回り"], ["ccw", "左回り"]], dir: "cw", amounts: AMOUNT_SPIN, amount: 360, lasting: true,
      run: (c) => box(c, [{ "--m3d-yaw": "0" }, { "--m3d-yaw": String((c.dir === "ccw" ? -1 : 1) * c.amount) }], { composite: "add" }) },
    swing3d: { label: "スイング", model: true, dur: 2000, ease: "smooth",
      run: (c) => box(c, [{ "--m3d-yaw": "0" }, { "--m3d-yaw": "32" }, { "--m3d-yaw": "-32" }, { "--m3d-yaw": "16" }, { "--m3d-yaw": "0" }], { composite: "add" }) },
    jump3d: { label: "ジャンプしてターン", model: true, dur: 1600, ease: "smooth",
      run: (c) => box(c, [{ "--m3d-yaw": "0", "--m3d-lift": "0" }, { "--m3d-lift": "0.35", offset: 0.45 }, { "--m3d-yaw": "360", "--m3d-lift": "0" }], { composite: "add" }) },
  };

  // Motion paths: points in slide pixels from where the object is, scaled to the object when it is added.
  const PATHS = {
    lineDown: { label: "直線（下へ）", pts: [[0, 0], [0, 1]] },
    lineUp: { label: "直線（上へ）", pts: [[0, 0], [0, -1]] },
    lineRight: { label: "直線（右へ）", pts: [[0, 0], [1, 0]] },
    lineLeft: { label: "直線（左へ）", pts: [[0, 0], [-1, 0]] },
    diagDown: { label: "斜め（右下へ）", pts: [[0, 0], [0.8, 0.8]] },
    diagUp: { label: "斜め（右上へ）", pts: [[0, 0], [0.8, -0.8]] },
    arcDown: { label: "円弧（下）", curve: true, pts: [[0, 0], [0.5, 0.4], [1, 0]] },
    arcUp: { label: "円弧（上）", curve: true, pts: [[0, 0], [0.5, -0.4], [1, 0]] },
    arcRight: { label: "円弧（右）", curve: true, pts: [[0, 0], [0.4, 0.5], [0, 1]] },
    arcLeft: { label: "円弧（左）", curve: true, pts: [[0, 0], [-0.4, 0.5], [0, 1]] },
    turnDown: { label: "ターン（下）", curve: true, pts: [[0, 0], [0.55, 0], [0.85, 0.15], [1, 0.45], [1, 0.9]] },
    turnUp: { label: "ターン（上）", curve: true, pts: [[0, 0], [0.55, 0], [0.85, -0.15], [1, -0.45], [1, -0.9]] },
    circle: { label: "円", curve: true, closed: true, pts: Array.from({ length: 8 }, (_, i) => { const a = (i / 8) * Math.PI * 2; return [0.5 * Math.sin(a), 0.5 - 0.5 * Math.cos(a)]; }) },
    square: { label: "正方形", closed: true, pts: [[0, 0], [0.7, 0], [0.7, 0.7], [0, 0.7]] },
    triangle: { label: "三角形", closed: true, pts: [[0, 0], [0.45, 0.78], [-0.45, 0.78]] },
    heart: { label: "ハート", curve: true, closed: true, pts: [[0, 0], [0.22, -0.2], [0.42, -0.08], [0.38, 0.18], [0, 0.55], [-0.38, 0.18], [-0.42, -0.08], [-0.22, -0.2]] },
    loop: { label: "ループ", curve: true, pts: [[0, 0], [0.4, 0], [0.62, -0.22], [0.5, -0.45], [0.32, -0.25], [0.55, 0], [1, 0]] },
    wave: { label: "波線", curve: true, pts: [[0, 0], [0.25, -0.18], [0.5, 0], [0.75, 0.18], [1, 0]] },
    zigzag: { label: "ジグザグ", pts: [[0, 0], [0.2, -0.15], [0.4, 0.15], [0.6, -0.15], [0.8, 0.15], [1, 0]] },
    bounceRight: { label: "バウンド", curve: true, pts: [[0, 0], [0.12, -0.35], [0.28, 0], [0.42, -0.2], [0.56, 0], [0.68, -0.09], [0.8, 0]] },
    custom: { label: "ユーザー設定パス", pts: [[0, 0], [1, 0]] },
  };

  const MEDIA = {
    play: { label: "再生", dur: 1 },
    pause: { label: "一時停止", dur: 1 },
    stop: { label: "停止", dur: 1 },
  };

  const FX = { in: IN, em: EM, out: IN, path: PATHS, media: MEDIA };
  const fxLabel = (cls, fx) => (cls === "out" ? IN[fx]?.out : FX[cls]?.[fx]?.label) || fx;
  const dirsOf = (cls, fx) => (cls === "out" ? IN[fx]?.outDirs || IN[fx]?.dirs : FX[cls]?.[fx]?.dirs) || null;
  const defaultDur = (cls, fx) => (cls === "path" ? 2000 : FX[cls]?.[fx]?.dur ?? 500);

  // ---------------------------------------------------------------- reading a timeline

  const num = (v, lo, hi, d) => (Number.isFinite(Number(v)) && v !== "" && v !== null ? Math.min(hi, Math.max(lo, Number(v))) : d);
  const HEX = /^#[0-9a-f]{6}$/i;
  const LAYOUT_TARGET = /^@(title|takeaway|g\d{1,2})$/;
  const isTarget = (el, objectIds, groupIds) => objectIds.has(el) || (el.startsWith("grp:") && groupIds.has(el.slice(4))) || LAYOUT_TARGET.test(el);

  /** A clean timeline: known effects only, numbers in range, targets that exist on the slide, ids unique. */
  function normalizeTimeline(list, slide) {
    const objects = Array.isArray(slide?.elements) ? slide.elements.filter((o) => o && typeof o.id === "string") : [];
    const objectIds = new Set(objects.map((o) => o.id));
    const groupIds = new Set(objects.map((o) => o.group).filter(Boolean));
    const kinds = new Map(objects.map((o) => [o.id, o.kind]));
    // ブックマーク時: a trigger "id@mark" names a bookmark of a video or a sound on the slide.
    const marks = new Map(objects.filter((o) => (o.kind === "video" || o.kind === "audio") && Array.isArray(o.bookmarks)).map((o) => [o.id, new Set(o.bookmarks.map((b) => b?.id).filter(Boolean))]));
    const seen = new Set();
    const out = [];
    for (const raw of Array.isArray(list) ? list : []) {
      if (!raw || typeof raw !== "object") continue;
      const cls = CLASSES[raw.cls] ? raw.cls : null;
      const el = typeof raw.el === "string" ? raw.el : "";
      if (!cls || !el || !isTarget(el, objectIds, groupIds)) continue;
      const fx = Object.hasOwn(FX[cls], raw.fx) ? raw.fx : null;
      if (!fx) continue;
      if (cls === "out" && IN[fx].noExit) continue;
      if (cls === "media" && !["video", "audio", "lottie"].includes(kinds.get(el))) continue;
      let id = typeof raw.id === "string" && /^[A-Za-z0-9_-]{1,32}$/.test(raw.id) ? raw.id : "";
      while (!id || seen.has(id)) id = `a${Math.random().toString(36).slice(2, 9)}`;
      seen.add(id);
      const entry = { id, el, cls, fx, start: STARTS[raw.start] ? raw.start : "click", dur: Math.round(num(raw.dur, cls === "media" ? 1 : 1, 60000, defaultDur(cls, fx))), delay: Math.round(num(raw.delay, 0, 60000, 0)) };
      const def = cls === "out" ? IN[fx] : FX[cls][fx];
      const dirs = dirsOf(cls, fx);
      if (dirs) entry.dir = dirs.some(([k]) => k === raw.dir) ? raw.dir : def.dir;
      if (def.amounts) entry.amount = num(raw.amount, 1, 3600, def.amount);
      if (def.color) entry.color = HEX.test(raw.color || "") ? raw.color.toLowerCase() : def.color;
      if (raw.repeat === "click" || raw.repeat === "slide") entry.repeat = raw.repeat;
      else if (Number(raw.repeat) > 1) entry.repeat = Math.round(num(raw.repeat, 1, 20, 1));
      if (raw.rewind === true) entry.rewind = true;
      if (AFTERS[raw.after] && raw.after !== "none" && cls !== "out" && cls !== "media") entry.after = raw.after;
      if (raw.sound && typeof raw.sound === "string" && /^[a-z]{2,16}$/.test(raw.sound)) entry.sound = raw.sound;
      if (raw.autoReverse === true && cls !== "media") entry.autoReverse = true;
      if (EASES[raw.ease] && raw.ease !== "auto") entry.ease = raw.ease;
      if (BY[raw.by] && raw.by !== "all" && ["in", "out", "em"].includes(cls) && !def.custom && !def.html && (objectIds.has(el) ? ["shape", "text"].includes(kinds.get(el)) : LAYOUT_TARGET.test(el))) entry.by = raw.by;
      if (raw.by === "item" && ["in", "out", "em"].includes(cls) && !def.custom && !def.html && kinds.get(el) === "smartart") entry.by = "item";
      if (typeof raw.trigger === "string" && /^[^@]+(@[A-Za-z0-9_-]{1,16})?$/.test(raw.trigger)) {
        const [tid, mark] = raw.trigger.split("@");
        if (objectIds.has(tid) && (!mark || marks.get(tid)?.has(mark))) entry.trigger = raw.trigger;
      }
      if (cls === "path") {
        const pts = (Array.isArray(raw.path?.pts) ? raw.path.pts : []).filter((p) => Array.isArray(p) && Number.isFinite(Number(p[0])) && Number.isFinite(Number(p[1]))).slice(0, 2000).map(([x, y]) => [Math.round(num(x, -6000, 6000, 0) * 10) / 10, Math.round(num(y, -6000, 6000, 0) * 10) / 10]);
        if (pts.length < 2) continue;
        entry.path = { pts, ...(raw.path.closed ? { closed: true } : {}), ...(raw.path.curve ? { curve: true } : {}) };
      }
      if (typeof raw.name === "string" && raw.name.trim()) entry.name = raw.name.trim().slice(0, 40);
      out.push(entry);
    }
    return out;
  }

  /** The animations in playing order: one group per click (group 0 plays as the slide arrives), and per trigger. */
  function timelinePlan(slide) {
    const list = normalizeTimeline(slide?.timeline, slide);
    const main = [[]];
    const triggers = new Map();
    for (const e of list) {
      if (e.trigger) {
        if (!triggers.has(e.trigger)) triggers.set(e.trigger, []);
        const seq = triggers.get(e.trigger);
        if (!seq.length || e.start === "click") seq.push([]);
        seq[seq.length - 1].push(e);
      } else if (e.start === "click") main.push([e]);
      else main[main.length - 1].push(e);
    }
    return { list, main: main.map(timed), triggers: new Map([...triggers].map(([k, seq]) => [k, seq.map(timed)])), clicks: main.length - 1, elements: Array.isArray(slide?.elements) ? slide.elements : [] };
  }

  /** When each animation of a group begins and ends (ms from the click). */
  function timed(group) {
    const items = [];
    let prev = null;
    for (const e of group) {
      const base = !prev ? 0 : e.start === "with" ? prev.base : prev.end;
      const begin = base + e.delay;
      const times = typeof e.repeat === "number" ? e.repeat : 1;
      const end = begin + e.dur * times * (e.autoReverse ? 2 : 1);
      const item = { e, base, begin, end };
      items.push(item);
      prev = item;
    }
    return { items, total: items.reduce((m, i) => Math.max(m, i.end), 0) };
  }

  /** True when the timeline animates the layout's items itself (the layout's own click build steps aside). */
  const timelineTakesLayout = (slide) => Array.isArray(slide?.timeline) && slide.timeline.some((e) => /^@g\d/.test(e?.el || ""));

  // ---------------------------------------------------------------- geometry helpers

  const px = (v) => `${Math.round(v * 100) / 100}px`;
  const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));

  /** How far the unit must travel to start just outside the slide. */
  function offstage(c) {
    const u = c.unit;
    const dx = /left/i.test(c.dir) ? -(u.x + u.w) - 8 : /right/i.test(c.dir) ? W - u.x + 8 : 0;
    const dy = /top/i.test(c.dir) ? -(u.y + u.h) - 8 : /bottom/i.test(c.dir) ? H - u.y + 8 : 0;
    return [dx, dy];
  }

  /** A region of the slide (x, y, w, h) as an inset() clip on this element. */
  function insetFor(c, [rx, ry, rw, rh]) {
    const s = c.self;
    const t = clamp(ry - s.y, 0, s.h);
    const l = clamp(rx - s.x, 0, s.w);
    const r = clamp(s.x + s.w - (rx + rw), 0, s.w);
    const b = clamp(s.y + s.h - (ry + rh), 0, s.h);
    return `inset(${px(t)} ${px(r)} ${px(b)} ${px(l)})`;
  }
  function wipeRegion(u, dir, p) {
    if (dir === "left") return [u.x, u.y, u.w * p, u.h];
    if (dir === "right") return [u.x + u.w * (1 - p), u.y, u.w * p, u.h];
    if (dir === "top") return [u.x, u.y, u.w, u.h * p];
    return [u.x, u.y + u.h * (1 - p), u.w, u.h * p];
  }
  /** Frames along p = 0…1: two for one element, more when several elements share one region (a group). */
  function sampled(c, fn, n = 11) {
    const k = c.multi ? n : 2;
    return Array.from({ length: k }, (_, i) => ({ offset: i / (k - 1), ...fn(i / (k - 1)) }));
  }
  // The unit's box in this element's own coordinates.
  const local = (c) => ({ x: c.unit.x - c.self.x, y: c.unit.y - c.self.y, w: c.unit.w, h: c.unit.h, cx: c.unit.x - c.self.x + c.unit.w / 2, cy: c.unit.y - c.self.y + c.unit.h / 2 });
  const poly = (pts, evenodd = false) => `polygon(${evenodd ? "evenodd, " : ""}${pts.map(([x, y]) => `${px(x)} ${px(y)}`).join(", ")})`;
  // Everything outside a shape (evenodd: the element's box with the shape cut out of it).
  const outside = (c, inner) => poly([[0, 0], [c.self.w, 0], [c.self.w, c.self.h], [0, c.self.h], [0, 0], ...inner, inner[0]], true);

  function splitFrames(c) {
    const u = local(c);
    const out = c.dir === "hOut" || c.dir === "vOut";
    const horizontal = c.dir === "hOut" || c.dir === "hIn";
    const band = (k) => (horizontal
      ? [[u.x - 2, u.cy - (k * u.h) / 2], [u.x + u.w + 2, u.cy - (k * u.h) / 2], [u.x + u.w + 2, u.cy + (k * u.h) / 2], [u.x - 2, u.cy + (k * u.h) / 2]]
      : [[u.cx - (k * u.w) / 2, u.y - 2], [u.cx + (k * u.w) / 2, u.y - 2], [u.cx + (k * u.w) / 2, u.y + u.h + 2], [u.cx - (k * u.w) / 2, u.y + u.h + 2]]);
    return [0, 1].map((p) => ({ clipPath: out ? poly(band(p)) : outside(c, band(1 - p)) }));
  }

  function shapeFrames(c) {
    const u = local(c);
    const out = /Out$/.test(c.dir);
    const kind = c.dir.replace(/(Out|In)$/, "");
    const reach = Math.hypot(u.w, u.h) / 2 + 2;
    const shapePts = (k) => {
      if (kind === "circle") return Array.from({ length: 32 }, (_, i) => { const a = (i / 32) * Math.PI * 2; return [u.cx + Math.cos(a) * reach * k, u.cy + Math.sin(a) * reach * k]; });
      if (kind === "box") return [[u.cx - (u.w / 2 + 2) * k, u.cy - (u.h / 2 + 2) * k], [u.cx + (u.w / 2 + 2) * k, u.cy - (u.h / 2 + 2) * k], [u.cx + (u.w / 2 + 2) * k, u.cy + (u.h / 2 + 2) * k], [u.cx - (u.w / 2 + 2) * k, u.cy + (u.h / 2 + 2) * k]];
      if (kind === "diamond") return [[u.cx, u.cy - (u.h + 4) * k], [u.cx + (u.w + 4) * k, u.cy], [u.cx, u.cy + (u.h + 4) * k], [u.cx - (u.w + 4) * k, u.cy]];
      // plus: a cross whose arms reach past the corners as it opens
      const a = (u.w / 2 + 2) * k * 1.6;
      const b = (u.h / 2 + 2) * k * 1.6;
      const t = Math.max(a, b) * 0.36;
      return [[u.cx - t, u.cy - b], [u.cx + t, u.cy - b], [u.cx + t, u.cy - t], [u.cx + a, u.cy - t], [u.cx + a, u.cy + t], [u.cx + t, u.cy + t], [u.cx + t, u.cy + b], [u.cx - t, u.cy + b], [u.cx - t, u.cy + t], [u.cx - a, u.cy + t], [u.cx - a, u.cy - t], [u.cx - t, u.cy - t]];
    };
    if (kind === "circle" && out) return [0, 1].map((p) => ({ clipPath: `circle(${px(reach * p)} at ${px(u.cx)} ${px(u.cy)})` }));
    return [0, 1].map((p) => ({ clipPath: out ? poly(shapePts(Math.max(p, 0.0001))) : outside(c, shapePts(1 - p)) }));
  }

  function wheelFrames(c) {
    const u = local(c);
    const spokes = Math.max(1, Number(c.dir) || 1);
    const reach = Math.hypot(u.w, u.h) / 2 + 4;
    const arcPts = 12;
    return Array.from({ length: 13 }, (_, f) => {
      const p = f / 12;
      const pts = [];
      for (let s = 0; s < spokes; s += 1) {
        const a0 = -Math.PI / 2 + (s / spokes) * Math.PI * 2;
        const sweep = (p * Math.PI * 2) / spokes;
        pts.push([u.cx, u.cy]);
        for (let i = 0; i <= arcPts; i += 1) { const a = a0 + (sweep * i) / arcPts; pts.push([u.cx + Math.cos(a) * reach, u.cy + Math.sin(a) * reach]); }
      }
      return { offset: p, clipPath: poly(pts) };
    });
  }

  // A seeded random so a slide's bars fall the same way every time (playing, going back, exporting).
  function seeded(text) {
    let x = 2166136261;
    for (const ch of String(text)) x = Math.imul(x ^ ch.charCodeAt(0), 16777619);
    return () => { x = Math.imul(x ^ (x >>> 15), 2246822507); x = Math.imul(x ^ (x >>> 13), 3266489909); x ^= x >>> 16; return (x >>> 0) / 4294967296; };
  }
  function barsFrames(c) {
    const u = local(c);
    const n = 24;
    const rand = seeded(c.seed || "bars");
    const at = Array.from({ length: n }, () => rand() * 0.75);
    const across = c.dir !== "v";
    return Array.from({ length: 13 }, (_, f) => {
      const p = f / 12;
      const pts = [];
      for (let i = 0; i < n; i += 1) {
        const k = clamp((p - at[i]) / 0.25, 0, 1);
        if (across) {
          const y0 = u.y + (u.h * i) / n;
          const y1 = y0 + (u.h / n) * k + (k >= 1 ? 0.6 : 0);
          pts.push([u.x - 2, y0], [u.x + u.w + 2, y0], [u.x + u.w + 2, y1], [u.x - 2, y1], [u.x - 2, y0]);
        } else {
          const x0 = u.x + (u.w * i) / n;
          const x1 = x0 + (u.w / n) * k + (k >= 1 ? 0.6 : 0);
          pts.push([x0, u.y - 2], [x1, u.y - 2], [x1, u.y + u.h + 2], [x0, u.y + u.h + 2], [x0, u.y - 2]);
        }
      }
      return { offset: p, clipPath: poly(pts) };
    });
  }

  function zoomFrames(c) {
    if (c.dir === "out") return [{ opacity: 0, transform: "translate(0px, 0px) scale(1.8)" }, { opacity: 1, transform: "translate(0px, 0px) scale(1)" }];
    if (c.dir === "center") {
      const dx = W / 2 - (c.unit.x + c.unit.w / 2);
      const dy = H / 2 - (c.unit.y + c.unit.h / 2);
      return [{ opacity: 0, transform: `translate(${px(dx)}, ${px(dy)}) scale(0.1)` }, { opacity: 1, transform: "translate(0px, 0px) scale(1)" }];
    }
    return [{ opacity: 0, transform: "translate(0px, 0px) scale(0.3)" }, { opacity: 1, transform: "translate(0px, 0px) scale(1)" }];
  }

  function bounceFrames(c) {
    const drop = -(c.unit.y + c.unit.h) - 8;
    const lift = Math.min(160, c.unit.h * 0.5 + 40);
    const steps = [[0, drop, "cubic-bezier(.55,0,1,.45)"], [0.38, 0, "cubic-bezier(0,.55,.45,1)"], [0.55, -lift, "cubic-bezier(.55,0,1,.45)"], [0.7, 0, "cubic-bezier(0,.55,.45,1)"], [0.81, -lift * 0.35, "cubic-bezier(.55,0,1,.45)"], [0.9, 0, "cubic-bezier(0,.55,.45,1)"], [0.95, -lift * 0.1, "cubic-bezier(.55,0,1,.45)"], [1, 0, "linear"]];
    return steps.map(([offset, y, easing]) => ({ offset, transform: `translate(0px, ${px(y)})`, easing }));
  }

  // Catmull-Rom through the points (a curve), or straight segments; closed paths come back to the start.
  function tracePath(path, samples = 240) {
    const pts = path.pts.map(([x, y]) => [Number(x), Number(y)]);
    if (path.closed && (pts[0][0] !== pts[pts.length - 1][0] || pts[0][1] !== pts[pts.length - 1][1])) pts.push([...pts[0]]);
    const dense = [];
    if (path.curve && pts.length > 2) {
      const at = (i) => (path.closed ? pts[(i + pts.length - 1) % (pts.length - 1)] : pts[clamp(i, 0, pts.length - 1)]);
      const per = Math.max(4, Math.ceil(samples / (pts.length - 1)));
      for (let i = 0; i < pts.length - 1; i += 1) {
        const p0 = path.closed && i === 0 ? pts[pts.length - 2] : at(i - 1);
        const [p1, p2] = [pts[i], pts[i + 1]];
        const p3 = path.closed && i === pts.length - 2 ? pts[1] : at(i + 2);
        for (let k = 0; k < per; k += 1) {
          const t = k / per;
          const t2 = t * t;
          const t3 = t2 * t;
          dense.push([0, 1].map((d) => 0.5 * (2 * p1[d] + (-p0[d] + p2[d]) * t + (2 * p0[d] - 5 * p1[d] + 4 * p2[d] - p3[d]) * t2 + (-p0[d] + 3 * p1[d] - 3 * p2[d] + p3[d]) * t3)));
        }
      }
      dense.push(pts[pts.length - 1]);
    } else {
      for (let i = 0; i < pts.length - 1; i += 1) for (let k = 0; k < 16; k += 1) dense.push([pts[i][0] + ((pts[i + 1][0] - pts[i][0]) * k) / 16, pts[i][1] + ((pts[i + 1][1] - pts[i][1]) * k) / 16]);
      dense.push(pts[pts.length - 1]);
    }
    return dense;
  }
  /** Points along the path evenly spaced by length (so the object moves at an even speed). */
  function pathPoints(path, n = 61) {
    const dense = tracePath(path);
    const lengths = [0];
    for (let i = 1; i < dense.length; i += 1) lengths.push(lengths[i - 1] + Math.hypot(dense[i][0] - dense[i - 1][0], dense[i][1] - dense[i - 1][1]));
    const total = lengths[lengths.length - 1] || 1;
    const out = [];
    let j = 0;
    for (let k = 0; k < n; k += 1) {
      const target = (total * k) / (n - 1);
      while (j < lengths.length - 2 && lengths[j + 1] < target) j += 1;
      const span = lengths[j + 1] - lengths[j] || 1;
      const t = clamp((target - lengths[j]) / span, 0, 1);
      out.push([dense[j][0] + (dense[j + 1][0] - dense[j][0]) * t, dense[j][1] + (dense[j + 1][1] - dense[j][1]) * t]);
    }
    return out;
  }
  /** An SVG path (d) of the motion path, from 0,0 (for the editor's overlay). */
  function pathD(path) {
    const pts = tracePath(path, 120);
    return pts.map(([x, y], i) => `${i ? "L" : "M"}${Math.round(x * 10) / 10} ${Math.round(y * 10) / 10}`).join(" ");
  }
  /** A preset motion path sized to the object it moves. */
  function pathPreset(key, o = {}) {
    const def = PATHS[key] || PATHS.lineDown;
    const size = clamp(Math.max(Number(o.w) || 0, Number(o.h) || 0) * 1.4, 240, 640);
    return { pts: def.pts.map(([x, y]) => [Math.round(x * size), Math.round(y * size)]), ...(def.closed ? { closed: true } : {}), ...(def.curve ? { curve: true } : {}) };
  }
  /** Where a path ends (the overlay's red mark). */
  const pathEnd = (path) => (path.closed ? [0, 0] : path.pts[path.pts.length - 1]);

  // ---------------------------------------------------------------- easing

  function bounceEase() {
    const out = (t) => {
      const n = 7.5625;
      const d = 2.75;
      if (t < 1 / d) return n * t * t;
      if (t < 2 / d) return n * (t -= 1.5 / d) * t + 0.75;
      if (t < 2.5 / d) return n * (t -= 2.25 / d) * t + 0.9375;
      return n * (t -= 2.625 / d) * t + 0.984375;
    };
    return `linear(${Array.from({ length: 41 }, (_, i) => `${Math.round(out(i / 40) * 1000) / 1000} ${Math.round((i / 40) * 1000) / 10}%`).join(", ")})`;
  }
  const EASE_CSS = { linear: "linear", smooth: "cubic-bezier(.45,0,.55,1)", out: "cubic-bezier(.22,.61,.36,1)", in: "cubic-bezier(.55,.06,.68,.19)" };
  let bounceCss = null;
  function easingOf(e, def) {
    let key = e.ease || def?.ease || "smooth";
    // An exit plays its entrance backwards: a gentle landing becomes a gentle take-off.
    if (!e.ease && e.cls === "out") key = key === "out" ? "in" : key === "in" ? "out" : key;
    if (key === "bounce") {
      if (bounceCss == null) bounceCss = root.CSS?.supports?.("animation-timing-function", "linear(0, 1)") ? bounceEase() : EASE_CSS.out;
      return bounceCss;
    }
    return EASE_CSS[key] || EASE_CSS.smooth;
  }

  // ---------------------------------------------------------------- finding what moves

  const cssId = (value) => (root.CSS?.escape ? root.CSS.escape(value) : String(value).replace(/["\\]/g, "\\$&"));
  const slideScale = (slideEl) => { const r = slideEl.getBoundingClientRect(); return r.width ? r.width / W : 1; };
  function measured(slideEl, el) {
    const k = slideScale(slideEl);
    const base = slideEl.getBoundingClientRect();
    const r = el.getBoundingClientRect();
    return { x: (r.left - base.left) / k, y: (r.top - base.top) / k, w: r.width / k, h: r.height / k };
  }
  const objectBox = (node) => ({ x: parseFloat(node.style.left) || 0, y: parseFloat(node.style.top) || 0, w: parseFloat(node.style.width) || 1, h: parseFloat(node.style.height) || 1 });
  function union(boxes) {
    if (!boxes.length) return { x: 0, y: 0, w: 1, h: 1 };
    const x = Math.min(...boxes.map((b) => b.x));
    const y = Math.min(...boxes.map((b) => b.y));
    return { x, y, w: Math.max(...boxes.map((b) => b.x + b.w)) - x, h: Math.max(...boxes.map((b) => b.y + b.h)) - y };
  }

  /** The parts an animation moves: [{ node (the object's box or the layout part), fx, move, text, self }] and their union. */
  function targetsOf(slideEl, plan, el) {
    const parts = [];
    const objectPart = (id) => {
      const node = slideEl.querySelector(`.hs-obj[data-el="${cssId(id)}"]`);
      if (!node) return;
      parts.push({ node, object: true, fx: node.querySelector(".hs-obj-fx") || node, move: node.querySelector(".hs-obj-move") || node, self: objectBox(node) });
    };
    if (el.startsWith("grp:")) for (const o of plan.elements) { if (o && o.group === el.slice(4) && !o.hidden) objectPart(o.id); }
    else if (el.startsWith("@")) {
      const nodes = el === "@title" ? [slideEl.querySelector('[data-field="title"]')]
        : el === "@takeaway" ? [slideEl.querySelector('[data-field="takeaway"]')]
          : [...slideEl.querySelectorAll(`[data-g="${Number(el.slice(2))}"]`)];
      for (const node of nodes.filter(Boolean)) parts.push({ node, object: false, fx: node, move: node, self: measured(slideEl, node) });
    } else objectPart(el);
    return { parts, unit: union(parts.map((p) => p.self)) };
  }

  /** Text pieces of a part for "by word / character / paragraph" (split once, kept for later animations). */
  function units(part, by) {
    if (by === "item") return [...part.node.querySelectorAll(".hs-sa-step:not(.hs-sa-fixed)")];
    const host = part.object ? part.node.querySelector(".hs-obj-tx") : part.node;
    if (!host) return [];
    if (by === "para") {
      const blocks = [...host.querySelectorAll(":scope > p, :scope > ul > li, :scope > ol > li, :scope > div")];
      return blocks.length ? blocks : [host];
    }
    const key = `hsSplit${by}`;
    if (!host.dataset[key]) {
      host.dataset[key] = "1";
      const walker = host.ownerDocument.createTreeWalker(host, 4);
      const texts = [];
      while (walker.nextNode()) if (walker.currentNode.nodeValue.trim()) texts.push(walker.currentNode);
      for (const node of texts) {
        if (node.parentElement?.classList.contains("hs-u")) continue;
        const pieces = by === "char" ? [...node.nodeValue] : wordsOf(node.nodeValue);
        const frag = host.ownerDocument.createDocumentFragment();
        for (const piece of pieces) {
          if (!piece.trim()) { frag.append(piece); continue; }
          const span = host.ownerDocument.createElement("span");
          span.className = "hs-u";
          span.dataset.u = by;
          span.textContent = piece;
          frag.append(span);
        }
        node.replaceWith(frag);
      }
    }
    return [...host.querySelectorAll(`.hs-u[data-u="${by}"]`)];
  }
  function wordsOf(text) {
    if (typeof Intl !== "undefined" && Intl.Segmenter) return [...new Intl.Segmenter("ja", { granularity: "word" }).segment(text)].map((s) => s.segment);
    return text.split(/(\s+)/);
  }

  // ---------------------------------------------------------------- effects on elements

  // Each helper returns the animations it started (Web Animations), so a group can finish them early.
  function animateEl(c, el, frames, opts = {}) {
    if (!el?.animate) return [];
    const list = typeof frames === "function" ? frames(el) : frames;
    try {
      return [el.animate(list, { duration: c.dur, easing: c.easing, iterations: c.iterations, direction: c.direction, fill: opts.fill ?? c.fill, composite: opts.composite || "replace" })];
    } catch { return []; }
  }
  const box = (c, frames, opts) => c.parts.flatMap((part) => animateEl({ ...c, self: part.self }, part.fx, frames, opts));
  function text(c, frames) {
    const out = [];
    for (const part of c.parts) {
      const host = part.object ? part.node.querySelector(".hs-obj-tx") : part.node;
      if (!host) continue;
      // Words with a colour of their own change too.
      const inner = frames && c.fx === "fontColor" ? [...host.querySelectorAll("[style*='color']")] : [];
      for (const el of [host, ...inner]) out.push(...animateEl(c, el, frames));
    }
    return out;
  }
  function paint(c, prop, color) {
    const out = [];
    for (const part of c.parts) {
      const shapes = part.object ? [...part.node.querySelectorAll(".hs-obj-geom path, .hs-obj-geom rect, .hs-obj-geom ellipse")] : [];
      const els = shapes.length ? shapes : part.object ? [] : [part.node];
      for (const el of els) {
        const cssProp = part.object ? prop : prop === "fill" ? "backgroundColor" : "borderColor";
        const current = getComputedStyle(el)[cssProp];
        if (part.object && (current === "none" || el.getAttribute(prop) === "none")) continue;
        out.push(...animateEl(c, el, [{ [cssProp]: current }, { [cssProp]: color }]));
      }
    }
    return out;
  }
  function colorPulse(c) {
    const out = [];
    for (const part of c.parts) {
      const shapes = part.object ? [...part.node.querySelectorAll(".hs-obj-geom path")].filter((el) => el.getAttribute("fill") && el.getAttribute("fill") !== "none") : [];
      if (shapes.length) for (const el of shapes) { const cur = getComputedStyle(el).fill; out.push(...animateEl(c, el, [{ fill: cur }, { fill: c.color, offset: 0.5 }, { fill: cur }])); }
      else {
        const host = part.object ? part.node.querySelector(".hs-obj-tx") || part.fx : part.node;
        const cur = getComputedStyle(host).color;
        out.push(...animateEl(c, host, [{ color: cur }, { color: c.color, offset: 0.5 }, { color: cur }]));
      }
    }
    return out;
  }


  // ---------------------------------------------------------------- what only HTML can do

  const textHost = (part) => (part.object ? part.node.querySelector(".hs-obj-tx") : part.node);
  /** The box shown while an effect draws its own frames (and hidden at the end of an exit). */
  function holdVisible(c, part, exit) {
    return animateEl({ ...c, easing: "linear", iterations: 1, direction: "normal" }, part.fx,
      exit ? [{ visibility: "visible" }, { visibility: "visible", offset: 0.999 }, { visibility: "hidden" }] : [{ visibility: "visible" }, { visibility: "visible" }], { fill: c.fill });
  }
  // Kinetic titles are filmed as text units, including the blur/perspective unavailable in the native presets.
  function kineticWords(c, style) {
    const out = [];
    for (const part of c.parts) {
      out.push(...holdVisible(c, part, false));
      const list = units(part, style === "zoom" ? "word" : "char");
      for (const [i, el] of list.entries()) {
        const start = Math.min(.48, i * (style === "zoom" ? 110 : 50) / c.dur);
        const end = Math.min(.99, start + (style === "zoom" ? 850 : 800) / c.dur);
        const from = style === "zoom"
          ? { opacity: 0, transform: "scale(2.4)", filter: "blur(10px)" }
          : { opacity: 0, transform: "perspective(500px) rotateX(-100deg)", filter: "blur(0px)" };
        const to = { opacity: 1, transform: "none", filter: "blur(0px)" };
        out.push(...animateEl({ ...c, easing: "linear" }, el, [
          { ...from, offset: 0 }, { ...from, offset: Math.max(.0001, start), easing: "cubic-bezier(.2,.8,.2,1)" },
          { ...to, offset: Math.max(start + .0001, end) }, { ...to, offset: 1 },
        ]));
      }
    }
    return out;
  }
  /**
   * An effect drawn frame by frame (letters changing, numbers counting): `clock` is the Web Animation that keeps
   * its time, so finishing a click early (or seeking) lands on the last frame and a reset puts things back.
   */
  function driven(clock, { frame, end, cancel }) {
    if (!clock) { end(); return; }
    let done = false;
    const tick = () => {
      if (done || clock.playState === "idle" || clock.playState === "finished") return;
      const t = Number(clock.currentTime) || 0;
      frame(clamp(t / (clock.effect?.getTiming?.().duration || 1), 0, 1));
      requestAnimationFrame(tick);
    };
    clock.finished.then(() => { done = true; end(); }, () => { done = true; cancel(); });
    if (typeof requestAnimationFrame === "function") requestAnimationFrame(tick);
  }
  const easeOut = (p) => 1 - (1 - p) ** 3;

  /** タイプライター: the letters come one by one behind a blinking caret (an exit deletes them from the end). */
  function typewriter(c, exit) {
    const out = [];
    for (const part of c.parts) {
      const [clock] = holdVisible(c, part, exit);
      const chars = units(part, "char");
      if (!chars.length) { out.push(clock); continue; }
      const doc = part.node.ownerDocument;
      const caret = doc.createElement("span");
      caret.className = "hs-caret";
      const show = (k) => chars.forEach((ch, i) => ch.classList.toggle("hs-anim-hide", exit ? i >= chars.length - k : i >= k));
      show(0);
      driven(clock, {
        frame: (p) => {
          const k = Math.round(p * chars.length);
          show(k);
          const at = exit ? chars[chars.length - k - 1] : chars[k - 1];
          if (at) at.after(caret); else chars[0].before(caret);
        },
        end: () => { show(chars.length); caret.remove(); },
        cancel: () => { for (const ch of chars) ch.classList.remove("hs-anim-hide"); caret.remove(); },
      });
      out.push(clock);
    }
    return out.filter(Boolean);
  }

  const GLYPHS = { wide: "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン", upper: "ABCDEFGHIJKLMNOPQRSTUVWXYZ", lower: "abcdefghijklmnopqrstuvwxyz", digit: "0123456789" };
  const glyphPool = (ch) => (/[぀-ヿ㐀-鿿！-～]/u.test(ch) ? GLYPHS.wide : /[A-Z]/.test(ch) ? GLYPHS.upper : /[a-z]/.test(ch) ? GLYPHS.lower : /\d/.test(ch) ? GLYPHS.digit : null);
  /** デコード: each letter flickers through random letters of its kind, then settles, left to right. */
  function decode(c) {
    const out = [];
    for (const part of c.parts) {
      const [clock] = holdVisible(c, part, false);
      const chars = units(part, "char");
      if (!chars.length) { out.push(clock); continue; }
      for (const ch of chars) if (ch.dataset.t == null) ch.dataset.t = ch.textContent;
      const n = chars.length;
      let swapped = 0;
      driven(clock, {
        frame: (p) => {
          const now = Date.now();
          const flicker = now - swapped > 55;
          if (flicker) swapped = now;
          chars.forEach((ch, i) => {
            const start = (i / n) * 0.62;
            const settle = start + 0.34;
            if (p < start) ch.classList.add("hs-anim-hide");
            else if (p >= settle) { ch.classList.remove("hs-anim-hide"); ch.textContent = ch.dataset.t; }
            else {
              ch.classList.remove("hs-anim-hide");
              if (flicker) ch.textContent = [...ch.dataset.t].map((x) => { const pool = glyphPool(x); return pool ? pool[Math.floor(Math.random() * pool.length)] : x; }).join("");
            }
          });
        },
        end: () => { for (const ch of chars) { ch.classList.remove("hs-anim-hide"); ch.textContent = ch.dataset.t; } },
        cancel: () => { for (const ch of chars) { ch.classList.remove("hs-anim-hide"); ch.textContent = ch.dataset.t; } },
      });
      out.push(clock);
    }
    return out.filter(Boolean);
  }

  /** マスクから立ち上がる: each word rises out of a slit under its own line (without words: the box rises). */
  function maskRise(c, exit) {
    const out = [];
    for (const part of c.parts) {
      out.push(...holdVisible(c, part, exit));
      const words = units(part, "word");
      if (!words.length) {
        let frames = [{ clipPath: "inset(100% 0 0 0)", transform: "translate(0px, 60px)" }, { clipPath: "inset(0% 0 0 0)", transform: "translate(0px, 0px)" }];
        if (exit) frames = frames.reverse();
        out.push(...animateEl(c, part.fx, frames));
        continue;
      }
      const n = words.length;
      const each = c.dur * (n > 1 ? 0.62 : 1);
      const gap = n > 1 ? (c.dur - each) / (n - 1) : 0;
      words.forEach((word, i) => {
        let frames = [{ transform: "translate(0px, 105%)", clipPath: "inset(0 0 100% 0)" }, { transform: "translate(0px, 0%)", clipPath: "inset(0 0 0% 0)" }];
        if (exit) frames = frames.reverse();
        const order = exit ? n - 1 - i : i;
        try { out.push(word.animate(frames, { duration: each, delay: order * gap, easing: c.easing, fill: exit ? "forwards" : "both" })); } catch { /* detached */ }
      });
    }
    return out;
  }

  /** 線を描く: outlines and lines are drawn as with a pen, then the fill and the words come in. */
  function drawIn(c, exit) {
    const out = [];
    const dir = (frames) => (exit ? [...frames].reverse().map((f) => { const g = { ...f }; if (g.offset != null) g.offset = 1 - g.offset; return g; }) : frames);
    for (const part of c.parts) {
      out.push(...holdVisible(c, part, exit));
      const node = part.node;
      if (node.querySelector?.(".hs-ochart, .hs-chart")) { out.push(...chartGrow({ ...c, parts: [part] }, { skipHold: true })); continue; }
      const stroked = (el) => el.getAttribute("stroke") && el.getAttribute("stroke") !== "none";
      // Lines and icons are all pen; a shape's outline is pen when it has one.
      const strokes = [...node.querySelectorAll(".hs-obj-geom path, .hs-obj-line path")].filter(stroked)
        .concat([...node.querySelectorAll(".hs-obj-icon :is(path, line, circle, rect, polyline, polygon, ellipse)")])
        .filter((el) => typeof el.getTotalLength === "function");
      const fills = [...node.querySelectorAll(".hs-obj-geom path")].filter((el) => el.getAttribute("fill") && el.getAttribute("fill") !== "none");
      const words = node.querySelector(".hs-obj-text");
      // A shape with a fill but no outline gets a pen line of its own, which fades as the fill comes in.
      const temp = [];
      if (!strokes.length && fills.length) {
        for (const el of fills) {
          if (typeof el.getTotalLength !== "function") continue;
          const pen = el.cloneNode(false);
          pen.setAttribute("fill", "none");
          pen.setAttribute("stroke", getComputedStyle(node).getPropertyValue("--accent").trim() || "#4472c4");
          pen.setAttribute("stroke-width", "4");
          pen.removeAttribute("stroke-dasharray");
          pen.classList.add("hs-pen");
          el.after(pen);
          temp.push(pen);
        }
      }
      const pens = strokes.length ? strokes : temp;
      if (!pens.length) {
        // Nothing to draw (a picture, a table): it is uncovered from the left.
        out.push(...animateEl(c, part.fx, dir([{ clipPath: "inset(0 100% 0 0)" }, { clipPath: "inset(0 0% 0 0)" }])));
        continue;
      }
      // Ink (描画) is drawn again stroke by stroke, in the order it was written (PowerPoint's 描画で再生).
      const ink = node.dataset?.kind === "ink";
      const total = pens.reduce((sum, el) => { try { return sum + el.getTotalLength(); } catch { return sum; } }, 0) || 1;
      let done = 0;
      for (const el of pens) {
        let len = 0;
        try { len = el.getTotalLength(); } catch { len = 0; }
        if (!len) continue;
        const dash = `${len} ${len}`;
        if (ink) {
          const start = done / total;
          done += len;
          const end = done / total;
          const frames = [{ strokeDasharray: dash, strokeDashoffset: `${len}`, offset: 0 }, { strokeDasharray: dash, strokeDashoffset: `${len}`, offset: start }, { strokeDasharray: dash, strokeDashoffset: "0", offset: Math.max(end, start + 0.001) }, { strokeDasharray: dash, strokeDashoffset: "0", offset: 1 }].filter((f, i, all) => i === 0 || f.offset > all[i - 1].offset || i === all.length - 1);
          out.push(...animateEl({ ...c, easing: "linear" }, el, dir(frames)));
          continue;
        }
        const frames = temp.includes(el)
          ? [{ strokeDasharray: dash, strokeDashoffset: `${len}`, opacity: 1 }, { strokeDasharray: dash, strokeDashoffset: "0", opacity: 1, offset: 0.65 }, { strokeDasharray: dash, strokeDashoffset: "0", opacity: 0 }]
          : [{ strokeDasharray: dash, strokeDashoffset: `${len}` }, { strokeDasharray: dash, strokeDashoffset: "0", offset: 0.7 }, { strokeDasharray: dash, strokeDashoffset: "0" }];
        out.push(...animateEl(c, el, dir(frames)));
      }
      if (temp.length) out[out.length - 1]?.finished.then(() => temp.forEach((pen) => pen.remove()), () => temp.forEach((pen) => pen.remove()));
      for (const el of fills) out.push(...animateEl(c, el, dir([{ fillOpacity: 0 }, { fillOpacity: 0, offset: 0.45 }, { fillOpacity: Number(el.getAttribute("fill-opacity") ?? 1) }])));
      if (words) out.push(...animateEl(c, words, dir([{ opacity: 0 }, { opacity: 0, offset: 0.6 }, { opacity: 1 }])));
      for (const el of node.querySelectorAll(".hs-obj-geom path:not(.hs-pen)")) if (!strokes.includes(el) && !fills.includes(el)) out.push(...animateEl(c, el, dir([{ opacity: 0 }, { opacity: 0, offset: 0.6 }, { opacity: 1 }])));
    }
    return out;
  }

  const NUMBER = /[-+−▲▼]?\d{1,3}(?:,\d{3})+(?:\.\d+)?|[-+−▲▼]?\d+(?:\.\d+)?/g;
  /** The numbers in a text, each wrapped once in a span that remembers its words. */
  function numberSpans(host) {
    if (!host) return [];
    if (!host.dataset.hsCount) {
      host.dataset.hsCount = "1";
      const walker = host.ownerDocument.createTreeWalker(host, 4);
      const texts = [];
      while (walker.nextNode()) if (/\d/.test(walker.currentNode.nodeValue)) texts.push(walker.currentNode);
      for (const node of texts) {
        const frag = host.ownerDocument.createDocumentFragment();
        let last = 0;
        const value = node.nodeValue;
        for (const m of value.matchAll(NUMBER)) {
          if (m.index > last) frag.append(value.slice(last, m.index));
          const span = host.ownerDocument.createElement("span");
          span.className = "hs-cnt";
          span.dataset.count = m[0];
          span.textContent = m[0];
          frag.append(span);
          last = m.index + m[0].length;
        }
        if (last < value.length) frag.append(value.slice(last));
        node.replaceWith(frag);
      }
    }
    return [...host.querySelectorAll(".hs-cnt")];
  }
  function countText(target, p) {
    const clean = target.replace(/[,\s]/g, "").replace("−", "-").replace(/^[▲+]/, "").replace(/^▼/, "-");
    const value = parseFloat(clean);
    if (!Number.isFinite(value)) return target;
    const decimals = (clean.split(".")[1] || "").length;
    const sign = /^[-−▼▲+]/.test(target) ? target[0] : "";
    const v = Math.abs(value) * p;
    const fixed = v.toFixed(decimals);
    const [int, frac] = fixed.split(".");
    return `${sign}${target.includes(",") ? Number(int).toLocaleString("en-US") : int}${frac ? `.${frac}` : ""}`;
  }
  /** カウントアップ: every number in the words counts up from zero (the words around them stay). */
  function countUp(c) {
    const out = [];
    for (const part of c.parts) {
      const [clock] = holdVisible(c, part, false);
      const spans = numberSpans(textHost(part));
      if (!spans.length) { out.push(...animateEl(c, part.fx, [{ opacity: 0 }, { opacity: 1 }]), clock); continue; }
      const set = (p) => { for (const sp of spans) sp.textContent = p >= 1 ? sp.dataset.count : countText(sp.dataset.count, easeOut(p)); };
      set(0);
      driven(clock, { frame: set, end: () => set(1), cancel: () => set(1) });
      out.push(clock);
    }
    return out.filter(Boolean);
  }

  /** グラフが伸びる: bars grow from their axis one after another, lines are drawn, slices open, then the labels. */
  function chartGrow(c, { skipHold = false } = {}) {
    const out = [];
    for (const part of c.parts) {
      if (!skipHold) out.push(...holdVisible(c, part, false));
      const node = part.node;
      const at = (el, frames, k = 0, span = 0.7) => {
        try { out.push(el.animate(frames, { duration: c.dur * span, delay: c.dur * (1 - span) * k, easing: c.easing, fill: "both" })); } catch { /* detached */ }
      };
      const bars = [...node.querySelectorAll(".hs-obar, .hs-bar")];
      const n = Math.max(1, bars.length - 1);
      bars.forEach((bar, i) => {
        const horizontal = bar.classList.contains("h");
        bar.style.transformBox = "fill-box";
        bar.style.transformOrigin = bar.dataset.center != null ? "50% 50%" : horizontal ? (bar.dataset.neg ? "100% 50%" : "0% 50%") : (bar.dataset.neg ? "50% 0%" : "50% 100%");
        at(bar, [{ transform: horizontal ? "scale(0, 1)" : "scale(1, 0)" }, { transform: "scale(1, 1)" }], i / n);
      });
      for (const line of node.querySelectorAll(".hs-oline, .hs-draw")) {
        let len = 0;
        try { len = line.getTotalLength(); } catch { len = 0; }
        if (len) at(line, [{ strokeDasharray: `${len} ${len}`, strokeDashoffset: `${len}` }, { strokeDasharray: `${len} ${len}`, strokeDashoffset: "0" }], 0, 0.85);
      }
      for (const area of node.querySelectorAll(".hs-oarea")) at(area, [{ clipPath: "inset(0 100% 0 0)" }, { clipPath: "inset(0 0% 0 0)" }], 0, 0.85);
      const pie = node.querySelector(".hs-ochart-pie");
      const slices = [...node.querySelectorAll(".hs-oslice, .hs-arc")];
      const [cx, cy] = (pie?.dataset.c || "").split(",").map(Number);
      slices.forEach((slice, i) => {
        if (Number.isFinite(cx)) { slice.style.transformBox = "view-box"; slice.style.transformOrigin = `${cx}px ${cy}px`; }
        at(slice, [{ opacity: 0, transform: "scale(0.55) rotate(-25deg)" }, { opacity: 1, transform: "scale(1) rotate(0deg)" }], i / Math.max(1, slices.length - 1), 0.6);
      });
      for (const el of node.querySelectorAll(".hs-ochart-labels text, .hs-omark, .hs-chart .hs-val, .hs-chart .hs-err, .hs-chart .hs-trend-eq, .hs-chart .hs-slice-label, .hs-chart .hs-leader, .hs-chart circle, .hs-chart .hs-marker")) at(el, [{ opacity: 0 }, { opacity: 0, offset: 0.7 }, { opacity: 1 }], 0, 1);
      if (!bars.length && !slices.length && !node.querySelector(".hs-oline, .hs-draw, .hs-oarea")) out.push(...animateEl(c, part.fx, [{ opacity: 0 }, { opacity: 1 }]));
    }
    return out;
  }

  /** 光が走る: a band of light sweeps across. */
  function shine(c) {
    const out = [];
    for (const part of c.parts) {
      const band = part.fx.ownerDocument.createElement("span");
      band.className = "hs-fx-shine";
      part.fx.append(band);
      const [a] = animateEl({ ...c, iterations: c.iterations }, band, [{ backgroundPosition: "130% 0" }, { backgroundPosition: "-30% 0" }], { fill: "none" });
      if (!a) { band.remove(); continue; }
      a.finished.then(() => band.remove(), () => band.remove());
      out.push(a);
    }
    return out;
  }
  /** 波紋: a ring spreads out from the object and fades. */
  function ripple(c) {
    const out = [];
    for (const part of c.parts) {
      const host = part.object ? part.node : part.node;
      const ring = host.ownerDocument.createElement("span");
      ring.className = "hs-fx-ripple";
      const size = Math.max(part.self.w, part.self.h);
      Object.assign(ring.style, { width: `${size}px`, height: `${size}px`, left: `${part.self.w / 2 - size / 2}px`, top: `${part.self.h / 2 - size / 2}px` });
      host.append(ring);
      const [a] = animateEl(c, ring, [{ transform: "scale(0.7)", opacity: 0.85 }, { transform: "scale(1.9)", opacity: 0 }], { fill: "none" });
      if (!a) { ring.remove(); continue; }
      a.finished.then(() => ring.remove(), () => ring.remove());
      out.push(a);
    }
    return out;
  }
  /** マーカーを引く: a pale-blue marker is drawn under the words, line by line. */
  function marker(c) {
    const out = [];
    for (const part of c.parts) {
      const host = textHost(part);
      if (!host) continue;
      const blocks = [...host.querySelectorAll(":scope > p, :scope > ul > li, :scope > ol > li")];
      for (const block of blocks.length ? blocks : [host]) {
        let mark = block.querySelector(":scope > .hs-mk");
        if (!mark) {
          mark = block.ownerDocument.createElement("span");
          mark.className = "hs-mk";
          while (block.firstChild) mark.append(block.firstChild);
          block.append(mark);
        }
        out.push(...animateEl(c, mark, [{ backgroundSize: "0% 100%" }, { backgroundSize: "100% 100%" }]));
      }
    }
    return out;
  }
  /** スポットライト: everything else on the slide dims until the next click. */
  function spotlight(c, state) {
    const slideEl = state?.slideEl;
    if (!slideEl) return [];
    const keep = new Set(c.parts.map((p) => p.node));
    const out = [];
    const others = [...slideEl.querySelectorAll(".hs-objects > .hs-obj")].filter((n) => !keep.has(n)).map((n) => n.querySelector(".hs-obj-move") || n);
    const frame = slideEl.querySelector(":scope > .hs-frame");
    for (const el of [...others, ...(frame && !c.parts.some((p) => !p.object) ? [frame] : [])]) out.push(...animateEl({ ...c, iterations: 1, direction: "normal" }, el, [{ filter: "opacity(1)" }, { filter: "opacity(0.16)" }], { fill: "forwards" }));
    state.spots = [...(state.spots || []), ...out];
    return out;
  }
  /** At the next click the spotlight lifts. */
  function releaseSpots(state) {
    for (const a of state.spots || []) { try { a.reverse(); } catch { a.cancel(); } }
    state.spots = [];
  }

  /** Start one animation now: the entrance/exit/emphasis/path/media it describes, on everything it moves. */
  function runEntry(state, e, { instant = false } = {}) {
    const { parts, unit } = targetsOf(state.slideEl, state.plan, e.el);
    if (!parts.length) return [];
    const def = e.cls === "out" ? IN[e.fx] : FX[e.cls][e.fx];
    const repeatTimes = typeof e.repeat === "number" ? e.repeat : e.repeat ? Infinity : 1;
    const c = {
      e, fx: e.fx, dir: e.dir, amount: e.amount, color: e.color, unit, parts, multi: parts.length > 1, seed: e.id,
      dur: instant ? Math.max(1, e.dur) : Math.max(1, e.dur), easing: easingOf(e, def),
      iterations: instant ? 1 : repeatTimes * (e.autoReverse ? 2 : 1), direction: e.autoReverse ? "alternate" : "normal",
      // Entrances, exits, lasting emphasis and paths keep their end; "rewind" puts things back as they were.
      fill: e.rewind ? "none" : "forwards",
    };
    if (e.autoReverse && instant) c.iterations = 2;
    let started = [];
    if (e.cls === "media") { if (!instant) media(parts, e.fx); return []; }
    if (e.cls === "in" || e.cls === "out") started = entrance(c, def, e.cls === "out");
    else if (e.cls === "em") {
      if (!def.lasting && !e.rewind && c.fill === "forwards") c.fill = "none";
      const by = e.by || def.by;
      if (by) started = byUnits(c, by, (cc) => def.run(cc, state));
      else started = def.run(c, state);
    } else if (e.cls === "path") {
      const pts = pathPoints(e.path);
      const frames = pts.map(([x, y]) => ({ translate: `${px(x)} ${px(y)}` }));
      started = c.parts.flatMap((part) => animateEl(c, part.move, frames, { composite: "add" }));
    }
    // (What waits hidden for its entrance stays marked: the entrance's frames show it while they apply, and an
    // entrance that rewinds leaves it hidden again.)
    // サウンド: the effect's sound as it starts (the player's own, as for slide transitions).
    if (e.sound && !instant) state.slideEl.hsAnimSound?.(e.sound);
    // アニメーションの後: once the effect has played (not when it is called off, nor after the slide was reset).
    const gen = state.gen;
    if (e.after && started.length) Promise.all(started.map((a) => a.finished)).then(() => { if (state.gen === gen) afterEffect(state, e, parts); }).catch(() => {});
    else if (e.after) afterEffect(state, e, parts);
    return started;
  }

  /** アニメーションの後: faded (暗くする), hidden now, or kept to be hidden at the next click. */
  function afterEffect(state, e, parts) {
    if (!parts.every((part) => part.fx.isConnected)) return;
    if (e.after === "hideNext") { state.hideNext = [...(state.hideNext || []), ...parts]; return; }
    const frames = e.after === "dim" ? [{ opacity: 0.35 }, { opacity: 0.35 }] : [{ visibility: "hidden" }, { visibility: "hidden" }];
    for (const part of parts) { try { part.fx.animate(frames, { duration: 1, fill: "forwards" }); } catch { /* detached */ } }
  }
  /** The next click: what waited for it is hidden. */
  function hideWaiting(state) {
    for (const part of state.hideNext || []) { try { part.fx.animate([{ visibility: "hidden" }, { visibility: "hidden" }], { duration: 1, fill: "forwards" }); } catch { /* detached */ } }
    state.hideNext = [];
  }

  /** An entrance (or, played backwards, an exit) on each part — or on its words or letters. */
  function entrance(c, def, exit) {
    if (def.custom) return def.custom(c, exit);
    const frameFor = (cc) => {
      let frames = def.kf(cc).map((f) => ({ ...f }));
      // Entrances keep the element shown from their first frame on; exits hide it at their last.
      if (exit) frames = frames.reverse().map((f) => { const g = { ...f }; if (g.offset != null) g.offset = 1 - g.offset; delete g.easing; return g; });
      return frames.map((f, i) => ({ ...f, visibility: exit && i === frames.length - 1 ? "hidden" : "visible" }));
    };
    const originFor = (cc, el) => {
      const o = typeof def.origin === "function" ? def.origin(cc) : def.origin;
      if (!o) return;
      const u = local(cc);
      const at = { center: [u.cx, u.cy], left: [u.x, u.cy], right: [u.x + u.w, u.cy], top: [u.cx, u.y], bottom: [u.cx, u.y + u.h] }[o] || [u.cx, u.cy];
      el.style.transformOrigin = `${px(at[0])} ${px(at[1])}`;
    };
    const by = c.e.by;
    if (by) {
      // The box itself shows at once (its fill, its frame); the words come in one after another.
      const out = [];
      for (const part of c.parts) out.push(...animateEl({ ...c, dur: 1, iterations: 1, direction: "normal" }, part.fx, exit ? [{ visibility: "visible" }, { visibility: "visible" }] : [{ visibility: "visible" }, { visibility: "visible" }], { fill: exit ? "none" : c.fill }));
      out.push(...byUnits(c, by, (cc) => cc.parts.flatMap((part) => { originFor(cc, part.fx); return animateEl(cc, part.fx, frameFor(cc)); }), { waitHidden: !exit }));
      if (exit) {
        // When the last word has gone, the box goes too.
        const last = out[out.length - 1];
        last?.finished.then(() => { for (const part of c.parts) animateEl({ ...c, dur: 1, iterations: 1, direction: "normal" }, part.fx, [{ visibility: "hidden" }, { visibility: "hidden" }], { fill: c.fill }); }).catch(() => {});
      }
      return out;
    }
    return c.parts.flatMap((part) => {
      const cc = { ...c, self: part.self };
      originFor(cc, part.fx);
      return animateEl(cc, part.fx, frameFor(cc));
    });
  }

  /** Run an effect on each word, letter or paragraph in turn (staggered as PowerPoint's "% delay between"). */
  function byUnits(c, by, run, { waitHidden = false } = {}) {
    const out = [];
    for (const part of c.parts) {
      const pieces = units(part, by);
      if (!pieces.length) { out.push(...run({ ...c, parts: [part] })); continue; }
      // An instant effect (アピール) still comes in letter by letter, like a typewriter.
      const gap = (c.dur < 100 ? 500 : c.dur) * (by === "char" ? 0.1 : by === "word" ? 0.2 : 0.6);
      pieces.forEach((piece, i) => {
        const self = measured(part.node.closest(".hs-slide") || part.node, piece);
        const sub = { node: piece, object: false, fx: piece, move: piece, self };
        const cc = { ...c, parts: [sub], unit: self, multi: false };
        // Words still to come in wait hidden; each one's frames show it when its turn starts.
        if (waitHidden) piece.classList.add("hs-anim-hide");
        for (const a of run(cc)) { try { a.effect.updateTiming({ delay: i * gap }); } catch { /* detached */ } out.push(a); }
      });
    }
    return out;
  }

  function media(parts, fx) {
    for (const part of parts) {
      const video = part.node.querySelector("video, audio");
      const frame = part.node.querySelector("iframe");
      const lottie = part.node.querySelector(".hs-lottie-host");
      if (video && E.mediaPlay) {
        if (fx === "play") E.mediaPlay(video);
        else E.mediaPause(video, { stop: fx === "stop" });
      } else if (video) {
        if (fx === "play") video.play().catch(() => { video.muted = true; video.play().catch(() => {}); });
        else { video.pause(); if (fx === "stop") video.currentTime = 0; }
      }
      if (frame) frame.contentWindow?.postMessage(JSON.stringify({ event: "command", func: fx === "play" ? "playVideo" : fx === "pause" ? "pauseVideo" : "stopVideo", args: [] }), "*");
      if (lottie?.hsAnim) { if (fx === "play") lottie.hsAnim.play(); else if (fx === "pause") lottie.hsAnim.pause(); else lottie.hsAnim.stop(); }
      else if (lottie && fx === "play") E.mountLottie?.(part.node, { play: true, frame: 0 }).then(() => lottie.hsAnim?.play());
    }
  }

  // ---------------------------------------------------------------- playing a slide

  function stateOf(slideEl) {
    const plan = slideEl.hsTimeline;
    if (!plan) return null;
    if (!slideEl.hsAnim) slideEl.hsAnim = { slideEl, plan, timers: [], running: [], pending: [], triggerAt: new Map() };
    return slideEl.hsAnim;
  }

  /** What is hidden before anything plays: everything whose first animation is an entrance. */
  function firstEntrances(plan) {
    const firsts = new Map();
    const order = [...plan.main.flatMap((g) => g.items), ...[...plan.triggers.values()].flatMap((seq) => seq.flatMap((g) => g.items))];
    for (const { e } of order) if (!firsts.has(e.el)) firsts.set(e.el, e);
    return [...firsts.values()].filter((e) => e.cls === "in").map((e) => e.el);
  }

  /** Put everything back to the start: animations off, entrances' targets hidden. */
  function reset(state) {
    for (const t of state.timers) clearTimeout(t);
    state.timers = [];
    state.running = [];
    state.pending = [];
    state.triggerAt = new Map();
    state.hideNext = [];
    state.lastInstant = [];
    state.gen = (state.gen || 0) + 1;
    const slideEl = state.slideEl;
    for (const el of slideEl.querySelectorAll(".hs-obj-fx, .hs-obj-move, .hs-obj-tx, .hs-obj-text, .hs-obj-geom *, .hs-obj-icon *, .hs-obj-line path, .hs-ochart *, .hs-chart *, .hs-mk, .hs-frame, [data-field], [data-g], .hs-u")) for (const a of el.getAnimations?.() || []) a.cancel();
    for (const el of slideEl.querySelectorAll(".hs-pen, .hs-caret, .hs-fx-shine, .hs-fx-ripple")) el.remove();
    state.spots = [];
    for (const el of slideEl.querySelectorAll(".hs-anim-hide")) el.classList.remove("hs-anim-hide");
    for (const target of firstEntrances(state.plan)) for (const part of targetsOf(slideEl, state.plan, target).parts) part.fx.classList.add("hs-anim-hide");
  }

  /** Play a group now (timed), or apply where it ends (instant). Returns how long it takes (ms). */
  function runGroup(state, group, { instant = false, offset = 0 } = {}) {
    if (!group) return 0;
    if (instant) {
      for (const item of [...group.items].sort((a, b) => a.begin - b.begin)) for (const a of runEntry(state, item.e, { instant: true })) { try { a.finish(); } catch { a.cancel(); } }
      state.lastInstant = group.items;
      return 0;
    }
    for (const item of group.items) {
      const go = () => {
        state.pending = state.pending.filter((p) => p !== item);
        const started = runEntry(state, item.e);
        state.running.push(...started.map((a) => ({ a, e: item.e })));
      };
      state.pending.push(item);
      if (item.begin + offset <= 0) go();
      else state.timers.push(setTimeout(go, item.begin + offset));
    }
    return group.total + offset;
  }

  /** True while a group is still playing (a click then finishes it instead of moving on). */
  function animBusy(slideEl) {
    const state = slideEl?.hsAnim;
    if (!state) return false;
    return state.pending.length > 0 || state.running.some(({ a, e }) => a.playState === "running" && Number.isFinite(a.effect?.getComputedTiming?.().iterations ?? 1) && e.repeat !== "slide");
  }
  /** Finish what is playing at once (pending animations jump to their end too). */
  function animFinish(slideEl) {
    const state = slideEl?.hsAnim;
    if (!state) return;
    for (const t of state.timers) clearTimeout(t);
    state.timers = [];
    const pending = [...state.pending].sort((a, b) => a.begin - b.begin);
    state.pending = [];
    for (const { a, e } of state.running) {
      if (e.repeat === "slide") continue;
      if (e.repeat === "click") { a.cancel(); continue; }
      try { a.finish(); } catch { a.cancel(); }
    }
    state.running = state.running.filter(({ e }) => e.repeat === "slide");
    for (const item of pending) for (const a of runEntry(state, item.e, { instant: true })) { try { a.finish(); } catch { a.cancel(); } }
  }

  /**
   * Start a slide's animations at click step `step` (0: as the slide arrives). Playing from the start runs the
   * group that needs no click; arriving later (going back, a deep-dive page closing) shows each group's end.
   * Returns how long the first group takes (0 when nothing plays).
   */
  function animStart(slideEl, step = 0, { animate = true, delay = 0 } = {}) {
    const state = stateOf(slideEl);
    if (!state) return 0;
    reset(state);
    bindTriggers(state);
    const upto = Math.min(step, state.plan.clicks);
    if (!animate || upto > 0) {
      for (let g = 0; g <= upto; g += 1) { if (g > 0) settleThenHide(state); runGroup(state, state.plan.main[g], { instant: true }); }
      return 0;
    }
    return runGroup(state, state.plan.main[0], { offset: Math.max(0, delay) });
  }

  /** Click step n (1-based): finish what still plays, then play group n. Returns its length (ms). */
  function animStep(slideEl, n, { animate = true } = {}) {
    const state = stateOf(slideEl);
    if (!state) return 0;
    animFinish(slideEl);
    releaseSpots(state);
    const group = state.plan.main[n];
    if (!group) return 0;
    settleThenHide(state);
    const reduced = root.matchMedia?.("(prefers-reduced-motion: reduce)").matches;
    if (!animate || reduced) { runGroup(state, group, { instant: true }); return 0; }
    return runGroup(state, group);
  }

  /** Apply the end of click step n without playing (thumbnails, the presenter view). */
  function animSeek(slideEl, n) {
    const state = stateOf(slideEl);
    if (!state) return;
    reset(state);
    for (let g = 0; g <= Math.min(n, state.plan.clicks); g += 1) { if (g > 0) settleThenHide(state); runGroup(state, state.plan.main[g], { instant: true }); }
  }
  /**
   * Before a click's group: what the last group left to hide at the next click goes. Effects finished at once
   * report their end a moment later, so their "hide at the next click" is collected here first.
   */
  function settleThenHide(state) {
    for (const { e } of state.lastInstant || []) if (e.after === "hideNext") for (const part of targetsOf(state.slideEl, state.plan, e.el).parts) if (!(state.hideNext || []).includes(part)) state.hideNext = [...(state.hideNext || []), part];
    state.lastInstant = [];
    hideWaiting(state);
  }

  /** Stop everything (leaving a slide). */
  function animStop(slideEl) {
    const state = slideEl?.hsAnim;
    if (!state) return;
    for (const t of state.timers) clearTimeout(t);
    state.timers = [];
    state.pending = [];
    for (const { a } of state.running) { try { a.cancel(); } catch { /* gone */ } }
    state.running = [];
  }

  /** Play a trigger's next group (after its last group it starts over, as a button pressed again). */
  function fireTrigger(state, key) {
    const seq = state.plan.triggers.get(key);
    if (!seq) return false;
    const at = state.triggerAt.get(key) || 0;
    const index = at >= seq.length ? 0 : at;
    if (at >= seq.length) for (const g of seq) for (const item of g.items) for (const part of targetsOf(state.slideEl, state.plan, item.e.el).parts) for (const a of part.fx.getAnimations?.() || []) a.cancel();
    state.triggerAt.set(key, index + 1);
    runGroup(state, seq[index]);
    return true;
  }

  // A trigger object plays its own sequence, one group per click on it, and does not move the slide on; a
  // bookmark trigger ("id@mark") plays its sequence as the video or sound passes that bookmark.
  function bindTriggers(state) {
    const slideEl = state.slideEl;
    if (!state.plan.triggers.size || slideEl.dataset.hsTriggers) return;
    slideEl.dataset.hsTriggers = "1";
    slideEl.addEventListener("click", (event) => {
      const node = event.target.closest?.(".hs-obj[data-trigger]");
      if (!node || !slideEl.contains(node)) return;
      if (!state.plan.triggers.has(node.dataset.el)) return;
      event.stopPropagation();
      event.preventDefault();
      fireTrigger(state, node.dataset.el);
    });
    for (const key of state.plan.triggers.keys()) {
      const [id, mark] = key.split("@");
      if (!mark) continue;
      const media = slideEl.querySelector(`.hs-obj[data-el="${cssId(id)}"] :is(video, audio)`);
      const t = media && E.mediaMarks ? E.mediaMarks(media).find((b) => b.id === mark)?.t : null;
      if (t == null) continue;
      let last = media.currentTime;
      media.addEventListener("play", () => { last = media.currentTime - 0.001; });
      media.addEventListener("seeked", () => { last = media.currentTime; });
      media.addEventListener("timeupdate", () => {
        const now = media.currentTime;
        // Passing the moment while playing (not a jump past it) plays the sequence.
        if (!media.seeking && last < t && now >= t && now - last < 1.5) fireTrigger(state, key);
        last = now;
      });
    }
  }

  // ---------------------------------------------------------------- mounting on a rendered slide

  /**
   * Called by render(): keep the plan on the slide, count its click steps after the layout's own (data-steps
   * becomes the total, data-lsteps the layout's), mark trigger objects, and, in a presentation, hide what
   * enters later before the slide is ever painted.
   */
  function timelineMount(slideEl, slide, ctx = {}) {
    if (!Array.isArray(slide?.timeline) || !slide.timeline.length) return;
    const plan = timelinePlan(slide);
    if (!plan.list.length) return;
    slideEl.hsTimeline = plan;
    const lsteps = Number(slideEl.dataset.steps || 0);
    slideEl.dataset.lsteps = String(lsteps);
    slideEl.dataset.steps = String(lsteps + plan.clicks);
    slideEl.dataset.timeline = String(plan.list.length);
    if (!ctx.live) return;
    for (const id of plan.triggers.keys()) if (!id.includes("@")) slideEl.querySelector(`.hs-obj[data-el="${cssId(id)}"]`)?.setAttribute("data-trigger", "");
    for (const target of firstEntrances(plan)) {
      for (const part of targetsOf(slideEl, plan, target).parts) part.fx.classList.add("hs-anim-hide");
    }
  }

  /** The layout's parts that can be animated on a rendered slide: [{ el: "@title", label }]. */
  function layoutTargets(slideEl) {
    const out = [];
    if (slideEl.querySelector('[data-field="title"]')) out.push({ el: "@title", label: "タイトル" });
    if (slideEl.querySelector('[data-field="takeaway"]')) out.push({ el: "@takeaway", label: "キーメッセージ" });
    const groups = new Set([...slideEl.querySelectorAll("[data-g]")].map((el) => Number(el.dataset.g)));
    for (const g of [...groups].sort((a, b) => a - b)) out.push({ el: `@g${g}`, label: `本文の項目 ${g + 1}` });
    return out;
  }

  Object.assign(E, {
    ANIM_CLASSES: CLASSES, ANIM_STARTS: STARTS, ANIM_AFTERS: AFTERS, ANIM_EASES: EASES, ANIM_BY: BY, ANIM_BY_SMARTART: BY_SMARTART, ANIM_REPEATS: REPEATS, ANIM_SPEEDS: SPEEDS,
    ANIM_IN: IN, ANIM_EM: EM, ANIM_PATHS: PATHS, ANIM_MEDIA: MEDIA, animLabel: fxLabel, animDirs: dirsOf, animDefaultDur: defaultDur, animIsHtml: (cls, fx) => Boolean((cls === "out" ? IN[fx] : FX[cls]?.[fx])?.html),
    normalizeTimeline, timelinePlan, timelineTakesLayout, timelineMount, layoutTargets,
    animStart, animStep, animSeek, animStop, animBusy, animFinish,
    pathPreset, pathPoints, pathD, pathEnd,
  });
})(typeof window !== "undefined" ? window : globalThis);
