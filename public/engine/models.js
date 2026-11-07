/*
 * 3D models (挿入 → 3D モデル): a GLB / glTF file or one of the built-in models, drawn with three.js (MIT) in the
 * object's box. A slide keeps a still picture of each model (its poster) for thumbnails, printing and the moment
 * before the live view is ready; on the stage and in a slide show the model is drawn live, turned to its view
 * (yaw, pitch, zoom, pan). The 3D animations (animate.js: 到着・ターンテーブル・スイング・ジャンプしてターン・退出)
 * animate CSS custom properties on the object (--m3d-yaw, --m3d-pitch, --m3d-lift); the drawing reads them every frame,
 * so 3D motion plays with the rest of the slide's timeline.
 * three.js comes from E.threeUrls: the studio serves it (/vendor/three/...), an exported file carries it inline.
 */
(function (root) {
  "use strict";
  const E = root.SlideEngine;
  if (!E) return;

  // The custom properties the 3D animations move (inherited, so the model inside the animated box reads them).
  if (root.CSS?.registerProperty) {
    for (const [name, initialValue] of [["--m3d-yaw", "0"], ["--m3d-pitch", "0"], ["--m3d-lift", "0"]]) {
      try { root.CSS.registerProperty({ name, syntax: "<number>", inherits: true, initialValue }); } catch { /* registered */ }
    }
  }

  /** The built-in models (ストック 3D モデル), each made of simple shapes in soft colours. */
  const BUILTINS = {
    cube: "立方体", sphere: "球", cylinder: "円柱", cone: "円すい", torus: "ドーナツ", store: "店舗", box: "段ボール箱", arrow: "矢印",
  };
  const COLORS = { blue: 0xc9d6ee, navy: 0x1f3864, gray: 0xb8bec8, brown: 0xd9c3a5, light: 0xeef2f8, green: 0x2e8b57 };

  E.threeUrls = E.threeUrls || { three: "/vendor/three/three.module.js", gltf: "/vendor/three/GLTFLoader.js" };
  let loading = null;
  /** three.js and its glTF loader (loaded once). */
  function loadThree() {
    if (!loading) {
      loading = Promise.all([import(E.threeUrls.three), import(E.threeUrls.gltf)]).then(([THREE, loader]) => ({ THREE, GLTFLoader: loader.GLTFLoader }));
      loading.catch(() => { loading = null; });
    }
    return loading;
  }

  function builtin(THREE, name) {
    const mat = (color, extra = {}) => new THREE.MeshStandardMaterial({ color, roughness: 0.55, metalness: 0.05, ...extra });
    const group = new THREE.Group();
    const add = (geometry, material, [x, y, z] = [0, 0, 0], rot = null) => {
      const mesh = new THREE.Mesh(geometry, material);
      mesh.position.set(x, y, z);
      if (rot) mesh.rotation.set(...rot);
      group.add(mesh);
      return mesh;
    };
    if (name === "cube") add(new THREE.BoxGeometry(1, 1, 1), mat(COLORS.blue));
    else if (name === "sphere") add(new THREE.SphereGeometry(0.6, 48, 32), mat(COLORS.blue));
    else if (name === "cylinder") add(new THREE.CylinderGeometry(0.5, 0.5, 1.1, 48), mat(COLORS.blue));
    else if (name === "cone") add(new THREE.ConeGeometry(0.55, 1.1, 48), mat(COLORS.blue));
    else if (name === "torus") add(new THREE.TorusGeometry(0.5, 0.2, 32, 64), mat(COLORS.blue), [0, 0, 0], [Math.PI / 2.6, 0, 0]);
    else if (name === "box") {
      add(new THREE.BoxGeometry(1.2, 0.8, 0.9), mat(COLORS.brown));
      add(new THREE.BoxGeometry(1.21, 0.06, 0.2), mat(0xc7ab86), [0, 0.4, 0]);
    } else if (name === "arrow") {
      add(new THREE.CylinderGeometry(0.16, 0.16, 1, 32), mat(COLORS.navy), [0, -0.25, 0]);
      add(new THREE.ConeGeometry(0.36, 0.55, 32), mat(COLORS.navy), [0, 0.5, 0]);
    } else {
      // 店舗: a small shop with a glass front, a sign band and a flat roof.
      add(new THREE.BoxGeometry(1.6, 0.8, 1), mat(COLORS.light));
      add(new THREE.BoxGeometry(1.62, 0.16, 1.02), mat(COLORS.navy), [0, 0.32, 0]);
      add(new THREE.BoxGeometry(1.64, 0.04, 1.04), mat(COLORS.gray), [0, 0.42, 0]);
      add(new THREE.BoxGeometry(1.1, 0.42, 0.02), mat(COLORS.blue, { roughness: 0.15, metalness: 0.2 }), [-0.15, -0.1, 0.51]);
      add(new THREE.BoxGeometry(0.28, 0.5, 0.02), mat(COLORS.gray), [0.58, -0.15, 0.51]);
      add(new THREE.BoxGeometry(1.9, 0.02, 1.3), mat(0xdfe3ea), [0, -0.41, 0]);
    }
    return group;
  }

  // Parsed files, once each (every view of the same file shares its geometry).
  const files = new Map();
  async function modelOf(src) {
    const { THREE, GLTFLoader } = await loadThree();
    if (src.startsWith("builtin:")) return builtin(THREE, src.slice(8));
    if (!files.has(src)) {
      const pending = new GLTFLoader().loadAsync(src).then((gltf) => gltf.scene);
      files.set(src, pending);
      pending.catch(() => files.delete(src));
    }
    return (await files.get(src)).clone(true);
  }

  const rad = (deg) => (deg * Math.PI) / 180;
  const num = (v, d = 0) => (Number.isFinite(Number(v)) ? Number(v) : d);

  /** A scene for a model: centred, sized to fit, lit softly (no shadows), in a transparent box. */
  async function stage(THREE, src) {
    const model = await modelOf(src);
    const pivot = new THREE.Group();
    const box = new THREE.Box3().setFromObject(model);
    const center = box.getCenter(new THREE.Vector3());
    const radius = box.getBoundingSphere(new THREE.Sphere()).radius || 1;
    model.position.sub(center);
    pivot.add(model);
    pivot.scale.setScalar(1 / radius);
    const scene = new THREE.Scene();
    scene.add(pivot);
    scene.add(new THREE.HemisphereLight(0xffffff, 0x9aa3b2, 2.2));
    const sun = new THREE.DirectionalLight(0xffffff, 1.7);
    sun.position.set(3, 5, 6);
    scene.add(sun);
    const camera = new THREE.PerspectiveCamera(30, 1, 0.01, 100);
    return { scene, camera, pivot };
  }

  /** Turn the model and place the camera: the view (yaw, pitch, zoom, pan) plus what the animations add. */
  function pose(THREE, s, view, { yaw = 0, pitch = 0, lift = 0 } = {}, aspect = 1) {
    s.pivot.rotation.set(rad(num(view.pitch) + pitch), rad(num(view.yaw) + yaw), rad(num(view.roll)), "XYZ");
    const fov = rad(s.camera.fov);
    // The whole model (radius 1) fits the narrower side, with a little room.
    const fit = 1.12 / Math.sin(fov / 2) / Math.min(1, aspect);
    const zoom = Math.min(4, Math.max(0.3, num(view.zoom, 1)));
    s.camera.aspect = aspect;
    s.camera.position.set(-num(view.panX) * 1.2, num(view.panY) * 1.2 - lift * 1.6, fit / zoom);
    s.camera.lookAt(-num(view.panX) * 1.2, num(view.panY) * 1.2 - lift * 1.6, 0);
    s.camera.updateProjectionMatrix();
  }

  const live = new Map();
  /**
   * Draw the models in `root` live (the stage, a slide show). Each `.hs-model` gets a canvas over its poster; it is
   * drawn again whenever its size, view or animation changes. Returns a promise for when all are drawn.
   */
  function mountModels(rootEl) {
    const boxes = [...(rootEl?.querySelectorAll?.(".hs-model[data-src]") || [])].filter((el) => !live.has(el) && el.dataset.src);
    return Promise.all(boxes.map(async (el) => {
      const entry = { stopped: false };
      live.set(el, entry);
      try {
        const { THREE } = await loadThree();
        const s = await stage(THREE, el.dataset.src);
        if (entry.stopped) return;
        const canvas = document.createElement("canvas");
        canvas.className = "hs-model-canvas";
        const renderer = new THREE.WebGLRenderer({ canvas, alpha: true, antialias: true, preserveDrawingBuffer: true });
        renderer.setClearColor(0x000000, 0);
        el.append(canvas);
        const view = JSON.parse(el.dataset.view || "{}");
        let last = "";
        const frame = () => {
          if (entry.stopped) return;
          entry.raf = requestAnimationFrame(frame);
          const w = el.clientWidth;
          const hgt = el.clientHeight;
          if (!w || !hgt) return;
          const cs = getComputedStyle(el);
          const motion = { yaw: num(cs.getPropertyValue("--m3d-yaw")), pitch: num(cs.getPropertyValue("--m3d-pitch")), lift: num(cs.getPropertyValue("--m3d-lift")) };
          const key = `${w}x${hgt}|${motion.yaw}|${motion.pitch}|${motion.lift}|${el.dataset.view}`;
          if (key === last) return;
          last = key;
          const v = el.dataset.view === JSON.stringify(view) ? view : JSON.parse(el.dataset.view || "{}");
          renderer.setPixelRatio(Math.min(2, root.devicePixelRatio || 1));
          renderer.setSize(w, hgt, false);
          pose(THREE, s, v, motion, w / hgt);
          renderer.render(s.scene, s.camera);
          el.classList.add("is-live");
        };
        Object.assign(entry, { renderer, canvas });
        frame();
      } catch (error) {
        el.classList.add("is-broken");
        el.title = `3D モデルを表示できません（${error.message || error}）`;
      }
    }));
  }

  /** Stop drawing the models in `root` (their canvases and graphics memory go). */
  function stopModels(rootEl) {
    for (const [el, entry] of live) {
      if (rootEl && !rootEl.contains(el) && el.isConnected) continue;
      entry.stopped = true;
      cancelAnimationFrame(entry.raf);
      try { entry.renderer?.dispose(); entry.renderer?.forceContextLoss(); } catch { /* gone */ }
      entry.canvas?.remove();
      el.classList.remove("is-live");
      live.delete(el);
    }
  }

  /** A still picture of a model at its view (the poster kept with the slide): a PNG data URL. */
  async function modelPoster(src, view = {}, { width = 640, height = 480 } = {}) {
    const { THREE } = await loadThree();
    const s = await stage(THREE, src);
    const canvas = document.createElement("canvas");
    const renderer = new THREE.WebGLRenderer({ canvas, alpha: true, antialias: true, preserveDrawingBuffer: true });
    try {
      renderer.setClearColor(0x000000, 0);
      renderer.setPixelRatio(1);
      renderer.setSize(width, height, false);
      pose(THREE, s, view, {}, width / height);
      renderer.render(s.scene, s.camera);
      return canvas.toDataURL("image/png");
    } finally {
      renderer.dispose();
      renderer.forceContextLoss();
    }
  }

  /** The 3D モデル ビュー presets (PowerPoint's gallery): yaw and pitch in degrees. */
  const MODEL_VIEWS = {
    front: ["正面", 0, 0], back: ["背面", 180, 0], left: ["左", -90, 0], right: ["右", 90, 0], top: ["上", 0, 80], bottom: ["下", 0, -80],
    upperLeft: ["左上前面", -35, 25], upperRight: ["右上前面", 35, 25], lowerLeft: ["左下前面", -35, -20], lowerRight: ["右下前面", 35, -20],
  };

  Object.assign(E, { MODEL_BUILTINS: BUILTINS, MODEL_VIEWS, loadThree, mountModels, stopModels, modelPoster });
})(typeof window !== "undefined" ? window : globalThis);
