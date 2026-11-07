// The rulers around the slide, as PowerPoint's (表示 → ルーラー): centimetres with zero at the middle of the slide,
// the selected objects' extent shaded on both rulers, and a mark following the pointer. They stay at the stage's
// edges while the slide scrolls (zoomed in) and follow every zoom and resize.

import * as ops from "./ops.mjs";

const SIZE = 18; // px: the rulers' thickness (app.css .ruler)

export function createRulers(editor, { E, stage, body }) {
  const top = document.createElement("canvas");
  const left = document.createElement("canvas");
  top.className = "ruler ruler-h";
  left.className = "ruler ruler-v";
  top.setAttribute("aria-hidden", "true");
  left.setAttribute("aria-hidden", "true");
  stage.append(top, left);
  let on = false;
  let pointer = null;
  let frame = 0;

  const slideRect = () => editor.state.wrap?.isConnected ? editor.state.wrap.getBoundingClientRect() : null;

  /** Ticks along one ruler: `from` is where the slide's left (top) edge is on the ruler, `k` screen px per slide px. */
  function draw(canvas, length, from, k, horizontal, band, mark) {
    const dpr = window.devicePixelRatio || 1;
    const w = horizontal ? length : SIZE;
    const hgt = horizontal ? SIZE : length;
    if (canvas.width !== Math.round(w * dpr) || canvas.height !== Math.round(hgt * dpr)) { canvas.width = Math.round(w * dpr); canvas.height = Math.round(hgt * dpr); }
    canvas.style.width = `${w}px`;
    canvas.style.height = `${hgt}px`;
    const c = canvas.getContext("2d");
    c.setTransform(dpr, 0, 0, dpr, 0, 0);
    c.clearRect(0, 0, w, hgt);
    c.fillStyle = "#f7f8fa";
    c.fillRect(0, 0, w, hgt);
    const span = (horizontal ? ops.W : ops.H) * k;
    // The slide itself is white on the ruler, as in PowerPoint.
    c.fillStyle = "#ffffff";
    if (horizontal) c.fillRect(from, 0, span, hgt); else c.fillRect(0, from, w, span);
    if (band) {
      c.fillStyle = "rgba(31, 56, 100, .16)";
      const a = from + band[0] * k;
      const b = from + band[1] * k;
      if (horizontal) c.fillRect(a, 0, b - a, hgt); else c.fillRect(0, a, w, b - a);
    }
    const cm = ops.PX_PER_CM * k;
    const centre = from + span / 2;
    const labelEvery = cm >= 24 ? 1 : cm >= 12 ? 2 : 5;
    const minor = cm >= 40 ? 4 : cm >= 16 ? 2 : 1;
    c.strokeStyle = "#9aa1ad";
    c.fillStyle = "#5b6270";
    c.font = "9px system-ui, sans-serif";
    c.textAlign = "center";
    c.textBaseline = "middle";
    c.lineWidth = 1;
    const max = Math.ceil(length / cm) + 2;
    c.beginPath();
    for (let i = -max; i <= max; i += 1) {
      for (let j = 0; j < minor; j += 1) {
        const at = centre + (i + j / minor) * cm;
        if (at < 0 || at > length) continue;
        const major = j === 0;
        const tick = major ? (i % labelEvery === 0 ? 0 : 6) : minor === 4 && j === 2 ? 9 : 12;
        if (major && i % labelEvery === 0) continue; // the number stands for this tick
        const p = Math.round(at) + 0.5;
        if (horizontal) { c.moveTo(p, SIZE - (SIZE - tick)); c.lineTo(p, SIZE); } else { c.moveTo(SIZE - (SIZE - tick), p); c.lineTo(SIZE, p); }
      }
    }
    c.stroke();
    for (let i = -max; i <= max; i += labelEvery) {
      const at = centre + i * cm;
      if (at < 6 || at > length - 6) continue;
      const text = String(Math.abs(i));
      if (horizontal) c.fillText(text, at, SIZE / 2);
      else { c.save(); c.translate(SIZE / 2, at); c.rotate(-Math.PI / 2); c.fillText(text, 0, 0); c.restore(); }
    }
    if (mark != null && mark >= 0 && mark <= length) {
      c.strokeStyle = "#1f3864";
      c.beginPath();
      const p = Math.round(mark) + 0.5;
      if (horizontal) { c.moveTo(p, 0); c.lineTo(p, SIZE); } else { c.moveTo(0, p); c.lineTo(SIZE, p); }
      c.stroke();
    }
    c.strokeStyle = "#e1e4ea";
    c.beginPath();
    if (horizontal) { c.moveTo(0, SIZE - 0.5); c.lineTo(w, SIZE - 0.5); } else { c.moveTo(SIZE - 0.5, 0); c.lineTo(SIZE - 0.5, hgt); }
    c.stroke();
  }

  function render() {
    frame = 0;
    if (!on) return;
    const room = body.getBoundingClientRect();
    const host = stage.getBoundingClientRect();
    const slide = slideRect();
    top.style.left = `${room.left - host.left}px`;
    top.style.top = `${room.top - host.top}px`;
    left.style.left = `${room.left - host.left}px`;
    left.style.top = `${room.top - host.top + SIZE}px`;
    if (!slide || !slide.width) { top.hidden = true; left.hidden = true; return; }
    top.hidden = false;
    left.hidden = false;
    const k = slide.width / ops.W;
    const chosen = editor.selectedObjects().filter((o) => !o.hidden);
    const box = chosen.length ? ops.bounds(chosen, (o) => E.lineEnds(o, editor.objects())) : null;
    draw(top, room.width, slide.left - room.left, k, true, box && [box.x, box.x + box.w], pointer && pointer.x - room.left);
    draw(left, room.height - SIZE, slide.top - room.top - SIZE, k, false, box && [box.y, box.y + box.h], pointer && pointer.y - room.top - SIZE);
  }
  const schedule = () => { if (on && !frame) frame = requestAnimationFrame(render); };

  body.addEventListener("scroll", schedule, { passive: true });
  body.addEventListener("pointermove", (event) => { if (!on) return; pointer = { x: event.clientX, y: event.clientY }; schedule(); });
  body.addEventListener("pointerleave", () => { pointer = null; schedule(); });
  if ("ResizeObserver" in window) new ResizeObserver(schedule).observe(body);
  window.addEventListener("resize", schedule);
  editor.subscribe(schedule);

  function show(value) {
    on = Boolean(value);
    stage.classList.toggle("with-rulers", on);
    top.hidden = !on;
    left.hidden = !on;
    if (on) { schedule(); requestAnimationFrame(() => editor.applyZoom()); }
  }
  show(false);
  return { show, refresh: schedule, get on() { return on; } };
}
