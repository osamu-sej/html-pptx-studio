// 共同編集の3方向マージ: pure functions shared by the browser (public/editor/coedit.mjs) and the server
// (server/rooms.mjs). The server merges too, so a change sent before another person's change arrived here does not
// overwrite it: each change says which version of the slide (or the order) it was made on.

const json = (value) => JSON.stringify(value ?? null);

/** A short fingerprint of a text (FNV-1a, 32 bits, hex): names one version of a slide or of the slide order. */
export function textHash(text) {
  let h = 0x811c9dc5;
  const s = String(text);
  for (let i = 0; i < s.length; i += 1) {
    h ^= s.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(16).padStart(8, "0");
}

/**
 * A new slide id (s + time + 7 random characters, as the rooms accept: [A-Za-z0-9_-]{1,32}) that none of `taken` has.
 * Many slides get theirs in the same millisecond, so the random part has to tell them apart.
 */
export function newSid(taken = new Set()) {
  const random = () => (globalThis.crypto?.getRandomValues ? globalThis.crypto.getRandomValues(new Uint32Array(1))[0] : Math.floor(Math.random() * 2 ** 32));
  for (;;) {
    const sid = `s${Date.now().toString(36).slice(-6)}${random().toString(36).padStart(7, "0")}`;
    if (!taken.has(sid)) return sid;
  }
}

/**
 * Every slide with its own id: a missing one gets one, and a repeated one a new one (links keep pointing at the first
 * slide with it). In place; gives back the slides.
 */
export function uniqueSids(slides) {
  const taken = new Set(slides.map((slide) => slide?.sid).filter(Boolean));
  const seen = new Set();
  for (const slide of slides) {
    if (!slide) continue;
    if (!slide.sid || seen.has(slide.sid)) {
      const sid = newSid(taken);
      taken.add(sid);
      slide.sid = sid;
    }
    seen.add(slide.sid);
  }
  return slides;
}

/** Three-way merge of one slide: our changes on top of theirs (objects by id, everything else field by field). */
export function mergeSlide(base, mine, theirs) {
  if (json(mine) === json(base)) return theirs;
  if (json(theirs) === json(base)) return mine;
  const out = { ...theirs };
  const keys = new Set([...Object.keys(base || {}), ...Object.keys(mine || {}), ...Object.keys(theirs || {})]);
  for (const key of keys) {
    if (key === "elements") continue;
    if (json(mine?.[key]) !== json(base?.[key])) { if (mine?.[key] === undefined) delete out[key]; else out[key] = mine[key]; }
  }
  // Objects: theirs, with the ones changed here put back, the ones added here added, the ones deleted here removed.
  const b = new Map((base?.elements || []).map((o) => [o.id, o]));
  const m = new Map((mine?.elements || []).map((o) => [o.id, o]));
  const t = theirs?.elements || [];
  const merged = [];
  for (const o of t) {
    if (b.has(o.id) && !m.has(o.id)) continue; // deleted here
    const changedHere = m.has(o.id) && json(m.get(o.id)) !== json(b.get(o.id));
    merged.push(changedHere ? m.get(o.id) : o);
  }
  const theirIds = new Set(t.map((o) => o.id));
  for (const [id, o] of m) if (!b.has(id) && !theirIds.has(id)) merged.push(o); // added here
  if (merged.length) out.elements = merged; else delete out.elements;
  return out;
}

/** The order of the slides when both sides may have moved, added or removed some. */
export function mergeOrder(baseSids, mineSids, theirSids) {
  if (json(mineSids) === json(baseSids)) return theirSids;
  const base = new Set(baseSids);
  const mine = new Set(mineSids);
  const removedHere = new Set(baseSids.filter((sid) => !mine.has(sid)));
  const out = (json(theirSids) === json(baseSids) ? mineSids : theirSids).filter((sid) => !removedHere.has(sid));
  // Slides added here go after the slide they follow here.
  mineSids.forEach((sid, i) => {
    if (base.has(sid) || out.includes(sid)) return;
    const before = mineSids.slice(0, i).reverse().find((s) => out.includes(s));
    out.splice(before ? out.indexOf(before) + 1 : 0, 0, sid);
  });
  return out;
}
