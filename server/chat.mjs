// Applies the AI's chat operations (numbered against the deck the AI saw) to produce a proposal.

const IMAGE_MARK = "[画像あり]";

/**
 * Deep-dive pages (slides with `drillOf`) belong to the nearest story slide above them and are not part of
 * the story (same rule as storyMap in public/engine/engine.js). Returns the owning slide of each page.
 */
export function drillParents(slides) {
  const parent = {};
  let last = -1;
  slides.forEach((slide, i) => {
    if (slide?.drillOf && last >= 0) parent[i] = last;
    else last = i;
  });
  return parent;
}

/** The AI never sees image data or user-only fields; carry them over from the slide it replaced. */
function restoreUserFields(content, original) {
  const slide = { ...content };
  for (const key of ["customImage", "image"]) {
    if (slide[key] === IMAGE_MARK) delete slide[key];
    const kept = original?.[key];
    if (!slide[key] && typeof kept === "string" && kept.startsWith("data:") && (key === "customImage" || slide.type === original.type)) slide[key] = kept;
  }
  // What the user placed by hand (a photo, a video, its position, objects and their animations, the slide's id
  // that links point at) always stays with the slide.
  if (original?.media && !slide.media) slide.media = original.media;
  // So do the slide's place in the show (hidden), its section and the review comments on it.
  for (const key of ["elements", "timeline", "sid", "transitionDur", "transitionSound", "transitionDir", "advance", "hidden", "section", "comments", "readingOrder", "background", "hideMaster", "notesRich"]) if (original?.[key] != null && slide[key] == null) slide[key] = original[key];
  if (original?.imagePlacement && slide.customImage === original.customImage && !slide.imagePlacement) slide.imagePlacement = original.imagePlacement;
  for (const key of ["animation", "photoMotion", "kinetic", "backdrop", "entrance", "emphasis", "transition", "hideTitle", "master", "sourceViewport"]) {
    if (original?.[key] && slide[key] == null && slide.type === original.type) slide[key] = original[key];
  }
  // Click-for-details text stays unless the AI rewrote it (an empty list removes it).
  if (original?.details?.length && slide.details == null && slide.type === original.type) slide.details = original.details;
  // A deep-dive page stays one.
  if (original?.drillOf && slide.drillOf === undefined) slide.drillOf = original.drillOf;
  return slide;
}

/**
 * @returns {{ slides: object[], items: {from: number|null, changed: boolean}[], deleted: number[], moved: boolean, changed: boolean }}
 * Throws an Error in words the AI can act on when the operations are not usable.
 */
export function applyChatOperations(slides, { operations = [], order } = {}) {
  const n = slides.length;
  const replaced = new Map();
  const deleted = new Set();
  const inserts = new Map();
  for (const op of operations) {
    const i = op.slide - 1;
    if (op.op === "insert") {
      if (i < 0 || i >= n - 1) throw new Error(`insert の slide は 1〜${n - 1} にしてください（最後の closing の後ろには追加できません）。`);
      if (!op.content) throw new Error("insert には content（新しいスライドの完全なJSON）が必要です。");
      if (["title", "closing"].includes(op.content.type)) throw new Error("追加するスライドの type は title と closing 以外にしてください。");
      inserts.set(i, [...(inserts.get(i) ?? []), op.content]);
      continue;
    }
    if (i < 0 || i >= n) throw new Error(`slide ${op.slide} はありません（全${n}枚）。`);
    if (op.op === "delete") {
      if (i === 0 || i === n - 1) throw new Error("表紙と最後のスライドは削除できません。");
      deleted.add(i);
      continue;
    }
    if (!op.content) throw new Error("replace には content（スライドの完全なJSON）が必要です。");
    if (i === 0 && op.content.type !== "title") throw new Error("1枚目は表紙なので type は title のままにしてください。");
    if (i === n - 1 && op.content.type !== "closing") throw new Error("最後のスライドの type は closing のままにしてください。");
    replaced.set(i, restoreUserFields(op.content, slides[i]));
  }

  // Each story slide travels with its deep-dive pages and the slides inserted after them: new deep-dive
  // pages join the slide's own, new story slides come after the whole group, and a deleted slide takes its
  // deep-dive pages with it.
  const parentOf = drillParents(slides);
  const heads = slides.map((_, i) => i).filter((i) => parentOf[i] == null);
  const headOf = (i) => parentOf[i] ?? i;
  let groups = heads.map((head) => {
    const members = slides.map((_, i) => i).filter((i) => i === head || parentOf[i] === head);
    const inserted = members.flatMap((i) => inserts.get(i) ?? []);
    return {
      from: head,
      drills: members.filter((i) => i !== head),
      newDrills: inserted.filter((slide) => slide.drillOf),
      inserted: inserted.filter((slide) => !slide.drillOf),
    };
  });
  let moved = false;
  if (order?.length) {
    const wanted = [...new Set(order.map((no) => no - 1).filter((i) => i >= 0 && i < n).map(headOf))];
    const rest = groups.filter((group) => !wanted.includes(group.from));
    const next = [...wanted.map((i) => groups.find((group) => group.from === i)), ...rest];
    if (next[0].from !== 0 || next.at(-1).from !== headOf(n - 1)) throw new Error("order では1枚目（表紙）を先頭、最後のスライドを末尾のままにしてください。");
    moved = next.some((group, index) => group.from !== groups[index].from);
    groups = next;
  }
  const items = [];
  const keep = (i) => items.push({ from: i, slide: replaced.get(i) ?? slides[i], changed: replaced.has(i) });
  for (const group of groups) {
    if (!deleted.has(group.from)) {
      keep(group.from);
      for (const i of group.drills) if (!deleted.has(i)) keep(i);
      for (const slide of group.newDrills) items.push({ from: null, slide, changed: true });
    } else for (const i of group.drills) deleted.add(i);
    for (const slide of group.inserted) items.push({ from: null, slide, changed: true });
  }
  if (items.length < 1 || items.length > 500) throw new Error(`変更後の枚数が${items.length}枚になります。1〜500枚にしてください。`);
  return {
    slides: items.map((item) => item.slide),
    items: items.map(({ from, changed }) => ({ from, changed })),
    deleted: [...deleted].sort((a, b) => a - b),
    moved,
    changed: replaced.size > 0 || inserts.size > 0 || deleted.size > 0 || moved,
  };
}

export function summarizeChange(applied) {
  const edited = applied.items.filter((item) => item.changed && item.from != null).length;
  const added = applied.items.filter((item) => item.from == null).length;
  return [
    edited ? `${edited}枚を修正` : "",
    added ? `${added}枚を追加` : "",
    applied.deleted.length ? `${applied.deleted.length}枚を削除` : "",
    applied.moved ? "順番を変更" : "",
  ].filter(Boolean).join("・");
}
