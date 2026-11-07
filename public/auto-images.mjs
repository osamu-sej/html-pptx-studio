// Aim for roughly one picture per three slides, while keeping charts and diagrams prominent.
const PHOTO_SLOTS = new Set(["title", "section", "closing", "hero", "statement", "content", "quote", "imageText"]);
const PHOTO_PRIORITY = {
  content: 90, hero: 85, imageText: 82, quote: 78, title: 74,
  section: 70, statement: 66, closing: 45,
};

export function hasSlidePicture(slide) {
  const src = slide.media?.src;
  return Boolean(
    (src && (slide.media.kind === "image" || (!slide.media.kind && /^(?:data:image\/|asset:)|\.(png|jpe?g|webp|gif)(?:[?#]|$)/i.test(src)))) ||
    (typeof slide.customImage === "string" && slide.customImage.startsWith("data:image/")) ||
    (typeof slide.image === "string" && slide.image.startsWith("data:image/")) ||
    (slide.visualAsset && slide.visualAsset !== "none")
  );
}

export function picturePlan(deck) {
  const slides = deck.slides ?? deck.slideData ?? [];
  const target = slides.length ? Math.max(1, Math.round(slides.length / 3)) : 0;
  const present = slides.filter(hasSlidePicture).length;
  const candidates = slides.flatMap((slide, index) => {
    if (hasSlidePicture(slide)) return [];
    const hasChart = slide.type === "imageText" && slide.image && typeof slide.image === "object";
    const score = (PHOTO_PRIORITY[slide.type] ?? 20) - (hasChart ? 65 : 0);
    return [{ index, score, slotted: PHOTO_SLOTS.has(slide.type) && !hasChart }];
  }).sort((a, b) => b.score - a.score || a.index - b.index);
  return { target, present, needed: Math.max(0, target - present), candidates };
}
