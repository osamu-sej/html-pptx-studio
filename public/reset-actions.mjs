// Remove presentation behaviour from every page while retaining its visible content.
// The caller keeps one undo snapshot before replacing the deck with this result.

export const STILL_MOTION = Object.freeze({
  entrance: "none", hover: "none", numbers: false, ambient: false,
  kinetic: "none", backdrop: "none", emphasis: "none", draw: false,
});

function playableMedia(media) {
  if (!media?.src) return false;
  return ["video", "youtube", "lottie"].includes(media.kind)
    || /(?:youtu\.be\/|youtube\.com\/|\.(?:mp4|webm|mov|m4v|json)(?:[?#]|$))/i.test(media.src)
    || /^data:(?:video\/|application\/json)/i.test(media.src);
}

export function resetDeckActions(source) {
  const deck = structuredClone(source);
  const removed = { slides: deck.slides.length, timeline: 0, links: 0, interactions: 0, details: 0, deepDives: 0, mediaAutoplay: 0 };
  deck.transition = "none";
  deck.motion = { ...deck.motion, ...STILL_MOTION };
  for (const slide of deck.slides) {
    removed.timeline += slide.timeline?.length ?? 0;
    removed.details += slide.details?.length ?? 0;
    if (slide.drillOf) removed.deepDives += 1;
    delete slide.timeline;
    delete slide.details;
    delete slide.drillOf;
    delete slide.transitionDur;
    delete slide.transitionSound;
    delete slide.transitionDir;
    delete slide.advance;
    slide.animation = "none";
    slide.photoMotion = "none";
    slide.entrance = "none";
    slide.emphasis = "none";
    slide.kinetic = "none";
    slide.backdrop = "none";
    slide.transition = "none";
    if (playableMedia(slide.media) && slide.media.autoplay !== false) {
      slide.media.autoplay = false;
      removed.mediaAutoplay += 1;
    }
    for (const object of slide.elements ?? []) {
      if (object.action) { delete object.action; removed.links += 1; }
      if (object.overAction) { delete object.overAction; removed.links += 1; }
      if (object.item) { delete object.item; removed.links += 1; }
      if (object.hover) { delete object.hover; removed.interactions += 1; }
      if (object.tip) { delete object.tip; removed.interactions += 1; }
      if (typeof object.loop === "string") { delete object.loop; removed.interactions += 1; }
      if (["video", "youtube", "lottie"].includes(object.kind) && object.autoplay !== false) {
        object.autoplay = false;
        removed.mediaAutoplay += 1;
      }
      // A sound that starts by itself or plays on across slides waits for a click on its icon instead.
      if (object.kind === "audio" && (object.autoplay || object.across)) {
        object.autoplay = false;
        delete object.across;
        removed.mediaAutoplay += 1;
      }
    }
  }
  return { deck, removed, changed: JSON.stringify(deck) !== JSON.stringify(source) };
}
