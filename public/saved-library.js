const DB_NAME = "html-pptx-studio-library";
const STORE = "decks";
const SKIP_FIELDS = new Set([
  "media", "customImage", "src", "visualAsset", "imagePlacement", "photoMotion",
  "animation", "kinetic", "backdrop", "entrance", "emphasis", "icon", "type",
  "chartType", "fit", "kind", "imagePosition", "target", "theme", "transition", "formula",
]);

export function normalizeSearch(text) {
  return String(text ?? "").normalize("NFKC").toLocaleLowerCase("ja-JP").replace(/\s+/g, " ").trim();
}

function textValues(value, key = "", out = []) {
  if (SKIP_FIELDS.has(key) || value == null) return out;
  if (typeof value === "string") {
    if (value && !/^(?:data:|idb:|https?:\/\/)/i.test(value)) out.push(value.replace(/\*\*/g, ""));
  } else if (Array.isArray(value)) {
    for (const item of value) textValues(item, "", out);
  } else if (typeof value === "object") {
    for (const [childKey, item] of Object.entries(value)) textValues(item, childKey, out);
  } else if (typeof value === "number") {
    out.push(String(value));
  }
  return out;
}

export function buildSearchIndex(deck) {
  return {
    deck: normalizeSearch([deck.title, deck.purpose, deck.audience, deck.memo].join(" ")),
    slides: deck.slides.map((slide) => normalizeSearch(textValues(slide).join(" "))),
  };
}

export function slideExcerpt(slide, query = "") {
  const values = textValues(slide);
  const terms = normalizeSearch(query).split(" ").filter(Boolean);
  const text = values.find((value) => terms.some((term) => normalizeSearch(value).includes(term))) || values.find(Boolean) || "";
  return text.length > 110 ? `${text.slice(0, 107)}…` : text;
}

export function searchSavedDecks(records, query) {
  const terms = normalizeSearch(query).split(" ").filter(Boolean);
  return records.map((record) => {
    const index = record.searchIndex || buildSearchIndex(record.deck);
    const all = `${index.deck} ${index.slides.join(" ")}`;
    if (!terms.every((term) => all.includes(term))) return null;
    const exactSlideMatches = terms.length
      ? index.slides.flatMap((text, number) => terms.every((term) => text.includes(term)) ? [number] : [])
      : record.deck.slides.map((_, number) => number);
    const slideMatches = exactSlideMatches.length ? exactSlideMatches : record.deck.slides.map((_, number) => number);
    return { record, slideMatches };
  }).filter(Boolean).sort((a, b) => b.record.updatedAt - a.record.updatedAt);
}

function openLibrary() {
  if (!globalThis.indexedDB) return Promise.reject(new Error("このブラウザでは資料を保存できません。"));
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(STORE, { keyPath: "id" });
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
    request.onblocked = () => reject(new Error("保存庫を開けません。ほかのタブを閉じて再試行してください。"));
  });
}

async function transaction(mode, method, value) {
  const db = await openLibrary();
  try {
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, mode);
      const request = tx.objectStore(STORE)[method](value);
      tx.oncomplete = () => resolve(request.result);
      tx.onerror = () => reject(tx.error || new Error("保存庫へのアクセスに失敗しました。"));
      tx.onabort = () => reject(tx.error || new Error("保存庫へのアクセスが中断されました。"));
    });
  } finally {
    db.close();
  }
}

export const listSavedDecks = () => transaction("readonly", "getAll");
export const getSavedDeck = (id) => transaction("readonly", "get", id);
export const putSavedDeck = (record) => transaction("readwrite", "put", record);
export const deleteSavedDeck = (id) => transaction("readwrite", "delete", id);
