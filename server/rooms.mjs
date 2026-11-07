// 共同編集 (co-editing): a shared copy of a deck that several browsers edit at once. A room is created from a deck and
// reached by an unguessable link; each browser listens to a stream of events (Server-Sent Events) and posts its
// changes as small operations: one slide replaced (by its sid), the order of the slides, or the deck's own settings.
// The server keeps the latest deck and passes every change on; browsers merge what arrives with what they are
// typing. Pictures, videos and sounds kept in a browser travel through the room (media by id). Rooms live in memory
// and are forgotten after a while without anyone in them; a browser that still has the deck can bring one back.
// オンライン プレゼンテーション uses a room of its own kind ("show"): viewers only watch. The presenter holds the room's
// key and posts where the show is (slide and build step); everyone watching follows.

import { randomBytes } from "node:crypto";
import { mergeOrder, mergeSlide, textHash } from "../public/editor/coedit-merge.mjs";

const ID = /^[A-Za-z0-9_-]{16,40}$/;
const CLIENT = /^[A-Za-z0-9_-]{4,40}$/;
const MEDIA_ID = /^[A-Za-z0-9_-]{1,64}$/;
// How many recent versions of each slide (and of the order) a room remembers, to merge a change made on one of them.
const KEEP_VERSIONS = 12;
const COLORS = ["#2b6be0", "#c2410c", "#0f766e", "#7c3aed", "#b45309", "#be185d", "#15803d", "#4338ca", "#a16207", "#0e7490"];

export const isRoomId = (id) => ID.test(String(id || ""));
export const newRoomId = () => randomBytes(18).toString("base64url");
const clean = (text, max) => String(text ?? "").replace(/[\u0000-\u001f<>]/g, "").trim().slice(0, max);

export function createRooms({ maxRooms = 60, maxDeckBytes = 20_000_000, maxMediaBytes = 300_000_000, idleMs = 12 * 3600_000, now = () => Date.now() } = {}) {
  const rooms = new Map();

  function validDeck(deck) {
    if (!deck || typeof deck !== "object" || !Array.isArray(deck.slides) || !deck.slides.length || deck.slides.length > 500) throw new Error("資料が読めません");
    const size = Buffer.byteLength(JSON.stringify(deck));
    if (size > maxDeckBytes) throw new Error("資料が大きすぎて共有できません（20MBまで）");
    // Every slide needs its id: changes are sent slide by slide.
    const seen = new Set();
    for (const slide of deck.slides) {
      if (!slide || typeof slide !== "object" || typeof slide.sid !== "string" || !/^[A-Za-z0-9_-]{1,32}$/.test(slide.sid) || seen.has(slide.sid)) throw new Error("スライドのIDがそろっていません");
      seen.add(slide.sid);
    }
    return deck;
  }

  /**
   * A new room for a deck (or, `id` given, the same room brought back after the server forgot it). `show` makes a
   * room to present online: it gets a key that only its presenter knows (needed to move the show or change the deck).
   */
  function create(deck, { id = null, show = false, key = null, fits = null } = {}) {
    sweep();
    if (id && rooms.has(id)) {
      const room = rooms.get(id);
      // A presenter bringing back their show after the server forgot it proves it with the key it had.
      if (room.kind === "show" && room.key !== key) throw new Error("この発表の鍵が違います");
      return room;
    }
    if (rooms.size >= maxRooms) throw new Error("共同編集の部屋がいっぱいです。しばらくしてからもう一度どうぞ");
    const room = { id: id && isRoomId(id) ? id : newRoomId(), deck: validDeck(structuredClone(deck)), version: 1, clients: new Map(), media: new Map(), mediaBytes: 0, touched: now(), createdAt: now(), seen: new Map(), orders: new Map() };
    rememberAll(room);
    if (show) Object.assign(room, { kind: "show", key: isRoomId(key) ? key : newRoomId(), fits: fitsOf(fits, room.deck), show: { index: 0, step: 0, sid: room.deck.slides[0].sid, live: false, at: now() } });
    rooms.set(room.id, room);
    return room;
  }
  const get = (id) => (isRoomId(id) ? rooms.get(id) || null : null);
  // A show carries the presenter's text fitting (font scale per slide and object), so viewers see the same slides.
  const fitsOf = (fits, deck) => (Array.isArray(fits) ? fits.slice(0, deck.slides.length).map((f) => (f && typeof f === "object" ? f : null)) : null);

  // The versions a change may have been made on: each slide's recent versions and the recent orders, by fingerprint.
  function remember(room, slide) {
    let versions = room.seen.get(slide.sid);
    if (!versions) room.seen.set(slide.sid, (versions = new Map()));
    const key = textHash(JSON.stringify(slide));
    versions.delete(key);
    versions.set(key, slide);
    if (versions.size > KEEP_VERSIONS) versions.delete(versions.keys().next().value);
  }
  function rememberOrder(room) {
    const sids = room.deck.slides.map((s) => s.sid);
    const key = textHash(JSON.stringify(sids));
    room.orders.delete(key);
    room.orders.set(key, sids);
    if (room.orders.size > KEEP_VERSIONS) room.orders.delete(room.orders.keys().next().value);
  }
  function rememberAll(room) {
    for (const slide of room.deck.slides) remember(room, slide);
    rememberOrder(room);
  }

  function people(room) {
    return [...room.clients.values()].map(({ id, name, uid, color, slide, sel, at }) => ({ id, name, uid, color, slide, sel, at }));
  }
  function send(client, event, data) {
    try { client.res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`); } catch { /* gone */ }
  }
  function broadcast(room, event, data, except = null) {
    for (const client of room.clients.values()) if (client.id !== except) send(client, event, data);
  }

  /** A browser joins: it gets the deck and who is there; the others learn it came. */
  function join(room, res, { client, name, uid }) {
    if (!CLIENT.test(String(client || ""))) throw new Error("bad client");
    const old = room.clients.get(client);
    if (old && old.res !== res) { try { old.res.end(); } catch { /* closed */ } }
    const used = new Set([...room.clients.values()].map((c) => c.color));
    const color = old?.color || COLORS.find((c) => !used.has(c)) || COLORS[room.clients.size % COLORS.length];
    const entry = { id: client, res, name: clean(name, 40) || "名前なし", uid: clean(uid, 24), color, slide: old?.slide || null, sel: old?.sel || [], at: now() };
    room.clients.set(client, entry);
    room.touched = now();
    send(entry, "snapshot", { deck: room.deck, version: room.version, people: people(room), you: { id: client, color }, kind: room.kind || "edit", show: room.show || null, fits: room.fits || null, votes: room.kind === "show" ? Object.fromEntries([...(room.votes?.keys() || [])].map((poll) => [poll, tally(room, poll)])) : undefined });
    broadcast(room, "people", { people: people(room) }, client);
    return entry;
  }
  function leave(room, client, res = null) {
    const entry = room.clients.get(client);
    if (!entry || (res && entry.res !== res)) return;
    room.clients.delete(client);
    room.touched = now();
    broadcast(room, "people", { people: people(room) });
  }

  /**
   * Apply a browser's changes to the room's deck; everyone else hears them. A change made on an older version of a
   * slide (or of the order) than the room has now — another person's change crossed it — is merged three ways with
   * what the room has, and the merged result goes back to the sender too.
   */
  function apply(room, client, ops) {
    if (room.kind === "show") throw new Error("発表を見る部屋では資料を変更できません");
    if (!Array.isArray(ops) || !ops.length || ops.length > 600) throw new Error("変更が読めません");
    const deck = room.deck;
    const applied = [];
    const merged = [];
    for (const op of ops) {
      if (op?.t === "slide" && op.slide && typeof op.slide === "object" && op.slide.sid === op.sid) {
        const i = deck.slides.findIndex((s) => s.sid === op.sid);
        if (i >= 0) {
          const current = deck.slides[i];
          const base = typeof op.base === "string" && op.base !== textHash(JSON.stringify(current)) ? room.seen.get(op.sid)?.get(op.base) : null;
          if (base) {
            const slide = { ...mergeSlide(base, op.slide, current), sid: op.sid };
            deck.slides[i] = slide;
            remember(room, slide);
            const out = { t: "slide", sid: op.sid, slide, merged: true };
            applied.push(out);
            merged.push(out);
            continue;
          }
          deck.slides[i] = op.slide;
        } else if (Number.isInteger(op.at)) deck.slides.splice(Math.max(0, Math.min(deck.slides.length, op.at)), 0, op.slide);
        else continue;
        remember(room, op.slide);
        applied.push(op);
      } else if (op?.t === "order" && Array.isArray(op.sids)) {
        const bySid = new Map(deck.slides.map((s) => [s.sid, s]));
        for (const s of Array.isArray(op.add) ? op.add : []) if (s && typeof s.sid === "string" && !bySid.has(s.sid)) bySid.set(s.sid, s);
        if (new Set(op.sids).size !== op.sids.length) continue;
        const now = deck.slides.map((s) => s.sid);
        const base = typeof op.base === "string" && op.base !== textHash(JSON.stringify(now)) ? room.orders.get(op.base) : null;
        const sids = base ? mergeOrder(base, op.sids, now) : op.sids;
        const next = sids.map((sid) => bySid.get(sid)).filter(Boolean);
        if (!next.length || next.length > 500) continue;
        deck.slides = next;
        for (const s of next) if (!room.seen.has(s.sid)) remember(room, s);
        rememberOrder(room);
        const out = base ? { t: "order", sids: next.map((s) => s.sid), add: op.add, merged: true } : op;
        applied.push(out);
        if (base) merged.push(out);
      } else if (op?.t === "deck" && op.meta && typeof op.meta === "object") {
        for (const [key, value] of Object.entries(op.meta)) {
          if (key === "slides") continue;
          if (value === null) delete deck[key]; else deck[key] = value;
        }
        applied.push(op);
      }
    }
    if (!applied.length) return room.version;
    if (Buffer.byteLength(JSON.stringify(deck)) > maxDeckBytes) throw new Error("資料が大きすぎます");
    room.version += 1;
    room.touched = now();
    broadcast(room, "ops", { client, version: room.version, ops: applied }, client);
    const sender = room.clients.get(client);
    if (merged.length && sender) send(sender, "ops", { client, version: room.version, ops: merged });
    return room.version;
  }

  /** Where someone is: the slide they are on and what they have selected (drawn in their colour for the others). */
  function presence(room, client, { slide, sel, name } = {}) {
    const entry = room.clients.get(client);
    if (!entry) return;
    entry.slide = typeof slide === "string" ? slide.slice(0, 32) : null;
    entry.sel = Array.isArray(sel) ? sel.filter((x) => typeof x === "string").slice(0, 50).map((x) => x.slice(0, 40)) : [];
    if (name) entry.name = clean(name, 40) || entry.name;
    entry.at = now();
    broadcast(room, "presence", { id: client, name: entry.name, color: entry.color, slide: entry.slide, sel: entry.sel }, client);
  }

  const isPresenter = (room, key) => room.kind === "show" && typeof key === "string" && key === room.key;

  /**
   * The presenter moved on: where the show is now (`live` false while the presenter is not presenting, `over` once
   * the online presentation has ended). Everyone watching follows.
   */
  function setShow(room, key, { index, step, sid, ended, over } = {}) {
    if (!isPresenter(room, key)) throw new Error("発表者だけが進められます");
    const i = Number.isInteger(index) && index >= 0 && index < room.deck.slides.length ? index : Math.max(0, room.deck.slides.findIndex((s) => s.sid === sid));
    room.show = { index: i, step: Number.isFinite(step) ? Math.max(0, Math.min(10_000, Math.floor(step))) : 0, sid: room.deck.slides[i]?.sid || null, live: !ended && !over, ...(over ? { over: true } : {}), at: now() };
    room.touched = now();
    broadcast(room, "show", room.show);
    // The end of the online presentation: the room goes at once (its viewers have heard; their streams close soon).
    if (over) {
      rooms.delete(room.id);
      setTimeout(() => { for (const client of room.clients.values()) { try { client.res.end(); } catch { /* gone */ } } }, 1500).unref?.();
    }
    return room.show;
  }

  /** The presenter changed the deck during the show (資料の更新): viewers get the new one. */
  function replaceDeck(room, key, deck, fits = null) {
    if (!isPresenter(room, key)) throw new Error("発表者だけが資料を更新できます");
    room.deck = validDeck(structuredClone(deck));
    room.fits = fitsOf(fits, room.deck);
    rememberAll(room);
    room.version += 1;
    room.touched = now();
    if (room.show && room.show.index >= room.deck.slides.length) room.show.index = room.deck.slides.length - 1;
    broadcast(room, "deck", { deck: room.deck, version: room.version, show: room.show, fits: room.fits });
    return room.version;
  }

  // PowerPoint Live: what the audience sends (reactions, votes) and the presenter's captions, passed to everyone.
  const REACTIONS = ["👍", "❤️", "👏", "😮", "💡", "😂"];
  function allowed(room, client, kind, perWindow, windowMs = 4000) {
    const entry = room.clients.get(client);
    if (!entry) return false;
    const t = now();
    entry.sent = (entry.sent || []).filter((x) => x.kind !== kind || t - x.at < windowMs);
    if (entry.sent.filter((x) => x.kind === kind).length >= perWindow) return false;
    entry.sent.push({ kind, at: t });
    return true;
  }
  /** A viewer's reaction (one of REACTIONS), a few every few seconds at most: everyone sees it float up. */
  function react(room, client, emoji) {
    if (room.kind !== "show" || !REACTIONS.includes(emoji)) throw new Error("リアクションが読めません");
    if (!allowed(room, client, "react", 6)) return false;
    broadcast(room, "react", { emoji, from: client });
    return true;
  }
  /** A vote in a poll (アンケート) on the deck's slides: one answer per viewer per poll (a new one replaces it). */
  function vote(room, client, { poll, option } = {}) {
    if (room.kind !== "show") throw new Error("投票できません");
    const object = room.deck.slides.flatMap((s) => s.elements || []).find((o) => o.kind === "poll" && o.id === poll);
    if (!object || !Number.isInteger(option) || option < 0 || option >= (object.options || []).length) throw new Error("投票が読めません");
    if (!room.clients.has(client) || !allowed(room, client, "vote", 10)) return null;
    room.votes = room.votes || new Map();
    const answers = room.votes.get(poll) || new Map();
    answers.set(client, option);
    room.votes.set(poll, answers);
    const counts = tally(room, poll);
    broadcast(room, "votes", { poll, counts });
    return counts;
  }
  function tally(room, poll) {
    const object = room.deck.slides.flatMap((s) => s.elements || []).find((o) => o.kind === "poll" && o.id === poll);
    const counts = new Array((object?.options || []).length).fill(0);
    for (const option of room.votes?.get(poll)?.values() || []) if (option < counts.length) counts[option] += 1;
    return counts;
  }
  /** The presenter's live captions (字幕), shown on every viewer's screen. */
  function caption(room, key, text) {
    if (!isPresenter(room, key)) throw new Error("発表者だけが字幕を送れます");
    broadcast(room, "caption", { text: clean(text, 300) });
  }

  function putMedia(room, mid, type, body) {
    if (!MEDIA_ID.test(mid)) throw new Error("bad media id");
    if (room.media.has(mid)) return;
    if (room.mediaBytes + body.length > maxMediaBytes) throw new Error("共有できるファイルの量を超えました");
    room.media.set(mid, { type: clean(type, 100) || "application/octet-stream", body });
    room.mediaBytes += body.length;
    room.touched = now();
  }
  const getMedia = (room, mid) => room.media.get(mid) || null;

  /** Rooms nobody has used for a while are forgotten. */
  function sweep() {
    for (const [id, room] of rooms) if (!room.clients.size && now() - room.touched > idleMs) rooms.delete(id);
  }
  function heartbeat() {
    for (const room of rooms.values()) for (const client of room.clients.values()) { try { client.res.write(": ping\n\n"); } catch { /* gone */ } }
    sweep();
  }

  /** The server is stopping: end every stream (the browsers reconnect and bring their rooms back). */
  function closeAll() {
    for (const room of rooms.values()) for (const client of room.clients.values()) { try { client.res.end(); } catch { /* gone */ } }
  }

  return { create, get, join, leave, apply, presence, setShow, replaceDeck, isPresenter, react, vote, caption, putMedia, getMedia, people, sweep, heartbeat, closeAll, rooms };
}
