// 共同編集 (co-editing), as PowerPoint's 共有: share the deck with a link and edit it together at the same time. The
// server keeps a room (server/rooms.mjs); this browser listens to it (Server-Sent Events) and posts its own changes
// slide by slide. What arrives is merged with what is being edited here: another person's change to a slide this
// browser has not touched simply replaces it; two people on the same slide keep each other's objects (a three-way
// merge by object id); the order of the slides and the deck's settings merge the same way. Undo stays one's own:
// changes from others are written into the undo history too, so ⌘Z never takes them back.
// People are shown with their name and colour: in the top bar, on the thumbnails of the slides they are on, and
// around the objects they have selected. Pictures, videos and sounds kept in a browser go through the room.

import { avatar, cleanName, initials, userId, userName, setUserName } from "./people.mjs";
import { mergeOrder, mergeSlide, textHash } from "./coedit-merge.mjs";

export { mergeOrder, mergeSlide, textHash };

const SEND_DELAY = 250;
const PRESENCE_DELAY = 300;
const json = (value) => JSON.stringify(value ?? null);
const newClientId = () => `c${(globalThis.crypto?.randomUUID?.() || `${Date.now()}${Math.random()}`).replace(/[^a-z0-9]/gi, "").slice(0, 20)}`;
const metaOf = (deck) => { const { slides, ...meta } = deck || {}; void slides; return meta; };

export function createCoedit(app) {
  const { h, E } = app;
  let room = null; // { id, es, client, color, version, base: Map, order: [], meta, people: Map, status }
  let sendTimer = 0;
  let presenceTimer = 0;
  let queue = [];
  let flushTimer = 0;
  const uploaded = new Set();
  const others = () => (room ? [...room.people.values()].filter((p) => p.id !== room.client) : []);

  // ---------------------------------------------------------------- starting, joining, leaving

  async function ensureName() {
    if (userName()) return userName();
    const typed = await app.ask("あなたの名前", "共同編集で相手に表示する名前（このブラウザに保存します）", "");
    return setUserName(typed || "") || "";
  }

  /** 共有: make a room from the deck on the stage and open its link for others. */
  async function start() {
    if (room) { showMenu(); return; }
    const deck = app.deck();
    if (!deck) return;
    if (!(await app.confirm("共同編集を始めます。リンクを知っている人は、この資料を同時に見て編集できます（部屋はサーバーに一時的に置かれ、しばらく誰もいないと閉じます）。始めますか？"))) return;
    if (!(await ensureName())) { app.toast("名前を入れると共同編集を始められます"); return; }
    app.ensureAllSids();
    try {
      const response = await fetch("/api/rooms", { method: "POST", credentials: "same-origin", headers: { "content-type": "application/json" }, body: JSON.stringify({ deck: app.deck() }) });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || `HTTP ${response.status}`);
      connect(result.id, { host: true });
      shareLink();
    } catch (error) {
      app.toast(`共同編集を始められません：${error.message}`);
    }
  }
  /** Opened from a link (…?room=ID): join the room, its deck comes with the first message. */
  async function join(id) {
    if (!(await ensureName())) { app.toast("名前を入れると共同編集に参加できます"); return; }
    connect(id, { host: false });
  }
  function connect(id, { host }) {
    disconnect({ quiet: true });
    room = { id, es: null, client: newClientId(), color: null, version: 0, base: new Map(), order: [], meta: null, people: new Map(), status: "connecting", host, joined: false };
    setUrl(id);
    open();
    render();
  }
  function open() {
    if (!room) return;
    const params = new URLSearchParams({ client: room.client, name: userName() || "名前なし", uid: userId() });
    const es = new EventSource(`/api/rooms/${encodeURIComponent(room.id)}?${params}`);
    room.es = es;
    es.addEventListener("snapshot", (event) => onSnapshot(JSON.parse(event.data)));
    es.addEventListener("ops", (event) => onOps(JSON.parse(event.data)));
    es.addEventListener("people", (event) => onPeople(JSON.parse(event.data).people));
    es.addEventListener("presence", (event) => onPresence(JSON.parse(event.data)));
    es.onerror = () => {
      if (!room || room.es !== es) return;
      room.status = "reconnecting";
      render();
      // A room the server no longer has (restarted) is brought back from this browser's copy.
      if (es.readyState === EventSource.CLOSED) setTimeout(revive, 1500);
      else fetch(`/api/rooms/${encodeURIComponent(room.id)}/presence`, { method: "POST", credentials: "same-origin", headers: { "content-type": "application/json" }, body: JSON.stringify({ client: room.client }) })
        .then((r) => r.json().catch(() => ({}))).then((r) => { if (r?.gone) revive(); }).catch(() => {});
    };
  }
  async function revive() {
    if (!room || room.status === "connected") return;
    const deck = app.deck();
    try {
      const response = await fetch("/api/rooms", { method: "POST", credentials: "same-origin", headers: { "content-type": "application/json" }, body: JSON.stringify({ id: room.id, deck }) });
      if (!response.ok) throw new Error();
      try { room.es?.close(); } catch { /* closed */ }
      uploaded.clear();
      open();
    } catch {
      setTimeout(revive, 5000);
    }
  }
  function disconnect({ quiet = false } = {}) {
    if (!room) return;
    clearTimeout(sendTimer);
    try { room.es?.close(); } catch { /* closed */ }
    navigator.sendBeacon?.(`/api/rooms/${encodeURIComponent(room.id)}/leave`, new Blob([JSON.stringify({ client: room.client })], { type: "application/json" }));
    room = null;
    queue = [];
    setUrl(null);
    render();
    app.redraw();
    if (!quiet) app.toast("共同編集から抜けました（資料はこのブラウザに残っています）");
  }
  function setUrl(id) {
    try {
      const url = new URL(location.href);
      if (id) url.searchParams.set("room", id); else url.searchParams.delete("room");
      history.replaceState(null, "", url);
    } catch { /* file: */ }
  }
  const link = () => { const url = new URL(location.href); url.search = ""; url.searchParams.set("room", room.id); url.hash = ""; return url.toString(); };
  async function shareLink() {
    if (!room) return;
    const text = link();
    let copied = false;
    try { await navigator.clipboard.writeText(text); copied = true; } catch { /* no clipboard */ }
    app.showLink(text, copied);
  }

  // ---------------------------------------------------------------- what arrives

  function onSnapshot({ deck, version, people, you }) {
    if (!room) return;
    room.status = "connected";
    room.version = version;
    room.color = you?.color || null;
    onPeople(people);
    if (!room.joined && !room.host) {
      // Joining: the room's deck replaces the one on the stage (that one stays in this browser's history).
      app.loadShared(deck);
      setBase(deck);
      room.joined = true;
      app.toast(`共同編集に参加しました（${others().length + 1}人）`);
    } else if (!room.joined) {
      setBase(deck);
      room.joined = true;
      uploadMedia(app.deck());
    } else {
      // Back after a break: whatever changed meanwhile is merged with what was done here.
      apply({ slides: deck.slides, meta: metaOf(deck) });
    }
    fetchMissingMedia(deck);
    sendPresence();
    render();
  }
  function setBase(deck) {
    room.base = new Map(deck.slides.map((s) => [s.sid, json(s)]));
    room.order = deck.slides.map((s) => s.sid);
    room.meta = json(metaOf(deck));
  }
  function onOps({ version, ops }) {
    if (!room) return;
    room.version = Math.max(room.version, version);
    queue.push(...ops);
    flush();
  }
  /** Changes from others wait while this browser types into a slide, then apply together. */
  function flush() {
    clearTimeout(flushTimer);
    if (!queue.length || !room) return;
    if (app.busy()) { flushTimer = setTimeout(flush, 350); return; }
    const ops = queue;
    queue = [];
    const slides = [];
    let order = null;
    let add = [];
    let meta = null;
    for (const op of ops) {
      if (op.t === "slide") slides.push(op.slide);
      else if (op.t === "order") { order = op.sids; add = [...add, ...(op.add || [])]; }
      else if (op.t === "deck") meta = { ...(meta || JSON.parse(room.meta || "{}")), ...op.meta };
    }
    apply({ slides, order, add, meta: meta && Object.fromEntries(Object.entries(meta).filter(([, v]) => v !== null)) });
    fetchMissingMedia({ slides: [...slides, ...add] });
  }
  /** Merge others' slides, order and settings into the deck here (no undo step; the undo history learns them too). */
  function apply({ slides = [], order = null, add = [], meta = null }) {
    const deck = app.deck();
    if (!deck || !room) return;
    const local = new Map(deck.slides.map((s) => [s.sid, s]));
    const merged = new Map(local);
    const theirs = new Map();
    for (const s of [...add, ...slides]) {
      if (!s?.sid) continue;
      theirs.set(s.sid, s);
      const base = room.base.has(s.sid) ? JSON.parse(room.base.get(s.sid)) : null;
      // Several changes to one slide in a row (another person's, then the room's merge of ours) build on each other.
      const mine = merged.get(s.sid);
      merged.set(s.sid, mine && base ? mergeSlide(base, mine, s) : s);
      room.base.set(s.sid, json(s));
    }
    let sids = deck.slides.map((s) => s.sid);
    if (order) {
      sids = mergeOrder(room.order, sids, order);
      room.order = order;
      for (const sid of [...room.base.keys()]) if (!order.includes(sid)) room.base.delete(sid);
    }
    let nextMeta = metaOf(deck);
    if (meta) {
      const base = JSON.parse(room.meta || "{}");
      const mineMeta = metaOf(deck);
      const keys = new Set([...Object.keys(base), ...Object.keys(meta), ...Object.keys(mineMeta)]);
      nextMeta = {};
      for (const key of keys) {
        const value = json(mineMeta[key]) !== json(base[key]) ? mineMeta[key] : meta[key];
        if (value !== undefined) nextMeta[key] = value;
      }
      room.meta = json(meta);
    }
    const next = { ...nextMeta, slides: sids.map((sid) => merged.get(sid)).filter(Boolean) };
    app.applyRemote(next, { slides: theirs, order, meta });
    // What was merged on top of theirs is ours to send.
    changed();
  }
  function onPeople(list) {
    if (!room) return;
    const before = new Set(room.people.keys());
    room.people = new Map((list || []).map((p) => [p.id, { ...room.people.get(p.id), ...p }]));
    for (const p of room.people.values()) if (!before.has(p.id) && p.id !== room.client && room.joined) app.toast(`${p.name}さんが参加しました`);
    for (const id of before) if (!room.people.has(id) && id !== room.client) app.toast("参加者が1人抜けました");
    render();
    app.redraw();
  }
  function onPresence(p) {
    if (!room) return;
    room.people.set(p.id, { ...room.people.get(p.id), ...p });
    render();
    app.redraw();
  }

  // ---------------------------------------------------------------- what is sent

  /** Something changed here (markChanged): the differences go to the room shortly. */
  function changed() {
    if (!room?.joined) return;
    clearTimeout(sendTimer);
    sendTimer = setTimeout(send, SEND_DELAY);
  }
  async function send() {
    if (!room?.joined || room.status !== "connected") { if (room) sendTimer = setTimeout(send, 1000); return; }
    // Changes from others that arrived meanwhile are merged in first (unless typing holds them; the room merges then).
    if (queue.length && !app.busy()) flush();
    app.ensureAllSids();
    const deck = app.deck();
    const ops = [];
    const sids = deck.slides.map((s) => s.sid);
    const known = new Set(room.order);
    // Each change names the version it was made on, so the room can merge it with a change that crossed it.
    if (json(sids) !== json(room.order)) ops.push({ t: "order", sids, add: deck.slides.filter((s) => !known.has(s.sid)), base: textHash(json(room.order)) });
    for (const s of deck.slides) {
      if (!known.has(s.sid)) continue;
      if (room.base.get(s.sid) !== json(s)) ops.push({ t: "slide", sid: s.sid, slide: s, base: room.base.has(s.sid) ? textHash(room.base.get(s.sid)) : null });
    }
    const meta = metaOf(deck);
    if (json(meta) !== room.meta) {
      const base = JSON.parse(room.meta || "{}");
      const patch = {};
      for (const key of new Set([...Object.keys(base), ...Object.keys(meta)])) if (json(meta[key]) !== json(base[key])) patch[key] = meta[key] === undefined ? null : meta[key];
      ops.push({ t: "deck", meta: patch });
    }
    if (!ops.length) return;
    // What is being sent becomes the base (if the send fails, the room is revived from this browser's copy).
    const snapshot = { order: room.order, base: new Map(room.base), meta: room.meta };
    room.order = sids;
    for (const s of deck.slides) room.base.set(s.sid, json(s));
    for (const sid of [...room.base.keys()]) if (!sids.includes(sid)) room.base.delete(sid);
    room.meta = json(meta);
    uploadMedia({ slides: ops.flatMap((op) => (op.slide ? [op.slide] : op.add || [])) });
    try {
      const response = await fetch(`/api/rooms/${encodeURIComponent(room.id)}/ops`, { method: "POST", credentials: "same-origin", headers: { "content-type": "application/json" }, body: JSON.stringify({ client: room.client, ops }) });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw Object.assign(new Error(result.error || `HTTP ${response.status}`), { gone: result.gone });
      if (room) room.version = Math.max(room.version, result.version || 0);
    } catch (error) {
      if (!room) return;
      Object.assign(room, snapshot);
      if (error.gone) { room.status = "reconnecting"; render(); revive(); }
      else { app.toast(`共同編集に送れませんでした：${error.message}`); sendTimer = setTimeout(send, 3000); }
    }
  }
  /** Where I am (the slide on the stage) and what I have selected, for the others. */
  function sendPresence() {
    if (!room?.joined) return;
    clearTimeout(presenceTimer);
    presenceTimer = setTimeout(() => {
      if (!room) return;
      const slide = app.deck()?.slides[app.index()]?.sid || null;
      fetch(`/api/rooms/${encodeURIComponent(room.id)}/presence`, { method: "POST", credentials: "same-origin", headers: { "content-type": "application/json" }, body: JSON.stringify({ client: room.client, slide, sel: app.selection(), name: userName() }) }).catch(() => {});
    }, PRESENCE_DELAY);
  }

  // ---------------------------------------------------------------- media through the room

  const mediaIds = (slides) => (slides || []).flatMap((s) => [s?.media?.src, s?.background?.image, ...(s?.elements || []).flatMap((o) => [o?.src, o?.fillImg])]).filter((src) => typeof src === "string" && src.startsWith("idb:"));
  async function uploadMedia(deck) {
    if (!room) return;
    for (const src of new Set(mediaIds(deck?.slides))) {
      if (uploaded.has(src)) continue;
      uploaded.add(src);
      try {
        const blob = await app.mediaBlob(src);
        if (!blob) continue;
        await fetch(`/api/rooms/${encodeURIComponent(room.id)}/media/${encodeURIComponent(src.slice(4))}`, { method: "PUT", credentials: "same-origin", headers: { "content-type": blob.type || "application/octet-stream" }, body: blob });
      } catch { uploaded.delete(src); }
    }
  }
  async function fetchMissingMedia(deck) {
    if (!room) return;
    let added = false;
    for (const src of new Set(mediaIds(deck?.slides))) {
      if (await app.hasMedia(src)) continue;
      try {
        const response = await fetch(`/api/rooms/${encodeURIComponent(room.id)}/media/${encodeURIComponent(src.slice(4))}`, { credentials: "same-origin" });
        if (!response.ok) continue;
        await app.putMediaAs(src, await response.blob());
        uploaded.add(src);
        added = true;
      } catch { /* try again with the next change */ }
    }
    if (added) app.mediaArrived();
  }

  // ---------------------------------------------------------------- showing people

  const bar = document.getElementById("coeditBar");
  function render() {
    const btn = document.getElementById("shareBtn");
    if (btn) {
      btn.classList.toggle("is-shared", Boolean(room));
      btn.replaceChildren(...(room ? [
        h("span", { class: ["co-dot", room.status] }),
        ...others().slice(0, 4).map((p) => { const a = avatar(h, p.name, { key: p.uid || p.id, size: 22, title: `${p.name}（${slideLabel(p.slide)}）` }); if (p.color) a.style.background = p.color; return a; }),
        h("span", {}, room.status === "connected" ? `${others().length + 1}人` : room.status === "connecting" ? "接続中…" : "再接続中…"),
      ] : [h("span", {}, "共有")]));
      btn.title = room ? `共同編集中（${others().length + 1}人）：参加者・リンクのコピー・抜ける` : "共同編集：リンクを知っている人と同時に編集します";
    }
    if (bar) bar.hidden = true;
  }
  const slideLabel = (sid) => { const i = app.deck()?.slides.findIndex((s) => s.sid === sid) ?? -1; return i >= 0 ? `スライド ${i + 1}` : "—"; };
  function showMenu() {
    if (!room) return;
    const btn = document.getElementById("shareBtn");
    const items = [
      { head: "共同編集" },
      { label: "リンクをコピー", icon: "link", run: () => shareLink() },
      "-", { head: `参加者（${others().length + 1}人）` },
      { label: `${userName() || "名前なし"}（自分）・${slideLabel(app.deck()?.slides[app.index()]?.sid)}`, icon: "people", run: () => {} },
      ...others().map((p) => ({ label: `${p.name}・${slideLabel(p.slide)}`, icon: "people", run: () => { const i = app.deck()?.slides.findIndex((s) => s.sid === p.slide); if (i >= 0) app.select(i); } })),
      "-", { label: "共同編集から抜ける", icon: "close", run: () => disconnect() },
    ];
    app.openMenu(btn, items);
  }
  /** Avatars on the thumbnails of the slides the others are on. */
  function decorateFilm(items) {
    if (!room) return;
    const deck = app.deck();
    for (const p of others()) {
      const i = deck?.slides.findIndex((s) => s.sid === p.slide);
      const item = i >= 0 ? [...items].find((el) => el.dataset.index === String(i)) : null;
      if (!item) continue;
      let holder = item.querySelector(".co-film");
      if (!holder) { holder = h("span", { class: "co-film" }); item.append(holder); }
      const a = avatar(h, p.name, { key: p.uid || p.id, size: 18, title: `${p.name}さんが表示中` });
      a.style.background = p.color || a.style.background;
      holder.append(a);
    }
  }
  /** Outlines around what the others have selected on this slide, in their colour, with their name. */
  function overlay(list, k) {
    if (!room) return [];
    const sid = app.deck()?.slides[app.index()]?.sid;
    const out = [];
    for (const p of others()) {
      if (p.slide !== sid || !p.sel?.length) continue;
      const chosen = list.filter((o) => p.sel.includes(o.id) && !o.hidden);
      if (!chosen.length) continue;
      const pts = chosen.flatMap((o) => (o.kind === "line" ? E.lineEnds(o, list) : E.corners(o)));
      const xs = pts.map((q) => q[0]);
      const ys = pts.map((q) => q[1]);
      const b = { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
      out.push(h("div", { class: "co-sel", style: { left: `${b.x * k - 4}px`, top: `${b.y * k - 4}px`, width: `${b.w * k + 8}px`, height: `${b.h * k + 8}px`, "--co": p.color || "#2b6be0" } },
        h("span", { class: "co-sel-name" }, p.name)));
    }
    return out;
  }

  return {
    start, join, disconnect, changed, sendPresence, decorateFilm, overlay, render, showMenu,
    get active() { return Boolean(room); }, get id() { return room?.id || null; },
    info: () => (room ? { id: room.id, status: room.status, version: room.version, people: room.people.size, joined: room.joined } : null),
    names: () => others().map((p) => p.name),
    rename: () => { sendPresence(); },
    initials, cleanName,
  };
}
