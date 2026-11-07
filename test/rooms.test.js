// 共同編集 (server/rooms.mjs): rooms by unguessable id, a snapshot on joining, changes passed on, presence, media.
import assert from "node:assert/strict";
import test from "node:test";

import { createRooms, isRoomId } from "../server/rooms.mjs";

/** A fake event stream: what the server wrote, as { event, data } pairs. */
function stream() {
  const events = [];
  return { events, write(chunk) { const m = /^event: (\w+)\ndata: (.*)\n\n$/s.exec(chunk); if (m) events.push({ event: m[1], data: JSON.parse(m[2]) }); }, end() { this.ended = true; } };
}
const deck = () => ({ title: "共有", theme: "clarity", slides: [{ type: "title", title: "表紙", sid: "s1" }, { type: "blank", title: "本文", sid: "s2" }, { type: "closing", sid: "s3" }] });

test("a room is made from a deck with slide ids, under an id nobody can guess", () => {
  const rooms = createRooms();
  const room = rooms.create(deck());
  assert.ok(isRoomId(room.id) && room.id.length >= 24);
  assert.notEqual(rooms.create(deck()).id, room.id);
  assert.equal(rooms.get(room.id), room);
  assert.equal(rooms.get("short"), null);
  assert.throws(() => rooms.create({ slides: [{ type: "title" }] }), /ID/);
  assert.throws(() => rooms.create({ slides: [{ sid: "a" }, { sid: "a" }] }), /ID/);
  // A browser that still has the deck brings a forgotten room back under the same link.
  const again = createRooms().create(deck(), { id: room.id });
  assert.equal(again.id, room.id);
});

test("joining gives the deck and who is there; changes go to the others only", () => {
  const rooms = createRooms();
  const room = rooms.create(deck());
  const a = stream();
  const b = stream();
  rooms.join(room, a, { client: "clientA", name: "山田", uid: "u1" });
  rooms.join(room, b, { client: "clientB", name: "佐藤", uid: "u2" });
  assert.equal(a.events[0].event, "snapshot");
  assert.equal(a.events[0].data.deck.slides.length, 3);
  assert.deepEqual(a.events.at(-1).data.people.map((p) => p.name), ["山田", "佐藤"], "A hears that B came");
  assert.notEqual(a.events[0].data.you.color, b.events[0].data.you.color, "each has a colour of their own");
  const version = rooms.apply(room, "clientA", [{ t: "slide", sid: "s2", slide: { type: "blank", title: "直した本文", sid: "s2" } }]);
  assert.equal(version, 2);
  assert.equal(room.deck.slides[1].title, "直した本文");
  const got = b.events.at(-1);
  assert.equal(got.event, "ops");
  assert.equal(got.data.client, "clientA");
  assert.equal(got.data.ops[0].slide.title, "直した本文");
  assert.ok(!a.events.some((e) => e.event === "ops"), "not echoed to the sender");
});

test("slides added, removed and reordered; the deck's own settings; nonsense ignored", () => {
  const rooms = createRooms();
  const room = rooms.create(deck());
  rooms.apply(room, "c1", [{ t: "order", sids: ["s1", "s4", "s3"], add: [{ type: "blank", title: "新しい", sid: "s4" }] }]);
  assert.deepEqual(room.deck.slides.map((s) => s.sid), ["s1", "s4", "s3"], "s4 in, s2 out");
  rooms.apply(room, "c1", [{ t: "deck", meta: { title: "新しい題", transition: null, slides: [] } }]);
  assert.equal(room.deck.title, "新しい題");
  assert.equal(room.deck.slides.length, 3, "slides are not deck settings");
  const before = room.version;
  rooms.apply(room, "c1", [{ t: "order", sids: ["s1", "s1"] }, { t: "slide", sid: "zz", slide: { sid: "other" } }, { t: "unknown" }]);
  assert.equal(room.version, before, "nothing applied, nothing broadcast");
  assert.throws(() => rooms.apply(room, "c1", "x"));
});

test("presence, leaving, media and forgetting empty rooms", () => {
  let clock = 0;
  const rooms = createRooms({ idleMs: 1000, now: () => clock });
  const room = rooms.create(deck());
  const a = stream();
  const b = stream();
  rooms.join(room, a, { client: "clientA", name: "山田" });
  rooms.join(room, b, { client: "clientB", name: "佐藤" });
  rooms.presence(room, "clientA", { slide: "s2", sel: ["o1", 5] });
  assert.deepEqual(b.events.at(-1), { event: "presence", data: { id: "clientA", name: "山田", color: room.clients.get("clientA").color, slide: "s2", sel: ["o1"] } });
  rooms.putMedia(room, "abc123", "audio/webm", Buffer.from("sound"));
  assert.equal(rooms.getMedia(room, "abc123").body.toString(), "sound");
  assert.throws(() => rooms.putMedia(room, "../x", "a", Buffer.from("")));
  rooms.leave(room, "clientA", a);
  assert.deepEqual(b.events.at(-1).data.people.map((p) => p.name), ["佐藤"]);
  rooms.leave(room, "clientB", b);
  clock = 5000;
  rooms.sweep();
  assert.equal(rooms.get(room.id), null, "forgotten after a while with nobody in it");
});

test("changes that cross are merged on the server: a change made on an older version does not undo the other one", async () => {
  const { textHash } = await import("../public/editor/coedit-merge.mjs");
  const h = (value) => textHash(JSON.stringify(value));
  const rooms = createRooms();
  const room = rooms.create({ title: "共有", slides: [{ type: "blank", sid: "s1", elements: [{ id: "t", kind: "text", text: "メモ" }] }, { type: "closing", sid: "s2" }] });
  const a = stream();
  const b = stream();
  rooms.join(room, a, { client: "clientA", name: "山田" });
  rooms.join(room, b, { client: "clientB", name: "佐藤" });
  const start = JSON.parse(JSON.stringify(room.deck.slides[0]));
  // B adds an ellipse; A, not having heard of it yet, edits the text on the version both started from.
  rooms.apply(room, "clientB", [{ t: "slide", sid: "s1", slide: { ...start, elements: [...start.elements, { id: "e", kind: "shape", shape: "ellipse" }] }, base: h(start) }]);
  rooms.apply(room, "clientA", [{ t: "slide", sid: "s1", slide: { ...start, elements: [{ id: "t", kind: "text", text: "メモ（追記）" }] }, base: h(start) }]);
  const els = room.deck.slides[0].elements;
  assert.deepEqual(els.map((o) => o.id), ["t", "e"], "both: A's text and B's ellipse");
  assert.equal(els[0].text, "メモ（追記）");
  // A hears the merged slide too (it only had its own), B hears it as usual.
  const toA = a.events.filter((e) => e.event === "ops").at(-1);
  assert.equal(toA.data.ops[0].merged, true);
  assert.deepEqual(toA.data.ops[0].slide.elements.map((o) => o.id), ["t", "e"]);
  assert.deepEqual(b.events.filter((e) => e.event === "ops").at(-1).data.ops[0].slide.elements.map((o) => o.id), ["t", "e"]);
  // A change made on the room's current version simply replaces it (nothing merged, nothing echoed).
  const now = JSON.parse(JSON.stringify(room.deck.slides[0]));
  const echoes = a.events.length;
  rooms.apply(room, "clientA", [{ t: "slide", sid: "s1", slide: { ...now, title: "題" }, base: h(now) }]);
  assert.equal(room.deck.slides[0].title, "題");
  assert.equal(a.events.length, echoes);
  // Orders: B moves s2 first while A adds s3 at the end, both on the first order.
  const first = ["s1", "s2"];
  rooms.apply(room, "clientB", [{ t: "order", sids: ["s2", "s1"], base: h(first) }]);
  rooms.apply(room, "clientA", [{ t: "order", sids: ["s1", "s2", "s3"], add: [{ type: "blank", sid: "s3" }], base: h(first) }]);
  // B's move kept; A's new slide follows the slide it followed for A (s2).
  assert.deepEqual(room.deck.slides.map((s) => s.sid), ["s2", "s3", "s1"]);
  assert.equal(a.events.filter((e) => e.event === "ops").at(-1).data.ops[0].merged, true, "A hears the merged order");
});

test("オンライン プレゼンテーション: only the presenter (with the key) moves the show; viewers follow and cannot edit", () => {
  const rooms = createRooms();
  const room = rooms.create(deck(), { show: true });
  assert.equal(room.kind, "show");
  assert.ok(isRoomId(room.key));
  const viewer = stream();
  rooms.join(room, viewer, { client: "viewer1", name: "視聴者" });
  const snap = viewer.events[0].data;
  assert.equal(snap.kind, "show");
  assert.deepEqual([snap.show.index, snap.show.step, snap.show.live], [0, 0, false]);
  assert.ok(!JSON.stringify(snap).includes(room.key), "viewers never see the key");
  assert.throws(() => rooms.setShow(room, "wrong-key-wrong-key", { index: 1 }), /発表者/);
  assert.throws(() => rooms.apply(room, "viewer1", [{ t: "slide", sid: "s1", slide: { sid: "s1" } }]), /変更できません/);
  rooms.setShow(room, room.key, { index: 1, step: 2 });
  const moved = viewer.events.at(-1);
  assert.equal(moved.event, "show");
  assert.deepEqual([moved.data.index, moved.data.step, moved.data.sid, moved.data.live], [1, 2, "s2", true]);
  // An index out of range falls back to the slide's id; ending the show is heard too.
  rooms.setShow(room, room.key, { index: 99, sid: "s3", ended: true });
  assert.deepEqual([viewer.events.at(-1).data.index, viewer.events.at(-1).data.live], [2, false]);
  // (the deck update below needs the room: the end comes last)
  // 資料の更新: the presenter sends a new deck; viewers get it.
  rooms.replaceDeck(room, room.key, { ...deck(), title: "更新した資料" });
  assert.equal(viewer.events.at(-1).event, "deck");
  assert.equal(viewer.events.at(-1).data.deck.title, "更新した資料");
  assert.throws(() => rooms.replaceDeck(room, null, deck()), /発表者/);
  rooms.setShow(room, room.key, { index: 0, ended: true, over: true });
  assert.equal(viewer.events.at(-1).data.over, true, "the end of the online presentation is heard");
  assert.equal(rooms.get(room.id), null, "and the room goes at once");
  // Bringing the show back after the server forgot it needs its key.
  const again = createRooms();
  assert.equal(again.create(deck(), { id: room.id, show: true, key: room.key }).key, room.key);
  assert.throws(() => again.create(deck(), { id: room.id }), /鍵/);
});

test("PowerPoint Live: reactions float to everyone (a few at a time), one vote per viewer per poll, captions from the presenter", () => {
  const rooms = createRooms();
  const pollDeck = { ...deck(), slides: [{ type: "blank", sid: "s1", elements: [{ id: "p1", kind: "poll", question: "どれがよい？", options: ["A案", "B案", "C案"] }] }] };
  const room = rooms.create(pollDeck, { show: true });
  const presenter = stream();
  const v1 = stream();
  const v2 = stream();
  rooms.join(room, presenter, { client: "presenter1", name: "発表者" });
  rooms.join(room, v1, { client: "viewer1", name: "視聴者" });
  rooms.join(room, v2, { client: "viewer2", name: "視聴者" });
  assert.equal(rooms.react(room, "viewer1", "👏"), true);
  assert.deepEqual(presenter.events.at(-1), { event: "react", data: { emoji: "👏", from: "viewer1" } });
  assert.throws(() => rooms.react(room, "viewer1", "<b>"), /リアクション/);
  for (let i = 0; i < 10; i += 1) rooms.react(room, "viewer1", "👍");
  assert.equal(presenter.events.filter((e) => e.event === "react").length, 6, "at most six every few seconds per viewer");
  // Votes: a viewer's new answer replaces their old one.
  assert.deepEqual(rooms.vote(room, "viewer1", { poll: "p1", option: 0 }), [1, 0, 0]);
  assert.deepEqual(rooms.vote(room, "viewer2", { poll: "p1", option: 2 }), [1, 0, 1]);
  assert.deepEqual(rooms.vote(room, "viewer1", { poll: "p1", option: 2 }), [0, 0, 2]);
  assert.deepEqual(presenter.events.at(-1), { event: "votes", data: { poll: "p1", counts: [0, 0, 2] } });
  assert.throws(() => rooms.vote(room, "viewer1", { poll: "p1", option: 5 }), /投票/);
  assert.throws(() => rooms.vote(room, "viewer1", { poll: "nope", option: 0 }), /投票/);
  // A viewer who joins later hears the counts so far.
  const late = stream();
  rooms.join(room, late, { client: "viewer3", name: "視聴者" });
  assert.deepEqual(late.events[0].data.votes, { p1: [0, 0, 2] });
  // Captions only from the presenter (with the key).
  rooms.caption(room, room.key, "本日は<お集まり>いただき");
  assert.deepEqual(v2.events.at(-1), { event: "caption", data: { text: "本日はお集まりいただき" } });
  assert.throws(() => rooms.caption(room, "wrong-key-wrong-key", "x"), /発表者/);
});
