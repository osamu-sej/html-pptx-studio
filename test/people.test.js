// The writer of a comment (people.mjs): initials for the avatar, a steady colour, and @mentions found in plain text.
import test from "node:test";
import assert from "node:assert/strict";
import { cleanName, colorOf, initials, mentionsIn, mentionNodes } from "../public/editor/people.mjs";

const h = (tag, attrs, ...kids) => ({ tag, cls: [attrs?.class].flat().filter(Boolean).join(" "), text: kids.join("") });

test("initials: two letters of a Western name, the first two characters of a Japanese one", () => {
  assert.equal(initials("Taro Yamada"), "TY");
  assert.equal(initials("sato"), "S");
  assert.equal(initials("山田 太郎"), "山田");
  assert.equal(initials(""), "?");
  assert.equal(cleanName("  <b>Ann</b>\n "), "bAnn/b");
  assert.equal(colorOf("山田 太郎"), colorOf("山田 太郎"));
});

test("@mentions: known names with spaces, other words, not e-mail addresses", () => {
  const known = ["山田 太郎", "佐藤"];
  assert.deepEqual(mentionsIn("@山田 太郎 表の単位は？ @佐藤、確認を", known), ["山田 太郎", "佐藤"]);
  assert.deepEqual(mentionsIn("mail me at a@b.jp", known), []);
  assert.deepEqual(mentionsIn("cc @kim please", known), ["kim"]);
  const nodes = mentionNodes(h, "@山田 太郎 確認", known, "山田 太郎");
  assert.equal(nodes[0].text, "@山田 太郎");
  assert.equal(nodes[0].cls, "pp-mention me");
  assert.equal(nodes[1], " 確認");
});
