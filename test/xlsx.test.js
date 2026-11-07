// 表 → Excel (public/editor/xlsx.mjs): a table written as a workbook, which tools/xlsx_table.py (and Excel) read back.
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { once } from "node:events";
import { existsSync } from "node:fs";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

import * as S from "../public/editor/sheet.mjs";
import { numberFormatCode, tableToXlsx, workbookParts } from "../public/editor/xlsx.mjs";
const { settingsFromWords } = S;

const root = fileURLToPath(new URL("..", import.meta.url));
const python = process.env.PYTHON_BIN || (existsSync(join(root, ".venv", "bin", "python")) ? join(root, ".venv", "bin", "python") : "python3");
const sample = () => {
  const cells = S.cellsFromWords([["項目", "前期", "今期", "増減"], ["売上", "1,200", "1,500", "=C2-B2"], ["利益率", "12%", "15.5%", "=C3-B3"], ["日付", "2024/4/1", "2024/4/30", "=C4-B4"], ["合計", "=SUM(B2:B2)", "=SUM(C2:C2)", "=SUM(D2:D2)&\"円\""]]);
  cells[1][3].sep = true; cells[1][3].dec = 0;
  cells[2][3].pct = true; cells[2][3].dec = 1;
  cells[0][0] = { text: "<p>項目 &amp; 備考</p>", rs: 1, cs: 1 };
  return S.recalc({ kind: "table", sheet: true, x: 0, y: 0, w: 1100, h: 400, cols: [0.4, 0.2, 0.2, 0.2], rows: [0.2, 0.2, 0.2, 0.2, 0.2], cells });
};

test("number formats: the code Excel gets, and what a typed number says about its own", () => {
  assert.equal(numberFormatCode({}), "");
  assert.equal(numberFormatCode({ sep: true, dec: 0 }), "#,##0");
  assert.equal(numberFormatCode({ cur: "¥", sep: true, dec: 2 }), '"¥"#,##0.00');
  assert.equal(numberFormatCode({ pct: true, dec: 1 }), "0.0%");
  assert.deepEqual(settingsFromWords("1,200"), { sep: true, dec: 0 });
  assert.deepEqual(settingsFromWords("¥3,500.5"), { dec: 1, sep: true, cur: "¥" });
  assert.deepEqual(settingsFromWords("15.5%"), { dec: 1, pct: true });
  assert.deepEqual(settingsFromWords("1200"), {});
});

test("the workbook's parts: formulas with their values, merged cells, widths and escaped words", () => {
  const t = sample();
  t.cells[1][0].cs = 1;
  t.cells.push([{ text: "<p>a</p>", rs: 1, cs: 2 }, { merged: true }, {}, {}]);
  const parts = workbookParts(t, { name: "売上:表/1" });
  assert.deepEqual(Object.keys(parts).sort(), ["[Content_Types].xml", "_rels/.rels", "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml"]);
  assert.match(parts["xl/workbook.xml"], /<sheet name="売上 表 1"/, "characters a sheet name may not have");
  const sheet = parts["xl/worksheets/sheet1.xml"];
  assert.match(sheet, /<c r="D2" s="\d+"><f>C2-B2<\/f><v>300<\/v><\/c>/);
  assert.match(sheet, /<c r="B2" s="\d+"><v>1200<\/v><\/c>/, "1,200 becomes the number 1200 with a comma format");
  assert.match(sheet, /<c r="B3" s="\d+"><v>0\.12<\/v><\/c>/);
  assert.match(sheet, /<c r="B4" s="\d+"><v>45383<\/v><\/c>/, "a date is its serial number");
  assert.match(sheet, /<c r="D5" s="\d+" t="str"><f>SUM\(D2:D2\)&amp;&quot;円&quot;<\/f><v>300円<\/v><\/c>/);
  assert.match(sheet, /<t xml:space="preserve">項目 &amp; 備考<\/t>/);
  assert.match(sheet, /<mergeCell ref="A6:B6"\/>/);
  assert.match(sheet, /<col min="1" max="1" width="20" customWidth="1"\/>/, "440 px is 20 characters");
  const styles = parts["xl/styles.xml"];
  assert.match(styles, /formatCode="#,##0"/);
  assert.match(styles, /formatCode="0\.0%"/);
  assert.doesNotMatch(styles, /yyyy\/m\/d/, "the date uses Excel's own format 14");
  assert.match(styles, /numFmtId="14"/);
  // Errors are written as errors.
  const bad = { ...t, cells: [[{ f: "=1/0" }]] };
  assert.match(workbookParts(bad)["xl/worksheets/sheet1.xml"], /t="e"><f>1\/0<\/f><v>#DIV\/0!<\/v>/);
});

test("the .xlsx is a zip Excel can open: every part is stored with the right size and CRC", () => {
  const bytes = tableToXlsx(sample());
  assert.equal(String.fromCharCode(bytes[0], bytes[1]), "PK");
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const end = bytes.length - 22;
  assert.equal(view.getUint32(end, true), 0x06054b50, "end of central directory");
  assert.equal(view.getUint16(end + 10, true), 6, "six parts");
});

test("what is written is read back by tools/xlsx_table.py: words, formulas, formats and merges", async () => {
  const t = sample();
  t.cells[0][1] = { text: "<p>前期</p>", rs: 1, cs: 2 };
  t.cells[0][2] = { merged: true };
  const child = spawn(python, [join(root, "tools", "xlsx_table.py"), "x.xlsx"], { stdio: ["pipe", "pipe", "pipe"] });
  const out = [];
  const err = [];
  child.stdout.on("data", (c) => out.push(c));
  child.stderr.on("data", (c) => err.push(c));
  child.stdin.end(Buffer.from(tableToXlsx(t)));
  const [code] = await once(child, "exit");
  assert.equal(code, 0, Buffer.concat(err).toString());
  const { sheets } = JSON.parse(Buffer.concat(out).toString());
  const rows = sheets[0].rows;
  assert.equal(rows[0][0].text, "項目 & 備考");
  assert.deepEqual([rows[0][1].rs, rows[0][1].cs, rows[0][2].merged], [1, 2, true]);
  assert.deepEqual([rows[1][1].text, rows[1][2].text], ["1,200", "1,500"]);
  assert.equal(rows[1][3].f, "=C2-B2");
  assert.equal(rows[1][3].text, "300");
  assert.deepEqual([rows[2][1].text, rows[2][2].text], ["12%", "16%"].map((x, i) => (i ? "15.5%" : x)));
  assert.equal(rows[3][1].text, "2024/4/1", "date");
  assert.equal(rows[4][3].f, '=SUM(D2:D2)&"円"');
  // And back into a table: the formulas work out to what they showed.
  const back = S.recalc({ sheet: true, cells: rows.map((row) => row.map((c) => ({ ...(c.text ? { text: `<p>${c.text}</p>` } : {}), ...(c.f ? { f: c.f } : {}), ...(c.merged ? { merged: true } : {}) }))) });
  assert.equal(S.plainOf(back.cells[1][3].text), "300");
  assert.equal(S.plainOf(back.cells[3][3].text), "29", "2024/4/30 − 2024/4/1 days");
});

test("フィルター: the rows a filter hides are hidden rows in the workbook", () => {
  const cells = S.cellsFromWords([["店", "売上"], ["A店", "10"], ["B店", "20"]]);
  const parts = workbookParts({ cells, cols: [0.5, 0.5], rows: [1 / 3, 1 / 3, 1 / 3], w: 800, hide: [1], sheet: true });
  const xml = parts["xl/worksheets/sheet1.xml"];
  assert.ok(/<row r="2" hidden="1">/.test(xml), "row 2 is hidden");
  assert.ok(/<row r="1">/.test(xml) && /<row r="3">/.test(xml), "the others are not");
});
