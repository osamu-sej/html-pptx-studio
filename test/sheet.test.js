// Excel スプレッドシート (public/editor/sheet.mjs): formulas in a table's cells, worked out into the words the cell shows;
// moving addresses when a formula is copied or rows and columns come and go; CSV; and the engine's normalization.
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import vm from "node:vm";
import test from "node:test";
import { parseHTML } from "linkedom";

import * as S from "../public/editor/sheet.mjs";
import * as ops from "../public/editor/ops.mjs";

const root = fileURLToPath(new URL("..", import.meta.url));
async function loadEngine() {
  const { window } = parseHTML("<!doctype html><html><head></head><body></body></html>");
  const icons = (await readFile(join(root, "public", "engine", "icons.json"), "utf8")).trim();
  const context = vm.createContext(window);
  vm.runInContext((await readFile(join(root, "public", "engine", "engine.js"), "utf8")).replace("/*__ICONS__*/{}", () => icons), context, { filename: "engine.js" });
  vm.runInContext(await readFile(join(root, "public", "engine", "objects.js"), "utf8"), context, { filename: "objects.js" });
  return window.SlideEngine;
}

/** A sheet table from rows of words ("=…" is a formula). */
const sheetOf = (rows) => ({ kind: "table", sheet: true, x: 0, y: 0, w: 800, h: 300, cols: rows[0].map(() => 1 / rows[0].length), rows: rows.map(() => 1 / rows.length), cells: S.cellsFromWords(rows) });
/** What every cell of the table shows, as plain words. */
const shown = (table) => S.recalc(table).cells.map((row) => row.map((cell) => S.plainOf(cell.text)));

test("addresses: A…Z, AA…, and $ marks", () => {
  assert.equal(S.colName(0), "A");
  assert.equal(S.colName(25), "Z");
  assert.equal(S.colName(26), "AA");
  assert.equal(S.colName(99), "CV");
  assert.equal(S.colIndex("AA"), 26);
  assert.equal(S.cellName(2, 1), "B3");
  assert.deepEqual(S.parseAddress("$B$3"), { r: 2, c: 1, ra: true, ca: true });
  assert.deepEqual(S.parseAddress("c10"), { r: 9, c: 2, ra: false, ca: false });
  assert.equal(S.parseAddress("A0"), null);
  assert.equal(S.parseAddress("12"), null);
});

test("parseNumber: Japanese ways of writing a number", () => {
  assert.equal(S.parseNumber("1,200"), 1200);
  assert.equal(S.parseNumber("¥1,200"), 1200);
  assert.equal(S.parseNumber("￥１，２００"), 1200, "full width");
  assert.equal(S.parseNumber("12%"), 0.12);
  assert.equal(S.parseNumber("△1,200"), -1200);
  assert.equal(S.parseNumber("▲3"), -3);
  assert.equal(S.parseNumber("(5)"), -5);
  assert.equal(S.parseNumber("-2.5"), -2.5);
  assert.equal(S.parseNumber("5円"), 5);
  assert.equal(S.parseNumber("1e3"), 1000);
  assert.equal(S.parseNumber("2024/4/1"), null);
  assert.equal(S.parseNumber("03-1234"), null);
  assert.equal(S.parseNumber("1,2"), null);
  assert.equal(S.parseNumber(""), null);
  assert.equal(S.parseNumber("-"), null);
});

test("formatNumber: digits, commas, percent and a currency sign", () => {
  assert.equal(S.formatNumber(0.1 + 0.2), "0.3", "no floating-point noise");
  assert.equal(S.formatNumber(1234567.891, { sep: true }), "1,234,567.891");
  assert.equal(S.formatNumber(2.5, { dec: 0 }), "3", "half away from zero");
  assert.equal(S.formatNumber(-2.5, { dec: 0 }), "-3");
  assert.equal(S.formatNumber(1.005, { dec: 2 }), "1.01");
  assert.equal(S.formatNumber(1234.5, { dec: 2, sep: true, cur: "¥" }), "¥1,234.50");
  assert.equal(S.formatNumber(0.256, { pct: true, dec: 1 }), "25.6%");
  assert.equal(S.formatNumber(-0.0001, { dec: 2 }), "0.00", "no minus zero");
  assert.equal(S.formatNumber(-1234, { sep: true }), "-1,234");
});

test("formulas: arithmetic, precedence and the way Excel groups them", () => {
  const t = sheetOf([["=1+2*3", "=(1+2)*3", "=2^3^2", "=-2^2", "=10/4", "=50%*8"], ["=\"a\"&\"b\"&1", "=1<2", "=\"abc\"=\"ABC\"", "=3<>3", "=TRUE", "=2>=2"]]);
  assert.deepEqual(shown(t), [["7", "9", "64", "4", "2.5", "4"], ["ab1", "TRUE", "TRUE", "FALSE", "TRUE", "TRUE"]]);
});

test("formulas: cell addresses and ranges, in any case and full width", () => {
  const t = sheetOf([["10", "20", "=a1+b1", "=ＳＵＭ（Ａ１：Ｂ１）"], ["3,000", "▲5", "=SUM(A1:B2)", "=SUM($A$1:B$2)"], ["", "", "=A3", "=sum(a1,b1,5)"]]);
  const out = S.recalc(t);
  assert.deepEqual(shown(t), [["10", "20", "30", "30"], ["3,000", "▲5", "3025", "3025"], ["", "", "0", "35"]]);
  assert.equal(out.cells[0][3].f, "=SUM(A1:B1)", "kept half-width and upper case");
});

test("functions: the ones a slide's table needs", () => {
  const t = sheetOf([
    ["4", "8", "15", "=SUM(A1:C1)", "=AVERAGE(A1:C1)", "=MIN(A1:C1)", "=MAX(A1:C1)", "=COUNT(A1:C1,\"x\")"],
    ["=ROUND(2.567,2)", "=ROUNDUP(2.341,1)", "=ROUNDDOWN(-2.399,1)", "=INT(-2.5)", "=ABS(-3)", "=SQRT(16)", "=POWER(2,10)", "=MOD(-7,3)"],
    ["=IF(A1>3,\"高\",\"低\")", "=IF(A1>9,1)", "=AND(1,0)", "=OR(0,1)", "=NOT(0)", "=IFERROR(1/0,\"n/a\")", "=MEDIAN(1,2,10,20)", "=PRODUCT(2,3,4)"],
    ["=LEN(\"売上高\")", "=LEFT(\"売上高\",2)", "=RIGHT(\"売上高\",1)", "=MID(\"売上高ABC\",3,2)", "=UPPER(\"abc\")", "=TRIM(\"  a   b \")", "=VALUE(\"1,200\")+1", "=CONCAT(\"a\",1,\"b\")"],
    ["=LARGE(A1:C1,2)", "=SMALL(A1:C1,1)", "=SUMPRODUCT(A1:C1,A1:C1)", "=COUNTA(A1:C1,D4)", "=COUNTBLANK(A6:C6)", "=ISNUMBER(A1)", "=ISBLANK(B6)", "=ISERROR(1/0)"],
    ["", "", "", "", "", "", "", ""],
  ]);
  assert.deepEqual(shown(t), [
    ["4", "8", "15", "27", "9", "4", "15", "3"],
    ["2.57", "2.4", "-2.3", "-3", "3", "4", "1024", "2"],
    ["高", "FALSE", "FALSE", "TRUE", "TRUE", "n/a", "6", "24"],
    ["3", "売上", "高", "高A", "ABC", "a b", "1201", "a1b"],
    ["8", "4", "305", "4", "3", "TRUE", "TRUE", "TRUE"],
    ["", "", "", "", "", "", "", ""],
  ]);
});

test("functions: conditions (SUMIF, COUNTIF, AVERAGEIF, SUMIFS, COUNTIFS) and lookups", () => {
  const t = sheetOf([
    ["東京", "100", "A"], ["大阪", "250", "B"], ["東京", "50", "B"], ["名古屋", "80", "A"],
    ["=SUMIF(A1:A4,\"東京\",B1:B4)", "=COUNTIF(B1:B4,\">=80\")", "=AVERAGEIF(C1:C4,\"A\",B1:B4)", "=SUMIFS(B1:B4,A1:A4,\"東京\",C1:C4,\"B\")"],
    ["=COUNTIFS(A1:A4,\"東*\",C1:C4,\"B\")", "=SUMIF(B1:B4,\"<100\")", "=COUNTIF(A1:A4,\"<>東京\")", "=COUNTIF(A1:A4,\"?阪\")"],
    ["=VLOOKUP(\"大阪\",A1:C4,2,FALSE)", "=VLOOKUP(\"京都\",A1:C4,2,FALSE)", "=MATCH(\"名古屋\",A1:A4,0)", "=INDEX(B1:B4,2)"],
    ["=VLOOKUP(120,B1:C4,2)", "=INDEX(A1:C4,3,1)", "=IFERROR(VLOOKUP(\"京都\",A1:C4,2,FALSE),\"なし\")", "=ROWS(A1:C4)*COLUMNS(A1:C4)"],
  ]);
  const out = shown(t);
  assert.deepEqual(out[4], ["150", "3", "90", "50"]);
  assert.deepEqual(out[5], ["1", "130", "2", "1"]);
  assert.deepEqual(out[6], ["250", "#N/A", "4", "250"]);
  assert.equal(out[7][1], "東京");
  assert.equal(out[7][2], "なし");
  assert.equal(out[7][3], "12");
});

test("errors: what each one says, and that they spread", () => {
  const t = sheetOf([["=1/0", "=A1+1", "=UNKNOWN(1)", "=1+", "=SQRT(-1)", "=\"a\"+1", "=#REF!+1", "=IFERROR(A1,0)"], ["=B2", "=A2", "=SUM(A1:B1)", "=COUNT(A1:B1)", "=NOSUCH", "=1,2", "=(1", "=A1:B1"]]);
  assert.deepEqual(shown(t), [
    ["#DIV/0!", "#DIV/0!", "#NAME?", "#ERROR!", "#NUM!", "#VALUE!", "#REF!", "0"],
    ["#CIRC!", "#CIRC!", "#DIV/0!", "0", "#NAME?", "#ERROR!", "#ERROR!", "#VALUE!"],
  ]);
  assert.ok(S.ERROR_HINTS["#DIV/0!"] && S.ERROR_HINTS["#CIRC!"] && S.ERROR_HINTS["#NAME?"]);
});

test("recalc: formula cells show their value, a plain table or a sheet without formulas comes back unchanged", () => {
  const plain = { kind: "table", cells: [[{ text: "<p>=1+1</p>" }]] };
  assert.equal(S.recalc(plain), plain, "not a sheet: nothing happens");
  const bare = sheetOf([["1", "2"]]);
  assert.equal(S.recalc(bare), bare, "no formulas");
  const t = sheetOf([["2", "=A1*3"]]);
  const out = S.recalc(t);
  assert.notEqual(out, t);
  assert.equal(out.cells[0][1].text, "<p>6</p>");
  assert.equal(t.cells[0][1].text, undefined, "the table given is not changed");
  assert.equal(S.recalc(out), out, "worked out already");
  // A value that came to be empty text (a result of "") clears the words.
  assert.equal(S.recalc(sheetOf([["=\"\""]])).cells[0][0].text, undefined);
});

test("recalc: a number cell's settings shape the result (桁数・桁区切り・%・通貨)", () => {
  const t = sheetOf([["1234.567", "=A1", "=A1", "=A1/10000", "=A1"]]);
  t.cells[0][1].dec = 1; t.cells[0][1].sep = true;
  t.cells[0][2].cur = "¥"; t.cells[0][2].dec = 0; t.cells[0][2].sep = true;
  t.cells[0][3].pct = true; t.cells[0][3].dec = 1;
  assert.deepEqual(shown(t)[0], ["1234.567", "1,234.6", "¥1,235", "12.3%", "1234.567"]);
  assert.deepEqual(S.numberSettings({ dec: 9, sep: true, cur: "€", pct: 1 }), { sep: true }, "only what the screen offers");
  assert.deepEqual(S.numberSettings({ dec: "2", cur: "$", pct: true }), { dec: 2, pct: true, cur: "$" });
});

test("recalc: a formula this file cannot work out keeps the words an imported sheet came with", () => {
  const t = sheetOf([["10", "=XLOOKUP(A1,A1,A1)", "=B1*2", "=Sheet2!A1"]]);
  t.cells[0][1].text = "<p>1,500</p>";
  t.cells[0][3].text = "<p>7</p>";
  const out = S.recalc(t);
  assert.equal(S.plainOf(out.cells[0][1].text), "1,500", "kept as it was saved");
  assert.equal(S.plainOf(out.cells[0][2].text), "3000", "and used by the cells that read it");
  assert.equal(S.plainOf(out.cells[0][3].text), "7");
  // The same formula typed by hand has no saved words: the error shows.
  assert.equal(shown(sheetOf([["=XLOOKUP(1,2,3)"]]))[0][0], "#NAME?");
});

test("formulas move: copying down and across, with $ holding an address in place", () => {
  assert.equal(S.shiftFormula("=A1+B$2+$C3+$D$4", 1, 1), "=B2+C$2+$C4+$D$4");
  assert.equal(S.shiftFormula("=SUM(A1:A3)", 2, 0), "=SUM(A3:A5)");
  assert.equal(S.shiftFormula("=A1", -1, 0), "=#REF!", "off the top of the sheet");
  assert.equal(S.shiftFormula("=\"A1\"&A1", 1, 0), "=\"A1\"&A2", "words in quotes are not addresses");
  assert.equal(S.shiftFormula("=LOG10(A1)+1E5", 1, 0), "=LOG10(A2)+1E5", "LOG10 and 1E5 are not addresses");
  const cells = S.cellsFromWords([["1", "2"], ["=A1*10", ""], ["", ""]]);
  const down = S.fillBlock(cells, 1, 0, 2, 0, "down");
  assert.equal(down[2][0].f, "=A2*10");
  const right = S.fillBlock(S.cellsFromWords([["=A2", "", ""], ["5", "", ""]]), 0, 0, 0, 2, "right");
  assert.deepEqual([right[0][1].f, right[0][2].f], ["=B2", "=C2"]);
});

test("formulas move: rows and columns that come and go", () => {
  assert.equal(S.insertRefs("=A1+A5+SUM(A1:A5)", "row", 2), "=A1+A6+SUM(A1:A6)", "a range grows when rows go inside it");
  assert.equal(S.insertRefs("=SUM(A1:A5)", "row", 5), "=SUM(A1:A5)", "but not after its last row");
  assert.equal(S.insertRefs("=B1+$C$1", "col", 1, 2), "=D1+$E$1");
  assert.equal(S.deleteRefs("=A1+A5+A9", "row", 1, 3), "=A1+A2+A6");
  assert.equal(S.deleteRefs("=A1+A3", "row", 2, 2), "=A1+#REF!");
  assert.equal(S.deleteRefs("=SUM(A2:A6)", "row", 2, 3), "=SUM(A2:A4)", "ranges shrink");
  assert.equal(S.deleteRefs("=SUM(A2:A3)", "row", 1, 4), "=SUM(#REF!)", "a range that is gone");
  assert.equal(S.deleteRefs("=SUM(A1:C1)", "col", 1, 1), "=SUM(A1:B1)");
  assert.deepEqual(S.formulaCells("=SUM(A1:B2)+C3").map(([r, c]) => S.cellName(r, c)), ["A1", "B1", "A2", "B2", "C3"]);
});

test("the table's own rows and columns carry their formulas along (ops)", () => {
  const t = sheetOf([["1", "2", "3"], ["4", "5", "=SUM(A1:C2)"], ["", "=A2", ""]]);
  const withRow = ops.tableInsertRow(t, 1);
  assert.equal(withRow.cells[2][2].f, "=SUM(A1:C3)");
  assert.equal(withRow.cells[3][1].f, "=A3");
  const withCol = ops.tableInsertCol(t, 0);
  assert.equal(withCol.cells[1][3].f, "=SUM(B1:D2)");
  const noRow = ops.tableDeleteRows(t, 0, 0);
  assert.equal(noRow.cells[0][2].f, "=SUM(A1:C1)");
  assert.equal(noRow.cells[1][1].f, "=A1", "the cell it reads moved up");
  assert.equal(ops.tableDeleteRows(t, 1, 1).cells[1][1].f, "=#REF!", "it read the row that went");
  const noCol = ops.tableDeleteCols(t, 0, 0);
  assert.equal(noCol.cells[2][0].f, "=#REF!", "the formula that read the deleted column");
  assert.equal(noCol.cells[1][1].f, "=SUM(#REF!:B2)".replace("#REF!:B2", "A1:B2"), "a range shrinks to what is left");
  assert.equal(S.recalc(noCol).cells[2][0].text, "<p>#REF!</p>");
  // A table that is not a sheet leaves its words alone.
  const plain = { ...t, sheet: undefined };
  assert.equal(ops.tableInsertRow(plain, 0).cells[2][2].f, "=SUM(A1:C2)");
});

test("typing: = starts a formula in a sheet, anything else is words; editing shows the formula", () => {
  const typed = S.enterText({ text: "<p>old</p>", bold: true }, "=SUM(A1:A3)");
  assert.deepEqual(typed, { bold: true, f: "=SUM(A1:A3)" }, "the old words go: the formula decides what shows");
  assert.equal(S.editText({ f: "=A1", text: "<p>5</p>" }), "=A1");
  assert.equal(S.editText({ text: "<p>売上 &amp; 利益</p>" }), "売上 & 利益");
  assert.deepEqual(S.enterText({ f: "=A1" }, "売上"), { text: "<p>売上</p>" }, "typing words over a formula takes the formula away");
  assert.deepEqual(S.enterText({}, "=1+1", { sheet: false }), { text: "<p>=1+1</p>" }, "not a sheet: just words");
  assert.deepEqual(S.enterText({ text: "<p>x</p>" }, "  "), {}, "empty");
  assert.equal(S.enterText({}, "＝ａ１＋１").f, "=A1+1", "full-width formula");
  assert.equal(S.enterText({}, "=").f, undefined, "= alone is words");
});

test("autoSum: the numbers above, else the numbers to the left", () => {
  const cells = S.cellsFromWords([["売上", "10"], ["a", "20"], ["b", "30"], ["合計", ""]]);
  assert.equal(S.autoSum(cells, 3, 1), "=SUM(B1:B3)");
  const side = S.cellsFromWords([["x", "1", "2", ""]]);
  assert.equal(S.autoSum(side, 0, 3), "=SUM(B1:C1)");
  assert.equal(S.autoSum(S.cellsFromWords([["a", ""]]), 0, 1), "");
});

test("CSV: quotes, tabs and semicolons, a BOM, CRLF, and a leading = that Excel would run", () => {
  assert.deepEqual(S.parseCsv("a,b\r\n\"x,1\",\"y\"\"z\"\r\n"), [["a", "b"], ["x,1", 'y"z']]);
  assert.deepEqual(S.parseCsv("﻿項目\t値\n売上\t100"), [["項目", "値"], ["売上", "100"]]);
  assert.deepEqual(S.parseCsv("a;b\nc;d"), [["a", "b"], ["c", "d"]]);
  assert.deepEqual(S.parseCsv("a,\"1\n2\"\n"), [["a", "1\n2"]], "a quoted line break stays in the cell");
  assert.deepEqual(S.parseCsv("a,b\n\n\n"), [["a", "b"]]);
  assert.equal(S.toCsv([["a", "b,c"], ["=1+1", "-5"], ["@x", 'q"q']]), "a,\"b,c\"\r\n'=1+1,-5\r\n'@x,\"q\"\"q\"");
  const t = sheetOf([["x", "=1+1"]]);
  assert.deepEqual(S.tableWords(S.recalc(t)), [["x", "2"]]);
  const cells = S.cellsFromWords([["a", "=1+1"], ["b"]]);
  assert.equal(cells[0][1].f, "=1+1");
  assert.equal(cells[1].length, 2, "short rows are filled out");
});

test("the engine keeps formulas and works them out once the editor has registered the sheet", async () => {
  const E = await loadEngine();
  E.sheetCalc = (table) => S.recalc(table, { inPlace: true });
  const raw = { kind: "table", sheet: true, x: 0, y: 0, w: 600, h: 200, cols: [0.5, 0.5], rows: [1], cells: [[{ text: "<p>4</p>" }, { f: " =a1*2 ", text: "<p>stale</p>", dec: 1, sep: true, cur: "¥" }]] };
  const o = E.normalizeObject(raw);
  assert.equal(o.sheet, true);
  assert.equal(o.cells[0][1].f, "=A1*2");
  assert.equal(o.cells[0][1].text, "<p>¥8.0</p>");
  assert.deepEqual([o.cells[0][1].dec, o.cells[0][1].sep, o.cells[0][1].cur], [1, true, "¥"]);
  const again = E.normalizeObject(o);
  assert.deepEqual(again.cells, o.cells, "normalizing twice changes nothing");
  // Not a sheet: the formula is dropped, the words stay.
  const plain = E.normalizeObject({ ...raw, sheet: undefined });
  assert.equal(plain.sheet, undefined);
  assert.equal(plain.cells[0][1].f, undefined);
  // Junk is not kept.
  const junk = E.normalizeObject({ ...raw, cells: [[{ f: "SUM(1)" }, { f: "=" + "1+".repeat(300) + "1", dec: 99, cur: "€", pct: "yes" }]] });
  assert.equal(junk.cells[0][0].f, undefined, "a formula starts with =");
  assert.equal(junk.cells[0][1].f, undefined, "and is short");
  assert.equal(junk.cells[0][1].dec, undefined);
  assert.equal(junk.cells[0][1].cur, undefined);
  assert.equal(junk.cells[0][1].pct, undefined);
});

test("number settings (ribbon): 桁数・桁区切り・％・¥ on a formula's result and on a typed number", () => {
  const apply = (cell, op) => { const patch = S.restyleNumber(cell, op); return patch && Object.fromEntries(Object.entries({ ...cell, ...patch }).filter(([, v]) => v !== undefined)); };
  // A formula: the settings are kept on the cell, the words are worked out again.
  const f = { f: "=A1/3", text: "<p>1234.5</p>" };
  assert.deepEqual(apply(f, "decMore"), { f: "=A1/3", text: "<p>1234.5</p>", dec: 2 }, "one more than the 1 shown");
  assert.deepEqual(apply({ ...f, dec: 2 }, "decLess"), { f: "=A1/3", text: "<p>1234.5</p>", dec: 1 });
  assert.equal(apply({ ...f, dec: 0 }, "decLess").dec, 0, "never below 0");
  assert.deepEqual(apply({ ...f, dec: 1 }, "sep"), { f: "=A1/3", text: "<p>1234.5</p>", dec: 1, sep: true });
  assert.equal(apply({ ...f, dec: 1, sep: true }, "sep").sep, undefined, "again takes it off");
  assert.equal(apply(f, "yen").cur, "¥");
  assert.equal(apply({ ...f, cur: "¥" }, "yen").cur, undefined);
  assert.equal(apply(f, "pct").pct, true);
  assert.equal(apply({ f: "=\"a\"&\"b\"", text: "<p>ab</p>" }, "pct"), null, "words have no number format");
  // A number typed in is rewritten as it is to be shown.
  const n = (text) => ({ text: `<p>${text}</p>` });
  const show = (cell, op) => S.plainOf(S.restyleNumber(cell, op).text);
  assert.equal(show(n("1200"), "sep"), "1,200");
  assert.equal(show(n("1,200"), "sep"), "1200", "off again");
  assert.equal(show(n("1,234.5"), "decMore"), "1,234.50");
  assert.equal(show(n("1,234.56"), "decLess"), "1,234.6");
  assert.equal(show(n("0.256"), "pct"), "26%", "whole percents, as in Excel");
  assert.equal(show(n("26%"), "pct"), "0.26", "and off again");
  assert.equal(show(n("¥1,200"), "yen"), "1,200");
  assert.equal(show(n("1200"), "yen"), "¥1200");
  assert.equal(S.restyleNumber(n("売上"), "sep"), null, "words stay as they are");
  assert.equal(S.restyleNumber({}, "sep"), null);
});

test("typing as a patch for ops.tableCells: what the typed words take away is set to undefined", () => {
  assert.deepEqual(S.cellPatch({ f: "=A1", dec: 2, bold: true }, "売上"), { dec: 2, bold: true, text: "<p>売上</p>", f: undefined });
  const t = sheetOf([["1", "=A1*2"]]);
  const next = ops.tableCells(t, 0, 1, 0, 1, (cell) => S.cellPatch(cell, "=A1*3"));
  assert.equal(next.cells[0][1].f, "=A1*3");
  assert.equal(next.cells[0][1].text, undefined, "the old value goes until it is worked out again");
  assert.equal(S.plainOf(S.recalc(next).cells[0][1].text), "3");
  const words = ops.tableCells(next, 0, 1, 0, 1, (cell) => S.cellPatch(cell, "メモ"));
  assert.equal(words.cells[0][1].f, undefined);
  assert.equal(S.plainOf(words.cells[0][1].text), "メモ");
});

test("並べ替え: numbers, then words, blanks last; formulas move with their row; the header and totals rows stay", () => {
  const rows = [["店", "売上", "倍"], ["C店", "30", "=B2*2"], ["A店", "10", "=B3*2"], ["", "", ""], ["B店", "20", "=B5*2"], ["合計", "=SUM(B2:B5)", ""]];
  const table = { ...sheetOf(rows), header: true, lastRow: true };
  assert.deepEqual(S.sortSpan(table), [1, 4], "under the header, above the totals");
  const asc = { ...table, cells: S.sortRows(table.cells, { col: 1, r0: 1, r1: 4 }) };
  assert.deepEqual(shown(asc).map((r) => r.join("|")), ["店|売上|倍", "A店|10|20", "B店|20|40", "C店|30|60", "||", "合計|60|"], "numbers up, the blank row last, the sum still covers them");
  assert.equal(asc.cells[1][2].f, "=B2*2", "a formula follows its row: B3 → B2");
  const desc = { ...table, cells: S.sortRows(table.cells, { col: 1, desc: true, r0: 1, r1: 4 }) };
  assert.deepEqual(shown(desc).map((r) => r[0]), ["店", "C店", "B店", "A店", "", "合計"], "blanks stay last in 降順 too");
  // Words: by the column of words, in Japanese order; the same order again is "nothing to do".
  const byName = S.sortRows(asc.cells, { col: 0, r0: 1, r1: 4 });
  assert.deepEqual(byName.slice(1, 5).map((r) => S.plainOf(r[0].text)), ["A店", "B店", "C店", ""]);
  assert.equal(S.sortRows(asc.cells, { col: 1, r0: 1, r1: 4 }), asc.cells, "already in order: the same cells back");
  // Mixed: numbers before words before TRUE; equal keys keep their order.
  const mixed = sheetOf([["x", "k"], ["a", "b"], ["b", "3"], ["c", "a"], ["d", "3"], ["e", "=1>0"]]);
  const sorted = S.sortRows(mixed.cells, { col: 1, r0: 1, r1: 5 });
  assert.deepEqual(sorted.slice(1).map((r) => S.plainOf(r[0].text)), ["b", "d", "c", "a", "e"], "3, 3 (stable), the word 'a' before 'b', then TRUE");
  // $ stays; a reference to a row that moves is rewritten like a copy.
  const abs = sheetOf([["k", "v"], ["2", "=$B$1"], ["1", "=B3+1"]]);
  const moved = S.sortRows(abs.cells, { col: 0, r0: 1, r1: 2 });
  assert.equal(moved[1][1].f, "=B2+1", "the relative one moved up a row");
  assert.equal(moved[2][1].f, "=$B$1", "the absolute one did not change");
  // Merged cells in the way, or one row only: nothing is done.
  const merged = sheetOf([["a", "b"], ["1", "x"], ["2", ""]]);
  merged.cells[1][1] = { ...merged.cells[1][1], rs: 2 };
  merged.cells[2][1] = { merged: true };
  assert.equal(S.sortRows(merged.cells, { col: 0, r0: 1, r1: 2, desc: true }), null);
  assert.equal(S.sortRows(table.cells, { col: 1, r0: 2, r1: 2 }), null);
  assert.equal(S.sortRows(table.cells, { col: 1, r0: 1, r1: 99 }), null);
});

test("フィルター: the rows that fail the condition are hidden; the header and totals rows stay; merged cells stop it", async () => {
  const rows = [["店", "売上"], ["A店", "10"], ["B店", "1,200"], ["C店", "30"], ["", ""], ["合計", "=SUM(B2:B5)"]];
  const base = { ...sheetOf(rows), header: true, lastRow: true };
  const hide = (filter) => S.filterRows({ ...base, cells: S.recalc(base).cells, filter });
  assert.deepEqual(hide({ col: 1, op: "gt", value: "15" }), [1, 4], "売上 > 15: 10 and the blank row go (1,200 is a number)");
  assert.deepEqual(hide({ col: 1, op: "le", value: "30" }), [2, 4]);
  assert.deepEqual(hide({ col: 0, op: "eq", value: "B店" }), [1, 3, 4], "words: equal means the same words");
  assert.deepEqual(hide({ col: 0, op: "ne", value: "B店" }), [2], "not equal: a blank is not B店");
  assert.deepEqual(hide({ col: 0, op: "contains", value: "店" }), [4], "contains: the blank row goes");
  assert.deepEqual(hide({ col: 1, op: "blank" }), [1, 2, 3]);
  assert.deepEqual(hide({ col: 1, op: "nonblank" }), [4]);
  assert.deepEqual(hide({ col: 1, op: "nonsense" }), [], "an unknown condition hides nothing");
  assert.deepEqual(hide(undefined), []);
  const merged = { ...base, filter: { col: 1, op: "gt", value: "15" } };
  merged.cells = base.cells.map((row) => row.map((cell) => ({ ...cell })));
  merged.cells[1][0].rs = 2;
  merged.cells[2][0] = { merged: true };
  assert.deepEqual(S.filterRows(merged), [], "a table with merged cells is not filtered");
  // applyFilter writes `hide` in place and takes it away again.
  const t = { ...base, cells: S.recalc(base).cells, filter: { col: 1, op: "gt", value: "15" } };
  S.applyFilter(t);
  assert.deepEqual([...t.hide], [1, 4]);
  delete t.filter;
  S.applyFilter(t);
  assert.equal(t.hide, undefined);
});

test("フィルター in the engine: `filter` is kept for a sheet, the hidden rows are not drawn, and the rows left fill the height", async () => {
  const E = await loadEngine();
  E.sheetCalc = (table) => { S.recalc(table, { inPlace: true }); S.applyFilter(table); };
  const rows = [["店", "売上"], ["A店", "10"], ["B店", "20"], ["C店", "30"]];
  const draw = (filter, extra = {}) => {
    const o = E.normalizeObject({ id: "t", kind: "table", sheet: true, x: 0, y: 0, w: 800, h: 400, cols: [0.5, 0.5], rows: [0.25, 0.25, 0.25, 0.25], cells: rows.map((r) => r.map((text) => ({ text: `<p>${text}</p>` }))), header: true, filter, ...extra });
    return { o };
  };
  const { o } = draw({ col: 1, op: "gt", value: "15" });
  assert.deepEqual(plain(o.filter), { col: 1, op: "gt", value: "15" });
  assert.deepEqual(plain(o.hide), [1], "A店 (10) is hidden");
  // Not a sheet, or a filter that makes no sense: nothing kept.
  assert.equal(draw({ col: 9, op: "gt", value: "1" }).o.filter, undefined, "a column that is not there");
  assert.equal(draw({ col: 0, op: "xx" }).o.filter, undefined, "an unknown condition");
  assert.equal(draw(undefined).o.hide, undefined);
  assert.equal(E.normalizeObject({ id: "u", kind: "table", x: 0, y: 0, w: 400, h: 200, cols: [1], rows: [1], cells: [[{ text: "<p>a</p>" }]], filter: { col: 0, op: "eq", value: "a" } }).filter, undefined, "a plain table has no filter");
  // A saved `hide` is kept as it was when nothing works the rows out (an exported show), and only valid rows.
  delete E.sheetCalc;
  const saved = E.normalizeObject({ id: "s", kind: "table", sheet: true, x: 0, y: 0, w: 800, h: 400, cols: [1], rows: [0.5, 0.5], cells: [[{ text: "<p>a</p>" }], [{ text: "<p>b</p>" }]], filter: { col: 0, op: "eq", value: "a" }, hide: [1, 1, 7, -1, "x"] });
  assert.deepEqual(plain(saved.hide), [1]);
  // Drawn: the table with the hidden row left out, the others sharing the height.
  const slide = { type: "blank", elements: [{ ...saved, hide: [1] }] };
  const el = E.render(slide, { mode: "present", index: 1, deck: { slides: [slide], theme: "clarity" } });
  const trs = [...el.querySelectorAll(".hs-otable tr")];
  assert.equal(trs.length, 1, "one row left");
  assert.equal(trs[0].getAttribute("data-r"), "0", "the rows keep their own numbers");
  assert.equal(trs[0].style.height, "400px", "and the whole height");
});
const plain = (v) => JSON.parse(JSON.stringify(v));
