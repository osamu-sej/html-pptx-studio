// 表 → Excel (.xlsx): a sheet table saved as a workbook Excel opens with its formulas alive (and brought back with
// 「Excel／CSV から」). Only what a slide table has is written: the words and numbers, formulas with the value they work out to,
// number formats (桁数・桁区切り・%・¥/$), merged cells and column widths. Pure: the file's parts are returned, `zipFiles` packs them.

import { cellName, makeSheet, numberSettings, parseDate, parseNumber, plainOf, settingsFromWords, showValue, SheetError } from "./sheet.mjs";
import { zipFiles } from "./imagexport.mjs";

const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
  // The XML 1.0 characters that may not be written at all.
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "");
const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

/** The number format an Excel cell gets for the settings (a code, or "" for General). */
export function numberFormatCode({ dec, sep, pct, cur } = {}) {
  if (dec == null && !sep && !pct && !cur) return "";
  const body = (sep ? "#,##0" : "0") + (dec ? `.${"0".repeat(dec)}` : "");
  return `${cur ? `"${cur}"` : ""}${body}${pct ? "%" : ""}`;
}
/** The parts of the workbook (path → xml text) for a table (a sheet or not: a plain table has no formulas). */
export function workbookParts(table, { name = "表" } = {}) {
  const sheet = makeSheet(table.cells);
  const formats = new Map([["", 0]]);
  const formatList = [];
  const styleOf = (code, dateFormat = false) => {
    const key = dateFormat ? "#date" : code;
    if (!formats.has(key)) { formats.set(key, formats.size); formatList.push([key, code]); }
    return formats.get(key);
  };
  // フィルター: the rows it hides are hidden in Excel too.
  const hidden = new Set(table.hide || []);
  const rows = table.cells.map((cells, r) => {
    const out = cells.map((cell, c) => {
      if (!cell || cell.merged) return "";
      const ref = cellName(r, c);
      const words = plainOf(cell.text);
      if (cell.f) {
        const value = sheet.cell(r, c);
        const own = numberSettings(cell);
        const s = styleOf(numberFormatCode(own));
        const f = `<f>${esc(cell.f.replace(/^=/, ""))}</f>`;
        if (value instanceof SheetError) return `<c r="${ref}" s="${s}" t="e">${f}<v>${esc(value.code)}</v></c>`;
        if (typeof value === "number") return `<c r="${ref}" s="${s}">${f}<v>${value}</v></c>`;
        if (typeof value === "boolean") return `<c r="${ref}" s="${s}" t="b">${f}<v>${value ? 1 : 0}</v></c>`;
        return `<c r="${ref}" s="${s}" t="str">${f}<v>${esc(showValue(value, cell))}</v></c>`;
      }
      if (!words) return "";
      const n = parseNumber(words);
      if (n != null) return `<c r="${ref}" s="${styleOf(numberFormatCode(settingsFromWords(words)))}"><v>${n}</v></c>`;
      const day = parseDate(words);
      if (day != null) return `<c r="${ref}" s="${styleOf("yyyy/m/d", true)}"><v>${day}</v></c>`;
      return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${esc(words)}</t></is></c>`;
    }).join("");
    return out ? `<row r="${r + 1}"${hidden.has(r) ? ' hidden="1"' : ""}>${out}</row>` : "";
  }).join("");
  const merges = [];
  table.cells.forEach((row, r) => row.forEach((cell, c) => {
    if (cell && (cell.rs > 1 || cell.cs > 1)) merges.push(`<mergeCell ref="${cellName(r, c)}:${cellName(r + (cell.rs || 1) - 1, c + (cell.cs || 1) - 1)}"/>`);
  }));
  // One character is about 22 px on a slide that is 1920 px wide.
  const widths = (table.cols || []).map((share) => Math.max(4, Math.min(60, Math.round(((share || 0) * (table.w || 960)) / 22))));
  const cols = widths.length ? `<cols>${widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("")}</cols>` : "";
  const numFmts = formatList.map(([key, code], i) => (key === "#date" ? "" : `<numFmt numFmtId="${164 + i}" formatCode="${esc(code)}"/>`)).filter(Boolean);
  const xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>', ...formatList.map(([key, code], i) => {
    const id = key === "#date" ? 14 : 164 + i;
    return `<xf numFmtId="${id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`;
  })];
  const styles = `${HEAD}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.join("")}</numFmts>` : ""}`
    + '<fonts count="1"><font><sz val="11"/><name val="Yu Gothic"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + `<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs></styleSheet>`;
  const sheetName = esc(String(name).replace(/[\\/?*[\]:]/g, " ").slice(0, 31) || "表");
  return {
    "[Content_Types].xml": `${HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>`
      + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',
    "_rels/.rels": `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
    "xl/workbook.xml": `${HEAD}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${sheetName}" sheetId="1" r:id="rId1"/></sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`,
    "xl/_rels/workbook.xml.rels": `${HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    "xl/styles.xml": styles,
    "xl/worksheets/sheet1.xml": `${HEAD}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cols}<sheetData>${rows}</sheetData>${merges.length ? `<mergeCells count="${merges.length}">${merges.join("")}</mergeCells>` : ""}</worksheet>`,
  };
}

/** The .xlsx file (bytes) for a table. */
export function tableToXlsx(table, options = {}) {
  const enc = new TextEncoder();
  return zipFiles(Object.entries(workbookParts(table, options)).map(([path, xml]) => ({ name: path, data: enc.encode(xml) })));
}
