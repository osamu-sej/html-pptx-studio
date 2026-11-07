#!/usr/bin/env python3
"""Read an Excel workbook (.xlsx) as tables the studio can put on a slide: tools/xlsx_table.py NAME < file.xlsx.

Every sheet becomes a grid of cells ({"text", "f", "dec", "sep", "pct", "cur"}, merged areas as rs/cs). A formula keeps its
text ("=SUM(B2:B4)") and the value Excel saved with it; numbers show the way the workbook formats them (桁数・桁区切り・%・¥・日付).
Only the standard library is used (an .xlsx is a zip of XML files).
"""

from __future__ import annotations

import datetime
import json
import re
import sys
import zipfile
from decimal import ROUND_HALF_UP, Decimal
from io import BytesIO
from xml.etree import ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships", "p": "http://schemas.openxmlformats.org/package/2006/relationships"}
MAX_ROWS = 500
MAX_COLS = 100
MAX_FILE = 60_000_000
MAX_PART = 120_000_000

BUILTIN = {0: "General", 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 11: "0.00E+00", 12: "# ?/?", 13: "# ??/??",
           14: "yyyy/m/d", 15: "d-mmm-yy", 16: "d-mmm", 17: "mmm-yy", 18: "h:mm AM/PM", 19: "h:mm:ss AM/PM", 20: "h:mm", 21: "h:mm:ss", 22: "yyyy/m/d h:mm",
           37: "#,##0 ;(#,##0)", 38: "#,##0 ;[Red](#,##0)", 39: "#,##0.00;(#,##0.00)", 40: "#,##0.00;[Red](#,##0.00)", 45: "mm:ss", 46: "[h]:mm:ss", 47: "mmss.0", 49: "@"}
WEEKDAYS = "月火水木金土日"


def col_index(name: str) -> int:
    n = 0
    for ch in name.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def col_name(i: int) -> str:
    n = i + 1
    out = ""
    while n > 0:
        n, m = divmod(n - 1, 26)
        out = chr(65 + m) + out
    return out


def split_ref(ref: str) -> tuple[int, int]:
    m = re.match(r"^\$?([A-Za-z]{1,3})\$?(\d+)$", ref)
    if not m:
        raise ValueError(ref)
    return int(m.group(2)) - 1, col_index(m.group(1))


# ---------------------------------------------------------------- shared formulas (a copy of the master, its addresses moved)

TOKEN = re.compile(r'"(?:[^"]|"")*"|(\$?[A-Z]{1,3}\$?\d{1,4})(?::(\$?[A-Z]{1,3}\$?\d{1,4}))?(?![A-Za-z0-9_.(])|\d+\.?\d*(?:[Ee][+-]?\d+)?|[A-Za-z_][A-Za-z0-9_.]*', re.I)


def shift_ref(text: str, dr: int, dc: int) -> str:
    m = re.match(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$", text)
    ca, col, ra, row = m.groups()
    c = col_index(col) + (0 if ca else dc)
    r = int(row) - 1 + (0 if ra else dr)
    if c < 0 or r < 0:
        return "#REF!"
    return f"{ca}{col_name(c)}{ra}{r + 1}"


def shift_formula(f: str, dr: int, dc: int) -> str:
    def swap(m: re.Match) -> str:
        if m.group(1) is None:
            return m.group(0)
        a = shift_ref(m.group(1), dr, dc)
        return a if m.group(2) is None else f"{a}:{shift_ref(m.group(2), dr, dc)}"

    return TOKEN.sub(swap, f)


# ---------------------------------------------------------------- number formats

def split_sections(code: str) -> list[str]:
    out, cur, quoted = [], "", False
    for ch in code:
        if ch == '"':
            quoted = not quoted
        if ch == ";" and not quoted:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def clean_code(code: str) -> str:
    """The format without quoted words, [colour]/[locale] marks and padding (_x, *x, \\x)."""
    code = re.sub(r'"[^"]*"', "", code)
    code = re.sub(r"\[[^\]]*\]", "", code)
    code = re.sub(r"[_*\\].", "", code)
    return code


def is_date_code(code: str) -> bool:
    c = clean_code(code).lower()
    if "general" in c:
        return False
    return bool(re.search(r"[ymdhs]|aaa", c))


def number_settings(code: str) -> dict:
    """{dec, sep, pct, cur} the studio keeps for a number format (a few things only: 桁数・桁区切り・%・¥/$)."""
    first = split_sections(code)[0]
    raw = first
    c = clean_code(first)
    out: dict = {}
    if c.strip().lower() in ("general", "@", ""):
        return out
    m = re.search(r"\.([0#?]+)", c)
    out["dec"] = len(m.group(1)) if m else 0
    if re.search(r"[0#?],[0#?]", c) or c.strip().startswith("#,"):
        out["sep"] = True
    if "%" in c:
        out["pct"] = True
    if re.search(r"[¥￥]", raw) or "[$¥" in raw or "[$￥" in raw:
        out["cur"] = "¥"
    elif "$" in raw.replace("[$", "") or re.search(r"\[\$\$", raw):
        out["cur"] = "$"
    return {k: v for k, v in out.items() if v is not False and v is not None}


def general(value: float) -> str:
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    text = format(value, ".10g")
    if "e" in text:
        text = format(value, "f").rstrip("0").rstrip(".")
    return text


def show_number(value: float, s: dict) -> str:
    n = value * 100 if s.get("pct") else value
    dec = s.get("dec")
    if dec is None:
        text = general(n)
    else:
        # Excel rounds half away from zero.
        shown = Decimal(repr(abs(n))).quantize(Decimal(1).scaleb(-dec), rounding=ROUND_HALF_UP)
        text = f"{shown:f}"
        if shown != 0 and n < 0:
            text = "-" + text
    neg = text.startswith("-")
    body = text[1:] if neg else text
    whole, _, frac = body.partition(".")
    if s.get("sep"):
        whole = f"{int(whole):,}" if whole.isdigit() else whole
    return f"{'-' if neg else ''}{s.get('cur', '')}{whole}{'.' + frac if frac else ''}{'%' if s.get('pct') else ''}"


def show_date(serial: float, code: str) -> str:
    base = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=serial)
    c = re.sub(r"\[[^\]]*\]", "", code)
    parts = re.split(r'("[^"]*")', c)
    pieces = []
    for i, part in enumerate(parts):
        if i % 2:
            pieces.append(("lit", part.strip('"')))
            continue
        for tok in re.finditer(r"yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|aaaa|aaa|hh|h|ss|s|AM/PM|A/P|.", part, re.I):
            pieces.append(("tok", tok.group(0)))
    out = []
    for i, (kind, tok) in enumerate(pieces):
        if kind == "lit":
            out.append(tok)
            continue
        t = tok.lower()
        if t == "yyyy":
            out.append(f"{base.year}")
        elif t == "yy":
            out.append(f"{base.year % 100:02d}")
        elif t in ("m", "mm"):
            # m next to h or s is minutes, otherwise the month.
            before = next((p[1].lower() for p in reversed(pieces[:i]) if p[0] == "tok" and p[1] not in ":" and p[1].strip()), "")
            after = next((p[1].lower() for p in pieces[i + 1:] if p[0] == "tok" and p[1] not in ":" and p[1].strip()), "")
            minutes = before in ("h", "hh") or after in ("s", "ss")
            value = base.minute if minutes else base.month
            out.append(f"{value:02d}" if t == "mm" else str(value))
        elif t == "mmm":
            out.append(f"{base.month}月")
        elif t == "mmmm":
            out.append(f"{base.month}月")
        elif t in ("d", "dd"):
            out.append(f"{base.day:02d}" if t == "dd" else str(base.day))
        elif t in ("aaa", "ddd"):
            out.append(WEEKDAYS[base.weekday()])
        elif t in ("aaaa", "dddd"):
            out.append(WEEKDAYS[base.weekday()] + "曜日")
        elif t in ("h", "hh"):
            out.append(f"{base.hour:02d}" if t == "hh" else str(base.hour))
        elif t in ("s", "ss"):
            out.append(f"{base.second:02d}" if t == "ss" else str(base.second))
        else:
            out.append(tok)
    return "".join(out)


# ---------------------------------------------------------------- the workbook

def text_of(si: ET.Element) -> str:
    """The words of a shared string or inline string: plain <t>, or the runs <r><t>; phonetic hints (<rPh>) are left out."""
    out = []
    for child in si:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag == "t":
            out.append(child.text or "")
        elif tag == "r":
            out.extend(t.text or "" for t in child.findall("m:t", NS))
    return "".join(out)


def read_workbook(data: bytes) -> list[dict]:
    if len(data) > MAX_FILE:
        raise ValueError("ファイルが大きすぎます")
    try:
        zf = zipfile.ZipFile(BytesIO(data))
    except zipfile.BadZipFile as error:
        raise ValueError("Excelのファイル（.xlsx）として読めません") from error
    names = set(zf.namelist())
    if "xl/workbook.xml" not in names:
        raise ValueError("Excelのファイル（.xlsx）として読めません")
    def parse(name: str) -> ET.Element:
        if zf.getinfo(name).file_size > MAX_PART:
            raise ValueError("ファイルが大きすぎます")
        return ET.fromstring(zf.read(name))

    shared: list[str] = []
    if "xl/sharedStrings.xml" in names:
        shared = [text_of(si) for si in parse("xl/sharedStrings.xml").findall("m:si", NS)]

    # Number formats: cell style index → format code.
    codes: dict[int, str] = {}
    xfs: list[int] = []
    if "xl/styles.xml" in names:
        styles = parse("xl/styles.xml")
        for nf in styles.findall("m:numFmts/m:numFmt", NS):
            codes[int(nf.get("numFmtId", "0"))] = nf.get("formatCode", "General")
        for xf in styles.findall("m:cellXfs/m:xf", NS):
            xfs.append(int(xf.get("numFmtId", "0")))

    def code_of(style: int) -> str:
        fmt = xfs[style] if 0 <= style < len(xfs) else 0
        return codes.get(fmt) or BUILTIN.get(fmt, "General")

    rels = {r.get("Id"): r.get("Target") for r in parse("xl/_rels/workbook.xml.rels").findall("p:Relationship", NS)} if "xl/_rels/workbook.xml.rels" in names else {}
    sheets = []
    for sh in parse("xl/workbook.xml").findall("m:sheets/m:sheet", NS):
        if sh.get("state") in ("hidden", "veryHidden"):
            continue
        target = rels.get(sh.get("{%s}id" % NS["r"]), "")
        path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        if path in names:
            sheets.append((sh.get("name") or f"Sheet{len(sheets) + 1}", path))
    return [read_sheet(parse(path), name, shared, code_of) for name, path in sheets]


def read_sheet(root: ET.Element, name: str, shared: list[str], code_of) -> dict:
    cells: dict[tuple[int, int], dict] = {}
    masters: dict[str, tuple[str, int, int]] = {}
    for row in root.findall("m:sheetData/m:row", NS):
        for c in row.findall("m:c", NS):
            ref = c.get("r")
            if not ref:
                continue
            try:
                r, col = split_ref(ref)
            except ValueError:
                continue
            if r >= MAX_ROWS or col >= MAX_COLS:
                continue
            kind = c.get("t", "n")
            style = int(c.get("s", "0") or 0)
            v = c.find("m:v", NS)
            raw = v.text if v is not None and v.text is not None else ""
            f_el = c.find("m:f", NS)
            cell: dict = {}
            code = code_of(style)
            settings = {} if is_date_code(code) else number_settings(code)
            text = ""
            if kind == "s":
                text = shared[int(raw)] if raw.isdigit() and int(raw) < len(shared) else ""
            elif kind == "inlineStr":
                is_el = c.find("m:is", NS)
                text = text_of(is_el) if is_el is not None else ""
            elif kind == "str":
                text = raw
            elif kind == "b":
                text = "TRUE" if raw == "1" else "FALSE"
            elif kind == "e":
                text = raw
            elif kind == "d":
                text = raw[:10].replace("-", "/")
            elif raw:
                try:
                    number = float(raw)
                except ValueError:
                    number = None
                if number is not None:
                    text = show_date(number, code) if is_date_code(code) else show_number(number, settings)
            if text:
                cell["text"] = text
            if kind == "n" and settings and f_el is not None:
                cell.update(settings)
            if f_el is not None and not is_date_code(code):
                formula = f_el.text or ""
                if f_el.get("t") == "shared" and f_el.get("si") is not None:
                    si = f_el.get("si")
                    if formula:
                        masters[si] = (formula, r, col)
                    elif si in masters:
                        base, mr, mc = masters[si]
                        formula = shift_formula(base, r - mr, col - mc)
                elif f_el.get("t") == "array":
                    formula = ""
                if formula and len(formula) < 500:
                    cell["f"] = "=" + formula
            if cell:
                cells[(r, col)] = cell

    merges = []
    for mc in root.findall("m:mergeCells/m:mergeCell", NS):
        try:
            a, b = (mc.get("ref") or "").split(":")
            (r0, c0), (r1, c1) = split_ref(a), split_ref(b)
        except ValueError:
            continue
        if r0 < MAX_ROWS and c0 < MAX_COLS:
            merges.append((r0, c0, min(r1, MAX_ROWS - 1), min(c1, MAX_COLS - 1)))

    if not cells and not merges:
        return {"name": name, "rows": [], "cols": []}
    rows_n = max([r for r, _ in cells] + [m[2] for m in merges]) + 1
    cols_n = max([c for _, c in cells] + [m[3] for m in merges]) + 1
    grid = [[dict(cells.get((r, c), {})) for c in range(cols_n)] for r in range(rows_n)]
    for r0, c0, r1, c1 in merges:
        if r1 <= r0 and c1 <= c0:
            continue
        first = grid[r0][c0]
        first["rs"], first["cs"] = r1 - r0 + 1, c1 - c0 + 1
        for r in range(r0, min(r1, rows_n - 1) + 1):
            for c in range(c0, min(c1, cols_n - 1) + 1):
                if (r, c) != (r0, c0):
                    grid[r][c] = {"merged": True}
    widths = [8.43] * cols_n
    for col in root.findall("m:cols/m:col", NS):
        try:
            lo, hi, w = int(col.get("min", "1")) - 1, int(col.get("max", "1")) - 1, float(col.get("width", "8.43"))
        except ValueError:
            continue
        for i in range(lo, min(hi, cols_n - 1) + 1):
            widths[i] = max(1.0, w)
    total = sum(widths) or 1.0
    return {"name": name, "rows": grid, "cols": [round(w / total, 5) for w in widths]}


def main() -> int:
    data = sys.stdin.buffer.read()
    sheets = [s for s in read_workbook(data) if s["rows"]]
    if not sheets:
        raise ValueError("表にできるデータがありませんでした")
    json.dump({"sheets": sheets}, sys.stdout, ensure_ascii=False)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except ValueError as error:
        print(f"ValueError: {error}", file=sys.stderr)
        sys.exit(2)
