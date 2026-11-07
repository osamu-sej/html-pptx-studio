#!/usr/bin/env python3
"""Turn an existing PPTX / PDF / DOCX into slideData the studio can edit and brush up.

PowerPoint keeps the most: slide boundaries, titles, bullet text, tables, chart data,
the main photo and speaker notes. PDF is read page by page (text only) and Word is
split at headings. The result is a faithful first draft, not a redesign; the
"AI brush-up" step rewrites it into a clear storyline afterwards.
"""

from __future__ import annotations

import base64
import io
import json
import re
import sys
import zipfile
from typing import Any

from lxml import etree
from PIL import Image

# Page furniture a template repeats on every slide: confidentiality marks, copyright lines, page numbers.
CHROME_TEXT = re.compile(r"^(秘（?[A-CＡ-Ｃ]）?|社内限り|社外秘|confidential|(©|\(c\)|copyright).*|\d{1,3}|STEP\s*\d+|[QA]|NEXT ACTION|KEY MESSAGE|\d{2})$", re.I)
DATE_LINE = re.compile(r"^(\d{4}\s*[年/.-]\s*\d{1,2}\s*月?(\s*[/.-]?\s*\d{1,2}\s*日?)?|令和\s*\d+\s*年.*)$")
CLOSING_TITLE = re.compile(r"(次のアクション|ネクストステップ|今後の(予定|進め方|アクション)|まとめ|おわりに|お願い|ありがとう|Next\s*Step)", re.I)
MAX_POINTS = 10
EMU_IN = 914400
SYMBOL_ONLY = re.compile(r"^[→←↑↓⇒⇔▶►▲▼▷◀＞>・\s]+$")
# A figure shown big on a KPI card: "120h/月", "-42%", "+1.2pt", "2.4億円", "4.5点".
VALUE = re.compile(r"^[+\-−▲▼△]?\s*[¥$]?\d[\d,.]*\s*(%|％|pt|ポイント|倍|[a-zA-Z]{0,3}(/[月年日週人店件])?|[万億千百]?[円件店人名時間分秒日週月年回個台点社]|か月|ヶ月)?$")
AGENDA_TITLE = re.compile(r"(アジェンダ|目次|本日の(内容|流れ|議題)|agenda)", re.I)
FAQ_TITLE = re.compile(r"(質問|FAQ|Ｑ＆Ａ|Q\s*&\s*A|問答)", re.I)


def clip(text: Any, limit: int) -> str:
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    return text if len(text) <= limit else text[: limit - 1] + "…"


def clean_line(text: str) -> str:
    """Drop a leading bullet mark. A minus sign in front of a figure ("-42%") is kept."""
    text = re.sub(r"^[\s・•●○■□◆◇▶►*]+", "", text)
    return re.sub(r"^[\-–—]+(?=\s|$|[^\d.])\s*", "", text).strip()


def image_data_url(blob: bytes, max_side: int = 1400) -> str | None:
    try:
        with Image.open(io.BytesIO(blob)) as img:
            if img.width < 200 or img.height < 120:
                return None  # icons and logos are not worth carrying over
            img = img.convert("RGB")
            scale = min(1.0, max_side / max(img.size))
            if scale < 1:
                img = img.resize((round(img.width * scale), round(img.height * scale)))
            out = io.BytesIO()
            img.save(out, "JPEG", quality=85)
            return "data:image/jpeg;base64," + base64.b64encode(out.getvalue()).decode()
    except Exception:
        return None


# ---------------------------------------------------------------- PowerPoint

def chart_spec(chart) -> dict[str, Any] | None:
    try:
        plot = chart.plots[0]
        categories = [clip(c, 80) for c in plot.categories]
        series = [(clip(s.name, 80), [float(v or 0) for v in s.values]) for s in plot.series]
    except Exception:
        return None
    if not categories or not series:
        return None
    kind = str(chart.chart_type).lower()
    title = clip(chart.chart_title.text_frame.text, 120) if chart.has_title else ""
    data: dict[str, Any] = {"title": title} if title else {}
    categories, series = categories[:20], [(name, values[:20]) for name, values in series[:8]]
    if "pie" in kind or "doughnut" in kind:
        return {"chartType": "donut", "data": {**data, "items": [{"label": c, "value": v} for c, v in zip(categories, series[0][1])]}}
    if "line" in kind:
        if len(series) == 1:
            return {"chartType": "line", "data": {**data, "items": [{"label": c, "value": v} for c, v in zip(categories, series[0][1])]}}
        return {"chartType": "multi-line", "data": {**data, "xAxisLabels": categories, "series": [{"label": n, "values": v} for n, v in series]}}
    if "stacked" in kind and len(series) > 1:
        chart_type = "100-stacked-bar" if "100" in kind else "stacked-bar"
        return {"chartType": chart_type, "data": {**data, "legendLabels": [n for n, _ in series],
                                                   "barData": [{"label": c, "values": [v[i] for _, v in series]} for i, c in enumerate(categories)]}}
    if len(series) == 1:
        return {"chartType": "bar", "data": {**data, "items": [{"label": c, "value": v} for c, v in zip(categories, series[0][1])]}}
    return {"chartType": "multi-line", "data": {**data, "xAxisLabels": categories, "series": [{"label": n, "values": v} for n, v in series]}}


def shape_lines(shape) -> list[str]:
    lines = []
    for paragraph in shape.text_frame.paragraphs:
        text = clean_line("".join(run.text for run in paragraph.runs) or paragraph.text)
        if text and not CHROME_TEXT.match(text) and not SYMBOL_ONLY.match(text):
            lines.append(("　" * min(paragraph.level, 2)) + text)
    return lines


def walk(shapes, to_slide=lambda x, y, w, h: (x, y, w, h)):
    """Every shape with its box on the slide; group members are mapped out of their group's own space."""
    from pptx_geometry import _xfrm_map

    for shape in shapes:
        if shape.shape_type == 6:  # group
            inner = _xfrm_map(shape)
            yield from walk(shape.shapes, lambda x, y, w, h, inner=inner: to_slide(*inner(x, y, w, h)))
        else:
            yield shape, to_slide(shape.left or 0, shape.top or 0, shape.width or 0, shape.height or 0)


DIAGRAM_URI = "http://schemas.openxmlformats.org/drawingml/2006/diagram"


def smartart_lines(slide, shape) -> list[str]:
    """Text of a SmartArt graphic, which lives in its own data part rather than on the slide."""
    try:
        rel_ids = shape._element.xpath(".//*[local-name()='relIds']")
        if not rel_ids:
            return []
        rid = rel_ids[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}dm")
        root = etree.fromstring(slide.part.related_part(rid).blob)
    except (KeyError, AttributeError, etree.XMLSyntaxError):
        return []
    lines = []
    for point in root.iter("{%s}pt" % DIAGRAM_URI):
        if point.get("type") not in (None, "node"):
            continue
        for p in point.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}p"):
            text = clean_line("".join(t.text or "" for t in p.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}t")))
            if text and not CHROME_TEXT.match(text):
                lines.append(text)
    return lines


def picture_of(shape):
    """The image of a picture or a filled picture placeholder, else None."""
    if shape.shape_type == 13 or getattr(shape, "is_placeholder", False):
        try:
            return shape.image
        except (AttributeError, ValueError, KeyError):
            return None
    return None


def looks(slide, shape) -> tuple[float, bool]:
    """(font size, bold) of a text box's first line: headings are bigger or bolder than what they head."""
    from pptx_geometry import _paragraph_size

    for paragraph in shape.text_frame.paragraphs:
        if paragraph.text.strip():
            runs = [r for r in paragraph.runs if r.text.strip()]
            return _paragraph_size(paragraph, shape, slide), bool(runs) and all(r.font.bold for r in runs)
    return 18.0, False


def heading_like(block: dict[str, Any]) -> bool:
    text = block["lines"][0].strip() if len(block["lines"]) == 1 else ""
    return bool(text) and len(text) <= 30 and not text.endswith("。")


def units_of(blocks: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Group a heading with the text right under it (a card's title and description, a KPI's value,
    label and note, a question and its answer) so they are read together, then put the groups in
    reading order: row by row, left to right."""
    blocks = sorted(blocks, key=lambda b: (b["y"], b["x"]))
    used: set[int] = set()
    units = []
    for i, head in enumerate(blocks):
        if i in used:
            continue
        used.add(i)
        members, last = [], head
        if heading_like(head):
            # A label with its text right beside it on the same row ("改革の背景と目的 | 人手不足下で…").
            # PDF boxes end where their text ends, so the gap may be wide when nothing lies in between.
            beside = []
            for j, c in enumerate(blocks):
                if j in used or c is head:
                    continue
                v_overlap = min(head["y"] + head["h"], c["y"] + c["h"]) - max(head["y"], c["y"])
                h_gap = c["x"] - (head["x"] + head["w"])
                weaker = c["size"] < head["size"] - 0.5 or (head["bold"] and not c["bold"])
                if v_overlap >= 0.5 * min(head["h"], c["h"]) and -0.1 * EMU_IN <= h_gap <= 1.6 * EMU_IN:
                    beside.append((h_gap, j, c, weaker))
            if beside:
                h_gap, j, c, weaker = min(beside, key=lambda item: item[0])
                if weaker and (h_gap <= 0.4 * EMU_IN or "bbox" in head):
                    members.append(c)
                    used.add(j)
        if heading_like(head) and not members:
            for j in range(i + 1, len(blocks)):
                c = blocks[j]
                if j in used:
                    continue
                overlap = min(head["x"] + head["w"], c["x"] + c["w"]) - max(head["x"], c["x"])
                gap = c["y"] - (last["y"] + last["h"])
                weaker = c["size"] < head["size"] - 0.5 or (head["bold"] and not c["bold"])
                # Text spanning several headings of one row (a footer line under three cards) belongs to none.
                spans_others = any(o is not head and abs(o["y"] - head["y"]) < 0.2 * EMU_IN and heading_like(o)
                                   and min(o["x"] + o["w"], c["x"] + c["w"]) - max(o["x"], c["x"]) > 0 for o in blocks)
                if (c["y"] > head["y"] + head["h"] * 0.5 and overlap >= 0.5 * min(head["w"], c["w"]) and not spans_others
                        and c["w"] >= 0.5 * head["w"] and -0.1 * EMU_IN <= gap <= 0.35 * EMU_IN and weaker):
                    members.append(c)
                    used.add(j)
                    last = c
                elif c["y"] > last["y"] + last["h"] + 0.35 * EMU_IN:
                    break
        lines = [line for m in members for line in m["lines"]]
        unit = {"x": head["x"], "y": head["y"], "w": head["w"]}
        if members and VALUE.match(head["lines"][0].strip()):
            unit.update(kind="kpi", value=head["lines"][0].strip(), label=lines[0], note=" ".join(lines[1:]))
        elif members:
            unit.update(kind="pair", head=head["lines"][0].strip(), lines=lines)
        else:
            unit.update(kind="text", lines=head["lines"])
        units.append(unit)
    # Reading order: units whose tops are within a third of an inch share a row.
    rows: list[list[dict[str, Any]]] = []
    for unit in sorted(units, key=lambda u: u["y"]):
        if rows and unit["y"] - rows[-1][0]["y"] < 0.33 * EMU_IN:
            rows[-1].append(unit)
        else:
            rows.append([unit])
    rows = [sorted(row, key=lambda u: u["x"]) for row in rows]
    # A list laid out in two or more columns of three or more rows reads down each column first.
    columns = {len(row) for row in rows}
    if len(rows) >= 3 and columns == {len(rows[0])} and len(rows[0]) >= 2 and all(
            abs(row[c]["x"] - rows[0][c]["x"]) < 0.2 * EMU_IN for row in rows for c in range(len(row))):
        return [row[c] for c in range(len(rows[0])) for row in rows]
    return [unit for row in rows for unit in row]


def unit_points(unit: dict[str, Any]) -> list[str]:
    if unit["kind"] == "kpi":
        return [f"{unit['label']}：{unit['value']}" + (f"（{unit['note']}）" if unit["note"] else "")]
    if unit["kind"] == "pair":
        rest = unit["lines"]
        return [f"{unit['head']}：{rest[0].strip()}"] + rest[1:] if rest else [unit["head"]]
    return unit["lines"]


COMMENT_RELS = ("/relationships/comments",)  # classic and modern (2018) PowerPoint comments both end like this


def slide_comments(slide, authors: dict[str, str]) -> list[dict[str, str]]:
    """Review comments on one slide, from classic and modern PowerPoint comment parts (replies included)."""
    from lxml import etree

    found = []
    for rel in slide.part.rels.values():
        if not rel.reltype.endswith(COMMENT_RELS) or rel.is_external:
            continue
        try:
            root = etree.fromstring(rel.target_part.blob)
        except (etree.XMLSyntaxError, AttributeError):
            continue
        for node in root.iter():
            name = etree.QName(node).localname
            if name not in {"cm", "reply"}:
                continue
            if name == "cm" and node.find("{*}text") is not None:  # classic: <p:cm authorId><p:text>
                text = node.find("{*}text").text or ""
            else:  # modern: <p188:cm authorId><p188:txBody>…<a:t>
                body = node.find("{*}txBody")
                if body is None:
                    continue
                text = "".join(t.text or "" for t in body.iter("{*}t"))
            text = re.sub(r"\s+", " ", text).strip()
            if text:
                found.append({"author": authors.get(str(node.get("authorId")), ""), "text": clip(text, 600)})
    return found


def comment_authors(prs) -> dict[str, str]:
    from lxml import etree

    authors = {}
    for part in prs.part.package.iter_parts():
        name = str(part.partname)
        if name in {"/ppt/commentAuthors.xml", "/ppt/authors.xml"}:
            try:
                root = etree.fromstring(part.blob)
            except etree.XMLSyntaxError:
                continue
            for node in root:
                if node.get("id") is not None:
                    authors[str(node.get("id"))] = node.get("name") or node.get("initials") or ""
    return authors


def read_pptx(data: bytes) -> tuple[str, list[dict[str, Any]], dict[str, int]]:
    from pptx import Presentation

    prs = Presentation(io.BytesIO(data))
    stats = {"tables": 0, "charts": 0, "images": 0, "notes": 0}
    raw = []
    authors = comment_authors(prs)
    for slide in prs.slides:
        title = ""
        if slide.shapes.title is not None and slide.shapes.title.has_text_frame:
            title = clip(slide.shapes.title.text_frame.text, 90)
        blocks, tables, chart, picture = [], [], None, None
        for shape, (x, y, w, h) in walk(slide.shapes):
            if slide.shapes.title is not None and shape.shape_id == slide.shapes.title.shape_id:
                continue
            image = picture_of(shape)
            if getattr(shape, "has_table", False) and shape.has_table:
                rows = [[clip(cell.text, 120) for cell in row.cells] for row in shape.table.rows]
                if rows and len(rows[0]) >= 2:
                    tables.append(rows)
            elif getattr(shape, "has_chart", False) and shape.has_chart:
                chart = chart or chart_spec(shape.chart)
            elif image is not None:
                if picture is None or w * h > picture[0]:
                    picture = (w * h, image)
            elif getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                lines = shape_lines(shape)
                if lines:
                    size, bold = looks(slide, shape)
                    blocks.append({"x": x, "y": y, "w": w, "h": h, "lines": lines, "size": size, "bold": bold})
            elif shape._element.xpath(".//*[local-name()='graphicData'][@uri='%s']" % DIAGRAM_URI):
                lines = smartart_lines(slide, shape)
                if lines:
                    blocks.append({"x": x, "y": y, "w": w, "h": h, "lines": lines, "size": 18.0, "bold": False})
        blocks.sort(key=lambda block: (block["y"], block["x"]))
        if not title and blocks:
            # No title placeholder: use the top-most short line.
            first = blocks[0]["lines"][0].strip()
            if len(first) <= 40:
                title = clip(first, 90)
                blocks[0] = {**blocks[0], "lines": blocks[0]["lines"][1:]}
                if not blocks[0]["lines"]:
                    blocks = blocks[1:]
        lead = ""
        if blocks and len(blocks[0]["lines"]) == 1 and blocks[0]["y"] < 1.9 * EMU_IN and 10 <= len(blocks[0]["lines"][0]) <= 70 and len(blocks) > 1:
            # A single short sentence right under the title reads as the slide's message.
            lead = blocks[0]["lines"][0].strip()
            blocks = blocks[1:]
        units = units_of(blocks)
        lines = [line for unit in units for line in unit_points(unit)]
        notes = ""
        if slide.has_notes_slide:
            notes = clip(slide.notes_slide.notes_text_frame.text, 1200)
        image = image_data_url(picture[1].blob) if picture else None
        raw.append({"title": title, "lead": lead, "lines": lines, "units": units, "tables": tables, "chart": chart, "image": image,
                    "notes": notes, "comments": slide_comments(slide, authors)})
    title = next((r["title"] for r in raw if r["title"]), "")
    return title, *convert_all(raw, stats)


def convert_all(raw: list[dict[str, Any]], stats: dict[str, Any]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Convert every source page; one page may become several slides when it holds more than one fits."""
    slides, comments = [], []
    for i, r in enumerate(raw):
        comments += [{"slide": len(slides), **c} for c in r.get("comments", [])]
        slides += convert(r, i, len(raw), stats)
    return slides, {**stats, "comments": comments}


def chunks(items: list[str], size: int) -> list[list[str]]:
    return [items[i:i + size] for i in range(0, len(items), size)] or [[]]


def continued(title: str, part: int) -> str:
    return clip(title, 90) if part == 0 else clip(f"{title}（続き）", 90)


def structured(r: dict[str, Any], title: str, common: dict[str, Any]) -> dict[str, Any] | None:
    """Rebuild a layout the page plainly had: KPI cards, a two-sided comparison, an agenda or a FAQ."""
    units = r.get("units") or []
    if not units:
        return None
    kinds = [u["kind"] for u in units]
    if 2 <= kinds.count("kpi") <= 4 and all(k == "kpi" for k in kinds):
        return {"type": "kpi", "title": clip(title, 90), **common,
                "items": [{"label": clip(u["label"], 60), "value": clip(u["value"], 40), **({"change": clip(u["note"], 60)} if u["note"] else {})} for u in units]}
    if len(units) == 2 and kinds == ["pair", "pair"] and abs(units[0]["y"] - units[1]["y"]) < 0.33 * EMU_IN and units[0]["lines"] and units[1]["lines"]:
        left, right = units
        return {"type": "compare", "title": clip(title, 90), **common,
                "leftTitle": clip(left["head"], 60), "rightTitle": clip(right["head"], 60),
                "leftItems": [clip(x.strip(), 180) for x in left["lines"][:8]], "rightItems": [clip(x.strip(), 180) for x in right["lines"][:8]]}
    if FAQ_TITLE.search(title) and 1 <= len(units) <= 4 and all(k == "pair" for k in kinds):
        mark = lambda text, letter: re.sub(rf"^[{letter}][.:：\s]*", "", text.strip())  # "Q" / "A" badges read into the text
        return {"type": "faq", "title": clip(title, 90), **common,
                "items": [{"q": clip(mark(u["head"], "QＱ"), 100), "a": clip(mark(" ".join(x.strip() for x in u["lines"]), "AＡ"), 220)} for u in units]}
    return None


def convert(r: dict[str, Any], index: int, total: int, stats: dict[str, int]) -> list[dict[str, Any]]:
    title = r["title"] or f"スライド{index + 1}"
    lines = [clip(line.strip(), 180) for line in r["lines"] if line.strip()]
    notes = r["notes"]
    if notes:
        stats["notes"] += 1
    common = {"notes": notes} if notes else {}
    if index == 0:
        slide = {"type": "title", "title": clip(title, 100), **common}
        candidates = [line for line in ([r.get("lead", "")] + lines) if line and not DATE_LINE.match(line)]
        if candidates:
            # Department/audience labels are short; the real subtitle is the longest line.
            slide["subtitle"] = clip(max(candidates[:4], key=len), 180)
        date = next((line for line in lines if DATE_LINE.match(line)), "")
        if date:
            slide["date"] = clip(date, 32)
        if r["image"]:
            slide["customImage"] = r["image"]
            stats["images"] += 1
        return [slide]
    if r.get("section"):
        return [{"type": "section", "title": clip(title, 90), **common}]
    if r.get("lead"):
        common["takeaway"] = clip(r["lead"], 120)
    if index == total - 1 and CLOSING_TITLE.search(title):
        message = ""
        for part in [r.get("lead", "")] + lines:
            message = _join(message, part.strip()) if message and part.strip() else message or part.strip()
        return [{"type": "closing", "title": clip(title, 90), "message": clip(message, 240), **({"notes": notes} if notes else {})}]
    out: list[dict[str, Any]] = []
    tables = r.get("tables") or ([r["table"]] if r.get("table") else [])
    if tables:
        # Every table, in pieces of 8 rows under a repeated header; nothing is cut off.
        for number, rows in enumerate(tables):
            stats["tables"] += 1
            headers = [cell or "—" for cell in rows[0][:6]]
            body = [row[: len(headers)] for row in rows[1:] if any(row)] or [["" for _ in headers]]
            for part, piece in enumerate(chunks(body, 8)):
                out.append({"type": "table", "title": continued(title, len(out)), "headers": headers, "rows": piece, **(common if not out else {})})
    elif r["chart"]:
        stats["charts"] += 1
        out.append({"type": "imageText", "title": clip(title, 90), "image": r["chart"], "points": lines[:8] or ["グラフから読み取れること"], **common})
        lines = lines[8:]
    elif AGENDA_TITLE.search(title) and 2 <= len(lines) <= 10:
        return [{"type": "agenda", "title": clip(title, 90), "items": lines, **common}]
    else:
        rebuilt = structured(r, title, common)
        if rebuilt:
            return [rebuilt]
    if tables:
        # Text beside a table: a short line becomes the table's message, anything longer its own slide.
        if len(lines) == 1 and "takeaway" not in out[0] and len(lines[0]) <= 120:
            out[0]["takeaway"], lines = lines[0], []
    for part, piece in enumerate(chunks(lines, MAX_POINTS)):
        if not piece and (out or part):
            break
        slide: dict[str, Any] = {"type": "content", "title": continued(title, len(out)), "points": piece, **(common if not out else {})}
        if len(piece) > 6:
            slide["twoColumn"] = True
        if r["image"] and not out and not r["chart"]:
            slide["customImage"] = r["image"]
            stats["images"] += 1
        out.append(slide)
    return out


# ---------------------------------------------------------------- PDF / Word

PDF_EMU = 12700  # PDF points → EMU, so PDF pages go through the same grouping as PowerPoint slides
BULLET_START = re.compile(r"^[・•●○■□◆◇▶►\-–—*※①-⑳]|^\(?\d{1,2}[.)．）]\s|^[0-9０-９]+\s*[.．、]")


CJK = "　-ヿ㐀-鿿豈-﫿＀-￯"
PDF_SPACE = re.compile(rf"(?<=[{CJK}]) +(?=[0-9A-Za-z%.,+\-−/])|(?<=[0-9A-Za-z%.,/]) +(?=[{CJK}])")


def _join(a: str, b: str) -> str:
    """Join two visual lines of one paragraph: Japanese runs on, Latin words keep a space."""
    return a + (" " if a[-1:].isascii() and a[-1:].isalnum() and b[:1].isascii() and b[:1].isalnum() else "") + b


def pdf_lines(page) -> list[dict[str, Any]]:
    """Every visual line of a PDF page with its box, size and weight. PDF text often carries a space
    wherever the font changes between Latin and Japanese ("120 時間"); those spaces are removed."""
    lines = []
    for block in page.get_text("dict", sort=True)["blocks"]:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            # Text from separate boxes on one baseline ("Q" beside its question) arrives as one line; a
            # wide gap between spans splits it again.
            pieces: list[list[dict[str, Any]]] = []
            for sp in line["spans"]:
                if not sp["text"].strip():
                    continue
                if pieces and sp["bbox"][0] - pieces[-1][-1]["bbox"][2] <= sp["size"] * 0.7:
                    pieces[-1].append(sp)
                else:
                    pieces.append([sp])
            for spans in pieces:
                text = PDF_SPACE.sub("", re.sub(r"\s+", " ", "".join(sp["text"] for sp in spans)).strip())
                x0, y0 = spans[0]["bbox"][0], min(sp["bbox"][1] for sp in spans)
                x1, y1 = spans[-1]["bbox"][2], max(sp["bbox"][3] for sp in spans)
                lines.append({"text": text, "x0": x0, "y0": y0, "x1": x1, "y1": y1, "size": round(max(sp["size"] for sp in spans), 1),
                              "bold": all(sp["flags"] & 16 or "bold" in sp.get("font", "").lower() for sp in spans)})
    return lines


def pdf_blocks(page) -> list[dict[str, Any]]:
    """Text of a PDF page regrouped into text boxes and paragraphs, like a PowerPoint slide's shapes.

    Exported slides often write every visual line separately, and two columns side by side come out
    interleaved. A line joins the box above it when it has the same size and weight, starts (or is
    centred) at the same place and follows closely; a paragraph ends where a line stops short of
    the box's right edge, ends a sentence, or the next line starts with a bullet."""
    boxes: list[dict[str, Any]] = []
    for line in sorted(pdf_lines(page), key=lambda l: (l["y0"], l["x0"])):
        size = line["size"]
        home = None
        for box in reversed(boxes):
            last = box["rows"][-1]
            gap = line["y0"] - last["y1"]
            # Same start (a bullet's wrapped lines hang about one character in), or centred alike.
            aligned = abs(line["x0"] - box["x0"]) <= size * 1.2 or abs((line["x0"] + line["x1"]) - (last["x0"] + last["x1"])) / 2 <= size * 0.6
            if -size * 0.3 <= gap <= size * 0.8 and aligned and abs(size - last["size"]) <= 0.6 and line["bold"] == last["bold"]:
                home = box
                break
        if home is None:
            boxes.append({"x0": line["x0"], "rows": [line]})
        else:
            home["rows"].append(line)
    out = []
    for box in boxes:
        rows = box["rows"]
        right = max(r["x1"] for r in rows)
        paragraphs: list[str] = []
        for i, row in enumerate(rows):
            prev = rows[i - 1] if i else None
            new = (prev is None or BULLET_START.match(row["text"]) or prev["text"].endswith(("。", "：", ":", "？", "！"))
                   or prev["x1"] < right - prev["size"] * 2.5)
            if new:
                paragraphs.append(row["text"])
            else:
                paragraphs[-1] = _join(paragraphs[-1], row["text"])
        lines = [clean_line(t) for t in paragraphs]
        lines = [t for t in lines if t and not CHROME_TEXT.match(t) and not SYMBOL_ONLY.match(t)]
        if lines:
            x0, y0 = min(r["x0"] for r in rows), min(r["y0"] for r in rows)
            x1, y1 = right, max(r["y1"] for r in rows)
            out.append({"x": x0 * PDF_EMU, "y": y0 * PDF_EMU, "w": (x1 - x0) * PDF_EMU, "h": (y1 - y0) * PDF_EMU,
                        "lines": lines, "size": rows[0]["size"], "bold": rows[0]["bold"], "bbox": (x0, y0, x1, y1)})
    return out


def read_pdf(data: bytes) -> tuple[str, list[dict[str, Any]], dict[str, int]]:
    import fitz  # PyMuPDF

    stats: dict[str, Any] = {"tables": 0, "charts": 0, "images": 0, "notes": 0}
    raw = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        pages = list(doc)[:50]
        found = [pdf_blocks(page) for page in pages]
        # Headers, footers and marks that repeat at the same place in the page margins are not content.
        seen: dict[tuple[str, int, int], int] = {}
        for page, blocks in zip(pages, found):
            margin = [b for b in blocks if not 0.15 < b["bbox"][1] / (page.rect.height or 1) < 0.85]
            for key in {(b["lines"][0], round(b["bbox"][0] / 20), round(b["bbox"][1] / 20)) for b in margin}:
                seen[key] = seen.get(key, 0) + 1
        repeated = {key for key, count in seen.items() if len(pages) >= 3 and count >= max(3, len(pages) * 0.5)}
        for page, blocks in zip(pages, found):
            height = page.rect.height or 1
            tables = []
            try:
                for table in page.find_tables().tables:
                    rows = [[clip(PDF_SPACE.sub("", cell or ""), 120) for cell in row] for row in table.extract()]
                    if len(rows) >= 2 and len(rows[0]) >= 2:
                        tables.append(rows)
                        box = fitz.Rect(table.bbox)
                        blocks = [b for b in blocks if not fitz.Rect(b["bbox"]).intersects(box)]
            except Exception:  # table detection is a bonus; text still comes through
                pass
            blocks = [b for b in blocks if (b["lines"][0], round(b["bbox"][0] / 20), round(b["bbox"][1] / 20)) not in repeated
                      and not (re.fullmatch(r"[\d\s/\-–]+", b["lines"][0]) and not 0.08 < b["bbox"][1] / height < 0.92)]
            if not blocks and not tables:
                continue
            # The title is the biggest text near the top of the page (else the first line).
            top = [b for b in blocks if b["bbox"][1] < height * 0.3 and len(b["lines"][0]) <= 60]
            body = sorted(b["size"] for b in blocks)
            head = max(top, key=lambda b: (b["size"], -b["bbox"][1]), default=None)
            if head is None or (len(blocks) > 1 and head["size"] < body[len(body) // 2] * 1.15):
                head = blocks[0] if blocks else None
            title = clip(head["lines"][0], 90) if head else ""
            if head:
                head["lines"] = head["lines"][1:]
                blocks = [b for b in blocks if b["lines"]]
            lead = ""
            blocks.sort(key=lambda b: (b["y"], b["x"]))
            if len(blocks) > 1 and len(blocks[0]["lines"]) == 1 and blocks[0]["bbox"][1] < height * 0.3 and 10 <= len(blocks[0]["lines"][0]) <= 70:
                lead = blocks[0]["lines"][0].strip()
                blocks = blocks[1:]
            units = units_of(blocks)
            image = None
            for info in page.get_image_info(xrefs=True):
                rect = fitz.Rect(info["bbox"])
                if info.get("xref") and abs(rect) > abs(page.rect) * 0.12:
                    try:
                        image = image_data_url(doc.extract_image(info["xref"])["image"])
                    except Exception:
                        image = None
                    break
            raw.append({"title": title, "lead": lead, "lines": [line for u in units for line in unit_points(u)], "units": units,
                        "tables": tables, "chart": None, "image": image, "notes": ""})
    title = raw[0]["title"] if raw else ""
    return title, *convert_all(raw, stats)


W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def docx_styles(archive: zipfile.ZipFile) -> dict[str, tuple[str, int | None]]:
    """styleId → (lower-case name, outline level), following basedOn so "見出し 2 (カスタム)" counts."""
    try:
        root = etree.fromstring(archive.read("word/styles.xml"))
    except (KeyError, etree.XMLSyntaxError):
        return {}
    raw = {}
    for style in root.iter(W + "style"):
        sid = style.get(W + "styleId")
        name = style.find(W + "name")
        level = style.find(f"{W}pPr/{W}outlineLvl")
        based = style.find(W + "basedOn")
        raw[sid] = ((name.get(W + "val") if name is not None else sid or "").lower(),
                    int(level.get(W + "val")) if level is not None else None,
                    based.get(W + "val") if based is not None else None)
    out = {}
    for sid, (name, level, based) in raw.items():
        seen = {sid}
        while level is None and based in raw and based not in seen:
            seen.add(based)
            level, based = raw[based][1], raw[based][2]
        out[sid] = (name, level)
    return out


def read_docx(data: bytes) -> tuple[str, list[dict[str, Any]], dict[str, int]]:
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
        styles = docx_styles(archive)
    body = root.find(W + "body")
    items: list[tuple[str, Any]] = []  # ("p", (text, heading level or None, bold)) / ("table", rows)
    for node in body if body is not None else []:
        if node.tag == W + "tbl":
            rows = [[clip(" ".join("".join(t.text or "" for t in p.iter(W + "t")) for p in cell.iter(W + "p")).strip(), 120)
                     for cell in row.findall(W + "tc")] for row in node.findall(W + "tr")]
            rows = [row for row in rows if any(row)]
            if rows:
                items.append(("table", rows))
            continue
        if node.tag != W + "p":
            continue
        text = clean_line("".join(t.text or "" for t in node.iter(W + "t")))
        if not text:
            continue
        sid = node.find(f"{W}pPr/{W}pStyle")
        name, level = styles.get(sid.get(W + "val") if sid is not None else "", ("", None))
        own = node.find(f"{W}pPr/{W}outlineLvl")
        if own is not None:
            level = int(own.get(W + "val"))
        if name.startswith(("title", "表題")):
            level = -1
        elif level is None and (name.startswith(("heading", "見出し")) or (sid is not None and sid.get(W + "val") in {"1", "2", "3"})):
            level = int(re.sub(r"\D", "", name) or 1) - 1
        runs = [r for r in node.iter(W + "r") if "".join(t.text or "" for t in r.iter(W + "t")).strip()]
        bold = bool(runs) and all(r.find(f"{W}rPr/{W}b") is not None and r.find(f"{W}rPr/{W}b").get(W + "val") not in ("0", "false") for r in runs)
        items.append(("p", (text, level if level is not None and level <= 2 else None, bold)))
    has_headings = any(kind == "p" and value[1] is not None for kind, value in items)
    sections: list[dict[str, Any]] = []
    doc_title = ""
    for kind, value in items:
        if kind == "table":
            if not sections:
                sections.append({"title": "", "lines": [], "tables": [], "chart": None, "image": None, "notes": ""})
            sections[-1]["tables"].append(value)
            continue
        text, level, bold = value
        # Without heading styles, a short bold line on its own is how people mark a heading.
        heading = level is not None or (not has_headings and bold and len(text) <= 40 and not text.endswith("。"))
        if level == -1 and not doc_title:
            doc_title = clip(text, 100)
        if heading or not sections:
            sections.append({"title": clip(text, 90), "lines": [], "tables": [], "chart": None, "image": None, "notes": "", "level": level})
        else:
            sections[-1]["lines"].append(text)
    for i, section in enumerate(sections):
        nxt = sections[i + 1] if i + 1 < len(sections) else None
        # A chapter heading followed straight by a sub-heading becomes a chapter divider.
        if i and not section["lines"] and not section["tables"] and nxt and (section.get("level") or 0) < (nxt.get("level") or 0):
            section["section"] = True
    stats = {"tables": 0, "charts": 0, "images": 0, "notes": 0}
    title = doc_title or (sections[0]["title"] if sections else "")
    return title, *convert_all(sections[:50], stats)


def main() -> int:
    name = (sys.argv[1] if len(sys.argv) > 1 else "").lower()
    mode = sys.argv[2] if len(sys.argv) > 2 else "layout"
    data = sys.stdin.buffer.read()
    if name.endswith(".pptx") and mode == "exact":
        # As it looks: every slide a 白紙 page of objects (tools/pptx_exact.py).
        from pptx_exact import read_pptx_exact

        result = read_pptx_exact(data)
        if not result["slideData"]:
            raise ValueError("取り込めるスライドがありませんでした（すべて非表示の可能性があります）")
        if len(result["slideData"]) == 1:
            result["slideData"].append({"type": "closing", "title": "おわり"})
        print(json.dumps(result, ensure_ascii=False))
        return 0
    if name.endswith(".pptx"):
        title, slides, stats = read_pptx(data)
        fidelity = "high"
    elif name.endswith(".pdf"):
        title, slides, stats = read_pdf(data)
        fidelity = "text"
    elif name.endswith(".docx"):
        title, slides, stats = read_docx(data)
        fidelity = "text"
    else:
        raise ValueError("取り込めるのは PowerPoint（.pptx）・PDF・Word（.docx）です")
    if not slides:
        raise ValueError("スライドにできる文章が見つかりませんでした（画像だけの資料の可能性があります）")
    comments = stats.pop("comments", [])
    slides = slides[:49]
    if slides[0]["type"] != "title":
        slides.insert(0, {"type": "title", "title": clip(title or "取り込んだ資料", 100)})
        comments = [{**c, "slide": c["slide"] + 1} for c in comments]
    comments = [c for c in comments if c["slide"] < len(slides)]
    if len(slides) < 2 or slides[-1]["type"] != "closing":
        slides.append({"type": "closing", "title": "次のアクション", "message": "【誰が・いつまでに・何をするか】"})
    print(json.dumps({"deckTitle": clip(title or slides[0].get("title") or "取り込んだ資料", 100),
                      "slideData": slides, "stats": {**stats, "slides": len(slides), "comments": len(comments)},
                      "comments": comments, "fidelity": fidelity},
                     ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
