"""Write a studio deck as an editable PowerPoint (.pptx): the inverse of pptx_exact.py.

The browser sends the deck as PowerPoint-shaped objects (public/editor/pptx.mjs): every slide is a list of objects in
slide pixels (1920 × 1080 = 13.333 × 7.5 in, so 1 px = 6350 EMU = 0.5 pt) with text sizes already fitted, pictures as
data URLs, and the deck's theme (its colours by slot and which colour is which theme colour). Shapes, text boxes,
tables, charts, pictures and lines become PowerPoint's own, editable parts; colours that are theme colours are written
as theme colours (schemeClr with lumMod/lumOff) so they follow the theme in PowerPoint too.

Usage: python tools/pptx_export.py < model.json > deck.pptx
"""
from __future__ import annotations

import base64
import binascii
import io
import json
import re
import sys
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData, XyChartData, BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.util import Emu

sys.path.insert(0, str(Path(__file__).resolve().parent))
import pptx_exact as imp  # noqa: E402  (the importer's tables, read the other way)
import pptx_timing as timing  # noqa: E402  (animations and transitions)

EMU_PX = 6350
SLIDE_W, SLIDE_H = 1920, 1080
NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
A = "{%s}" % NS["a"]
P = "{%s}" % NS["p"]
R = "{%s}" % NS["r"]
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
NSDECL = f'xmlns:a="{NS["a"]}" xmlns:p="{NS["p"]}" xmlns:r="{NS["r"]}"'

# The studio's shape names are PowerPoint's preset names, except the few the importer renamed.
PRESETS: dict[str, str] = {}
for _ppt, _studio in imp.SHAPE_MAP.items():
    PRESETS.setdefault(_studio, _ppt.strip())
PRESETS.update({name: name for name in imp.SHAPE_MAP.values() if name in imp.SHAPE_MAP})
DASH_OUT = {"solid": None, "dot": "sysDot", "sysDot": "sysDot", "dash": "dash", "sysDash": "sysDash", "lgDash": "lgDash",
            "dashDot": "dashDot", "lgDashDot": "lgDashDot", "lgDashDotDot": "lgDashDotDot", "roundDot": "sysDot", "longDash": "lgDash"}
ARROW_SIZES = {1: "sm", 2: "med", 3: "lg"}
UL_MARKS = {"disc": "●", "circle": "○", "square": "■", "diamond": "◆", "arrow": "➢", "check": "✓", "star": "★", "dash": "–"}
OL_SCHEMES = {"decimal": "arabicPeriod", "paren": "arabicParenR", "circled": "circleNumDbPlain", "lower-alpha": "alphaLcPeriod",
              "upper-alpha": "alphaUcPeriod", "lower-roman": "romanLcPeriod", "upper-roman": "romanUcPeriod"}
U_LINES = {"double": "dbl", "thick": "heavy", "dotted": "dotted", "dashed": "dash", "wavy": "wavy"}
ANCHORS = {"top": "t", "middle": "ctr", "bottom": "b"}
# Words as the browser's Japanese word breaker counts them, near enough: a run of Latin letters or digits, a run of
# kana or kanji (two at most), or any other mark.
WORD = re.compile(r"[A-Za-z0-9'’]+|[\u3040-\u30ff\u4e00-\u9fff]{1,2}|[^\s]")
ALIGNS = {"left": "l", "center": "ctr", "right": "r", "justify": "just", "distributed": "dist"}


def emu(px: float) -> int:
    return int(round(float(px) * EMU_PX))


def esc(text: str) -> str:
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def xml(fragment: str):
    """A DrawingML/PresentationML fragment as an element (namespaces declared on a wrapper)."""
    return etree.fromstring(f"<w {NSDECL}>{fragment}</w>")[0]


def data_url(src: str) -> tuple[bytes, str] | None:
    m = re.match(r"^data:([\w/+.-]+)?(;base64)?,(.*)$", src or "", re.S)
    if not m:
        return None
    try:
        raw = base64.b64decode(m.group(3)) if m.group(2) else m.group(3).encode("utf-8")
    except (binascii.Error, ValueError):
        return None
    return raw, (m.group(1) or "application/octet-stream")


# ---------------------------------------------------------------- colours

class Colors:
    """Hex colours → DrawingML colour elements; a theme colour is written as one (slot, lumMod, lumOff)."""

    def __init__(self, theme: dict[str, Any]):
        self.refs = {k.lower(): v for k, v in (theme.get("refs") or {}).items()}

    def el(self, hex_color: str | None, alpha: float | None = None) -> str:
        if not hex_color or hex_color == "none":
            return ""
        h = hex_color.lower()
        a = f'<a:alpha val="{int(round(max(0.0, min(1.0, alpha)) * 100000))}"/>' if alpha is not None and alpha < 0.999 else ""
        ref = self.refs.get(h)
        if ref:
            mods = ""
            if ref.get("mod") is not None and ref["mod"] != 100000:
                mods += f'<a:lumMod val="{int(ref["mod"])}"/>'
            if ref.get("off"):
                mods += f'<a:lumOff val="{int(ref["off"])}"/>'
            return f'<a:schemeClr val="{ref["slot"]}">{mods}{a}</a:schemeClr>'
        return f'<a:srgbClr val="{h.lstrip("#").upper()}">{a}</a:srgbClr>'

    def solid(self, hex_color: str | None, alpha: float | None = None) -> str:
        c = self.el(hex_color, alpha)
        return f"<a:solidFill>{c}</a:solidFill>" if c else "<a:noFill/>"


# ---------------------------------------------------------------- rich text → paragraphs

class RichParser(HTMLParser):
    """The studio's rich text (objects.js sanitizeRich: p, ul/ol/li, b, i, u, s, sub, sup, span[style], a, br) as
    paragraphs of runs: [{"attrs": {...}, "runs": [{"text", "b", "i", "u", "s", "baseline", "color", "size", "face", "href", "hl"}]}]."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.paras: list[dict[str, Any]] = []
        self.stack: list[dict[str, Any]] = [{}]
        self.lists: list[dict[str, Any]] = []
        self.para: dict[str, Any] | None = None

    def _open(self, attrs: dict[str, Any]):
        self.para = {"attrs": attrs, "runs": []}
        self.paras.append(self.para)

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        style = dict(kv.split(":", 1) for kv in (a.get("style") or "").split(";") if ":" in kv)
        style = {k.strip(): v.strip() for k, v in style.items()}
        if tag in ("ul", "ol"):
            self.lists.append({"tag": tag, "style": a.get("data-style") or "", "start": int(a.get("start") or 1), "mark": a.get("data-mark") or "",
                               "msize": a.get("data-msize") or "", "n": 0})
            return
        if tag in ("p", "li"):
            pa: dict[str, Any] = {"align": style.get("text-align"), "indent": int(a.get("data-indent") or 0), "bullet": a.get("data-bullet") or "",
                                  "marL": _px(style.get("margin-left")), "first": _px(style.get("text-indent"))}
            if style.get("text-align-last") == "justify":
                pa["align"] = "distributed"
            if tag == "li" and self.lists:
                lst = self.lists[-1]
                pa["list"] = dict(lst)
                pa["indent"] = max(pa["indent"], len(self.lists) - 1)
                lst["n"] += 1
            self._open(pa)
            return
        if tag == "br":
            if self.para is None:
                self._open({})
            self.para["runs"].append({"br": True, **self.stack[-1]})
            return
        top = dict(self.stack[-1])
        if tag == "b":
            top["b"] = True
        elif tag == "i":
            top["i"] = True
        elif tag == "u":
            top["u"] = a.get("data-u") or True
        elif tag == "s":
            top["s"] = True
        elif tag == "sub":
            top["baseline"] = -25000
        elif tag == "sup":
            top["baseline"] = 30000
        elif tag == "a":
            top["href"] = a.get("href") or ""
        if style.get("color"):
            top["color"] = style["color"]
        if style.get("background-color"):
            top["hl"] = style["background-color"]
        if style.get("font-size"):
            top["size"] = _px(style["font-size"])
        if style.get("font-family"):
            top["face"] = style["font-family"].strip("\"' ")
        if style.get("font-weight"):
            top["b"] = style["font-weight"] in ("700", "bold")
        if style.get("font-style"):
            top["i"] = style["font-style"] == "italic"
        deco = style.get("text-decoration", "")
        if "underline" in deco:
            top["u"] = True
        if "line-through" in deco:
            top["s"] = True
        if deco == "none":
            top["u"] = top["s"] = False
        self.stack.append(top)

    def handle_endtag(self, tag):
        if tag in ("ul", "ol"):
            if self.lists:
                self.lists.pop()
            return
        if tag in ("p", "li"):
            self.para = None
            return
        if tag == "br":
            return
        if len(self.stack) > 1:
            self.stack.pop()

    def handle_data(self, data):
        if not data:
            return
        if self.para is None:
            if not data.strip():
                return
            self._open({})
        self.para["runs"].append({"text": data, **self.stack[-1]})


def speaker_png() -> bytes:
    """A plain speaker icon (a sound's picture when the browser drew none)."""
    from PIL import Image, ImageDraw
    img = Image.new("RGBA", (96, 96), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.polygon([(14, 36), (34, 36), (56, 16), (56, 80), (34, 60), (14, 60)], fill=(68, 114, 196, 255))
    for r in (14, 26):
        d.arc((56 - r, 48 - r, 56 + r, 48 + r), -50, 50, fill=(68, 114, 196, 255), width=6)
    out = io.BytesIO()
    img.save(out, "PNG")
    return out.getvalue()


def _px(value: str | None) -> float | None:
    m = re.match(r"^\s*(-?[\d.]+)px\s*$", value or "")
    return float(m.group(1)) if m else None


def parse_rich(html: str) -> list[dict[str, Any]]:
    parser = RichParser()
    parser.feed(html or "")
    parser.close()
    return parser.paras or [{"attrs": {}, "runs": []}]


# ---------------------------------------------------------------- the writer

FONT_FACES = {"body": ("+mn-lt", "+mn-ea"), "heading": ("+mj-lt", "+mj-ea")}


class Writer:
    def __init__(self, model: dict[str, Any]):
        self.model = model
        self.theme = model.get("theme") or {}
        self.colors = Colors(self.theme)
        self.fonts = model.get("fonts") or {}
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = Emu(SLIDE_W * EMU_PX), Emu(SLIDE_H * EMU_PX)
        self.next_id = 2
        self.sound_parts: dict[str, Any] = {}
        self.tag_count = 0
        self.stats = {"slides": 0, "objects": 0, "pictures": 0, "tables": 0, "charts": 0, "rasterized": 0}

    # ---- theme
    def write_theme(self):
        part = self.prs.slide_master.part.part_related_by(RT.THEME)
        root = etree.fromstring(part.blob)
        c = self.theme.get("colors") or {}
        scheme = root.find(f".//{A}clrScheme")
        if scheme is not None and c:
            scheme.set("name", str(self.theme.get("name") or "HTML PPTX Studio"))
            slots = {"dk1": c.get("tx1"), "lt1": c.get("bg1"), "dk2": c.get("tx2"), "lt2": c.get("bg2"),
                     **{f"accent{i + 1}": v for i, v in enumerate(c.get("accents") or [])}, "hlink": c.get("hlink"), "folHlink": c.get("folHlink")}
            for name, value in slots.items():
                slot = scheme.find(A + name)
                if slot is None or not value:
                    continue
                for kid in list(slot):
                    slot.remove(kid)
                slot.append(xml(f'<a:srgbClr val="{value.lstrip("#").upper()}"/>'))
        f = self.theme.get("fonts") or {}
        for tag, face in (("majorFont", f.get("heading")), ("minorFont", f.get("body"))):
            node = root.find(f".//{A}{tag}")
            if node is None or not face:
                continue
            for kid in ("latin", "ea"):
                el = node.find(A + kid)
                if el is not None:
                    el.set("typeface", face)
        part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    # ---- ids
    def new_id(self) -> int:
        """The next free shape id on the slide (python-pptx numbers the pictures, tables and charts it adds the same way)."""
        if self._slide is not None:
            return self._slide.shapes._next_shape_id
        self.next_id += 1
        return self.next_id

    # ---- fills, lines, effects
    def fill_xml(self, o: dict[str, Any]) -> str:
        if o.get("gradient") and (o["gradient"].get("stops")):
            g = o["gradient"]
            stops = "".join(f'<a:gs pos="{int(round(float(st.get("at", 0)) * 100000))}">{self.colors.el(st.get("color"), st.get("opacity"))}</a:gs>' for st in g["stops"])
            ang = int(round(((float(g.get("angle", 90))) % 360) * 60000))
            return f'<a:gradFill rotWithShape="1"><a:gsLst>{stops}</a:gsLst><a:lin ang="{ang}" scaled="0"/></a:gradFill>'
        if o.get("pattern"):
            pt = o["pattern"]
            kind = pt.get("kind") or "pct50"
            return f'<a:pattFill prst="{esc(kind)}"><a:fgClr>{self.colors.el(pt.get("fg") or "#000000")}</a:fgClr><a:bgClr>{self.colors.el(pt.get("bg") or "#ffffff")}</a:bgClr></a:pattFill>'
        fill = o.get("fill")
        if not fill or fill == "none":
            return "<a:noFill/>"
        alpha = o.get("fillOpacity")
        if o.get("opacity") is not None:
            alpha = (alpha if alpha is not None else 1.0) * float(o["opacity"])
        return self.colors.solid(fill, alpha)

    def line_xml(self, o: dict[str, Any], *, line: bool = False) -> str:
        stroke = o.get("stroke")
        if not stroke or stroke == "none":
            return "<a:ln><a:noFill/></a:ln>"
        w = emu(float(o.get("strokeW") or 2))
        attrs = [f'w="{w}"']
        if o.get("cap"):
            attrs.append(f'cap="{ {"round": "rnd", "square": "sq", "flat": "flat"}.get(o["cap"], "flat") }"')
        if o.get("cmpd"):
            attrs.append(f'cmpd="{esc(o["cmpd"])}"')
        alpha = o.get("strokeOpacity")
        if o.get("opacity") is not None:
            alpha = (alpha if alpha is not None else 1.0) * float(o["opacity"])
        inner = self.colors.solid(stroke, alpha)
        dash = DASH_OUT.get(o.get("dash") or "solid")
        if dash:
            inner += f'<a:prstDash val="{dash}"/>'
        join = o.get("join")
        if join == "round":
            inner += "<a:round/>"
        elif join == "bevel":
            inner += "<a:bevel/>"
        elif join == "miter":
            inner += '<a:miter lim="800000"/>'
        if line:
            for key, tag in (("head", "headEnd"), ("tail", "tailEnd")):
                kind = o.get(key)
                if kind and kind != "none":
                    size = ARROW_SIZES.get(int(o.get(f"{key}Size") or 2), "med")
                    inner += f'<a:{tag} type="{esc(imp.ARROWS.get(kind, kind))}" w="{size}" len="{size}"/>'
        return f'<a:ln {" ".join(attrs)}>{inner}</a:ln>'

    def effects_xml(self, o: dict[str, Any]) -> str:
        out = ""
        if o.get("glow"):
            g = o["glow"]
            out += f'<a:glow rad="{emu(float(g.get("r") or 10))}">{self.colors.el(g.get("color") or "#ffc000", g.get("opacity"))}</a:glow>'
        if o.get("shadow"):
            s = o["shadow"]
            dx, dy = float(s.get("dx") or 0), float(s.get("dy") or 0)
            import math
            dist = emu(math.hypot(dx, dy))
            ang = int(round((math.degrees(math.atan2(dy, dx)) % 360) * 60000))
            out += (f'<a:outerShdw blurRad="{emu(float(s.get("blur") or 0))}" dist="{dist}" dir="{ang}" algn="ctr" rotWithShape="0">'
                    f'{self.colors.el(s.get("color") or "#000000", s.get("opacity", 0.35))}</a:outerShdw>')
        if o.get("reflect"):
            r = o["reflect"]
            out += (f'<a:reflection blurRad="6350" stA="{int(round(float(r.get("opacity", 0.4)) * 100000))}" endA="300" '
                    f'endPos="{int(round(float(r.get("size", 0.5)) * 100000))}" dist="{emu(float(r.get("gap") or 0))}" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>')
        if o.get("soft"):
            out += f'<a:softEdge rad="{emu(float(o["soft"]))}"/>'
        return f"<a:effectLst>{out}</a:effectLst>" if out else ""

    def xfrm(self, o: dict[str, Any]) -> str:
        attrs = []
        if o.get("rot"):
            attrs.append(f'rot="{int(round(float(o["rot"]) * 60000)) % 21600000}"')
        if o.get("flipH"):
            attrs.append('flipH="1"')
        if o.get("flipV"):
            attrs.append('flipV="1"')
        return (f'<a:xfrm {" ".join(attrs)}><a:off x="{emu(o["x"])}" y="{emu(o["y"])}"/>'
                f'<a:ext cx="{max(1, emu(o["w"]))}" cy="{max(1, emu(o["h"]))}"/></a:xfrm>')

    def geometry(self, o: dict[str, Any]) -> str:
        shape = o.get("shape") or "rect"
        if shape == "custom" and (o.get("path") or {}).get("pts"):
            return self.cust_geom(o)
        prst = PRESETS.get(shape, shape if shape in imp.SHAPE_MAP else "rect")
        av = ""
        adj = o.get("adj")
        if isinstance(adj, list) and adj:
            if shape in imp.ADJ_ANGLES:
                vals = [int(round(v * 60000)) if is_angle else int(round(v * 100000)) for v, is_angle in zip(adj, imp.ADJ_ANGLES[shape])]
            elif shape in imp.ADJ_SCALE:
                scale = imp.ADJ_SCALE[shape]
                vals = [int(round(float(v) / scale * 100000)) for v in adj]
            else:
                vals = []
            names = ["adj"] if len(vals) == 1 else [f"adj{i + 1}" for i in range(len(vals))]
            av = "".join(f'<a:gd name="{n}" fmla="val {v}"/>' for n, v in zip(names, vals))
        return f'<a:prstGeom prst="{esc(prst)}"><a:avLst>{av}</a:avLst></a:prstGeom>'

    def cust_geom(self, o: dict[str, Any]) -> str:
        path = o["path"]
        w, h = max(1, emu(o["w"])), max(1, emu(o["h"]))
        parts = [path.get("pts") or []] + [p.get("pts") if isinstance(p, dict) else p for p in (path.get("parts") or [])]
        out = ""
        filled = bool(path.get("closed"))
        for pts in parts:
            if not pts or len(pts) < 2:
                continue
            seg = f'<a:moveTo><a:pt x="{int(pts[0][0] * w)}" y="{int(pts[0][1] * h)}"/></a:moveTo>'
            for x, y in pts[1:]:
                seg += f'<a:lnTo><a:pt x="{int(x * w)}" y="{int(y * h)}"/></a:lnTo>'
            if filled:
                seg += "<a:close/>"
            no_fill = "" if filled else ' fill="none"'
            out += f'<a:path w="{w}" h="{h}"{no_fill}>{seg}</a:path>'
        return f'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>{out}</a:pathLst></a:custGeom>'

    # ---- text
    def face(self, key_or_name: str | None) -> tuple[str, str] | None:
        if not key_or_name:
            return None
        if key_or_name in FONT_FACES:
            return FONT_FACES[key_or_name]
        name = (self.fonts.get(key_or_name) or key_or_name)
        return name, name

    def run_props(self, o: dict[str, Any], run: dict[str, Any], *, tag: str = "a:rPr", base_size: float) -> str:
        size = run.get("size") or base_size
        attrs = [f'lang="ja-JP"', f'sz="{int(round(size * 50))}"']
        bold = run.get("b", o.get("bold"))
        if bold:
            attrs.append('b="1"')
        elif bold is False:
            attrs.append('b="0"')
        if run.get("i", o.get("italic")):
            attrs.append('i="1"')
        u = run.get("u", o.get("underline"))
        if u:
            attrs.append(f'u="{U_LINES.get(o.get("uline") or "", "sng") if u is True else U_LINES.get(u, "sng")}"')
        if run.get("s", o.get("strike")):
            attrs.append('strike="sngStrike"')
        if run.get("baseline"):
            attrs.append(f'baseline="{int(run["baseline"])}"')
        if o.get("caps") == "all":
            attrs.append('cap="all"')
        elif o.get("caps") == "small":
            attrs.append('cap="small"')
        if o.get("ls"):
            attrs.append(f'spc="{int(round(float(o["ls"]) * size * 50))}"')
        inner = ""
        if o.get("toutline"):
            t = o["toutline"]
            inner += f'<a:ln w="{emu(float(t.get("w") or 2))}">{self.colors.solid(t.get("color"))}</a:ln>'
        color = run.get("color") or o.get("color")
        if o.get("hollow") and not run.get("color"):
            inner += "<a:noFill/>"
        elif color:
            inner += self.colors.solid(color)
        fx = ""
        if o.get("tglow"):
            g = o["tglow"]
            fx += f'<a:glow rad="{emu(float(g.get("r") or 6))}">{self.colors.el(g.get("color") or "#ffc000", g.get("opacity"))}</a:glow>'
        if o.get("tshadow"):
            import math
            s = o["tshadow"]
            dx, dy = float(s.get("dx") or 0), float(s.get("dy") or 0)
            fx += (f'<a:outerShdw blurRad="{emu(float(s.get("blur") or 0))}" dist="{emu(math.hypot(dx, dy))}" dir="{int(round((math.degrees(math.atan2(dy, dx)) % 360) * 60000))}" algn="ctr" rotWithShape="0">'
                   f'{self.colors.el(s.get("color") or "#000000", s.get("opacity", 0.4))}</a:outerShdw>')
        if fx:
            inner += f"<a:effectLst>{fx}</a:effectLst>"
        if run.get("hl"):
            inner += f"<a:highlight>{self.colors.el(run['hl'])}</a:highlight>"
        face = self.face(run.get("face") or o.get("fontFace") or o.get("font") or "body")
        if face:
            inner += f'<a:latin typeface="{esc(face[0])}"/><a:ea typeface="{esc(face[1])}"/><a:cs typeface="{esc(face[0])}"/>'
        if run.get("href") and self._slide is not None:
            rid = self._slide.part.relate_to(run["href"], RT.HYPERLINK, is_external=True)
            inner += f'<a:hlinkClick r:id="{rid}"/>'
        return f'<{tag} {" ".join(attrs)}>{inner}</{tag}>'

    def paragraphs(self, o: dict[str, Any], html: str | None) -> str:
        base = float(o.get("fs") or 36)  # px (a run's own size is px too)
        lh = float(o.get("lh") or 1.35)
        out = ""
        units: list[list[int]] = []
        for para in parse_rich(html or ""):
            pa = para["attrs"]
            runs = para["runs"]
            biggest = max([float(r.get("size") or 0) for r in runs if r.get("size")] + [base])
            ppr = []
            align = pa.get("align") or o.get("align")
            if ALIGNS.get(align):
                ppr.append(f'algn="{ALIGNS[align]}"')
            lvl = int(pa.get("indent") or 0)
            if lvl:
                ppr.append(f'lvl="{min(lvl, 8)}"')
            lst = pa.get("list")
            mar = emu(pa["marL"]) if pa.get("marL") else 0
            first = emu(pa["first"]) if pa.get("first") else 0
            bullet = ""
            if lst or pa.get("bullet"):
                hang = int(round(biggest * 1.2 * EMU_PX))
                mar = mar or hang * (lvl + 1)
                first = first or -hang
                if lst and lst["tag"] == "ol":
                    scheme = OL_SCHEMES.get(lst["style"] or "decimal", "arabicPeriod")
                    start = int(lst.get("start") or 1)
                    start_at = f' startAt="{start}"' if start > 1 else ""
                    bullet = f'<a:buAutoNum type="{scheme}"{start_at}/>'
                else:
                    char = pa.get("bullet") or UL_MARKS.get((lst or {}).get("style") or "disc", "●")
                    bullet = f'<a:buFont typeface="Arial"/><a:buChar char="{esc(char)}"/>'
                if lst and lst.get("msize"):
                    bullet = f'<a:buSzPct val="{int(float(lst["msize"]) * 1000)}"/>' + bullet
                if lst and lst.get("mark"):
                    mark_color = (self.theme.get("marks") or {}).get(lst["mark"])
                    if mark_color:
                        bullet = f"<a:buClr>{self.colors.el(mark_color)}</a:buClr>" + bullet
            if mar:
                ppr.append(f'marL="{mar}"')
            if first:
                ppr.append(f'indent="{first}"')
            # The line pitch the browser drew (CSS line-height × the largest size), fixed so PowerPoint breaks the same.
            spacing = f'<a:lnSpc><a:spcPts val="{int(round(biggest / 2 * lh * 100))}"/></a:lnSpc>'
            body = spacing + (bullet or ("<a:buNone/>" if not pa.get("list") else ""))
            ptag = f'<a:pPr {" ".join(ppr)}>{body}</a:pPr>'
            rtxt = ""
            for run in runs:
                if run.get("br"):
                    rtxt += f"<a:br>{self.run_props(o, run, base_size=base)}</a:br>"
                    continue
                text = run.get("text") or ""
                if not text:
                    continue
                rtxt += f"<a:r>{self.run_props(o, run, base_size=base)}<a:t>{esc(text)}</a:t></a:r>"
            end = self.run_props(o, runs[-1] if runs else {}, tag="a:endParaRPr", base_size=base)
            out += f"<a:p>{ptag}{rtxt}{end}</a:p>"
            words = "".join(run.get("text") or "" for run in runs)
            units.append([len(re.sub(r"\s", "", words)), len(WORD.findall(words))])
        self.last_paras = units
        return out or f"<a:p>{self.run_props(o, {}, tag='a:endParaRPr', base_size=base)}</a:p>"

    @staticmethod
    def one_line(o: dict[str, Any]) -> bool:
        """Words the browser set on one line: PowerPoint keeps them on one line too (its glyphs may be a little wider)."""
        html = o.get("text") or ""
        if not html or re.search(r"<br|<li|</p>\s*<p", html) or o.get("vertical"):
            return False
        sizes = [float(v) for v in re.findall(r"font-size:\s*([\d.]+)px", html)] + [float(o.get("fs") or 36)]
        pad = o.get("pad") or [7, 14, 7, 14]
        inner = float(o.get("h") or 0) - float(pad[0]) - float(pad[2])
        return inner < max(sizes) * float(o.get("lh") or 1.35) * 1.5

    def body_pr(self, o: dict[str, Any]) -> str:
        pad = o.get("pad") or [7, 14, 7, 14]
        t, r, b, l = (emu(v) for v in pad)
        wrap = "none" if o.get("wrap") is False or (o.get("kind") == "text" and self.one_line(o)) else "square"
        attrs = [f'lIns="{l}"', f'tIns="{t}"', f'rIns="{r}"', f'bIns="{b}"', f'anchor="{ANCHORS.get(o.get("valign") or "top", "t")}"',
                 f'wrap="{wrap}"', 'rtlCol="0"']
        if o.get("vertical"):
            attrs.append('vert="eaVert"')
        if o.get("cols") and int(o["cols"]) > 1:
            attrs.append(f'numCol="{int(o["cols"])}" spcCol="{emu(float(o.get("colGap") or 28))}"')
        return f'<a:bodyPr {" ".join(attrs)}><a:noAutofit/></a:bodyPr>'

    # ---- objects
    def sp(self, o: dict[str, Any], *, text_box: bool = False) -> str:
        sid = self.new_id()
        name = esc(o.get("name") or f"{'テキスト ボックス' if text_box else '図形'} {sid}")
        descr = f' descr="{esc(o["alt"])}"' if o.get("alt") else ""
        tx_box = ' txBox="1"' if text_box else ""
        nv = f'<p:nvSpPr><p:cNvPr id="{sid}" name="{name}"{descr}/><p:cNvSpPr{tx_box}/><p:nvPr/></p:nvSpPr>'
        sppr = f"<p:spPr>{self.xfrm(o)}{self.geometry(o)}{self.fill_xml(o)}{self.line_xml(o)}{self.effects_xml(o)}</p:spPr>"
        tx = f"<p:txBody>{self.body_pr(o)}<a:lstStyle/>{self.paragraphs(o, o.get('text'))}</p:txBody>"
        o["_spid"] = sid
        o["_paras"] = list(self.last_paras) if o.get("text") else []
        o["_bg"] = "noFill" not in self.fill_xml(o) or "noFill" not in self.line_xml(o)
        return f"<p:sp {NSDECL}>{nv}{sppr}{tx}</p:sp>"

    def cxn(self, o: dict[str, Any]) -> str:
        x1, y1, x2, y2 = (float(o[k]) for k in ("x1", "y1", "x2", "y2"))
        box = {"x": min(x1, x2), "y": min(y1, y2), "w": max(1.0, abs(x2 - x1)), "h": max(1.0, abs(y2 - y1)), "flipH": x2 < x1, "flipV": y2 < y1}
        sid = self.new_id()
        route = {"elbow": "bentConnector3", "curve": "curvedConnector3"}.get(o.get("route") or "", "straightConnector1")
        nv = f'<p:nvCxnSpPr><p:cNvPr id="{sid}" name="{esc(o.get("name") or f"コネクタ {sid}")}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
        sppr = f'<p:spPr>{self.xfrm(box)}<a:prstGeom prst="{route}"><a:avLst/></a:prstGeom>{self.line_xml(o, line=True)}{self.effects_xml(o)}</p:spPr>'
        o["_spid"] = sid
        return f"<p:cxnSp {NSDECL}>{nv}{sppr}</p:cxnSp>"

    def picture(self, slide, o: dict[str, Any]) -> None:
        got = data_url(o.get("src") or "")
        if not got:
            return
        raw, mime = got
        stream = io.BytesIO(raw)
        if mime == "image/svg+xml":
            png = data_url(o.get("png") or "")
            if not png:
                return
            stream = io.BytesIO(png[0])
        pic = slide.shapes.add_picture(stream, Emu(emu(o["x"])), Emu(emu(o["y"])), Emu(max(1, emu(o["w"]))), Emu(max(1, emu(o["h"]))))
        el = pic._element
        sppr = el.spPr
        xf = sppr.find(A + "xfrm")
        if o.get("rot"):
            xf.set("rot", str(int(round(float(o["rot"]) * 60000)) % 21600000))
        if o.get("flipH"):
            xf.set("flipH", "1")
        if o.get("flipV"):
            xf.set("flipV", "1")
        crop = o.get("crop") or {}
        if any(crop.get(k) for k in ("l", "t", "r", "b")):
            blip_fill = el.find(P + "blipFill")
            src = blip_fill.find(A + "srcRect")
            if src is None:
                src = etree.SubElement(blip_fill, A + "srcRect")
                blip_fill.remove(src)
                blip_fill.insert(1, src)
            for k in ("l", "t", "r", "b"):
                if crop.get(k):
                    src.set(k, str(int(round(float(crop[k]) * 100000))))
        if o.get("mask"):
            geom = sppr.find(A + "prstGeom")
            if geom is not None:
                geom.set("prst", PRESETS.get(o["mask"], o["mask"]))
        if o.get("opacity") is not None and float(o["opacity"]) < 1:
            blip = el.find(f"{P}blipFill/{A}blip")
            blip.append(xml(f'<a:alphaModFix amt="{int(round(float(o["opacity"]) * 100000))}"/>'))
        line = self.line_xml(o)
        if "noFill" not in line:
            sppr.append(xml(line))
        fx = self.effects_xml(o)
        if fx:
            sppr.append(xml(fx))
        if o.get("alt"):
            el.nvPicPr.cNvPr.set("descr", o["alt"])
        if o.get("name"):
            el.nvPicPr.cNvPr.set("name", o["name"])
        o["_spid"] = int(el.nvPicPr.cNvPr.get("id"))
        self.stats["pictures"] += 1

    def movie(self, slide, o: dict[str, Any]) -> None:
        """A video PowerPoint plays itself (its first picture is the poster the browser drew)."""
        got = data_url(o["src"])
        if not got:
            return
        raw, mime = got
        poster = data_url(o.get("poster") or "")
        mov = slide.shapes.add_movie(io.BytesIO(raw), Emu(emu(o["x"])), Emu(emu(o["y"])), Emu(max(1, emu(o["w"]))), Emu(max(1, emu(o["h"]))),
                                     poster_frame_image=io.BytesIO(poster[0]) if poster else None, mime_type=mime if mime.startswith("video/") else "video/mp4")
        if o.get("name"):
            mov._element.nvPicPr.cNvPr.set("name", o["name"])
        o["_spid"] = mov.shape_id
        self.media_ext(mov._element, o)
        self.stats["pictures"] += 1

    def media_ext(self, el, o: dict[str, Any]) -> None:
        """再生 options PowerPoint keeps with the media itself (p14:media): trimming, fading, bookmarks."""
        media = el.find(f".//{{{P14}}}media")
        if media is None:
            return
        for kid in list(media):
            media.remove(kid)
        ms = lambda v: str(int(round(float(v) * 1000)))  # noqa: E731
        if o.get("trimStart") or o.get("trimEnd"):
            trim = etree.SubElement(media, f"{{{P14}}}trim")
            if o.get("trimStart"):
                trim.set("st", ms(o["trimStart"]))
            if o.get("trimEnd"):
                trim.set("end", ms(o["trimEnd"]))
        if o.get("fadeIn") or o.get("fadeOut"):
            fade = etree.SubElement(media, f"{{{P14}}}fade")
            if o.get("fadeIn"):
                fade.set("in", ms(o["fadeIn"]))
            if o.get("fadeOut"):
                fade.set("out", ms(o["fadeOut"]))
        marks = [m for m in (o.get("bookmarks") or []) if isinstance(m, dict) and m.get("t") is not None]
        if marks:
            lst = etree.SubElement(media, f"{{{P14}}}bmkLst")
            for m in marks:
                bm = etree.SubElement(lst, f"{{{P14}}}bmk")
                bm.set("name", str(m.get("name") or m.get("id") or "")[:40])
                bm.set("time", ms(m["t"]))

    def sound(self, slide, o: dict[str, Any]) -> None:
        """A sound PowerPoint plays itself: its icon (as the browser drew it) on the slide, the file inside."""
        got = data_url(o.get("src") or "")
        if not got:
            return
        raw, mime = got
        ext = {"audio/mpeg": "mp3", "audio/mp3": "mp3", "audio/wav": "wav", "audio/x-wav": "wav", "audio/mp4": "m4a", "audio/x-m4a": "m4a",
               "audio/aac": "m4a", "audio/ogg": "ogg", "audio/webm": "webm", "audio/flac": "flac"}.get(mime, "mp3")
        part = Part(self.prs.part.package.next_media_partname(ext), mime if mime.startswith("audio/") else "audio/mpeg", self.prs.part.package, raw)
        rid_media = slide.part.relate_to(part, RT.MEDIA)
        rid_audio = slide.part.relate_to(part, RT.AUDIO)
        icon = data_url(o.get("poster") or "") or (speaker_png(), "image/png")
        _, rid_img = slide.part.get_or_add_image_part(io.BytesIO(icon[0]))
        sid = self.new_id()
        name = esc(o.get("name") or f"オーディオ {sid}")
        el = etree.fromstring(
            f'<p:pic {NSDECL} xmlns:p14="{P14}"><p:nvPicPr><p:cNvPr id="{sid}" name="{name}"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr>'
            f'<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:audioFile r:link="{rid_audio}"/>'
            f'<p:extLst><p:ext uri="{{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}}"><p14:media r:embed="{rid_media}"/></p:ext></p:extLst></p:nvPr></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid_img}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f'<p:spPr>{self.xfrm(o)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>')
        slide.shapes._spTree.append(el)
        o["_spid"] = sid
        self.media_ext(el, o)
        self.stats["pictures"] += 1

    def table(self, slide, o: dict[str, Any]) -> None:
        cells = o.get("cells") or []
        n_rows = len(cells)
        n_cols = max((len(r) for r in cells), default=0)
        if not n_rows or not n_cols:
            return
        frame = slide.shapes.add_table(n_rows, n_cols, Emu(emu(o["x"])), Emu(emu(o["y"])), Emu(max(1, emu(o["w"]))), Emu(max(1, emu(o["h"]))))
        tbl = frame.table
        cols = o.get("cols") or [1 / n_cols] * n_cols
        rows = o.get("rows") or [1 / n_rows] * n_rows
        for i, share in enumerate(cols[:n_cols]):
            tbl.columns[i].width = Emu(max(1, emu(float(share) * float(o["w"]))))
        for i, share in enumerate(rows[:n_rows]):
            tbl.rows[i].height = Emu(max(1, emu(float(share) * float(o["h"]))))
        tbl_pr = tbl._tbl.tblPr
        for flag, key in (("firstRow", "header"), ("bandRow", "banded"), ("firstCol", "firstCol"), ("lastRow", "lastRow")):
            on = o.get(key, key in ("header", "banded"))
            tbl_pr.set(flag, "1" if on is not False and on else "0")
        style_id = (self.model.get("tableStyles") or {}).get(o.get("style") or "medium2")
        sid = tbl_pr.find(A + "tableStyleId")
        if sid is None:
            sid = etree.SubElement(tbl_pr, A + "tableStyleId")
        sid.text = style_id or imp.DEFAULT_TABLE_STYLE
        look_of = self.table_look(o, n_rows, n_cols)
        for r, row in enumerate(cells):
            for c in range(n_cols):
                cell_data = row[c] if c < len(row) else {}
                cell = tbl.cell(r, c)
                if cell_data.get("covered"):
                    continue
                st = look_of(r, c)
                cs, rs = int(cell_data.get("cs") or 1), int(cell_data.get("rs") or 1)
                if cs > 1 or rs > 1:
                    cell.merge(tbl.cell(min(n_rows - 1, r + rs - 1), min(n_cols - 1, c + cs - 1)))
                tc = cell._tc
                tx = tc.find(A + "txBody")
                for p in tx.findall(A + "p"):
                    tx.remove(p)
                look = {"fs": o.get("fs", 28), "color": cell_data.get("color") or st.get("color") or o.get("color"), "lh": o.get("lh", 1.35), "font": o.get("font"),
                        "align": cell_data.get("align"), "bold": cell_data.get("bold", st.get("bold"))}
                for p in etree.fromstring(f"<w {NSDECL}>{self.paragraphs(look, cell_data.get('text') or '')}</w>"):
                    tx.append(p)
                tcpr = tc.get_or_add_tcPr()
                tcpr.set("anchor", ANCHORS.get(cell_data.get("valign") or "middle", "ctr"))
                # The cell's padding as the slide draws it (engine.css .hs-otable td: .3em .55em).
                fs_px = float(o.get("fs") or 28)
                pad = [emu(fs_px * 0.3), emu(fs_px * 0.55), emu(fs_px * 0.3), emu(fs_px * 0.55)]
                tcpr.set("marT", str(pad[0])); tcpr.set("marR", str(pad[1])); tcpr.set("marB", str(pad[2])); tcpr.set("marL", str(pad[3]))
                for side, tag in (("bl", "lnL"), ("br", "lnR"), ("bt", "lnT"), ("bb", "lnB")):
                    b = cell_data.get(side) or st.get(side)
                    if b:
                        tcpr.append(xml(f'<a:{tag} w="{emu(float(b.get("w") or 2))}">{self.colors.solid(b.get("c"))}</a:{tag}>'))
                    else:
                        tcpr.append(xml(f'<a:{tag} w="0"><a:noFill/></a:{tag}>'))
                fill = cell_data.get("fill") or st.get("fill")
                tcpr.append(xml(self.colors.solid(fill) if fill else "<a:noFill/>"))
        o["_spid"] = frame.shape_id
        self.stats["tables"] += 1

    def table_look(self, o: dict[str, Any], n_rows: int, n_cols: int):
        """Each cell's fill, borders and words in the table's style, as the slide draws it (engine.css .ts-*)."""
        lk = o.get("look") or {}
        a, t20, t40, lt, dk, mid = (lk.get(k) for k in ("a", "t20", "t40", "lt", "dk", "mid"))
        style = o.get("style") or "medium2"
        header = o.get("header", True) is not False
        banded = o.get("banded", True) is not False
        first_col = bool(o.get("firstCol"))
        last_row = bool(o.get("lastRow"))
        line = lambda c, w=1: {"c": c, "w": w} if c else None

        def look(r: int, c: int) -> dict[str, Any]:
            st: dict[str, Any] = {}
            is_head = header and r == 0
            is_last = last_row and r == n_rows - 1 and not is_head
            data_r = r - (1 if header else 0)
            band = banded and not is_head and not is_last and data_r % 2 == 1
            if style in ("medium2", "medium2a2"):
                st["fill"] = t40 if band else t20
                for side in ("bl", "br", "bt", "bb"):
                    st[side] = line(lt)
                if is_head or is_last or (first_col and c == 0):
                    st.update({"fill": a, "color": lt, "bold": True})
                if is_head:
                    st["bb"] = line(lt, 3)
                if is_last:
                    st["bt"] = line(lt, 3)
            elif style == "light1":
                if band:
                    st["fill"] = t20
                if r == 0:
                    st["bt"] = line(a)
                if r == n_rows - 1:
                    st["bb"] = line(a)
                if is_head:
                    st.update({"bold": True, "bb": line(a)})
                if is_last:
                    st.update({"bt": line(a), "bold": True})
            elif style == "light2":
                st["bb"] = line(a)
                if r == 0:
                    st["bt"] = line(a)
                if c == 0:
                    st["bl"] = line(a)
                if c == n_cols - 1:
                    st["br"] = line(a)
                if is_head:
                    st.update({"fill": a, "color": lt, "bold": True})
                if is_last:
                    st.update({"bt": line(a, 2), "bold": True})
            elif style == "grid":
                for side in ("bl", "br", "bt", "bb"):
                    st[side] = line(dk)
                if is_head:
                    st["bold"] = True
            elif style == "lines":
                st["bb"] = line(mid)
                if r == 0:
                    st["bt"] = line(dk, 2)
                if r == n_rows - 1:
                    st["bb"] = line(dk, 2)
                if is_head:
                    st.update({"bold": True, "bb": line(dk, 2)})
            else:
                if is_head:
                    st["bold"] = True
            if first_col and c == 0 or is_last:
                st["bold"] = True
            return st
        return look

    CHART_TYPES = {"bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "clustered-bar": XL_CHART_TYPE.COLUMN_CLUSTERED, "stacked-bar": XL_CHART_TYPE.COLUMN_STACKED,
                   "hbar": XL_CHART_TYPE.BAR_CLUSTERED, "stacked-hbar": XL_CHART_TYPE.BAR_STACKED, "line": XL_CHART_TYPE.LINE_MARKERS,
                   "multi-line": XL_CHART_TYPE.LINE_MARKERS, "area": XL_CHART_TYPE.AREA, "stacked-area": XL_CHART_TYPE.AREA_STACKED,
                   "pie": XL_CHART_TYPE.PIE, "donut": XL_CHART_TYPE.DOUGHNUT, "radar": XL_CHART_TYPE.RADAR_MARKERS,
                   "scatter": XL_CHART_TYPE.XY_SCATTER, "bubble": XL_CHART_TYPE.BUBBLE}

    def chart(self, slide, o: dict[str, Any]) -> bool:
        ch = o.get("chart") or {}
        kind = self.CHART_TYPES.get(ch.get("type") or "bar")
        if kind is None:
            return False
        labels = [str(v) for v in (ch.get("labels") or [])]
        series = ch.get("series") or []
        if not series:
            return False
        if kind == XL_CHART_TYPE.XY_SCATTER:
            data = XyChartData()
            for s in series:
                ser = data.add_series(s.get("name") or "")
                for x, y in zip(s.get("xs") or range(1, len(s.get("values") or []) + 1), s.get("values") or []):
                    ser.add_data_point(float(x), float(y))
        elif kind == XL_CHART_TYPE.BUBBLE:
            data = BubbleChartData()
            for s in series:
                ser = data.add_series(s.get("name") or "")
                for i, y in enumerate(s.get("values") or []):
                    ser.add_data_point(float((s.get("xs") or [])[i] if i < len(s.get("xs") or []) else i + 1), float(y), float((s.get("sizes") or [])[i] if i < len(s.get("sizes") or []) else 1))
        else:
            data = CategoryChartData()
            data.categories = labels
            for s in (series[:1] if kind in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT) else series):
                data.add_series(s.get("name") or "", [float(v) if v is not None else None for v in (s.get("values") or [])])
        frame = slide.shapes.add_chart(kind, Emu(emu(o["x"])), Emu(emu(o["y"])), Emu(max(1, emu(o["w"]))), Emu(max(1, emu(o["h"]))), data)
        chart = frame.chart
        opts = ch.get("opts") or {}
        if ch.get("title"):
            chart.has_title = True
            chart.chart_title.text_frame.text = str(ch["title"])
        else:
            chart.has_title = False
        legend = opts.get("legend") or "top"
        if legend == "none" or (len(series) < 2 and kind not in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT)):
            chart.has_legend = False
        else:
            chart.has_legend = True
            chart.legend.position = {"top": XL_LEGEND_POSITION.TOP, "bottom": XL_LEGEND_POSITION.BOTTOM, "right": XL_LEGEND_POSITION.RIGHT,
                                     "left": XL_LEGEND_POSITION.LEFT}.get(legend, XL_LEGEND_POSITION.TOP)
            chart.legend.include_in_layout = False
        if opts.get("labels") is not False:
            plot = chart.plots[0]
            plot.has_data_labels = True
            plot.data_labels.show_value = True
        colors = o.get("seriesColors") or []
        plot = chart.plots[0]
        point_colors = o.get("pointColors") or []
        if point_colors and len(plot.series) == 1 and kind not in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
            for i, point in enumerate(plot.series[0].points):
                if i < len(point_colors) and point_colors[i]:
                    point.format.fill.solid()
                    self._paint(point.format.fill, point_colors[i])
        if kind in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
            for i, point in enumerate(plot.series[0].points):
                if i < len(colors):
                    sppr = point.format._element.get_or_add_spPr() if hasattr(point.format, "_element") else None
                    point.format.fill.solid()
                    self._paint(point.format.fill, colors[i])
        else:
            for i, ser in enumerate(plot.series):
                if i < len(colors):
                    if kind in (XL_CHART_TYPE.LINE_MARKERS, XL_CHART_TYPE.RADAR_MARKERS, XL_CHART_TYPE.XY_SCATTER):
                        ser.format.line.fill.solid()
                        self._paint(ser.format.line.fill, colors[i])
                    else:
                        ser.format.fill.solid()
                        self._paint(ser.format.fill, colors[i])
        if opts.get("gap") is not None and hasattr(plot, "gap_width"):
            plot.gap_width = int(opts["gap"])
        if opts.get("overlap") is not None and hasattr(plot, "overlap"):
            plot.overlap = int(opts["overlap"])
        font = chart.font
        font.size = Emu(int(float(o.get("fs") or 24) / 2 * 12700))
        o["_spid"] = frame.shape_id
        self.stats["charts"] += 1
        return True

    def _paint(self, fill, hex_color: str):
        """A python-pptx fill set to a theme colour when it is one, else the colour itself."""
        fore = fill._fill._solidFill if hasattr(fill, "_fill") else None
        el = xml(self.colors.el(hex_color))
        if fore is not None:
            for kid in list(fore):
                fore.remove(kid)
            fore.append(el)

    def background(self, slide, bg: dict[str, Any] | None):
        if not bg:
            return
        if bg.get("image"):
            got = data_url(bg["image"])
            if not got:
                return
            image_part, rid = slide.part.get_or_add_image_part(io.BytesIO(got[0]))
            alpha = f'<a:alphaModFix amt="{int(round((1 - float(bg.get("transparency") or 0)) * 100000))}"/>' if bg.get("transparency") else ""
            fill = (f'<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{rid}">{alpha}</a:blip>'
                    + ('<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>' if bg.get("tile") else "<a:stretch><a:fillRect/></a:stretch>")
                    + "</a:blipFill>")
        elif bg.get("color"):
            fill = self.colors.solid(bg["color"])
        else:
            return
        csld = slide._element.find(P + "cSld")
        old = csld.find(P + "bg")
        if old is not None:
            csld.remove(old)
        csld.insert(0, xml(f"<p:bg><p:bgPr>{fill}<a:effectLst/></p:bgPr></p:bg>"))

    def notes(self, slide, text: str):
        if text:
            slide.notes_slide.notes_text_frame.text = text

    def add_object(self, slide, o: dict[str, Any]):
        kind = o.get("kind")
        tree = slide.shapes._spTree
        self.stats["objects"] += 1
        if kind in ("shape", "text"):
            tree.append(etree.fromstring(self.sp(o, text_box=kind == "text")))
        elif kind == "line":
            tree.append(etree.fromstring(self.cxn(o)))
        elif kind == "table":
            self.table(slide, o)
        elif kind == "chart" and self.chart(slide, o):
            pass
        elif kind == "video" and (o.get("src") or "").startswith("data:"):
            self.movie(slide, o)
        elif kind == "audio":
            if (o.get("src") or "").startswith("data:"):
                self.sound(slide, o)
            else:
                self.stats["objects"] -= 1
        elif o.get("src") or o.get("png"):
            # Pictures, and what the browser drew as a picture (icons, SmartArt it could not split, equations, ink…).
            if kind != "image":
                self.stats["rasterized"] += 1
            self.picture(slide, o)
        else:
            self.stats["objects"] -= 1

    # ---- groups
    @staticmethod
    def box_of(el) -> tuple[int, int, int, int] | None:
        xf = el.find(f"{P}spPr/{A}xfrm")
        if xf is None:
            xf = el.find(f"{P}xfrm")
        if xf is None:
            xf = el.find(f"{P}grpSpPr/{A}xfrm")
        if xf is None:
            return None
        off, ext = xf.find(A + "off"), xf.find(A + "ext")
        if off is None or ext is None:
            return None
        x, y = int(off.get("x")), int(off.get("y"))
        return x, y, x + int(ext.get("cx")), y + int(ext.get("cy"))

    def group_shapes(self, slide, objects: list[dict[str, Any]], alone: set[str] | None = None) -> dict[str, dict[str, Any]]:
        """
        The studio's groups as PowerPoint groups (one part to select, move and animate) where their members lie one
        after another; a group split by other parts keeps its members apart (they still move together).
        """
        tree = slide.shapes._spTree
        members: dict[str, list[dict[str, Any]]] = {}
        for o in objects:
            if o.get("group") and o.get("_el") is not None:
                members.setdefault(str(o["group"]), []).append(o)
        out: dict[str, dict[str, Any]] = {}
        for gid, list_ in members.items():
            spids = [o["_spid"] for o in list_ if o.get("_spid")]
            boxes = [b for b in (self.box_of(o["_el"]) for o in list_) if b]
            box = {"x": 0.0, "y": 0.0, "w": 1.0, "h": 1.0}
            if boxes:
                x1, y1 = min(b[0] for b in boxes), min(b[1] for b in boxes)
                x2, y2 = max(b[2] for b in boxes), max(b[3] for b in boxes)
                box = {"x": x1 / EMU_PX, "y": y1 / EMU_PX, "w": max(1, x2 - x1) / EMU_PX, "h": max(1, y2 - y1) / EMU_PX}
            index = sorted(list(tree).index(o["_el"]) for o in list_)
            if len(list_) < 2 or index[-1] - index[0] + 1 != len(list_) or not boxes or any(str(o.get("id")) in (alone or set()) for o in list_):
                out[gid] = {"spids": spids, "box": box}
                continue
            sid = self.new_id()
            x, y, cx, cy = emu(box["x"]), emu(box["y"]), emu(box["w"]), emu(box["h"])
            grp = etree.fromstring(
                f'<p:grpSp {NSDECL}><p:nvGrpSpPr><p:cNvPr id="{sid}" name="グループ {sid}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
                f'<p:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/><a:chOff x="{x}" y="{y}"/><a:chExt cx="{cx}" cy="{cy}"/></a:xfrm></p:grpSpPr></p:grpSp>')
            tree.insert(index[0], grp)
            for o in sorted(list_, key=lambda m: list(tree).index(m["_el"])):
                grp.append(o["_el"])
            out[gid] = {"spids": [sid], "box": box, "group": True, "el": grp}
        return out

    # ---- animations, transitions, sounds
    def morph_names(self, slides: list[dict[str, Any]]) -> None:
        """変形: the pairs the studio morphs get one name on both slides ("!!…"), so PowerPoint pairs them the same."""
        names: dict[tuple[int, str], str] = {}
        n = 0
        for i, s in enumerate(slides):
            for a_id, b_id in s.get("morph") or []:
                key = (i - 1, str(a_id))
                if key not in names:
                    n += 1
                    names[key] = f"!!部品{n}"
                names[(i, str(b_id))] = names[key]
        for (i, oid), name in names.items():
            if 0 <= i < len(slides):
                for o in slides[i].get("objects") or []:
                    if str(o.get("id")) == oid:
                        o["name"] = name

    def sound_rel(self, slide):
        """A callback for the timing: the rId of one of the studio's sounds (made by the browser as WAV) on this slide."""
        made: dict[str, str] = {}

        def rel(name: str) -> str | None:
            if name in made:
                return made[name]
            got = data_url((self.model.get("sounds") or {}).get(name) or "")
            if not got:
                return None
            part = self.sound_parts.get(name)
            if part is None:
                part = Part(self.prs.part.package.next_media_partname("wav"), "audio/wav", self.prs.part.package, got[0])
                self.sound_parts[name] = part
            made[name] = slide.part.relate_to(part, RT.AUDIO)
            return made[name]

        return rel

    # ---- what only HTML draws, as video
    def tag(self, slide, el, name: str, value: str) -> None:
        """A PowerPoint tag on a picture or video (p:custDataLst): kept by PowerPoint as it saves the file."""
        self.tag_count += 1
        blob = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:tagLst xmlns:p="{NS["p"]}">'
                f'<p:tag name="{esc(name)}" val="{esc(value)}"/></p:tagLst>').encode("utf-8")
        part = Part(PackURI(f"/ppt/tags/tag{self.tag_count}.xml"), "application/vnd.openxmlformats-officedocument.presentationml.tags+xml",
                    self.prs.part.package, blob)
        rid = slide.part.relate_to(part, RT.TAGS)
        nvpr = el.find(f"{P}nvPicPr/{P}nvPr")
        holder = xml(f'<p:custDataLst><p:tags r:id="{rid}"/></p:custDataLst>')
        ext = nvpr.find(P + "extLst")
        if ext is not None:
            ext.addprevious(holder)
        else:
            nvpr.append(holder)

    @staticmethod
    def top_level(el):
        """The slide-level shape an element belongs to (itself, or the group it lies in)."""
        while el is not None and el.getparent() is not None and el.getparent().tag != P + "spTree":
            el = el.getparent()
        return el

    def clips(self, slide, s: dict[str, Any], objects: list[dict[str, Any]], groups: dict[str, dict[str, Any]]) -> tuple[list, list]:
        """
        An HTML-only effect the browser filmed (entry["clip"]): its video lies just under the part it shows, hidden
        until the effect's moment; then it plays while the part steps aside, and at its end the part (still the
        editable one) comes back where the video ends. The studio's own effect is kept in the video's tag
        (HPPTX_CLIP), so importing the file again gives the effect back on the part.
        Returns the timeline as PowerPoint plays it and the videos to list with the media.
        """
        by_id = {str(o.get("id")): o for o in objects}
        out, media = [], []
        for e in s.get("timeline") or []:
            c = e.get("clip")
            el_key = str(e.get("el") or "")
            target_el = None
            if el_key.startswith("grp:") and groups.get(el_key[4:]):
                g = groups[el_key[4:]]
                target_el = g.get("el") if g.get("group") else next((o.get("_el") for o in objects if str(o.get("group")) == el_key[4:] and o.get("_el") is not None), None)
            elif by_id.get(el_key) is not None:
                target_el = by_id[el_key].get("_el")
            got = data_url((c or {}).get("src") or "")
            if not c or target_el is None or not got:
                out.append({k: v for k, v in e.items() if k != "clip"})
                continue
            box = c["box"]
            poster = data_url(c.get("poster") or "")
            mov = slide.shapes.add_movie(io.BytesIO(got[0]), Emu(emu(box["x"])), Emu(emu(box["y"])), Emu(max(1, emu(box["w"]))), Emu(max(1, emu(box["h"]))),
                                         poster_frame_image=io.BytesIO(poster[0]) if poster else None, mime_type="video/mp4")
            vel = mov._element
            vid = f"clip:{e.get('id')}"
            vel.nvPicPr.cNvPr.set("name", f"HTMLの動き {e.get('fx')}")
            # Under the part it shows (anything above the part stays above the video).
            vel.getparent().remove(vel)
            self.top_level(target_el).addprevious(vel)
            spids = [o["_spid"] for o in objects if o.get("_spid") and (str(o.get("id")) == el_key or (el_key.startswith("grp:") and str(o.get("group")) == el_key[4:]))]
            cls, d0, total = e.get("cls"), int(e.get("delay") or 0), int(c.get("dur") or e.get("dur") or 500)
            base = {k: e[k] for k in ("trigger",) if e.get(k)}
            # The video's first effect is its entrance (PowerPoint keeps it hidden until then), and it plays with it.
            seq = [{"el": vid, "cls": "in", "fx": "appear", "start": e.get("start") or "click", "delay": d0, "dur": 1, **base},
                   {"el": vid, "cls": "media", "fx": "play", "start": "with", "delay": d0, "dur": 1, **base}]
            helpers = []
            if cls in ("out", "em"):
                helpers.append({"el": el_key, "cls": "out", "fx": "appear", "start": "with", "delay": d0, "dur": 1, **base})
            if not c.get("loop"):
                if cls == "in" or (cls == "em" and not c.get("lasting")):
                    helpers.append({"el": el_key, "cls": "in", "fx": "appear", "start": "with", "delay": d0 + total, "dur": 1, **base,
                                    **({"after": e["after"]} if e.get("after") else {})})
                if not (cls == "em" and c.get("lasting")):
                    seq.append({"el": vid, "cls": "out", "fx": "appear", "start": "with", "delay": d0 + total, "dur": 1, **base})
            out.extend(seq[:2] + helpers + seq[2:])
            original = {k: v for k, v in e.items() if k != "clip"}
            self.tag(slide, vel, "HPPTX_CLIP", json.dumps({"entry": original, "target": spids, "helpers": [{"cls": h["cls"], "delay": h["delay"]} for h in helpers]},
                                                          ensure_ascii=False, separators=(",", ":")))
            objects.append({"id": vid, "kind": "video", "_spid": mov.shape_id, "x": box["x"], "y": box["y"], "w": box["w"], "h": box["h"]})
            media.append({"id": vid, "kind": "video", "muted": True, "loop": bool(c.get("loop")), "clip": True})
            self.stats["clips"] = self.stats.get("clips", 0) + 1
        return out, media

    def motion(self, slide, s: dict[str, Any], objects: list[dict[str, Any]], groups: dict[str, dict[str, Any]]) -> None:
        objects = list(objects)
        timeline, clip_media = self.clips(slide, s, objects, groups)
        s = {**s, "timeline": timeline}
        targets: dict[str, timing.Target] = {}
        media = list(clip_media)
        for o in objects:
            if not o.get("_spid"):
                continue
            kind = o.get("kind") or "shape"
            if kind == "line":
                x1, y1, x2, y2 = (float(o.get(k) or 0) for k in ("x1", "y1", "x2", "y2"))
                box = {"x": min(x1, x2), "y": min(y1, y2), "w": abs(x2 - x1) or 1, "h": abs(y2 - y1) or 1}
            else:
                box = {k: float(o.get(k) or 0) for k in ("x", "y", "w", "h")}
            targets[str(o.get("id"))] = timing.Target([o["_spid"]], box, kind=kind, paras=o.get("_paras"), has_text=bool(o.get("text")),
                                                      media=kind if kind in ("video", "audio") else None, has_bg=bool(o.get("_bg")))
            if kind in ("video", "audio") and not str(o.get("id")).startswith("clip:"):
                media.append({"id": str(o.get("id")), **{k: o.get(k) for k in ("kind", "autoplay", "loop", "volume", "muted", "across", "hideIdle", "hideIcon", "fullscreen")}})
        for gid, g in groups.items():
            targets[f"grp:{gid}"] = timing.Target(g["spids"], g["box"], kind="group" if g.get("group") else "split")
        rel = self.sound_rel(slide)
        t = timing.Timing(targets, rel)
        xml_timing = t.write(s.get("timeline") or [], media)
        self.stats["approximated"] = self.stats.get("approximated", 0) + t.approximated
        self.stats["animations"] = self.stats.get("animations", 0) + len(s.get("timeline") or []) - t.skipped
        root = slide._element
        for old in [*root.findall(P + "transition"), *root.findall(P + "timing"), *root.findall(f"{{{timing.MC}}}AlternateContent")]:
            root.remove(old)
        anchor = root.find(P + "clrMapOvr")
        if anchor is None:
            anchor = root.find(P + "cSld")
        at = list(root).index(anchor) + 1
        tr = timing.transition_xml(s, rel)
        if tr:
            root.insert(at, xml(tr))
            at += 1
            if s.get("transition") == "morph":
                self.stats["morph"] = self.stats.get("morph", 0) + 1
        if xml_timing:
            root.insert(at, xml(xml_timing))

    def write(self) -> bytes:
        self.write_theme()
        layout = self.prs.slide_layouts[6]
        slides = self.model.get("slides") or []
        self.morph_names(slides)
        for s in slides:
            slide = self.prs.slides.add_slide(layout)
            self._slide = slide
            for ph in list(slide.placeholders):
                ph._element.getparent().remove(ph._element)
            self.background(slide, s.get("background"))
            tree = slide.shapes._spTree
            objects = s.get("objects") or []
            for o in objects:
                before = len(tree)
                self.add_object(slide, o)
                if len(tree) > before:
                    o["_el"] = tree[-1]
            # PowerPoint animates (and starts triggers from) whole shapes only: a group one of whose members moves on
            # its own, is a trigger or plays media stays as its members.
            alone = {str(e.get("el")) for e in s.get("timeline") or []} | {str(e.get("trigger") or "").split("@")[0] for e in s.get("timeline") or []}
            alone |= {str(o.get("id")) for o in objects if o.get("kind") in ("video", "audio")}
            groups = self.group_shapes(slide, objects, alone)
            self.motion(slide, s, objects, groups)
            self.notes(slide, s.get("notes") or "")
            if s.get("hidden"):
                slide._element.set("show", "0")
            self.stats["slides"] += 1
        self._slide = None
        core = self.prs.core_properties
        core.title = str(self.model.get("title") or "")[:255]
        core.author = "HTML PPTX Studio"
        out = io.BytesIO()
        self.prs.save(out)
        return out.getvalue()

    _slide = None
    last_paras: list[list[int]] = []


def write_pptx(model: dict[str, Any]) -> tuple[bytes, dict[str, int]]:
    w = Writer(model)
    data = w.write()
    return data, w.stats


def main() -> int:
    model = json.loads(sys.stdin.buffer.read().decode("utf-8"))
    data, stats = write_pptx(model)
    sys.stdout.buffer.write(data)
    sys.stderr.write(json.dumps(stats) + "\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
