"""PowerPoint への書き出し (tools/pptx_export.py): a model the browser sends becomes an editable .pptx, and the studio's
own importer (tools/pptx_exact.py) reads back what was written."""
import base64
import io
import sys
import unittest
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from pptx_exact import read_pptx_exact  # noqa: E402
from pptx_export import write_pptx, parse_rich  # noqa: E402

A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def png(color=(36, 81, 230), size=(40, 20)) -> str:
    out = io.BytesIO()
    Image.new("RGB", size, color).save(out, "PNG")
    return "data:image/png;base64," + base64.b64encode(out.getvalue()).decode()


THEME = {
    "id": "clarity", "name": "クリア",
    "colors": {"bg1": "#ffffff", "tx1": "#0e1726", "bg2": "#f4f6fa", "tx2": "#3a4556",
               "accents": ["#2451e6", "#f08a3c", "#13a89e", "#e8b415", "#d65b8b", "#6c7a92"], "hlink": "#2451e6", "folHlink": "#6b7689"},
    "fonts": {"heading": "Zen Kaku Gothic New", "body": "Noto Sans JP"},
    "refs": {"#2451e6": {"slot": "accent1"}, "#ffffff": {"slot": "bg1"}, "#0e1726": {"slot": "tx1"},
             "#d3dcfa": {"slot": "accent1", "mod": 20000, "off": 80000}, "#a7b9f5": {"slot": "accent1", "mod": 40000, "off": 60000}},
    "marks": {"accent": "#2451e6", "muted": "#3a4556"},
}
LOOK = {"a": "#2451e6", "t20": "#d3dcfa", "t40": "#a7b9f5", "lt": "#ffffff", "dk": "#0e1726", "mid": "#868b92"}

MODEL = {
    "title": "書き出しの確認",
    "theme": THEME,
    "tableStyles": {"medium2": "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"},
    "slides": [
        {
            "background": {"color": "#d3dcfa"},
            "notes": "ここで結論を話す。",
            "objects": [
                {"id": "box", "kind": "shape", "shape": "roundRect", "adj": [0.25], "x": 100, "y": 100, "w": 600, "h": 300, "fill": "#2451e6", "stroke": "none",
                 "color": "#ffffff", "fs": 40, "lh": 1.4, "align": "center", "valign": "middle", "rot": 15,
                 "text": '<p>課題<b>太字</b></p><ul><li>一つ目</li><li>二つ目</li></ul>'},
                {"id": "tb", "kind": "text", "x": 800, "y": 100, "w": 900, "h": 80, "fill": "none", "stroke": "none", "color": "#0e1726", "fs": 32, "lh": 1.35,
                 "pad": [7, 14, 7, 14], "text": '<p><span style="color: #f08a3c; font-size: 48px">大きい</span>文字 <a href="https://example.com/">リンク</a></p>'},
                {"id": "ln", "kind": "line", "x1": 900, "y1": 600, "x2": 300, "y2": 700, "stroke": "#a7b9f5", "strokeW": 6, "tail": "triangle", "dash": "dash"},
                {"id": "arc", "kind": "shape", "shape": "arc", "adj": [135, 45], "x": 1300, "y": 300, "w": 300, "h": 300, "fill": "none", "stroke": "#2451e6", "strokeW": 20, "cap": "round"},
                {"id": "pic", "kind": "image", "x": 1200, "y": 700, "w": 400, "h": 200, "src": png(), "crop": {"l": 0.1, "r": 0.1}},
                {"id": "icon", "kind": "picture", "from": "icon", "x": 1700, "y": 700, "w": 100, "h": 100, "src": png((255, 0, 0), (10, 10))},
            ],
        },
        {
            "hidden": True,
            "objects": [
                {"id": "t", "kind": "table", "x": 100, "y": 100, "w": 800, "h": 300, "fs": 28, "color": "#0e1726", "cols": [0.5, 0.5], "rows": [0.34, 0.33, 0.33],
                 "look": LOOK, "cells": [[{"text": "<p>項目</p>"}, {"text": "<p>値</p>"}], [{"text": "<p>A</p>"}, {"text": "<p>10</p>"}], [{"text": "<p>B</p>"}, {"text": "<p>20</p>"}]]},
                {"id": "c", "kind": "chart", "x": 1000, "y": 100, "w": 800, "h": 500, "fs": 24,
                 "chart": {"type": "bar", "labels": ["営業", "人事", "経理"], "series": [{"name": "活用率", "values": [91, 88, 31]}]},
                 "seriesColors": ["#2451e6"], "pointColors": ["#2451e6", "#d3dcfa", "#d3dcfa"]},
            ],
        },
    ],
}


class ExportTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.data, cls.stats = write_pptx(MODEL)
        cls.prs = Presentation(io.BytesIO(cls.data))
        cls.back = read_pptx_exact(cls.data)

    def test_a_widescreen_deck_with_every_slide(self):
        self.assertEqual((self.prs.slide_width, self.prs.slide_height), (12192000, 6858000), "13.333 × 7.5 in")
        self.assertEqual(len(self.prs.slides), 2)
        self.assertEqual(self.prs.core_properties.title, "書き出しの確認")
        self.assertEqual(self.stats["charts"], 1)
        self.assertEqual(self.stats["tables"], 1)

    def test_the_theme_is_a_powerpoint_theme(self):
        part = self.prs.slide_master.part.part_related_by("http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
        root = etree.fromstring(part.blob)
        scheme = root.find(f".//{A}clrScheme")
        self.assertEqual(scheme.find(f"{A}accent1/{A}srgbClr").get("val"), "2451E6")
        self.assertEqual(scheme.find(f"{A}dk1/{A}srgbClr").get("val"), "0E1726")
        self.assertEqual(scheme.find(f"{A}lt1/{A}srgbClr").get("val"), "FFFFFF")
        self.assertEqual(root.find(f".//{A}majorFont/{A}ea").get("typeface"), "Zen Kaku Gothic New")
        self.assertEqual(root.find(f".//{A}minorFont/{A}latin").get("typeface"), "Noto Sans JP")

    def test_theme_colours_are_written_as_theme_colours(self):
        sp = next(s for s in self.prs.slides[0].shapes if s.shape_type is not None and s.has_text_frame and "課題" in s.text_frame.text)
        fill = sp._element.spPr.find(f"{A}solidFill/{A}schemeClr")
        self.assertEqual(fill.get("val"), "accent1", "the fill follows the theme")
        bg = self.prs.slides[0]._element.find(f"{P}cSld/{P}bg/{P}bgPr/{A}solidFill/{A}schemeClr")
        self.assertEqual(bg.get("val"), "accent1")
        self.assertEqual(bg.find(f"{A}lumMod").get("val"), "20000")
        self.assertEqual(bg.find(f"{A}lumOff").get("val"), "80000")

    def test_shapes_keep_their_geometry_and_words(self):
        box = next(o for o in self.back["slideData"][0]["elements"] if o.get("shape") == "roundRect")
        for key, want in (("x", 100), ("y", 100), ("w", 600), ("h", 300)):
            self.assertAlmostEqual(box[key], want, delta=1, msg=key)
        self.assertAlmostEqual(box["rot"], 15, delta=0.1)
        self.assertEqual(box["fill"], "#2451e6")
        self.assertIn("課題", box["text"])
        self.assertIn("一つ目", box["text"])
        self.assertTrue(box["adj"] and abs(box["adj"][0] - 0.25) < 0.01, f"the corner: {box.get('adj')}")

    def test_text_runs_sizes_line_pitch_and_links(self):
        tb = next(s for s in self.prs.slides[0].shapes if s.has_text_frame and "大きい" in s.text_frame.text)
        body = tb._element.txBody
        self.assertEqual(body.bodyPr.get("wrap"), "none", "words the browser set on one line stay on one line")
        runs = body.findall(f".//{A}r")
        self.assertEqual(runs[0].find(f"{A}rPr").get("sz"), "2400", "48 px is 24 pt")
        self.assertEqual(runs[1].find(f"{A}rPr").get("sz"), "1600", "the box's 32 px is 16 pt")
        self.assertEqual(body.find(f".//{A}lnSpc/{A}spcPts").get("val"), str(round(24 * 1.35 * 100)), "the line pitch the browser drew")
        self.assertIsNotNone(runs[-1].find(f"{A}rPr/{A}hlinkClick"), "a link")
        latin = runs[1].find(f"{A}rPr/{A}latin").get("typeface")
        self.assertEqual(latin, "+mn-lt", "the theme's body font")

    def test_lists_become_bullets(self):
        paras = parse_rich('<p>a</p><ul data-style="square"><li>b</li></ul><ol start="3"><li>c</li></ol>')
        self.assertEqual([p["attrs"].get("list", {}).get("tag") for p in paras], [None, "ul", "ol"])
        box = next(s for s in self.prs.slides[0].shapes if s.has_text_frame and "一つ目" in s.text_frame.text)
        self.assertEqual(box._element.txBody.findall(f".//{A}buChar")[0].get("char"), "●")

    def test_lines_and_arcs(self):
        els = self.back["slideData"][0]["elements"]
        line = next(o for o in els if o["kind"] == "line")
        self.assertEqual((round(line["x1"]), round(line["y1"]), round(line["x2"]), round(line["y2"])), (900, 600, 300, 700), "the direction kept (a flipped connector)")
        self.assertEqual(line.get("tail"), "triangle")
        arc = next(o for o in els if o.get("shape") == "arc")
        self.assertEqual([round(v) for v in arc["adj"]], [135, 45], "the arc's start and end angles")

    def test_pictures_and_pictures_of_what_powerpoint_has_not(self):
        pics = [o for o in self.back["slideData"][0]["elements"] if o["kind"] == "image"]
        self.assertEqual(len(pics), 2)
        photo = min(pics, key=lambda o: o["x"])
        self.assertAlmostEqual(photo["crop"]["l"], 0.1, places=2)
        self.assertEqual(self.stats["rasterized"], 1, "the icon went in as a picture")

    def test_tables_carry_the_style_cell_by_cell(self):
        tbl = next(s for s in self.prs.slides[1].shapes if s.has_table).table
        self.assertEqual(tbl._tbl.tblPr.find(f"{A}tableStyleId").text, "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}")
        head = tbl.cell(0, 0)._tc.tcPr.find(f"{A}solidFill/{A}schemeClr")
        self.assertEqual(head.get("val"), "accent1", "the header in the accent")
        band = tbl.cell(2, 0)._tc.tcPr.find(f"{A}solidFill/{A}schemeClr")
        self.assertEqual(band.find(f"{A}lumMod").get("val"), "40000", "the second data row banded")
        head_run = tbl.cell(0, 0)._tc.find(f".//{A}r/{A}rPr")
        self.assertEqual(head_run.get("b"), "1")
        self.assertEqual(head_run.find(f"{A}solidFill/{A}schemeClr").get("val"), "bg1", "white words on the accent")

    def test_charts_are_powerpoint_charts_with_their_colours(self):
        chart = next(s for s in self.prs.slides[1].shapes if s.has_chart).chart
        self.assertEqual(list(chart.plots[0].categories), ["営業", "人事", "経理"])
        self.assertEqual(list(chart.series[0].values), [91.0, 88.0, 31.0])
        pts = chart._chartSpace.findall(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}dPt")
        self.assertEqual(len(pts), 3, "each bar its own colour")

    def test_notes_and_hidden_slides(self):
        self.assertEqual(self.prs.slides[0].notes_slide.notes_text_frame.text, "ここで結論を話す。")
        self.assertEqual(self.prs.slides[1]._element.get("show"), "0")
        self.assertTrue(self.back["slideData"][1]["hidden"])


if __name__ == "__main__":
    unittest.main()
