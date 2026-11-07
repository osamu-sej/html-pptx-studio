"""PowerPoint のアニメーションと画面切り替え (tools/pptx_timing.py through tools/pptx_export.py): every effect, option and
transition the studio has is written as PowerPoint's own, every slide part is valid against the ECMA-376 schema, and
the studio's importer (tools/pptx_exact.py) reads the same animations back."""
import base64
import io
import re
import struct
import sys
import unittest
import zipfile
from pathlib import Path

from lxml import etree

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))

from pptx_exact import read_pptx_exact  # noqa: E402
from pptx_export import write_pptx  # noqa: E402
import pptx_timing  # noqa: E402

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
SCHEMA = etree.XMLSchema(etree.parse(str(ROOT / "test" / "ooxml-xsd" / "pml.xsd")))
ANIMATE = (ROOT / "public" / "engine" / "animate.js").read_text(encoding="utf-8")
ENGINE = (ROOT / "public" / "engine" / "engine.js").read_text(encoding="utf-8")


def catalogue(name: str) -> dict[str, str]:
    """The effects of one of animate.js's tables (IN, EM): key → the rest of its line."""
    body = ANIMATE[ANIMATE.index(f"  const {name} = {{"):]
    body = body[: body.index("\n  };")]
    return {m.group(1): m.group(2) for m in re.finditer(r"^    (\w+): \{ (.*)$", body, re.M)}


IN = catalogue("IN")
EM = catalogue("EM")


def wav() -> str:
    data = b"\0\0" * 800
    head = b"RIFF" + struct.pack("<I", 36 + len(data)) + b"WAVEfmt " + struct.pack("<IHHIIHH", 16, 1, 1, 8000, 16000, 2, 16) + b"data" + struct.pack("<I", len(data))
    return "data:audio/wav;base64," + base64.b64encode(head + data).decode()


def shape(oid, x=100, y=100, w=300, h=200, **more):
    return {"id": oid, "kind": "shape", "shape": "rect", "x": x, "y": y, "w": w, "h": h, "fill": "#4472c4", "text": "<p>一つ目</p><p>二つ目の段落</p>", "fs": 32, **more}


def slides_of(data: bytes):
    """Each slide part (parsed), mc:AlternateContent replaced by its Fallback (as a 2007 reader sees it)."""
    out = []
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = sorted((n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)), key=lambda n: int(re.findall(r"\d+", n)[0]))
        for name in names:
            root = etree.fromstring(z.read(name))
            raw = etree.fromstring(z.read(name))
            for alt in list(root.iter(f"{{{MC}}}AlternateContent")):
                fb = alt.find(f"{{{MC}}}Fallback")
                parent = alt.getparent()
                i = list(parent).index(alt)
                parent.remove(alt)
                for k, kid in enumerate(list(fb) if fb is not None else []):
                    parent.insert(i + k, kid)
            out.append((name, root, raw))
    return out


class EveryEffect(unittest.TestCase):
    """One slide per class of effect, every effect of the studio's catalogue on it, each with its directions."""

    @classmethod
    def setUpClass(cls):
        slides = []
        cls.expected = []
        # Entrances and exits, each direction the studio offers.
        for cls_ in ("in", "out"):
            objects, timeline = [], []
            for k, (fx, line) in enumerate(IN.items()):
                if cls_ == "out" and "noExit: true" in line:
                    continue
                dirs = re.findall(r'\["(\w+)", "', line.split("outDirs")[0]) if "dirs:" in line else [None]
                for d in dirs or [None]:
                    oid = f"{cls_}{len(objects)}"
                    objects.append(shape(oid, x=(len(objects) % 6) * 300, y=(len(objects) // 6 % 5) * 210))
                    e = {"id": f"e{oid}", "el": oid, "cls": cls_, "fx": fx, "start": "click" if len(timeline) % 3 == 0 else "after", "dur": 600, "delay": 0}
                    if d:
                        e["dir"] = d
                    timeline.append(e)
            slides.append({"objects": objects, "timeline": timeline})
        # Emphasis, with amounts and colours.
        objects, timeline = [], []
        for fx, line in EM.items():
            oid = f"em{len(objects)}"
            objects.append(shape(oid, x=(len(objects) % 6) * 300, y=(len(objects) // 6) * 210))
            e = {"id": f"e{oid}", "el": oid, "cls": "em", "fx": fx, "start": "click", "dur": 1000, "delay": 0}
            if "amounts:" in line:
                e["amount"] = 180 if "SPIN" in line else 50 if "FADE" in line else 150
            if "dirs:" in line:
                e["dir"] = re.findall(r'\["(\w+)", "', line)[-1]
            if "color:" in line:
                e["color"] = "#c00000"
            timeline.append(e)
        slides.append({"objects": objects, "timeline": timeline})
        # Paths (straight, curved, closed), by paragraph / word / letter, timing options, sounds, after, triggers, media.
        objects = [shape("p1"), shape("p2", x=500), {"id": "t1", "kind": "text", "x": 100, "y": 500, "w": 800, "h": 200, "text": "<p>Hello studio world</p>", "fs": 40},
                   shape("btn", x=1500, y=800, w=200, h=120),
                   {"id": "v1", "kind": "video", "x": 1000, "y": 400, "w": 640, "h": 360, "src": "data:video/mp4;base64," + base64.b64encode(b"\0" * 64).decode(),
                    "trimStart": 1.5, "fadeIn": 0.5, "bookmarks": [{"id": "m1", "t": 2.25, "name": "要点"}], "loop": True, "volume": 0.5},
                   {"id": "a1", "kind": "audio", "x": 1700, "y": 100, "w": 120, "h": 120, "src": "data:audio/mpeg;base64," + base64.b64encode(b"\0" * 64).decode(), "across": 3, "hideIcon": True}]
        timeline = [
            {"id": "x1", "el": "p1", "cls": "path", "fx": "custom", "start": "click", "dur": 2000, "delay": 0, "path": {"pts": [[0, 0], [300, 0]]}},
            {"id": "x2", "el": "p2", "cls": "path", "fx": "custom", "start": "with", "dur": 2000, "delay": 250, "path": {"pts": [[0, 0], [200, 150], [400, 0]], "curve": True}},
            {"id": "x3", "el": "p1", "cls": "path", "fx": "custom", "start": "after", "dur": 2000, "delay": 0, "path": {"pts": [[0, 0], [200, 0], [200, 200]], "closed": True, "curve": True}},
            {"id": "x4", "el": "p1", "cls": "in", "fx": "fade", "start": "click", "dur": 500, "delay": 0, "by": "para"},
            {"id": "x5", "el": "t1", "cls": "in", "fx": "fade", "start": "click", "dur": 500, "delay": 0, "by": "word"},
            {"id": "x6", "el": "p2", "cls": "in", "fx": "appear", "start": "click", "dur": 1, "delay": 0, "by": "char"},
            {"id": "x7", "el": "p2", "cls": "em", "fx": "pulse", "start": "click", "dur": 500, "delay": 0, "repeat": 3, "autoReverse": True, "ease": "in", "sound": "chime"},
            {"id": "x8", "el": "p1", "cls": "em", "fx": "spin", "start": "click", "dur": 1000, "delay": 0, "repeat": "click", "after": "dim"},
            {"id": "x9", "el": "t1", "cls": "em", "fx": "teeter", "start": "click", "dur": 1000, "delay": 0, "repeat": "slide", "after": "hide"},
            {"id": "x10", "el": "p2", "cls": "em", "fx": "growShrink", "start": "click", "dur": 1000, "delay": 0, "after": "hideNext", "rewind": True},
            {"id": "x11", "el": "t1", "cls": "out", "fx": "fade", "start": "click", "dur": 500, "delay": 0},
            {"id": "x12", "el": "p1", "cls": "em", "fx": "spin", "start": "click", "dur": 800, "delay": 0, "trigger": "btn"},
            {"id": "x13", "el": "p2", "cls": "em", "fx": "pulse", "start": "with", "dur": 800, "delay": 0, "trigger": "btn"},
            {"id": "x14", "el": "v1", "cls": "media", "fx": "play", "start": "click", "dur": 1, "delay": 0},
            {"id": "x15", "el": "v1", "cls": "media", "fx": "pause", "start": "after", "dur": 1, "delay": 2000},
        ]
        slides.append({"objects": objects, "timeline": timeline, "transition": "push", "transitionDir": "up", "transitionSound": "chime", "advance": 4})
        # Every transition, each option; one morph pair.
        options = dict(re.findall(r"^    (\w+): \[(.*)\],$", ENGINE[ENGINE.index("const TRANSITION_OPTIONS"):ENGINE.index("const TRANSITION_SOUNDS")], re.M))
        kinds = re.findall(r"(\w+): \"", ENGINE[ENGINE.index("const TRANSITIONS = {"):ENGINE.index("const TRANSITION_OPTIONS")])
        cls.transitions = []
        for kind in kinds:
            for d in re.findall(r'\["(\w+)", "', options.get(kind, "")) or [None]:
                slides.append({"objects": [shape("m", x=100 + 50 * len(cls.transitions))], "transition": kind, "transitionDir": d,
                               "transitionSound": "stop" if kind == "fade" else None, **({"morph": [["m", "m"]]} if kind == "morph" else {})})
                cls.transitions.append((kind, d))
        cls.model = {"title": "every effect", "theme": {}, "slides": slides, "sounds": {"chime": wav()}}
        cls.timelines = [s.get("timeline") or [] for s in slides]
        cls.data, cls.stats = write_pptx(cls.model)
        cls.parts = slides_of(cls.data)
        cls.back = read_pptx_exact(cls.data)

    def test_every_slide_part_is_valid(self):
        for name, root, _ in self.parts:
            ok = SCHEMA.validate(root)
            self.assertTrue(ok, f"{name}: {[e.message for e in list(SCHEMA.error_log)[:3]]}")

    def test_animations_target_whole_shapes(self):
        for name, root, _ in self.parts:
            tops = {el.get("id") for el in root.findall(f"{P}cSld/{P}spTree/*/*/{P}cNvPr")}
            targets = {t.get("spid") for t in root.iter(f"{P}spTgt")}
            self.assertTrue(targets <= tops, f"{name}: {targets - tops}")

    def test_shape_ids_are_unique_on_every_slide(self):
        for name, _, raw in self.parts:
            ids = [el.get("id") for el in raw.iter(f"{P}cNvPr")]
            self.assertEqual(len(ids), len(set(ids)), f"{name}: {ids}")

    def test_every_effect_has_a_preset(self):
        # Each studio effect is either PowerPoint's own (a preset number) or written as the nearest one and counted.
        for fx, line in IN.items():
            if "html: true" in line or "model: true" in line:
                self.assertIn(fx, pptx_timing.NEAREST_IN, fx)
            else:
                self.assertIn(fx, {**pptx_timing.ENTRANCE_IDS, "shape": 0}, fx)
        for fx, line in EM.items():
            if "html: true" in line or "model: true" in line:
                self.assertIn(fx, pptx_timing.NEAREST_EM, fx)
            else:
                self.assertIn(fx, pptx_timing.EMPHASIS_IDS, fx)
        near = sum(1 for t in self.timelines for e in t if e["fx"] in pptx_timing.NEAREST_IN or e["fx"] in pptx_timing.NEAREST_EM)
        self.assertEqual(self.stats["approximated"], near)

    def test_entrances_exits_and_emphasis_read_back(self):
        for k in range(3):
            # One entry per part (an effect written one paragraph at a time comes back as one per paragraph).
            got = list({e["el"]: e for e in reversed(self.back["slideData"][k]["timeline"])}.values())[::-1]
            want = self.timelines[k]
            self.assertEqual(len(got), len(want))
            for g, w in zip(got, want):
                label = f"{w['cls']} {w['fx']} {w.get('dir')}"
                self.assertEqual(g["cls"], w["cls"], label)
                self.assertEqual(g["start"], w["start"], label)
                if w["fx"] in pptx_timing.NEAREST_IN or w["fx"] in pptx_timing.NEAREST_EM:
                    continue
                self.assertEqual(g["fx"], w["fx"], label)
                if w.get("dir") and w["fx"] not in ("growShrink",):
                    self.assertEqual(g.get("dir"), w["dir"], label)
                if w["fx"] in ("spin", "growShrink", "transparency"):
                    self.assertEqual(g.get("amount"), w["amount"], label)
                if w.get("color"):
                    self.assertEqual(g.get("color"), w["color"], label)

    def test_options_triggers_and_media_read_back(self):
        els = {o["id"]: o for o in self.back["slideData"][3]["elements"]}
        entries = self.back["slideData"][3]["timeline"]
        paths = [e for e in entries if e["cls"] == "path"]
        self.assertEqual(len(paths), 3)
        self.assertEqual(paths[1].get("delay"), 250)
        # The curve comes back as points along it (through its middle point).
        self.assertTrue(any(abs(x - 200) < 3 and abs(y - 150) < 3 for x, y in paths[1]["path"]["pts"]), paths[1]["path"]["pts"])
        # By paragraph comes back once per paragraph; by word and by letter as PowerPoint's own iteration.
        self.assertEqual([e.get("by") for e in entries if e["cls"] == "in"], ["para", "para", "word", "char", "char"])
        pulse = next(e for e in entries if e["fx"] == "pulse" and not e.get("trigger"))
        self.assertEqual((pulse.get("repeat"), pulse.get("autoReverse"), pulse.get("ease")), (3, True, "in"))
        spin = next(e for e in entries if e["fx"] == "spin" and not e.get("trigger"))
        self.assertEqual((spin.get("repeat"), spin.get("after")), ("click", "dim"))
        teeter = next(e for e in entries if e["fx"] == "teeter")
        self.assertEqual((teeter.get("repeat"), teeter.get("after"), teeter.get("dur")), ("slide", "hide", 1000))
        grow = next(e for e in entries if e["fx"] == "growShrink")
        self.assertTrue(grow.get("rewind"))
        triggered = [e for e in entries if e.get("trigger")]
        self.assertEqual([e["fx"] for e in triggered], ["spin", "pulse"])
        self.assertEqual([e["start"] for e in triggered], ["click", "with"])
        self.assertEqual(len({e["trigger"] for e in triggered}), 1)
        self.assertEqual([e["fx"] for e in entries if e["cls"] == "media"], ["play", "pause"])
        self.assertEqual(next(e for e in entries if e["fx"] == "pause").get("delay"), 2000)
        video = next(o for o in els.values() if o["kind"] == "video")
        self.assertEqual((video.get("trimStart"), video.get("fadeIn"), video.get("loop"), video.get("volume")), (1.5, 0.5, True, 0.5))
        self.assertEqual([b["name"] for b in video.get("bookmarks") or []], ["要点"])
        audio = next(o for o in els.values() if o["kind"] == "audio")
        self.assertEqual((audio.get("across"), audio.get("hideIcon")), (3, True))

    def test_the_effect_sound_and_the_transition_sound(self):
        _, _, raw = self.parts[3]
        snds = raw.findall(f".//{P}sndTgt")
        self.assertEqual(len(snds), 1)
        self.assertIsNotNone(raw.find(f".//{P}stSnd"))
        with zipfile.ZipFile(io.BytesIO(self.data)) as z:
            rels = z.read("ppt/slides/_rels/slide4.xml.rels").decode()
            self.assertEqual(rels.count("relationships/audio"), 2)  # the sound (shared by both) and the audio object
            self.assertTrue(any(n.endswith(".wav") for n in z.namelist()))
        self.assertEqual(self.back["slideData"][3].get("transitionSound"), "chime")
        self.assertEqual(self.back["slideData"][3].get("advance"), 4)

    def test_transitions_read_back(self):
        for (kind, d), s in zip(self.transitions, self.back["slideData"][4:]):
            if kind == "none":
                self.assertNotIn("transition", s)
                continue
            want = {"drill": "zoom"}.get(kind, kind)
            self.assertEqual(s.get("transition"), want, f"{kind} {d}")
            default = {"slide": "right", "push": "down", "wipe": "right", "flip": "left", "blinds": "horizontal", "curtain": "vertical"}.get(kind)
            if d and d != default:
                self.assertEqual(s.get("transitionDir"), d, f"{kind} {d}")

    def test_morph_pairs_share_a_name(self):
        k = [i for i, (kind, _) in enumerate(self.transitions) if kind == "morph"][0] + 4
        a = self.back["slideData"][k - 1]["elements"][0]["name"]
        b = self.back["slideData"][k]["elements"][0]["name"]
        self.assertTrue(a.startswith("!!"))
        self.assertEqual(a, b)


class Groups(unittest.TestCase):
    def test_a_group_is_one_powerpoint_group_and_moves_as_one(self):
        model = {"theme": {}, "slides": [{"objects": [
            shape("a", group="g"), shape("b", x=450, group="g"), shape("c", x=900),
            shape("d", y=600, group="h"), shape("e", x=450, y=600), shape("f", x=900, y=600, group="h")],
            "timeline": [{"id": "1", "el": "grp:g", "cls": "in", "fx": "wipe", "dir": "left", "start": "click", "dur": 500, "delay": 0},
                         {"id": "2", "el": "grp:h", "cls": "in", "fx": "fade", "start": "click", "dur": 500, "delay": 0}]}]}
        data, _ = write_pptx(model)
        (_, root, _), = slides_of(data)
        self.assertTrue(SCHEMA.validate(root), SCHEMA.error_log)
        grps = root.findall(f"{P}cSld/{P}spTree/{P}grpSp")
        self.assertEqual(len(grps), 1)  # "h" lies around "e": its members stay where they are
        back = read_pptx_exact(data)["slideData"][0]
        self.assertEqual([e["el"].startswith("grp:") for e in back["timeline"]], [True, False, False][: len(back["timeline"])] or [True])
        # The split group's members both move on that click (the second one with the first).
        second = [e for e in back["timeline"] if not e["el"].startswith("grp:")]
        self.assertEqual([e["start"] for e in second], ["click", "with"])

    def test_a_group_whose_member_moves_on_its_own_stays_apart(self):
        # PowerPoint animates whole shapes only (a group's member cannot have an animation of its own).
        model = {"theme": {}, "slides": [{"objects": [shape("a", group="g"), shape("b", x=450, group="g")],
                                          "timeline": [{"id": "1", "el": "grp:g", "cls": "in", "fx": "fade", "start": "click", "dur": 500, "delay": 0},
                                                       {"id": "2", "el": "b", "cls": "em", "fx": "pulse", "start": "click", "dur": 500, "delay": 0}]}]}
        data, _ = write_pptx(model)
        (_, root, _), = slides_of(data)
        self.assertEqual(root.findall(f"{P}cSld/{P}spTree/{P}grpSp"), [])
        tops = {el.get("id") for el in root.findall(f"{P}cSld/{P}spTree/*/*/{P}cNvPr")}
        targets = {t.get("spid") for t in root.iter(f"{P}spTgt")}
        self.assertTrue(targets <= tops, (targets, tops))


if __name__ == "__main__":
    unittest.main()


class FilmedEffects(unittest.TestCase):
    """What only HTML draws comes as the browser's film (entry["clip"]): a video under its part, shown and played at the
    effect's moment, the part (still editable) coming back at its end; importing the file gives the effect back."""

    @classmethod
    def setUpClass(cls):
        from PIL import Image
        out = io.BytesIO()
        Image.new("RGB", (40, 20), (255, 255, 255)).save(out, "PNG")
        poster = "data:image/png;base64," + base64.b64encode(out.getvalue()).decode()
        film = lambda box, dur, **more: {"src": "data:video/mp4;base64," + base64.b64encode(b"\0\0\0\x18ftypmp42" + b"\0" * 64).decode(),  # noqa: E731
                                         "poster": poster, "box": box, "dur": dur, "lasting": False, "loop": False, **more}
        objects = [shape("under", x=80, y=80, w=900, h=500), {"id": "tw", "kind": "text", "x": 100, "y": 100, "w": 800, "h": 120, "text": "<p>タイプライター</p>", "fs": 48},
                   shape("over", x=600, y=150, w=200, h=200), shape("mk", x=100, y=600, w=800, h=150)]
        timeline = [
            {"id": "a1", "el": "tw", "cls": "in", "fx": "typewriter", "start": "click", "dur": 1600, "delay": 0, "clip": film({"x": 88, "y": 88, "w": 824, "h": 144}, 1600)},
            {"id": "a2", "el": "tw", "cls": "em", "fx": "shine", "start": "after", "dur": 1000, "delay": 300, "clip": film({"x": 88, "y": 88, "w": 824, "h": 144}, 1000)},
            {"id": "a3", "el": "mk", "cls": "em", "fx": "marker", "start": "click", "dur": 800, "delay": 0, "clip": film({"x": 88, "y": 588, "w": 824, "h": 174}, 800, lasting=True)},
            {"id": "a4", "el": "tw", "cls": "out", "fx": "fade", "start": "click", "dur": 500, "delay": 0},
        ]
        cls.model = {"theme": {}, "slides": [{"objects": objects, "timeline": timeline}]}
        cls.data, cls.stats = write_pptx(cls.model)
        (cls.name, cls.root, cls.raw), = slides_of(cls.data)
        cls.back = read_pptx_exact(cls.data)["slideData"][0]

    def test_valid_and_counted(self):
        self.assertTrue(SCHEMA.validate(self.root), [e.message for e in list(SCHEMA.error_log)[:3]])
        self.assertEqual(self.stats["clips"], 3)
        self.assertEqual(self.stats.get("approximated", 0), 0)

    def test_each_video_lies_under_its_part_and_keeps_the_effect(self):
        tree = self.raw.find(f"{P}cSld/{P}spTree")
        names = [el.find(".//" + P + "cNvPr").get("name") for el in tree if el.find(".//" + P + "cNvPr") is not None]
        videos = [el for el in tree if el.find(f"{P}nvPicPr/{P}nvPr/{{http://schemas.openxmlformats.org/drawingml/2006/main}}videoFile") is not None]
        self.assertEqual(len(videos), 3)
        order = [el.find(".//" + P + "cNvPr").get("id") for el in tree if el.find(".//" + P + "cNvPr") is not None]
        tw = next(el for el in tree if el.find(f".//{P}txBody") is not None and "タイプライター" in "".join(el.itertext()))
        tw_at = list(tree).index(tw)
        # Both films of the typewriter's part lie right under it, the shape drawn over it stays over the films.
        self.assertTrue(all(list(tree).index(v) < tw_at for v in videos[:2]), (names, order))
        for v in videos:
            self.assertIsNotNone(v.find(f"{P}nvPicPr/{P}nvPr/{P}custDataLst/{P}tags"))
        with zipfile.ZipFile(io.BytesIO(self.data)) as z:
            tags = [n for n in z.namelist() if n.startswith("ppt/tags/")]
            self.assertEqual(len(tags), 3)
            self.assertIn("HPPTX_CLIP", z.read(tags[0]).decode())

    def test_the_films_show_and_play_at_the_effects_moment(self):
        vids = {el.find(f"{P}nvPicPr/{P}cNvPr").get("id") for el in self.raw.iter(f"{P}pic") if el.find(f"{P}nvPicPr/{P}nvPr/{{http://schemas.openxmlformats.org/drawingml/2006/main}}videoFile") is not None}
        effects = [(c.get("presetClass"), c.get("presetID"), c.get("nodeType"), next(c.iter(f"{P}spTgt")).get("spid"), c.find(f"{P}stCondLst/{P}cond").get("delay"))
                   for c in self.raw.iter(f"{P}cTn") if c.get("presetClass")]
        first = [e for e in effects if e[3] in vids][:2]
        # A film's first effect is its entrance (PowerPoint keeps it hidden until then), and it plays with it.
        self.assertEqual([(e[0], e[2]) for e in first], [("entr", "clickEffect"), ("mediacall", "withEffect")])
        # The typewriter's part comes in (appear) as its film ends, and the film goes.
        self.assertIn(("entr", "1", "withEffect", "1600"), [(e[0], e[1], e[2], e[4]) for e in effects if e[3] not in vids])
        # A click on a film does nothing (no interactive sequence of its own).
        self.assertFalse(any(c.get("nodeType") == "interactiveSeq" for c in self.raw.iter(f"{P}cTn")))

    def test_importing_gives_the_effects_back(self):
        kinds = [o["kind"] for o in self.back["elements"]]
        self.assertNotIn("video", kinds)
        got = [(e["cls"], e["fx"], e["start"], e.get("delay", 0)) for e in self.back["timeline"]]
        self.assertEqual(got, [("in", "typewriter", "click", 0), ("em", "shine", "after", 300), ("em", "marker", "click", 0), ("out", "fade", "click", 0)])
        els = {e["fx"]: e["el"] for e in self.back["timeline"]}
        self.assertEqual(els["typewriter"], els["shine"])
        self.assertEqual(els["typewriter"], els["fade"])


class FilmedBackdrop(unittest.TestCase):
    """背景の動き comes as the browser's seamless loop: a video behind every part, playing with the slide and looping;
    importing the file gives the slide its moving backdrop back."""

    @classmethod
    def setUpClass(cls):
        loop = {"src": "data:video/mp4;base64," + base64.b64encode(b"\0\0\0\x18ftypmp42" + b"\0" * 64).decode(), "kind": "particles"}
        cls.data, cls.stats = write_pptx({"theme": {}, "slides": [{"objects": [shape("a"), shape("b", x=500)], "backdropVideo": loop,
                                                                   "timeline": [{"id": "1", "el": "a", "cls": "in", "fx": "fade", "start": "click", "dur": 500, "delay": 0}]}]})
        (cls.name, cls.root, cls.raw), = slides_of(cls.data)

    def test_behind_everything_and_valid(self):
        self.assertTrue(SCHEMA.validate(self.root), [e.message for e in list(SCHEMA.error_log)[:3]])
        tree = self.raw.find(f"{P}cSld/{P}spTree")
        parts = [kid for kid in tree if kid.tag not in (P + "nvGrpSpPr", P + "grpSpPr")]
        self.assertEqual(parts[0].find(f"{P}nvPicPr/{P}cNvPr").get("name"), "背景の動き")
        self.assertEqual(self.stats["backdrops"], 1)

    def test_plays_with_the_slide_and_loops(self):
        spid = self.raw.find(f".//{P}pic/{P}nvPicPr/{P}cNvPr").get("id")
        calls = [c for c in self.raw.iter(f"{P}cTn") if c.get("presetClass") == "mediacall" and next(c.iter(f"{P}spTgt")).get("spid") == spid]
        self.assertEqual(len(calls), 1)
        self.assertNotEqual(calls[0].get("nodeType"), "clickEffect")
        node = next(v for v in self.raw.iter(f"{P}video") if v.find(f".//{P}spTgt").get("spid") == spid)
        self.assertEqual(node.find(f"{P}cMediaNode/{P}cTn").get("repeatCount"), "indefinite")
        self.assertEqual(node.find(f"{P}cMediaNode").get("mute"), "1")
        self.assertFalse(any(c.get("nodeType") == "interactiveSeq" for c in self.raw.iter(f"{P}cTn")))

    def test_importing_gives_the_backdrop_back(self):
        back = read_pptx_exact(self.data)["slideData"][0]
        self.assertEqual(back.get("backdrop"), "particles")
        self.assertNotIn("video", [o["kind"] for o in back["elements"]])
        self.assertEqual([(e["cls"], e["fx"]) for e in back["timeline"]], [("in", "fade")])
