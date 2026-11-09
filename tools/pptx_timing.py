"""A slide's animations and its transition as PowerPoint writes them (p:timing, p:transition): the inverse of the
importer's timeline_of and transition_of (pptx_exact.py).

The studio's timeline (animate.js) is a list of effects in order, each started on a click, with the previous one or
after it, or by clicking a trigger object. PowerPoint keeps the same thing as a tree: the main sequence holds one group
per click; each click holds the moments its effects start at (the delays are absolute, in ms from the click); each
moment holds its effects, each with its preset (presetClass, presetID, presetSubtype) and the behaviours that draw it
(set, anim, animEffect, animScale, animRot, animClr, animMotion, cmd). The behaviours are written so the slide moves as
the studio moves it (the same starts, durations, delays, directions and amounts), and the preset numbers are
PowerPoint's own, so the アニメーション ウィンドウ names each effect and its 効果のオプション as the studio does.

Effects only HTML can draw (typewriter, decode, blur, a line drawing itself, a number counting up, a chart growing,
shine, ripple, marker, spotlight, the 3D model's own moves) have no PowerPoint preset. Until they are recorded as video
(their own part only), each is written as the nearest PowerPoint effect and counted in `approximated`.
"""
from __future__ import annotations

from typing import Any, Callable

SLIDE_W, SLIDE_H = 1920, 1080

# ---------------------------------------------------------------- presets (the studio's effect → PowerPoint's number)

# Entrances and exits: presetID, and the subtype for the studio's direction option.
FLY_SUB = {"top": 1, "right": 2, "topRight": 3, "bottom": 4, "bottomRight": 6, "left": 8, "topLeft": 9, "bottomLeft": 12}
WIPE_IN = {"top": (1, "down"), "right": (2, "left"), "bottom": (4, "up"), "left": (8, "right")}
# An exit plays the entrance backwards; PowerPoint names its option by where the wipe starts (the other side).
WIPE_OUT_SUB = {"top": 4, "right": 8, "bottom": 1, "left": 2}
SPLIT = {"hOut": (42, "barn(outHorizontal)"), "hIn": (26, "barn(inHorizontal)"), "vOut": (37, "barn(outVertical)"), "vIn": (21, "barn(inVertical)")}
SHAPES = {"circle": (6, "circle"), "box": (4, "box"), "diamond": (8, "diamond"), "plus": (13, "plus")}
STRETCH_SUB = {"across": 10, "left": 8, "right": 2, "top": 1, "bottom": 4}
ENTRANCE_IDS = {"appear": 1, "flyIn": 2, "randomBars": 14, "fade": 10, "split": 16, "stretch": 17, "wheel": 21, "wipe": 22,
                "zoom": 53, "bounce": 26, "growTurn": 31, "pinwheel": 35, "rise": 37, "floatIn": 42, "swivel": 45, "spinner": 49,
                "expand": 55}
EMPHASIS_IDS = {"fillColor": 1, "fontColor": 3, "growShrink": 6, "lineColor": 7, "spin": 8, "transparency": 9, "boldFlash": 10,
                "underline": 18, "darken": 24, "desaturate": 25, "pulse": 26, "colorPulse": 27, "brighten": 30, "teeter": 32,
                "wave": 34, "blink": 35}
# What only HTML draws → the nearest PowerPoint effect (until the part is recorded as a video).
NEAREST_IN = {"layoutSpin": ("swivel", {"dir": "h"}), "layoutSwirl": ("pinwheel", {}),
              "layoutFlip": ("swivel", {"dir": "v"}),
              "kineticZoom": ("zoom", {"dir": "in"}), "kineticFlip": ("swivel", {"dir": "h"}),
              "typewriter": ("appear", {"by": "char"}), "decode": ("appear", {"by": "char"}), "maskRise": ("rise", {}),
              "blurIn": ("fade", {}), "draw": ("wipe", {"dir": "left"}), "countUp": ("fade", {}), "chartGrow": ("wipe", {"dir": "bottom"}),
              "arrive3d": ("flyIn", {"dir": "bottom"})}
NEAREST_EM = {"shine": "pulse", "ripple": "pulse", "marker": "colorPulse", "spotlight": "pulse", "turntable3d": "spin",
              "swing3d": "teeter", "jump3d": "pulse"}
MEDIA_IDS = {"play": 1, "pause": 2, "stop": 3}

# The studio's easing (the effect's own when the entry names none) → PowerPoint's smooth start and end.
EASES = {"smooth": (50000, 50000), "out": (0, 100000), "in": (100000, 0), "linear": (0, 0), "bounce": (0, 100000)}
DEFAULT_EASE = {"appear": "linear", "fade": "smooth", "flyIn": "out", "floatIn": "out", "wipe": "linear", "split": "linear",
                "shape": "out", "wheel": "linear", "randomBars": "linear", "zoom": "out", "swivel": "out", "bounce": "linear",
                "growTurn": "out", "stretch": "out", "rise": "out", "spinner": "out", "pinwheel": "out", "expand": "out"}

# ---------------------------------------------------------------- transitions

# The studio's transition → PowerPoint's element (ns, name) and how its direction option is written.
TRANSITIONS = {"fade": ("p", "fade"), "slide": ("p", "cover"), "push": ("p", "push"), "wipe": ("p", "wipe"), "zoom": ("p", "zoom"),
               "circle": ("p", "circle"), "blinds": ("p", "blinds"), "curtain": ("p", "split"), "flip": ("p14", "flip"),
               "dive": ("p14", "flythrough"), "drill": ("p", "zoom"), "morph": ("p159", "morph")}
TRANSITION_DEFAULT_DIR = {"slide": "right", "push": "down", "wipe": "right", "flip": "left", "blinds": "horizontal", "curtain": "vertical"}
# The studio names where the new slide comes from; PowerPoint the way it moves.
TOWARD = {"right": "l", "left": "r", "down": "u", "up": "d"}
TRANSITION_MS = {"wipe": 920, "circle": 920, "drill": 560, "push": 680, "flip": 940, "dive": 860, "blinds": 920, "curtain": 920, "morph": 900}
P14 = "http://schemas.microsoft.com/office/powerpoint/2010/main"
P159 = "http://schemas.microsoft.com/office/powerpoint/2015/09/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _r(v: float, digits: int = 4) -> str:
    """A number as PowerPoint writes it in a formula (no exponent, no trailing zeros)."""
    text = f"{v:.{digits}f}".rstrip("0").rstrip(".")
    return text if text not in ("", "-0") else "0"


def transition_xml(slide: dict[str, Any], sound_rel: Callable[[str], str | None] | None = None) -> str | None:
    """The slide's p:transition (in mc:AlternateContent when PowerPoint 2010's or 2016's), or None for none."""
    kind = slide.get("transition") or "fade"
    advance = slide.get("advance")
    adv = f' advTm="{int(round(float(advance) * 1000))}"' if isinstance(advance, (int, float)) and advance >= 0 else ""
    sound = slide.get("transitionSound")
    snd = ""
    if sound == "stop":
        snd = "<p:sndAc><p:endSnd/></p:sndAc>"
    elif sound and sound_rel:
        rid = sound_rel(sound)
        if rid:
            snd = f'<p:sndAc><p:stSnd><p:snd r:embed="{rid}" name="{sound}.wav"/></p:stSnd></p:sndAc>'
    if kind == "none" or kind not in TRANSITIONS:
        if not adv and not snd:
            return None
        return f"<p:transition{adv}>{snd}</p:transition>"
    ns, name = TRANSITIONS[kind]
    way = slide.get("transitionDir") or TRANSITION_DEFAULT_DIR.get(kind)
    attrs = ""
    if kind in ("slide", "push", "wipe"):
        attrs = f' dir="{TOWARD.get(way, "l")}"'
    elif kind == "flip":
        attrs = f' dir="{"r" if way == "right" else "l"}"'
    elif kind == "blinds":
        attrs = f' dir="{"vert" if way == "vertical" else "horz"}"'
    elif kind == "curtain":
        attrs = f' orient="{"horz" if way == "horizontal" else "vert"}" dir="out"'
    elif kind == "dive":
        attrs = ' dir="in"'
    elif kind == "morph":
        attrs = ' option="byObject"'
    dur = slide.get("transitionDur")
    ms = int(dur) if isinstance(dur, (int, float)) and dur > 0 else TRANSITION_MS.get(kind, 600)
    spd = "fast" if ms <= 500 else "med" if ms <= 750 else "slow"
    effect = f"<{ns}:{name}{attrs}/>"
    if ns == "p":
        # PowerPoint 2010 and later read the exact duration (p14:dur); older ones the speed.
        plain = f'<p:transition spd="{spd}"{adv}>{effect}{snd}</p:transition>'
        timed = f'<p:transition spd="{spd}" p14:dur="{ms}"{adv}>{effect}{snd}</p:transition>'
        return (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:p14="{P14}" Requires="p14">{timed}</mc:Choice>'
                f"<mc:Fallback>{plain}</mc:Fallback></mc:AlternateContent>")
    requires = "p159" if ns == "p159" else "p14"
    decl = f'xmlns:p14="{P14}"' + (f' xmlns:p159="{P159}"' if ns == "p159" else "")
    choice = f'<p:transition spd="{spd}" p14:dur="{ms}"{adv}>{effect}{snd}</p:transition>'
    return (f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice {decl} Requires="{requires}">{choice}</mc:Choice>'
            f'<mc:Fallback><p:transition spd="{spd}"{adv}><p:fade/>{snd}</p:transition></mc:Fallback></mc:AlternateContent>')


# ---------------------------------------------------------------- targets

class Target:
    """What an effect moves: one or more shapes (a group the writer could not keep as one), with their box."""

    def __init__(self, spids: list[int], box: dict[str, float], *, kind: str = "shape", paras: list[int] | None = None,
                 has_text: bool = False, media: str | None = None, has_bg: bool = False):
        self.spids = spids
        self.has_bg = has_bg
        self.box = box
        self.kind = kind
        # Units per paragraph: (letters, words), for "by paragraph / word / letter".
        self.paras = paras or []
        self.has_text = has_text
        self.media = media


def sp_tgt(spid: int, para: int | None = None) -> str:
    if para is None:
        return f'<p:spTgt spid="{spid}"/>'
    return f'<p:spTgt spid="{spid}"><p:txEl><p:pRg st="{para}" end="{para}"/></p:txEl></p:spTgt>'


# ---------------------------------------------------------------- behaviours

class Builder:
    def __init__(self):
        self.n = 0

    def id(self) -> int:
        self.n += 1
        return self.n

    def bhvr(self, tgt: str, dur: int, names: tuple[str, ...] = (), *, delay: int = 0, fill: str | None = "hold",
             additive: str | None = None, override: str | None = None, extra: str = "") -> str:
        attrs = (f' additive="{additive}"' if additive else "") + (f' override="{override}"' if override else "")
        fill_attr = f' fill="{fill}"' if fill else ""
        st = f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>' if delay else ""
        names_xml = ("<p:attrNameLst>" + "".join(f"<p:attrName>{n}</p:attrName>" for n in names) + "</p:attrNameLst>") if names else ""
        return f'<p:cBhvr{attrs}><p:cTn id="{self.id()}" dur="{dur}"{fill_attr}{extra}>{st}</p:cTn><p:tgtEl>{tgt}</p:tgtEl>{names_xml}</p:cBhvr>'

    def set(self, tgt: str, attr: str, value: str, *, delay: int = 0, dur: int = 1, fill: str = "hold", extra: str = "") -> str:
        return f'<p:set>{self.bhvr(tgt, dur, (attr,), delay=delay, fill=fill, extra=extra)}<p:to><p:strVal val="{value}"/></p:to></p:set>'

    def anim(self, tgt: str, attr: str, keys: list[tuple[int, str]], dur: int, *, delay: int = 0, additive: str | None = None,
             fmla: str | None = None, calc: str = "lin", value_type: str = "num") -> str:
        tavs = ""
        for i, (tm, val) in enumerate(keys):
            f_attr = f' fmla="{fmla}"' if fmla and i == 0 else ""
            tavs += f'<p:tav tm="{tm}"{f_attr}><p:val><p:strVal val="{val}"/></p:val></p:tav>'

        return (f'<p:anim calcmode="{calc}" valueType="{value_type}">{self.bhvr(tgt, dur, (attr,), delay=delay, additive=additive)}'
                f"<p:tavLst>{tavs}</p:tavLst></p:anim>")

    def effect(self, tgt: str, dur: int, filt: str, transition: str = "in", *, delay: int = 0, pr: str | None = None) -> str:
        pr_attr = f' prLst="{pr}"' if pr else ""
        return f'<p:animEffect transition="{transition}" filter="{filt}"{pr_attr}>{self.bhvr(tgt, dur, delay=delay, fill=None)}</p:animEffect>'

    def scale(self, tgt: str, dur: int, x: float, y: float, *, delay: int = 0, extra: str = "") -> str:
        return f'<p:animScale>{self.bhvr(tgt, dur, delay=delay, extra=extra)}<p:by x="{int(round(x * 1000))}" y="{int(round(y * 1000))}"/></p:animScale>'

    def rot(self, tgt: str, dur: int, degrees: float, *, delay: int = 0) -> str:
        return f'<p:animRot by="{int(round(degrees * 60000))}">{self.bhvr(tgt, dur, ("r",), delay=delay)}</p:animRot>'

    def color(self, tgt: str, attr: str, dur: int, *, to: str | None = None, hsl: tuple[int, int, int] | None = None,
              delay: int = 0, extra: str = "") -> str:
        target = f'<p:by><p:hsl h="{hsl[0]}" s="{hsl[1]}" l="{hsl[2]}"/></p:by>' if hsl else f'<p:to><a:srgbClr val="{(to or "#ffc000").lstrip("#").upper()}"/></p:to>'
        return f'<p:animClr clrSpc="{"hsl" if hsl else "rgb"}" dir="cw">{self.bhvr(tgt, dur, (attr,), delay=delay, override="childStyle", extra=extra)}{target}</p:animClr>'

    def motion(self, tgt: str, dur: int, path: str) -> str:
        return (f'<p:animMotion origin="layout" path="{path}" pathEditMode="relative" ptsTypes="">'
                f'{self.bhvr(tgt, dur, ("ppt_x", "ppt_y"))}<p:rCtr x="0" y="0"/></p:animMotion>')

    def cmd(self, tgt: str, command: str) -> str:
        return f'<p:cmd type="call" cmd="{command}">{self.bhvr(tgt, 1)}</p:cmd>'


# ---------------------------------------------------------------- one effect's behaviours

class Effect:
    """The context one effect is drawn with: its target, length, direction, amount, colour, and whether it leaves."""

    def __init__(self, b: Builder, tgt: str, e: dict[str, Any], box: dict[str, float]):
        self.b = b
        self.tgt = tgt
        self.e = e
        self.dur = max(1, int(e.get("dur") or 500))
        self.out = e.get("cls") == "out"
        self.dir = e.get("dir")
        self.box = box

    # Visibility: an entrance shows the part as it starts, an exit hides it as it ends.
    def vis(self) -> str:
        if self.out:
            return self.b.set(self.tgt, "style.visibility", "hidden", delay=max(0, self.dur - 1))
        return self.b.set(self.tgt, "style.visibility", "visible")

    def fade(self) -> str:
        return self.b.effect(self.tgt, self.dur, "fade", "out" if self.out else "in")

    def filt(self, name: str) -> str:
        return self.b.effect(self.tgt, self.dur, name, "out" if self.out else "in")

    def keys(self, attr: str, keys: list[tuple[int, str]], *, fmla: str | None = None) -> str:
        """Key values from hidden (0) to shown (100000); an exit runs them backwards."""
        if self.out:
            keys = [(100000 - tm, val) for tm, val in reversed(keys)]
        return self.b.anim(self.tgt, attr, keys, self.dur, additive="base", fmla=fmla)

    def move(self, attr: str, hidden: str, shown: str) -> str:
        return self.keys(attr, [(0, hidden), (100000, shown)])


def entrance(c: Effect, fx: str) -> tuple[int, int, list[str]]:
    """(presetID, presetSubtype, behaviours) of an entrance (or, with c.out, the exit of the same name)."""
    d = c.dir or ""
    box = c.box
    if fx == "appear":
        return 1, 0, [c.vis()]
    if fx == "fade":
        return 10, 0, [c.vis(), c.fade()]
    if fx == "flyIn":
        d = d or "bottom"
        x = "0-#ppt_w/2" if "left" in d.lower() else "1+#ppt_w/2" if "right" in d.lower() else "#ppt_x"
        y = "0-#ppt_h/2" if "top" in d.lower() else "1+#ppt_h/2" if "bottom" in d.lower() else "#ppt_y"
        return 2, FLY_SUB.get(d, 4), [c.vis(), c.move("ppt_x", x, "#ppt_x"), c.move("ppt_y", y, "#ppt_y")]
    if fx == "floatIn":
        down = d == "down"
        return (47 if down else 42), 0, [c.vis(), c.fade(), c.move("ppt_y", f"#ppt_y{'-' if down else '+'}{_r(120 / SLIDE_H)}", "#ppt_y")]
    if fx == "wipe":
        sub, way = WIPE_IN.get(d or "bottom", WIPE_IN["bottom"])
        return 22, (WIPE_OUT_SUB.get(d or "bottom", 1) if c.out else sub), [c.vis(), c.filt(f"wipe({way})")]
    if fx == "split":
        sub, name = SPLIT.get(d or "vOut", SPLIT["vOut"])
        return 16, sub, [c.vis(), c.filt(name)]
    if fx == "shape":
        kind = (d or "circleOut").replace("Out", "").replace("In", "")
        pid, name = SHAPES.get(kind, SHAPES["circle"])
        way = "out" if (d or "circleOut").endswith("Out") else "in"
        return pid, 32 if way == "out" else 16, [c.vis(), c.filt(f"{name}({way})")]
    if fx == "wheel":
        spokes = int(d) if str(d).isdigit() else 1
        return 21, spokes, [c.vis(), c.filt(f"wheel({spokes})")]
    if fx == "randomBars":
        vertical = d == "v"
        return 14, 5 if vertical else 10, [c.vis(), c.filt(f"randombar({'vertical' if vertical else 'horizontal'})")]
    if fx == "zoom":
        if d == "out":
            return 23, 32, [c.vis(), c.move("ppt_w", "#ppt_w*1.8", "#ppt_w"), c.move("ppt_h", "#ppt_h*1.8", "#ppt_h"), c.fade()]
        if d == "center":
            return 53, 528, [c.vis(), c.move("ppt_w", "#ppt_w*0.1", "#ppt_w"), c.move("ppt_h", "#ppt_h*0.1", "#ppt_h"),
                             c.move("ppt_x", "0.5", "#ppt_x"), c.move("ppt_y", "0.5", "#ppt_y"), c.fade()]
        return 53, 16, [c.vis(), c.move("ppt_w", "#ppt_w*0.3", "#ppt_w"), c.move("ppt_h", "#ppt_h*0.3", "#ppt_h"), c.fade()]
    if fx == "swivel":
        attr = "ppt_h" if d == "v" else "ppt_w"
        size = "#ppt_h" if d == "v" else "#ppt_w"
        # The width (or height) follows the turn: 720° is four half-turns, ending face on.
        fmla = f"{size}*sin(4.5*pi*(1-$))" if c.out else f"{size}*sin(4.5*pi*$)"
        return 45, 10 if d != "v" else 5, [c.vis(), c.b.anim(c.tgt, attr, [(0, "0"), (100000, "1")], c.dur, fmla=fmla), c.fade()]
    if fx == "bounce":
        drop = (box["y"] + box["h"] + 8) / SLIDE_H
        lift = min(160, box["h"] * 0.5 + 40) / SLIDE_H
        steps = [(0, -drop), (38000, 0), (55000, -lift), (70000, 0), (81000, -lift * 0.35), (90000, 0), (95000, -lift * 0.1), (100000, 0)]
        return 26, 0, [c.vis(), c.keys("ppt_y", [(tm, f"#ppt_y{'+' if v >= 0 else '-'}{_r(abs(v))}" if v else "#ppt_y") for tm, v in steps])]
    if fx == "growTurn":
        return 31, 0, [c.vis(), c.move("ppt_w", "0", "#ppt_w"), c.move("ppt_h", "0", "#ppt_h"), c.move("style.rotation", "#ppt_r-90", "#ppt_r"), c.fade()]
    if fx == "stretch":
        d = d or "across"
        parts = [c.vis()]
        if d in ("top", "bottom"):
            parts.append(c.move("ppt_h", "0", "#ppt_h"))
            parts.append(c.move("ppt_y", f"#ppt_y{'-' if d == 'top' else '+'}#ppt_h/2", "#ppt_y"))
        else:
            parts.append(c.move("ppt_w", "0", "#ppt_w"))
            if d != "across":
                parts.append(c.move("ppt_x", f"#ppt_x{'-' if d == 'left' else '+'}#ppt_w/2", "#ppt_x"))
        return 17, STRETCH_SUB.get(d, 10), parts
    if fx == "rise":
        return 37, 0, [c.vis(), c.fade(), c.keys("ppt_y", [(0, f"#ppt_y+{_r(140 / SLIDE_H)}"), (80000, f"#ppt_y-{_r(12 / SLIDE_H)}"), (100000, "#ppt_y")])]
    if fx == "spinner":
        return 49, 0, [c.vis(), c.move("ppt_w", "#ppt_w*0.6", "#ppt_w"), c.move("ppt_h", "#ppt_h*0.6", "#ppt_h"), c.move("style.rotation", "#ppt_r-180", "#ppt_r"), c.fade()]
    if fx == "pinwheel":
        return 35, 0, [c.vis(), c.move("ppt_w", "0", "#ppt_w"), c.move("ppt_h", "0", "#ppt_h"), c.move("style.rotation", "#ppt_r-720", "#ppt_r"), c.fade()]
    if fx == "expand":
        return 55, 0, [c.vis(), c.move("ppt_w", "#ppt_w*0.6", "#ppt_w"), c.fade()]
    return 10, 0, [c.vis(), c.fade()]


def emphasis(c: Effect, fx: str) -> tuple[int, int, list[str]]:
    b, t, dur, e = c.b, c.tgt, c.dur, c.e
    amount = float(e.get("amount") or 0)
    color = e.get("color")
    if fx == "pulse":
        return 26, 0, [b.scale(t, max(1, dur // 2), 108, 108, extra=' autoRev="1"')]
    if fx == "colorPulse":
        return 27, 0, [b.color(t, "fillcolor", max(1, dur // 2), to=color or "#ffc000", extra=' autoRev="1"'),
                       b.set(t, "fill.type", "solid", dur=dur), b.set(t, "fill.on", "true", dur=dur)]
    if fx == "teeter":
        # Back and forth: +7°, -7°, +5°, -5°, home (as the studio swings it).
        swings = [7, -14, 12, -10, 5]
        step = max(1, dur // len(swings))
        return 32, 0, [b.rot(t, step, deg, delay=i * step) for i, deg in enumerate(swings)]
    if fx == "spin":
        deg = (amount or 360) * (-1 if c.dir == "ccw" else 1)
        return 8, 0, [b.rot(t, dur, deg)]
    if fx == "growShrink":
        k = amount or 150
        x, y = (k, 100) if c.dir == "h" else (100, k) if c.dir == "v" else (k, k)
        return 6, 0, [b.scale(t, dur, x, y)]
    if fx == "transparency":
        alpha = _r(1 - (amount or 50) / 100, 2)
        return 9, 0, [b.set(t, "style.opacity", alpha, dur=dur), b.effect(t, dur, "image", "in", pr=f"opacity: {alpha} ")]
    if fx == "fillColor":
        return 1, 2, [b.color(t, "fillcolor", dur, to=color or "#ffc000"), b.set(t, "fill.type", "solid", dur=dur), b.set(t, "fill.on", "true", dur=dur)]
    if fx == "lineColor":
        return 7, 2, [b.color(t, "stroke.color", dur, to=color or "#c00000"), b.set(t, "stroke.on", "true", dur=dur)]
    if fx == "fontColor":
        return 3, 2, [b.color(t, "style.color", dur, to=color or "#c00000")]
    if fx in ("brighten", "darken", "desaturate"):
        hsl = {"brighten": (0, 0, 30000), "darken": (0, 0, -30000), "desaturate": (0, -100000, 0)}[fx]
        return EMPHASIS_IDS[fx], 0, [b.color(t, "fillcolor", dur, hsl=hsl), b.color(t, "style.color", dur, hsl=hsl)]
    if fx == "boldFlash":
        return 10, 0, [b.set(t, "style.fontWeight", "bold", dur=max(1, int(dur * 0.8)), delay=int(dur * 0.1), fill="remove")]
    if fx == "blink":
        return 35, 0, [b.anim(t, "style.opacity", [(0, "1"), (25000, "0"), (50000, "1"), (75000, "0"), (100000, "1")], dur)]
    if fx == "wave":
        return 34, 0, [b.anim(t, "ppt_y", [(0, "#ppt_y"), (50000, f"#ppt_y-{_r(28 / SLIDE_H)}"), (100000, "#ppt_y")], dur, additive="base")]
    if fx == "underline":
        return 18, 0, [b.set(t, "style.textDecorationUnderline", "true", dur=dur)]
    return 26, 0, [b.scale(t, max(1, dur // 2), 108, 108, extra=' autoRev="1"')]


def path_string(path: dict[str, Any]) -> str:
    """The studio's motion path (pixel offsets from where the part is) in PowerPoint's units (fractions of the slide)."""
    pts = [(float(x), float(y)) for x, y in path.get("pts") or []]
    if path.get("closed") and pts and pts[0] != pts[-1]:
        pts.append(pts[0])
    fx = lambda p: f"{_r(p[0] / SLIDE_W, 5)} {_r(p[1] / SLIDE_H, 5)}"  # noqa: E731
    if not pts:
        return "M 0 0 E"
    out = [f"M {fx(pts[0])}"]
    if path.get("curve") and len(pts) > 2:
        # The studio draws a Catmull-Rom curve through the points: the same curve as cubic Béziers.
        n = len(pts)

        def at(i: int):
            if path.get("closed"):
                return pts[(i + n - 1) % (n - 1)]
            return pts[max(0, min(n - 1, i))]

        for i in range(n - 1):
            p0 = pts[n - 2] if path.get("closed") and i == 0 else at(i - 1)
            p1, p2 = pts[i], pts[i + 1]
            p3 = pts[1] if path.get("closed") and i == n - 2 else at(i + 2)
            c1 = (p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6)
            c2 = (p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6)
            out.append(f"C {fx(c1)} {fx(c2)} {fx(p2)}")
    else:
        out.extend(f"L {fx(p)}" for p in pts[1:])
    return " ".join(out) + " E"


# ---------------------------------------------------------------- the whole timing

NODE = {"click": "clickEffect", "with": "withEffect", "after": "afterEffect"}


class Timing:
    """Writes one slide's p:timing and p:bldLst."""

    def __init__(self, targets: dict[str, Target], sound_rel: Callable[[str], str | None] | None = None):
        self.b = Builder()
        self.targets = targets
        self.sound_rel = sound_rel
        self.builds: dict[tuple[int, int], str] = {}
        self.grp: dict[int, int] = {}
        self.approximated = 0
        self.skipped = 0

    # ---- one effect (it may become several: one per shape of a split group, one per paragraph)
    def next_grp(self, spid: int) -> int:
        """Every effect on a shape is its own build (grpId), numbered per shape as PowerPoint does."""
        self.grp[spid] = self.grp.get(spid, -1) + 1
        return self.grp[spid]

    def nearest(self, e: dict[str, Any], target: Target) -> tuple[dict[str, Any], str]:
        """The effect as PowerPoint can draw it: HTML-only ones become the nearest preset (and are counted)."""
        cls, fx = e.get("cls"), e.get("fx") or ""
        e = dict(e)
        if cls in ("in", "out") and fx in NEAREST_IN:
            near, opts = NEAREST_IN[fx]
            for k, v in opts.items():
                if k != "by":
                    e[k] = v
                elif target.kind in ("shape", "text") and any(u and u[0] for u in target.paras) and not e.get("by"):
                    # Letter by letter over the effect's own length (a typewriter types its text in that time).
                    letters = max(1, sum(u[0] for u in target.paras if u))
                    e["by"], e["gap"] = v, max(10, int(int(e.get("dur") or 500) / letters))
                    e["dur"] = 1
            if near == "appear":
                e["ease"] = "linear"
            self.approximated += 1
            return e, near
        if cls == "em" and fx in NEAREST_EM:
            self.approximated += 1
            return e, NEAREST_EM[fx]
        return e, fx

    def effect_pars(self, e: dict[str, Any], node: str) -> list[str]:
        target = self.targets.get(e.get("el") or "")
        if not target or not target.spids:
            self.skipped += 1
            return []
        e, fx = self.nearest(e, target)
        cls = e.get("cls")
        by = e.get("by")
        dur = max(1, int(e.get("dur") or 500))
        units: list[tuple[int, int | None, int]] = []  # (spid, paragraph, delay from the effect's own)
        paras = [(i, u) for i, u in enumerate(target.paras) if u and u[0]]
        if by in ("para", "word", "char") and cls in ("in", "out", "em") and len(target.spids) == 1 and paras:
            spid = target.spids[0]
            if by == "para":
                gap = int((500 if dur < 100 else dur) * 0.6)
                units = [(spid, i, k * gap) for k, (i, _) in enumerate(paras)]
            else:
                gap = e.get("gap") or (500 if dur < 100 else dur) * (0.1 if by == "char" else 0.2)
                e["gap"] = gap
                if not target.has_bg:
                    # A text box with no ground: PowerPoint runs through its words itself.
                    units = [(spid, None, 0)]
                else:
                    # A filled shape stays as it is: its paragraphs letter by letter (word by word), each starting
                    # where the one before ended.
                    at = 0.0
                    for i, (letters, words) in paras:
                        units.append((spid, i, int(at)))
                        at += gap * (letters if by == "char" else words)
        else:
            by = None
            units = [(spid, None, 0) for spid in target.spids]
        grps: dict[int, int] = {}
        pars = []
        for k, (spid, para, offset) in enumerate(units):
            if spid not in grps:
                grps[spid] = self.next_grp(spid)
            xml = self.one(e, fx, sp_tgt(spid, para), target, spid, node if k == 0 else "withEffect", delay=int(e.get("delay") or 0) + offset,
                           iterate=by if by in ("word", "char") else None, paragraph=para is not None, grp=grps[spid])
            if xml:
                pars.append(xml)
        return pars

    def one(self, e: dict[str, Any], fx: str, tgt: str, target: Target, spid: int, node: str, *, delay: int, iterate: str | None,
            paragraph: bool, grp: int) -> str:
        cls = e.get("cls")
        b = self.b
        c = Effect(b, tgt, e, target.box)
        if cls in ("in", "out"):
            pid, sub, parts = entrance(c, fx)
            klass = "entr" if cls == "in" else "exit"
        elif cls == "em":
            pid, sub, parts = emphasis(c, fx)
            klass = "emph"
        elif cls == "path":
            pid, sub, parts, klass = 0, 0, [b.motion(tgt, c.dur, path_string(e.get("path") or {}))], "path"
        elif cls == "media":
            if not target.media:
                self.skipped += 1
                return ""
            fx = fx if fx in MEDIA_IDS else "play"
            command = {"play": "playFrom(0.0)", "pause": "togglePause", "stop": "stop"}[fx]
            return self.wrap(MEDIA_IDS[fx], 0, "mediacall", [b.cmd(tgt, command)], node, delay=delay, grp=None)
        else:
            self.skipped += 1
            return ""
        # How long the effect keeps going: repeats, back and forth, until the next click or the slide's end.
        attrs = ""
        repeat = e.get("repeat")
        times = int(repeat) if isinstance(repeat, (int, float)) and repeat > 1 else 1
        if times > 1:
            attrs += f' repeatCount="{times * 1000}"'
        elif repeat in ("click", "slide"):
            attrs += ' repeatCount="indefinite"'
        if e.get("autoReverse"):
            attrs += ' autoRev="1"'
        ease = e.get("ease") or (DEFAULT_EASE.get(fx, "smooth") if cls in ("in", "out") else "smooth")
        if cls == "out" and not e.get("ease"):
            ease = {"out": "in", "in": "out"}.get(ease, ease)
        accel, decel = EASES.get(ease, EASES["smooth"])
        if accel:
            attrs += f' accel="{accel}"'
        if decel:
            attrs += f' decel="{decel}"'
        if e.get("rewind") and cls in ("in", "em", "path"):
            attrs += ' fill="remove"'
        end = '<p:endCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>' if repeat == "click" else ""
        it = ""
        if iterate:
            it = f'<p:iterate type="{"lt" if iterate == "char" else "wd"}"><p:tmAbs val="{int(round(float(e.get("gap") or 50)))}"/></p:iterate>'
        # アニメーションの後: dimmed or hidden once the effect is over.
        after = e.get("after")
        if after in ("dim", "hide") and cls not in ("out", "media"):
            total = c.dur * times * (2 if e.get("autoReverse") else 1)
            if after == "dim":
                parts.append(b.set(tgt, "style.opacity", "0.35", delay=total, extra=' afterEffect="1"'))
            else:
                parts.append(b.set(tgt, "style.visibility", "hidden", delay=total, extra=' afterEffect="1"'))
        sound = e.get("sound")
        rid = self.sound_rel(sound) if sound and self.sound_rel else None
        if target.kind in ("shape", "text"):
            bg = ' animBg="1"' if target.kind == "shape" else ""
            self.builds[(spid, grp)] = (f'<p:bldP spid="{spid}" grpId="{grp}" build="p"/>' if paragraph
                                        else f'<p:bldP spid="{spid}" grpId="{grp}"{bg}/>')
        elif target.kind == "chart":
            self.builds[(spid, grp)] = f'<p:bldGraphic spid="{spid}" grpId="{grp}"><p:bldAsOne/></p:bldGraphic>'
        return self.wrap(pid, sub, klass, parts, node, delay=delay, grp=grp, attrs=attrs, end=end, iterate=it, sound=(rid, sound) if rid else None)

    def wrap(self, pid: int, sub: int, klass: str, parts: list[str], node: str, *, delay: int, grp: int | None, attrs: str = "",
             end: str = "", iterate: str = "", sound: tuple[str, str] | None = None) -> str:
        cid = self.b.id()
        fill = "" if ' fill="' in attrs else ' fill="hold"'
        grp_attr = f' grpId="{grp}"' if grp is not None else ""
        sub_tn = ""
        if sound:
            # The effect's sound (効果のオプション → サウンド), starting with it.
            rid, name = sound
            sub_tn = (f'<p:subTnLst><p:audio><p:cMediaNode vol="80000" showWhenStopped="0"><p:cTn id="{self.b.id()}" display="0" masterRel="sameClick">'
                      f'<p:stCondLst><p:cond evt="begin" delay="0"><p:tn val="{cid}"/></p:cond></p:stCondLst>'
                      f'<p:endCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst></p:cTn>'
                      f'<p:tgtEl><p:sndTgt r:embed="{rid}" name="{name}.wav"/></p:tgtEl></p:cMediaNode></p:audio></p:subTnLst>')
        return (f'<p:par><p:cTn id="{cid}" presetID="{pid}" presetClass="{klass}" presetSubtype="{sub}"{fill}{grp_attr} nodeType="{node}"{attrs}>'
                f'<p:stCondLst><p:cond delay="{delay}"/></p:stCondLst>{end}{iterate}<p:childTnLst>{"".join(parts)}</p:childTnLst>{sub_tn}</p:cTn></p:par>')

    # ---- a click group: its moments, each with its effects
    def click_group(self, group: list[dict[str, Any]], *, first_node: str | None = None, hide_before: list[dict[str, Any]] | None = None) -> str:
        moments: list[tuple[int, list[str]]] = []
        prev = None
        for i, e in enumerate(group):
            dur = max(1, int(e.get("dur") or 500))
            repeat = e.get("repeat")
            times = int(repeat) if isinstance(repeat, (int, float)) and repeat >= 1 else 1
            if prev is None:
                base = 0
            elif e.get("start") == "with":
                base = prev["base"]
            else:
                base = prev["end"]
            begin = base + int(e.get("delay") or 0)
            end = begin + dur * times * (2 if e.get("autoReverse") else 1)
            item = {"base": base, "end": end}
            if i == 0:
                node = first_node or NODE.get(e.get("start") or "click", "clickEffect")
            else:
                node = NODE["with"] if e.get("start") == "with" else NODE["after"]
            pars = self.effect_pars(e, node)
            if not pars:
                # Nothing to draw (its part did not make it): the effects after it keep their times.
                prev = item
                continue
            if moments and moments[-1][0] == base:
                moments[-1][1].extend(pars)
            else:
                moments.append((base, pars))
            prev = item
        # 次のアニメーションのクリック時に非表示にする: the parts the last click left go as this one starts.
        if hide_before:
            hides = []
            for e in hide_before:
                target = self.targets.get(e.get("el") or "")
                for spid in (target.spids if target else []):
                    hides.append(self.b.set(sp_tgt(spid), "style.visibility", "hidden", extra=' afterEffect="1"'))
            if hides:
                par = f'<p:par><p:cTn id="{self.b.id()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{"".join(hides)}</p:childTnLst></p:cTn></p:par>'
                if moments and moments[0][0] == 0:
                    moments[0][1].insert(0, par)
                else:
                    moments.insert(0, (0, [par]))
        if not moments:
            return ""
        inner = "".join(
            f'<p:par><p:cTn id="{self.b.id()}" fill="hold"><p:stCondLst><p:cond delay="{at}"/></p:stCondLst><p:childTnLst>{"".join(pars)}</p:childTnLst></p:cTn></p:par>'
            for at, pars in moments)
        return inner

    def outer(self, inner: str, *, arrival: bool = False) -> str:
        cid = self.b.id()
        cond = '<p:cond delay="indefinite"/>' + ('<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>' if arrival else "")
        return f'<p:par><p:cTn id="{cid}" fill="hold"><p:stCondLst>{cond}</p:stCondLst><p:childTnLst>{inner}</p:childTnLst></p:cTn></p:par>'

    def write(self, timeline: list[dict[str, Any]], media: list[dict[str, Any]]) -> str | None:
        """The slide's p:timing (or None when nothing moves or plays)."""
        main: list[list[dict[str, Any]]] = [[]]
        triggers: dict[str, list[list[dict[str, Any]]]] = {}
        # A video or a sound that plays as the slide arrives starts first.
        for m in media:
            if m.get("autoplay"):
                main[0].append({"el": m["id"], "cls": "media", "fx": "play", "start": "with" if main[0] else "after", "dur": 1})
        for e in timeline:
            trig = e.get("trigger")
            if trig:
                if "@" in trig:
                    # ブックマーク時: PowerPoint starts these from the media's bookmark; until written, they play on a click of the media.
                    trig = trig.split("@")[0]
                    self.approximated += 1
                seq = triggers.setdefault(trig, [])
                if not seq or e.get("start") == "click":
                    seq.append([])
                seq[-1].append(e)
            elif e.get("start") == "click":
                main.append([e])
            else:
                main[-1].append(e)
        b = self.b
        b.n = 2  # 1 is the root, 2 the main sequence (an arrival's onBegin names it)
        groups = []
        hide_next: list[dict[str, Any]] = []
        for k, group in enumerate(main):
            if not group and k == 0:
                continue
            inner = self.click_group(group, hide_before=hide_next if k else None)
            hide_next = [e for e in group if e.get("after") == "hideNext" and e.get("cls") not in ("out", "media")]
            if inner:
                groups.append(self.outer(inner, arrival=k == 0))
        seqs = []
        for trig, seq in triggers.items():
            target = self.targets.get(trig)
            if not target or not target.spids:
                self.skipped += sum(len(g) for g in seq)
                continue
            seqs.append(self.interactive(target.spids[0], [self.click_group(g, first_node="clickEffect") for g in seq]))
        # A video plays and pauses as it is clicked (PowerPoint's 自動 / クリック時 for an inserted video), unless the
        # slide already starts it with a trigger of its own.
        for m in media:
            target = self.targets.get(m["id"])
            # A filmed HTML effect plays on its own moment only (a click on it does nothing).
            if not target or m["id"] in triggers or m.get("clip"):
                continue
            spid = target.spids[0]
            toggle = self.wrap(MEDIA_IDS["pause"], 0, "mediacall", [b.cmd(sp_tgt(spid), "togglePause")], "clickEffect", delay=0, grp=None)
            inner = f'<p:par><p:cTn id="{b.id()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{toggle}</p:childTnLst></p:cTn></p:par>'
            seqs.append(self.interactive(spid, [inner]))
        nodes = "".join(self.media_node(m) for m in media if self.targets.get(m["id"]))
        if not groups and not seqs and not nodes:
            return None
        main_seq = ""
        if groups:
            main_seq = (f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>{"".join(groups)}</p:childTnLst></p:cTn>'
                        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
                        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>')
        builds = "".join(self.builds[k] for k in sorted(self.builds))
        bld = f"<p:bldLst>{builds}</p:bldLst>" if builds else ""
        return (f'<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
                f"{main_seq}{''.join(seqs)}{nodes}</p:childTnLst></p:cTn></p:par></p:tnLst>{bld}</p:timing>")

    def interactive(self, spid: int, clicks: list[str]) -> str:
        b = self.b
        cid = b.id()
        body = "".join(f'<p:par><p:cTn id="{b.id()}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>{inner}</p:childTnLst></p:cTn></p:par>'
                       for inner in clicks if inner)
        click = f'<p:cond evt="onClick" delay="0"><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cond>'
        return (f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="{cid}" restart="whenNotActive" fill="hold" evtFilter="cancelBubble" nodeType="interactiveSeq">'
                f'<p:stCondLst>{click}</p:stCondLst><p:endSync evt="end" delay="0"><p:rtn val="all"/></p:endSync><p:childTnLst>{body}</p:childTnLst></p:cTn>'
                f"<p:nextCondLst>{click}</p:nextCondLst></p:seq>")

    def media_node(self, m: dict[str, Any]) -> str:
        """A video's or a sound's playing options (再生): volume, across slides, loop, hidden while stopped, full screen."""
        target = self.targets[m["id"]]
        spid = target.spids[0]
        vol = m.get("volume")
        attrs = f' vol="{int(round(float(vol) * 100000))}"' if isinstance(vol, (int, float)) and vol < 1 else ' vol="80000"'
        if m.get("muted"):
            attrs += ' mute="1"'
        if m.get("across"):
            attrs += f' numSld="{int(m["across"])}"'
        if m.get("hideIdle") or m.get("hideIcon"):
            attrs += ' showWhenStopped="0"'
        repeat = ' repeatCount="indefinite"' if m.get("loop") else ""
        tag = "audio" if m.get("kind") == "audio" else "video"
        full = ' fullScrn="1"' if tag == "video" and m.get("fullscreen") else ""
        end = ('<p:endCondLst><p:cond evt="onStopAudio" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:endCondLst>' if tag == "audio" else "")
        node = (f'<p:cMediaNode{attrs}><p:cTn id="{self.b.id()}"{repeat} fill="hold" display="0"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst>{end}</p:cTn>'
                f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cMediaNode>')
        return f"<p:{tag}{full}>{node}</p:{tag}>"
