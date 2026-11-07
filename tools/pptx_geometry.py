"""Where things are and how big their text is in a PowerPoint file (shared by the import tools)."""

from __future__ import annotations

from pptx.oxml.ns import qn


def _xfrm_map(group):
    """Child → parent coordinate mapping of a group shape (groups keep their members in their own space)."""
    xfrm = group._element.find(qn("p:grpSpPr") + "/" + qn("a:xfrm"))
    if xfrm is None:
        return lambda x, y, w, h: (x, y, w, h)

    def val(tag, a, b):
        el = xfrm.find(qn(tag))
        return (int(el.get(a, 0)), int(el.get(b, 0))) if el is not None else (0, 0)

    (ox, oy), (ex, ey) = val("a:off", "x", "y"), val("a:ext", "cx", "cy")
    (cx, cy), (cw, ch) = val("a:chOff", "x", "y"), val("a:chExt", "cx", "cy")
    sx, sy = (ex / cw if cw else 1.0), (ey / ch if ch else 1.0)
    return lambda x, y, w, h: (ox + (x - cx) * sx, oy + (y - cy) * sy, w * sx, h * sy)


def _level_size(container, level: int) -> float | None:
    """Font size (pt) a list-style container (lstStyle / txStyles entry) gives paragraphs of `level`."""
    if container is None:
        return None
    for tag in (f"a:lvl{level + 1}pPr", "a:lvl1pPr", "a:defPPr"):
        ppr = container.find(qn(tag))
        rpr = ppr.find(qn("a:defRPr")) if ppr is not None else None
        if rpr is not None and rpr.get("sz"):
            return int(rpr.get("sz")) / 100
    return None


def _inherited_size(shape, slide, level: int) -> float:
    """Where PowerPoint looks for a paragraph's size when its runs do not say: the shape's own list
    style, then its layout and master placeholders, then the master's text styles, then the
    presentation default (18pt)."""
    body = getattr(shape, "_element", None)
    body = body.find(".//" + qn("a:lstStyle")) if body is not None else None
    size = _level_size(body, level)
    if size:
        return size
    master = slide.slide_layout.slide_master
    styles = master._element.find(qn("p:txStyles"))
    if getattr(shape, "is_placeholder", False) and shape.is_placeholder:
        try:
            idx, kind = shape.placeholder_format.idx, shape.placeholder_format.type
            layout_ph = next((ph for ph in slide.slide_layout.placeholders if ph.placeholder_format.idx == idx), None)
            for ph in (layout_ph, getattr(layout_ph, "_base_placeholder", None)):
                if ph is not None:
                    size = _level_size(ph._element.find(".//" + qn("a:lstStyle")), level)
                    if size:
                        return size
            style = "p:titleStyle" if kind in (1, 3) else "p:bodyStyle"
            size = _level_size(styles.find(qn(style)) if styles is not None else None, level)
            if size:
                return size
        except (AttributeError, KeyError, ValueError):
            pass
    default = slide.part.package.presentation_part.presentation._element.find(qn("p:defaultTextStyle"))
    return _level_size(default, level) or 18.0


def _paragraph_size(paragraph, shape, slide) -> float:
    for run in paragraph.runs:
        if run.font.size:
            return run.font.size.pt
    end = paragraph._p.find(qn("a:endParaRPr"))
    if end is not None and end.get("sz"):
        return int(end.get("sz")) / 100
    return _inherited_size(shape, slide, paragraph.level)
