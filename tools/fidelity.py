"""Measures how closely the studio's 見た目どおりに取り込む import looks like the original PowerPoint deck.

qa/fidelity.mjs drives this: it renders the original with LibreOffice (``render``), reads the words the original holds
(``texts``), draws the imported deck in the studio engine, and then ``compare`` scores each slide:

* ``ssim``    structural similarity of the two pictures (0..1, grey, 480 x 270) - does it look the same.
* ``blocks``  how close the coarse light/dark layout is (0..1, 48 x 27 cells) - are things in the right place and colour.
* ``text``    the share of the original's characters the studio's slide still shows (0..1) - was any word lost.

Each slide also gets a picture "original | studio | difference" so the worst ones can be looked at.

    python3 tools/fidelity.py render  DECK.pptx OUTDIR            -> OUTDIR/page-001.png ... + meta.json
    python3 tools/fidelity.py texts   DECK.pptx                   -> JSON on stdout
    python3 tools/fidelity.py compare ORIGDIR STUDIODIR OUTDIR    -> OUTDIR/report.json + OUTDIR/pair-001.png ...
"""

from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata

PAGE_W = 960  # both pictures are drawn this wide (16:9 -> 960 x 540)
CMP_W, CMP_H = 480, 270


def soffice_binary() -> str:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    raise SystemExit("LibreOffice (soffice) is not installed")


def render(deck: str, outdir: str) -> None:
    """The original deck as LibreOffice draws it: one PNG per visible slide."""
    import fitz  # pymupdf

    os.makedirs(outdir, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        # A private profile so parallel runs and a user's own LibreOffice never share a lock.
        profile = os.path.join(tmp, "profile")
        subprocess.run(
            [soffice_binary(), f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "pdf", "--outdir", tmp, deck],
            check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=600,
        )
        pdf = os.path.join(tmp, os.path.splitext(os.path.basename(deck))[0] + ".pdf")
        doc = fitz.open(pdf)
        pages = []
        for i, page in enumerate(doc):
            scale = PAGE_W / page.rect.width
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            name = f"page-{i + 1:03d}.png"
            pix.save(os.path.join(outdir, name))
            pages.append({"file": name, "w": pix.width, "h": pix.height})
        with open(os.path.join(outdir, "meta.json"), "w", encoding="utf-8") as f:
            json.dump({"pages": pages}, f)


def iter_shapes(shapes):
    for shape in shapes:
        yield shape
        if shape.shape_type == 6:  # group
            yield from iter_shapes(shape.shapes)


def texts(deck: str) -> dict:
    """Per visible slide, the pieces of text the original holds (shapes, tables, groups) and how many pictures it shows."""
    from pptx import Presentation

    prs = Presentation(deck)
    out = []
    for slide in prs.slides:
        element = slide._element
        if element.get("show") == "0":
            continue  # a hidden slide is not in the PDF
        pieces, pictures = [], 0
        for shape in iter_shapes(slide.shapes):
            if shape.shape_type == 13:
                pictures += 1
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    value = "".join(run.text for run in para.runs)
                    if value.strip():
                        pieces.append(value)
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for para in cell.text_frame.paragraphs:
                            value = "".join(run.text for run in para.runs)
                            if value.strip():
                                pieces.append(value)
        out.append({"text": pieces, "pictures": pictures})
    return {"slides": out}


def normalise(value: str) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", value))


def text_recall(pieces: list[str], shown: str) -> tuple[float, list[str]]:
    """Share of the original's characters whose whole piece (or, for long ones, most of its words) is still on the slide."""
    shown_n = normalise(shown)
    total = got = 0
    missing = []
    for piece in pieces:
        p = normalise(piece)
        if not p:
            continue
        total += len(p)
        if p in shown_n:
            got += len(p)
            continue
        # The studio may break a paragraph differently (a line break, a bullet mark): count what is there by chunks.
        step = 6
        chunks = [p[i:i + step] for i in range(0, len(p), step)]
        have = sum(len(c) for c in chunks if c in shown_n)
        got += have
        if have < len(p) * 0.8:
            missing.append(piece[:60])
    return (got / total if total else 1.0), missing


def load_gray(path: str, size=(CMP_W, CMP_H)):
    import numpy as np
    from PIL import Image

    img = Image.open(path).convert("RGB")
    return np.asarray(img.resize(size, Image.LANCZOS).convert("L"), dtype=np.float64)


def box_filter(a, k):
    """Mean over k x k windows (valid positions), by cumulative sums."""
    import numpy as np

    c = np.cumsum(np.cumsum(np.pad(a, ((1, 0), (1, 0))), axis=0), axis=1)
    return (c[k:, k:] - c[:-k, k:] - c[k:, :-k] + c[:-k, :-k]) / (k * k)


def ssim(a, b, k=7) -> float:
    c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
    ma, mb = box_filter(a, k), box_filter(b, k)
    saa = box_filter(a * a, k) - ma * ma
    sbb = box_filter(b * b, k) - mb * mb
    sab = box_filter(a * b, k) - ma * mb
    m = ((2 * ma * mb + c1) * (2 * sab + c2)) / ((ma * ma + mb * mb + c1) * (saa + sbb + c2))
    return float(m.mean())


def blocks(a, b, grid=(48, 27)) -> float:
    import numpy as np

    def cells(x):
        h, w = x.shape
        gh, gw = grid[1], grid[0]
        return x[: h // gh * gh, : w // gw * gw].reshape(gh, h // gh, gw, w // gw).mean(axis=(1, 3))

    return float(1 - np.abs(cells(a) - cells(b)).mean() / 255)


def pair_image(orig: str, studio: str, out: str, score: float) -> None:
    """original | studio | difference, side by side (the difference is brighter where they disagree)."""
    import numpy as np
    from PIL import Image, ImageDraw

    a = Image.open(orig).convert("RGB").resize((PAGE_W // 2, PAGE_W * 9 // 32))
    b = Image.open(studio).convert("RGB").resize(a.size)
    diff = np.abs(np.asarray(a, dtype=np.int16) - np.asarray(b, dtype=np.int16)).max(axis=2)
    heat = np.clip(255 - diff * 3, 0, 255).astype(np.uint8)
    d = Image.fromarray(np.stack([np.full_like(heat, 255), heat, heat], axis=2))
    sheet = Image.new("RGB", (a.width * 3 + 20, a.height + 24), (40, 40, 40))
    for i, im in enumerate((a, b, d)):
        sheet.paste(im, (i * (a.width + 10), 24))
    draw = ImageDraw.Draw(sheet)
    for i, label in enumerate(("original (LibreOffice)", "studio", f"difference  ssim {score:.3f}")):
        draw.text((i * (a.width + 10) + 4, 6), label, fill=(255, 255, 255))
    sheet.save(out)


def compare(orig_dir: str, studio_dir: str, outdir: str) -> dict:
    os.makedirs(outdir, exist_ok=True)
    meta = json.load(open(os.path.join(orig_dir, "meta.json"), encoding="utf-8"))
    studio = json.load(open(os.path.join(studio_dir, "studio.json"), encoding="utf-8"))
    originals = json.load(open(os.path.join(orig_dir, "texts.json"), encoding="utf-8"))["slides"]
    rows = []
    count = min(len(meta["pages"]), len(studio["slides"]))
    for i in range(count):
        orig_png = os.path.join(orig_dir, meta["pages"][i]["file"])
        studio_png = os.path.join(studio_dir, studio["slides"][i]["file"])
        a, b = load_gray(orig_png), load_gray(studio_png)
        s = ssim(a, b)
        recall, missing = text_recall(originals[i]["text"] if i < len(originals) else [], studio["slides"][i].get("text", ""))
        pair = f"pair-{i + 1:03d}.png"
        pair_image(orig_png, studio_png, os.path.join(outdir, pair), s)
        pictures = originals[i]["pictures"] if i < len(originals) else 0
        rows.append({
            "slide": i + 1,
            "ssim": round(s, 4),
            "blocks": round(blocks(a, b), 4),
            "text": round(recall, 4),
            "missing": missing[:5],
            "pictures": pictures,
            "studioPictures": studio["slides"][i].get("pictures", 0),
            "pair": pair,
        })
    summary = {
        "slides": len(rows),
        "originalPages": len(meta["pages"]),
        "studioSlides": len(studio["slides"]),
        "ssimMean": round(sum(r["ssim"] for r in rows) / len(rows), 4) if rows else 0,
        "ssimMin": min((r["ssim"] for r in rows), default=0),
        "blocksMean": round(sum(r["blocks"] for r in rows) / len(rows), 4) if rows else 0,
        "textMean": round(sum(r["text"] for r in rows) / len(rows), 4) if rows else 1,
        "textMin": min((r["text"] for r in rows), default=1),
        "pictures": sum(r["pictures"] for r in rows),
        "studioPictures": sum(r["studioPictures"] for r in rows),
    }
    report = {"summary": summary, "rows": rows}
    with open(os.path.join(outdir, "report.json"), "w", encoding="utf-8") as f:
        json.dump(report, f, ensure_ascii=False, indent=1)
    return report


def main(argv: list[str]) -> int:
    if len(argv) >= 4 and argv[1] == "render":
        render(argv[2], argv[3])
        # The words the original holds are saved beside its pictures so compare needs only the directories.
        with open(os.path.join(argv[3], "texts.json"), "w", encoding="utf-8") as f:
            json.dump(texts(argv[2]), f, ensure_ascii=False)
        return 0
    if len(argv) >= 3 and argv[1] == "texts":
        json.dump(texts(argv[2]), sys.stdout, ensure_ascii=False)
        return 0
    if len(argv) >= 5 and argv[1] == "compare":
        report = compare(argv[2], argv[3], argv[4])
        json.dump(report["summary"], sys.stdout, ensure_ascii=False)
        return 0
    print(__doc__)
    return 2


if __name__ == "__main__":
    sys.exit(main(sys.argv))
