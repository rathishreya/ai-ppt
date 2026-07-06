"""In-place high-end polish for decks that ALREADY have a real master.

Unlike the rebuild path (compose.py), this keeps the source presentation intact — its masters,
layouts, theme, backgrounds, logo — and only *redesigns the plain text slides* onto that master.
Complex slides (charts, diagrams/groups, pictures, multi-table) are preserved untouched.
"""
from __future__ import annotations
import zipfile
from lxml import etree
from pptx import Presentation

from .ingest import load_deck
from .brand import extract_brand
from .classify import classify_slide
from . import color as C
from . import compose as X

_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _brand_from_source(path, deck):
    """Prefer the deck's REAL theme palette (a designed master) over color clustering."""
    b = extract_brand(deck)
    try:
        z = zipfile.ZipFile(path)
        themes = sorted(n for n in z.namelist() if n.startswith("ppt/theme/") and n.endswith(".xml"))
        t = etree.fromstring(z.read(themes[0]))
        cs = t.find(".//" + _A + "clrScheme")
        col = {}
        for c in cs:
            name = etree.QName(c).localname
            srgb = c.find(_A + "srgbClr")
            sysc = c.find(_A + "sysClr")
            if srgb is not None:
                col[name] = srgb.get("val")
            elif sysc is not None:
                col[name] = sysc.get("lastClr") or "000000"
        # only override if this looks like a real (non-default-Office) theme
        if col.get("accent1", "").upper() not in ("4472C4", ""):
            b.accent = col.get("accent1", b.accent)
            b.secondary = col.get("accent2", b.secondary)
            b.dark = col.get("dk2", b.dark)
            b.ink = b.dark
            if col.get("lt2"):
                b.light = C.mix(col["lt2"], "FFFFFF", 0.35)
            b.neutral = C.mix(b.dark, "FFFFFF", 0.45)
    except Exception:
        pass
    return b


def _is_complex(sir) -> bool:
    kinds = {sh.kind for sh in sir.shapes}
    if kinds & {"picture", "group", "other"}:
        return True
    if sum(1 for sh in sir.shapes if sh.kind == "table") >= 2:
        return True
    return False


def _clear(slide):
    for sp in list(slide.shapes):
        sp._element.getparent().remove(sp._element)


def polish(src: str, out: str):
    prs = Presentation(src)
    deck = load_deck(src)
    B = _brand_from_source(src, deck)
    slides = list(prs.slides)
    redesigned, preserved = [], []
    for i, slide in enumerate(slides):
        if i >= len(deck.slides):
            break
        sir = deck.slides[i]
        if _is_complex(sir):
            preserved.append(i + 1)
            continue
        plan = classify_slide(sir, deck)
        _clear(slide)
        # redesign onto the existing master layout (no bg, no footer — the master provides them)
        X.build_content(slide, deck, plan, B, draw_bg=False, draw_footer=True)
        redesigned.append(i + 1)
    prs.save(out)
    return {"redesigned": redesigned, "preserved": preserved, "brand": B}
