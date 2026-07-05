"""Write a real theme into the composed deck so the *master* carries the brand.

'Follow the slide master' — we formalize the derived BrandSpec into the presentation's theme
(fontScheme major/minor + clrScheme). The composer then styles from the same tokens, so the
whole deck is governed by one design system rather than 275 ad-hoc shapes.
"""
from __future__ import annotations
from lxml import etree
from pptx.oxml import parse_xml
from pptx.oxml.ns import qn
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from .brand import BrandSpec


def _set_srgb(parent, hexv: str):
    """Replace parent's color child with an <a:srgbClr val=..>."""
    for child in list(parent):
        parent.remove(child)
    el = parent.makeelement(qn("a:srgbClr"), {"val": hexv.lstrip("#")})
    parent.append(el)


def apply_theme(prs, brand: BrandSpec) -> None:
    master = prs.slide_masters[0]
    theme_part = master.part.part_related_by(RT.THEME)
    root = parse_xml(theme_part.blob)  # <a:theme>

    fs = root.find(".//" + qn("a:fontScheme"))
    if fs is not None:
        for tag, font in (("a:majorFont", brand.heading_font), ("a:minorFont", brand.body_font)):
            latin = fs.find(qn(tag) + "/" + qn("a:latin"))
            if latin is not None:
                latin.set("typeface", font)

    cs = root.find(".//" + qn("a:clrScheme"))
    if cs is not None:
        mapping = {
            "dk1": "000000", "lt1": "FFFFFF",
            "dk2": brand.dark, "lt2": brand.light,
            "accent1": brand.accent, "accent2": brand.secondary,
            "accent3": brand.dark, "accent4": brand.neutral,
            "accent5": brand.light, "accent6": brand.accent,
            "hlink": brand.secondary, "folHlink": brand.neutral,
        }
        for name, hexv in mapping.items():
            el = cs.find(qn("a:" + name))
            if el is None:
                continue
            # dk1/lt1 use sysClr; keep those, only recolor the scheme accents/dk2/lt2
            if name in ("dk1", "lt1"):
                continue
            _set_srgb(el, hexv)

    theme_part._blob = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
