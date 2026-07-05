"""Compose the output deck from the source IR + BrandSpec.

python-pptx only (no LibreOffice). Every source text run is re-placed verbatim and restyled to
the design system; nothing is invented. Text-dense bodies use PowerPoint auto-fit (normAutofit)
so the deck never overflows even though we render nothing here — PowerPoint sizes it on open.

v1 scope: full design system (background, eyebrow, serif headline, subtitle, body, footer, tables)
applied consistently to all slides. Pictures are carried in a side rail; richer components
(icon-rings, card grids, textures, chevrons) are the next elevation pass.
"""
from __future__ import annotations
import io
import math

from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.oxml.ns import qn

from .ir import DeckIR, ShapeIR, ParaIR
from .brand import BrandSpec
from .classify import classify_slide, SlidePlan
from . import color as C

IN = 914400


def _emu(v_in: float) -> Emu:
    return Emu(int(round(v_in * IN)))


def _rgb(hex_: str) -> RGBColor:
    return RGBColor.from_string(hex_.lstrip("#"))


# ---------------------------------------------------------------- low-level draw

def _no_line(shape):
    shape.line.fill.background()


def _solid(shape, hex_):
    shape.fill.solid()
    shape.fill.fore_color.rgb = _rgb(hex_)


def bg(slide, deck: DeckIR, hex_):
    r = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, deck.width_emu, deck.height_emu)
    _solid(r, hex_)
    _no_line(r)
    return r


def rrect(slide, x, y, w, h, fill=None, radius=0.06, line=None, line_w=1.0):
    sp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, _emu(x), _emu(y), _emu(w), _emu(h))
    try:
        sp.adjustments[0] = radius
    except Exception:
        pass
    if fill:
        _solid(sp, fill)
    else:
        sp.fill.background()
    if line:
        sp.line.color.rgb = _rgb(line)
        sp.line.width = Pt(line_w)
    else:
        _no_line(sp)
    sp.shadow.inherit = False
    return sp


def hrule(slide, x, y, w, hex_, weight=1.2, dash=None):
    ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, _emu(x), _emu(y), _emu(x + w), _emu(y))
    ln.line.color.rgb = _rgb(hex_)
    ln.line.width = Pt(weight)
    if dash:
        lnel = ln.line._get_or_add_ln()
        lnel.append(lnel.makeelement(qn("a:prstDash"), {"val": dash}))
    return ln


def icon_ring(slide, x, y, d, ring_hex, icon_blob=None):
    ov = slide.shapes.add_shape(MSO_SHAPE.OVAL, _emu(x), _emu(y), _emu(d), _emu(d))
    ov.fill.background()
    ov.line.color.rgb = _rgb(ring_hex)
    ov.line.width = Pt(1.6)
    ov.shadow.inherit = False
    placed = False
    if icon_blob:
        s = d * 0.5
        off = (d - s) / 2
        try:
            slide.shapes.add_picture(io.BytesIO(icon_blob), _emu(x + off), _emu(y + off), _emu(s), _emu(s))
            placed = True
        except Exception:
            pass
    if not placed:  # accent dot so empty rings read as intentional
        dd = d * 0.30
        o = (d - dd) / 2
        dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, _emu(x + o), _emu(y + o), _emu(dd), _emu(dd))
        _solid(dot, ring_hex)
        _no_line(dot)
        dot.shadow.inherit = False
    return ov


def bg_image(slide, deck, blob):
    try:
        slide.shapes.add_picture(io.BytesIO(blob), 0, 0, deck.width_emu, deck.height_emu)
    except Exception:
        pass


def _set_tracking(run, pts: float):
    """Letter-spacing (a:rPr @spc in 1/100 pt) — not exposed by python-pptx."""
    run._r.get_or_add_rPr().set("spc", str(int(pts * 100)))


def add_text(slide, x, y, w, h, paras, *, anchor=MSO_ANCHOR.TOP, autofit=False, wrap=True):
    """paras: list of dicts:
       {runs:[{text,color,bold,italic,font,size}], size, color, font, bold, italic,
        align, space_after, space_before, line_spacing, tracking}
    """
    tb = slide.shapes.add_textbox(_emu(x), _emu(y), _emu(w), _emu(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, 0)
    if autofit:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for i, spec in enumerate(paras):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = spec.get("align", PP_ALIGN.LEFT)
        if spec.get("space_after") is not None:
            p.space_after = Pt(spec["space_after"])
        if spec.get("space_before") is not None:
            p.space_before = Pt(spec["space_before"])
        p.line_spacing = spec.get("line_spacing", 1.08)
        base_size = spec.get("size", 13)
        base_color = spec.get("color")
        base_font = spec.get("font")
        for rspec in spec["runs"]:
            r = p.add_run()
            r.text = rspec["text"]
            f = r.font
            f.name = rspec.get("font", base_font)
            f.size = Pt(rspec.get("size", base_size))
            b = rspec.get("bold", spec.get("bold"))
            if b is not None:
                f.bold = b
            it = rspec.get("italic", spec.get("italic"))
            if it is not None:
                f.italic = it
            col = rspec.get("color", base_color)
            if col:
                f.color.rgb = _rgb(col)
            if spec.get("tracking"):
                _set_tracking(r, spec["tracking"])
    return tb


# ---------------------------------------------------------------- text estimation

def _est_lines(text: str, size_pt: float, width_in: float) -> int:
    if not text.strip():
        return 1
    char_w_in = size_pt * 0.50 / 72.0
    cpl = max(8, int(width_in / char_w_in))
    lines = 0
    for seg in text.split("\n"):
        lines += max(1, math.ceil(len(seg) / cpl))
    return lines


def _block_h_in(text, size_pt, width_in, line_spacing=1.12, pad=0.06) -> float:
    return _est_lines(text, size_pt, width_in) * (size_pt * line_spacing / 72.0) + pad


# ---------------------------------------------------------------- body rendering

def _para_spec_one(para: ParaIR, B: BrandSpec, base_color: str, size=None,
                   space_after=7, lead_color=None) -> dict | None:
    """One styled paragraph spec, preserving runs. Bold lead-in -> accent."""
    runs = [r for r in para.runs if r.text != ""]
    if not runs:
        return None
    has_following_plain = any(not r.bold for r in runs[1:])
    lead_color = lead_color or B.accent
    rspecs = []
    for j, r in enumerate(runs):
        is_lead = bool(r.bold) and j == 0 and has_following_plain
        rspecs.append({"text": r.text, "color": lead_color if is_lead else base_color,
                       "bold": bool(r.bold)})
    return {"runs": rspecs, "size": size or B.scale["body"], "font": B.body_font,
            "color": base_color, "line_spacing": 1.14, "space_after": space_after}


def _para_specs(shape: ShapeIR, B: BrandSpec) -> list[dict]:
    out = []
    for para in shape.paras:
        s = _para_spec_one(para, B, B.ink)
        if s:
            out.append(s)
    return out


def _body_items(plan: SlidePlan) -> list[ParaIR]:
    """Flatten body into a list of non-empty paragraphs (one per 'item')."""
    items = []
    for sh in plan.body:
        for para in sh.paras:
            if para.text.strip():
                items.append(para)
    return items


def _para_len(para: ParaIR) -> int:
    return len(para.text.strip())


def _is_label(para: ParaIR) -> bool:
    """A short, (near) all-caps standalone label like 'WHAT AI71 BRINGS'."""
    t = para.text.strip()
    if not t or len(t) > 34:
        return False
    lowers = sum(1 for c in t if c.islower())
    return lowers <= 2


def _section_label(slide, B, x, y, w, para, on_dark):
    spec = {"runs": [{"text": para.text.strip(), "color": B.accent, "bold": True}],
            "size": B.scale["eyebrow"] + 1, "font": B.body_font, "tracking": 1.4, "space_after": 0}
    tb = add_text(slide, x, y, 2.6, 0.3, [spec])
    rule_col = "34597B" if on_dark else C.mix(B.neutral, "FFFFFF", 0.35)
    hrule(slide, x + 2.7, y + 0.14, w - 2.7, rule_col, 0.9)


def _render_icon_list(slide, deck, B, x, y, w, h, items, icons, on_dark):
    n = max(1, len(items))
    row_h = h / n
    text_color = "FFFFFF" if on_dark else B.ink
    lead_color = B.accent
    div_color = "34597B" if on_dark else C.mix(B.neutral, "FFFFFF", 0.35)
    d = min(0.66, row_h * 0.62)
    tx = x + d + 0.28
    tw = w - (d + 0.28) - 0.1
    for i, para in enumerate(items):
        ry = y + i * row_h
        icon_ring(slide, x, ry + (row_h - d) / 2, d, B.accent, icons[i] if i < len(icons) else None)
        spec = _para_spec_one(para, B, text_color, lead_color=lead_color, space_after=0)
        if spec:
            add_text(slide, tx, ry, tw, row_h, [spec], anchor=MSO_ANCHOR.MIDDLE, autofit=True)
        if i < n - 1:
            hrule(slide, tx, ry + row_h, tw, div_color, 0.9, dash="dash")


def _render_card_grid(slide, deck, B, x, y, w, h, items):
    n = len(items)
    cols = 3 if n in (3, 6, 9) else (2 if n in (2, 4) else min(3, n))
    rows = math.ceil(n / cols)
    gap = 0.22
    cw = (w - gap * (cols - 1)) / cols
    ch = (h - gap * (rows - 1)) / rows
    for i, para in enumerate(items):
        r, c = divmod(i, cols)
        cx, cy = x + c * (cw + gap), y + r * (ch + gap)
        rrect(slide, cx, cy, cw, ch, fill=B.light, radius=0.08)
        # accent tab
        rrect(slide, cx, cy, 0.09, ch, fill=B.accent, radius=0.0)
        spec = _para_spec_one(para, B, B.ink, size=B.scale["body"], space_after=0)
        if spec:
            add_text(slide, cx + 0.24, cy + 0.16, cw - 0.4, ch - 0.32, [spec],
                     anchor=MSO_ANCHOR.TOP, autofit=True)


# ---------------------------------------------------------------- footer

def _footer(slide, deck: DeckIR, plan: SlidePlan, B: BrandSpec, on_dark: bool):
    y = deck.height_in - 0.46
    ml, mr = 0.62, 0.62
    cw = deck.width_in - ml - mr
    col = "FFFFFF" if on_dark else B.neutral
    hrule(slide, ml, y - 0.10, cw, (B.neutral if not on_dark else "31597B"), 0.75)
    # split footer shapes by horizontal position
    W = deck.width_emu
    left = [s for s in plan.footers if (s.left or 0) < 0.45 * W]
    right = [s for s in plan.footers if (s.left or 0) >= 0.45 * W]
    if left:
        add_text(slide, ml, y, cw * 0.6, 0.32,
                 [{"runs": [{"text": "   ".join(s.text.replace(chr(10), " ") for s in left)}],
                   "size": B.scale["footer"], "font": B.body_font, "color": col, "space_after": 0}])
    if right:
        add_text(slide, ml + cw * 0.4, y, cw * 0.6, 0.32,
                 [{"runs": [{"text": "   ".join(s.text.replace(chr(10), " ") for s in right)}],
                   "size": B.scale["footer"], "font": B.body_font, "color": col,
                   "align": PP_ALIGN.RIGHT, "space_after": 0}])


# ---------------------------------------------------------------- table

def _table(slide, shape: ShapeIR, x, y, w, B: BrandSpec, max_h: float = 4.6):
    grid = shape.table
    rows, cols = len(grid), max(len(r) for r in grid)
    h = min(max_h, 0.42 * rows + 0.1)
    gt = slide.shapes.add_table(rows, cols, _emu(x), _emu(y), _emu(w), _emu(h)).table
    gt.first_row = True
    for ri, row in enumerate(grid):
        for ci in range(cols):
            cell = gt.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(B.dark if ri == 0 else (B.paper if ri % 2 else B.light))
            runs = row[ci] if ci < len(row) else []
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            if runs:
                for r in runs:
                    rr = p.add_run(); rr.text = r.text
                    rr.font.size = Pt(B.scale["small"])
                    rr.font.name = B.body_font
                    rr.font.bold = (ri == 0)
                    rr.font.color.rgb = _rgb("FFFFFF" if ri == 0 else B.ink)
            else:
                p.text = ""
    return y + h


# ---------------------------------------------------------------- side rail (pictures)

def _place_pictures(slide, plan: SlidePlan, x, y, w, h):
    pics = [p for p in plan.pictures if p.image_blob]
    if not pics:
        return
    n = len(pics)
    ph = h / n
    for i, p in enumerate(pics[:6]):
        try:
            iw = (p.width or IN) / IN
            ih = (p.height or IN) / IN
            ar = iw / ih if ih else 1.0
            draw_h = min(ph - 0.12, w / ar)
            draw_w = draw_h * ar
            slide.shapes.add_picture(io.BytesIO(p.image_blob),
                                     _emu(x + (w - draw_w) / 2), _emu(y + i * ph),
                                     _emu(draw_w), _emu(draw_h))
        except Exception:
            pass


# ---------------------------------------------------------------- slide builders

def _wordmark(slide, text, x, y, color, B, size=17):
    add_text(slide, x, y, 2.2, 0.42,
             [{"runs": [{"text": text}], "size": size, "font": B.heading_font,
               "color": color, "bold": True, "tracking": 1.0, "space_after": 0}])


def _gather_nonchrome(plan: SlidePlan) -> list[ShapeIR]:
    """All text shapes except wordmark/footers, de-duped and in reading order.
    Guarantees every text run gets rendered (protects invariant I2)."""
    cand = []
    for s in (plan.title, plan.eyebrow, plan.subtitle):
        if s is not None:
            cand.append(s)
    cand += plan.body
    seen, out = set(), []
    for s in sorted(cand, key=lambda s: ((s.top or 0), (s.left or 0))):
        if id(s) in seen:
            continue
        seen.add(id(s)); out.append(s)
    return out


def _split_headline(shapes: list[ShapeIR]) -> tuple[str, str]:
    """Big shapes -> headline; the rest -> supporting text. Preserves all text."""
    tmax = max((s.max_size_pt for s in shapes), default=0.0)
    head = [s for s in shapes if tmax > 0 and s.max_size_pt >= 0.55 * tmax]
    rest = [s for s in shapes if s not in head]
    headline = " ".join(s.text.replace("\n", " ") for s in head).strip()
    support = "   ".join(s.text.replace("\n", " ") for s in rest).strip()
    return headline, support


def build_title(slide, deck, plan, B, icons=None, texture=None):
    bg(slide, deck, B.dark)
    if texture:
        bg_image(slide, deck, texture)
    rrect(slide, deck.width_in - 1.5, deck.height_in - 1.5, 0.9, 0.9, fill=None,
          radius=0.5, line=B.accent, line_w=2.0)
    ml = 0.9
    if plan.wordmark:
        _wordmark(slide, plan.wordmark.text.strip(), ml, 0.55, "FFFFFF", B, size=18)
    headline, support = _split_headline(_gather_nonchrome(plan))
    y = deck.height_in * 0.30
    if headline:
        th = _block_h_in(headline, B.scale["display"], deck.width_in - 2 * ml, 1.06)
        add_text(slide, ml, y, deck.width_in - 2 * ml, th + 0.2,
                 [{"runs": [{"text": headline}], "size": B.scale["display"], "font": B.heading_font,
                   "color": "FFFFFF", "line_spacing": 1.04, "space_after": 0}])
        y += th + 0.35
    if support:
        add_text(slide, ml, y, deck.width_in - 2 * ml, 1.4,
                 [{"runs": [{"text": support}], "size": B.scale["subtitle"], "font": B.body_font,
                   "color": "C7D2DE", "line_spacing": 1.22, "space_after": 0}])
    _footer(slide, deck, plan, B, on_dark=True)


def build_closing(slide, deck, plan, B, icons=None, texture=None):
    bg(slide, deck, B.dark)
    if texture:
        bg_image(slide, deck, texture)
    rrect(slide, deck.width_in - 1.5, deck.height_in - 1.5, 0.9, 0.9, fill=None,
          radius=0.5, line=B.accent, line_w=2.0)
    if plan.wordmark:
        _wordmark(slide, plan.wordmark.text.strip(), 0.9, 0.55, "FFFFFF", B, size=18)
    headline, support = _split_headline(_gather_nonchrome(plan))
    add_text(slide, 0.9, deck.height_in * 0.38, deck.width_in - 1.8, 1.4,
             [{"runs": [{"text": headline}], "size": B.scale["display"], "font": B.heading_font,
               "color": "FFFFFF", "space_after": 6}] +
             ([{"runs": [{"text": support}], "size": B.scale["subtitle"], "font": B.body_font,
                "color": "C7D2DE", "space_after": 0}] if support else []),
             anchor=MSO_ANCHOR.MIDDLE)


def build_section(slide, deck, plan, B, icons=None, texture=None):
    bg(slide, deck, B.dark)
    if texture:
        bg_image(slide, deck, texture)
    ml = 0.9
    if plan.eyebrow:
        add_text(slide, ml, deck.height_in * 0.34, deck.width_in - 2 * ml, 0.35,
                 [{"runs": [{"text": plan.eyebrow.text.strip()}], "size": B.scale["eyebrow"],
                   "font": B.body_font, "color": B.secondary, "bold": True, "tracking": 2.2,
                   "space_after": 0}])
    shapes = [s for s in _gather_nonchrome(plan) if s is not plan.eyebrow]
    headline, support = _split_headline(shapes)
    add_text(slide, ml, deck.height_in * 0.42, deck.width_in - 2 * ml, 2.0,
             [{"runs": [{"text": headline}], "size": B.scale["display"], "font": B.heading_font,
               "color": "FFFFFF", "line_spacing": 1.04, "space_after": 8}] +
             ([{"runs": [{"text": support}], "size": B.scale["subtitle"], "font": B.body_font,
                "color": "C7D2DE", "space_after": 0}] if support else []))
    _footer(slide, deck, plan, B, on_dark=True)


def build_content(slide, deck, plan, B, icons=None, texture=None):
    icons = icons or []
    bg(slide, deck, B.paper)
    ml, mr, mt = 0.62, 0.62, 0.5
    cw = deck.width_in - ml - mr
    y = mt
    # eyebrow
    if plan.eyebrow:
        add_text(slide, ml, y, cw, 0.3,
                 [{"runs": [{"text": plan.eyebrow.text.strip()}], "size": B.scale["eyebrow"],
                   "font": B.body_font, "color": B.secondary, "bold": True, "tracking": 2.0,
                   "space_after": 0}])
        y += 0.32
    # title
    if plan.title:
        th = _block_h_in(plan.title.text, B.scale["h1"], cw, 1.06)
        add_text(slide, ml, y, cw, th + 0.15,
                 [{"runs": [{"text": plan.title.text.replace("\n", " ")}], "size": B.scale["h1"],
                   "font": B.heading_font, "color": B.dark, "line_spacing": 1.05, "space_after": 0}])
        y += th + 0.10
    hrule(slide, ml, y, 0.7, B.accent, 2.4)
    y += 0.16
    # subtitle
    if plan.subtitle:
        sh = _block_h_in(plan.subtitle.text, B.scale["subtitle"], cw, 1.2)
        add_text(slide, ml, y, cw, sh + 0.1,
                 [{"runs": [{"text": plan.subtitle.text.replace("\n", " ")}], "size": B.scale["subtitle"],
                   "font": B.body_font, "color": B.neutral, "line_spacing": 1.18, "space_after": 0}])
        y += sh + 0.16

    body_bottom = deck.height_in - 0.62
    items = _body_items(plan)

    # a long, non-lead leading paragraph becomes a callout bar
    callout = None
    if items and _para_len(items[0]) >= 90 and not (items[0].runs and items[0].runs[0].bold):
        callout = items[0]
        items = items[1:]
    if callout:
        chh = _block_h_in(callout.text, B.scale["body"], cw - 0.5, 1.2) + 0.2
        chh = min(chh, 1.3)
        rrect(slide, ml, y, cw, chh, fill=B.light, radius=0.10)
        spec = _para_spec_one(callout, B, B.ink, space_after=0)
        add_text(slide, ml + 0.25, y + 0.1, cw - 0.5, chh - 0.2, [spec], anchor=MSO_ANCHOR.MIDDLE, autofit=True)
        y += chh + 0.22

    avail = body_bottom - y
    lens = [_para_len(p) for p in items] or [0]
    on_dark = bool(plan.eyebrow and plan.eyebrow.text.strip().upper().startswith("PHASE"))

    if plan.tables:
        # table slide: compact body list above the table
        if items:
            specs = [s for p in items if (s := _para_spec_one(p, B, B.ink))]
            bh = min(avail * 0.45, 2.2)
            add_text(slide, ml, y, cw, max(0.5, bh), specs, anchor=MSO_ANCHOR.TOP, autofit=True)
            y += bh + 0.18
        for tb in plan.tables:
            a = body_bottom - y
            if a < 0.9:
                y = body_bottom - 0.9; a = 0.9
            y = _table(slide, tb, ml, y, cw, B, max_h=a) + 0.15
    elif items and 3 <= len(items) <= 6 and max(lens) <= 115:
        _render_card_grid(slide, deck, B, ml, y, cw, min(avail, 3.6), items)
    elif items and 2 <= len(items) <= 7 and sum(lens) <= 1500:
        # icon-list inside a callout card (dark for PHASE slides, light otherwise)
        card_fill = B.dark if on_dark else B.light
        rrect(slide, ml, y, cw, avail, fill=card_fill, radius=0.06)
        pad = 0.34
        iy = y + pad
        ih = avail - 2 * pad
        # a leading label ('WHAT AI71 BRINGS') becomes a section header + rule
        if items and _is_label(items[0]):
            _section_label(slide, B, ml + pad, iy, cw - 2 * pad, items[0], on_dark)
            iy += 0.42
            ih -= 0.42
            items = items[1:]
        _render_icon_list(slide, deck, B, ml + pad, iy, cw - 2 * pad, ih, items, icons, on_dark)
    elif items:
        specs = [s for p in items if (s := _para_spec_one(p, B, B.ink))]
        add_text(slide, ml, y, cw, max(0.6, avail), specs, anchor=MSO_ANCHOR.TOP, autofit=True)

    _footer(slide, deck, plan, B, on_dark=False)


BUILDERS = {"title": build_title, "closing": build_closing, "section": build_section,
            "content": build_content}


def compose(deck: DeckIR, brand: BrandSpec) -> Presentation:
    from .designsystem import apply_theme
    from .assets import make_contour_texture, recolor_icon
    prs = Presentation()
    prs.slide_width = Emu(deck.width_emu)
    prs.slide_height = Emu(deck.height_emu)
    apply_theme(prs, brand)   # master carries the derived fonts + palette
    blank = prs.slide_layouts[6]
    texture = make_contour_texture(1280, 720, brand.dark, C.mix(brand.dark, "FFFFFF", 0.22))
    for s in deck.slides:
        plan = classify_slide(s, deck)
        icons = [b for p in plan.pictures if p.image_blob and (b := recolor_icon(p.image_blob, brand.accent))]
        slide = prs.slides.add_slide(blank)
        BUILDERS.get(plan.stype, build_content)(slide, deck, plan, brand, icons, texture)
    return prs
