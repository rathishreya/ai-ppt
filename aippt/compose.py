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
import copy

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
TITLE_IDX, DATE_IDX, FOOTER_IDX = 0, 10, 11


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
    # send to back so it sits behind the layout's placeholders
    spTree = slide.shapes._spTree
    spTree.remove(r._element)
    spTree.insert(2, r._element)
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
        pic = slide.shapes.add_picture(io.BytesIO(blob), 0, 0, deck.width_emu, deck.height_emu)
        spTree = slide.shapes._spTree
        spTree.remove(pic._element)
        spTree.insert(2, pic._element)  # behind placeholders
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
    _write_paras(tf, paras)
    return tb


def _write_paras(tf, paras):
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


# ---- placeholder helpers (so the deck follows the master) ----

def _get_ph(slide, idx):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


def _clone_ph(slide, idx):
    ph = _get_ph(slide, idx)
    if ph is not None:
        return ph
    for lph in slide.slide_layout.placeholders:
        if lph.placeholder_format.idx == idx:
            slide.shapes._spTree.append(copy.deepcopy(lph._element))
            return _get_ph(slide, idx)
    return None


def _del_ph_except(slide, keep):
    for ph in list(slide.placeholders):
        if ph.placeholder_format.idx not in keep:
            ph._element.getparent().remove(ph._element)


def _fill_ph(slide, idx, x, y, w, h, paras, *, anchor=MSO_ANCHOR.TOP, autofit=False, clone=False):
    """Put text into a real placeholder (repositioned/styled) so the slide follows the master."""
    ph = _clone_ph(slide, idx) if clone else _get_ph(slide, idx)
    if ph is None:
        return add_text(slide, x, y, w, h, paras, anchor=anchor, autofit=autofit)
    ph.left, ph.top, ph.width, ph.height = _emu(x), _emu(y), _emu(w), _emu(h)
    tf = ph.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
        setattr(tf, m, Emu(0))
    if autofit:
        tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    _write_paras(tf, paras)
    return ph


# ---------------------------------------------------------------- text estimation

def _est_lines(text: str, size_pt: float, width_in: float, cw_factor: float = 0.52) -> int:
    if not text.strip():
        return 1
    char_w_in = size_pt * cw_factor / 72.0
    cpl = max(8, int(width_in / char_w_in))
    lines = 0
    for seg in text.split("\n"):
        lines += max(1, math.ceil(len(seg) / cpl))
    return lines


def _block_h_in(text, size_pt, width_in, line_spacing=1.16, pad=0.08, cw_factor=0.52) -> float:
    return _est_lines(text, size_pt, width_in, cw_factor) * (size_pt * line_spacing / 72.0) + pad


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
    from .icons import icon_for
    d = min(0.72, row_h * 0.66)
    arrow_x = x + d + 0.12
    asz = 0.06
    tx = x + d + 0.42
    tw = w - (tx - x) - 0.05
    for i, para in enumerate(items):
        ry = y + i * row_h
        cy = ry + row_h / 2
        blob = icon_for(para.text, B.accent, i)
        icon_ring(slide, x, cy - d / 2, d, B.accent, blob)
        _chevron_arrow(slide, arrow_x, cy, asz, B.accent)
        spec = _para_spec_one(para, B, text_color, lead_color=lead_color, space_after=0)
        if spec:
            add_text(slide, tx, ry, tw, row_h, [spec], anchor=MSO_ANCHOR.MIDDLE, autofit=True)
        if i < n - 1:
            hrule(slide, tx, ry + row_h, tw, div_color, 0.9, dash="dash")


def _chevron_arrow(slide, cx, cy, s, hex_):
    for (x1, y1), (x2, y2) in (((cx, cy - s), (cx + s, cy)), ((cx + s, cy), (cx, cy + s))):
        ln = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, _emu(x1), _emu(y1), _emu(x2), _emu(y2))
        ln.line.color.rgb = _rgb(hex_)
        ln.line.width = Pt(2.0)


def _render_flow(slide, deck, B, x, y, w, h, items):
    """Horizontal chevron flowchart for short step/process sequences."""
    n = len(items)
    h = min(h, 1.7)
    overlap = 0.22
    cw = (w + overlap * (n - 1)) / n
    palette = [B.dark, B.secondary]
    for i, para in enumerate(items):
        cx = x + i * (cw - overlap)
        sp = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, _emu(cx), _emu(y), _emu(cw), _emu(h))
        _solid(sp, palette[i % 2])
        _no_line(sp)
        sp.shadow.inherit = False
        tf = sp.text_frame
        tf.word_wrap = True
        for m in ("margin_left", "margin_right", "margin_top", "margin_bottom"):
            setattr(tf, m, Emu(0))
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        for r in para.runs:
            if r.text:
                run = p.add_run(); run.text = r.text
                run.font.name = B.body_font; run.font.size = Pt(B.scale["small"])
                run.font.bold = True; run.font.color.rgb = _rgb("FFFFFF")
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE


def _group_for_cards(items):
    """Merge a short 'header' paragraph with the following longer one so cards carry a
    heading + body (fewer, better-padded cards instead of many cramped fragments)."""
    groups, i = [], 0
    while i < len(items):
        cur = items[i]
        nxt = items[i + 1] if i + 1 < len(items) else None
        if (_para_len(cur) <= 32 and nxt is not None and _para_len(nxt) > 32
                and not (cur.runs and cur.runs[0].bold)):
            groups.append([cur, nxt]); i += 2
        else:
            groups.append([cur]); i += 1
    return groups


def _columns_from_shapes(plan, deck):
    """Detect the source's real column layout from body-shape x-positions.
    Returns (wide_headers, columns) where columns is a left-to-right list of shape-lists
    (each top-ordered), or None if there aren't clear columns."""
    W = deck.width_emu
    body = [s for s in plan.body if s.plain_len > 0]
    narrow = [s for s in body if (s.width or 0) < 0.55 * W]
    wide = [s for s in body if (s.width or 0) >= 0.55 * W]
    if len(narrow) < 3:
        return None
    lefts = sorted(set(round((s.left or 0) / IN, 2) for s in narrow))
    bands = []
    for L in lefts:
        if bands and L - bands[-1][-1] < 1.6:
            bands[-1].append(L)
        else:
            bands.append([L])
    if not (2 <= len(bands) <= 4):
        return None
    centers = [(b[0] + b[-1]) / 2 for b in bands]

    def band_of(s):
        L = (s.left or 0) / IN
        return min(range(len(centers)), key=lambda i: abs(L - centers[i]))

    cols = [[] for _ in bands]
    for s in sorted(narrow, key=lambda s: (s.top or 0)):
        cols[band_of(s)].append(s)
    cols = [c for c in cols if c]
    if len(cols) < 2 or max(len(c) for c in cols) < 1:
        return None
    return wide, cols


def _render_columns(slide, deck, B, x, y, w, h, columns, B_dark=False):
    """Render source columns side by side as light cards (header + items), preserving structure."""
    n = len(columns)
    gap = 0.24
    cw = (w - gap * (n - 1)) / n
    for ci, col in enumerate(columns):
        cx = x + ci * (cw + gap)
        paras = [p for sh in col for p in sh.paras if p.text.strip()]
        if not paras:
            continue
        rrect(slide, cx, y, cw, h, fill=B.light, radius=0.07)
        rrect(slide, cx, y, 0.08, h, fill=B.accent, radius=0.0)
        specs = []
        for j, para in enumerate(paras):
            is_head = (j == 0)
            sp = _para_spec_one(para, B, B.dark if is_head else B.ink,
                                size=(B.scale["body"] + 1 if is_head else B.scale["body"]),
                                space_after=(5 if is_head else 4))
            if sp and is_head:
                for r in sp["runs"]:
                    r["bold"] = True; r["color"] = B.dark
            if sp:
                specs.append(sp)
        add_text(slide, cx + 0.26, y + 0.18, cw - 0.44, h - 0.36, specs,
                 anchor=MSO_ANCHOR.TOP, autofit=True)


def _render_card_grid(slide, deck, B, x, y, w, h, items):
    groups = _group_for_cards(items)
    n = len(groups)
    cols = 3 if n >= 5 else (2 if n in (2, 4) else min(3, max(1, n)))
    rows = math.ceil(n / cols)
    gap = 0.26
    cw = (w - gap * (cols - 1)) / cols
    ch = (h - gap * (rows - 1)) / rows
    padx, pady = 0.32, 0.16
    for i, paras in enumerate(groups):
        r, c = divmod(i, cols)
        cx, cy = x + c * (cw + gap), y + r * (ch + gap)
        rrect(slide, cx, cy, cw, ch, fill=B.light, radius=0.07)
        rrect(slide, cx, cy, 0.08, ch, fill=B.accent, radius=0.0)  # accent tab
        specs = []
        for j, para in enumerate(paras):
            is_head = (j == 0 and len(paras) > 1)
            sp = _para_spec_one(para, B, B.dark if is_head else B.ink,
                                size=B.scale["body"], space_after=(3 if is_head else 0))
            if sp and is_head:
                for rspec in sp["runs"]:
                    rspec["bold"] = True
            if sp:
                specs.append(sp)
        if specs:
            add_text(slide, cx + padx, cy + pady, cw - padx - 0.18, ch - 2 * pady, specs,
                     anchor=MSO_ANCHOR.MIDDLE, autofit=True)


# ---------------------------------------------------------------- footer

def _footer(slide, deck: DeckIR, plan: SlidePlan, B: BrandSpec, on_dark: bool):
    y = deck.height_in - 0.46
    ml, mr = 0.62, 0.62
    cw = deck.width_in - ml - mr
    col = "FFFFFF" if on_dark else B.neutral
    hrule(slide, ml, y - 0.10, cw, (B.neutral if not on_dark else "31597B"), 0.75)
    W = deck.width_emu
    left = [s for s in plan.footers if (s.left or 0) < 0.45 * W]
    right = [s for s in plan.footers if (s.left or 0) >= 0.45 * W]
    # footer text -> real FOOTER / DATE placeholders (cloned) so it follows the master
    if left:
        _fill_ph(slide, DATE_IDX, ml, y, cw * 0.62, 0.32,
                 [{"runs": [{"text": "   ".join(s.text.replace(chr(10), " ") for s in left)}],
                   "size": B.scale["footer"], "font": B.body_font, "color": col, "space_after": 0}],
                 clone=True)
    if right:
        _fill_ph(slide, FOOTER_IDX, ml + cw * 0.38, y, cw * 0.62, 0.32,
                 [{"runs": [{"text": "   ".join(s.text.replace(chr(10), " ") for s in right)}],
                   "size": B.scale["footer"], "font": B.body_font, "color": col,
                   "align": PP_ALIGN.RIGHT, "space_after": 0}], clone=True)


# ---------------------------------------------------------------- table

def _table(slide, shape: ShapeIR, x, y, w, B: BrandSpec, max_h: float = 4.6):
    grid = shape.table
    rows, cols = len(grid), max(len(r) for r in grid)
    h = min(max_h, 0.42 * rows + 0.1)
    # scale cell font so tall tables fit the allotted height (avoid footer collision)
    row_h = h / max(rows, 1)
    fsize = max(7.5, min(B.scale["small"], round(row_h * 72 * 0.28, 1)))
    gt = slide.shapes.add_table(rows, cols, _emu(x), _emu(y), _emu(w), _emu(h)).table
    gt.first_row = True
    from pptx.util import Emu as _E
    for ri, row in enumerate(grid):
        try:
            gt.rows[ri].height = _E(int(row_h * IN))
        except Exception:
            pass
        for ci in range(cols):
            cell = gt.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(B.dark if ri == 0 else (B.paper if ri % 2 else B.light))
            cell.margin_left = Pt(7); cell.margin_right = Pt(7)
            cell.margin_top = Pt(3); cell.margin_bottom = Pt(3)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            runs = row[ci] if ci < len(row) else []
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            if runs:
                for r in runs:
                    rr = p.add_run(); rr.text = r.text
                    rr.font.size = Pt(fsize)
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


def _logo_mark(slide, x, y, d, B):
    ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, _emu(x), _emu(y), _emu(d), _emu(d))
    ring.fill.background(); ring.line.color.rgb = _rgb("FFFFFF"); ring.line.width = Pt(1.5)
    ring.shadow.inherit = False
    ch = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, _emu(x + d * 0.30), _emu(y + d * 0.32),
                                _emu(d * 0.5), _emu(d * 0.36))
    ch.fill.background(); ch.line.color.rgb = _rgb(B.accent); ch.line.width = Pt(2.0)
    ch.shadow.inherit = False


def build_title(slide, deck, plan, B, icons=None, texture=None):
    if texture:
        bg_image(slide, deck, texture)
    else:
        bg(slide, deck, B.dark)
    _logo_mark(slide, deck.width_in - 1.55, deck.height_in - 1.5, 0.95, B)
    ml = 0.9
    if plan.wordmark:
        _wordmark(slide, plan.wordmark.text.strip(), ml, 0.55, "FFFFFF", B, size=20)
    headline, support = _split_headline(_gather_nonchrome(plan))
    hw = deck.width_in * 0.72
    hsize = min(56, round(B.scale["display"] * 1.24))
    y = deck.height_in * 0.24
    if headline:
        th = _block_h_in(headline, hsize, hw, line_spacing=1.1, cw_factor=0.62)
        _fill_ph(slide, TITLE_IDX, ml, y, hw, th + 0.15,
                 [{"runs": [{"text": headline}], "size": hsize, "font": B.heading_font,
                   "color": "FFFFFF", "line_spacing": 1.06, "space_after": 0}])
        y += th + 0.28
    if support:
        add_text(slide, ml, y, deck.width_in * 0.8, 1.4,
                 [{"runs": [{"text": support}], "size": B.scale["subtitle"], "font": B.body_font,
                   "color": "C7D2DE", "line_spacing": 1.24, "space_after": 0}])
    _footer(slide, deck, plan, B, on_dark=True)


def build_closing(slide, deck, plan, B, icons=None, texture=None):
    if texture:
        bg_image(slide, deck, texture)
    else:
        bg(slide, deck, B.dark)
    rrect(slide, deck.width_in - 1.5, deck.height_in - 1.5, 0.9, 0.9, fill=None,
          radius=0.5, line=B.accent, line_w=2.0)
    if plan.wordmark:
        _wordmark(slide, plan.wordmark.text.strip(), 0.9, 0.55, "FFFFFF", B, size=18)
    headline, support = _split_headline(_gather_nonchrome(plan))
    _fill_ph(slide, TITLE_IDX, 0.9, deck.height_in * 0.38, deck.width_in - 1.8, 1.4,
             [{"runs": [{"text": headline}], "size": B.scale["display"], "font": B.heading_font,
               "color": "FFFFFF", "space_after": 0}], anchor=MSO_ANCHOR.MIDDLE)
    if support:
        add_text(slide, 0.9, deck.height_in * 0.56, deck.width_in - 1.8, 1.0,
                 [{"runs": [{"text": support}], "size": B.scale["subtitle"], "font": B.body_font,
                   "color": "C7D2DE", "space_after": 0}])


def build_section(slide, deck, plan, B, icons=None, texture=None):
    if texture:
        bg_image(slide, deck, texture)
    else:
        bg(slide, deck, B.dark)
    ml = 0.9
    if plan.eyebrow:
        add_text(slide, ml, deck.height_in * 0.34, deck.width_in - 2 * ml, 0.35,
                 [{"runs": [{"text": plan.eyebrow.text.strip()}], "size": B.scale["eyebrow"],
                   "font": B.body_font, "color": B.secondary, "bold": True, "tracking": 2.2,
                   "space_after": 0}])
    shapes = [s for s in _gather_nonchrome(plan) if s is not plan.eyebrow]
    headline, support = _split_headline(shapes)
    _fill_ph(slide, TITLE_IDX, ml, deck.height_in * 0.42, deck.width_in - 2 * ml, 1.6,
             [{"runs": [{"text": headline}], "size": B.scale["display"], "font": B.heading_font,
               "color": "FFFFFF", "line_spacing": 1.04, "space_after": 0}])
    if support:
        add_text(slide, ml, deck.height_in * 0.64, deck.width_in - 2 * ml, 1.2,
                 [{"runs": [{"text": support}], "size": B.scale["subtitle"], "font": B.body_font,
                   "color": "C7D2DE", "space_after": 0}])
    _footer(slide, deck, plan, B, on_dark=True)


def build_content(slide, deck, plan, B, icons=None, texture=None):
    icons = icons or []
    bg(slide, deck, B.paper)
    ml, mr, mt = 0.62, 0.62, 0.5
    cw = deck.width_in - ml - mr
    footer_top = deck.height_in - 0.66
    y = mt
    # eyebrow
    if plan.eyebrow:
        add_text(slide, ml, y, cw, 0.3,
                 [{"runs": [{"text": plan.eyebrow.text.strip()}], "size": B.scale["eyebrow"],
                   "font": B.body_font, "color": B.secondary, "bold": True, "tracking": 2.0,
                   "space_after": 0}])
        y += 0.34
    # title -> real TITLE placeholder (follows the master).
    if plan.title:
        th = _block_h_in(plan.title.text, B.scale["h1"], cw, line_spacing=1.1, cw_factor=0.53)
        _fill_ph(slide, TITLE_IDX, ml, y, cw, th + 0.08,
                 [{"runs": [{"text": plan.title.text.replace("\n", " ")}], "size": B.scale["h1"],
                   "font": B.heading_font, "color": B.dark, "line_spacing": 1.06, "space_after": 0}])
        y += th + 0.12
    hrule(slide, ml, y, 0.7, B.accent, 2.4)
    y += 0.18
    items = _body_items(plan)
    cols_data = None if plan.tables else _columns_from_shapes(plan, deck)
    intro = None
    has_label = False
    lens = [0]
    if cols_data:
        mode = "columns"
    else:
        # a long, non-lead leading paragraph = intro (rendered as a callout bar)
        if items and _para_len(items[0]) >= 90 and not (items[0].runs and items[0].runs[0].bold):
            intro = items[0]; items = items[1:]
        lens = [_para_len(p) for p in items] or [0]
        has_label = bool(items and _is_label(items[0]))
        if plan.tables:
            mode = "table"
        elif items and 3 <= len(items) <= 6 and max(lens) <= 40:
            mode = "flow"
        elif items and 2 <= len(items) <= 8 and sum(lens) <= 2100 and max(lens) > 40:
            mode = "iconlist"
        elif items:
            mode = "grid"
        else:
            mode = "none"
    is_phase = bool(plan.eyebrow and plan.eyebrow.text.strip().upper().startswith("PHASE"))
    on_dark = mode == "iconlist" and (is_phase or has_label or intro is not None)

    # subtitle: light callout bar on dark-card slides, plain neutral text otherwise
    if plan.subtitle:
        if on_dark:
            y = _callout_bar(slide, ml, y, cw, plan.subtitle.text.replace("\n", " "), B, bold=True)
        else:
            sh = _block_h_in(plan.subtitle.text, B.scale["subtitle"], cw, line_spacing=1.2, cw_factor=0.53)
            add_text(slide, ml, y, cw, sh + 0.05,
                     [{"runs": [{"text": plan.subtitle.text.replace("\n", " ")}], "size": B.scale["subtitle"],
                       "font": B.body_font, "color": B.neutral, "line_spacing": 1.2, "space_after": 0}])
            y += sh + 0.18
    if intro:
        y = _callout_bar(slide, ml, y, cw, None, B, para=intro)

    avail = footer_top - y

    if mode == "columns":
        wide, columns = cols_data
        # full-width headers become a compact intro block above the columns
        wparas = [p for s in wide for p in s.paras if p.text.strip()]
        if wparas:
            wh = min(_block_h_in(" ".join(p.text for p in wparas), B.scale["body"], cw,
                                 cw_factor=0.53) + 0.1, 0.95)
            wspecs = [s for p in wparas if (s := _para_spec_one(p, B, B.ink))]
            add_text(slide, ml, y, cw, wh, wspecs, anchor=MSO_ANCHOR.TOP, autofit=True)
            y += wh + 0.16
            avail = footer_top - y
        _render_columns(slide, deck, B, ml, y, cw, avail, columns)
    elif mode == "table":
        tb = plan.tables[0]
        nrows = len(tb.table) if tb.table else 0
        if items:
            specs = [s for p in items if (s := _para_spec_one(p, B, B.ink, size=B.scale["small"]))]
            frac = 0.22 if nrows >= 7 else 0.32
            bh = min((footer_top - y) * frac, 1.4)
            add_text(slide, ml, y, cw, max(0.35, bh), specs, anchor=MSO_ANCHOR.TOP, autofit=True)
            y += bh + 0.12
        tmax = footer_top - y - 0.05
        if tmax >= 0.7:
            _table(slide, tb, ml, y, cw, B, max_h=tmax)
    elif mode == "flow":
        _render_flow(slide, deck, B, ml, y + 0.2, cw, min(avail, 1.7), items)
    elif mode == "iconlist":
        rrect(slide, ml, y, cw, avail, fill=(B.dark if on_dark else B.light), radius=0.06)
        pad = 0.34
        iy, ih = y + pad, avail - 2 * pad
        if has_label:
            _section_label(slide, B, ml + pad, iy, cw - 2 * pad, items[0], on_dark)
            iy += 0.44; ih -= 0.44
            items = items[1:]
        _render_icon_list(slide, deck, B, ml + pad, iy, cw - 2 * pad, ih, items, icons, on_dark)
    elif mode == "grid":
        _render_card_grid(slide, deck, B, ml, y, cw, min(avail, 4.3), items)

    _footer(slide, deck, plan, B, on_dark=False)


def _callout_bar(slide, x, y, w, text, B, *, para=None, bold=False):
    """A light rounded callout bar; returns the new y. Preserves the run text if `para` given."""
    src = para.text if para else text
    chh = min(_block_h_in(src, B.scale["body"], w - 0.5, cw_factor=0.53) + 0.16, 1.2)
    rrect(slide, x, y, w, chh, fill=B.light, radius=0.10)
    if para is not None:
        spec = _para_spec_one(para, B, B.ink, space_after=0)
    else:
        spec = {"runs": [{"text": text, "bold": bold}], "size": B.scale["body"],
                "font": B.body_font, "color": B.dark, "line_spacing": 1.16, "space_after": 0}
    add_text(slide, x + 0.26, y + 0.07, w - 0.52, chh - 0.14, [spec], anchor=MSO_ANCHOR.MIDDLE, autofit=True)
    return y + chh + 0.16


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
    L = {l.name: l for l in prs.slide_masters[0].slide_layouts}
    texture = make_contour_texture(1280, 720, brand.dark, C.mix(brand.dark, "FFFFFF", 0.32))
    layout_for = {"title": "Title Slide", "section": "Section Header",
                  "closing": "Section Header", "content": "Title and Content"}
    for s in deck.slides:
        plan = classify_slide(s, deck)
        icons = [b for p in plan.pictures if p.image_blob and (b := recolor_icon(p.image_blob, brand.accent))]
        layout = L.get(layout_for.get(plan.stype, "Title and Content"), blank)
        slide = prs.slides.add_slide(layout)
        _del_ph_except(slide, {0, 10, 11, 12})  # keep title + footer chrome; drop body/subtitle prompts
        BUILDERS.get(plan.stype, build_content)(slide, deck, plan, brand, icons, texture)
    return prs
