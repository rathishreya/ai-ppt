"""Load a .pptx into the DeckIR (python-pptx + lxml; never LibreOffice)."""
from __future__ import annotations
from pptx import Presentation
from pptx.util import Emu
from pptx.enum.shapes import MSO_SHAPE_TYPE

from .ir import DeckIR, SlideIR, ShapeIR, ParaIR, RunIR


def _run_color_hex(run) -> str | None:
    try:
        c = run.font.color
        if c is not None and c.type is not None and str(c.type) == "MSO_THEME_COLOR.NOT_THEME_COLOR":
            return str(c.rgb)
    except Exception:
        pass
    # also handle explicit rgb without theme type
    try:
        c = run.font.color
        if c is not None and c.rgb is not None:
            return str(c.rgb)
    except Exception:
        pass
    return None


def _para_to_ir(p) -> ParaIR:
    runs = []
    for r in p.runs:
        runs.append(RunIR(
            text=r.text,
            font=r.font.name,
            size_pt=(r.font.size.pt if r.font.size is not None else None),
            bold=r.font.bold,
            italic=r.font.italic,
            color_hex=_run_color_hex(r),
        ))
    # paragraph-level fallback: if no runs but text present (rare), keep the text
    if not runs and (p.text or "").strip():
        runs.append(RunIR(text=p.text))
    align = None
    try:
        align = str(p.alignment) if p.alignment is not None else None
    except Exception:
        pass
    return ParaIR(runs=runs, level=p.level or 0, bullet=False, align=align)


def _fill_hex(shape) -> str | None:
    try:
        f = shape.fill
        if f.type is not None and f.fore_color is not None and f.fore_color.type is not None:
            if str(f.fore_color.type) == "MSO_THEME_COLOR.NOT_THEME_COLOR":
                return str(f.fore_color.rgb)
    except Exception:
        pass
    return None


def _shape_to_ir(shape) -> ShapeIR:
    st = shape.shape_type
    geom = dict(
        left=int(shape.left) if shape.left is not None else None,
        top=int(shape.top) if shape.top is not None else None,
        width=int(shape.width) if shape.width is not None else None,
        height=int(shape.height) if shape.height is not None else None,
    )
    try:
        rot = float(shape.rotation)
    except Exception:
        rot = 0.0

    if st == MSO_SHAPE_TYPE.GROUP:
        sub = [_shape_to_ir(s) for s in shape.shapes]
        # flatten group children as siblings (keeps text order) but tag kind
        g = ShapeIR(kind="group", **geom, rot=rot)
        g.paras = [pp for c in sub for pp in c.paras]
        # attach children table/image via first child that has them (best-effort)
        return g

    if st == MSO_SHAPE_TYPE.TABLE or shape.has_table:
        tbl = shape.table
        grid = []
        for row in tbl.rows:
            cells = []
            for cell in row.cells:
                cruns = []
                for p in cell.text_frame.paragraphs:
                    for r in p.runs:
                        cruns.append(RunIR(text=r.text, font=r.font.name,
                                           size_pt=(r.font.size.pt if r.font.size else None),
                                           bold=r.font.bold, italic=r.font.italic,
                                           color_hex=_run_color_hex(r)))
                    if not cruns and (p.text or "").strip():
                        cruns.append(RunIR(text=p.text))
                cells.append(cruns)
            grid.append(cells)
        return ShapeIR(kind="table", table=grid, **geom, rot=rot)

    if st == MSO_SHAPE_TYPE.PICTURE:
        blob = ext = None
        try:
            img = shape.image
            blob, ext = img.blob, img.ext
        except Exception:
            pass
        return ShapeIR(kind="picture", image_blob=blob, image_ext=ext, **geom, rot=rot)

    if shape.has_text_frame:
        paras = [_para_to_ir(p) for p in shape.text_frame.paragraphs]
        return ShapeIR(kind="text", paras=paras, fill_hex=_fill_hex(shape), **geom, rot=rot)

    return ShapeIR(kind="other", fill_hex=_fill_hex(shape), **geom, rot=rot)


def load_deck(path: str) -> DeckIR:
    prs = Presentation(path)
    deck = DeckIR(width_emu=prs.slide_width, height_emu=prs.slide_height)
    for i, slide in enumerate(prs.slides):
        sir = SlideIR(index=i)
        for shape in slide.shapes:
            sir.shapes.append(_shape_to_ir(shape))
        try:
            if slide.has_notes_slide:
                sir.notes = slide.notes_slide.notes_text_frame.text or ""
        except Exception:
            pass
        deck.slides.append(sir)
    return deck
