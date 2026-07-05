"""Per-slide classification: slide type + role of each shape.

Heuristic and *defensive*: anything not confidently a title/eyebrow/subtitle/footer/wordmark
falls into `body`, so no text is ever stranded (protects invariant I2 downstream).
"""
from __future__ import annotations
from dataclasses import dataclass, field

from .ir import DeckIR, SlideIR, ShapeIR


@dataclass
class SlidePlan:
    index: int
    stype: str = "content"                     # title | section | closing | content
    wordmark: ShapeIR | None = None
    eyebrow: ShapeIR | None = None
    title: ShapeIR | None = None
    subtitle: ShapeIR | None = None
    footers: list[ShapeIR] = field(default_factory=list)
    body: list[ShapeIR] = field(default_factory=list)
    tables: list[ShapeIR] = field(default_factory=list)
    pictures: list[ShapeIR] = field(default_factory=list)


def _pos(s: ShapeIR):
    return (s.top or 0, s.left or 0)


def classify_slide(slide: SlideIR, deck: DeckIR) -> SlidePlan:
    W, H = deck.width_emu, deck.height_emu
    texts = [s for s in slide.shapes if s.kind == "text" and s.plain_len > 0]
    tables = [s for s in slide.shapes if s.kind == "table"]
    pics = [s for s in slide.shapes if s.kind == "picture"]
    for s in texts:
        # group shapes carry text too
        pass
    texts.sort(key=_pos)

    plan = SlidePlan(index=slide.index, tables=tables, pictures=pics)

    # wordmark: a short 'AI71'-like token near the top-left
    for s in texts:
        t = s.text.strip().lower().replace(" ", "")
        if t in ("ai71", "ai7l", "ai7i") or (len(s.text.strip()) <= 6 and (s.top or 0) < 0.16 * H
                                             and (s.left or 0) < 0.30 * W and s.max_size_pt <= 30):
            plan.wordmark = s
            break

    # footers: small text in the bottom band
    plan.footers = [s for s in texts if (s.top or 0) > 0.86 * H and s is not plan.wordmark]

    pool = [s for s in texts if s is not plan.wordmark and s not in plan.footers]

    # title: the largest text on the slide (tie-break: higher up)
    if pool:
        plan.title = max(pool, key=lambda s: (round(s.max_size_pt, 1), -(s.top or 0)))

    # eyebrow: short line at/above the title
    if plan.title:
        ty = plan.title.top or 0
        above = [s for s in pool if s is not plan.title and (s.top or 0) <= ty + 1 and s.plain_len <= 52]
        if above:
            plan.eyebrow = min(above, key=lambda s: abs((s.top or 0) - ty))

    # subtitle: the block right below the title, not tiny, smaller-or-equal to title
    if plan.title:
        below = sorted([s for s in pool if s not in (plan.title, plan.eyebrow)
                        and (s.top or 0) > (plan.title.top or 0)], key=_pos)
        if below and below[0].plain_len >= 18 and below[0].max_size_pt <= plan.title.max_size_pt:
            plan.subtitle = below[0]

    plan.body = [s for s in pool if s not in (plan.title, plan.eyebrow, plan.subtitle)]
    plan.body.sort(key=_pos)

    # slide type
    alltext = " ".join(s.text for s in texts).strip()
    if slide.index == 0:
        plan.stype = "title"
    elif len(alltext) <= 44 and plan.title and plan.title.max_size_pt >= 28 and not plan.body and not tables:
        plan.stype = "closing" if slide.index == len(deck.slides) - 1 else "section"
    elif plan.eyebrow and plan.eyebrow.text.strip().lower().startswith("phase") \
            and len(plan.body) == 0 and not tables:
        plan.stype = "section"
    return plan
