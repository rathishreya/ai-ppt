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

    # footers: the recurring SHORT footer line, or a copyright/confidential line (any length).
    # Long slide-specific *content* notes stay in the body.
    _FH = ("confidential", "proposal", "خاص", "©", "copyright", "permission", "all rights")
    plan.footers = [s for s in texts if (s.top or 0) > 0.85 * H and s is not plan.wordmark
                    and (s.plain_len <= 55 or any(h in s.text.lower() for h in _FH))]

    pool = [s for s in texts if s is not plan.wordmark and s not in plan.footers]

    # The master puts a recurring HEADER (topic line) in a fixed top band. It is the
    # slide's title *per the master* even though it is rarely the biggest font on the slide.
    header = None
    band = [s for s in pool if (s.top or 0) < 0.10 * H]
    if band:
        header = min(band, key=lambda s: (s.top or 0))

    # A headline is a SINGLE dominant block below the header (e.g. "Minimum Tasks:").
    # If several shapes share the largest font, they are a list/group — NOT a headline,
    # so the header stays the title and those shapes fall to the body.
    lower = [s for s in pool if s is not header]
    headline = None
    if lower:
        mx = max(round(s.max_size_pt, 1) for s in lower)
        biggest = [s for s in lower if round(s.max_size_pt, 1) >= mx - 0.6]
        cand = min(biggest, key=lambda s: (s.top or 0))
        t = cand.text.strip()
        looks_listy = t[:1] in "-•*·(" or t[:2] in ("1.", "2.", "3.", "4.")
        if len(biggest) == 1 and not looks_listy and cand.plain_len <= 120 \
                and (header is None or mx >= (header.max_size_pt or 0)):
            headline = cand

    if headline is not None:
        plan.title, plan.eyebrow = headline, header
    elif header is not None:
        plan.title = header
    elif pool:
        plan.title = max(pool, key=lambda s: (round(s.max_size_pt, 1), -(s.top or 0)))

    # subtitle: a genuine lead SENTENCE right below the title — never a formula/list sibling.
    if plan.title:
        below = sorted([s for s in pool if s not in (plan.title, plan.eyebrow)
                        and (s.top or 0) > (plan.title.top or 0)], key=_pos)
        if below:
            b0 = below[0]
            alone = sum(1 for s in below if abs(s.max_size_pt - b0.max_size_pt) < 1) == 1
            formulaic = ("\t" in b0.text) or ("=" in b0.text)
            single_para = len([p for p in b0.paras if p.text.strip()]) <= 1
            if b0.plain_len >= 18 and b0.max_size_pt <= plan.title.max_size_pt \
                    and alone and single_para and not formulaic:
                plan.subtitle = b0

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
