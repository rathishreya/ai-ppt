"""Text-preservation & slide-parity guard (invariants I1/I2).

The gate is fail-closed: the pipeline refuses to emit if the output's text inventory is not
equal to the source's. "Equal" is a *fair* comparison — whitespace-normalized, order-independent
within a slide (a word multiset) — so restyling/reflow is allowed but no word may be added or
dropped. We also track exact non-space character counts as a second signal.
"""
from __future__ import annotations
import re
from collections import Counter
from dataclasses import dataclass

from .ir import DeckIR, SlideIR

_WS = re.compile(r"\s+")


def _strip_glyphs(s: str) -> str:
    """Decorative glyphs (Wingdings/private-use icons, object-replacement, control chars) are
    NOT text — the compose step drops them, so the guard ignores them on BOTH sides to stay fair."""
    out = []
    for ch in s:
        o = ord(ch)
        if o < 0x20 and ch not in "\t\n\r":
            continue
        if 0xE000 <= o <= 0xF8FF or o in (0xFFFC, 0xFFFD):
            continue
        out.append(ch)
    return "".join(out)


def _norm(s: str) -> str:
    return _WS.sub(" ", _strip_glyphs(s)).strip()


def slide_words(slide: SlideIR) -> Counter:
    """Multiset of whitespace-delimited tokens across all text on the slide."""
    words: Counter = Counter()
    for sh in slide.shapes:
        for tok in _norm(sh.text).split(" "):
            if tok:
                words[tok] += 1
    return words


def slide_charcount(slide: SlideIR) -> int:
    return sum(len(_WS.sub("", _strip_glyphs(sh.text))) for sh in slide.shapes)


@dataclass
class GuardReport:
    ok: bool
    slide_count_src: int
    slide_count_out: int
    per_slide: list[dict]

    def summary(self) -> str:
        lines = [f"slide parity: src={self.slide_count_src} out={self.slide_count_out} "
                 f"-> {'OK' if self.slide_count_src == self.slide_count_out else 'FAIL'}"]
        bad = [p for p in self.per_slide if not p["ok"]]
        lines.append(f"text preservation: {len(self.per_slide) - len(bad)}/{len(self.per_slide)} slides OK")
        for p in bad:
            lines.append(f"  slide {p['index']+1}: MISSING={p['missing']} EXTRA={p['extra']} "
                         f"(chars src={p['chars_src']} out={p['chars_out']})")
        lines.append(f"OVERALL: {'PASS' if self.ok else 'FAIL'}")
        return "\n".join(lines)


def verify(src: DeckIR, out: DeckIR) -> GuardReport:
    per = []
    ok_all = (len(src.slides) == len(out.slides))
    for i in range(min(len(src.slides), len(out.slides))):
        ws, wo = slide_words(src.slides[i]), slide_words(out.slides[i])
        missing = ws - wo   # in source, not in output (dropped/altered)
        extra = wo - ws     # in output, not in source (added/altered)
        ok = (not missing and not extra)
        per.append(dict(
            index=i, ok=ok,
            missing=dict(list(missing.items())[:12]),
            extra=dict(list(extra.items())[:12]),
            chars_src=slide_charcount(src.slides[i]),
            chars_out=slide_charcount(out.slides[i]),
        ))
        ok_all = ok_all and ok
    return GuardReport(ok=ok_all, slide_count_src=len(src.slides),
                       slide_count_out=len(out.slides), per_slide=per)
