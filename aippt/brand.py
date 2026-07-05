"""Derive a BrandSpec (the 'Design DNA') from a DeckIR — deterministically.

Palette: area/frequency-weighted clustering of every fill & font color, mapped to roles
(dark/primary, accent, secondary, neutrals, light bg) by HSL heuristics with brand-safe
fallbacks. Typography: detect the heading (serif) and body (sans) families and collapse the
deck's many ad-hoc sizes into a disciplined modular scale.
"""
from __future__ import annotations
from collections import Counter
from dataclasses import dataclass, field, asdict
import json

from .ir import DeckIR
from . import color as C

SERIF_HINTS = ("cambria", "times", "georgia", "garamond", "serif", "playfair", "merriweather",
               "newsreader", "pt serif", "source serif", "constantia", "palatino", "book antiqua",
               "minion", "caslon", "didot", "noto serif")


def _is_serif(name: str) -> bool:
    n = (name or "").lower()
    return any(h in n for h in SERIF_HINTS)


@dataclass
class BrandSpec:
    # colors (hex, no '#')
    dark: str = "22314F"          # primary dark / navy
    accent: str = "E8623D"        # accent / orange
    secondary: str = "5CB8A5"     # secondary / teal
    light: str = "F1F2F4"         # light card / bg tint
    paper: str = "FFFFFF"         # page background
    ink: str = "1E2A44"           # default text on light
    neutral: str = "6B7280"       # muted text
    # fonts
    heading_font: str = "Georgia"
    body_font: str = "Calibri"
    # type scale (pt)
    scale: dict = field(default_factory=lambda: {
        "display": 40, "h1": 30, "h2": 20, "eyebrow": 12.5,
        "subtitle": 15, "body": 13, "small": 10.5, "footer": 9,
    })
    # detected raw palette (for audit)
    palette_raw: list = field(default_factory=list)

    def text_on(self, bg_hex: str) -> str:
        return C.best_text_on(bg_hex, light=self.paper, dark=self.ink)

    def to_json(self) -> str:
        return json.dumps(asdict(self), indent=2)


def _collect_swatches(deck: DeckIR) -> list[C.Swatch]:
    weight: Counter = Counter()
    for s in deck.slides:
        for sh in s.shapes:
            # fills weighted by shape area (in²), fonts by run count
            if sh.fill_hex:
                area = 1.0
                if sh.width and sh.height:
                    area = max(0.2, (sh.width * sh.height) / (914400.0 ** 2))
                weight[sh.fill_hex.upper()] += area
            for p in sh.paras:
                for r in p.runs:
                    if r.color_hex:
                        weight[r.color_hex.upper()] += 0.25
    return [C.Swatch(hex=h, weight=w) for h, w in weight.most_common()]


def _is_warm_vivid(hex_: str) -> bool:
    h, l, s = C.rgb_to_hls(C.hex_to_rgb(hex_))
    return s >= 0.35 and (0 <= h <= 45 or h >= 330) and 0.30 <= l <= 0.75


def _pick_dark(swatches: list[C.Swatch]) -> str | None:
    """Prefer the dominant true navy: bluish hue, mid-dark lightness, high weight."""
    best, best_score = None, 0.0
    for sw in swatches:
        h, l, s = sw.hls
        if l > 0.42:
            continue
        bluish = 1.0 if 180 <= h <= 260 else (0.4 if s < 0.15 else 0.15)
        # favor lightness ~0.15-0.28 (readable navy, not near-black), and weight
        light_fit = 1.0 - min(1.0, abs(l - 0.22) / 0.22)
        score = bluish * (0.5 + 0.5 * light_fit) * (0.5 + 0.5 * min(1.0, sw.weight / 10.0))
        if score > best_score:
            best, best_score = sw.hex, score
    return best


def _pick_role(swatches: list[C.Swatch], role: str, exclude: set[str]) -> str | None:
    best, best_score = None, 0.0
    for sw in swatches:
        if sw.hex in exclude:
            continue
        sc = C.role_scores(sw)[role] * (0.6 + 0.4 * min(1.0, sw.weight))
        if sc > best_score:
            best, best_score = sw.hex, sc
    return best


def extract_brand(deck: DeckIR) -> BrandSpec:
    spec = BrandSpec()
    swatches = _collect_swatches(deck)
    spec.palette_raw = [{"hex": s.hex, "weight": round(s.weight, 2)} for s in swatches[:20]]

    used: set[str] = set()
    # dark/primary: prefer a real navy (bluish, mid-dark, high weight) over near-black
    dark = _pick_dark(swatches) or _pick_role(swatches, "dark", used)
    if dark:
        spec.dark = dark; used.add(dark)
    # accent: the deck's own warm accent if it has one; else synthesize the complement
    accent = _pick_role([s for s in swatches if not C.is_neutral(s.hex)], "accent", used)
    if accent and _is_warm_vivid(accent):
        spec.accent = accent; used.add(accent)
    else:
        spec.accent = C.complement_accent(spec.dark)  # navy -> orange, generalizable
    secondary = _pick_role([s for s in swatches if not C.is_neutral(s.hex) and s.hex != spec.accent],
                           "secondary", used)
    if secondary:
        spec.secondary = secondary; used.add(secondary)

    # ink = darkest bluish; neutral = a mid gray; light = near-white tint
    spec.ink = spec.dark
    spec.neutral = C.mix(spec.dark, "FFFFFF", 0.45)
    spec.light = C.mix(spec.dark, "FFFFFF", 0.94)

    # fonts: most common serif = heading; most common sans = body
    font_counts: Counter = Counter()
    big_font_counts: Counter = Counter()
    for s in deck.slides:
        for sh in s.shapes:
            for p in sh.paras:
                for r in p.runs:
                    if r.font:
                        font_counts[r.font] += 1
                        if (r.size_pt or 0) >= 18:
                            big_font_counts[r.font] += 1
    serifs = [f for f, _ in font_counts.most_common() if _is_serif(f)]
    sans = [f for f, _ in font_counts.most_common() if not _is_serif(f)]
    # prefer a serif that appears in big text (headings)
    big_serifs = [f for f, _ in big_font_counts.most_common() if _is_serif(f)]
    if big_serifs:
        spec.heading_font = big_serifs[0]
    elif serifs:
        spec.heading_font = serifs[0]
    if sans:
        spec.body_font = sans[0]

    # type scale: anchor on observed max heading size, keep a clean modular ramp
    sizes = [r.size_pt for s in deck.slides for sh in s.shapes for p in sh.paras
             for r in p.runs if r.size_pt]
    if sizes:
        display = min(46, max(34, max(sizes)))
        spec.scale = {
            "display": round(display), "h1": round(display * 0.72),
            "h2": round(display * 0.5), "eyebrow": 12.5, "subtitle": round(display * 0.38),
            "body": 13, "small": 10.5, "footer": 9,
        }
    return spec
