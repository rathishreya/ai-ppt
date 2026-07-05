"""Color utilities: hex/RGB/HSL conversion and brand-role classification.

Pure functions, no dependencies beyond stdlib. Used by brand extraction to turn the
raw palette found in a deck into named roles (primary/dark, accent, secondary, neutrals).
"""
from __future__ import annotations
import colorsys
from dataclasses import dataclass


def hex_to_rgb(h: str) -> tuple[int, int, int]:
    h = h.lstrip("#")
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def rgb_to_hex(rgb: tuple[int, int, int]) -> str:
    r, g, b = (max(0, min(255, int(round(v)))) for v in rgb)
    return f"{r:02X}{g:02X}{b:02X}"


def rgb_to_hls(rgb: tuple[int, int, int]) -> tuple[float, float, float]:
    r, g, b = (v / 255.0 for v in rgb)
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    return h * 360.0, l, s  # hue in degrees, lightness/sat in [0,1]


def hls_to_hex(h_deg: float, l: float, s: float) -> str:
    r, g, b = colorsys.hls_to_rgb((h_deg % 360) / 360.0, max(0, min(1, l)), max(0, min(1, s)))
    return rgb_to_hex((r * 255, g * 255, b * 255))


def complement_accent(dark_hex: str, l: float = 0.57, s: float = 0.68) -> str:
    """A warm complementary accent derived from the primary hue (e.g. navy -> orange)."""
    h, _, _ = rgb_to_hls(hex_to_rgb(dark_hex))
    return hls_to_hex(h + 180, l, s)


def relative_luminance(rgb: tuple[int, int, int]) -> float:
    """WCAG relative luminance."""
    def chan(c: float) -> float:
        c /= 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = rgb
    return 0.2126 * chan(r) + 0.7152 * chan(g) + 0.0722 * chan(b)


def contrast_ratio(a: tuple[int, int, int], b: tuple[int, int, int]) -> float:
    la, lb = relative_luminance(a), relative_luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def best_text_on(bg_hex: str, light="FFFFFF", dark="14213A") -> str:
    """Return whichever of light/dark text has better contrast on bg (WCAG-aware)."""
    bg = hex_to_rgb(bg_hex)
    return light if contrast_ratio(hex_to_rgb(light), bg) >= contrast_ratio(hex_to_rgb(dark), bg) else dark


def mix(a_hex: str, b_hex: str, t: float) -> str:
    """Linear blend a->b by t in [0,1]."""
    a, b = hex_to_rgb(a_hex), hex_to_rgb(b_hex)
    return rgb_to_hex(tuple(a[i] + (b[i] - a[i]) * t for i in range(3)))


@dataclass
class Swatch:
    hex: str
    weight: float  # accumulated area/frequency weight

    @property
    def rgb(self):
        return hex_to_rgb(self.hex)

    @property
    def hls(self):
        return rgb_to_hls(self.rgb)


def is_neutral(hex_: str, sat_thresh=0.12) -> bool:
    _, _, s = rgb_to_hls(hex_to_rgb(hex_))
    return s < sat_thresh


def role_scores(sw: Swatch) -> dict[str, float]:
    """Heuristic affinity of a swatch to each brand role (higher = better fit)."""
    h, l, s = sw.hls
    scores = {}
    # dark primary: dark & (bluish or neutral-dark)
    scores["dark"] = (1 - l) * (0.6 + 0.4 * (1 if 180 <= h <= 260 else 0.3))
    # accent: vivid warm (orange/red) OR simply the most saturated non-blue
    warm = 1.0 if (0 <= h <= 45 or h >= 330) else 0.2
    scores["accent"] = s * (0.4 + 0.6 * warm) * (0.5 + 0.5 * (1 - abs(l - 0.55)))
    # secondary: saturated green/teal/cyan
    teal = 1.0 if 140 <= h <= 200 else 0.15
    scores["secondary"] = s * teal * (0.5 + 0.5 * (1 - abs(l - 0.55)))
    # light bg
    scores["light"] = l * (1 - s)
    return scores
