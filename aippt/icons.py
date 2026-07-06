"""Bundled vector icon set — drawn programmatically (Pillow), license-free & on-brand.

No network/asset dependency. Each icon is a simple, recognizable line glyph rendered to a
transparent PNG in the accent color, then keyword-mapped to list items so every icon-ring
carries a meaningful symbol.
"""
from __future__ import annotations
import io
import math
from PIL import Image, ImageDraw

_cache: dict[tuple, bytes] = {}


def _rgb(hex_):
    h = hex_.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def render_icon(name: str, hex_: str, sz: int = 240) -> bytes:
    key = (name, hex_, sz)
    if key in _cache:
        return _cache[key]
    img = Image.new("RGBA", (sz, sz), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    col = _rgb(hex_) + (255,)
    lw = max(4, sz // 22)
    S = sz

    def L(x1, y1, x2, y2, w=lw):
        d.line([x1 * S, y1 * S, x2 * S, y2 * S], fill=col, width=w)

    def C(cx, cy, r, fill=False, w=lw):
        bb = [(cx - r) * S, (cy - r) * S, (cx + r) * S, (cy + r) * S]
        d.ellipse(bb, outline=col, width=w, fill=(col if fill else None))

    def R(x1, y1, x2, y2, rad=0.05, fill=False, w=lw):
        d.rounded_rectangle([x1 * S, y1 * S, x2 * S, y2 * S], radius=rad * S,
                            outline=col, width=w, fill=(col if fill else None))

    def P(pts, fill=False, w=lw, closed=True):
        xy = [(x * S, y * S) for x, y in pts]
        if fill:
            d.polygon(xy, fill=col)
        else:
            d.line(xy + ([xy[0]] if closed else []), fill=col, width=w, joint="curve")

    def rect(x1, y1, x2, y2, fill=True):
        d.rectangle([x1 * S, y1 * S, x2 * S, y2 * S], fill=(col if fill else None),
                    outline=col, width=lw)

    n = name
    if n == "search":
        C(0.42, 0.42, 0.24); L(0.60, 0.60, 0.82, 0.82)
    elif n == "chart":
        L(0.22, 0.18, 0.22, 0.80); L(0.22, 0.80, 0.84, 0.80)
        rect(0.32, 0.58, 0.44, 0.80); rect(0.50, 0.44, 0.62, 0.80); rect(0.68, 0.30, 0.80, 0.80)
    elif n == "document":
        R(0.30, 0.16, 0.70, 0.84, 0.05)
        for yy in (0.34, 0.44, 0.54, 0.64):
            L(0.38, yy, 0.62, yy, max(4, lw - 3))
    elif n == "clipboard":
        R(0.30, 0.20, 0.70, 0.84, 0.05); R(0.42, 0.13, 0.58, 0.24, 0.03)
        for yy in (0.40, 0.50, 0.60):
            L(0.38, yy, 0.62, yy, max(4, lw - 3))
    elif n == "shield":
        P([(0.5, 0.14), (0.80, 0.26), (0.80, 0.52), (0.5, 0.86), (0.20, 0.52), (0.20, 0.26)])
        L(0.38, 0.48, 0.47, 0.58); L(0.47, 0.58, 0.64, 0.36)
    elif n == "people":
        C(0.36, 0.36, 0.12); C(0.64, 0.36, 0.12)
        P([(0.20, 0.74), (0.26, 0.56), (0.46, 0.56), (0.52, 0.74)], closed=False)
        P([(0.48, 0.74), (0.54, 0.56), (0.74, 0.56), (0.80, 0.74)], closed=False)
    elif n == "gear":
        C(0.5, 0.5, 0.20); C(0.5, 0.5, 0.09)
        for k in range(8):
            a = k * math.pi / 4
            L(0.5 + 0.20 * math.cos(a), 0.5 + 0.20 * math.sin(a),
              0.5 + 0.32 * math.cos(a), 0.5 + 0.32 * math.sin(a))
    elif n == "target":
        C(0.5, 0.5, 0.30); C(0.5, 0.5, 0.18); C(0.5, 0.5, 0.05, fill=True)
    elif n == "check":
        C(0.5, 0.5, 0.32); L(0.36, 0.52, 0.46, 0.63); L(0.46, 0.63, 0.66, 0.40)
    elif n == "chat":
        R(0.18, 0.22, 0.82, 0.62, 0.10); P([(0.34, 0.62), (0.34, 0.78), (0.50, 0.62)], fill=True)
        for xx in (0.34, 0.50, 0.66):
            C(xx, 0.42, 0.03, fill=True)
    elif n == "megaphone":
        P([(0.22, 0.40), (0.58, 0.28), (0.58, 0.64), (0.22, 0.52)])
        P([(0.58, 0.31), (0.78, 0.24), (0.78, 0.68), (0.58, 0.61)], closed=True)
        L(0.30, 0.54, 0.34, 0.72)
    elif n == "trophy":
        P([(0.34, 0.20), (0.66, 0.20), (0.62, 0.50), (0.38, 0.50)])
        d.arc([0.60 * S, 0.22 * S, 0.82 * S, 0.44 * S], -90, 90, fill=col, width=lw)
        d.arc([0.18 * S, 0.22 * S, 0.40 * S, 0.44 * S], 90, 270, fill=col, width=lw)
        L(0.5, 0.50, 0.5, 0.64); L(0.36, 0.78, 0.64, 0.78); L(0.42, 0.64, 0.58, 0.64)
    elif n == "calendar":
        R(0.20, 0.22, 0.80, 0.82, 0.05); L(0.20, 0.36, 0.80, 0.36)
        L(0.34, 0.14, 0.34, 0.28); L(0.66, 0.14, 0.66, 0.28)
    elif n == "idea":
        C(0.5, 0.40, 0.20); L(0.42, 0.62, 0.58, 0.62); L(0.44, 0.70, 0.56, 0.70)
        L(0.46, 0.78, 0.54, 0.78)
    elif n == "book":
        L(0.5, 0.24, 0.5, 0.80)
        P([(0.5, 0.24), (0.24, 0.30), (0.24, 0.78), (0.5, 0.72)], closed=True)
        P([(0.5, 0.24), (0.76, 0.30), (0.76, 0.78), (0.5, 0.72)], closed=True)
    elif n == "flag":
        L(0.30, 0.16, 0.30, 0.84); P([(0.30, 0.20), (0.72, 0.28), (0.30, 0.44)], fill=True)
    elif n == "building":
        P([(0.5, 0.16), (0.20, 0.34), (0.80, 0.34)], fill=True)
        for xx in (0.28, 0.42, 0.58, 0.72):
            L(xx, 0.38, xx, 0.74)
        L(0.18, 0.78, 0.82, 0.78)
    elif n == "handshake":
        P([(0.16, 0.46), (0.40, 0.40), (0.60, 0.52), (0.50, 0.62)], closed=False)
        P([(0.84, 0.46), (0.60, 0.40), (0.44, 0.52)], closed=False)
        C(0.5, 0.54, 0.02, fill=True)
    elif n == "star":
        pts = []
        for k in range(10):
            a = -math.pi / 2 + k * math.pi / 5
            r = 0.32 if k % 2 == 0 else 0.14
            pts.append((0.5 + r * math.cos(a), 0.5 + r * math.sin(a)))
        P(pts, fill=True)
    elif n == "grid":
        R(0.22, 0.22, 0.46, 0.46, 0.04); R(0.54, 0.22, 0.78, 0.46, 0.04)
        R(0.22, 0.54, 0.46, 0.78, 0.04); R(0.54, 0.54, 0.78, 0.78, 0.04)
    elif n == "arrow":
        L(0.20, 0.5, 0.74, 0.5); L(0.74, 0.5, 0.60, 0.38); L(0.74, 0.5, 0.60, 0.62)
    elif n == "coin":
        C(0.5, 0.5, 0.30); L(0.5, 0.34, 0.5, 0.66); C(0.5, 0.5, 0.14)
    elif n == "mail":
        R(0.18, 0.30, 0.82, 0.70, 0.03)
        P([(0.18, 0.33), (0.5, 0.54), (0.82, 0.33)], closed=False)
    elif n == "globe":
        C(0.5, 0.5, 0.30); L(0.20, 0.5, 0.80, 0.5)
        d.ellipse([0.36 * S, 0.20 * S, 0.64 * S, 0.80 * S], outline=col, width=lw)
    else:  # generic mark
        C(0.5, 0.5, 0.18, fill=True)

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    _cache[key] = buf.getvalue()
    return buf.getvalue()


# keyword -> icon (checked in order; first hit wins)
_KEYWORDS = [
    ("assess", "search"), ("diagnos", "search"), ("audit", "search"), ("insight", "search"),
    ("baseline", "chart"), ("measur", "chart"), ("metric", "chart"), ("adoption", "chart"),
    ("productiv", "chart"), ("result", "chart"), ("data", "chart"), ("evidence", "chart"),
    ("survey", "clipboard"), ("assessment report", "document"), ("report", "document"),
    ("document", "document"), ("draft", "document"), ("sow", "document"),
    ("governance", "shield"), ("privacy", "shield"), ("risk", "shield"), ("responsible", "shield"),
    ("secur", "shield"), ("complian", "shield"),
    ("team", "people"), ("facilitat", "people"), ("stakeholder", "people"), ("workforce", "people"),
    ("role", "people"), ("bilingual", "chat"), ("communic", "chat"), ("message", "chat"),
    ("campaign", "megaphone"), ("awareness", "megaphone"), ("announce", "megaphone"),
    ("champion", "trophy"), ("recogn", "trophy"), ("award", "trophy"), ("incentive", "trophy"),
    ("badge", "trophy"), ("leaderboard", "trophy"),
    ("method", "gear"), ("process", "gear"), ("tool", "gear"), ("platform", "grid"),
    ("dashboard", "grid"), ("system", "grid"),
    ("framework", "book"), ("learn", "book"), ("training", "book"), ("session", "book"),
    ("kpi", "target"), ("goal", "target"), ("target", "target"), ("outcome", "target"),
    ("priorit", "target"), ("account", "check"), ("verif", "check"), ("proven", "check"),
    ("valid", "check"), ("quality", "star"), ("proof", "check"),
    ("timeline", "calendar"), ("week", "calendar"), ("schedule", "calendar"), ("calendar", "calendar"),
    ("milestone", "flag"), ("deliver", "flag"), ("handover", "flag"),
    ("institution", "building"), ("government", "building"), ("sovereign", "building"),
    ("national", "building"), ("public-sector", "building"),
    ("partner", "handshake"), ("pedigree", "handshake"), ("cross-sector", "handshake"),
    ("idea", "idea"), ("why", "idea"), ("reason", "idea"), ("purpose", "idea"),
    ("cost", "coin"), ("roi", "coin"), ("budget", "coin"), ("value", "coin"), ("time", "coin"),
    ("flow", "arrow"), ("sequenc", "arrow"), ("link", "arrow"), ("practice", "gear"),
]

_DEFAULT_CYCLE = ["target", "check", "chart", "gear", "people", "book"]


def pick_icon_name(text: str, index: int = 0) -> str:
    t = text.lower()
    for kw, icon in _KEYWORDS:
        if kw in t:
            return icon
    return _DEFAULT_CYCLE[index % len(_DEFAULT_CYCLE)]


def icon_for(text: str, hex_: str, index: int = 0) -> bytes:
    return render_icon(pick_icon_name(text, index), hex_)
