"""Deterministic on-brand asset generation (Pillow) — no network, no AI required.

- contour texture: soft flowing navy background for title/section/closing slides
- icon recolor: tint the deck's own icons to an accent color for icon-rings
Everything is seeded so re-runs are byte-stable.
"""
from __future__ import annotations
import io
import math
import random

from PIL import Image, ImageDraw, ImageFilter


def _rgb(hex_):
    h = hex_.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def make_contour_texture(w: int, h: int, base_hex: str, tint_hex: str, seed: int = 7) -> bytes:
    """Layered soft blobs + faint contour arcs → a calm 'flowing' navy field."""
    rnd = random.Random(seed)
    base = _rgb(base_hex)
    tint = _rgb(tint_hex)
    img = Image.new("RGB", (w, h), base)

    # soft lighter blobs (like the reference's waves), heavily blurred and low-opacity
    for _ in range(9):
        layer = Image.new("L", (w, h), 0)
        d = ImageDraw.Draw(layer)
        cx, cy = rnd.uniform(0.1, 1.05) * w, rnd.uniform(-0.2, 0.95) * h
        rx, ry = rnd.uniform(0.35, 0.85) * w, rnd.uniform(0.35, 0.7) * h
        d.ellipse([cx - rx, cy - ry, cx + rx, cy + ry], fill=255)
        layer = layer.filter(ImageFilter.GaussianBlur(radius=w // 10))
        alpha = rnd.uniform(0.10, 0.26)
        tinted = Image.new("RGB", (w, h), tint)
        img = Image.composite(tinted, img, layer.point(lambda p: int(p * alpha)))

    # concentric contour arcs (top-right) for the 'flowing wave' feel
    arc = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    ad = ImageDraw.Draw(arc)
    ox, oy = rnd.uniform(0.6, 1.0) * w, rnd.uniform(0.0, 0.35) * h
    for i in range(1, 26):
        r = i * (w // 24)
        ad.ellipse([ox - r, oy - r, ox + r, oy + r], outline=(tint[0], tint[1], tint[2], 40), width=3)
    arc = arc.filter(ImageFilter.GaussianBlur(radius=2))
    img = Image.alpha_composite(img.convert("RGBA"), arc).convert("RGB")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def recolor_icon(blob: bytes, hex_: str) -> bytes | None:
    """Tint an icon (any shape) to a solid accent color, preserving its silhouette."""
    try:
        im = Image.open(io.BytesIO(blob)).convert("RGBA")
    except Exception:
        return None
    r, g, b, a = im.split()
    # if the icon is opaque, derive a mask from luminance (assume light glyph)
    if a.getextrema()[0] == 255 and a.getextrema()[1] == 255:
        gray = im.convert("L")
        a = gray  # brighter -> more opaque
    col = _rgb(hex_)
    solid = Image.new("RGBA", im.size, (col[0], col[1], col[2], 0))
    solid.putalpha(a)
    buf = io.BytesIO()
    solid.save(buf, format="PNG")
    return buf.getvalue()
