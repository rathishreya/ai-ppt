"""Intermediate representation of a deck.

The IR decouples the rest of the pipeline from python-pptx / OOXML quirks. It captures
exactly what we need: ordered text runs (for preservation + brand + classification),
per-run formatting cues, and shape geometry. Non-text payloads (image bytes, table grids)
are carried through so composition can re-place them.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Optional

EMU_PER_IN = 914400
EMU_PER_PT = 12700


@dataclass
class RunIR:
    text: str
    font: Optional[str] = None       # explicit typeface or None (inherited)
    size_pt: Optional[float] = None  # explicit size or None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color_hex: Optional[str] = None  # explicit solid RGB or None


@dataclass
class ParaIR:
    runs: list[RunIR] = field(default_factory=list)
    level: int = 0
    bullet: bool = False
    align: Optional[str] = None

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)


@dataclass
class ShapeIR:
    kind: str                        # 'text' | 'table' | 'picture' | 'group' | 'other'
    paras: list[ParaIR] = field(default_factory=list)
    # geometry in EMU (may be None if inherited — rare for our source's free shapes)
    left: Optional[int] = None
    top: Optional[int] = None
    width: Optional[int] = None
    height: Optional[int] = None
    # payloads
    fill_hex: Optional[str] = None
    image_blob: Optional[bytes] = None
    image_ext: Optional[str] = None
    table: Optional[list[list[list[RunIR]]]] = None  # rows -> cells -> runs
    rot: float = 0.0

    @property
    def text(self) -> str:
        if self.table:
            return "\n".join(
                "\t".join("".join(r.text for r in cell) for cell in row) for row in self.table
            )
        return "\n".join(p.text for p in self.paras)

    @property
    def max_size_pt(self) -> float:
        sizes = [r.size_pt for p in self.paras for r in p.runs if r.size_pt]
        return max(sizes) if sizes else 0.0

    @property
    def plain_len(self) -> int:
        return len(self.text.strip())


@dataclass
class SlideIR:
    index: int
    shapes: list[ShapeIR] = field(default_factory=list)
    notes: str = ""
    slide_type: str = "content"      # set by classifier

    def text_shapes(self) -> list[ShapeIR]:
        return [s for s in self.shapes if s.plain_len > 0 and s.kind in ("text", "table")]


@dataclass
class DeckIR:
    slides: list[SlideIR] = field(default_factory=list)
    width_emu: int = 12192000
    height_emu: int = 6858000

    @property
    def width_in(self) -> float:
        return self.width_emu / EMU_PER_IN

    @property
    def height_in(self) -> float:
        return self.height_emu / EMU_PER_IN
