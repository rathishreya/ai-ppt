# ai-ppt

Turn a client's "basic design" `.pptx` into a **high-end, on-brand deck** — **without adding or
removing a single slide or any text** — by deriving a real design system from the deck itself and
rebuilding every slide to obey it.

See [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) and [`docs/ROADMAP.md`](docs/ROADMAP.md) for the design.

## Principles (hard invariants — the pipeline fails *closed*)

- **Slide parity** — output has exactly the same number of slides, in order.
- **Text preservation** — every text run is kept verbatim; no word added or removed.
- **PowerPoint-native** — built with `python-pptx` only. **No LibreOffice** anywhere in the pipeline.

## Install

```bash
pip install -r requirements.txt
```

## Use

```bash
# structural audit of a deck (master/theme/shape mix)
python -m aippt.cli audit  in.pptx

# derived BrandSpec (the "Design DNA": palette, fonts, type scale)
python -m aippt.cli extract in.pptx

# rebuild high-end; verifies invariants and refuses to certify on violation
python -m aippt.cli build  in.pptx -o out.pptx

# wireframe preview PNGs (Pillow only — no LibreOffice)
python -m aippt.cli preview out.pptx -o preview --only 1,2,3
```

## Status

- **v1 (this branch):** full pipeline — ingest → derive brand → build master-consistent deck →
  invariant gate. Consistent design system (navy/orange/teal, serif headings, eyebrows, footers,
  styled tables). Text preservation enforced on every run.
- **Next (elevation):** dark callout cards, orange icon-rings, dashed dividers, section textures,
  card grids — to match and then exceed the reference.
