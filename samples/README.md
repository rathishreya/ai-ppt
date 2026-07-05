# Samples (local calibration decks)

This directory holds sample presentations used to **calibrate and QA** the pipeline
(brand extraction, layout mapping, visual QA). It is the reference set the phases in
[`../docs/ROADMAP.md`](../docs/ROADMAP.md) are measured against.

**The deck binaries are intentionally _not_ committed** (see repo `.gitignore`): sample
decks may be confidential client material, so `*.pptx` / `*.pdf` / images under `samples/`
stay local to the working environment.

## Expected layout

```
samples/
  AI71/
    AI71_source.pptx            # client "basic design" deck (input)         [gitignored]
    AI71_reference_highend.pdf  # high-end reference to match/beat (target)  [gitignored]
```

To work locally, drop the deck files into the matching folder. If you want a sample deck
tracked in git anyway, remove its pattern from `.gitignore` or add it with `git add -f`.
