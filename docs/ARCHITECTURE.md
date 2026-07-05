# AI-PPT — Architecture

> **Mission.** Take a client's `.pptx` ("basic design", sometimes with comments) and produce a
> **world‑class, high‑end deck** — *without adding or removing any slide, and without adding or
> removing any text* — where **every slide obeys one coherent design system (a real "master")**.
> The system must **generalize** to arbitrary decks, not just the AI71 sample.

This document is the agreed technical design. It is deliberately detailed; the companion
[`ROADMAP.md`](./ROADMAP.md) sequences the build.

---

## 1. Problem reframing (what the AI71 sample taught us)

We audited the provided sample (`samples/AI71/AI71_source.pptx`) before designing:

| Finding | Evidence | Consequence for architecture |
|---|---|---|
| The embedded **master is empty** | master `name=""`, **1 blank layout `DEFAULT`, 0 placeholders**; theme = stock Office (Calibri / Calibri Light, default palette) | "Follow the *embedded* master" literally would **delete** the design. We must **establish** a master, not read one. |
| The "design" is **faked per‑shape** | **275 free‑floating text boxes**, **0 placeholders used**, **20 distinct font sizes** (8.5→44 pt), fonts forced to Cambria/Calibri per shape | The core job is to **replace ad‑hoc per‑shape styling with one enforced system**. |
| Content is rich & mixed | 26 slides, 413 text runs, 385 auto‑shapes, 49 pictures, 9 tables, notes on ~half | Remap/compose must handle text, images, tables (and later charts/SmartArt). |
| The high‑end reference is a **redesign, not a reflow** | reference PDF = **14 pages** (source = 26); introduces serif display type, orange accents, icon‑ring cards, dashed dividers, contour textures, consistent footer | The look we must reach is a **full design system**, and the reference **broke slide‑count preservation** — we will not. |

**Decisions locked with the user:**

1. **Design‑system source → *derive from the deck*.** Extract the brand DNA already present
   (navy/teal/orange palette, serif+sans pairing, the AI71 logo, eyebrow/card/footer patterns),
   formalize it into a real master + type scale + component library, then rebuild every slide to obey it.
2. **Designer brain → *deterministic engine + optional AI*.** Reproducible, offline‑capable core;
   AI (text advisor, image generation, vision critic) is a **pluggable, optional** enhancement.
3. **Layout freedom → *full remap to placeholders*.** Rebuild each slide onto our constructed layouts.
4. **Visual additions → *on‑brand additions allowed*** (backgrounds, dividers, icons, textures,
   AI imagery where it helps) — **never adds text or slides**.
5. **Preserve → 26 slides, every text run.** Strict. The reference's 26→14 merge is off‑limits.

---

## 2. Hard invariants (the system fails *closed* on any violation)

These are non‑negotiable gates enforced in code, verified on every run:

- **I1 — Slide parity.** `count(out) == count(in)` and slide order preserved.
- **I2 — Text preservation.** The normalized **text inventory** (every run of visible text, per
  slide → shape → paragraph → run) is **equal** before and after. No word/character added or removed.
  - *In scope:* text in placeholders, text boxes, table cells, grouped shapes, chart data labels/titles.
  - *Out of scope:* text baked into raster images (we cannot and must not fabricate it).
  - Comparison is **whitespace‑normalized** and **order‑aware within a slide** (a fair equality that
    ignores trailing spaces/line‑wrap but not content).
- **I3 — Validity.** Output opens in PowerPoint **and** LibreOffice; package validates.
- **Fail‑closed:** if any invariant is violated, the pipeline **errors and refuses to emit** (or emits
  to a quarantine path with a diff), rather than shipping a broken/altered deck.

`textguard` (below) is the module that makes I1/I2 mechanical, not aspirational.

---

## 3. Pipeline (stages)

```
          ┌──────────────────────────────────────────────────────────────────────────┐
          │                              aippt.pipeline                                │
          └──────────────────────────────────────────────────────────────────────────┘
 (1) Ingest ─► (2) Text Inventory ─► (3) Brand Extraction ─► (4) Design‑System Build
      │             (baseline I2)          "Design DNA"          theme+master+layouts+components
      ▼                                        │                          │
   Deck IR                                 BrandSpec  ───────────────────►│
      │                                                                   ▼
 (5) Classify ─► (6) Layout Select + Content Map (remap) ─► (7) Compose (deterministic + optional AI)
      │                                                                   │
      ▼                                                                   ▼
 (8) Render + Visual‑QA loop  ◄──── adjust ────►  (7)                 output.pptx (draft)
      │  (LibreOffice → PNG; overflow/contrast/overlap/margins)
      ▼
 (9) Invariant Gate (re‑extract text, check I1/I2/I3; fail‑closed) ─► (10) Emit .pptx (+ PDF/PNG + QA report)
```

**Stage detail**

1. **Ingest & IR.** Open with `python-pptx` (+ `lxml` for what python-pptx doesn't expose). Build an
   intermediate representation (`Deck → Slide → Element → Paragraph → Run`) capturing text, geometry
   (EMU), current styling, shape kind, image bytes/refs, table/chart structure. The IR decouples the
   rest of the pipeline from OOXML quirks.
2. **Text inventory & fingerprint.** Serialize the canonical text inventory and hash it. This is the
   **I2 baseline** compared against at stage 9.
3. **Brand extraction ("Design DNA").** Analyze the *whole* deck to derive a **BrandSpec** (§5).
4. **Design‑system construction.** Turn BrandSpec into a **real** OOXML design system in the output:
   a proper **theme** (fontScheme + clrScheme), a **slide master** with true placeholders and
   `txStyles` (the type scale + bullets), a set of **layouts**, and stampable **components** (§6).
5. **Per‑slide classification.** Classify each element's **role** (title/eyebrow/body/bullets/metric/
   caption/footer/image/table/chart/logo) and the slide's **type** (title/section/statement/agenda/
   1‑col/2‑col/content+visual/metrics/table/chart/closing) from structural signals (size rank,
   position, length, shape kind, repetition across deck).
6. **Layout selection + content mapping (the remap).** Choose the best target layout for the slide
   type; assign each classified element to a target **placeholder**; resolve over/under‑flow.
7. **Composition.** Instantiate the layout; place **preserved runs verbatim** into placeholders;
   apply the design system (theme‑referenced fonts, type‑scale sizes, palette colors, spacing/grid,
   hierarchy, alignment); add on‑brand visuals. Deterministic rules do the bulk; optional AI refines.
8. **Render + visual‑QA loop.** Render each output slide (LibreOffice headless → PDF → PNG) and run
   automated checks; feed defects back to (7) to adjust; bounded iterations to converge (§8).
9. **Invariant gate.** Re‑extract the text inventory from the *output*, compare to the baseline
   (I2), assert slide parity (I1), validate the package (I3). Fail‑closed.
10. **Emit.** Write final `.pptx`, plus optional PDF/PNG previews and a machine‑readable **QA report**.

---

## 4. Key OOXML/`python-pptx` realities that shape the design

(Confirmed in the research pass; these drive concrete choices.)

- **`python-pptx` does not resolve inheritance.** `run.font.size/name/bold` return `None` when a value
  is *inherited*. So "what will actually render" must be computed by us by walking
  slide → layout → master → theme via `lxml`. → We build the IR with **explicit resolved values**.
- **There is no `slide.set_layout`.** The slide→layout link is a *relationship* only. → Instead of
  mutating a slide onto a foreign layout, we **rebuild** each output slide **on the layouts we
  construct** (stage 4), copying preserved runs in. This sidesteps the hardest OOXML operation.
- **Placeholder matching:** slide→layout by `idx`; layout→master by `type`. Our constructed layouts
  use clean, known idx/type maps so composition is deterministic.
- **"Conform" is often *deletion of overrides*** (drop `a:xfrm`/`a:latin`/`sz` to re‑inherit) — a
  technique we use for tables/charts we keep in place, but the primary path is rebuild‑onto‑master.
- **Autofit** (`a:bodyPr` `spAutoFit`/`normAutofit fontScale`) is central to the overflow‑QA loop.

---

## 5. Brand extraction → `BrandSpec`

Deterministic analysis of the whole deck yields a single JSON contract consumed by everything downstream:

- **Palette.** Cluster all colors used (shape fills, text colors, and — later — logo pixels) →
  identify **primary/dark**, **background/light**, **accent(s)**, and a **neutral ramp**. For AI71 we
  expect to recover navy + orange + teal + grays. Enforce **WCAG‑AA** foreground/background pairings.
- **Typography.** Detect **heading** vs **body** families (AI71: Cambria serif / Calibri sans) and the
  size distribution → collapse 20 ad‑hoc sizes into a **disciplined modular type scale** (~6 steps:
  display, H1, H2, body, caption, eyebrow) mapped to roles.
- **Logo & marks.** Detect recurring picture(s) / wordmark ("AI71") and brand motifs (the contour
  texture) for consistent placement and regeneration.
- **Recurring components.** Detect the **eyebrow** label, **footer** string
  (`AI71 · … · Confidential (خاص) · N`), **card** and **page‑number** patterns.
- **Tokens.** Spacing scale, corner radius, grid (columns/margins/gutters), divider styles.

`BrandSpec` is versioned and inspectable (`extract` CLI emits it + a swatch/type preview), so a human
(or the optional AI advisor) can sanity‑check the DNA before a full rebuild.

---

## 6. Design‑system construction (the derived "master")

From `BrandSpec` we write real OOXML:

- **Theme** — `fontScheme` (major=heading, minor=body) and `clrScheme` (dk/lt/accents) so runs can
  reference `+mj-lt`/`+mn-lt` and scheme colors instead of hardcoded values.
- **Slide master** — true placeholders + `p:txStyles` encoding the type scale and bullet system.
- **Layouts** — a curated set: `TITLE`, `SECTION`, `STATEMENT/QUOTE`, `AGENDA`, `CONTENT_1COL`,
  `CONTENT_2COL`, `CONTENT_VISUAL`, `METRICS/KPI`, `TABLE`, `CHART`, `CLOSING` — each placing
  placeholders on a shared grid.
- **Components** — parametric builders (`eyebrow`, `footer`, `card`, `icon_ring`, `divider`,
  `bg_texture`) the composer stamps. This is what makes output *consistent by construction*.

A `design-system` command renders a reference deck of the derived system for review.

---

## 7. The deterministic "designer brain" (starter rule set)

The engine encodes expert judgment as rules (top ~20, expanded in ROADMAP Phase 6):

1. One modular **type scale**; nothing off‑scale. 2. Strong **size/weight/color hierarchy** without
changing words. 3. **60/30/10** color discipline (neutral/primary/accent). 4. **WCAG‑AA** contrast
enforced. 5. Consistent **safe margins** + column **grid**. 6. Generous, *budgeted* **whitespace**.
7. **Optical alignment** to the grid. 8. One idea per region; **chunk** long bodies. 9. **Eyebrow →
headline → content** rhythm on content slides. 10. **Consistent footers / page numbers / eyebrows**
across the deck. 11. Uniform **iconography** (one style, recolored to accent). 12. **Cards** for
parallel items; consistent padding/radius/elevation. 13. **Dashed/rule dividers** for scannability.
14. **Data‑ink** cleanup on tables/charts; palette‑tinted. 15. **Section dividers** get texture/
color blocks. 16. Images **reframed** to grid, optional brand duotone. 17. Bilingual **EN/AR** aware
(RTL‑safe). 18. Bullets → disciplined indents/leading. 19. Title slide = focal hierarchy + motif.
20. **Nothing overflows**; autofit + reflow guarantee fit (enforced by QA loop).

Deterministic first → same input yields same output. AI only *proposes*; rules and invariants *dispose*.

---

## 8. Rendering & visual‑QA loop

- **Renderer:** LibreOffice headless (`soffice --headless -env:UserInstallation=… --convert-to pdf`
  then `pdftoppm` → PNG). *(Environment note: `libreoffice-impress` was missing and is now installed;
  rendering works.)*
- **Structural checks (from OOXML):** text bbox vs placeholder box (overflow), off‑slide geometry,
  shape overlap, margin/grid conformance, autofit state.
- **Pixel checks (from PNG):** contrast ratios (text vs local background), clipped glyphs, low‑res or
  stretched images, visual balance.
- **Loop:** `compose → render → critique → adjust` (shrink‑to‑fit, reflow, re‑chunk, re‑place),
  **bounded** iterations with a convergence/So‑far‑best tracker to avoid oscillation.
- **Optional:** an LLM **vision critic** on the PNG as an extra aesthetic gate (flagged off by default).

---

## 9. Optional AI & asset generation (pluggable, never required)

Everything works with **zero network / zero AI**. Behind clean interfaces:

- **Text advisor** (LLM): disambiguate role/slide‑type, suggest headline emphasis. *Needs API key.*
- **Image/asset provider:** on‑brand backgrounds/textures/hero images. Providers:
  **deterministic vector textures** (offline default) · **Vibiz MCP** (`vibiz_generate_image`, available
  in dev) · **Pollinations** (deploy‑time; blocked by this dev proxy). All palette‑constrained + cached
  for determinism.
- **Vision critic** (LLM): aesthetic scoring in the QA loop.

If a provider is absent, the system degrades gracefully to deterministic output — still high‑end.

---

## 10. Module map

```
aippt/
  ir.py                  # Deck/Slide/Element/Paragraph/Run dataclasses (resolved values)
  ingest.py              # pptx -> IR (python-pptx + lxml)
  textguard.py           # text inventory, fingerprint, I1/I2 verification (fail-closed)
  brand/
    spec.py              # BrandSpec schema + type scale + tokens
    extract.py           # palette / typography / logo / motif / components -> BrandSpec
  designsystem/
    theme.py             # write theme xml (fontScheme, clrScheme)
    master.py            # build master + layouts + placeholders + txStyles
    components.py        # eyebrow/footer/card/icon_ring/divider/bg_texture builders
  classify.py            # element role + slide type classification
  layout/
    select.py            # slide-type -> layout
    map.py               # content -> placeholder mapping (remap), over/under-flow
  compose.py             # build output slides (place runs verbatim, style, add visuals)
  qa/
    render.py            # LibreOffice render wrapper
    checks.py            # overflow/contrast/overlap/margin/image metrics
    loop.py              # propose->render->critique->adjust controller
  ai/
    base.py              # provider interfaces (no-op defaults)
    text_advisor.py  image_provider.py  vision_critic.py
  pipeline.py            # orchestrates stages 1-10
  config.py              # tunables (grid, margins, scale ratios, QA thresholds, flags)
  cli.py                 # audit | extract | design-system | build | qa commands
tests/
  invariants/            # golden tests: text-preservation & slide-parity on AI71 + others
  fixtures/              # small synthetic decks for edge cases
docs/                    # ARCHITECTURE.md, ROADMAP.md
samples/AI71/            # source + reference (calibration)
```

---

## 11. Tech stack

- **Python 3.11**, `python-pptx` 1.0.2 (+ `lxml` for uncovered OOXML), `Pillow` (pixel QA / logo color).
- **LibreOffice 24.2** (Impress) + `poppler-utils` (`pdftoppm`) for render‑based QA.
- Optional: **Node 22** (only if a JS renderer/helper is ever wanted — not core).
- Optional AI: Anthropic SDK (advisor/critic) · image provider abstraction (Vibiz MCP / Pollinations /
  vector). All optional and flag‑gated.

---

## 12. Top risks & mitigations

| Risk | Mitigation |
|---|---|
| Text accidentally altered/dropped in remap | `textguard` fail‑closed gate (I2) on every run; golden tests |
| LibreOffice ≠ PowerPoint rendering | Prefer **structural** overflow checks; treat pixels as advisory; keep autofit on |
| Over‑aggressive redesign loses meaning | Role/type classification + human‑reviewable BrandSpec & mapping reports; conservative defaults |
| Generalization (not just AI71) | IR + BrandSpec abstractions; edge‑case fixtures; test on multiple decks in Phase 7 |
| SmartArt/charts/grouped shapes | Phase them in; preserve text, restyle conservatively; explicit "unsupported → passthrough" path |
| Arabic / RTL | RTL‑aware text handling from the start; bidi‑safe layouts |
| Determinism with AI in the loop | AI only *proposes*; rules/invariants decide; cache all AI/asset outputs |

---

## 13. What we will **not** do

- Not add/remove slides. Not add/remove/paraphrase text. Not invent text inside images.
- Not depend on any network or API key for a correct, good‑looking result.
- Not ship a deck that fails an invariant.
```
