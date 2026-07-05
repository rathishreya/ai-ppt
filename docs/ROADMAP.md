# AI-PPT — Implementation Roadmap

Incremental build. **Every phase ends with the invariants green** (I1 slide parity, I2 text
preservation, I3 valid package) and a **visible render diff** so quality is observable, not asserted.
The AI71 sample is the calibration target throughout; Phase 7 proves generalization on other decks.

Legend — **Goal** · **Deliverable** (what you can run/see) · **Exit criteria** (done when).

---

## Phase 0 — Foundations & safety harness
**Goal.** Stand up the skeleton and, first of all, the *guardrails*.
- Repo scaffold (`aippt/` package, `pyproject`, deps pinned), `CLAUDE.md`, config.
- **IR** model + `ingest` (pptx → IR with resolved styling).
- **`textguard`**: text inventory + fingerprint + **I1/I2 verifier** (fail‑closed).
- `qa/render.py`: LibreOffice render wrapper (already validated in this environment).
- **`audit` CLI**: report a deck's master/theme/placeholder conformance (formalizes the throwaway
  script we already used on AI71).

**Deliverable.** `aippt audit samples/AI71/AI71_source.pptx` → structured report; `aippt render …` → PNGs.
**Exit.** Round‑trip (ingest → emit unchanged) passes I1/I2/I3 on AI71 with a golden test.

## Phase 1 — Brand extraction (Design DNA)
**Goal.** Derive `BrandSpec` from the deck deterministically.
- Palette clustering (fills, text, later logo pixels) → primary/bg/accent/neutral, WCAG‑AA pairings.
- Typography detection → 6‑step modular type scale from the 20 ad‑hoc sizes.
- Logo/motif detection; recurring eyebrow/footer/card/page‑number patterns; spacing/grid tokens.

**Deliverable.** `aippt extract …` → `brandspec.json` + a **swatch + type‑scale preview** image.
**Exit.** On AI71 the spec recovers navy/orange/teal, serif+sans, the logo, and the footer pattern.

## Phase 2 — Design‑system construction (the derived master)
**Goal.** Turn `BrandSpec` into real OOXML.
- Write **theme** (fontScheme, clrScheme); build **master** + `txStyles`; build the **layout set**;
  build **components** (eyebrow/footer/card/icon_ring/divider/bg_texture).

**Deliverable.** `aippt design-system brandspec.json` → a rendered **design‑system reference deck**.
**Exit.** Reference deck renders on‑brand and consistent; theme fonts/colors resolve correctly.

## Phase 3 — Classification & layout mapping (the remap)
**Goal.** Decide, per slide, *what goes where*.
- Element **role** + slide **type** classification from structural signals.
- `layout/select` (type → layout) and `layout/map` (content → placeholders, over/under‑flow rules).

**Deliverable.** `aippt plan …` → per‑slide **mapping report** (chosen layout + element→placeholder table).
**Exit.** All 26 AI71 slides get a sensible layout + mapping; every text element is accounted for (I2 dry‑run).

## Phase 4 — Composition v1 (deterministic rebuild)
**Goal.** First full high‑end output.
- Rebuild all 26 slides onto the derived layouts; place **preserved runs verbatim**; apply
  fonts/scale/palette/spacing/grid/hierarchy; basic on‑brand visuals (bg, footer, eyebrow, cards).

**Deliverable.** `aippt build samples/AI71/AI71_source.pptx -o out.pptx` → full styled deck + PNGs.
**Exit.** 26 slides; **I1/I2/I3 green**; side‑by‑side render shows a clear, consistent redesign.

## Phase 5 — Visual‑QA loop
**Goal.** Make it *provably* clean.
- Structural + pixel checks (overflow, off‑slide, overlap, margins, contrast, image DPI).
- `compose → render → critique → adjust` controller (autofit/reflow/re‑chunk), bounded + best‑so‑far.

**Deliverable.** `aippt build --qa …` → deck + **QA report** (zero criticals) + before/after metrics.
**Exit.** No overflow/overlap/contrast criticals on AI71; loop converges deterministically.

## Phase 6 — High‑end elevation (beat the reference)
**Goal.** Push past the reference PDF.
- Generative/vector **textures & motifs**, full **icon system**, richer cards/dividers, refined
  micro‑typography, strict 60/30/10 color, section‑divider treatments, image duotone.
- Wire **optional AI**: image provider (Vibiz MCP / Pollinations / vector) + LLM vision critic — flag‑gated.

**Deliverable.** `aippt build --elevate …`; a documented A/B (source vs reference vs ours).
**Exit.** Blind review rates our output ≥ the reference on hierarchy, consistency, polish.

## Phase 7 — Generalization & hardening
**Goal.** Works on *any* deck, not just AI71.
- Edge cases: tables, charts, SmartArt/groups, RTL/Arabic, dense/sparse slides, missing brand cues.
- Config surface, docs, packaging, CLI polish, test matrix over multiple real decks.

**Deliverable.** Test suite over ≥3 diverse decks; `aippt` usable end‑to‑end; docs complete.
**Exit.** Invariants green across the matrix; graceful degradation on unsupported constructs.

---

### Sequencing notes
- **Guardrails before beauty:** Phase 0's `textguard` gate exists before any transform touches a deck.
- **Deterministic before AI:** Phases 4–5 must look good with **no** network/AI; Phase 6 only *adds*.
- **Calibrate on AI71, prove on others:** every phase is measured on AI71; Phase 7 generalizes.
- Each phase is independently reviewable and leaves `main` in a runnable state.
```
