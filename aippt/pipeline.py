"""End-to-end: source .pptx -> derived brand -> rebuilt high-end deck -> invariant gate."""
from __future__ import annotations
from dataclasses import dataclass

from .ingest import load_deck
from .brand import extract_brand, BrandSpec
from .compose import compose
from .textguard import verify, GuardReport


@dataclass
class BuildResult:
    brand: BrandSpec
    report: GuardReport
    out_path: str


def build(source: str, out_path: str, fail_closed: bool = True) -> BuildResult:
    src = load_deck(source)
    brand = extract_brand(src)
    prs = compose(src, brand)
    prs.save(out_path)
    out = load_deck(out_path)          # re-ingest the emitted file and verify
    report = verify(src, out)
    if fail_closed and not report.ok:
        # keep the file for inspection but signal failure loudly
        raise AssertionError("INVARIANT VIOLATION — refusing to certify output:\n" + report.summary())
    return BuildResult(brand=brand, report=report, out_path=out_path)
