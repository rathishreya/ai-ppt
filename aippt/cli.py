"""aippt command-line interface.

  aippt audit   <in.pptx>              structural audit of master/theme/conformance
  aippt extract <in.pptx>              print derived BrandSpec (Design DNA)
  aippt build   <in.pptx> -o out.pptx  rebuild high-end, verify invariants
"""
from __future__ import annotations
import argparse
import sys

from .ingest import load_deck
from .brand import extract_brand
from .textguard import slide_words, slide_charcount


def _cmd_extract(args):
    deck = load_deck(args.input)
    print(extract_brand(deck).to_json())


def _cmd_audit(args):
    deck = load_deck(args.input)
    words = sum(sum(slide_words(s).values()) for s in deck.slides)
    chars = sum(slide_charcount(s) for s in deck.slides)
    print(f"slides={len(deck.slides)}  words={words}  chars={chars}  "
          f"size={deck.width_in:.2f}x{deck.height_in:.2f}in")
    for s in deck.slides:
        kinds = {}
        for sh in s.shapes:
            kinds[sh.kind] = kinds.get(sh.kind, 0) + 1
        print(f"  slide {s.index+1:2d}: shapes={len(s.shapes)} {kinds}")


def _cmd_build(args):
    from .pipeline import build
    try:
        res = build(args.input, args.output, fail_closed=not args.no_gate)
    except AssertionError as e:
        print(str(e)); sys.exit(2)
    print("BRAND:", res.brand.heading_font, "/", res.brand.body_font,
          "| dark", res.brand.dark, "accent", res.brand.accent, "secondary", res.brand.secondary)
    print(res.report.summary())
    print("wrote", res.out_path)


def main(argv=None):
    ap = argparse.ArgumentParser(prog="aippt")
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name, fn in (("audit", _cmd_audit), ("extract", _cmd_extract)):
        p = sub.add_parser(name); p.add_argument("input"); p.set_defaults(func=fn)
    pb = sub.add_parser("build")
    pb.add_argument("input")
    pb.add_argument("-o", "--output", default="out.pptx")
    pb.add_argument("--no-gate", action="store_true", help="don't fail on invariant violation")
    pb.set_defaults(func=_cmd_build)
    po = sub.add_parser("polish", help="in-place polish for decks WITH a real master "
                                       "(keeps master + complex slides, redesigns text slides)")
    po.add_argument("input")
    po.add_argument("-o", "--output", default="polished.pptx")
    po.set_defaults(func=_cmd_polish)
    pp = sub.add_parser("preview", help="wireframe PNGs via Pillow (no LibreOffice)")
    pp.add_argument("input")
    pp.add_argument("-o", "--out-prefix", default="preview")
    pp.add_argument("--only", default=None, help="comma-separated 1-based slide numbers")
    pp.set_defaults(func=_cmd_preview)
    args = ap.parse_args(argv)
    args.func(args)


def _cmd_polish(args):
    from .inplace import polish
    from .ingest import load_deck
    from .textguard import verify
    r = polish(args.input, args.output)
    rep = verify(load_deck(args.input), load_deck(args.output))
    print(f"redesigned {len(r['redesigned'])} text slides, preserved {len(r['preserved'])} complex slides")
    print(f"text preservation: {'PASS' if rep.ok else 'FAIL'}")
    print("wrote", args.output)


def _cmd_preview(args):
    from .preview import render
    only = [int(x) for x in args.only.split(",")] if args.only else None
    render(args.input, args.out_prefix, only)


if __name__ == "__main__":
    main()
