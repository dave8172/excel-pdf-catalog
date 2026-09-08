"""Build a catalog PDF from a Shopify store URL, from the command line.

    ./.venv/bin/python scripts/shopify_to_catalog.py https://someshop.com
    ./.venv/bin/python scripts/shopify_to_catalog.py someshop.com --refine --products 24

`--refine` turns on the one optional model call (see `brand_refiner`); without
it the run is entirely deterministic and costs nothing. `--compare` runs both
and reports what the model changed, which is how we find out whether it earns
its place.
"""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import brand_refiner  # noqa: E402
import shopify_catalog  # noqa: E402
from shopify_catalog import StoreError  # noqa: E402


def run(url: str, out: Path, products: int, refine: bool) -> shopify_catalog.CatalogResult:
    work = Path(tempfile.mkdtemp(prefix="shopify_catalog_"))
    started = time.monotonic()
    try:
        result = shopify_catalog.build_catalog(
            url,
            work_dir=work,
            output_path=out,
            max_products=products,
            refine=brand_refiner.refine if refine else None,
            status_callback=lambda message: print(f"    {message}"),
        )
    finally:
        shutil.rmtree(work, ignore_errors=True)

    print(f"    -> {out}  ({out.stat().st_size / 1024:.0f}KB, "
          f"{result.product_count} products, {time.monotonic() - started:.1f}s)")
    print(f"       brand   : {result.brand.name!r}  {result.brand.primary} / {result.brand.secondary}")
    print(f"       logo    : {result.brand.logo_url or '(typeset name)'}")
    for note in result.notes:
        print(f"       note    : {note}")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("url")
    parser.add_argument("--out", type=Path, default=None)
    parser.add_argument("--products", type=int, default=shopify_catalog.DEFAULT_MAX_PRODUCTS)
    parser.add_argument("--refine", action="store_true", help="use the model for logo and colours")
    parser.add_argument("--compare", action="store_true", help="run both ways and diff the branding")
    parser.add_argument("--detect", action="store_true",
                        help="also report what the box detector makes of the generated template")
    args = parser.parse_args()

    stem = shopify_catalog.normalize_store_url(args.url).split("//", 1)[-1].replace(".", "-")
    out = args.out or REPO / "exports" / f"{stem}-catalog.pdf"
    out.parent.mkdir(parents=True, exist_ok=True)

    try:
        if args.compare:
            print("rules only:")
            plain = run(args.url, out.with_name(out.stem + "-rules.pdf"), args.products, refine=False)
            print("\nwith the refiner:")
            if not brand_refiner.available():
                print("    ANTHROPIC_API_KEY is not set — skipping")
                return 0
            refined = run(args.url, out.with_name(out.stem + "-refined.pdf"), args.products, refine=True)

            print("\ndifference:")
            for label, a, b in (
                ("logo", plain.brand.logo_url, refined.brand.logo_url),
                ("primary", plain.brand.primary, refined.brand.primary),
                ("secondary", plain.brand.secondary, refined.brand.secondary),
                ("tagline", plain.brand.tagline, refined.brand.tagline),
            ):
                mark = "  same" if a == b else "CHANGED"
                print(f"    {mark} {label:10s} {a or '-'}\n              {'' if a == b else (b or '-')}")
            return 0

        run(args.url, out, args.products, refine=args.refine)
        return 0
    except StoreError as error:
        print(f"\n{error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
