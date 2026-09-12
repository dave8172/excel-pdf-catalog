"""A store URL in, a *report* out — the catalog as a document rather than a grid.

topdf's own engine (`simple_catalog.py`) draws product boxes onto a template and
is the right thing for a price list you print and hand over. This is the other
shape a shop asks for: a cover that leads with a number, a page that shows what
the range actually looks like before the products arrive, the grid, then the
same lines as a list.

**The structure is not ours.** It comes from `profexpdf`, a neutral report
engine that knows about covers, figures, cards and grouped tables and knows
nothing about shops. Everything in this file is the half that *is* ours: what
counts as a product, what a price is, how a range is grouped, which four
numbers are worth a strip at the top, and every word on the page.

The line to hold when changing this: if a change would need profexpdf to learn
the word "product", it belongs here instead. If it would need this file to
learn the word "millimetre", it belongs there.

    python3 profex_catalog.py https://some-store.example -o report.pdf

Needs Node 22+ and a Chrome, which is a heavier dependency than the reportlab
path — a print peaks at a few hundred MB. That is why this is a command and is
not wired into the hosted export, where one gunicorn worker is already holding
the memory ceiling for the whole box.
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

from PIL import Image

import shopify_catalog as shop

PROFEXPDF = Path(os.environ.get("PROFEXPDF_HOME", Path(__file__).resolve().parent.parent / "profexpdf"))

# Cards get an image each and every image is embedded in the PDF, so the long
# edge is capped here rather than in the report engine: a 2048px product shot
# is 400KB of base64 that nobody can see the benefit of on a 60mm card.
CARD_IMAGE_PX = 700
LOGO_PX = 320


@dataclass
class Line:
    """One sellable thing, in this file's words rather than Shopify's."""
    title: str
    blurb: str
    price: float | None
    price_text: str
    kind: str
    vendor: str
    image: Path | None
    handle: str
    url: str


def _price(product: dict) -> float | None:
    for variant in product.get("variants") or []:
        raw = str(variant.get("price") or "").strip()
        try:
            value = float(raw)
        except ValueError:
            continue
        if value > 0:
            return value
    return None


def read_lines(store: shop.StoreData, *, symbol: str, limit: int, work_dir: Path) -> list[Line]:
    """Shopify's JSON into Lines, with the photos already on disk.

    Same de-duplication rule as the grid engine: shops publish colourways as
    separate products, and where the colour is not in the title the cards come
    out indistinguishable.
    """
    images = work_dir / "img"
    images.mkdir(parents=True, exist_ok=True)
    shop.prefetch_images([
        shop._shopify_sized(str((p.get("images") or [{}])[0].get("src") or ""), CARD_IMAGE_PX)
        for p in store.products if (p.get("images") or [])
    ])

    lines: list[Line] = []
    seen: set[str] = set()
    for product in store.products:
        if len(lines) >= limit:
            break
        if not shop.is_real_product(product):
            continue
        key = re.sub(r"\s+", " ", str(product.get("title") or "")).strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)

        value = _price(product)
        src = str((product.get("images") or [{}])[0].get("src") or "")
        path = _stage_image(shop._shopify_sized(src, CARD_IMAGE_PX), images, len(lines)) if src else None
        handle = str(product.get("handle") or "")
        lines.append(Line(
            title=shop._plain_text(str(product.get("title") or "")).strip() or "Untitled",
            blurb=shop.short_description(product),
            price=value,
            price_text=f"{symbol}{value:,.2f}".rstrip("0").rstrip(".") if value else "",
            kind=(str(product.get("product_type") or "").strip() or "Other"),
            vendor=str(product.get("vendor") or "").strip(),
            handle=handle,
            url=f"{store.origin}/products/{handle}" if handle else store.origin,
            image=path,
        ))
    return lines


def _stage_image(url: str, into: Path, index: int) -> Path | None:
    """Re-encode a fetched photo into a file the report engine can embed.

    Re-encoding rather than handing over the cached bytes is deliberate: the
    engine takes local files only and matches the type to the extension, and
    passing a stranger's bytes through Pillow is a cheap boundary on the way.
    """
    import catalog_exporter

    try:
        catalog_exporter.ensure_cache_dir(catalog_exporter.IMAGE_CACHE_DIR)
        cached = catalog_exporter.cache_path_for_source(url)
        if not cached.exists():
            cached.write_bytes(catalog_exporter.IMAGE_FETCHER(url))
        with Image.open(cached) as img:
            img.load()
            flat = Image.new("RGB", img.size, "white")
            flat.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA", "P") else None)
            flat.thumbnail((CARD_IMAGE_PX, CARD_IMAGE_PX))
            out = into / f"{index:03d}.jpg"
            flat.save(out, "JPEG", quality=82)
            return out
    except Exception:
        # A missing photo is a card without one, exactly as in the grid engine.
        return None


def _stage_logo(brand: shop.BrandProfile, work_dir: Path) -> Path | None:
    if brand.logo is None:
        return None
    try:
        logo = brand.logo.convert("RGBA")
        logo.thumbnail((LOGO_PX, LOGO_PX))
        out = work_dir / "logo.png"
        logo.save(out, "PNG")
        return out
    except Exception:
        return None


def _bands(values: list[float], symbol: str) -> list[dict]:
    """Price bands, chosen from the data rather than fixed.

    A fixed £0–50/50–100/… puts every line of a £15 candle shop in one bucket
    and every line of a furniture shop in the last. Quintile edges rounded to
    something a person would say keep the shape visible at any price point.
    """
    if not values:
        return []
    ordered = sorted(values)
    edges = [ordered[int(len(ordered) * f)] for f in (0.2, 0.4, 0.6, 0.8)]
    edges = sorted({_round_nice(e) for e in edges})
    if not edges:
        return []
    rows, low = [], None
    for edge in edges:
        n = len([v for v in ordered if (low is None or v > low) and v <= edge])
        rows.append({"label": f"{symbol}{_num(low)}–{_num(edge)}" if low is not None else f"up to {symbol}{_num(edge)}", "value": n})
        low = edge
    rows.append({"label": f"over {symbol}{_num(low)}", "value": len([v for v in ordered if v > low])})
    return rows


def _round_nice(value: float) -> float:
    for step in (5, 10, 25, 50, 100, 250, 500, 1000):
        if value <= step * 10:
            return max(step, round(value / step) * step)
    return round(value / 1000) * 1000


def _num(value: float) -> str:
    return f"{value:,.0f}" if float(value).is_integer() else f"{value:,.2f}"


def _sentence(text: str, limit: int = 220) -> str:
    """A shop's tagline, cut at a sentence rather than at a character count.

    The brand reader caps the tagline for a template that has one line for it,
    which put "...arabica and robusta co" under a cover headline. On a page
    with room, a half-word is worse than a shorter sentence.
    """
    clean = re.sub(r"\s+", " ", str(text or "")).strip()
    if not clean:
        return ""
    if len(clean) <= limit and clean[-1] in ".!?":
        return clean
    cut = clean[:limit]
    end = max(cut.rfind(". "), cut.rfind("! "), cut.rfind("? "))
    if end > 40:
        return cut[:end + 1]
    # No sentence ended in range: keep whole words and say it was cut.
    return cut[:cut.rfind(" ")].rstrip(",;:—- ") + "…" if " " in cut else ""


def build_spec(store: shop.StoreData, brand: shop.BrandProfile, lines: list[Line],
               *, symbol: str, work_dir: Path) -> dict:
    priced = [l for l in lines if l.price is not None]
    values = [l.price for l in priced]
    kinds: dict[str, list[Line]] = {}
    for line in lines:
        kinds.setdefault(line.kind, []).append(line)
    ranked = sorted(kinds.items(), key=lambda kv: -len(kv[1]))
    domain = re.sub(r"^https?://", "", store.origin).rstrip("/")
    median = sorted(values)[len(values) // 2] if values else None
    logo = _stage_logo(brand, work_dir)

    figures: list[dict] = [{
        "type": "stats",
        "stats": [
            {"value": str(len(lines)), "label": "lines in the catalogue"},
            {"value": str(len(kinds)), "label": "ranges"},
            {"value": f"{symbol}{_num(median)}" if median else "—", "label": "median price"},
            {"value": f"{symbol}{_num(min(values))}–{_num(max(values))}" if values else "—", "label": "price range"},
        ],
    }]

    if len(ranked) > 1:
        figures.append({
            "type": "bars",
            "title": "What the catalogue covers",
            "caption": f"Lines in each range, across {len(lines)} in total.",
            "labelWidth": "38mm",
            "split": 2 if len(ranked) > 9 else 1,
            "rows": [{"label": name, "value": len(items), "display": str(len(items))}
                     for name, items in ranked[:16]],
        })

    bands = _bands(values, symbol)
    if len(bands) > 1:
        unpriced = len(lines) - len(priced)
        rows = bands + ([{"label": "no price", "value": unpriced, "tone": "absent"}] if unpriced else [])
        figures.append({
            "type": "columns",
            "title": "Where the range sits on price",
            "caption": f"Lines by price band. The bands come from this catalogue, not from a fixed scale.",
            "kind": "scale",
            "rows": rows,
            "meter": {"part": len(priced), "whole": len(lines),
                      "label": f"**{len(priced)} of {len(lines)}** lines publish a price"} if unpriced else None,
            "note": "A band is a fifth of the catalogue, rounded to a figure a person would say — "
                    "so the shape shows where the range actually is, whatever it sells.",
        })

    vendors: dict[str, int] = {}
    for line in lines:
        if line.vendor:
            vendors[line.vendor] = vendors.get(line.vendor, 0) + 1
    if len(vendors) > 1:
        top = sorted(vendors.items(), key=lambda kv: -kv[1])[:8]
        figures.append({
            "type": "bars",
            "title": "Who makes it",
            "caption": f"{sum(vendors.values())} lines carry a maker; {len(vendors)} makers in all.",
            "labelWidth": "38mm",
            "rows": [{"label": name, "value": n, "display": str(n)} for name, n in top],
        })

    figures = [f for f in figures if f]
    pages: list[dict] = [{
        "type": "cover",
        "eyebrow": f"Catalogue · {domain}",
        "title": brand.name,
        "lede": _sentence(brand.tagline)
                or f"Every line published on {domain}, with its photograph, its description and its price.",
        "hero": {"value": str(len(lines)), "label": "lines in this catalogue"},
        "stats": [
            {"value": str(len(kinds)), "label": "ranges"},
            {"value": f"{symbol}{_num(median)}" if median else "—", "label": "median price"},
            {"value": f"{len(priced)}", "label": "with a price"},
            {"value": f"{sum(1 for l in lines if l.image)}", "label": "with a photo"},
        ],
        "columns": [
            {
                "heading": "What this is",
                "items": [
                    f"**Every line was read from {domain} itself** — the shop's own published "
                    "catalogue, not a third-party listing.",
                    "**Prices are the shop's published price** for the first variant that has one. "
                    "Options and sizes can change it.",
                    "**A line with no price is shown as such**, rather than guessed at.",
                ],
            },
            {
                "heading": "What is inside",
                "ordered": True,
                "items": [
                    {"text": "**The range in numbers** — what it covers and where it sits on price.",
                     "note": "page 2"},
                    {"text": "**The catalogue** — every line, with its photograph.", "note": "from page 3"},
                    {"text": "**The list** — the same lines as a table, grouped by range.",
                     "note": "at the back"},
                ],
                "note": f"Read from {domain}. Prices and availability change; the shop is the authority.",
            },
        ],
    }]

    if figures:
        pages.append({
            "type": "section",
            "title": "The range in numbers",
            "intro": "What the catalogue holds, and what it costs. Every figure is counted from "
                     "the lines in this file, so nothing here disagrees with the pages that follow.",
            "figures": figures,
        })

    pages.append({
        "type": "cards",
        "title": "The catalogue",
        "intro": f"Every line published on {domain}.",
        "perRow": 3,
        "items": [{
            "image": {"file": str(line.image)} if line.image else None,
            "title": line.title,
            "subtitle": " · ".join(x for x in (line.kind, line.vendor) if x) or None,
            "body": line.blurb or None,
            "feature": {"value": line.price_text or "—", "label": "each"},
        } for line in lines],
    })

    groups = []
    for name, items in sorted(kinds.items(), key=lambda kv: kv[0].lower()):
        priced_here = [l for l in items if l.price is not None]
        groups.append({
            "title": name,
            "meta": f"{len(items)} {'line' if len(items) == 1 else 'lines'}"
                    + (f" · {symbol}{_num(min(l.price for l in priced_here))}–"
                       f"{_num(max(l.price for l in priced_here))}" if priced_here else ""),
            "rows": [{
                "cells": [
                    {"text": line.title, "lead": True, "link": line.url},
                    line.vendor or None,
                    line.price_text or None,
                ],
                "note": line.blurb or None,
            } for line in sorted(items, key=lambda l: l.title.lower())],
        })

    pages.append({
        "type": "table",
        "title": "The list",
        "intro": "The same lines as a table, grouped by range and alphabetical within it. "
                 "Every name links to the page it was read from.",
        "orientation": "portrait",
        "columns": [
            {"label": "Line", "width": "62mm"},
            {"label": "Maker", "width": "38mm"},
            {"label": "Price", "width": "26mm", "align": "right"},
        ],
        "groups": groups,
    })

    return {
        "brand": {
            "name": brand.name,
            "url": store.origin,
            "color": brand.primary,
            "logo": {"file": str(logo)} if logo else None,
        },
        "document": {
            "title": f"{brand.name} — catalogue",
            "footer": f"{brand.name} · read from {domain}",
            "empty": "not published",
            "watermark": {"text": brand.name},
        },
        "pages": pages,
    }


def build_report(store_url: str, *, output_path: Path, work_dir: Path | None = None,
                 max_products: int = 60, refine=None) -> Path:
    work = Path(work_dir or tempfile.mkdtemp(prefix="profex-"))
    work.mkdir(parents=True, exist_ok=True)

    origin = shop.normalize_store_url(store_url)
    store = shop.fetch_store(origin, max_products=max_products)
    symbol = shop.currency_symbol(store.meta)
    lines = read_lines(store, symbol=symbol, limit=max_products, work_dir=work)
    if not lines:
        raise shop.StoreError("None of that store's published items look like catalogue lines.")
    brand = shop.read_brand(store, refine=refine)

    spec = build_spec(store, brand, lines, symbol=symbol, work_dir=work)
    spec_path = work / "spec.json"
    spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")

    engine = PROFEXPDF / "bin" / "profexpdf.mjs"
    if not engine.exists():
        raise RuntimeError(f"profexpdf not found at {PROFEXPDF} — set PROFEXPDF_HOME")
    result = subprocess.run(
        ["node", str(engine), str(spec_path), "-o", str(output_path)],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        raise RuntimeError(f"profexpdf failed:\n{result.stderr.strip() or result.stdout.strip()}")
    return output_path


if __name__ == "__main__":
    args = sys.argv[1:]
    if not args:
        print(__doc__.strip().splitlines()[-3].strip(), file=sys.stderr)
        sys.exit(1)
    url = args[0]
    out = Path(args[args.index("-o") + 1]) if "-o" in args else Path("catalogue.pdf")
    cap = int(args[args.index("--max") + 1]) if "--max" in args else 60
    keep = Path(args[args.index("--work") + 1]) if "--work" in args else None
    print(f"reading {url} ...", file=sys.stderr)
    path = build_report(url, output_path=out, work_dir=keep, max_products=cap)
    print(f"pdf   -> {path}")
