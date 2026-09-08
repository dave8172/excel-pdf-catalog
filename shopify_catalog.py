"""A Shopify store URL in, a finished catalog PDF out.

topdf's original shape asked for two files a stranger does not have lying
around: a product spreadsheet in our five columns, and a template PDF with
empty boxes drawn on it. That is two jobs of work before the tool does any of
its own. This module removes both by reading them off the store.

**The pipeline is deterministic.** Everything that makes the catalog -- the
products, prices, photos, currency, shop name and location -- comes from
Shopify's own public JSON endpoints, not from a model:

    /meta.json               shop name, city, province, country, currency
    /products.json           title, body_html, images, variants, prices

One step resists rules, and only one: looking at a homepage and deciding which
of a dozen images is the brand's logo and what its colours are. Measured on
real stores, the obvious heuristics get it wrong about half the time -- Death
Wish Coffee's page has four images with "logo" in the tag and three of them are
press badges (BuzzFeed, HuffPost, Yahoo); Tentree's first match is a
certification mark. So that step, and nothing else, can call a model.

`read_brand(..., refine=None)` is the whole of it. Pass no refiner and the
tool still works end to end on rules alone, in a neutral palette -- which is
the point. The refiner is a quality knob we can measure, not a dependency in
the middle of the pipeline.

**Not every Shopify-backed store is reachable.** A headless storefront -- the
shop runs on Shopify but the site is Next.js or similar -- serves neither
endpoint (Gymshark 403s, Kotn 404s). Those are detected and reported as such
rather than failing obscurely.
"""

from __future__ import annotations

import colorsys
import csv
import html as html_module
import io
import json
import re
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable
from urllib.parse import urljoin, urlsplit, urlunsplit

from PIL import Image
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as pdf_canvas

from catalog_exporter import NORMAL_EXPORT_RENDER_DPI, get_pdf_page_count, render_pdf_page
from simple_catalog import find_product_boxes
from web.security import ImageFetchError, fetch_remote_document, fetch_remote_image

# The catalog's five columns, in topdf's own spelling.
HEADERS = ["Image URL", "Name", "Description", "Price", "Case Size"]

# Products are fetched a page at a time; Shopify caps `limit` at 250.
PRODUCTS_PER_REQUEST = 250
MAX_PRODUCT_PAGES = 4

# A catalog is built synchronously and every product costs one image download,
# so the ceiling is a wall-clock decision, not a layout one. Twelve boxes fit a
# page; 36 is three pages and about a minute.
DEFAULT_MAX_PRODUCTS = 36

# Shopify's CDN resizes on request. Asking for the display size instead of the
# 2048px original is the single biggest saving in the whole run -- both the
# download and the memory the exporter then holds.
PRODUCT_IMAGE_WIDTH = 700
LOGO_IMAGE_WIDTH = 600

DESCRIPTION_MAX_CHARS = 88

PAGE_W, PAGE_H = A4


class StoreError(Exception):
    """Something about the store means we cannot build a catalog from it.

    Worded for the person who pasted the URL: they chose the input, so they are
    the one who can fix it.
    """


# ---------------------------------------------------------------------------
# Reaching the store
# ---------------------------------------------------------------------------

def normalize_store_url(raw: str) -> str:
    """Fold whatever was pasted into a bare `https://host` origin.

    People paste the address bar, so what arrives is as likely to be
    `myshop.com/collections/all?page=2` or a bare `myshop.com` as it is a
    clean origin.
    """
    text = (raw or "").strip()
    if not text:
        raise StoreError("Paste your Shopify store address first.")
    if "://" not in text:
        text = f"https://{text}"

    parts = urlsplit(text)
    if parts.scheme not in ("http", "https"):
        raise StoreError("That needs to be a web address starting with https://.")
    if not parts.hostname or "." not in parts.hostname:
        raise StoreError(f"{raw.strip()!r} does not look like a store address.")
    if parts.username or parts.password:
        raise StoreError("Remove the username and password from the address.")

    # Always https: a store that only answers on http still redirects, and the
    # fetcher follows that itself.
    return urlunsplit(("https", parts.netloc, "", "", ""))


def _get_text(url: str, accept: str) -> str:
    return fetch_remote_document(url, accept=accept)[1].decode("utf-8", "replace")


def _get_json(url: str) -> object:
    body = _get_text(url, "application/json")
    try:
        return json.loads(body)
    except json.JSONDecodeError as error:
        # A storefront that isn't Shopify-served answers these paths with its
        # own HTML 200 rather than a 404, so a parse failure here is the same
        # diagnosis as a 404 and gets the same message.
        raise StoreError(_NOT_SHOPIFY) from error


_NOT_SHOPIFY = (
    "That address does not answer Shopify's public product feed. It works with stores running "
    "on a normal Shopify storefront — if the shop uses Shopify only for checkout behind its own "
    "custom site, its products are not published this way and the tool cannot read them."
)


@dataclass
class StoreData:
    origin: str
    meta: dict
    products: list[dict]
    home_html: str


def fetch_store(origin: str, *, max_products: int = DEFAULT_MAX_PRODUCTS) -> StoreData:
    """Pull the three public documents a catalog needs."""
    try:
        raw_meta = _get_json(f"{origin}/meta.json")
    except ImageFetchError:
        # meta.json is the newer of the two endpoints and some themes do not
        # serve it. products.json below is the one that decides the verdict.
        raw_meta = {}
    meta = raw_meta if isinstance(raw_meta, dict) else {}

    products: list[dict] = []
    try:
        for page in range(1, MAX_PRODUCT_PAGES + 1):
            payload = _get_json(f"{origin}/products.json?limit={PRODUCTS_PER_REQUEST}&page={page}")
            batch = payload.get("products") if isinstance(payload, dict) else None
            if not batch:
                break
            products.extend(batch)
            # Stop as soon as there is comfortably more than the catalog needs;
            # the filter downstream discards some, so take a margin.
            if len(products) >= max_products * 3 or len(batch) < PRODUCTS_PER_REQUEST:
                break
    except ImageFetchError as error:
        # The guard's own words never reach the caller. "port 8020 is not
        # allowed" and "169.254.169.254 is not a public address" are accurate,
        # and they answer a port-scan for whoever asked -- so the reason stays
        # in the journal and the reply says only that the store did not answer.
        raise StoreError(_NOT_SHOPIFY) from error

    if not products:
        raise StoreError(
            "No published products came back from that store. An empty, password-protected or "
            "not-yet-launched store has nothing to put in a catalog."
        )

    try:
        home_html = _get_text(f"{origin}/", "text/html")
    except ImageFetchError:
        home_html = ""  # branding degrades to the neutral template; products still work

    return StoreData(origin=origin, meta=meta, products=products, home_html=home_html)


# ---------------------------------------------------------------------------
# Colour
# ---------------------------------------------------------------------------

NEUTRAL_PRIMARY = "#2B3440"
NEUTRAL_SECONDARY = "#6B7684"
BOX_LINE_WIDTH = 1.6

# The detector counts a pixel as ink when it is dark (luma <= 200) or colourful
# (saturation >= 70/255). A box outline in a pale brand tint would satisfy
# neither and the template would silently produce nothing, so every line colour
# goes through `_detectable` first.
DETECTABLE_MAX_LUMA = 170


def _rgb(colour: str) -> tuple[int, int, int]:
    value = colour.lstrip("#")
    if len(value) == 3:
        value = "".join(c * 2 for c in value)
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))  # type: ignore[return-value]


def _hex(rgb: tuple[int, int, int]) -> str:
    return "#{:02X}{:02X}{:02X}".format(*(max(0, min(255, round(c))) for c in rgb))


def _luma(colour: str) -> float:
    r, g, b = _rgb(colour)
    return 0.299 * r + 0.587 * g + 0.114 * b


def _shift(colour: str, factor: float) -> str:
    """Lighten (factor > 1) or darken (factor < 1) while keeping the hue."""
    r, g, b = (c / 255 for c in _rgb(colour))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = max(0.0, min(1.0, l * factor))
    return _hex(tuple(c * 255 for c in colorsys.hls_to_rgb(h, l, s)))


def _detectable(colour: str) -> str:
    """Darken a colour until the box detector is certain to see it as ink."""
    current = colour
    for _ in range(8):
        if _luma(current) <= DETECTABLE_MAX_LUMA:
            return current
        current = _shift(current, 0.78)
    return NEUTRAL_SECONDARY


def _readable_on(background: str) -> str:
    return "#FFFFFF" if _luma(background) < 150 else "#16202B"


def _valid_hex(value: object) -> str | None:
    if isinstance(value, str) and re.fullmatch(r"#?[0-9a-fA-F]{6}", value.strip()):
        return "#" + value.strip().lstrip("#").upper()
    return None


def palette_of(image: Image.Image, colours: int = 8) -> list[tuple[int, str, float, float]]:
    """Dominant colours of an image as (pixels, hex, saturation, lightness).

    Flattened onto white first: a logo is usually a transparent PNG, and the
    alpha would otherwise quantize into phantom colours.
    """
    rgba = image.convert("RGBA")
    flattened = Image.alpha_composite(Image.new("RGBA", rgba.size, (255, 255, 255, 255)), rgba)
    sample = flattened.convert("RGB")
    sample.thumbnail((160, 160))
    quantized = sample.quantize(colors=colours, method=Image.MEDIANCUT).convert("RGB")

    out = []
    for count, (r, g, b) in sorted(quantized.getcolors(colours * 4) or [], reverse=True):
        _h, lightness, saturation = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
        out.append((count, _hex((r, g, b)), saturation, lightness))
    return out


def brand_colours_from(image: Image.Image) -> tuple[str, str] | None:
    """The most-used genuinely coloured tones in a logo, if it has any.

    Returns None for a monochrome mark -- which is common, and is why the
    neutral palette has to be a real outcome rather than an error path.
    """
    branded = [
        (count, colour)
        for count, colour, saturation, lightness in palette_of(image)
        if saturation > 0.22 and 0.10 < lightness < 0.88
    ]
    if not branded:
        return None
    primary = branded[0][1]
    secondary = branded[1][1] if len(branded) > 1 else _shift(primary, 1.55)
    return primary, secondary


# ---------------------------------------------------------------------------
# Branding
# ---------------------------------------------------------------------------

# Images that carry the word "logo" but belong to somebody else. Stores put
# press mentions and certification marks in the footer, and those tags match
# every naive "find the logo" rule.
FOREIGN_LOGO_HINTS = (
    "buzzfeed", "huffpost", "huffington", "yahoo", "forbes", "vogue", "gq-", "cnn", "nbc",
    "today-show", "wsj", "nytimes", "techcrunch", "cosmopolitan", "esquire", "menshealth",
    "as-seen", "featured", "press", "award", "certif", "b-corp", "bcorp", "climate",
    "science-based", "1percent", "onepercent", "payment", "visa", "mastercard", "paypal",
    "amex", "klarna", "afterpay", "shop-pay", "google", "facebook", "instagram", "tiktok",
    "trustpilot", "yotpo", "judge-me", "stamped",
)


@dataclass
class LogoCandidate:
    source: str          # which rule produced it, for the refiner and for logs
    url: str
    image: Image.Image | None = None
    note: str = ""

    def describe(self, index: int) -> dict:
        size = f"{self.image.width}x{self.image.height}" if self.image else "unreadable"
        palette = [colour for _n, colour, _s, _l in palette_of(self.image)[:4]] if self.image else []
        return {
            "index": index,
            "found_by": self.source,
            "filename": urlsplit(self.url).path.rsplit("/", 1)[-1][:70],
            "size": size,
            "palette": palette,
        }


@dataclass
class BrandProfile:
    name: str
    tagline: str = ""
    address: str = ""
    contact: str = ""
    primary: str = NEUTRAL_PRIMARY
    secondary: str = NEUTRAL_SECONDARY
    logo: Image.Image | None = None
    logo_url: str = ""
    notes: list[str] = field(default_factory=list)

    @property
    def line(self) -> str:
        return _readable_on(self.primary)


def _shopify_sized(url: str, width: int) -> str:
    """Ask Shopify's CDN for a display-sized copy instead of the original.

    Unescapes first: a URL scraped out of an attribute arrives with its
    ampersands entity-encoded, and `?v=1&amp;width=600` sends the CDN a
    parameter literally named `amp;width`.
    """
    url = html_module.unescape(url)
    if "cdn/shop" not in url and "cdn.shopify.com" not in url:
        return url
    parts = urlsplit(url)
    query = [bit for bit in parts.query.split("&") if bit and not bit.startswith("width=")]
    query.append(f"width={width}")
    return urlunsplit((parts.scheme, parts.netloc, parts.path, "&".join(query), ""))


def _text_of(html: str, pattern: str) -> str:
    match = re.search(pattern, html, re.I | re.S)
    if not match:
        return ""
    return html_module.unescape(re.sub(r"\s+", " ", match.group(1))).strip()


def logo_candidates(html: str, origin: str) -> list[LogoCandidate]:
    """Every image on the homepage that could plausibly be the brand mark.

    Deliberately generous. Narrowing this list is the judgement call, and it is
    made later -- by `_pick_logo` on rules, or by the refiner on a model.
    """
    seen: set[str] = set()
    found: list[LogoCandidate] = []

    def add(source: str, raw_url: str) -> None:
        if not raw_url or raw_url.startswith("data:"):
            return
        url = urljoin(origin + "/", html_module.unescape(raw_url).strip().split()[0])
        if url in seen or urlsplit(url).scheme not in ("http", "https"):
            return
        seen.add(url)
        found.append(LogoCandidate(source=source, url=url))

    # The header is where a logo lives; searching it first means the rule-only
    # path usually gets the right answer before the footer's press badges.
    header = re.search(r"<header\b.*?</header>", html, re.I | re.S)
    for scope, label in ((header.group(0) if header else "", "header"), (html, "page")):
        for tag in re.findall(r"<img\b[^>]*>", scope, re.I):
            if not re.search(r"logo|wordmark|brand", tag, re.I):
                continue
            attribute = re.search(r'\b(?:src|data-src|data-srcset|srcset)\s*=\s*["\']([^"\']+)', tag, re.I)
            if attribute:
                add(f"img[logo] in {label}", attribute.group(1))

    add("apple-touch-icon", _text_of(html, r'<link[^>]+rel=["\'][^"\']*apple-touch-icon[^"\']*["\'][^>]+href=["\']([^"\']+)'))
    add("favicon", _text_of(html, r'<link[^>]+rel=["\'][^"\']*\bicon\b[^"\']*["\'][^>]+href=["\']([^"\']+)'))
    add("og:image", _text_of(html, r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)'))
    return found


def _load_candidates(candidates: list[LogoCandidate], limit: int = 6) -> list[LogoCandidate]:
    """Download and open each candidate, dropping the ones that are not images.

    SVG is the common casualty: plenty of themes use an inline or linked SVG
    wordmark, which Pillow cannot open and reportlab cannot place. Those stores
    fall back to a typeset name, which is a perfectly good catalog header.
    """
    loaded = []
    for candidate in candidates[:limit]:
        try:
            raw = fetch_remote_image(_shopify_sized(candidate.url, LOGO_IMAGE_WIDTH))
            with Image.open(io.BytesIO(raw)) as opened:
                opened.load()
                candidate.image = opened.convert("RGBA")
        except (ImageFetchError, OSError, ValueError) as error:
            candidate.note = f"unreadable ({type(error).__name__})"
            continue
        loaded.append(candidate)
    return loaded


def _pick_logo(candidates: list[LogoCandidate], shop_name: str) -> LogoCandidate | None:
    """The rules-only choice: score each candidate, take the best.

    This is the step the measurements said rules are weakest at, so the scoring
    is written to fail toward "no logo" rather than toward a confident wrong
    one -- a typeset shop name beats somebody else's badge on our cover.
    """
    words = [word for word in re.split(r"[^a-z0-9]+", shop_name.lower()) if len(word) > 2]

    def score(candidate: LogoCandidate) -> float:
        filename = urlsplit(candidate.url).path.lower()
        points = 0.0
        if "header" in candidate.source:
            points += 4
        elif "img[logo]" in candidate.source:
            points += 2
        elif candidate.source == "apple-touch-icon":
            points += 1.5
        elif candidate.source == "favicon":
            points += 1
        if any(word in filename for word in words):
            points += 3
        if any(hint in filename for hint in FOREIGN_LOGO_HINTS):
            points -= 6
        image = candidate.image
        if image is not None:
            # A wordmark is wide; a square 32px favicon is a last resort.
            ratio = image.width / max(1, image.height)
            if 1.6 <= ratio <= 8:
                points += 1.5
            if min(image.size) < 48:
                points -= 1
        return points

    ranked = sorted(candidates, key=score, reverse=True)
    return ranked[0] if ranked and score(ranked[0]) > 0 else None


def _address_from_meta(meta: dict) -> str:
    parts = [meta.get("city"), meta.get("province"), meta.get("country")]
    return ", ".join(str(part) for part in parts if part)


def currency_symbol(meta: dict) -> str:
    """The store's own currency mark, if the exporter can print it.

    `money_format` arrives as something like `${{amount}}` or `&pound;{{amount}}`.
    `simple_catalog.format_price` keeps a leading symbol only when it recognises
    it, so anything else (`CHF`, `kr`) is dropped rather than mangled.
    """
    raw = html_module.unescape(str(meta.get("money_format") or ""))
    lead = raw.split("{{")[0].strip()
    return lead if lead and lead in "£$€¥₹₽₩₪" else ""


def read_brand(
    store: StoreData,
    *,
    refine: Callable[[dict], dict] | None = None,
) -> BrandProfile:
    """Work out what the catalog should look like.

    Rules do the whole job. `refine` -- when given -- is handed the shortlist
    and may overrule the logo choice and the colours; anything it returns is
    validated before it is believed, and a bad answer leaves the rule result
    standing.
    """
    html = store.home_html
    name = str(store.meta.get("name") or "").strip() or _text_of(html, r"<title[^>]*>(.*?)</title>").split("|")[0].strip()
    name = name or urlsplit(store.origin).hostname or "Catalog"

    profile = BrandProfile(
        name=name,
        tagline=_text_of(html, r'<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)')[:110],
        address=_address_from_meta(store.meta),
    )

    candidates = _load_candidates(logo_candidates(html, store.origin))
    chosen = _pick_logo(candidates, name)
    suggested: tuple[str | None, str | None] = (None, None)

    if refine is not None and candidates:
        try:
            verdict = refine({
                "shop_name": name,
                "page_title": _text_of(html, r"<title[^>]*>(.*?)</title>")[:120],
                "meta_description": profile.tagline,
                "candidates": [candidate.describe(index) for index, candidate in enumerate(candidates)],
            })
        except Exception as error:  # a branding hint is never worth failing the export over
            profile.notes.append(f"brand refiner unavailable ({type(error).__name__}); used rules")
        else:
            index = verdict.get("logo_index")
            if isinstance(index, int) and 0 <= index < len(candidates):
                chosen = candidates[index]
                profile.notes.append(f"refiner chose candidate {index} ({chosen.source})")
            elif index is None:
                chosen = None
                profile.notes.append("refiner found no real logo; using the name")
            suggested = (_valid_hex(verdict.get("primary_color")), _valid_hex(verdict.get("secondary_color")))
            tagline = verdict.get("tagline")
            if isinstance(tagline, str) and tagline.strip():
                profile.tagline = tagline.strip()[:110]

    if chosen is not None and chosen.image is not None:
        profile.logo = chosen.image
        profile.logo_url = chosen.url

    # The model chooses *which image*; the pixels decide what colour it is.
    # Splitting it that way is not tidiness -- asked for both, the refiner
    # named a near-black for Death Wish Coffee, whose logo it had just
    # correctly identified and whose red is right there in it. Judgement is
    # what the model is for; measuring a colour is not.
    pair = brand_colours_from(profile.logo) if profile.logo is not None else None
    if pair:
        profile.primary, profile.secondary = pair
        profile.notes.append(f"colours read from the logo ({profile.primary})")
    elif all(suggested):
        profile.primary, profile.secondary = suggested  # type: ignore[assignment]
        profile.notes.append(f"logo is monochrome; refiner suggested {profile.primary}")
    else:
        profile.notes.append("logo is monochrome and unrefined; neutral palette")

    if profile.logo is None:
        profile.notes.append("no usable logo image; the shop name is typeset instead")

    # Whatever produced them, the colours have to survive contact with the
    # detector and stay readable.
    profile.secondary = _detectable(profile.secondary)
    return profile


# ---------------------------------------------------------------------------
# Products
# ---------------------------------------------------------------------------

# Shopify carts routinely contain things that are not products: shipping
# protection, gift cards, tips, warranties. They are recognisable without
# guessing -- nothing physical ships, or the type says so outright.
#
# Matched against `product_type` only, as a substring, because that field is
# written by the shop and reads like "return,package_protection" or "Gift Card".
NON_PRODUCT_TYPES = (
    "gift card", "gift_card", "giftcard", "package_protection", "shipping protection",
    "insurance", "warranty", "donation",
)

# Tags are matched as *whole tags*, never as substrings. Getting that wrong is
# not hypothetical: "return" as a substring matches `loop::returnable => true`,
# a returns-app tag that sits on virtually every real product, and it silently
# removed 215 of Allbirds' 294 items before anyone looked at the count.
NON_PRODUCT_TAGS = frozenset({
    "gift card", "gift_card", "giftcard", "package protection", "package_protection",
    "shipping protection", "redo-package-protection", "route-protection", "tip", "donation",
})


def _plain_text(raw_html: str) -> str:
    text = re.sub(r"<(script|style)\b.*?</\1>", " ", raw_html or "", flags=re.I | re.S)
    text = re.sub(r"<[^>]+>", " ", text)
    return re.sub(r"\s+", " ", html_module.unescape(text)).strip()


def short_description(product: dict) -> str:
    """A card-sized line, in the store's own words.

    The store already wrote a description; the only problem is length. Taking
    its first sentence and cutting on a word boundary keeps the shop's voice,
    which reads better on a catalog card than a paraphrase would.
    """
    text = _plain_text(product.get("body_html") or "")
    if not text:
        return str(product.get("product_type") or "").strip()

    sentence = re.split(r"(?<=[.!?])\s+", text)[0]
    if len(sentence) <= DESCRIPTION_MAX_CHARS:
        return sentence.rstrip(".")

    clipped = text[:DESCRIPTION_MAX_CHARS].rsplit(" ", 1)[0].rstrip(",;:-")
    return f"{clipped}…"


def is_real_product(product: dict) -> bool:
    variants = product.get("variants") or []
    if not variants:
        return False
    if not (product.get("images") or []):
        return False  # a catalog card with no photo is not worth a slot

    kind = str(product.get("product_type") or "").lower()
    if any(marker in kind for marker in NON_PRODUCT_TYPES):
        return False
    if {str(tag).strip().lower() for tag in product.get("tags") or []} & NON_PRODUCT_TAGS:
        return False

    # Nothing that ships is nothing to put in a product catalog. This is what
    # catches "Free Returns Coverage" and friends without a name blocklist.
    #
    # It deliberately does *not* also require a price. A price of 0.00 on a
    # shipping product means "price on request", which is normal in B2B -- 18
    # of Metrixplus Instruments' 97 items are quoted that way -- and dropping
    # them silently removes real products from the owner's own catalog. The
    # card simply omits the price line, exactly as it does for a blank Price
    # cell in an uploaded spreadsheet.
    return True


def _price_of(variant: dict) -> str:
    raw = str(variant.get("price") or "").strip()
    try:
        return raw if float(raw) > 0 else ""
    except ValueError:
        return ""


def product_rows(products: list[dict], *, symbol: str, limit: int) -> list[list[str]]:
    """Turn Shopify's JSON into the five columns topdf already exports.

    Products with an identical title are collapsed to the first. Shops publish
    colourways as separate products, and where the colour is not in the title
    the cards come out indistinguishable -- Tentree's "Lake Tentree T-Shirt"
    fills three slots on one page with the same picture and price. A shop that
    *does* put the colour in the title (Allbirds does) has distinct titles and
    is untouched.
    """
    rows: list[list[str]] = []
    seen_titles: set[str] = set()
    for product in products:
        if len(rows) >= limit:
            break
        if not is_real_product(product):
            continue

        title_key = re.sub(r"\s+", " ", str(product.get("title") or "")).strip().lower()
        if title_key in seen_titles:
            continue
        seen_titles.add(title_key)

        images = product.get("images") or []
        variants = product.get("variants") or []
        price = next((_price_of(variant) for variant in variants if _price_of(variant)), "")

        rows.append([
            _shopify_sized(str(images[0].get("src") or ""), PRODUCT_IMAGE_WIDTH),
            _plain_text(str(product.get("title") or "")).strip() or "Untitled product",
            short_description(product),
            f"{symbol}{price}" if price else "",
            "",  # Case Size has no Shopify equivalent; the card omits it
        ])
    return rows


def write_product_file(rows: list[list[str]], path: Path) -> Path:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(HEADERS)
        writer.writerows(rows)
    return path


# ---------------------------------------------------------------------------
# The template
# ---------------------------------------------------------------------------

MARGIN = 42.0
GUTTER = 14.0
COLUMNS = 3
COVER_ROWS = 3
INNER_ROWS = 4

COVER_TOP = PAGE_H - 156
INNER_TOP = PAGE_H - 86
GRID_BOTTOM = 52.0

COVER_CAPACITY = COVER_ROWS * COLUMNS
INNER_CAPACITY = INNER_ROWS * COLUMNS


def _box_height(top: float, rows: int) -> float:
    return (top - GRID_BOTTOM - (rows - 1) * GUTTER) / rows


def plan_pages(count: int) -> list[tuple[str, int]]:
    """How many boxes each template page should carry, for exactly `count` products.

    The uploaded-template path cannot know this -- the user drew the boxes, and
    a half-empty last page is their layout. Here the template is ours, so an
    empty box is a defect: the first run left nine of them on the back page.
    Drawing exactly as many boxes as there are products also means the
    exporter's "repeat the last page" rule never has to fire.
    """
    if count <= COVER_CAPACITY:
        return [("cover", count)]

    pages = [("cover", COVER_CAPACITY)]
    remaining = count - COVER_CAPACITY
    while remaining > 0:
        take = min(INNER_CAPACITY, remaining)
        pages.append(("inner", take))
        remaining -= take
    return pages


LOGO_PLATE_PADDING = 7.0


def _baked_background(image: Image.Image) -> str | None:
    """The background colour a logo carries with it, if it has one.

    A transparent PNG has none -- it takes whatever is behind it, which is what
    we want. A JPEG has its background baked in (Shopify even pads them:
    `pad_color=ffffff`), and dropping that straight onto a dark brand band
    leaves the mark looking like a sticker somebody stuck on. Sampling the
    border tells the two apart.
    """
    if image.getchannel("A").getextrema()[0] < 250:
        return None  # genuinely transparent somewhere; composite it

    edge = image.convert("RGB")
    width, height = edge.size
    step = max(1, min(width, height) // 12)
    samples = (
        [edge.getpixel((x, 0)) for x in range(0, width, step)]
        + [edge.getpixel((x, height - 1)) for x in range(0, width, step)]
        + [edge.getpixel((0, y)) for y in range(0, height, step)]
        + [edge.getpixel((width - 1, y)) for y in range(0, height, step)]
    )
    average = tuple(sum(channel) / len(samples) for channel in zip(*samples))
    spread = max(max(abs(pixel[i] - average[i]) for i in range(3)) for pixel in samples)
    return _hex(average) if spread < 28 else None  # a uniform border, so it is a background


def _draw_logo(page: pdf_canvas.Canvas, brand: BrandProfile, x: float, y: float, max_w: float, max_h: float) -> float:
    """Place the logo, or typeset the name. Returns the width it used."""
    image = brand.logo
    if image is None:
        page.setFillColor(HexColor(brand.line))
        page.setFont("Helvetica-Bold", 19)
        name = _clip_words(brand.name, 34)
        page.drawString(x, y - 7, name)
        return page.stringWidth(name, "Helvetica-Bold", 19)

    plate = _baked_background(image)
    # The plate steals room, so the mark itself is measured against what is
    # left rather than shrinking the lockup below the space allowed.
    inset = LOGO_PLATE_PADDING if plate else 0.0
    scale = min((max_w - 2 * inset) / image.width, (max_h - 2 * inset) / image.height)
    width, height = image.width * scale, image.height * scale

    if plate:
        # A rounded plate in the logo's own background colour: the same pixels,
        # but now reading as a deliberate lockup instead of a pasted rectangle.
        page.setFillColor(HexColor(plate))
        page.roundRect(x, y - height / 2 - inset, width + 2 * inset, height + 2 * inset,
                       4, stroke=0, fill=1)
        x += inset

    behind = plate or brand.primary
    flattened = Image.alpha_composite(Image.new("RGBA", image.size, _rgb(behind) + (255,)), image)
    page.drawImage(ImageReader(flattened.convert("RGB")), x, y - height / 2,
                   width=width, height=height, mask=None)
    return width + (2 * inset if plate else 0)


def _to_pixels(x: float, y: float, width: float, height: float) -> tuple[int, int, int, int]:
    """A reportlab rectangle as the pixel box the exporter fills.

    Two changes of frame at once: points to pixels at the export render dpi,
    and reportlab's bottom-left origin to Pillow's top-left one.
    """
    scale = NORMAL_EXPORT_RENDER_DPI / 72.0
    left = round(x * scale)
    top = round((PAGE_H - y - height) * scale)
    return left, top, left + round(width * scale), top + round(height * scale)


def _draw_boxes(
    page: pdf_canvas.Canvas, brand: BrandProfile, count: int, top: float, box_h: float
) -> list[tuple[int, int, int, int]]:
    """Exactly `count` empty outlined boxes, row-major, and where each one is.

    Box height is passed in rather than divided out of the space available, so
    a part-filled final page keeps the same card size as the full pages before
    it instead of stretching three products down the whole sheet.

    The returned pixel boxes go straight to the exporter, so the drawing and
    the filling come from one calculation rather than two that have to agree.
    """
    box_w = (PAGE_W - 2 * MARGIN - (COLUMNS - 1) * GUTTER) / COLUMNS
    boxes: list[tuple[int, int, int, int]] = []

    page.setStrokeColor(HexColor(_detectable(brand.secondary)))
    page.setLineWidth(BOX_LINE_WIDTH)
    for index in range(count):
        row, column = divmod(index, COLUMNS)
        x = MARGIN + column * (box_w + GUTTER)
        y = top - (row + 1) * box_h - row * GUTTER
        page.roundRect(x, y, box_w, box_h, 6, stroke=1, fill=0)
        # Inset by the stroke so a product is never drawn over its own outline.
        boxes.append(_to_pixels(x + BOX_LINE_WIDTH, y + BOX_LINE_WIDTH,
                                box_w - 2 * BOX_LINE_WIDTH, box_h - 2 * BOX_LINE_WIDTH))
    return boxes


def _draw_footer(page: pdf_canvas.Canvas, brand: BrandProfile, store_domain: str) -> None:
    page.setFillColor(HexColor(brand.secondary))
    page.setFont("Helvetica", 7.5)
    pieces = [piece for piece in (brand.address, brand.contact, store_domain) if piece]
    page.drawCentredString(PAGE_W / 2, 24, "   ·   ".join(pieces)[:150])


def _clip_words(text: str, limit: int) -> str:
    """Cut to `limit` on a word boundary. Mid-word truncation looks like a bug."""
    text = (text or "").strip()
    if len(text) <= limit:
        return text
    return text[:limit].rsplit(" ", 1)[0].rstrip(",;:-") + "…"


def build_template(
    brand: BrandProfile, store_domain: str, path: Path, product_count: int
) -> tuple[list[list[tuple[int, int, int, int]]], list[int]]:
    """Draw a catalog template in the store's own branding, sized to the products.

    Page 1 is a cover carrying the logo and a title, so it holds nine products;
    every page after it is the denser twelve-box page. The last page gets
    exactly the boxes still needed.

    Returns the boxes for each *output* page and which physical template page
    each of those is drawn on. **At most three physical pages are ever drawn**
    — cover, full inner, and a short final inner — however long the catalog is,
    because every full inner page is the same picture. A thousand products is
    then 84 output pages built from 3 renders rather than 84, which is the
    difference between 20MB of memory and 549MB of it.
    """
    pages = plan_pages(product_count)

    # Physical designs, deduplicated: a full inner page is drawn once and used
    # by every output page that is full.
    designs: list[tuple[str, int]] = []
    sources: list[int] = []
    for spec in pages:
        if spec not in designs:
            designs.append(spec)
        sources.append(designs.index(spec))

    page = pdf_canvas.Canvas(str(path), pagesize=A4)
    ink = brand.line
    per_design: list[list[tuple[int, int, int, int]]] = []

    for kind, boxes in designs:
        if kind == "cover":
            page.setFillColor(HexColor(brand.primary))
            page.rect(0, PAGE_H - 132, PAGE_W, 132, stroke=0, fill=1)
            _draw_logo(page, brand, MARGIN, PAGE_H - 64, max_w=210, max_h=54)

            page.setFillColor(HexColor(ink))
            page.setFont("Helvetica-Bold", 21)
            page.drawRightString(PAGE_W - MARGIN, PAGE_H - 58, "Product Catalogue")
            if brand.tagline:
                page.setFont("Helvetica", 9)
                page.drawRightString(PAGE_W - MARGIN, PAGE_H - 74, _clip_words(brand.tagline, 76))

            top, box_h = COVER_TOP, _box_height(COVER_TOP, COVER_ROWS)
        else:
            page.setFillColor(HexColor(brand.primary))
            page.rect(0, PAGE_H - 62, PAGE_W, 62, stroke=0, fill=1)
            _draw_logo(page, brand, MARGIN, PAGE_H - 31, max_w=132, max_h=30)
            page.setFillColor(HexColor(ink))
            page.setFont("Helvetica", 9)
            page.drawRightString(PAGE_W - MARGIN, PAGE_H - 35, _clip_words(brand.name, 46))

            top, box_h = INNER_TOP, _box_height(INNER_TOP, INNER_ROWS)

        per_design.append(_draw_boxes(page, brand, boxes, top=top, box_h=box_h))
        _draw_footer(page, brand, store_domain)
        page.showPage()

    page.save()
    return [per_design[index] for index in sources], sources


def detected_boxes(path: Path) -> list[int]:
    """What the box detector makes of our own template, page by page.

    Not used by the export -- that takes the coordinates we drew at. This is
    for the CLI's `--detect` check, which is how the two stayed honest with
    each other while the template design was being worked out.
    """
    counts = []
    for number in range(1, get_pdf_page_count(path) + 1):
        rendered = render_pdf_page(path, number, dpi=NORMAL_EXPORT_RENDER_DPI)
        try:
            counts.append(len(find_product_boxes(rendered)))
        finally:
            rendered.close()
    return counts


# ---------------------------------------------------------------------------
# The whole thing
# ---------------------------------------------------------------------------

# Image downloads are the whole cost of a run: 36 products took 62s cold and
# 8.5s once cached, so ~90% of the wall clock was waiting on a CDN one file at
# a time. It is pure network wait, so threads are the right tool. Twelve is
# chosen against the CDN's patience and this box's file handles, not its CPU.
IMAGE_FETCH_WORKERS = 12


def prefetch_images(urls: list[str], *, progress: Callable[[int, int], None] | None = None) -> None:
    """Warm the exporter's on-disk image cache in parallel, before the render loop.

    Deliberately a *prewarm* rather than a rewrite of the drawing code: the
    exporter already checks the cache first and falls back to drawing a card
    without a photo, so filling the cache concurrently makes the serial loop
    fast while leaving its error handling exactly as it was. A failure here is
    not raised -- it just means that one product is fetched (and fails) again
    in the loop, which is where the existing handling lives.
    """
    import catalog_exporter
    from catalog_exporter import cache_path_for_source, download_image, ensure_cache_dir

    wanted = [url for url in dict.fromkeys(urls) if url.startswith(("http://", "https://"))]
    if not wanted:
        return
    ensure_cache_dir(catalog_exporter_image_cache())

    # Read the fetcher off the module rather than importing the name: `app.py`
    # rebinds `IMAGE_FETCHER` at startup to the SSRF-guarded one, and a
    # from-import here would capture whatever it was at import time -- which in
    # the CLI is None. That mistake wrote empty files into the shared cache and
    # every product silently lost its photo, so the fallback is now the same
    # `download_image` the exporter itself falls back to.
    fetcher = catalog_exporter.IMAGE_FETCHER or download_image
    done = 0

    def fetch(url: str) -> None:
        target = cache_path_for_source(url)
        if target.exists():
            return
        data = fetcher(url)
        if not data:
            return  # never cache an empty body; a poisoned entry never expires
        # Write via a per-thread temporary file and rename: several products can
        # share a photo, and two threads writing the same cache path directly
        # would interleave into a corrupt image.
        scratch = target.with_suffix(target.suffix + f".{threading.get_ident():x}.part")
        scratch.write_bytes(data)
        scratch.replace(target)

    with ThreadPoolExecutor(max_workers=IMAGE_FETCH_WORKERS) as pool:
        for future in as_completed([pool.submit(fetch, url) for url in wanted]):
            done += 1
            try:
                future.result()
            except Exception:
                pass  # the render loop retries this one and degrades the card
            if progress:
                progress(done, len(wanted))


def catalog_exporter_image_cache() -> Path:
    from catalog_exporter import IMAGE_CACHE_DIR

    return IMAGE_CACHE_DIR


@dataclass
class CatalogResult:
    output: Path
    brand: BrandProfile
    product_count: int
    store_name: str
    notes: list[str]


def build_catalog(
    store_url: str,
    *,
    work_dir: Path,
    output_path: Path,
    max_products: int = DEFAULT_MAX_PRODUCTS,
    refine: Callable[[dict], dict] | None = None,
    status_callback: Callable[[str], None] | None = None,
) -> CatalogResult:
    """Store URL in, catalog PDF out."""
    from simple_catalog import export_catalog

    def say(message: str) -> None:
        if status_callback:
            status_callback(message)

    origin = normalize_store_url(store_url)
    domain = urlsplit(origin).hostname or ""

    say("Reading the store...")
    store = fetch_store(origin, max_products=max_products)

    say("Picking out the products...")
    rows = product_rows(store.products, symbol=currency_symbol(store.meta), limit=max_products)
    if not rows:
        raise StoreError(
            "None of that store's published items look like catalog products — they had no photo, "
            "no price, or nothing that ships."
        )

    say("Working out the branding...")
    brand = read_brand(store, refine=refine)

    say("Drawing a template in the store's colours...")
    template = work_dir / "template.pdf"
    boxes, sources = build_template(brand, domain, template, len(rows))

    product_file = write_product_file(rows, work_dir / "products.csv")

    say(f"Fetching {len(rows)} product photos...")
    prefetch_images(
        [row[0] for row in rows],
        progress=lambda done, total: say(f"Fetching product photos... {done} of {total}"),
    )

    export_catalog(
        excel_path=product_file,
        template_pdf=template,
        output_path=output_path,
        max_products=max_products,
        known_boxes=boxes,
        page_sources=sources,
        status_callback=status_callback,
    )

    return CatalogResult(
        output=output_path,
        brand=brand,
        product_count=len(rows),
        store_name=brand.name,
        notes=brand.notes,
    )
