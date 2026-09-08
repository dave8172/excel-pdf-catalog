"""topdf's own catalog engine — the public product.

Deliberately separate from `catalog_exporter`, which is the client tool's
engine and is not touched by anything here. Nothing in this module knows about
profit-on-return badges, NEW flashes, promotion cards, quality modes, or the
client's fixed reference grid — all of those are that client's design and stay
on the client deployment.

What this one does instead:

* **Five columns, plain English.** Image URL and Name are required; Description,
  Price and Case Size are optional and simply omitted from the card when absent.
* **It measures the template.** Rather than assuming boxes sit at known
  coordinates, it finds the rectangles that are actually drawn on each template
  page and still empty inside. Any number of rows, any number of columns,
  anywhere on the page.
* **No options.** One card design, one quality. The only two inputs are the
  spreadsheet and the template PDF.

Shared with the client engine only where the code is generic and carries no
design: PDF page rendering, the SSRF-guarded image fetch, white-trimming, font
lookup and PDF assembly.
"""

from __future__ import annotations

import re
import shutil
import tempfile
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Callable

from PIL import Image, ImageChops, ImageDraw, ImageFont

from catalog_exporter import (
    NORMAL_EXPORT_RENDER_DPI,
    _iter_rows_from_source,
    get_pdf_page_count,
    load_product_image,
    merge_pdf_files,
    render_pdf_page,
    save_page_pdf,
    trim_product_image,
    try_font,
)

# --- the spreadsheet -------------------------------------------------------

COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "image": ("image url", "image", "image src", "image link", "photo", "photo url", "picture", "img"),
    "name": ("name", "product", "product name", "title", "item", "item name"),
    "description": ("description", "desc", "details", "detail", "product description", "subtitle"),
    "price": ("price", "unit price", "cost", "rrp", "amount"),
    "case_size": ("case size", "case", "pack size", "pack", "units per case", "case qty", "quantity"),
}
REQUIRED_COLUMNS = ("image", "name")
COLUMN_DISPLAY_NAMES = {
    "image": "Image URL",
    "name": "Name",
    "description": "Description",
    "price": "Price",
    "case_size": "Case Size",
}

CURRENCY_SYMBOLS = "£$€¥₹₽₩₪"
DEFAULT_CURRENCY_SYMBOL = "£"


@dataclass
class Product:
    image: str = ""
    name: str = ""
    description: str = ""
    price: str = ""
    case_size: str = ""


def normalize_header(value: object) -> str:
    """Fold a header cell down to letters, digits and single spaces.

    Looser than the client engine's version on purpose: `Image_URL`,
    `image-url` and `Image URL ` all have to land on the same key, because the
    person filling this sheet in has never seen the tool before.
    """
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").strip().lower()).strip()


def _cell(value: object) -> str:
    return "" if value in (None, "") else str(value).strip()


def format_price(value: object) -> str:
    """Print a price as typed, but tidied. A bare number gets the default symbol."""
    raw = _cell(value)
    if not raw:
        return ""
    has_symbol = raw[:1] in set(CURRENCY_SYMBOLS)
    symbol = raw[0] if has_symbol else DEFAULT_CURRENCY_SYMBOL
    body = (raw[1:] if has_symbol else raw).strip().replace(",", "")
    try:
        number = Decimal(body)
    except InvalidOperation:
        return raw
    return f"{symbol}{format(number.quantize(Decimal('0.01')), 'f')}"


def format_case_size(value: object) -> str:
    """`24` becomes `24 per case`; `24 x 500ml` is left alone."""
    raw = _cell(value)
    if not raw:
        return ""
    try:
        number = Decimal(raw)
    except InvalidOperation:
        return raw
    whole = number.quantize(Decimal("1")) if number == number.to_integral() else number.normalize()
    return f"{format(whole, 'f')} per case"


def resolve_columns(header_row: tuple[object, ...]) -> dict[str, int]:
    """Map our five field names onto column indexes in the uploaded sheet."""
    present = {normalize_header(value): index for index, value in enumerate(header_row) if normalize_header(value)}
    resolved: dict[str, int] = {}
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in present:
                resolved[field] = present[alias]
                break
    return resolved


def load_products(excel_path: Path, max_products: int | None = None) -> list[Product]:
    rows = _iter_rows_from_source(excel_path)
    try:
        header_row = next(rows)
    except StopIteration:
        raise ValueError("That product file is empty. Row 1 should be the header row.") from None

    columns = resolve_columns(tuple(header_row))
    missing = [COLUMN_DISPLAY_NAMES[field] for field in REQUIRED_COLUMNS if field not in columns]
    if missing:
        raise ValueError(
            f"The product file needs {' and '.join(missing)} in row 1. "
            f"Row 1 has to be the header row, and the headers it can use are: "
            f"{', '.join(COLUMN_DISPLAY_NAMES.values())}."
        )

    products: list[Product] = []
    for row in rows:
        if not any(_cell(cell) for cell in row):
            continue  # a blank line in the middle of the sheet is not a product

        def value(field: str) -> str:
            index = columns.get(field)
            if index is None or index >= len(row):
                return ""
            return _cell(row[index])

        name = value("name")
        image = value("image")
        if not name and not image:
            continue
        products.append(
            Product(
                image=image,
                name=name or "Untitled product",
                description=value("description"),
                price=format_price(value("price")),
                case_size=format_case_size(value("case_size")),
            )
        )
        if max_products is not None and len(products) > max_products:
            raise ValueError(
                f"That file has more than {max_products} products. "
                f"Split the list and run it twice."
            )

    if not products:
        raise ValueError("No products were found under the header row.")
    return products


# --- finding the boxes on a template page ----------------------------------
#
# The client engine looks for boxes at coordinates it already knows. This one
# cannot: a stranger's template is whatever they drew. So it reads the page.
#
# The method is a projection profile. A drawn box edge is a long unbroken run
# of ink along one row or column of pixels; body text and photographs are not.
# Collect the rows and columns that contain such a run, treat every adjacent
# pair as a candidate rectangle, and keep the ones whose four edges are inked
# and whose middle is still blank. The blank-middle half is what stops a solid
# header band or a dark footer strip from being filled with products.

INK_LUMA_MAX = 200          # a pixel at least this dark counts as drawn
INK_SATURATION_MIN = 70     # ...or this colourful, so brand-coloured rules count
LINE_MAX_THICKNESS = 6      # px; anything inked wider than this is a fill, not a rule
MIN_BOX_WIDTH_RATIO = 0.07
MIN_BOX_HEIGHT_RATIO = 0.045
MIN_EDGE_INK_RATIO = 0.55
MIN_INTERIOR_CLEAR_RATIO = 0.90
MAX_BOXES_PER_PAGE = 80


def _ink_mask(page: Image.Image) -> Image.Image:
    """A 1-bit-ish mask (0 or 255 in mode L) of everything drawn on the page."""
    rgb = page if page.mode == "RGB" else page.convert("RGB")
    dark = rgb.convert("L").point(lambda value: 255 if value < INK_LUMA_MAX else 0)
    saturated = rgb.convert("HSV").getchannel("S").point(
        lambda value: 255 if value > INK_SATURATION_MIN else 0
    )
    return ImageChops.lighter(dark, saturated)


def _thin_ink(mask: Image.Image, dx: int, dy: int) -> Image.Image:
    """Ink that does *not* continue `LINE_MAX_THICKNESS` px in both directions.

    A drawn rule is thin; a filled band is not. Without this the whole point of
    the projection profile collapses — a solid header band is an unbroken run
    of ink down every column beneath it, so every column reads as a border.
    Keeping only ink with white a few pixels to either side leaves the rules
    and the glyph strokes, and drops the fills.
    """
    before = ImageChops.offset(mask, dx, dy)
    after = ImageChops.offset(mask, -dx, -dy)
    return ImageChops.subtract(mask, ImageChops.darker(before, after))


def _lines_with_long_runs(data: bytes, span: int, count: int, min_run: int) -> list[int]:
    """Indexes of the rows of `data` holding an unbroken ink run of `min_run`.

    `data` is the mask's raw bytes, so a row is a plain slice and the run
    lengths come from one C-level split. Doing this per pixel in Python is
    seconds per page; this is milliseconds.
    """
    hits: list[int] = []
    for index in range(count):
        line = data[index * span : (index + 1) * span]
        if b"\xff" not in line:
            continue
        if max(len(part) for part in line.split(b"\x00")) >= min_run:
            hits.append(index)
    return hits


def _group_adjacent(indexes: list[int], tolerance: int = 2) -> list[int]:
    """Collapse each run of neighbouring line indexes to its centre."""
    if not indexes:
        return []
    centres: list[int] = []
    start = previous = indexes[0]
    for index in indexes[1:]:
        if index - previous <= tolerance:
            previous = index
            continue
        centres.append((start + previous) // 2)
        start = previous = index
    centres.append((start + previous) // 2)
    return centres


def _edge_ink_ratio(mask: Image.Image, box: tuple[int, int, int, int]) -> float:
    """The weakest of the four edges, as a fraction of samples that are inked."""
    left, top, right, bottom = box
    width, height = mask.size
    inset = max(1, round(min(right - left, bottom - top) * 0.02))

    def ratio(points: list[tuple[int, int]]) -> float:
        if not points:
            return 0.0
        hit = 0
        for x, y in points:
            px = min(max(x, 0), width - 1)
            py = min(max(y, 0), height - 1)
            # A border can be a pixel off from the detected centre line, so a
            # sample counts if any of its immediate neighbours is inked.
            if any(
                mask.getpixel((min(max(px + dx, 0), width - 1), min(max(py + dy, 0), height - 1)))
                for dx in (-2, -1, 0, 1, 2)
                for dy in (-2, -1, 0, 1, 2)
            ):
                hit += 1
        return hit / len(points)

    step_x = max(1, (right - left) // 16)
    step_y = max(1, (bottom - top) // 16)
    xs = list(range(left + inset, right - inset, step_x))
    ys = list(range(top + inset, bottom - inset, step_y))
    return min(
        ratio([(x, top) for x in xs]),
        ratio([(x, bottom) for x in xs]),
        ratio([(left, y) for y in ys]),
        ratio([(right, y) for y in ys]),
    )


def _interior_clear_ratio(mask: Image.Image, box: tuple[int, int, int, int]) -> float:
    """How much of the middle of the rectangle is blank."""
    left, top, right, bottom = box
    inset_x = max(1, round((right - left) * 0.14))
    inset_y = max(1, round((bottom - top) * 0.14))
    region = mask.crop((left + inset_x, top + inset_y, right - inset_x, bottom - inset_y))
    total = region.width * region.height
    if total <= 0:
        return 0.0
    return region.histogram()[0] / total  # value 0 is "no ink"


def find_product_boxes(page: Image.Image) -> list[tuple[int, int, int, int]]:
    """Every empty, drawn rectangle on the page, in reading order."""
    width, height = page.size
    mask = _ink_mask(page)

    # A horizontal rule is ink that is thin *vertically*, and vice versa.
    rows = _group_adjacent(
        _lines_with_long_runs(
            _thin_ink(mask, 0, LINE_MAX_THICKNESS).tobytes(),
            width,
            height,
            round(width * MIN_BOX_WIDTH_RATIO),
        )
    )
    # TRANSPOSE, not ROTATE_90: reflecting along the main diagonal puts column
    # x of the page at row x of the result, so the indexes that come back are
    # already x coordinates and need no un-flipping.
    columns = _group_adjacent(
        _lines_with_long_runs(
            _thin_ink(mask, LINE_MAX_THICKNESS, 0).transpose(Image.Transpose.TRANSPOSE).tobytes(),
            height,
            width,
            round(height * MIN_BOX_HEIGHT_RATIO),
        )
    )

    min_width = round(width * MIN_BOX_WIDTH_RATIO)
    min_height = round(height * MIN_BOX_HEIGHT_RATIO)

    boxes: list[tuple[int, int, int, int]] = []
    for top, bottom in zip(rows, rows[1:]):
        if bottom - top < min_height:
            continue
        for left, right in zip(columns, columns[1:]):
            if right - left < min_width:
                continue
            box = (left, top, right, bottom)
            if _edge_ink_ratio(mask, box) < MIN_EDGE_INK_RATIO:
                continue
            if _interior_clear_ratio(mask, box) < MIN_INTERIOR_CLEAR_RATIO:
                continue
            boxes.append(box)
            if len(boxes) >= MAX_BOXES_PER_PAGE:
                return boxes
    return boxes


# --- drawing one product ---------------------------------------------------

NAME_COLOR = "#1A1A1A"
DESCRIPTION_COLOR = "#5F6368"
PRICE_COLOR = "#1A1A1A"
PLACEHOLDER_COLOR = "#C9CCD1"


def _wrap(draw: ImageDraw.ImageDraw, text: str, font: ImageFont.ImageFont, max_width: float) -> list[str]:
    lines: list[str] = []
    current = ""
    for word in text.split():
        trial = f"{current} {word}".strip()
        if not current or draw.textlength(trial, font=font) <= max_width:
            current = trial
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    return lines


def _fit_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    max_width: float,
    max_lines: int,
    start_size: int,
    min_size: int,
    bold: bool,
) -> tuple[ImageFont.ImageFont, list[str]]:
    """Largest size at or below `start_size` that fits in `max_lines`."""
    size = max(start_size, min_size)
    while size > min_size:
        font = try_font(size, bold=bold)
        lines = _wrap(draw, text, font, max_width)
        if len(lines) <= max_lines:
            return font, lines
        size -= 1
    font = try_font(min_size, bold=bold)
    lines = _wrap(draw, text, font, max_width)
    if len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1].rstrip(" ,.") + "…"
    return font, lines


def _line_height(font: ImageFont.ImageFont) -> int:
    ascent, descent = font.getmetrics() if hasattr(font, "getmetrics") else (10, 2)
    return ascent + descent


def draw_product(page: Image.Image, product: Product, image: Image.Image | None, box: tuple[int, int, int, int]) -> None:
    """One card: photo on top, then name, description and the price line."""
    left, top, right, bottom = box
    width = right - left
    height = bottom - top
    pad = max(4, round(min(width, height) * 0.07))
    inner_left = left + pad
    inner_right = right - pad
    inner_top = top + pad
    inner_bottom = bottom - pad
    text_width = inner_right - inner_left

    draw = ImageDraw.Draw(page)

    name_size = min(max(round(height * 0.085), 11), 44)
    name_font, name_lines = _fit_text(
        draw, product.name, text_width, 2, name_size, max(9, round(name_size * 0.62)), bold=True
    )

    description_lines: list[str] = []
    description_font = None
    if product.description:
        description_font, description_lines = _fit_text(
            draw,
            product.description,
            text_width,
            2,
            max(9, round(name_size * 0.76)),
            max(8, round(name_size * 0.55)),
            bold=False,
        )

    meta = " · ".join(part for part in (product.price, product.case_size) if part)
    meta_lines: list[str] = []
    meta_font = None
    if meta:
        meta_font, meta_lines = _fit_text(
            draw, meta, text_width, 1, max(9, round(name_size * 0.88)), max(8, round(name_size * 0.55)), bold=True
        )

    gap = max(2, round(height * 0.012))
    blocks: list[tuple[ImageFont.ImageFont, list[str], str]] = [(name_font, name_lines, NAME_COLOR)]
    if description_lines and description_font is not None:
        blocks.append((description_font, description_lines, DESCRIPTION_COLOR))
    if meta_lines and meta_font is not None:
        blocks.append((meta_font, meta_lines, PRICE_COLOR))

    text_height = sum(len(lines) * _line_height(font) for font, lines, _ in blocks) + gap * (len(blocks) - 1)

    # The photo takes whatever the text does not, and never less than a third
    # of the card -- a long name should shrink the type, not swallow the image.
    image_bottom = inner_bottom - text_height - max(gap * 2, round(height * 0.03))
    image_area = (inner_left, inner_top, inner_right, max(inner_top + round(height * 0.30), image_bottom))

    _draw_photo(page, draw, image, image_area)

    y = inner_bottom - text_height
    for font, lines, color in blocks:
        for line in lines:
            draw.text(((inner_left + inner_right) / 2, y), line, font=font, fill=color, anchor="ma")
            y += _line_height(font)
        y += gap


def _draw_photo(
    page: Image.Image,
    draw: ImageDraw.ImageDraw,
    image: Image.Image | None,
    area: tuple[int, int, int, int],
) -> None:
    left, top, right, bottom = area
    width = right - left
    height = bottom - top
    if width <= 2 or height <= 2:
        return

    if image is None:
        # No photo is a normal outcome — a broken link, or a row that never had
        # one. Say so quietly rather than leaving a hole the reader has to guess at.
        font = try_font(max(8, round(min(width, height) * 0.10)))
        draw.text(
            ((left + right) / 2, (top + bottom) / 2),
            "no image",
            font=font,
            fill=PLACEHOLDER_COLOR,
            anchor="mm",
        )
        return

    trimmed = trim_product_image(image)
    if trimmed.width < 1 or trimmed.height < 1:
        return
    scale = min(width / trimmed.width, height / trimmed.height)
    size = (max(1, round(trimmed.width * scale)), max(1, round(trimmed.height * scale)))
    resized = trimmed.resize(size, Image.LANCZOS)
    position = (round(left + (width - size[0]) / 2), round(top + (height - size[1]) / 2))
    page.paste(resized, position, resized)


# --- the export ------------------------------------------------------------


def export_catalog(
    *,
    excel_path: Path,
    template_pdf: Path,
    output_path: Path,
    max_products: int | None = None,
    known_boxes: list[list[tuple[int, int, int, int]]] | None = None,
    page_sources: list[int] | None = None,
    status_callback: Callable[[str], None] | None = None,
    progress_callback: Callable[[int, int], None] | None = None,
) -> Path:
    """Fill the template's boxes with the spreadsheet's products.

    Template pages are used in the order they appear. If products are still
    left after the last one, the last page that *has* boxes repeats until they
    run out — which makes the common shape (a branded cover, then a denser page
    that repeats) work without the tool needing to be told about it.

    `known_boxes` skips detection for a template this program drew itself, and
    exists because measuring your own drawing is a category error. Detection is
    how the tool copes with a *stranger's* PDF; when `shopify_catalog` generates
    the template it already knows where every box is, to the point. Re-deriving
    those coordinates from pixels only adds ways to be wrong — a brand logo
    whose left edge happened to land on the grid's left column silently cost a
    whole column of products, and no tuning of the detector makes that class of
    accident impossible. Uploaded templates are unaffected and still measure.

    `page_sources` goes with it: `known_boxes` is then indexed by *output* page
    and `page_sources[i]` names the template page output page `i` is drawn on.
    That lets a generated template hold two physical designs and still produce
    eighty-four pages, which is the difference between 13MB of rendered
    template and 549MB of it.
    """

    def say(message: str) -> None:
        if status_callback:
            status_callback(message)

    say("Reading the product file...")
    products = load_products(excel_path, max_products=max_products)

    page_count = get_pdf_page_count(template_pdf)
    say(f"Reading the template ({page_count} page{'s' if page_count != 1 else ''})...")

    # A rendered A4 page at 150dpi is 6.5MB, and this used to hold one per
    # template page for the whole export. That is fine for the handful of pages
    # somebody uploads and fatal for a generated one: an 84-page template (1000
    # products) is 549MB of identical pictures on a box whose MemoryHigh is
    # 550MB. Renders are now cached by template page, so the cost is the number
    # of *distinct designs* — two, for every generated template — rather than
    # the length of the catalog.
    rendered_cache: dict[int, Image.Image] = {}

    def template_page(index: int) -> Image.Image:
        if index not in rendered_cache:
            rendered_cache[index] = render_pdf_page(
                template_pdf, index + 1, dpi=NORMAL_EXPORT_RENDER_DPI
            )
        return rendered_cache[index]

    try:
        if known_boxes is not None:
            # A template we drew: `known_boxes` is per *output* page and
            # `page_sources` says which design each of those uses, so there is
            # nothing to detect and nothing to plan.
            plan = list(page_sources) if page_sources is not None else list(range(len(known_boxes)))
            boxes_for_output = known_boxes
        else:
            page_boxes = [
                find_product_boxes(template_page(index)) for index in range(page_count)
            ]
            if not any(page_boxes):
                raise ValueError(
                    "No product boxes were found in that template. The tool fills rectangles that are "
                    "drawn as visible outlines and left empty inside — check the boxes are actually "
                    "drawn (not just white space), and that nothing is sitting inside them."
                )

            # The page that repeats is the last one with boxes on it — not simply
            # the last one, or a template ending in a terms page would repeat that
            # forever and never place another product.
            repeating = next(index for index in range(page_count - 1, -1, -1) if page_boxes[index])

            # Which template page each output page is built from.
            plan = []
            remaining = len(products)
            for index in range(page_count):
                plan.append(index)
                remaining -= len(page_boxes[index])
                if remaining <= 0:
                    break
            while remaining > 0:
                plan.append(repeating)
                remaining -= len(page_boxes[repeating])
            boxes_for_output = [page_boxes[index] for index in plan]

        say(f"Placing {len(products)} products across {len(plan)} page{'s' if len(plan) != 1 else ''}...")

        work_dir = Path(tempfile.mkdtemp(prefix="topdf_"))
        try:
            rendered_paths: list[Path] = []
            cursor = 0
            for position, template_index in enumerate(plan, start=1):
                page = template_page(template_index).copy()
                for box in boxes_for_output[position - 1]:
                    if cursor >= len(products):
                        break
                    product = products[cursor]
                    cursor += 1
                    image = load_product_image(excel_path, product.image)
                    draw_product(page, product, image, box)
                    if image is not None:
                        image.close()

                page_path = work_dir / f"page-{position:03d}.pdf"
                save_page_pdf(page, page_path)
                page.close()
                rendered_paths.append(page_path)
                if progress_callback:
                    progress_callback(position, len(plan))

            say("Assembling the PDF...")
            merge_pdf_files(rendered_paths, output_path)
        finally:
            shutil.rmtree(work_dir, ignore_errors=True)
    finally:
        for page in rendered_cache.values():
            page.close()

    return output_path
