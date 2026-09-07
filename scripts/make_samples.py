"""Generate every sample asset the Guide hands out.

Nothing here is derived from a client file. The product images are drawn from
scratch, the product rows are invented, and the template PDF is built from the
exporter's own reference grid constants -- so the sample template is correct by
construction rather than by copying one that happens to work.

    python scripts/make_samples.py

Writes into web/static/samples/. Safe to re-run; it overwrites in place.
"""

from __future__ import annotations

import csv
import subprocess
import sys
import tempfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont
from reportlab.lib.colors import HexColor
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

from catalog_exporter import (  # noqa: E402  (path setup has to come first)
    P6_FIRST_PAGE_ROWS,
    P6_LAST_PAGE_ROWS,
    P6_MIDDLE_PAGE_ROWS,
    REFERENCE_COLUMNS,
    REFERENCE_PAGE1_ROWS,
    REFERENCE_PAGE2_ROWS,
    REFERENCE_PAGE_SIZE,
    classify_page_layout,
    detect_grid_rows,
    render_pdf_page,
    try_font,
)

SAMPLES = REPO / "web" / "client_static" / "samples"
PRODUCTS = SAMPLES / "products"

# The sample template is A4 because that is what someone building their own
# template will reach for. The exporter scales its reference grid to whatever
# page size it is given, so the boxes land exactly where they are drawn here.
PAGE_W, PAGE_H = A4
GRID_COLOR = "#8A9199"
# The exporter finds a product box by sampling for ink ~2px inside each edge of
# where the box should be, on a 72dpi render. A hairline centred on the box
# outline half-misses that; the stroke is therefore drawn fully *inside* the
# outline, so the box keeps its exact outer size and the ink lands where the
# detector looks. Thinner than ~2.5pt here and the sample template stops being
# recognised as a template at all.
GRID_WIDTH = 3.2
ACCENT = "#E31B23"

# Public URL the sample product file points at. The images are served by this
# app itself, so the sample works end-to-end without depending on anyone's CDN.
# The client tool's samples moved out of the public static folder on
# 2026-09-07 -- they imitate that client's catalog design and must not be
# reachable from stuffs.bid. They are served on the client host only.
SAMPLE_IMAGE_BASE = "https://excelpdf.duckdns.org/client-assets/samples/products"

# Invented catalog: 14 rows, so the output spills onto a second page and shows
# both template page styles being used.
# Titles are kept short on purpose. A card is a small box, and the promotion
# style gives roughly half of it to the price panel -- real catalog copy is
# terse for the same reason.
SAMPLE_PRODUCTS = [
    # (title, por, case, pack, price, sale_price, is_new, shape, colour)
    ("Alpine Spring Water", 32, 24, "500ml", "10.80", "0.65", False, "bottle", "#3E8FD0"),
    ("Redbrook Cola", 28, 24, "330ml", "9.60", "0.59", True, "can", "#C0392B"),
    ("Harvest Oat Cereal", 41, 12, "450g", "14.40", "1.99", False, "box", "#D68910"),
    ("Meadow Gold Butter", 24, 20, "250g", "22.00", "1.45", False, "tub", "#F1C40F"),
    ("Bramley Pasta Sauce", 36, 12, "500g", "11.40", "1.29", True, "jar", "#A93226"),
    ("Ridgeway Crisps", 44, 24, "150g", "16.80", "0.99", False, "bag", "#2E86AB"),
    ("Orchard Apple Juice", 30, 12, "1L", "13.20", "1.49", False, "carton", "#7DAF3F"),
    ("Northwind Coffee", 38, 8, "227g", "19.60", "3.49", True, "bag", "#6E4B2A"),
    ("Cloudbank Yoghurt", 26, 12, "450g", "12.00", "1.35", False, "tub", "#EAEDED"),
    ("Baytown Ketchup", 33, 12, "570g", "13.80", "1.59", False, "bottle", "#C0392B"),
    ("Oakfield Biscuits", 39, 24, "400g", "15.60", "0.85", False, "box", "#B9770E"),
    ("Summit Energy", 47, 24, "250ml", "17.40", "0.95", True, "can", "#27AE60"),
    ("Fernvale Olive Oil", 29, 6, "500ml", "21.00", "4.29", False, "bottle", "#5D6D1E"),
    ("Redstone Chilli", 52, 12, "45g", "10.20", "1.19", False, "jar", "#922B21"),
]

# Header row the Guide documents. These are the friendly aliases -- the long
# Shopify metafield names still work, they are just not what a new user has.
HEADERS = [
    "Image Src",
    "Title",
    "POR",
    "Created At",
    "Case Size",
    "Pack Size",
    "Price",
    "Sale Price",
]


def image_filename(index: int) -> str:
    return f"product-{index + 1:02d}.png"


# --------------------------------------------------------------------------
# Product images
# --------------------------------------------------------------------------

def draw_product_image(shape: str, colour: str, label: str, path: Path) -> None:
    """A flat, obviously-illustrative product on white.

    White background is deliberate: the exporter trims white before placing the
    image, which is exactly what a real product cut-out gets.
    """
    width, height = 620, 820
    image = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(image)
    fill = HexColor(colour)
    body = (int(fill.red * 255), int(fill.green * 255), int(fill.blue * 255))
    shade = tuple(max(0, channel - 38) for channel in body)
    cx = width // 2

    if shape == "bottle":
        draw.rounded_rectangle((cx - 62, 90, cx + 62, 210), radius=18, fill=shade)
        draw.polygon([(cx - 62, 200), (cx + 62, 200), (cx + 150, 330), (cx - 150, 330)], fill=body)
        draw.rounded_rectangle((cx - 150, 320, cx + 150, 730), radius=40, fill=body)
    elif shape == "can":
        draw.rounded_rectangle((cx - 145, 150, cx + 145, 690), radius=54, fill=body)
        draw.ellipse((cx - 145, 120, cx + 145, 190), fill=shade)
    elif shape == "box":
        draw.polygon([(cx - 170, 210), (cx + 120, 150), (cx + 190, 200), (cx - 100, 262)], fill=shade)
        draw.rectangle((cx - 170, 250, cx + 100, 720), fill=body)
        draw.polygon([(cx + 100, 250), (cx + 190, 195), (cx + 190, 665), (cx + 100, 720)], fill=shade)
    elif shape == "tub":
        draw.polygon([(cx - 165, 300), (cx + 165, 300), (cx + 130, 700), (cx - 130, 700)], fill=body)
        draw.rounded_rectangle((cx - 180, 255, cx + 180, 315), radius=22, fill=shade)
    elif shape == "jar":
        draw.rounded_rectangle((cx - 105, 175, cx + 105, 250), radius=14, fill=shade)
        draw.rounded_rectangle((cx - 155, 240, cx + 155, 700), radius=46, fill=body)
    elif shape == "carton":
        draw.polygon([(cx - 130, 200), (cx, 130), (cx + 130, 200), (cx, 265)], fill=shade)
        draw.rectangle((cx - 130, 200, cx + 130, 715), fill=body)
    else:  # bag
        draw.polygon(
            [(cx - 190, 250), (cx + 190, 250), (cx + 155, 720), (cx - 155, 720)],
            fill=body,
        )
        draw.polygon([(cx - 190, 250), (cx - 120, 175), (cx + 120, 175), (cx + 190, 250)], fill=shade)

    # A white label band with the product's initials -- enough to read as a
    # product, not enough to pretend it is a real brand.
    draw.rounded_rectangle((cx - 128, 430, cx + 128, 560), radius=16, fill="white")
    initials = "".join(word[0] for word in label.split()[:2]).upper()
    font = try_font(74, bold=True)
    box = draw.textbbox((0, 0), initials, font=font)
    draw.text(
        (cx - (box[2] - box[0]) / 2, 495 - (box[3] - box[1]) / 2 - box[1]),
        initials,
        font=font,
        fill=(60, 60, 60),
    )

    image.save(path, "PNG", optimize=True)


# --------------------------------------------------------------------------
# Product data files
# --------------------------------------------------------------------------

def sample_rows() -> list[list[str]]:
    rows = []
    for index, (title, por, case, pack, price, sale, is_new, _shape, _colour) in enumerate(SAMPLE_PRODUCTS):
        rows.append(
            [
                f"{SAMPLE_IMAGE_BASE}/{image_filename(index)}",
                title,
                por,
                "NEW" if is_new else "",
                case,
                pack,
                price,
                sale,
            ]
        )
    return rows


def write_xlsx(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Products"
    sheet.append(HEADERS)

    header_fill = PatternFill("solid", fgColor="1F2933")
    for column in range(1, len(HEADERS) + 1):
        cell = sheet.cell(row=1, column=column)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(vertical="center")

    for row in sample_rows():
        sheet.append(row)

    for column, width in enumerate([58, 34, 8, 12, 11, 11, 10, 11], start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    sheet.freeze_panes = "A2"

    workbook.save(path)


def write_csv(path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(HEADERS)
        writer.writerows(sample_rows())


# --------------------------------------------------------------------------
# Template PDF
# --------------------------------------------------------------------------

def scaled_rect(left: int, top: int, right: int, bottom: int) -> tuple[float, float, float, float]:
    """Reference-grid box -> reportlab rect (origin bottom-left)."""
    sx = PAGE_W / REFERENCE_PAGE_SIZE[0]
    sy = PAGE_H / REFERENCE_PAGE_SIZE[1]
    x = left * sx
    width = (right - left) * sx
    height = (bottom - top) * sy
    y = PAGE_H - (bottom * sy)
    return x, y, width, height


def draw_grid(page: pdf_canvas.Canvas, rows: list[tuple[int, int]]) -> None:
    page.setStrokeColor(HexColor(GRID_COLOR))
    page.setLineWidth(GRID_WIDTH)
    half = GRID_WIDTH / 2
    for top, bottom in rows:
        for left, right in REFERENCE_COLUMNS:
            x, y, width, height = scaled_rect(left, top, right, bottom)
            page.rect(x + half, y + half, width - GRID_WIDTH, height - GRID_WIDTH, stroke=1, fill=0)


def draw_page_one_header(page: pdf_canvas.Canvas) -> None:
    """Branding band for page 1.

    Kept clear of the reference grid's top row (y 200-783) on purpose: the
    exporter tells a cover page from a middle page by checking whether that top
    row has boxes in it, and a busy header there reads as boxes.
    """
    sy = PAGE_H / REFERENCE_PAGE_SIZE[1]
    page.setFillColor(HexColor(ACCENT))
    page.rect(0, PAGE_H - (150 * sy), PAGE_W, 150 * sy, stroke=0, fill=1)

    page.setFillColor(HexColor("#FFFFFF"))
    page.setFont("Helvetica-Bold", 15)
    page.drawString(34, PAGE_H - (105 * sy), "YOUR BRAND HERE")

    page.setFillColor(HexColor("#1F2933"))
    page.setFont("Helvetica-Bold", 26)
    page.drawString(34, PAGE_H - (470 * sy), "Catalog title goes here")
    page.setFillColor(HexColor("#6B7280"))
    page.setFont("Helvetica", 11)
    page.drawString(34, PAGE_H - (560 * sy), "Replace this band with your own artwork. Leave the boxes where they are.")


def draw_footer(page: pdf_canvas.Canvas, text: str) -> None:
    page.setFillColor(HexColor("#9AA0A6"))
    page.setFont("Helvetica", 8)
    page.drawCentredString(PAGE_W / 2, 16, text)


def write_template_pdf(path: Path) -> None:
    page = pdf_canvas.Canvas(str(path), pagesize=A4)
    page.setTitle("Catalog template (sample)")

    # Page 1 -- cover style: header band on top, four rows of boxes below.
    draw_page_one_header(page)
    draw_grid(page, REFERENCE_PAGE1_ROWS)
    draw_footer(page, "Sample catalog template - page 1 (cover style)")
    page.showPage()

    # Page 2 -- middle style: five rows, grid starts at the top of the page.
    draw_grid(page, REFERENCE_PAGE2_ROWS)
    draw_footer(page, "Sample catalog template - page 2 (middle style)")
    page.showPage()

    page.save()


def write_single_page_template_pdf(path: Path) -> None:
    """The simplest template there is: one page, repeated for the whole catalog.

    A one-page template is not matched against the cover/middle shapes at all —
    the exporter measures which rows of boxes are actually on it and uses that
    same row set for every page. So this one is drawn dense: a slim brand bar
    that can survive being repeated, and the full five rows beneath it.
    """
    sy = PAGE_H / REFERENCE_PAGE_SIZE[1]
    page = pdf_canvas.Canvas(str(path), pagesize=A4)
    page.setTitle("Catalog template - one page (sample)")

    # Anything here has to look right on *every* page, so it is a bar rather
    # than a cover treatment, and it stays clear of the first row at y=200.
    page.setFillColor(HexColor(ACCENT))
    page.rect(0, PAGE_H - (150 * sy), PAGE_W, 150 * sy, stroke=0, fill=1)
    page.setFillColor(HexColor("#FFFFFF"))
    page.setFont("Helvetica-Bold", 13)
    page.drawString(34, PAGE_H - (105 * sy), "YOUR BRAND HERE")

    draw_grid(page, REFERENCE_PAGE2_ROWS)
    draw_footer(page, "Sample one-page template - repeated for every page")
    page.showPage()
    page.save()


def write_promotion_template_pdf(path: Path) -> None:
    """A second template, shaped for the promotion card style.

    The promotion card carries a price panel beside the photo, so it is taller:
    the exporter gives it one row fewer per page, and a *different* row set for
    each of the three page roles (`P6_FIRST/MIDDLE/LAST_PAGE_ROWS`). Those three
    sets do not overlap, which is why this template is three pages and the plain
    one is two -- a 2-page promotion template would have the exporter drawing
    last-page products into rows the cover page has no boxes in.

    Page 2 is the awkward one. A middle page is *recognised* by ink sitting in
    the top row band, but the promotion layout never places a product there --
    so that band carries a section banner instead of boxes. It reads as a middle
    page, and the banner survives into the output because the exporter only
    whitens the band its own slots span.
    """
    sy = PAGE_H / REFERENCE_PAGE_SIZE[1]
    page = pdf_canvas.Canvas(str(path), pagesize=A4)
    page.setTitle("Catalog template - promotion style (sample)")

    # Page 1 -- cover. Boxes sit one row lower than in the plain template.
    draw_page_one_header(page)
    page.setFillColor(HexColor("#6B7280"))
    page.setFont("Helvetica", 10)
    page.drawString(34, PAGE_H - (1150 * sy), "Room for a hero image, an intro, or nothing at all.")
    draw_grid(page, P6_FIRST_PAGE_ROWS)
    draw_footer(page, "Sample promotion template - page 1 (cover)")
    page.showPage()

    # Page 2 -- the repeating middle page, with the banner standing in for the
    # row of boxes the exporter looks for but never fills.
    banner_top, banner_bottom = 200, 783
    page.setFillColor(HexColor(ACCENT))
    page.rect(0, PAGE_H - (banner_bottom * sy), PAGE_W, (banner_bottom - banner_top) * sy, stroke=0, fill=1)
    page.setFillColor(HexColor("#FFFFFF"))
    page.setFont("Helvetica-Bold", 22)
    page.drawString(34, PAGE_H - (560 * sy), "SECTION BANNER")
    draw_grid(page, P6_MIDDLE_PAGE_ROWS)
    draw_footer(page, "Sample promotion template - page 2 (middle, repeats)")
    page.showPage()

    # Page 3 -- back page: one row fewer again, leaving room for a contact band.
    draw_grid(page, P6_LAST_PAGE_ROWS)
    page.setFillColor(HexColor("#1F2933"))
    page.rect(0, PAGE_H - (3203 * sy), PAGE_W, (3203 - 2700) * sy, stroke=0, fill=1)
    page.setFillColor(HexColor("#FFFFFF"))
    page.setFont("Helvetica-Bold", 15)
    page.drawString(34, PAGE_H - (2900 * sy), "Back page — contact details, terms, whatever you need")
    draw_footer(page, "Sample promotion template - page 3 (back page)")
    page.showPage()

    page.save()


def verify_single_page_template(path: Path, expected_rows: int) -> None:
    """A one-page template is judged by the rows the exporter measures on it."""
    page = render_pdf_page(path, 1, dpi=72)
    try:
        rows = detect_grid_rows(page)
    finally:
        page.close()
    status = "ok" if len(rows) == expected_rows else "MISMATCH"
    print(f"  one-page: {len(rows)} rows detected (want {expected_rows}) tops={[r[0] for r in rows]}  [{status}]")
    if len(rows) != expected_rows:
        raise SystemExit(
            f"One-page sample template detects {len(rows)} rows, not {expected_rows}. "
            "Check the brand bar is not overlapping the first row."
        )


def verify_template(path: Path, expected: dict[int, str] | None = None) -> None:
    """Run the exporter's own page classifier over the sample we just wrote."""
    expected = expected or {1: "first", 2: "toprow"}
    for page_number, want in expected.items():
        page = render_pdf_page(path, page_number, dpi=72)
        try:
            result = classify_page_layout(page)
        finally:
            page.close()
        status = "ok" if result.kind == want else "MISMATCH"
        print(
            f"  page {page_number}: kind={result.kind} (want {want}) "
            f"first={result.first_score:.2f} middle={result.middle_score:.2f}  [{status}]"
        )
        if result.kind != want:
            raise SystemExit(
                f"Sample template page {page_number} classifies as {result.kind!r}, not {want!r}. "
                "Adjust the grid or the header band and re-run."
            )


# --------------------------------------------------------------------------
# Preview images for the Guide
# --------------------------------------------------------------------------

def write_pdf_previews(pdf_path: Path, stem: str, pages: int, dpi: int = 74) -> None:
    for page_number in range(1, pages + 1):
        page = render_pdf_page(pdf_path, page_number, dpi=dpi)
        try:
            bordered = Image.new("RGB", (page.width + 2, page.height + 2), "#D5D8DC")
            bordered.paste(page, (1, 1))
            bordered.save(SAMPLES / f"{stem}-page{page_number}.png", "PNG", optimize=True)
        finally:
            page.close()


def write_output_preview(template: str, form: str, stem: str) -> None:
    """Run the sample product file through the real exporter and photograph it.

    This doubles as an end-to-end check: if a sample template or the sample
    columns are wrong, this step fails rather than shipping a broken guide.
    Image paths are local here so the generator does not depend on the site
    being up; the shipped sample file uses public URLs, which is what a real
    product file has.
    """
    import catalog_exporter

    with tempfile.TemporaryDirectory() as work_dir:
        work = Path(work_dir)
        excel_path = work / "sample.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(HEADERS)
        for index, row in enumerate(sample_rows()):
            local = list(row)
            local[0] = str(PRODUCTS / image_filename(index))
            sheet.append(local)
        workbook.save(excel_path)

        output = work / "output.pdf"
        catalog_exporter.export_catalog(
            excel_path=excel_path,
            template_pdf=SAMPLES / template,
            output_path=output,
            quality="normal",
            template_form=form,
        )
        page_count = catalog_exporter.get_pdf_page_count(output)
        print(f"  {form}: {page_count} page(s)")
        write_pdf_previews(output, stem, min(page_count, 2))


def main() -> None:
    SAMPLES.mkdir(parents=True, exist_ok=True)
    PRODUCTS.mkdir(parents=True, exist_ok=True)

    print("product images...")
    for index, (title, _por, _case, _pack, _price, _sale, _new, shape, colour) in enumerate(SAMPLE_PRODUCTS):
        draw_product_image(shape, colour, title, PRODUCTS / image_filename(index))
    print(f"  {len(SAMPLE_PRODUCTS)} images")

    print("product data files...")
    write_xlsx(SAMPLES / "catalog-products-template.xlsx")
    write_csv(SAMPLES / "catalog-products-template.csv")

    print("template pdfs...")
    write_single_page_template_pdf(SAMPLES / "catalog-template-onepage.pdf")
    verify_single_page_template(SAMPLES / "catalog-template-onepage.pdf", expected_rows=5)
    write_template_pdf(SAMPLES / "catalog-template.pdf")
    verify_template(SAMPLES / "catalog-template.pdf")
    write_promotion_template_pdf(SAMPLES / "catalog-template-promotion.pdf")
    verify_template(
        SAMPLES / "catalog-template-promotion.pdf",
        {1: "first", 2: "toprow", 3: "first"},
    )

    print("previews...")
    write_pdf_previews(SAMPLES / "catalog-template-onepage.pdf", "onepage-template-preview", 1)
    write_pdf_previews(SAMPLES / "catalog-template.pdf", "template-preview", 2)
    write_pdf_previews(SAMPLES / "catalog-template-promotion.pdf", "promo-template-preview", 3)

    print("end-to-end sample exports...")
    write_output_preview("catalog-template-onepage.pdf", "por-title", "example-onepage")
    write_output_preview("catalog-template.pdf", "por-title", "example-output")
    write_output_preview("catalog-template-promotion.pdf", "promotion", "example-promo")

    print("done.")


if __name__ == "__main__":
    main()
