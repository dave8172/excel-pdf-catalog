# Excel To PDF Catalog Exporter

This project is a focused Python exporter for building catalog PDFs from:

- an input Excel file with product data
- a PDF template
- optional badge/label image assets

There is no GUI. You run `export_pdf.py` directly.

## Project Layout

- `export_pdf.py`: main file you edit and run
- `catalog_exporter.py`: export engine
- `templates/`: source Excel, template PDF, and label assets
- `exports/`: generated output PDFs

## Setup

Install Python dependencies:

```powershell
python -m pip install -r requirements.txt
```

Install Poppler and make sure `pdftoppm` and `pdfunite` are available on `PATH`.

## Run

Open `export_pdf.py` and edit the `User inputs` section:

```python
INPUT_EXCEL_FILE = TEMPLATES_DIR / "P6 Promotions_v3.xlsx"
INPUT_TEMPLATE_PDF = TEMPLATES_DIR / "Spring.pdf"
OUTPUT_PDF_FILE = EXPORTS_DIR / "catalog_normal.pdf"
EXPORT_QUALITY = "normal"
TEMPLATE_FORM = "auto"
```

Then run:

```powershell
python .\export_pdf.py
```

You can also override values from the command line:

```powershell
python .\export_pdf.py `
  --excel ".\templates\P6 Promotions_v3.xlsx" `
  --template ".\templates\Spring.pdf" `
  --output ".\exports\custom_output.pdf" `
  --quality normal `
  --template-form auto
```

## Saved Template Forms

Use `--template-form` when you want a specific saved layout instead of automatic detection.

- `por-title`: W15/Week 18 style. Uses only product image, title, POR, and the NEW marker from Excel. Product images are trimmed, centered horizontally, and sized to fill the card image area. The POR badge is placed at the top right using `por img.png`; the NEW badge is placed at the top left using `new img.png`.
- `promotion`: Promotion-style catalogs (Spring, P7 BBQ, and later campaigns). Uses price, sale price, POR panel, title, pack size, and horizontal/vertical image-position handling.
- `auto`: auto-detects from the Excel columns. If the sheet contains a `main price` or `retail sale price` column → `promotion` style. Otherwise → `por-title` style.

Example for the promotion form (P7 BBQ):

```powershell
python .\export_pdf.py `
  --excel ".\templates\P7 Promotions_v4.xlsx" `
  --template ".\templates\P7 BBQ.pdf" `
  --output ".\exports\P7 BBQ normal output.pdf" `
  --quality normal `
  --template-form promotion
```

Example for the Week 18-style form:

```powershell
python .\export_pdf.py `
  --excel ".\templates\Weekly Restocked & New Lines 26 - Week 18.xlsx" `
  --template ".\templates\Week 18.pdf" `
  --output ".\exports\Week 18 normal output.pdf" `
  --quality normal `
  --template-form por-title
```

## Optional Asset Inputs

`export_pdf.py` also exposes the badge image paths used by the renderer:

- `POR_BADGE_IMAGE_FILE`
- `NEW_BADGE_IMAGE_FILE`
- `P6_POR_PANEL_IMAGE_FILE`

If you want to use different label artwork for another catalog, change those paths before running the script.

## Excel Data

The exporter reads the required product fields from the Excel header row.

For the `por-title` form, the layout uses:

- image from `Image Src`
- title from `Title`
- POR from `Metafield: custom.tier_c_profit_on_return [number_decimal]`
- NEW badge when `Created At` contains `NEW`

For the `promotion` form, the layout uses:

- main price from column `E`
- POR from column `F`
- sale price from column `G`
- image orientation from column `J`

## Current Included Sources

The `templates/` folder currently keeps the active source files and shared assets:

- `P6 Promotions_v3.xlsx`
- `Spring.pdf`
- `ref pic p6.png`
- `por img p6.png`
- `por img.png`
- `new img.png`

## Notes

- `normal` export is the default and is usually the best choice for day-to-day runs.
- `high` export is available if you want a sharper template-preserving output.
- Output folders are created automatically if they do not exist.
