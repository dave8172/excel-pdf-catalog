from __future__ import annotations

import argparse
from pathlib import Path

from catalog_exporter import TEMPLATE_FORM_VARIANTS, configure_asset_paths, export_catalog


PROJECT_DIR = Path(__file__).resolve().parent
TEMPLATES_DIR = PROJECT_DIR / "templates"
EXPORTS_DIR = PROJECT_DIR / "exports"

# User inputs
# Edit these values directly when you want to run a different job.
INPUT_EXCEL_FILE = TEMPLATES_DIR / "P6 Promotions_v3.xlsx"
INPUT_TEMPLATE_PDF = TEMPLATES_DIR / "Spring.pdf"
OUTPUT_PDF_FILE = EXPORTS_DIR / "catalog_normal.pdf"
EXPORT_QUALITY = "normal"
TEMPLATE_FORM = "auto"

# Optional label assets used by the renderer.
# Set any of these to None if you do not want to override the built-in default.
POR_BADGE_IMAGE_FILE = TEMPLATES_DIR / "por img.png"
NEW_BADGE_IMAGE_FILE = TEMPLATES_DIR / "new img.png"
P6_POR_PANEL_IMAGE_FILE = TEMPLATES_DIR / "por img p6.png"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a catalog PDF from an input Excel file and template PDF.")
    parser.add_argument("--excel", type=Path, help="Override the Excel source file.")
    parser.add_argument("--template", type=Path, help="Override the template PDF.")
    parser.add_argument("--output", type=Path, help="Override the output PDF path.")
    parser.add_argument("--quality", choices=["normal", "high"], help="Override the export quality.")
    parser.add_argument(
        "--template-form",
        choices=["auto", *TEMPLATE_FORM_VARIANTS.keys()],
        help="Use a named saved template form, such as 'por-title'.",
    )
    parser.add_argument("--por-badge", type=Path, help="Override the standard POR badge image.")
    parser.add_argument("--new-badge", type=Path, help="Override the NEW badge image.")
    parser.add_argument("--p6-por-panel", type=Path, help="Override the P6 POR panel image.")
    return parser.parse_args()


def resolve_path(path: Path | None, *, base_dir: Path) -> Path | None:
    if path is None:
        return None
    return path if path.is_absolute() else (base_dir / path).resolve()


def require_file(path: Path | None, label: str) -> Path:
    if path is None:
        raise FileNotFoundError(f"{label} was not provided.")
    if not path.exists():
        raise FileNotFoundError(f"{label} not found: {path}")
    return path


def print_status(message: str) -> None:
    print(message)


def main() -> None:
    args = parse_args()

    excel_path = require_file(
        resolve_path(args.excel, base_dir=Path.cwd()) if args.excel else resolve_path(INPUT_EXCEL_FILE, base_dir=PROJECT_DIR),
        "Excel file",
    )
    template_path = require_file(
        resolve_path(args.template, base_dir=Path.cwd()) if args.template else resolve_path(INPUT_TEMPLATE_PDF, base_dir=PROJECT_DIR),
        "Template PDF",
    )
    output_path = (
        resolve_path(args.output, base_dir=Path.cwd()) if args.output else resolve_path(OUTPUT_PDF_FILE, base_dir=PROJECT_DIR)
    )
    quality = args.quality or EXPORT_QUALITY
    template_form = args.template_form or TEMPLATE_FORM

    por_badge_path = (
        resolve_path(args.por_badge, base_dir=Path.cwd())
        if args.por_badge
        else resolve_path(POR_BADGE_IMAGE_FILE, base_dir=PROJECT_DIR)
    )
    new_badge_path = (
        resolve_path(args.new_badge, base_dir=Path.cwd())
        if args.new_badge
        else resolve_path(NEW_BADGE_IMAGE_FILE, base_dir=PROJECT_DIR)
    )
    p6_por_panel_path = (
        resolve_path(args.p6_por_panel, base_dir=Path.cwd())
        if args.p6_por_panel
        else resolve_path(P6_POR_PANEL_IMAGE_FILE, base_dir=PROJECT_DIR)
    )

    if por_badge_path is not None:
        require_file(por_badge_path, "POR badge image")
    if new_badge_path is not None:
        require_file(new_badge_path, "NEW badge image")
    if p6_por_panel_path is not None:
        require_file(p6_por_panel_path, "P6 POR panel image")

    output_path.parent.mkdir(parents=True, exist_ok=True)

    configure_asset_paths(
        por_badge_image_file=por_badge_path,
        new_badge_image_file=new_badge_path,
        p6_por_panel_image_file=p6_por_panel_path,
    )

    print(f"Excel: {excel_path}")
    print(f"Template: {template_path}")
    print(f"Output: {output_path}")
    print(f"Quality: {quality}")
    print(f"Template form: {template_form}")

    created = export_catalog(
        excel_path=excel_path,
        template_pdf=template_path,
        output_path=output_path,
        quality=quality,
        template_form=template_form,
        status_callback=print_status,
    )
    print(f"Created: {created}")


if __name__ == "__main__":
    main()
