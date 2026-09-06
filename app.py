"""Hosted catalog PDF exporter.

Served at **https://stuffs.bid/topdf**, which is a rewrite in front of this
app's own origin, **https://topdf.stuffs.bid/topdf**. Both URLs render the same
pages; the browser posts the actual export straight to the origin, because a
finished catalog is far bigger and far slower than a proxy hop should carry.
See `URL_PREFIX` below and this project's CLAUDE.md for the why.

Everything here is the public-facing wrapper. The engine is `catalog_exporter`,
which is shared with the command-line entry point and knows nothing about HTTP.
"""

from __future__ import annotations

import io
import json
import logging
import os
import shutil
import time
import uuid
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, render_template, request, send_file
from PIL import Image
from werkzeug.middleware.proxy_fix import ProxyFix

import catalog_exporter
from catalog_exporter import TEMPLATE_FORM_VARIANTS, export_catalog, get_pdf_page_count
from web.limits import Busy, ExportSlot, Quota
from web.security import fetch_remote_image, looks_like_pdf, looks_like_text, looks_like_xlsx

APP_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = APP_DIR / "uploads"
LOG_DIR = APP_DIR / "logs"
UPLOADS_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# The app lives under a path prefix so the same deployment answers correctly
# whether it is reached directly or through the stuffs.bid rewrite. This mirrors
# the `basePath` convention the other showcase zones on that domain use.
URL_PREFIX = "/topdf"
ORIGIN_URL = os.environ.get("TOPDF_ORIGIN", "https://topdf.stuffs.bid")
ALLOWED_ORIGINS = {"https://stuffs.bid", ORIGIN_URL}

# --- limits -----------------------------------------------------------------
# Sized for a 2GB VPS shared with several other services, where one export has
# been measured peaking around 350MB. See projects/memory/resource-constraints.md.
MAX_CONTENT_LENGTH = 30 * 1024 * 1024
MAX_TEMPLATE_PAGES = 12
MAX_PRODUCTS = 400
SNIFF_BYTES = 8192
QUOTA_WINDOWS = {"hourly": (6, 3600), "daily": (25, 86400)}

JOB_MAX_AGE_SECONDS = 3600
IMAGE_CACHE_MAX_BYTES = 300 * 1024 * 1024
IMAGE_CACHE_MAX_AGE_SECONDS = 24 * 3600

TEMPLATE_FORM_CHOICES = ["auto", *TEMPLATE_FORM_VARIANTS.keys()]
QUALITY_CHOICES = ("normal", "high")

# Backstop against a decompression bomb: a 12MB PNG download is within the
# fetch limit and can still expand to something that eats the box. Pillow warns
# at this figure and refuses at twice it, so the effective ceiling is ~120M
# pixels — set above any plausible template render (a high-quality A4 page is
# under 9M) and below what would exhaust memory. The fetch size cap in
# web/security.py is the primary control; this catches what slips past it.
Image.MAX_IMAGE_PIXELS = 60_000_000

catalog_exporter.configure_image_loading(
    fetcher=fetch_remote_image,
    allow_local_paths=False,
    max_products=MAX_PRODUCTS,
)

app = Flask(
    __name__,
    static_folder="web/static",
    static_url_path=f"{URL_PREFIX}/static",
    template_folder="web/templates",
)
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH
# nginx is the only thing that talks to this process, and it overwrites the
# forwarding headers on every request, so exactly one hop is trusted.
app.wsgi_app = ProxyFix(app.wsgi_app, x_for=1, x_proto=1, x_host=0)

quota = Quota(LOG_DIR / "quota.sqlite3", QUOTA_WINDOWS)
usage_log = logging.getLogger("topdf.usage")
usage_log.setLevel(logging.INFO)
usage_log.propagate = False
_usage_handler = logging.FileHandler(LOG_DIR / "usage.jsonl")
_usage_handler.setFormatter(logging.Formatter("%(message)s"))
usage_log.addHandler(_usage_handler)


def record_usage(**fields: object) -> None:
    """One line per export attempt. No IPs, no filenames, no file contents.

    This exists to answer one question -- does anyone actually use this -- and
    is deliberately too thin to answer any other.
    """
    fields["ts"] = round(time.time())
    usage_log.info(json.dumps(fields, separators=(",", ":")))


# ---------------------------------------------------------------------------
# Housekeeping
# ---------------------------------------------------------------------------

def sweep_old_jobs() -> None:
    cutoff = time.time() - JOB_MAX_AGE_SECONDS
    if not UPLOADS_DIR.exists():
        return
    for child in UPLOADS_DIR.iterdir():
        try:
            if child.is_dir() and child.stat().st_mtime < cutoff:
                shutil.rmtree(child, ignore_errors=True)
        except FileNotFoundError:
            continue


def prune_image_cache() -> None:
    """Keep the shared downloaded-image cache from growing without a bound.

    The cache is keyed by URL and shared across exports, which is what makes a
    re-run of the same catalog fast. Public traffic means it is also an
    unbounded write target, so it gets an age limit and a size ceiling.
    """
    cache_dir = catalog_exporter.IMAGE_CACHE_DIR
    if not cache_dir.exists():
        return
    cutoff = time.time() - IMAGE_CACHE_MAX_AGE_SECONDS
    entries: list[tuple[float, int, Path]] = []
    total = 0
    for child in cache_dir.iterdir():
        try:
            stat = child.stat()
        except FileNotFoundError:
            continue
        if not child.is_file():
            continue
        if stat.st_mtime < cutoff:
            child.unlink(missing_ok=True)
            continue
        entries.append((stat.st_mtime, stat.st_size, child))
        total += stat.st_size

    if total <= IMAGE_CACHE_MAX_BYTES:
        return
    for _mtime, size, path in sorted(entries):
        path.unlink(missing_ok=True)
        total -= size
        if total <= IMAGE_CACHE_MAX_BYTES:
            break


# ---------------------------------------------------------------------------
# Request plumbing
# ---------------------------------------------------------------------------

@app.after_request
def add_common_headers(response: Response) -> Response:
    origin = request.headers.get("Origin")
    if origin in ALLOWED_ORIGINS:
        response.headers["Access-Control-Allow-Origin"] = origin
        response.headers["Vary"] = "Origin"
        # Without this the page cannot read the filename off the download.
        response.headers["Access-Control-Expose-Headers"] = "Content-Disposition"
    # stuffs.bid also hosts private tooling; nothing on this domain is offered
    # to search engines, consistent with the other showcase zones there.
    response.headers["X-Robots-Tag"] = "noindex, nofollow"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    return response


@app.errorhandler(413)
def too_large(_error: object) -> tuple[Response, int]:
    limit_mb = MAX_CONTENT_LENGTH // (1024 * 1024)
    return jsonify(error=f"Those files add up to more than {limit_mb} MB. Please upload smaller files."), 413


@app.route("/", methods=["GET"])
def root() -> Response:
    return redirect(URL_PREFIX, code=302)


@app.route("/robots.txt", methods=["GET"])
def robots() -> Response:
    return Response("User-agent: *\nDisallow: /\n", mimetype="text/plain")


@app.route(f"{URL_PREFIX}/healthz", methods=["GET"])
def healthz() -> Response:
    return jsonify(ok=True)


# Registered without the trailing slash, and matching both forms, because the
# Next.js rewrite in front of this normalises `/topdf/` to `/topdf`. A Flask
# rule written as `/topdf/` would redirect that straight back -- an infinite
# loop, and one whose Location header leaks this origin's hostname into the
# visitor's address bar.
@app.route(URL_PREFIX, methods=["GET"], strict_slashes=False)
def index() -> str:
    sweep_old_jobs()
    return render_template(
        "tool.html",
        template_forms=TEMPLATE_FORM_CHOICES,
        origin=ORIGIN_URL,
        prefix=URL_PREFIX,
        max_mb=MAX_CONTENT_LENGTH // (1024 * 1024),
    )


@app.route(f"{URL_PREFIX}/guide", methods=["GET"])
def guide() -> str:
    return render_template(
        "guide.html",
        prefix=URL_PREFIX,
        max_products=MAX_PRODUCTS,
        max_template_pages=MAX_TEMPLATE_PAGES,
        max_mb=MAX_CONTENT_LENGTH // (1024 * 1024),
    )


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

class RejectedUpload(Exception):
    """A problem with what was uploaded, safe to show the person who uploaded it."""


def validated_upload(field: str, allowed_suffixes: tuple[str, ...], label: str) -> tuple[object, str, bytes]:
    uploaded = request.files.get(field)
    if not uploaded or not uploaded.filename:
        raise RejectedUpload(f"Please choose a {label}.")

    suffix = Path(uploaded.filename).suffix.lower()
    if suffix not in allowed_suffixes:
        allowed = " or ".join(allowed_suffixes)
        raise RejectedUpload(f"The {label} must be {allowed}.")

    head = uploaded.stream.read(SNIFF_BYTES)
    uploaded.stream.seek(0)
    if not head:
        raise RejectedUpload(f"That {label} is empty.")

    # The extension is a claim; the first bytes are evidence.
    if suffix == ".xlsx" and not looks_like_xlsx(head):
        raise RejectedUpload("That .xlsx does not look like an Excel workbook. Re-save it from Excel as .xlsx.")
    if suffix == ".pdf" and not looks_like_pdf(head):
        raise RejectedUpload("That .pdf does not look like a PDF file.")
    if suffix == ".csv" and not looks_like_text(head):
        raise RejectedUpload("That .csv does not look like text. Re-save it from Excel as CSV.")

    return uploaded, suffix, head


@app.route(f"{URL_PREFIX}/generate", methods=["POST", "OPTIONS"])
def generate() -> Response | tuple[Response, int]:
    if request.method == "OPTIONS":
        response = Response(status=204)
        response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
        response.headers["Access-Control-Allow-Headers"] = "Content-Type"
        response.headers["Access-Control-Max-Age"] = "86400"
        return response

    started = time.monotonic()
    quality = request.form.get("quality", "normal")
    template_form = request.form.get("template_form", "auto")

    try:
        if quality not in QUALITY_CHOICES:
            raise RejectedUpload("Unknown quality setting.")
        if template_form not in TEMPLATE_FORM_CHOICES:
            raise RejectedUpload("Unknown template form.")

        excel_file, excel_suffix, _ = validated_upload(
            "excel", (".xlsx", ".csv"), "product file (.xlsx or .csv)"
        )
        template_file, _, _ = validated_upload("template", (".pdf",), "template PDF")
    except RejectedUpload as rejection:
        record_usage(ok=False, reason="rejected", ms=round((time.monotonic() - started) * 1000))
        return jsonify(error=str(rejection)), 400

    allowed, message = quota.check_and_record(quota.subject(request.remote_addr), "generate")
    if not allowed:
        record_usage(ok=False, reason="quota")
        return jsonify(error=message), 429

    sweep_old_jobs()
    prune_image_cache()

    job_id = uuid.uuid4().hex
    job_dir = UPLOADS_DIR / job_id
    job_dir.mkdir(parents=True, exist_ok=True)

    excel_path = job_dir / f"input{excel_suffix}"
    template_path = job_dir / "template.pdf"
    output_path = job_dir / "output.pdf"
    download_name = f"{catalog_exporter.sanitize_filename(Path(excel_file.filename).stem)} catalog.pdf"

    try:
        excel_file.save(excel_path)
        template_file.save(template_path)

        try:
            page_count = get_pdf_page_count(template_path)
        except Exception:
            return jsonify(error="That template PDF could not be read. Try re-exporting it."), 400
        if page_count > MAX_TEMPLATE_PAGES:
            return jsonify(
                error=(
                    f"That template has {page_count} pages. A template needs at most "
                    f"{MAX_TEMPLATE_PAGES} — it is the page design to repeat, not the finished catalog."
                )
            ), 400

        try:
            with ExportSlot():
                export_catalog(
                    excel_path=excel_path,
                    template_pdf=template_path,
                    output_path=output_path,
                    quality=quality,
                    template_form=template_form,
                )
        except Busy:
            record_usage(ok=False, reason="busy")
            return jsonify(
                error="Someone else's catalog is generating right now. Please try again in a minute."
            ), 503
        except ValueError as invalid:
            # The engine raises ValueError for the things the uploader can
            # actually fix -- missing columns, an unusable template -- and its
            # wording is already aimed at them.
            record_usage(ok=False, reason="invalid_input", ms=round((time.monotonic() - started) * 1000))
            return jsonify(error=str(invalid)), 400
        except Exception:
            # Anything else is a bug or a broken file, and its text can carry
            # server paths. The person gets a reference; the detail goes to the
            # journal, where only the operator can read it.
            reference = job_id[:8]
            app.logger.exception("export failed (ref %s)", reference)
            record_usage(ok=False, reason="error", ms=round((time.monotonic() - started) * 1000))
            return jsonify(
                error=(
                    "The export failed on this file. If the product file and template both match "
                    f"the Guide, this is a bug — quote reference {reference}."
                )
            ), 500

        payload = output_path.read_bytes()
    finally:
        # Read into memory and delete now, rather than deleting on a timer after
        # the response streams: nothing of a visitor's catalog outlives the
        # request, and there is no window where a crash leaves it on disk.
        shutil.rmtree(job_dir, ignore_errors=True)

    elapsed_ms = round((time.monotonic() - started) * 1000)
    record_usage(
        ok=True,
        ms=elapsed_ms,
        quality=quality,
        form=template_form,
        template_pages=page_count,
        out_kb=round(len(payload) / 1024),
        via="proxy" if request.headers.get("Origin") == "https://stuffs.bid" else "direct",
    )
    return send_file(
        io.BytesIO(payload),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=download_name,
    )


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8020, debug=False)
