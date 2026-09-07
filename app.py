"""Two catalog tools, one process.

**Public — topdf.** `https://stuffs.bid/topdf` (a Next.js rewrite in front of
`https://topdf.stuffs.bid/topdf`, which is this app). Five plain-English
columns, no options, one card design. Its engine is `simple_catalog`.

**Client.** `https://excelpdf.duckdns.org`. The original tool, with the card
styles, badge designs, quality choice and fixed reference grid that belong to
one client's catalog. Its engine is `catalog_exporter`, untouched.

The split was made 2026-09-07: the public product had been carrying the
client's design, sample assets and column vocabulary, which is not ours to
publish. They now share only this process, the export semaphore and the
upload validation — the box has 2GB and one export peaks near 350MB, so a
second service is the thing that must not happen, not a second engine.

Which tool a request gets is decided by its Host header; see `is_client_host`.
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

from flask import (
    Flask,
    Response,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    send_from_directory,
)
from PIL import Image
from werkzeug.middleware.proxy_fix import ProxyFix

import catalog_exporter
import simple_catalog
from catalog_exporter import TEMPLATE_FORM_VARIANTS, export_catalog, get_pdf_page_count
from web.limits import Busy, ExportSlot, Quota
from web.security import fetch_remote_image, looks_like_pdf, looks_like_text, looks_like_xlsx

APP_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = APP_DIR / "uploads"
LOG_DIR = APP_DIR / "logs"
CLIENT_STATIC_DIR = APP_DIR / "web" / "client_static"
UPLOADS_DIR.mkdir(exist_ok=True)
LOG_DIR.mkdir(exist_ok=True)

# --- addresses --------------------------------------------------------------
# The public app lives under a path prefix so the same deployment answers
# correctly whether it is reached directly or through the stuffs.bid rewrite.
# This mirrors the `basePath` convention the other showcase zones use.
URL_PREFIX = "/topdf"
ORIGIN_URL = os.environ.get("TOPDF_ORIGIN", "https://topdf.stuffs.bid")
PUBLIC_BASE = os.environ.get("TOPDF_PUBLIC_BASE", "https://stuffs.bid")
CLIENT_HOST = os.environ.get("TOPDF_CLIENT_HOST", "excelpdf.duckdns.org")
CLIENT_URL = f"https://{CLIENT_HOST}"
ALLOWED_ORIGINS = {PUBLIC_BASE, ORIGIN_URL, CLIENT_URL}

# --- limits -----------------------------------------------------------------
# Sized for a 2GB VPS shared with several other services, where one export has
# been measured peaking around 350MB. See projects/memory/resource-constraints.md.
MAX_CONTENT_LENGTH = 30 * 1024 * 1024
MAX_TEMPLATE_PAGES = 12
MAX_PRODUCTS = 400
SNIFF_BYTES = 8192
# Strangers get a tight allowance; the client host is one person doing a known
# weekly job and should never meet a quota wall mid-catalog.
PUBLIC_QUOTA_WINDOWS = {"hourly": (6, 3600), "daily": (25, 86400)}
CLIENT_QUOTA_WINDOWS = {"hourly": (30, 3600), "daily": (120, 86400)}

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

public_quota = Quota(LOG_DIR / "quota.sqlite3", PUBLIC_QUOTA_WINDOWS)
client_quota = Quota(LOG_DIR / "quota.sqlite3", CLIENT_QUOTA_WINDOWS)

usage_log = logging.getLogger("topdf.usage")
usage_log.setLevel(logging.INFO)
usage_log.propagate = False
_usage_handler = logging.FileHandler(LOG_DIR / "usage.jsonl")
_usage_handler.setFormatter(logging.Formatter("%(message)s"))
usage_log.addHandler(_usage_handler)


def is_client_host() -> bool:
    """True when this request arrived on the client tool's own hostname."""
    return request.host.split(":")[0].lower() == CLIENT_HOST


def record_usage(**fields: object) -> None:
    """One line per export attempt. No IPs, no filenames, no file contents.

    This exists to answer one question -- does anyone actually use the public
    tool -- and is deliberately too thin to answer any other. `app` separates
    the two tools so client runs never look like demand.
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
    # The public pages are indexable on purpose — the tool is published to find
    # out whether anyone outside one client wants it, and a page search engines
    # may not read cannot answer that. The client host is not: it is one
    # person's working tool and has no reason to be in an index. Duplicate
    # content between this origin and stuffs.bid is handled by the canonical
    # tag plus this origin's robots.txt, not by hiding the pages.
    if is_client_host():
        response.headers["X-Robots-Tag"] = "noindex, nofollow"
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    return response


@app.errorhandler(413)
def too_large(_error: object) -> tuple[Response, int]:
    limit_mb = MAX_CONTENT_LENGTH // (1024 * 1024)
    return jsonify(error=f"Those files add up to more than {limit_mb} MB. Please upload smaller files."), 413


@app.route("/robots.txt", methods=["GET"])
def robots() -> Response:
    """Keep both of *this app's* hostnames out of search results.

    The public pages are indexable at `https://stuffs.bid/topdf`, which is a
    different hostname serving the identical HTML through a rewrite. A crawler
    only reads this file if it reached `topdf.stuffs.bid` or the client host
    directly, and the right answer at either is "not here."
    """
    return Response("User-agent: *\nDisallow: /\n", mimetype="text/plain")


@app.route(f"{URL_PREFIX}/healthz", methods=["GET"])
def healthz() -> Response:
    return jsonify(ok=True)


# ---------------------------------------------------------------------------
# Pages
# ---------------------------------------------------------------------------

@app.route("/", methods=["GET"])
def root() -> Response | str:
    if not is_client_host():
        return redirect(URL_PREFIX, code=302)
    sweep_old_jobs()
    return render_template(
        "client.html",
        template_forms=TEMPLATE_FORM_CHOICES,
        generate_url=f"{CLIENT_URL}/generate",
        max_products=MAX_PRODUCTS,
        max_mb=MAX_CONTENT_LENGTH // (1024 * 1024),
    )


@app.route("/client-assets/<path:filename>", methods=["GET"])
def client_assets(filename: str) -> Response:
    """The client tool's own sample files, served only on the client host.

    They imitate that client's catalog design, so they must not be reachable
    from any public URL — which is why they live outside the app's static
    folder rather than being served from it with a different prefix.
    """
    if not is_client_host():
        return Response("Not found", status=404)
    return send_from_directory(CLIENT_STATIC_DIR, filename)


# Registered without the trailing slash, and matching both forms, because the
# Next.js rewrite in front of this normalises `/topdf/` to `/topdf`. A Flask
# rule written as `/topdf/` would redirect that straight back -- an infinite
# loop, and one whose Location header leaks this origin's hostname into the
# visitor's address bar.
@app.route(URL_PREFIX, methods=["GET"], strict_slashes=False)
def index() -> Response | str:
    if is_client_host():
        return redirect("/", code=302)
    sweep_old_jobs()
    return render_template(
        "landing.html",
        origin=ORIGIN_URL,
        prefix=URL_PREFIX,
        public_base=PUBLIC_BASE,
        canonical=f"{PUBLIC_BASE}{URL_PREFIX}",
        max_products=MAX_PRODUCTS,
        max_mb=MAX_CONTENT_LENGTH // (1024 * 1024),
    )


@app.route(f"{URL_PREFIX}/guide", methods=["GET"])
def guide() -> Response | str:
    if is_client_host():
        return redirect("/", code=302)
    return render_template(
        "guide.html",
        prefix=URL_PREFIX,
        public_base=PUBLIC_BASE,
        canonical=f"{PUBLIC_BASE}{URL_PREFIX}/guide",
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


def preflight_response() -> Response:
    response = Response(status=204)
    response.headers["Access-Control-Allow-Methods"] = "POST, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type"
    response.headers["Access-Control-Max-Age"] = "86400"
    return response


def run_export(*, tool: str, run) -> Response | tuple[Response, int]:
    """Everything both tools do around the export itself.

    `run(excel_path, template_path, output_path)` is the only difference
    between them, and it is where the two engines diverge completely.
    """
    started = time.monotonic()

    try:
        excel_file, excel_suffix, _ = validated_upload(
            "excel", (".xlsx", ".csv"), "product file (.xlsx or .csv)"
        )
        template_file, _, _ = validated_upload("template", (".pdf",), "template PDF")
    except RejectedUpload as rejection:
        record_usage(tool=tool, ok=False, reason="rejected", ms=round((time.monotonic() - started) * 1000))
        return jsonify(error=str(rejection)), 400

    quota = client_quota if tool == "client" else public_quota
    allowed, message = quota.check_and_record(quota.subject(request.remote_addr), f"generate:{tool}")
    if not allowed:
        record_usage(tool=tool, ok=False, reason="quota")
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
                run(excel_path, template_path, output_path)
        except Busy:
            record_usage(tool=tool, ok=False, reason="busy")
            return jsonify(
                error="Someone else's catalog is generating right now. Please try again in a minute."
            ), 503
        except ValueError as invalid:
            # Both engines raise ValueError for the things the uploader can
            # actually fix -- missing columns, an unusable template -- and the
            # wording is already aimed at them.
            record_usage(tool=tool, ok=False, reason="invalid_input", ms=round((time.monotonic() - started) * 1000))
            return jsonify(error=str(invalid)), 400
        except Exception:
            # Anything else is a bug or a broken file, and its text can carry
            # server paths. The person gets a reference; the detail goes to the
            # journal, where only the operator can read it.
            reference = job_id[:8]
            app.logger.exception("export failed (ref %s)", reference)
            record_usage(tool=tool, ok=False, reason="error", ms=round((time.monotonic() - started) * 1000))
            return jsonify(
                error=(
                    "The export failed on this file. If the product file and template both match "
                    f"the guide, this is a bug — quote reference {reference}."
                )
            ), 500

        payload = output_path.read_bytes()
    finally:
        # Read into memory and delete now, rather than deleting on a timer after
        # the response streams: nothing of a visitor's catalog outlives the
        # request, and there is no window where a crash leaves it on disk.
        shutil.rmtree(job_dir, ignore_errors=True)

    record_usage(
        tool=tool,
        ok=True,
        ms=round((time.monotonic() - started) * 1000),
        template_pages=page_count,
        out_kb=round(len(payload) / 1024),
        via="proxy" if request.headers.get("Origin") == PUBLIC_BASE else "direct",
    )
    return send_file(
        io.BytesIO(payload),
        mimetype="application/pdf",
        as_attachment=True,
        download_name=download_name,
    )


@app.route(f"{URL_PREFIX}/generate", methods=["POST", "OPTIONS"])
def generate_public() -> Response | tuple[Response, int]:
    """topdf's export. No options: one card design, one quality."""
    if request.method == "OPTIONS":
        return preflight_response()
    if is_client_host():
        return jsonify(error="Not found."), 404

    def run(excel_path: Path, template_path: Path, output_path: Path) -> None:
        simple_catalog.export_catalog(
            excel_path=excel_path,
            template_pdf=template_path,
            output_path=output_path,
            max_products=MAX_PRODUCTS,
        )

    return run_export(tool="topdf", run=run)


@app.route("/generate", methods=["POST", "OPTIONS"])
def generate_client() -> Response | tuple[Response, int]:
    """The client tool's export, with its card styles and quality choice."""
    if request.method == "OPTIONS":
        return preflight_response()
    if not is_client_host():
        return jsonify(error="Not found."), 404

    quality = request.form.get("quality", "normal")
    template_form = request.form.get("template_form", "auto")
    if quality not in QUALITY_CHOICES:
        return jsonify(error="Unknown quality setting."), 400
    if template_form not in TEMPLATE_FORM_CHOICES:
        return jsonify(error="Unknown template form."), 400

    def run(excel_path: Path, template_path: Path, output_path: Path) -> None:
        export_catalog(
            excel_path=excel_path,
            template_pdf=template_path,
            output_path=output_path,
            quality=quality,
            template_form=template_form,
        )

    return run_export(tool="client", run=run)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8020, debug=False)
