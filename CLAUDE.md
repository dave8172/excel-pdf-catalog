# excel-to-pdf — Catalog PDF Exporter Web UI

> Flask upload page wrapping the existing `export_pdf.py` / `catalog_exporter.py` engine, so catalog PDFs can be generated from a phone browser instead of the owner's VS Code terminal. Deployed for the owner's own Upwork client work — not client-facing directly (client only ever receives the finished PDF, shared manually).

## What it is

- Repo: git@github.com:dave8172/excel-pdf-catalog.git (private)
- Also mirrored on a Windows PC at `t:\Docs\Upwork\excel to pdf\excel-pdf` (VS Code + Claude Code). Edits can happen from either copy — GitHub is the source of truth; just `git pull` wherever you didn't make the change before editing there again. This VPS copy is the deployed runtime, so changes made here still need `git push` (see Deploy below) to reach the Windows copy.
- Input: a product data file (`.xlsx` or `.csv`, same columns either way) + a template PDF, uploaded via a mobile-friendly single-page form (`app.py`). CSV support (2026-08-23) added via `_iter_rows_from_source()` in `catalog_exporter.py`, which both `resolve_template_variant()` and `load_products()` now go through instead of calling `load_workbook()` directly — handles UTF-8-BOM/cp1252 encoding and comma/semicolon delimiters (Excel's own "Save As CSV" quirks). Verified by the owner in production the same day using real client catalogs "W34" and "P12" (weekly/period catalog naming convention this client uses) — both generated correctly from CSV input.
- Output: composed catalog PDF, streamed back as a download. Nothing is retained server-side — each upload gets `uploads/<uuid>/`, deleted ~10s after the response is sent, plus an hourly sweep of anything older than 1h as a backstop.
- No login/auth by owner's explicit choice — URL is unlisted (DuckDNS subdomain), not indexed, no persistent data at rest to leak.

## Current state (2026-08-17)

- Live at https://excelpdf.duckdns.org
- Runs as systemd service `excel-pdf.service` (gunicorn, **1 worker**, 300s timeout) bound to `127.0.0.1:8020`. This VPS has 2GB RAM total shared across several other services — see `/root/projects/memory/resource-constraints.md`. 1 worker is intentional, not a bug: `ps`/`systemctl status` will always show 2 processes (gunicorn master + the 1 worker) — that's normal gunicorn architecture, not 2 redundant instances.
- (2026-08-17) `--max-requests 100 --max-requests-jitter 20` added to ExecStart — recycles the worker periodically so per-process RSS doesn't creep up across repeated PDF exports (Pillow/reportlab/pypdf don't always release freed memory back to the OS within a long-lived process). Also added `MemoryHigh=400M` (soft cgroup throttle) and `OOMScoreAdjust=300` (bias the kernel to kill this best-effort personal tool before critical services like postgres/nginx/sshd/tailscaled if the box ever hits real memory pressure). Trigger: a service restart logged `342.3M memory peak, 157.2M memory swap peak` during a real export on a box with only ~200MB free RAM at the time.
- nginx (`/etc/nginx/sites-available/excel-pdf`) reverse-proxies + terminates TLS (certbot/Let's Encrypt, auto-renew)
- Deploy = `git pull` in this folder + `systemctl restart excel-pdf` (see below)

## How to run / update

```bash
cd /root/projects/excel-to-pdf
git pull
./.venv/bin/pip install -r requirements.txt   # only if requirements.txt changed
systemctl restart excel-pdf
systemctl status excel-pdf --no-pager
journalctl -u excel-pdf -n 50 --no-pager      # logs
```

## Notable fixes

- `catalog_exporter.py`'s `try_font()` originally only checked Windows font paths (`C:/Windows/Fonts/arial.ttf`). Added a Linux fallback to `/usr/share/fonts/truetype/liberation/LiberationSans-*.ttf` (metric-compatible with Arial, installed via `fonts-liberation`) — without it, text would silently fall back to Pillow's tiny bitmap default font on this VPS.
- (2026-08-17) The page's "Generating..." status used to get stuck forever even on a *successful* export. Cause: `/generate` returns the PDF via `send_file(..., as_attachment=True)`, which sets `Content-Disposition: attachment`; a browser receiving that from a plain HTML form POST downloads the file in the background but does **not** navigate the page, so the JS that set "Generating..." on submit never got a chance to reset it. (Error responses happened to reset fine, since a non-attachment HTML response *does* navigate.) Fixed by submitting via `fetch` instead and handling the blob/download and UI reset in JS explicitly, for both success and error paths.

## Dependencies on this VPS

- `poppler-utils` (pdftoppm/pdfunite) — system package
- `fonts-liberation` — system package
- Python deps in this project's own `.venv` (not shared with other projects)
