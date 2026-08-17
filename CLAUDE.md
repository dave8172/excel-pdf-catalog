# excel-to-pdf — Catalog PDF Exporter Web UI

> Flask upload page wrapping the existing `export_pdf.py` / `catalog_exporter.py` engine, so catalog PDFs can be generated from a phone browser instead of the owner's VS Code terminal. Deployed for the owner's own Upwork client work — not client-facing directly (client only ever receives the finished PDF, shared manually).

## What it is

- Repo: git@github.com:dave8172/excel-pdf-catalog.git (private)
- Also mirrored on a Windows PC at `t:\Docs\Upwork\excel to pdf\excel-pdf` (VS Code + Claude Code). Edits can happen from either copy — GitHub is the source of truth; just `git pull` wherever you didn't make the change before editing there again. This VPS copy is the deployed runtime, so changes made here still need `git push` (see Deploy below) to reach the Windows copy.
- Input: an Excel file (product data) + a template PDF, uploaded via a mobile-friendly single-page form (`app.py`).
- Output: composed catalog PDF, streamed back as a download. Nothing is retained server-side — each upload gets `uploads/<uuid>/`, deleted ~10s after the response is sent, plus an hourly sweep of anything older than 1h as a backstop.
- No login/auth by owner's explicit choice — URL is unlisted (DuckDNS subdomain), not indexed, no persistent data at rest to leak.

## Current state (2026-08-17)

- Live at https://excelpdf.duckdns.org
- Runs as systemd service `excel-pdf.service` (gunicorn, 2 workers, 300s timeout) bound to `127.0.0.1:8020`
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
