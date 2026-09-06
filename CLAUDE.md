# excel-to-pdf — Catalog PDF Exporter

> Turns a product spreadsheet plus a template PDF into a finished catalog PDF. Built for the owner's own Upwork client work, **public since 2026-09-06** at `https://stuffs.bid/topdf` as one of the micro-tools on that domain.

## What it is

- Repo: git@github.com:dave8172/excel-pdf-catalog.git (private)
- Also mirrored on a Windows PC at `t:\Docs\Upwork\excel to pdf\excel-pdf` (VS Code + Claude Code). Edits can happen from either copy — GitHub is the source of truth; just `git pull` wherever you didn't make the change before editing there again. This VPS copy is the deployed runtime, so changes made here still need `git push` to reach the Windows copy.
- Input: a product file (`.xlsx` or `.csv`) + a template PDF. Output: the composed catalog PDF, streamed back as a download.
- Nothing is retained. Each upload gets `uploads/<uuid>/`; the finished PDF is read into memory and the directory is deleted in the same request's `finally`, with an hourly sweep as a backstop.

## Where it lives

Two hostnames, on purpose:

| URL | What it is |
|---|---|
| `https://stuffs.bid/topdf` | **The public address.** A Next.js rewrite on the `stuffboard` deployment. |
| `https://topdf.stuffs.bid/topdf` | The origin — this app, on this VPS. Serves the same pages, and receives the exports directly. |
| `https://excelpdf.duckdns.org` | Retired 2026-09-06. 301s to `stuffs.bid/topdf`, nothing else. |

**Only the pages go through the rewrite; the export does not.** `tool.html` posts to the origin's absolute URL. Two hard platform limits make the proxy hop the wrong place for it: Vercel caps a proxied body at 4.5MB (a template PDF plus a finished catalog routinely exceeds that) and times an origin out at two minutes (a large catalog with cold image downloads exceeds that too). Routing the export through Vercel would break exactly the big jobs the tool exists for. CORS on the origin allows `https://stuffs.bid`, and exposes `Content-Disposition` so the page can read the filename off the download.

The app is served under the `/topdf` **path prefix** so both hostnames serve identical URLs — the same `basePath` convention the other showcase zones on that domain use. The index route is registered with `strict_slashes=False`: Next.js normalises `/topdf/` to `/topdf`, and a Flask rule written as `/topdf/` would redirect it straight back — an infinite loop whose `Location` also leaks the origin hostname into the address bar.

## Layout

```
catalog_exporter.py     the engine. 2,200 lines, shared with the CLI, knows nothing about HTTP
export_pdf.py           command-line entry point
app.py                  the hosted app: routes, validation, quotas, error mapping
web/security.py         SSRF-guarded image fetch + upload sniffing
web/limits.py           SQLite per-visitor quota + the one-export-at-a-time semaphore
web/templates/          base.html, tool.html, guide.html
web/static/samples/     everything the Guide hands out — all generated, none from a client
scripts/make_samples.py regenerates that folder end to end
```

## The Guide (`/topdf/guide`)

Added 2026-09-06 with the public launch. Without it a stranger cannot succeed: the template has to satisfy a page-layout detector they cannot see, and the required columns were named after one client's Shopify metafields.

**No client file is used anywhere in it.** `scripts/make_samples.py` draws the product images, invents the 14 products, and builds both template PDFs from the exporter's own reference-grid constants — so the samples are correct by construction rather than by copying one that happens to work. Re-run it after any change to the grid constants or the sample data:

```bash
./.venv/bin/python scripts/make_samples.py
```

It **verifies as it goes** — it runs the exporter's own `classify_page_layout` over each template page and refuses to finish if a page classifies wrong, then exports both samples for real to produce the "what comes back" screenshots. A broken guide fails the script instead of shipping.

Two things that took a round to get right, both encoded in the script's comments:

- **Border weight.** The detector samples for ink ~2px inside each box edge on a 72dpi render. A hairline centred on the outline half-misses it. The strokes are drawn fully *inside* the outline at 3.2pt, so the box keeps its exact outer size and the ink lands where the detector looks.
- **The promotion template is three pages, and the plain one is two.** `P6_FIRST/MIDDLE/LAST_PAGE_ROWS` are three *different* row sets; a 2-page promotion template makes the exporter draw last-page products into rows the cover page has no boxes in. Its page 2 also carries a banner in the top-row band rather than boxes — a middle page is *recognised* by ink there, but the promotion layout never fills it.

## Hardening for public use (2026-09-06)

The engine's two most dangerous behaviours were harmless while its author wrote the input files:

- **`Image Src` was fetched with `urlopen`.** On a box also running Postgres, an internal wiki and Tailscale, that is an SSRF primitive handed to whoever uploads. `web/security.py` now resolves the host, rejects the name if *any* answer is non-public (including 100.64/10, which `ipaddress.is_private` misses and Tailscale uses), and then connects to the address it vetted — so the second lookup a DNS-rebinding attack needs never happens. Redirects are followed manually, three hops max, each re-validated.
- **A non-URL `Image Src` was opened as a filesystem path.** `configure_image_loading(allow_local_paths=False)` turns that off for the hosted app only; the CLI keeps it, because there the operator owns both the machine and the spreadsheet.

Also added: upload magic-byte sniffing, a 30MB body cap, 400 products, 12 template pages, a 12MB-per-image fetch cap, a Pillow pixel ceiling, a bounded image cache, and generic error text with a reference id for anything that is not a `ValueError` the engine raised for the uploader (those are already worded for them and are passed through).

**Quotas and concurrency are separate problems solved separately** (`web/limits.py`): 6 exports/hour and 25/day per hashed IP in SQLite, because it must survive `--max-requests` recycling; and a `BoundedSemaphore(1)` around the export itself, because a single export peaks near 350MB on a 2GB box. nginx adds a cheap edge limit in front of both.

Column names were also generalised — `POR`, `Price`, `Image` and friends now resolve alongside the original `Metafield: custom.tier_c_profit_on_return [number_decimal]` spellings, which still take priority. A price that arrives with its own currency symbol keeps it instead of being forced to £.

## Running and deploying

```bash
cd /root/projects/excel-to-pdf
git pull
./.venv/bin/pip install -r requirements.txt   # only if requirements.txt changed
systemctl restart excel-pdf
journalctl -u excel-pdf -n 50 --no-pager
```

**A change to anything under `web/templates/` needs the restart too** — Jinja caches compiled templates outside debug mode, so an edited page keeps serving the old copy with no error to notice.

Service: `excel-pdf.service` — gunicorn, **1 process, 4 gthread threads**, bound to `127.0.0.1:8020`. One process is what bounds memory; the threads exist so a visitor loading the page is not stuck behind someone else's export. `MemoryHigh=550M`, `OOMScoreAdjust=300`, `--max-requests 200` to stop Pillow/reportlab RSS creeping. See `/root/projects/memory/resource-constraints.md`.

nginx: `/etc/nginx/sites-available/topdf` (+ rate-limit zones in `conf.d/topdf-limits.conf`). **Its `listen` is pinned to `65.20.79.200:443`, not `0.0.0.0`** — tailscaled already holds `:443` on its own interface, so a wildcard bind fails with EADDRINUSE and nginx silently keeps running the old config. Certbot writes `listen 443 ssl` on reconfiguration; if TLS breaks after a renewal, check that line first.

## Usage measurement

`logs/usage.jsonl` (gitignored), one line per export attempt: outcome, duration, template pages, output size, and whether it came via the proxy or direct. No IPs, no filenames, no content — deliberately too thin to answer anything except *does anyone use this*, which is the only reason the tool is public.

## System dependencies

`poppler-utils` (pdftoppm/pdfunite), `fonts-liberation`. Python deps in this project's own `.venv`.

## Notable fixes

- `try_font()` originally only checked Windows font paths. A Linux fallback to `/usr/share/fonts/truetype/liberation/` was added — without it text silently fell back to Pillow's bitmap default.
- (2026-08-17) "Generating..." used to stick forever even on success: `send_file(as_attachment=True)` downloads without navigating, so a plain form POST never reset the UI. Fixed by submitting via `fetch` and handling the blob in JS. The current page keeps that and adds an elapsed-seconds counter, because a minute of silence reads as a hang.
- (2026-08-23) CSV support via `_iter_rows_from_source()`, handling UTF-8-BOM/cp1252 and comma/semicolon delimiters. Verified in production on real client catalogs the same day.

## If it grows

The export is synchronous, and a second visitor during an export gets a worded 503 rather than a queue. That is the right trade at this traffic. If the usage log ever shows real concurrent demand, the engine already exposes `status_callback`/`progress_callback` — the upgrade is a job id plus polling, not a rewrite.
