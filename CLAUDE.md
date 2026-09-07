# excel-to-pdf — two catalog tools, one process

> A product spreadsheet plus a template PDF becomes a finished catalog PDF. **Two tools do that here, and keeping them apart is the point of the current design** — see the split below before changing anything.

## The split (2026-09-07) — read this first

The tool was built for one Upwork client's catalog and went public on 2026-09-06 exactly as it was. That was wrong: the public product was carrying **that client's design** — the profit-on-return and NEW badge artwork, the promotion card, their Shopify metafield column names, their fixed page grid, and sample assets built to imitate all of it. None of that is ours to publish.

So there are now two tools, and they share only a process:

| | **topdf** (public) | **client** |
|---|---|---|
| Address | `stuffs.bid/topdf` | `excelpdf.duckdns.org` |
| Engine | `simple_catalog.py` | `catalog_exporter.py` (untouched) |
| Columns | Image URL, Name, Description, Price, Case Size | `Image Src`, `Title`, `POR`, metafield spellings, … |
| Card | one design, no badges | POR badge, NEW flash, promotion price panel |
| Options | none | card style (`auto`/`por-title`/`promotion`), quality (`normal`/`high`) |
| Grid | measured off the template | fixed reference rows/columns |
| Samples | `web/static/samples/` | `web/client_static/samples/` |
| Indexed | yes | no |

**The rules that keep them apart, all enforced in code:**

- **`catalog_exporter.py` is the client's and is not modified for public features.** The public engine imports only generic helpers from it — PDF page render/merge, the hardened image fetch, white-trim, font lookup. Nothing design-bearing.
- **`web/client_static/` is outside Flask's static folder** and is served by one route that 404s unless the request arrived on the client host. Putting those files under `web/static/` with a different prefix would still expose them; this does not.
- **Each export endpoint 404s on the wrong host** (`generate_public` on the client host, `generate_client` on the public one), so neither engine is reachable from the other's address.
- **One process, not two.** The box is 2GB and one export peaks near 350MB, so a second gunicorn service is the thing that must not happen. Both tools go through the same `BoundedSemaphore(1)`.

**Why one process and not two repos:** the split is about what is *published*, not about isolation for its own sake. A second service would double the memory floor on a box that cannot afford it, and the host check is a two-line function.

## What it is

- Repo: git@github.com:dave8172/excel-pdf-catalog.git (private)
- Also mirrored on a Windows PC at `t:\Docs\Upwork\excel to pdf\excel-pdf` (VS Code + Claude Code). Edits can happen from either copy — GitHub is the source of truth; just `git pull` wherever you didn't make the change before editing there again. This VPS copy is the deployed runtime, so changes made here still need `git push` to reach the Windows copy.
- Input: a product file (`.xlsx` or `.csv`) + a template PDF. Output: the composed catalog PDF, streamed back as a download.
- Nothing is retained. Each upload gets `uploads/<uuid>/`; the finished PDF is read into memory and the directory is deleted in the same request's `finally`, with an hourly sweep as a backstop.

## Where it lives

| URL | What it is |
|---|---|
| `https://stuffs.bid/topdf` | **topdf's public address.** A Next.js rewrite on the `stuffboard` deployment. |
| `https://topdf.stuffs.bid/topdf` | The origin — this app, on this VPS. Serves the same pages, and receives the exports directly. |
| `https://excelpdf.duckdns.org` | **The client tool.** Restored 2026-09-07, having been a redirect to `stuffs.bid/topdf` for one day. |

**Only the pages go through the rewrite; the export does not.** `landing.html` posts to the origin's absolute URL. Two hard platform limits make the proxy hop the wrong place for it: Vercel caps a proxied body at 4.5MB (a template PDF plus a finished catalog routinely exceeds that) and times an origin out at two minutes (a large catalog with cold image downloads exceeds that too). Routing the export through Vercel would break exactly the big jobs the tool exists for. CORS on the origin allows `https://stuffs.bid`, and exposes `Content-Disposition` so the page can read the filename off the download.

The app is served under the `/topdf` **path prefix** so both hostnames serve identical URLs — the same `basePath` convention the other showcase zones on that domain use. The index route is registered with `strict_slashes=False`: Next.js normalises `/topdf/` to `/topdf`, and a Flask rule written as `/topdf/` would redirect it straight back — an infinite loop whose `Location` also leaks the origin hostname into the address bar.

## Layout

```
catalog_exporter.py            the CLIENT engine. 2,200 lines, shared with the CLI, no HTTP
simple_catalog.py              the PUBLIC engine. Own columns, own card, own grid detection
export_pdf.py                  command-line entry point (client engine)
app.py                         both tools: host routing, validation, quotas, error mapping
web/security.py                SSRF-guarded image fetch + upload sniffing
web/limits.py                  SQLite per-visitor quota + the one-export-at-a-time semaphore
web/templates/base.html        public chrome (nav, meta, canonical)
web/templates/landing.html     topdf's landing page — the form lives on it
web/templates/guide.html       topdf's guide
web/templates/client.html      the client tool, standalone: no shared base, noindex
web/static/samples/            topdf's samples. Teal, workshop products, five columns
web/client_static/samples/     the client tool's samples. Never served on a public host
scripts/make_public_samples.py regenerates web/static/samples end to end, and verifies it
scripts/make_samples.py        regenerates web/client_static/samples the same way
```

## topdf's engine — `simple_catalog.py`

Five columns, two required (`Image URL`, `Name`); `Description`, `Price` and `Case Size` are
optional and simply do not appear on the card when absent. One card design, one quality, no
options anywhere in the UI.

**The part worth understanding is the grid detection**, because it is what makes this a
product rather than the client's tool with the labels changed. The client engine knows where
boxes are; this one has to find them on a stranger's template. It builds an ink mask, keeps
only ink that is *thin* (`_thin_ink`), takes the rows and columns holding a long unbroken run
of it, treats every adjacent pair of lines as a candidate rectangle, and keeps the ones whose
four edges are inked and whose middle is still blank.

Two things in there are load-bearing and were each found by a failing template:

- **The thinness filter.** Without it a solid header band makes *every column beneath it*
  look like a vertical rule — a column through the band is one long unbroken run of ink — and
  detection returns nothing at all. Keeping only ink with white a few pixels either side
  leaves the rules and drops the fills.
- **The blank-middle test.** Edge ink alone cannot tell a box from a banner. This is what
  stops products being drawn over a section banner or a dark footer strip.

Performance without numpy: the mask goes to `bytes` once and run lengths come from
`bytes.split(b"\x00")`, which is C-speed. Per-pixel Python over a 150dpi A4 page is seconds;
this is milliseconds. Columns use `Image.Transpose.TRANSPOSE` (not `ROTATE_90`) so the
indexes that come back are already x coordinates and need no un-flipping.

Pagination rule, chosen so it needs no setting: template pages are used in order, and **the
last page that has boxes** repeats until the products run out. A cover plus a repeating inner
page therefore just works. It has to be the last page *with boxes* — a template ending in a
terms page would otherwise repeat that forever and place nothing.

## topdf's landing page (2026-09-07)

`/topdf` was the upload form and nothing else, titled "Catalog PDF Exporter" with a
`noindex` on it — so a stranger arriving cold had no idea what it made, and nobody could
arrive cold in the first place. It is now a real product page: hero with a finished export
beside it, before/after pairs, the form itself, the five columns, how-it-works, who it is
for, features, and an FAQ. **The form stays on `/topdf`** rather than moving behind a
landing page, so the tool is never more than one screen away.

**It is indexable now, and that is the point** — the tool is public to find out whether the
job it does is wanted outside one client, and a page search engines may not read cannot
answer that. Three things were in the way and all three are gone: the `noindex` meta tag,
the blanket `X-Robots-Tag` in `add_common_headers`, and a title that named the product
instead of the task. `stuffs.bid` is otherwise unaffected — see that project's CLAUDE.md.

Two hostnames serve identical HTML, so every page carries `<link rel="canonical">` pointing
at `PUBLIC_BASE` (`https://stuffs.bid`), and the origin's own `robots.txt` still disallows
everything — a crawler only reads that file if it reached `topdf.stuffs.bid` directly, and
the right answer there is "not here."

**One CSS collision worth remembering:** the new site header was first written as
`nav.bar`, and `.bar` is already the export progress element — 3px tall, `overflow:hidden`,
`--panel-2` background. The header inherited the background and painted a stray band behind
itself on every page. It renders subtly enough that reading the screenshot missed it; a
pixel probe found it. The header is `nav.site`.

## The Guide (`/topdf/guide`)

Added 2026-09-06 with the public launch, rewritten 2026-09-07 for topdf's own five columns. Without it a stranger cannot succeed: the template has to satisfy a detector they cannot see.

### topdf's samples — `scripts/make_public_samples.py`

Its own palette (teal, `#0F5257`), its own products (workshop and site supplies, not the client's grocery catalog), its own templates and its own five columns. **Nothing is shared with the client samples, deliberately** — the whole reason the split exists is that the public product must not carry the client's look.

```bash
./.venv/bin/python scripts/make_public_samples.py
```

It **verifies as it goes**: every template it draws is run back through `simple_catalog.find_product_boxes` and the script fails rather than shipping a template the tool cannot read. It caught the header-band detection bug — the run printed `found 0 product boxes, expected 12` instead of publishing a guide whose own sample does not work.

### The client tool's samples — `scripts/make_samples.py`

Unchanged apart from its output path (`web/client_static/samples/`) and the sample image base URL, which now points at the client host. Re-run it after any change to the grid constants or the sample data. It verifies with `classify_page_layout` the same way.

Two things that took a round to get right there, both encoded in the script's comments:

- **Border weight.** The detector samples for ink ~2px inside each box edge on a 72dpi render. A hairline centred on the outline half-misses it. The strokes are drawn fully *inside* the outline at 3.2pt, so the box keeps its exact outer size and the ink lands where the detector looks.
- **The promotion template is three pages, and the plain one is two.** `P6_FIRST/MIDDLE/LAST_PAGE_ROWS` are three *different* row sets; a 2-page promotion template makes the exporter draw last-page products into rows the cover page has no boxes in. Its page 2 also carries a banner in the top-row band rather than boxes — a middle page is *recognised* by ink there, but the promotion layout never fills it.

## Single-page templates (2026-09-06)

A one-page template was already *accepted* — and produced a broken second page. `resolve_template_page_numbers` returns `(1, 1, 1)`, so the same page then got the fixed row set for each role in turn, and those row sets differ: a cover-shaped page used as the last page had products drawn at row 200, where it has a header rather than boxes. Three products floated over the branding and the bottom row came out empty.

**The fix is to stop assuming and measure.** `detect_grid_rows(page)` returns which of `REFERENCE_PAGE2_ROWS` — the superset every other row set is drawn from — actually have empty boxes on the page. `single_page_grid_rows()` applies it **only when the three page indices are identical**, and `get_template_slots(..., rows=…)` takes the override. Multi-page templates are deliberately untouched: real client templates are built against the fixed role sets, and re-deriving those from pixels would put working catalogs at the mercy of a detector.

A row counts only when its boxes are drawn **and still empty inside** (`slot_interior_clear_ratio`). Border sampling alone cannot tell a box from a solid banner, and both the promotion template's section banner and its dark back-page band sit exactly in a row's band — without the interior check, products would be laid over them.

**The same mismatch was live one layer up, in the 2-page path** (fixed straight after, once it showed up in the guide's own sample output). `resolve_template_page_numbers` returns `(first, middle, first)` for a 2-page template — the cover is the final page too — but the final page was still given `REFERENCE_LAST_ROWS`, whose first row is 200. On a cover page that band is the header, so the leftover products printed on top of the branding with no boxes around them, and the twelve real boxes stayed empty. `final_page_role()` now picks the role from the page rather than from the position in the catalog: when `last_index == first_index` the final page uses the cover's rows. Slot counts are unchanged (both row sets are 4×3), so the page-splitting decisions are identical — only the positions move. The footer-safe path is untouched and still correct on a cover page, because it whitens the band and redraws the grid itself.

Consequences worth knowing:
- Any number of rows works, in any position — a 3-row page was rejected before as `invalid` and is fine now.
- Single-page templates bypass the p6/plain row distinction entirely, so both card styles fit one.
- The footer-safe reflow is skipped (as it is for p6): every page of a single-page export has the same design, so the last page keeps the same rows.
- `single_page_grid_rows` returns `None` if it finds nothing, falling back to the role sets rather than producing a page with no slots — resolve-time detection runs at 72dpi and export-time at 150, and they should never disagree, but a fallback costs one line.

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

`logs/usage.jsonl` (gitignored), one line per export attempt: **which tool** (`topdf` or `client`), outcome, duration, template pages, output size, and whether it came via the proxy or direct. No IPs, no filenames, no content — deliberately too thin to answer anything except *does anyone use this*, which is the only reason the public tool exists.

**The `tool` field is the point of it now.** Before the split, the owner's own weekly client catalogs and a stranger's curiosity landed in the same counter, so the file could not answer the one question it was written for. Filter on `tool == "topdf"` for demand; `client` runs are work, not signal.

## System dependencies

`poppler-utils` (pdftoppm/pdfunite), `fonts-liberation`. Python deps in this project's own `.venv`.

## Notable fixes

- `try_font()` originally only checked Windows font paths. A Linux fallback to `/usr/share/fonts/truetype/liberation/` was added — without it text silently fell back to Pillow's bitmap default.
- (2026-08-17) "Generating..." used to stick forever even on success: `send_file(as_attachment=True)` downloads without navigating, so a plain form POST never reset the UI. Fixed by submitting via `fetch` and handling the blob in JS. The current page keeps that and adds an elapsed-seconds counter, because a minute of silence reads as a hang.
- (2026-08-23) CSV support via `_iter_rows_from_source()`, handling UTF-8-BOM/cp1252 and comma/semicolon delimiters. Verified in production on real client catalogs the same day.

## If it grows

The export is synchronous, and a second visitor during an export gets a worded 503 rather than a queue. That is the right trade at this traffic. If the usage log ever shows real concurrent demand, the engine already exposes `status_callback`/`progress_callback` — the upgrade is a job id plus polling, not a rewrite.
