# Catalog PDF Exporter

Turns a product spreadsheet (`.xlsx` or `.csv`) plus a template PDF into a finished,
branded catalog PDF — one product per box, across as many pages as the list needs.

Live at **https://stuffs.bid/topdf**, with a guide and downloadable templates at
**https://stuffs.bid/topdf/guide**.

- `catalog_exporter.py` — the engine (layout detection, image placement, PDF composition)
- `export_pdf.py` — command-line entry point
- `app.py` + `web/` — the hosted app

See `CLAUDE.md` for architecture, deployment and the reasoning behind both.
