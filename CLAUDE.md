# excel-to-pdf — two catalog tools, one process

> A Shopify store URL — or a product spreadsheet plus a template PDF — becomes a finished catalog PDF. **Two tools do that here, and keeping them apart is the point of the current design** — see the split below before changing anything.

## Before shipping a visual change

```
node ../groundwork/preflight/render.mjs https://stuffs.bid/topdf
```

**[groundwork](../groundwork)** is the shared foundation across products — a render preflight, dated vendor facts (Cloudflare, Search Console, Dodo), and one rules file per domain. Technique that is true of the next product goes there; anything about catalogs stays here.

Run 2026-09-11: the public page passes at phone, tablet and desktop. One finding, and it belongs to stuffboard rather than here — **an unknown path under the domain returns an error page that declares no `color-scheme`**, so Chrome on Android force-darkens it.


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
- Input: **either** a Shopify store URL (topdf reads the products and generates the template — see
  the Shopify section) **or** a product file (`.xlsx`/`.csv`) plus a template PDF. Output either way:
  the composed catalog PDF, streamed back as a download.
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
shopify_catalog.py             store URL -> products + branding + generated template -> catalog
brand_refiner.py               the one optional model call: which image is the logo
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
scripts/shopify_to_catalog.py  build a catalog from a store URL on the command line
scripts/make_public_samples.py regenerates web/static/samples end to end, and verifies it
scripts/make_samples.py        regenerates web/client_static/samples the same way
```

## The third engine — `profex_catalog.py` (2026-09-12)

**A store URL in, a *report* out** rather than a grid: a cover that leads with a number, a page of figures about the range, the products as cards, then the same lines as a grouped table with page numbers and a running brand.

```bash
./.venv/bin/python profex_catalog.py https://some-store.example -o catalogue.pdf --max 24
```

**The structure is not this repo's.** It comes from **profexpdf** — a separate, private report engine that knows about covers, figures, cards and grouped tables and nothing about shops. `profex_catalog.py` is the half that *is* ours: what counts as a product, what a price is, how a range is grouped, which four numbers earn a strip at the top, and every word on the page.

**The line, and it is the point of the split:** if a change would need profexpdf to learn the word "product", it belongs in `profex_catalog.py`. If it would need `profex_catalog.py` to learn the word "millimetre", it belongs in profexpdf. Nothing shop-shaped goes back upstream — that is the same rule `catalog_exporter.py` lives under, one level up.

`PROFEXPDF_HOME` points at that checkout; it defaults to a sibling directory named `profexpdf`.

**It is a command, not a hosted path, and that is deliberate.** profexpdf prints through headless Chrome, which peaks at a few hundred MB. This box is 2GB with `MemoryHigh=550M` on one gunicorn process that already peaks near 350MB for a reportlab export — a Chrome inside a request is the thing that must not happen here. Wiring it into `/topdf` needs a separate worker with its own memory budget, or a queue, first.

**What it needed from the store that the grid engine does not:** `product_type` and `vendor` (to group and to chart), the product handle (every name in the table links back to the page it was read from), and photos re-encoded through Pillow into local files, because the engine embeds local files or data URIs and never fetches a URL itself.

Two things this shook out, both fixed here rather than there:

- **`brand.tagline` is cut to a character count** for a template that has one line for it, which put *"…arabica and robusta co"* under a cover headline. `_sentence()` cuts at a sentence instead.
- **Price bands are computed from the catalogue, not fixed.** A fixed $0–50/50–100 scale puts every line of a candle shop in the first bucket and every line of a furniture shop in the last; quintile edges rounded to a figure a person would say keep the shape visible at any price point.

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

## Shopify: a store URL in, a catalog out (2026-09-08)

topdf asked for two files a stranger does not have: a product list in our five columns, and a
template PDF with empty boxes drawn on it. That is two jobs of work before the tool does any of
its own, and it is why the tool was not actually usable by the people it is aimed at.
`shopify_catalog.py` removes both by reading them off the store. `POST /topdf/from-shopify`, and
the form is now the first thing on the landing page.

**The pipeline is deterministic.** Everything on the page comes from Shopify's own public JSON:
`/meta.json` for shop name, city, country and `money_format`; `/products.json` for titles,
`body_html`, images and variant prices. No model touches the product data, the prices or the copy
— the description is the store's own first sentence, clipped on a word boundary, which keeps the
shop's voice and costs nothing. Junk items ("Free Returns Coverage", gift cards) are filtered on
`requires_shipping` rather than a name blocklist.

**One step resists rules, and only that step calls a model.** Deciding which of a dozen homepage
images is the brand's own logo is a judgement call, not a parsing problem, and it was measured
failing: Death Wish Coffee's page has four images tagged "logo" and three are press badges
(BuzzFeed, HuffPost, Yahoo); Tentree's first match is a Science Based Targets certification mark.
Each is fixable with one more rule, and the next store breaks the next rule. So `brand_refiner.py`
gets that shortlist and nothing else — Haiku, one call, ~800 tokens, a fraction of a cent.

**It is off by default and degrades to a complete catalog.** `read_brand(refine=None)` still
produces a finished PDF in a neutral palette; the refiner is a quality knob we can measure, not a
dependency in the middle of the pipeline. `scripts/shopify_to_catalog.py --compare` runs both ways
and diffs the branding, which is how the split below was found.

**The model picks the image; the pixels pick the colour.** Asked for both, the refiner named a
near-black for Death Wish Coffee — whose logo it had just correctly identified, and whose red is
right there in it. So `brand_colours_from()` quantizes the chosen logo and takes its most-used
genuinely saturated tone, and the refiner's colours are only the fallback for a monochrome mark
(Tentree, Allbirds). Judgement is what the model is for; measuring a colour is not.

**Generated templates do not go through the detector.** `simple_catalog.export_catalog` takes an
optional `known_boxes`, and this path passes the coordinates it drew at. Measuring your own drawing
is a category error — detection exists to cope with a *stranger's* PDF. It was not theoretical: the
logo is drawn at the left margin, the grid's left column starts at that same 42pt, and at 150dpi
both land on pixel column 87 — the detector lost that column and a third of every page's products
silently vanished. No tuning makes that class of accident impossible. Uploaded templates are
untouched and still measure.

**The template is sized to the products.** `plan_pages()` draws exactly as many boxes as there are
products (cover 9, then 12 a page), because a template *we* generate with nine empty boxes on the
back page is our defect, not the user's layout. The first run shipped exactly that.

**A JPEG logo gets a plate.** `_baked_background()` samples the border: a transparent PNG has no
background of its own and is composited onto the brand band, but a JPEG carries one (Shopify even
pads them, `pad_color=ffffff`) and dropping that on a dark band looks like a sticker. A rounded
plate in the logo's own background colour makes the same pixels read as a deliberate lockup.

**Not every Shopify-backed store is reachable.** A headless storefront — Shopify for checkout,
Next.js or Sanity for the site — serves neither endpoint (Gymshark 403s, Kotn 404s). That is
detected and reported as such, rather than failing obscurely. Stock themes are the vast majority
and they work.

**SSRF matters more here than it did for images.** The store URL is typed straight into a box on
the page, which is the most direct such handle this app has; `web/security.py` grew
`fetch_guarded`/`fetch_remote_document` so the JSON and HTML fetches get the same
resolve-then-connect vetting as product images. The guard's own wording never reaches the caller —
"port 8020 is not allowed" is an accurate answer to a port scan — so the reply says only that the
store did not answer.

### Scaling it (2026-09-08, same day)

A store URL that took 62s server-side was cut by nginx at 60s and reported to the user as a
timeout. Three things were wrong and all three are fixed; a fourth is still open.

**nginx matched the wrong location.** `location = /topdf/generate` is an *exact* match with
`proxy_read_timeout 300`; `/topdf/from-shopify` did not exist when that was written, so it fell
through to `location /` and its 60s. The block is now `~ ^/topdf/(generate|from-shopify)$`. The
work had actually finished — the usage log has `ok:true, ms:62344` for the run the user saw fail.

**Rendered template pages were held one per output page.** A 150dpi A4 page is 6.5MB, so a
1000-product catalog wanted 84 of them — 549MB against a `MemoryHigh` of 550M. `export_catalog`
now renders lazily into a cache keyed by template page, and `build_template` emits **at most three
physical designs** (cover, full inner, short final) with a `page_sources` map saying which design
each output page uses. Memory is now a function of distinct designs, not catalog length.

**Image downloads were serial and were ~90% of the wall clock.** `prefetch_images()` warms the
exporter's on-disk cache through a 12-thread pool before the render loop. It is a *prewarm*, not a
rewrite: the loop still reads the cache and still degrades a failed photo to a card without one.

Measured after all three, on a cleared cache: **999 products, 135s, peak RSS 168MB, 84 pages,
12.2MB.** 36 products went 62s → 7.4s.

**Still open: the run is synchronous.** 135s holds a browser tab, holds the single export slot
(everyone else gets a 503 meanwhile), and sits under gunicorn's 300s ceiling with no progress
reporting. That is the next thing, and it is a job model, not a bigger timeout.

### Three filter bugs the scale test exposed

Each silently removed real products, which is worse than failing:

- **`"return"` matched as a substring** against the concatenated type+tags blob hit the tag
  `loop::returnable => true`, which sits on nearly every product of a shop using the Loop returns
  app. It deleted **215 of Allbirds' 294 products**. `product_type` is now matched as a substring
  (Shopify writes `return,package_protection` there) and tags only as **whole tags**.
- **A zero price was treated as junk.** On a shipping product it means "price on request", which
  is normal B2B — **18 of Metrixplus Instruments' 97 items** are quoted that way. Price is no
  longer required; the card omits the line, as it does for a blank Price cell.
- **Colourways published as separate products with identical titles** filled pages with the same
  card. Tentree's 999 real products are 371 distinct titles. Deduped on normalised title, which
  leaves shops that put the colour *in* the title (Allbirds) untouched.

### Three more the artifact showed, that the data did not

Reading a generated catalog as an outsider — not the diff, not the logs:

- **Prices on an Indian store printed as £.** `money_format` is `Rs. {{amount}}`, which is not a
  symbol `simple_catalog.format_price` recognises, so it fell through to that module's `£`
  default. Currency now resolves from the **ISO code** in `meta.json` (`CURRENCY_BY_CODE`), and
  an unknown one prints nothing rather than guessing: a bare number reads as "ask us", the wrong
  symbol is a false claim about the price.
- **Then ₹ rendered as a tofu box.** Liberation Sans, which is what `try_font` finds on this box,
  has no U+20B9. `_font_can_render()` draws the character and compares it against `.notdef`, and
  an unrenderable symbol falls back to the shop's own wording (`Rs. `). Swapping the card font
  would have changed the typography of every catalog ever made here, and `try_font` is the client
  engine's besides.
- **Every Metrixplus description began "Product Overview".** That is an `<h2>` being flattened
  into the sentence. `BOILERPLATE_OPENERS` strips a leading heading. Preferring the first `<p>`
  was tried first and was wrong in the other direction — Death Wish Coffee's tagline is an `<h4>`
  ("Keep it under wraps.") and is the best line on the card. What separates them is whether the
  heading *says* anything, not which tag it is.

### Bounds

**`SHOPIFY_MAX_PRODUCTS = 250`**, with **`SHOPIFY_BUDGET_SECONDS = 240`** enforced by the build on
itself between phases and once a page. The cap is a choice about how long one visitor may hold the
single export slot, not about what the machine can do — 250 products measures at 28s and 999 at
135s. The budget matters because the run is synchronous: without it the only limit is gunicorn's
300s, and reaching *that* kills the worker mid-request and gives a dropped connection with no
explanation, which is the exact failure this route shipped with. Past the deadline `prefetch_images`
abandons the remaining downloads rather than cancelling them — each is a blocking socket read of up
to 12s, so the pool drains far faster — and those products keep their card and lose only the photo.

**Deliberately still synchronous** (decided 2026-09-08). A job model was scoped and declined: it
would mean holding finished PDFs on disk for the browser to collect, and "nothing is retained" is
a promise the page makes. As it stands the job directory is deleted in the request's own `finally`
and that promise is literally true. The cost is that one build blocks others for up to ~40s and
there is no progress reporting. If that becomes the complaint, the answer is a job model with an
explicit retention line on the page — not a bigger timeout.

The old wall-clock note, for reference: every product is one cold CDN download, and before the
concurrency and memory work 36 products took 62s over three pages. Product images are requested at `?width=700` and the logo at
600, which is the single biggest saving in the run. Same public quota as the upload path; usage
logs as `tool="topdf-shopify"` so this demand signal stays separable from the other two tools.

The Anthropic key lives in a gitignored `.env`, read by the service through `EnvironmentFile=-`
— the `-` matters, because the service must still start when it is absent.

## The landing page redesign (2026-09-08)

The page shipped as one long column at a single margin, with the tool behind a *Paste my store
address* button that scrolled somewhere else. A stranger arriving saw a wall of prose and a button
whose job was to reveal the product. Three things changed.

**The paster is the hero.** The store-address field is the first interactive thing on the page,
above the fold, and there is exactly one of it — the form moved out of `#make` into the hero rather
than being duplicated, because two `id="storeUrl"` would have broken the JS outright. `#make` is
now only the upload path, introduced as *"Not on Shopify?"*.

**The page has weather.** Two fixed layers on `body::before/::after`: three colour blooms for
depth, and a 58px grid masked to fade out before it reaches the content — without the mask it tiles
the viewport and reads as graph paper. Both are `position: fixed` (one paint, no scroll cost) and
`pointer-events: none`. Buttons, the number badges and the section kickers take a single
accent→`--accent-2` gradient; that violet is only ever a gradient partner, never a flat fill.

**Sections have air and a shape.** 104px apart, each opening with a gradient kicker over a real
headline, separated by a hairline that fades out at both ends. The old duplicate "how it works"
section — written for the upload flow — was removed rather than renamed; the new one describes the
paste-a-URL flow, and the guide still has the detail.

**The nav is full-bleed and sticky.** Constraining the whole `nav.site` to `.wrap` painted a
floating slab with hard edges against the background, which read as a rendering fault. The bar
spans the viewport and its contents sit in `.wrap {{ self.wrapclass() }}`, so they align with the
column on every page rather than only the landing one.

### Three defects a headless probe found that reading the page did not

Worth keeping the probe (`Runtime.evaluate` for duplicate ids, overflow and console errors, plus a
sweep of every `src`/`href` for non-200s) — each of these is invisible in a screenshot:

- **Two `id="how"`**, because the rewrite inserted a new section and the old one was below the
  replaced range.
- **`og:image` 404** — it pointed at `example-output-page1.png`, a *client* sample filename, so
  every shared link had a broken preview from the 2026-09-07 split onward. Now the hero export.
- **No favicon**, so every visit logged a `/favicon.ico` 404 and the browser tab was blank. Now an
  inline SVG data URI: no file to serve, no route, no request.

### The in/out carousel (2026-09-08)

"What comes out" was two before/after pairs stacked under headings like *"One template page,
repeated"*, and readers could not tell which picture was input and which was output. Eight images
on screen at once, and the jargon named the mechanism rather than the outcome.

It is now one carousel with **two** slides, each the same shape: **You give → You get**. The
repetition is the teaching device, and only one pair is on screen at a time, so the eye compares
two things instead of eight. Slide 1 is the Shopify path — its "give" is a mock of the hero's own
address field, because the input is a URL and there is no picture of one. Slide 2 is the upload
path.

It shipped with three, and a cover-page slide was cut: it showed the same mechanism as slide 2
(a template with empty boxes, filled) and reading them in sequence made the section drag rather
than clarify. The point survives as a sentence in slide 2's caption, which is where a variation on
a mechanism belongs once the mechanism is shown. The two slides that remain are the two *decisions*
a visitor actually makes — paste an address, or bring your own design.

No library: three slides in a flex track, one `translateX`, dots built from the slide list. Details
that are not obvious:

- **Off-screen slides get `inert`**, or Tab walks focus into content nobody can see.
- **The viewport's height follows the active slide.** A flex track stretches every slide to the
  tallest one, which on a phone left slide 1 — one address field and one page — with 340px of dead
  space beneath it. A `ResizeObserver` on the slides re-takes the height, because it is not final
  until the images decode.
- **The step is measured, not assumed** (`slides[0]` width + the 14px gap), so it is re-taken on
  resize with the transition suppressed for that one frame.
- Arrow keys are bound to the carousel element, not the document, so they do not fight the
  browser's own scrolling; swipe requires horizontal intent so a vertical drag still scrolls.

### The hero shot

`scripts/make_public_samples.py` grew `write_hero_shot()`, which builds the hero image through
**`shopify_catalog.build_template`** — the Shopify path's own generator — for an invented shop
(Northgate Supply Co.). The before/after samples still say "YOUR LOGO HERE" and should: the blank
template beside them is the thing you download and put your own logo on. In the hero that
placeholder was wrong twice over — the headline sells the Shopify flow, and a placeholder logo is
the clearest sign a page is a demo rather than a product. The shop is invented deliberately;
publishing a real store's catalog as our own marketing is not ours to do, however good it looks.

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
