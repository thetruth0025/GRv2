# BOM Supplier Analyzer

Upload a bill of materials and get **lead time, cost, stock availability and lifecycle status**
for every part, from **DigiKey**, **Mouser** and **TrustedParts**, side by side in one table.

Each supplier gets its own column group, so you can see at a glance which one is cheaper for a
given line, which one can ship today, and which parts are heading for end of life.

TrustedParts is an aggregator rather than a distributor: it searches many authorized distributors
at once. Its column quotes whichever distributor is cheapest for your quantity, and expanding a row
lists **every** distributor it found, with stock, minimum order and price for each.

Two front ends over the same engine:

- **`python3 server.py`** — a browser app at <http://localhost:8787> for interactive work.
- **`python3 bom.py my-bom.csv -o comparison.xlsx`** — a CLI for scripting and one-shot runs.

---

## What it does

**Reads the BOM you already have.** CSV, TSV or Excel (`.xlsx`) — the `.xlsx` reader is `zipfile`
plus `xml.etree`, so no spreadsheet library is needed. Column headings are detected
automatically — `Qty`, `MPN`, `Mfr. Part #`, `RefDes` and the other spellings EDA tools emit all
map to the right field, and a title block above the header row does not confuse it. If a guess is
wrong you can remap any column from a dropdown without re-uploading.

**Looks each part up at every supplier.** One query per distinct part number per supplier, run
with bounded parallelism and cached, so a 200-line BOM with repeated part numbers does not burn
through a free-tier quota. TrustedParts accepts up to 50 parts per request, so that whole BOM
costs it four calls rather than two hundred.

**Puts the answers side by side.** For each supplier: stock on hand, lead time, unit price at your
quantity, extended price for the line, and lifecycle status. The cheapest and the soonest are
badged, and a verdict column says which supplier to use and why.

**Tells you what to worry about.** Obsolete and NRND parts, lines no supplier can cover, lines
where you would have to split the order across suppliers, and lead times past twelve weeks are all
called out per line and totalled at the top.

**Prices what you would actually buy.** Minimum order quantities and packaging multiples are
applied before pricing, and the correct price break is used for the resulting quantity. Where
DigiKey lists several packaging options, the one that can actually ship your quantity for the
lowest total cost is chosen — a reel with a lower unit price but a 5,000-piece minimum is not a
better way to buy 500.

**Exports the comparison.** One row per BOM line with every field for every supplier, respecting
the current filter. When an aggregator is in play the `.xlsx` gains a second **Distributors** sheet
carrying one row per distributor offer, and CSV gets a `-distributors` companion file.

### Attribution

TrustedParts require that a **publicly available** application displaying their API data shows the
words *"Powered by"* followed by their logo, linked back to trustedparts.com, and that the link is
followable — no `rel="nofollow"`, no temporary redirect. The app implements this:

- A visible attribution block sits with the results, and each expanded row carries a per-part link.
- Links use `rel="noopener"` only, deliberately **not** `nofollow`.
- The target follows their guidance: a single-part run links to that part's TrustedParts page (from
  the `Links` section of the response), a multi-part run links to their home page.

**One thing you must supply:** their logo. It is their trademark, so it is not bundled here —
see [`public/TRUSTEDPARTS-LOGO.md`](public/TRUSTEDPARTS-LOGO.md). Until you add
`public/trustedparts-logo.svg` the attribution renders as the linked words "Powered by
TrustedParts.com", which is visible attribution but not the logo their guidance asks for. Add the
file before deploying publicly. Running on localhost for your own use is not covered by the
requirement.

---

## Why there is a backend

The frontend is a static page, but it talks to a small Python server rather than to the suppliers
directly. Two reasons, both hard blockers:

- **Credentials.** DigiKey uses OAuth 2.0 client credentials and Mouser uses an API key. Either one
  shipped to a browser is public. They stay on the server.
- **CORS.** Neither API sends the headers a browser needs to read a cross-origin response, so a
  direct call from page JavaScript fails regardless of credentials.

The server is **pure standard library** — no `pip install`, no virtualenv, nothing to vendor. It
needs nothing but Python 3.8 or newer, which macOS and every Linux distribution already ship.

---

## Setup

### 1. Get API credentials

Both are free. Either one alone is enough to start — the app just shows one supplier column.

| Supplier | Where | What you need |
| --- | --- | --- |
| DigiKey | [developer.digikey.com](https://developer.digikey.com/) | Create an organization and an app, add the **Product Information** API to it, then copy the **Client ID** and **Client Secret** |
| Mouser | [mouser.com/api-hub](https://www.mouser.com/api-hub/) | Request a **Search API** key (not the Order API key) |
| TrustedParts | [trustedparts.com](https://www.trustedparts.com/) | Register, then request API access under **My Account → Additional Features**. The key arrives as *Trial* with an expiry — email `user-requests@trustedparts.com` to move it to *Active* |

DigiKey issues both Sandbox and Production keys. Sandbox returns fixed demo data, so use the
Production pair unless you are specifically testing against the sandbox.

### 2. Configure

```bash
cd bom-analyzer
cp .env.example .env
# edit .env and paste in the credentials you have
```

### 3. Run

```bash
python3 server.py          # web app on http://localhost:8787
python3 bom.py --help      # or use the command line
```

The status pills under the web app's title confirm which suppliers are connected.

To try the interface before you have keys, click **Load sample BOM** — parsing and column mapping
work without any credentials; only the lookup needs them.

---

## The command line

```bash
python3 bom.py samples/sample-bom.csv -o comparison.xlsx
```

```
Analyzing 10 parts from sample-bom.csv against DigiKey and Mouser…

ROW  PART                QTY  DIGIKEY STOCK  LEAD          EXT  MOUSER STOCK  LEAD         EXT  BEST     LIFECYCLE
  2  GRM188R71H104KA93D  300         21,450  in stock  $154.80        57,763  in stock  $85.50  Mouser   Obsolete
  3  RC0603FR-0710KL     500         17,902  in stock   $33.60        16,703  in stock $154.50  DigiKey  Active
  …

Summary
  Lines analyzed         10 across 1,100 units
  DigiKey cart           $317.54  (1 not carried)
  Mouser cart            $364.40  (2 not carried)
  Cheapest per line      $259.34  (saves $58.20 vs. single-sourcing)
  Stock risk             0
  Lifecycle risk         7

Needs attention
  • GRM188R71H104KA93D  Obsolete — find a replacement; 81.1% price spread between suppliers
  • STM32F103C8T6       Not Recommended for New Designs; Single source — only DigiKey carries it
```

The written `.xlsx` has a frozen header, autofilter, one column group per supplier and every
number stored as a number, so it sorts and totals correctly in Excel. `.csv` and `.json` come out
of the same column layout — `-f json` gives you the whole analysis for scripting.

**Options worth knowing:**

| Flag | What it does |
| --- | --- |
| `-o FILE` | Write the comparison (`.xlsx`, `.csv` or `.json`; `-f` overrides the extension) |
| `-b N` | Building N units — multiplies every BOM quantity by N before pricing |
| `-s digikey` / `-s mouser` / `-s trustedparts` | Query one supplier only (repeatable) |
| `--limit N` | Analyze just the first N parts, e.g. to sanity-check a big BOM |
| `--list-columns` | Show the detected headers and mapping, then exit |
| `--mpn-column COL` | Force a column by header name or 0-based index (one flag per field) |
| `--fail-on risk` | Exit non-zero when any line is flagged — for CI and scripts |
| `--no-cache`, `--clear-cache` | Bypass or empty the cached supplier answers |

Reading part numbers from stdin works too, one per line with an optional quantity:

```bash
printf 'STM32F103C8T6, 25\nLM358DR, 100\n' | python3 bom.py - -o quote.xlsx
```

If nothing is found, `--list-columns` is the fastest way to see what the header detection made of
your file and which flag to reach for:

```
  IDX  HEADER                           MAPPED TO
  0    Item
  1    Reference                        reference
  2    Qty                              quantity
  4    Manufacturer Part Number         mpn
```

---

## Using the web app

1. **Load your BOM** — drop a file on the upload area, or paste a list of part numbers with
   optional quantities (`GRM188R71H104KA93D, 100`).
2. **Check the columns** — confirm the detected mapping against the preview and fix anything wrong.
3. **Analyze** — progress streams back per query. Repeat runs come from cache and are near instant.

In the results table, the number in the **Stock** column is what that supplier holds in the
packaging option you would buy; `short` means it is below your quantity. **Lead time** shows
`In stock` when the supplier can ship now, with the factory lead time underneath for reference.
Click the arrow on any row to see supplier part numbers, packaging, minimum order quantities, the
full price break ladder with your break highlighted, and links to the product page and datasheet.

---

## Configuration

All settings live in `.env`; `.env.example` documents each one. The ones worth knowing about:

| Variable | Default | Purpose |
| --- | --- | --- |
| `CACHE_TTL_HOURS` | `6` | How long a supplier answer is reused before being re-fetched |
| `LOOKUP_CONCURRENCY` | `3` | Parallel supplier requests (worker threads) — raising it is faster but invites rate limiting |
| `TRUSTEDPARTS_DISTRIBUTORS` | *(all)* | Comma-separated distributor names to restrict TrustedParts to |
| `TRUSTEDPARTS_IN_STOCK_ONLY` | `false` | Return only distributors holding stock |
| `TRUSTEDPARTS_USE_CACHED_DATA` | `false` | Use TrustedParts' cached data instead of real-time distributor feeds |
| `MAX_PARTS_PER_REQUEST` | `500` | Largest BOM accepted in one run |
| `ALLOWED_ORIGINS` | `*` | Restrict which origins may call the API if you host the frontend separately |
| `CACHE_FILE` | `./.cache/parts.json` | On-disk cache location, or `none` to keep it in memory |

### Quotas

Free tiers are limited (both suppliers cap a free key in the low thousands of calls per day). The
app minimises calls in three ways: duplicate part numbers across BOM lines collapse into one query,
answers are cached on disk for `CACHE_TTL_HOURS`, and the cache stores a quantity-independent
catalog record — so re-running the same BOM at a different quantity reprices without any new API
calls. Failed lookups are never cached, so a rate limit does not poison the next run.

### Hosting the frontend separately

`public/` is a static PWA and can be served from anywhere. Set the backend URL in the app's
Settings panel, and set `ALLOWED_ORIGINS` on the server to the origin serving the page.

---

## Development

```bash
python3 -m unittest discover -s tests -t .   # 145 tests, no network access required
python3 server.py                            # http://localhost:8787
python3 bom.py samples/sample-bom.csv        # the CLI
```

```
server.py                 HTTP server, static hosting, API routes, SSE streaming
bom.py                    Command-line front end: argument parsing, terminal output
bomlib/spreadsheet.py     CSV/TSV parsing, .xlsx reading, header detection, column mapping
bomlib/digikey.py         OAuth 2.0 + Product Information V4 → catalog record
bomlib/mouser.py          Search API v1 → catalog record
bomlib/trustedparts.py    Inventory API v2 (aggregator, batched) → catalog record
bomlib/normalize.py       Lead time, lifecycle and price-break normalization; cross-supplier comparison
bomlib/lookup.py          Deduplication, caching, concurrency, BOM roll-up
bomlib/cache.py           TTL cache with disk persistence
bomlib/http_client.py     urllib with timeout, retry and backoff
bomlib/report.py          One column layout shared by the CSV, .xlsx and terminal output
bomlib/xlsx_writer.py     Minimal styled .xlsx writer (zipfile + hand-built XML)
public/                   The web frontend (plain HTML/CSS/JS, no build step)
```

`server.py` and `bom.py` are both thin shells over `bomlib/`: they parse input, call
`LookupService`, and render. Neither knows anything about DigiKey or Mouser specifically.

The server is `ThreadingHTTPServer`; supplier lookups run on a `ThreadPoolExecutor` bounded by
`LOOKUP_CONCURRENCY`. Since the work is entirely network-bound, threads sidestep the GIL and the
async machinery alike.

Supplier responses are converted into a common **catalog record** — the quantity-independent facts
about a product, including every packaging option with its own stock, minimum, multiple and price
ladder. An aggregator fits the same shape: each distributor's listing becomes a packaging option
tagged with the distributor's name, so the one total-order-cost rule ranks across distributors and
packagings at once, and the per-distributor breakdown is just that list regrouped. Pricing for a specific BOM line happens afterwards, in `record_to_offer`. That split is what
lets the cache serve a re-run at a different quantity, and it keeps all the comparison logic in one
supplier-agnostic place.

The tests cover the parts that are easy to get quietly wrong: price break selection at a quantity,
minimum-order and packaging-multiple arithmetic, the lead time and lifecycle vocabularies of both
suppliers, packaging choice, header detection against four different BOM dialects, `.xlsx`
container reading, `.xlsx` writing, the CSV formula-injection guard, the CLI's column overrides
and exit codes, and the comparison verdicts. Supplier clients are tested against recorded response
shapes, so no network access is needed.

---

## Limitations

- **Stock and pricing move constantly.** Confirm on the supplier's own page before ordering; the
  detail panel links straight to it.
- **Lead time is what the supplier publishes**, which is usually the manufacturer's factory quote,
  not a promise for your order.
- **Matching is by manufacturer part number.** A part is only reported when the returned MPN is a
  credible match, and the detail panel says when the match was not exact — but verify anything
  surprising.
- **TrustedParts reports no lead time.** Its API carries stock but no lead time field, so that
  column reads as unknown for TrustedParts rather than guessing. Lines it stocks still compare
  correctly, because stock on hand outranks any quoted lead time.
- **TrustedParts lifecycle is a risk rating, and gated.** `LifecycleRisk` and `SupplyChainRisk` are
  only populated if your account is approved for them, and they are *risk grades*, not statuses. A
  grade like "Low" is shown as a risk field and deliberately does **not** drive the lifecycle
  column — calling a part obsolete on the strength of a risk grade would assert something the API
  never said. Text that genuinely names a status ("Obsolete") is promoted.
- **The attribution `Links` section is undocumented.** TrustedParts' attribution guide describes a
  `Links` array of `{Key, SearchToken, Manufacturer, Url}` with one entry keyed `Primary`, but their
  published OpenAPI schema does not include it. It is read defensively — from the response root or a
  part result, tolerating that shape and the `{Type, Url}` shape the schema uses elsewhere — and its
  absence simply falls back to the part's `ProductUrl`, then their home page. If the live API turns
  out to place it somewhere else, that is the code to adjust.
- **Adding more suppliers.** A client needs `id`, `name`, `configured` and `fetch_record(part)`
  returning a catalog record, then gets added to the list in `server.py` and `bom.py`. Set
  `batch_size` and provide `fetch_records(parts)` instead to have it queried in batches. Everything
  downstream — comparison, roll-up, table columns, exports — is written against the supplier list
  rather than against any supplier by name.
