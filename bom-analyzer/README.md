# BOM Supplier Analyzer

Upload a bill of materials and get **lead time, cost, stock availability and lifecycle status**
for every part, from **DigiKey** and **Mouser**, side by side in one table.

Each supplier gets its own column group, so you can see at a glance which one is cheaper for a
given line, which one can ship today, and which parts are heading for end of life.

---

## What it does

**Reads the BOM you already have.** CSV, TSV or Excel (`.xlsx`). Column headings are detected
automatically — `Qty`, `MPN`, `Mfr. Part #`, `RefDes` and the other spellings EDA tools emit all
map to the right field, and a title block above the header row does not confuse it. If a guess is
wrong you can remap any column from a dropdown without re-uploading.

**Looks each part up at both suppliers.** One query per distinct part number per supplier, run
with bounded parallelism and cached, so a 200-line BOM with repeated part numbers does not burn
through a free-tier quota.

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

**Exports the comparison.** One CSV with every field for every supplier, respecting the current
filter.

---

## Why there is a backend

The frontend is a static page, but it talks to a small Node server rather than to the suppliers
directly. Two reasons, both hard blockers:

- **Credentials.** DigiKey uses OAuth 2.0 client credentials and Mouser uses an API key. Either one
  shipped to a browser is public. They stay on the server.
- **CORS.** Neither API sends the headers a browser needs to read a cross-origin response, so a
  direct call from page JavaScript fails regardless of credentials.

The server is ~400 lines, has **no dependencies**, and needs nothing but Node 18 or newer.

---

## Setup

### 1. Get API credentials

Both are free. Either one alone is enough to start — the app just shows one supplier column.

| Supplier | Where | What you need |
| --- | --- | --- |
| DigiKey | [developer.digikey.com](https://developer.digikey.com/) | Create an organization and an app, add the **Product Information** API to it, then copy the **Client ID** and **Client Secret** |
| Mouser | [mouser.com/api-hub](https://www.mouser.com/api-hub/) | Request a **Search API** key (not the Order API key) |

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
npm start
```

Then open <http://localhost:8787>. The status pills under the title confirm which suppliers are
connected.

To try the interface before you have keys, click **Load sample BOM** — parsing and column mapping
work without any credentials; only the lookup needs them.

---

## Using it

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
| `LOOKUP_CONCURRENCY` | `3` | Parallel supplier requests — raising it is faster but invites rate limiting |
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
npm test     # 66 tests, no network access required
npm start    # http://localhost:8787
```

```
server.js              HTTP server, static hosting, API routes, SSE streaming
lib/spreadsheet.js     CSV/TSV parsing, .xlsx reading, header detection, column mapping
lib/digikey.js         OAuth 2.0 + Product Information V4 → catalog record
lib/mouser.js          Search API v1 → catalog record
lib/normalize.js       Lead time, lifecycle and price-break normalization; cross-supplier comparison
lib/lookup.js          Deduplication, caching, concurrency, BOM roll-up
lib/cache.js           TTL cache with disk persistence
lib/http.js            fetch with timeout, retry and backoff
public/                The frontend (no build step)
```

Supplier responses are converted into a common **catalog record** — the quantity-independent facts
about a product, including every packaging option with its own stock, minimum, multiple and price
ladder. Pricing for a specific BOM line happens afterwards, in `recordToOffer`. That split is what
lets the cache serve a re-run at a different quantity, and it keeps all the comparison logic in one
supplier-agnostic place.

The tests cover the parts that are easy to get quietly wrong: price break selection at a quantity,
minimum-order and packaging-multiple arithmetic, the lead time and lifecycle vocabularies of both
suppliers, packaging choice, header detection against four different BOM dialects, `.xlsx`
container reading, and the comparison verdicts. Supplier clients are tested against recorded
response shapes, so no network access is needed.

---

## Limitations

- **Stock and pricing move constantly.** Confirm on the supplier's own page before ordering; the
  detail panel links straight to it.
- **Lead time is what the supplier publishes**, which is usually the manufacturer's factory quote,
  not a promise for your order.
- **Matching is by manufacturer part number.** A part is only reported when the returned MPN is a
  credible match, and the detail panel says when the match was not exact — but verify anything
  surprising.
- **Two suppliers.** The architecture takes more: a client needs `id`, `name`, `configured` and
  `fetchRecord(part)` returning a catalog record, then gets added to the list in `server.js`.
  Everything downstream — comparison, roll-up, table columns, CSV export — is written against the
  supplier list rather than against DigiKey and Mouser by name.
