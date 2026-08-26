# BOM Supplier Analyzer

Upload a bill of materials and get **lead time, cost, stock availability and lifecycle status**
for every part, from **DigiKey**, **Mouser** and **TrustedParts**, side by side in one table.

Each supplier gets its own column group, so you can see at a glance which one is cheaper for a
given line, which one can ship today, and which parts are heading for end of life.

TrustedParts is an aggregator rather than a distributor: it searches many authorized distributors
at once. Its column quotes whichever distributor is cheapest for your quantity, and expanding a row
lists **every** distributor it found, with stock, minimum order and price for each.

**A lookup form, not just an uploader.** Fill in a part number, an optional description and a
quantity, and press Enter to compare it across all three suppliers — no file, no column mapping. As
many rows as you like, and a column pasted straight out of a spreadsheet fills them for you.

**Several BOMs at once.** Load as many as you like — each gets its own tab, its own results, and
its own filter and search state, so switching between them restores exactly the view you left.
Results stay separate per BOM, but a part number is only ever looked up once: whichever BOM you
analyze first keeps it, and the rest report it as already covered instead of spending a second API
call on the same answer.

Two front ends over the same engine:

- **`python3 server.py`** — a browser app at <http://localhost:8787> for interactive work.
- **`python3 bom.py my-bom.csv -o comparison.xlsx`** — a CLI for scripting and one-shot runs;
  `python3 bom.py --part STM32F103C8T6` answers a single part without a file.

---

## What it does

**Reads the BOM you already have.** CSV, TSV or Excel (`.xlsx`) — the `.xlsx` reader is `zipfile`
plus `xml.etree`, so no spreadsheet library is needed. Column headings are detected
automatically — `Qty`, `MPN`, `Mfr. Part #`, `RefDes` and the other spellings EDA tools emit all
map to the right field, and a title block above the header row does not confuse it. If a guess is
wrong you can remap any column from a dropdown without re-uploading.

**Cleans the cells on the way in.** Part numbers arrive padded: indented in the cell, copied out of
a datasheet PDF, exported from an ERP. Spaces and tabs either side are trimmed, and so are the
characters that only *look* like spaces — zero-width space, byte-order mark, soft hyphen, word
joiner, the bidi marks. Those are the ones that matter, because Unicode does not classify them as
whitespace, so nothing strips them and they are invisible in a spreadsheet: you cannot find them to
delete, and the part matches nothing. A padded part number and a clean one are then the same part,
which is the difference between one line at the full quantity and two lookups that each half-price
the buy.

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

**Looks each part up once, across every BOM.** Load five boards that share a decoupling capacitor
and it is quoted once, on whichever BOM you analyze first; the others report it as already covered
rather than paying for the lookup again. The same part on two lines of one BOM becomes a single
line with the quantities added and the reference designators joined, because that is one purchase.

**Honours the BOM's own skip column.** If a column called *Skip to Production* (or *Skip*, *Do Not
Source*, and the usual spellings around them) is present, lines marked `YES` in it are never sent to
a supplier. Whoever wrote the BOM decided that deliberately, which beats any rule here — so it is
the first thing checked, and the reason quotes the cell.

**Ignores your own part numbers.** Assemblies, cables, drawings and bare boards are not distributor
stock, so part numbers starting with `ASY0`, `CBL0`, `DES0` or `PCB0` never reach a supplier. The
list is configurable, nothing is silently dropped, and every skipped line stays visible with the
reason it was skipped.

**Finds alternatives to a part in trouble.** Nexar (Altium) answers a different question from the
three suppliers: not what a part costs, but what could be used instead. It runs only for parts the
comparison has already found to be obsolete, NRND, end of life or simply unavailable — and only for
the ones you tick — so a free-tier quota is spent on the parts that need it rather than on every
line of a healthy BOM. What it finds fills the DMSMS form's Suggested Replacement column.

**Fills a DMSMS form per program.** Parts whose supply is ending — obsolete, discontinued, end of
life, last time buy, NRND — are collected across every analyzed BOM into a picker. Tick the ones
belonging to a program, name it, and out comes a DMSMS case form as Excel. A part can sit on three
boards and belong to one program, so which parts go on which form is yours to decide, not a rule.

**Produces a report you can hand to somebody.** The **Summary report** button opens a one-page
view: the headline numbers, what each supplier's cart would cost, the lines that need a decision,
and a concise per-part table. It prints to PDF, or exports to a styled Excel workbook — one BOM or
all of them at once.

**Exports the comparison.** One row per BOM line with every field for every supplier, respecting
the current filter. The `.xlsx` opens on a **Report** sheet, followed by a concise **Parts** sheet,
the full **Comparison**, a **Distributors** sheet when an aggregator is in play, and a **Skipped**
sheet listing whatever was screened out. CSV gets the same tables as companion files.

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

All are free to start. Any one supplier alone is enough — the app just shows fewer columns.

| Supplier | Where | What you need |
| --- | --- | --- |
| DigiKey | [developer.digikey.com](https://developer.digikey.com/) | Create an organization and an app, add the **Product Information** API to it, then copy the **Client ID** and **Client Secret** |
| Mouser | [mouser.com/api-hub](https://www.mouser.com/api-hub/) | Request a **Search API** key (not the Order API key) |
| TrustedParts | [trustedparts.com](https://www.trustedparts.com/) | Register, then request API access under **My Account → Additional Features**. The key arrives as *Trial* with an expiry — email `user-requests@trustedparts.com` to move it to *Active* |
| Nexar *(optional)* | [nexar.com](https://nexar.com/) | Create an application and copy its **Client ID** and **Client Secret**. Only needed for **Find alternatives** — the rest of the app works without it |

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

The written `.xlsx` opens on a **Report** sheet — title, headline numbers, what each supplier's
cart costs and the lines that need a decision — followed by a concise **Parts** sheet, the full
**Comparison** with one column group per supplier, a **Distributors** sheet when an aggregator is
in play, and a **Skipped** sheet listing anything screened out. Every sheet after the report has a
frozen header and autofilter, and every number is stored as a number, so it sorts and totals
correctly in Excel. `.csv` writes the same tables as companion files beside it; `-f json` gives you
the whole analysis, screened-out lines included, for scripting.

**Options worth knowing:**

| Flag | What it does |
| --- | --- |
| `-o FILE` | Write the comparison (`.xlsx`, `.csv` or `.json`; `-f` overrides the extension) |
| `-b N` | Building N units — multiplies every BOM quantity by N before pricing |
| `-p MPN` / `--part MPN` | Look up a part number directly, no BOM file (repeatable; `--part "STM32F103C8T6,25"` sets a quantity) |
| `-s digikey` / `-s mouser` / `-s trustedparts` | Query one supplier only (repeatable) |
| `--limit N` | Analyze just the first N parts, e.g. to sanity-check a big BOM |
| `--ignore-prefix PREFIX` | Skip part numbers starting with PREFIX (repeatable); replaces the `ASY0`/`CBL0`/`DES0`/`PCB0` default |
| `--no-ignore-prefixes` | Look up in-house part numbers too |
| `--no-merge-duplicates` | Keep repeated part numbers as separate lines instead of adding their quantities |
| `--ignore-skip-column` | Look up lines the BOM marks YES in a skip-to-production column |
| `--show-skipped` | List every skipped line and why |
| `--dmsms FILE --program NAME` | Write a DMSMS case form for every at-risk part |
| `--dmsms-status STATUS` | Narrow the form to given lifecycle statuses (repeatable) |
| `--alternatives` | Ask Nexar what could replace each at-risk part, and fill in Suggested Replacement |
| `--list-columns` | Show the detected headers and mapping, then exit |
| `--mpn-column COL` | Force a column by header name or 0-based index (one flag per field) |
| `--fail-on risk` | Exit non-zero when any line is flagged — for CI and scripts |
| `--no-cache`, `--clear-cache` | Bypass or empty the cached supplier answers |

For a one-off question there is no need for a file at all:

```bash
python3 bom.py --part STM32F103C8T6
python3 bom.py --part "STM32F103C8T6,25" --part "LM358DR,100" -o quote.xlsx
```

As in the web app's search box, a part named with `--part` is looked up as asked — the in-house
prefixes do not apply unless you pass `--ignore-prefix` yourself.

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

### Looking a part up

Fill in a row and press **Enter**. That is the whole flow — no file, no column mapping. The results
land in the same comparison table, with the same filters, report and exports as a BOM.

Each row has three fields:

| Field | |
| --- | --- |
| **Part number** | The only one that is required, and the only one the suppliers see |
| **Description** | Yours to carry through to the results, the report and the exports. It does not affect matching — the suppliers are searched on the part number alone |
| **Qty needed** | Drives the price break, the extended price and whether stock covers you. Defaults to 1, which is enough to compare unit price, stock and lifecycle status |

Rows are added as you fill them, and **+ Add part** and the **×** on each row cover the rest.
Pasting a column into a part-number box spreads it across the rows: a tab-separated copy out of a
spreadsheet lands in the right fields, and which value is which is worked out from what it looks
like — the number is the quantity, the first thing that is not a number is the part, and the rest is
the description. What landed is then sitting in the grid where you can see and correct it.

Searches share one tab, replacing it each time, and the rows stay filled in — so adding another part
to the comparison is a row and Enter away. Loaded BOMs are untouched by this.

**A lookup is answered as asked.** Neither the in-house prefixes nor another BOM's claim applies to
a part number you typed yourself: screening is there to stop automatic waste, not to answer a direct
question with nothing. Looking up `ASY0-1234` looks up `ASY0-1234`.

### Analyzing a BOM

1. **Load your BOMs** — drop one or more files on the upload area; multi-select works. Each file
   becomes its own tab; a file that cannot be read gets a tab saying so rather than aborting the
   batch.
2. **Check the columns** — confirm the detected mapping against the preview and fix anything wrong.
   The mapping belongs to the BOM you are viewing. A *Skip to Production* column is detected like
   any other, appears in the preview, and can be pointed at a different column or unmapped
   entirely — and the summary line says how many lines it will remove before you spend a call on
   them.
3. **Analyze** — either the BOM on screen, or **Analyze all pending** to work through the loaded
   BOMs one after another. Progress streams back per query, and each tab shows its own line count,
   cheapest-mix total and how many lines need review. Repeat runs come from cache and are near
   instant.

Switching tabs swaps the whole view — summary tiles, table, filters, search box and expanded rows —
to that BOM. Results, exports and totals are always scoped to the BOM you are looking at; the CSV
export is named after it. Closing a tab discards only that BOM.

In the results table, the number in the **Stock** column is what that supplier holds in the
packaging option you would buy; `short` means it is below your quantity. **Lead time** shows
`In stock` when the supplier can ship now, with the factory lead time underneath for reference.
Click the arrow on any row to see supplier part numbers, packaging, minimum order quantities, the
full price break ladder with your break highlighted, and links to the product page and datasheet.

The header rows stay put while the parts scroll under them, the part number stays pinned to the
left, and **Verdict** stays pinned to the right, so the column the comparison builds up to is
never the one you have to go looking for. Drag the bottom edge of the table to make it taller or
shorter.

Above the table, a line reports anything that was not looked up — lines the BOM marked skip to
production, in-house part numbers, lines merged into another line, and parts an earlier BOM already
claimed. Open it to see each one and why.

### The skip-to-production column

A line is skipped when that column holds any of the values a spreadsheet uses for yes:

```
YES   Y   TRUE   T   1   X   ✓   ✔        (any case, spacing ignored)
```

Everything else leaves the line in — blank, `NO`, `N/A`, `-`, a date, a note, a half-answer like
`yes if late`. Dropping a part on the strength of a value nobody recognised is the wrong way to be
wrong, so the rule only fires on an unambiguous yes.

The mark beats the in-house prefix rule, so a line that is both reports the mark: it is the more
useful reason. A marked line also claims nothing, so if another loaded BOM uses the same part it is
still free to look it up. `--ignore-skip-column` on the CLI turns the rule off, and unmapping the
column in the browser does the same.

Typed lookups ignore the column entirely, as they ignore every other screening rule: a part number
you entered by hand is a direct question, and it gets answered.

### Finding alternatives

**Find alternatives** lists every analyzed part worth replacing — obsolete, discontinued, end of
life, NRND, last time buy, plus any part no supplier carries or nobody holds the quantity for —
with the reason beside it. The ones that are gone are ticked; a part that is merely short on stock
is listed and left for you, because that is a judgement call.

Tick what you want asked about and the selected parts go to Nexar in one request. Each comes back
with the part Nexar matched — shown on purpose, since a suggestion is only worth as much as the
match it came from — and its alternatives, with manufacturer, stock, median price, factory lead time
and the key specifications, so you can see *why* something is being offered rather than taking the
word for it. Answers are cached like every other lookup, so asking twice is free.

What it finds flows into the DMSMS form's **Suggested Replacement** column, which is the point of
going looking. On the command line, `--alternatives` does the same for every at-risk part:

```bash
python3 bom.py my-bom.csv --dmsms falcon-ii.xlsx --program "Falcon II" --alternatives
```

Nexar needs its own credentials (`NEXAR_CLIENT_ID`, `NEXAR_CLIENT_SECRET`); the button says so if
they are missing. It is never called during a BOM analysis.

### The DMSMS form

**DMSMS form** collects every at-risk part across all analyzed BOMs — obsolete, discontinued, end of
life, last time buy and NRND — and lists them grouped by the BOM they came from, with the part
number, reference designators, quantity, status, distributor stock and a suggested risk.

Parts that are actually gone (obsolete, discontinued, end of life) start ticked; the ones still
buyable are listed but left for you. Tick what belongs to the program, fill in the case details, and
**Generate form** downloads the workbook. Everything except the program name is remembered, so the
next program is a name, a set of ticks and a click — the program is deliberately never pre-filled,
because that is how the wrong name ends up on a form.

The workbook is a DMSMS case form laid out on the SD-22 case data elements:

| Filled from the analysis | Left blank for you |
| --- | --- |
| Part number, manufacturer, description, reference designators | CAGE code |
| Next higher assembly, quantity per assembly | Lifetime buy quantity |
| Lifecycle status, **which supplier reported it**, date obtained | Last-time-buy date |
| Distributor stock, preferred source, unit and extended price | Resolution option |
| Suggested replacement, suggested risk | Disposition / remarks |

The blank columns are decisions, not lookups, so the form leaves them empty rather than inventing
them. **Status Source** matters: the analyzer reports whichever supplier gives a part its worst
standing, so a part DigiKey calls obsolete is never shown as active because Mouser has not caught
up — and naming the supplier is what makes the entry checkable. **Suggested Risk** is derived from
the reported status and whether distributor stock covers the quantity; the form says on its face
that it is a starting point, not a determination.

From the command line, `--dmsms` writes the same form with every at-risk part on it:

```bash
python3 bom.py my-bom.csv --dmsms falcon-ii.xlsx --program "Falcon II"
python3 bom.py my-bom.csv --dmsms obsolete-only.xlsx --program "Osprey" --dmsms-status Obsolete
```

There is no picking on the command line — that is what the browser is for.

### The summary report

**Summary report** opens a one-page view of the BOM you are looking at:

* the headline numbers — lines, best-mix total, stock risk, lifecycle risk, not found, skipped;
* what each supplier's whole cart would cost, and what splitting the order line by line saves;
* every line that needs a decision, with the reason;
* a concise per-part table: quantity, which supplier to buy from, unit and extended price, lead
  time and lifecycle status — plus, when several BOMs are open, what the others need of the same
  part;
* everything that was skipped.

**Export Excel** downloads it as a styled workbook. With more than one BOM analyzed, **Excel · all
BOMs** puts every one of them in a single workbook with the tabs prefixed by BOM name.
**Print / PDF** prints the report on its own, on white, without the rest of the app.

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
| `IGNORE_PART_PREFIXES` | `ASY0,CBL0,DES0,PCB0` | Part-number prefixes treated as in-house and never sent to a supplier. Empty value looks up everything |
| `MAX_PARTS_PER_REQUEST` | `500` | Largest BOM accepted in one run, counted **after** screening |
| `ALLOWED_ORIGINS` | `*` | Restrict which origins may call the API if you host the frontend separately |
| `CACHE_FILE` | `./.cache/parts.json` | On-disk cache location, or `none` to keep it in memory |

### Quotas

Free tiers are limited (both suppliers cap a free key in the low thousands of calls per day). The
app minimises calls in six ways: lines the BOM marks skip to production are never sent, in-house
part numbers never leave the building, repeated part numbers within a BOM collapse into one line, a part already quoted for another loaded BOM is not
quoted again, answers are cached on disk for `CACHE_TTL_HOURS`, and the cache stores a
quantity-independent catalog record — so re-running the same BOM at a different quantity reprices
without any new API calls. Failed lookups are never cached, so a rate limit does not poison the
next run.

### Hosting the frontend separately

`public/` is static and can be served from anywhere. Set the backend URL in the app's Settings
panel, and set `ALLOWED_ORIGINS` on the server to the origin serving the page.

Serve `index.html`, `app.js` and `styles.css` with the same caching policy — `server.py` sends
`no-cache` for all three. They are one unit, and caching them for different lengths of time is how a
browser ends up running last week's script against this week's page. If that ever happens the app
says so and offers to clear the old copy, rather than dying silently.

### If the page does not respond

A page whose script never started looks like a page with a slow backend: the status pill sits on
"Checking backend…" and no button does anything. `index.html` watches for that — it is the one file
a stale cache cannot hide — and puts up a notice with a button that clears the old copy and reloads.
**Ctrl+Shift+R** (**Cmd+Shift+R** on a Mac) does the same thing by hand, as does **Clear browser
copy** in Settings.

Versions before this one installed a service worker that caused exactly this after every update.
Opening the app once removes it for good.

---

## Development

```bash
python3 -m unittest discover -s tests -t .   # 301 tests, no network access required
python3 server.py                            # http://localhost:8787
python3 bom.py --part STM32F103C8T6          # look one part up
python3 bom.py samples/sample-bom.csv        # or a whole BOM
```

```
server.py                 HTTP server, static hosting, API routes, SSE streaming
bom.py                    Command-line front end: argument parsing, terminal output
bomlib/spreadsheet.py     CSV/TSV parsing, .xlsx reading, cell cleaning, header detection, column mapping
bomlib/digikey.py         OAuth 2.0 + Product Information V4 → catalog record
bomlib/mouser.py          Search API v1 → catalog record
bomlib/trustedparts.py    Inventory API v2 (aggregator, batched) → catalog record
bomlib/nexar.py           Nexar GraphQL: what could be used instead of a part in trouble
bomlib/normalize.py       Lead time, lifecycle and price-break normalization; cross-supplier comparison
bomlib/prepare.py         Screening: the BOM's skip column, in-house prefixes, duplicates, parts another BOM owns
bomlib/lookup.py          Deduplication, caching, concurrency, BOM roll-up
bomlib/cache.py           TTL cache with disk persistence
bomlib/http_client.py     urllib with timeout, retry and backoff
bomlib/report.py          The report and comparison layouts shared by CSV, .xlsx and terminal output
bomlib/dmsms.py           The DMSMS case form: what is at risk, and what the analyst still decides
bomlib/xlsx_writer.py     Minimal styled .xlsx writer (zipfile + hand-built XML)
public/                   The web frontend (plain HTML/CSS/JS, no build step)
public/sw.js              A service worker that exists only to uninstall an earlier one
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
and exit codes, the comparison verdicts, screening — that a prefix only counts at the start of
a part number, that `ASY1` is not `ASY0`, that merging adds quantities without mutating the caller's
lines, and that a BOM of nothing but in-house numbers costs zero API calls — and cell cleaning,
including every flavour of padding through both readers and the fact that a padded part and a clean
one merge rather than splitting in two. Supplier clients are tested against recorded response
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
- **Which BOM owns a shared part depends on the order you analyze them.** A part quoted for one BOM
  is not quoted again for the others, so its price and lead time reflect the quantity of whichever
  BOM you ran first. The report names the other BOMs and what they need, but it does not reprice
  the line at the combined quantity — buying the total in one order would usually be cheaper still.
- **The Nexar query has not been run against a live schema.** It was written from Nexar's Supply
  schema in a sandbox with no route to `api.nexar.com`, so the field names are documented rather
  than verified. GraphQL fails loudly — a wrong field comes back as *"Cannot query field X"* — and
  that message is shown as-is rather than swallowed, so a mismatch is visible immediately. Point
  `NEXAR_QUERY_FILE` at your own query to fix one without waiting for a code change.
- **An alternative is a suggestion, not a substitute.** Nexar's `similarParts` is matched on
  attributes, not on your circuit. The specs travel with each suggestion so you can judge it, and
  nothing here has been checked against your board.
- **The DMSMS form reports, it does not adjudicate.** Lifecycle status is what a distributor's API
  said today, which can lag a manufacturer's PCN and can differ between distributors for the same
  part. The form names its source and the date so the entry can be checked; confirm against the
  manufacturer before a case turns into a buy.
- **The skip column only understands yes.** It fires on `YES`/`Y`/`TRUE`/`T`/`1`/`X`/`✓` and
  nothing else, so a column using some other convention will be read as "source everything". Check
  the summary line before analyzing: it says how many lines the column will remove.
- **Cell cleaning keeps one interior space.** A run of spaces inside a value collapses to one, but a
  single space between characters is left alone: a few real part numbers carry one, and guessing
  which is which would break more than it fixed.
- **Screening is by prefix, not by meaning.** `ASY0`, `CBL0`, `DES0` and `PCB0` are assumed to be
  in-house numbering. If a real manufacturer part number starts with one of those, set
  `IGNORE_PART_PREFIXES` to something narrower, clear it to let everything through, or just look the
  part up directly — a hand-entered lookup is never screened.
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
