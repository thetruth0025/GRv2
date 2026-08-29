# BOM Supplier Analyzer

Upload a bill of materials and get **lead time, cost, stock availability and lifecycle status**
for every part, from **DigiKey**, **Mouser**, **TrustedParts** and **Nexar**, side by side in one
table.

Each supplier gets its own column group, so you can see at a glance which one is cheaper for a
given line, which one can ship today, and which parts are heading for end of life.

TrustedParts and Nexar are aggregators rather than distributors: each searches many authorized
sellers at once. Their columns quote whichever seller is cheapest for your quantity, and expanding
a row lists **every** seller they found, with stock, minimum order and price for each.

**A lookup form, not just an uploader.** Fill in a part number, an optional description and a
quantity, and press Enter to compare it across every configured supplier — no file, no column
mapping. As
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

**Matches on the part number, not on something like it.** A keyword search returns neighbours: ask
for `LM358` and `LM358DR` comes back, which is a different device in a different package. Only the
number asked for counts — case and punctuation are folded, because `RC0603FR-0710KL` and
`RC0603FR0710KL` are one part, but nothing else is. When a supplier's nearest answer is rejected the
table says so and names it, so a part that is merely spelled differently never reads as one nobody
carries.

**Looks each part up at every supplier.** One query per distinct part number per supplier, run
with bounded parallelism and cached, so a 200-line BOM with repeated part numbers does not burn
through a free-tier quota. TrustedParts accepts up to 50 parts per request, so that whole BOM
costs it four calls rather than two hundred.

**Puts the answers side by side.** For each supplier: stock on hand, lead time, unit price at your
quantity, extended price for the line, and lifecycle status. The cheapest and the soonest are
badged, and a verdict column says which supplier to use and why.

**Tells you what to worry about, and only that.** Obsolete and NRND parts, lines whose combined
stock across every supplier falls short of what the build needs, and lead times past twelve weeks
are called out per line and totalled at the top. A line no *single* supplier can cover is not one
of them, and neither is a part nobody stocks but everybody can order in a fortnight — see below.

**Works out how many to order from each supplier.** Four suppliers holding 80 each cover a need for
200 between them; the tool says who to buy how many from, and what the split costs.

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

**Prices the alternates your BOM already approved.** If a column names approved substitutes —
*Alternate Part Number*, *Alt P/N*, *Second Source*, and the usual spellings — each one is looked up
alongside the primary and reduced to the answer you need: could it cover this line today. Somebody
with the schematic in front of them already decided those parts fit, so they outrank anything an
algorithm suggests: they lead the DMSMS form's Suggested Replacement column, get their own sheet in
the workbook, and a line whose primary is in trouble but whose alternate is stocked says so in the
table without being expanded.

**Finds alternatives to a part in trouble.** Besides its supplier column, Nexar answers a second
question: not what a part costs, but what could be used instead. That one runs only for parts the
comparison has already found to be obsolete, NRND, end of life or simply unavailable — and only for
the ones you tick — so a metered quota is spent on the parts that need it rather than on every line
of a healthy BOM. What it finds fills the DMSMS form's Suggested Replacement column. It is a
separate query from part search and fails separately: `similarParts` is not on every Nexar plan, and
a plan without it still gives you the supplier column.

**Says when each part can arrive, and from whom.** The **Lead times** button bands every part
looked up — from a BOM or typed by hand — by how soon it can actually be here. Stock on hand counts
as zero days, because it ships today whatever the factory quotes behind it; where several suppliers
can deliver in the same time, the cheapest of them is the one named. In stock or inside three weeks
is green, three to eight weeks yellow, longer orange, and a part no supplier carries is red and says
so. Exports to Excel with the same shading.

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
| Nexar *(optional)* | [nexar.com](https://nexar.com/) | Create an application **with Supply access** and copy its **Client ID** and **Client Secret**. A Design application will not work: it signs a user in rather than authenticating itself, and the part data lives in the Supply API. Adds a fourth supplier column and powers **Find alternatives**; every other supplier works without it |

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
| `--match exact` / `--match relaxed` | Whether a supplier answer must be the same part number, or may be the closest returned (default: `exact`) |
| `--ignore-prefix PREFIX` | Skip part numbers starting with PREFIX (repeatable); replaces the `ASY0`/`CBL0`/`DES0`/`PCB0` default |
| `--no-ignore-prefixes` | Look up in-house part numbers too |
| `--no-merge-duplicates` | Keep repeated part numbers as separate lines instead of adding their quantities |
| `--ignore-skip-column` | Look up lines the BOM marks YES in a skip-to-production column |
| `--no-alternates` | Do not look up the approved alternates named in the BOM's alternates column |
| `--show-skipped` | List every skipped line and why |
| `--lead-time FILE` | Write a long-lead-times report: every part banded by how soon it can arrive |
| `--dmsms FILE --program NAME` | Write a DMSMS case form for every at-risk part |
| `--dmsms-status STATUS` | Narrow the form to given lifecycle statuses (repeatable) |
| `--alternatives` | Ask Nexar what could replace each at-risk part, and fill in Suggested Replacement |
| `--nexar-check` | Test the Nexar credentials on their own and print what the token endpoint said |
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
   them. An *Alternate Part Number* column is detected the same way, and can be
   remapped or unmapped just as easily.
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

### Short, and splitting an order

**A line is short only when every supplier's stock, added together, still cannot cover it.** Four
suppliers holding 80 pieces each are not a shortage against a need for 200 — they are 320 pieces
and three purchase orders. Only the combined position counts, so a line that used to be flagged
because no *single* supplier could cover it now reports what to do instead:

> Split 200 across 3 suppliers: Mouser 80, DigiKey 80, TrustedParts 40

The split is worked out for you. Expand the row and **Split this order** lists each supplier, how
many to take, how many that actually buys once minimum order quantities and packaging multiples are
applied, what they hold, and what each order costs — with the total underneath.

**Lead time decides who supplies it, and price only settles a tie.** Every unit in a split is drawn
from stock on hand, so every one of them ships today whatever factory lead time sits behind that
supplier's shelf — which means the whole field is tied on speed and price is all that is left to
order it by. Cheapest shelves are drawn down first. Nothing is ever ordered from a slower supplier
to save money: a supplier who can ship today is always drawn on before one who quotes four weeks,
whatever the difference in price.

That has a consequence worth knowing: where one supplier holds the whole quantity but at a higher
price, the tool will still propose a split if drawing part of it from a cheaper shelf costs less.
Both orders ship today, so nothing is given up on speed — but it is two purchase orders instead of
one, and the row says so.

Where it shows up:

- **In the table** — each supplier's stock cell carries a `take 80` badge, so the quantities are on
  the row without expanding it. The old per-supplier `short` badge is gone: whether one supplier
  alone can cover the line is no longer the question being asked.
- **In the verdict** — the split named in full, as a note rather than a warning.
- **In the stat strip** — *Stock risk* counts only lines the combined stock cannot cover.
- **In the report and the workbook** — a **Split orders** sheet with one row per purchase order, and
  the parts table pricing the whole split rather than one supplier's share of it.
- **On the lead-time report** — a split-covered line reads *In stock, split*, because it is. Banding
  it at the factory lead time of whichever supplier happened to be quickest would be wrong twice
  over: too slow, and from the wrong supplier.

**Short on stock and unobtainable are different problems, and only the second is an emergency.** A
part nobody has on a shelf but everybody can order in a fortnight is a normal factory order, so it
reads *"No stock on hand — soonest is 2 weeks from DigiKey"* at warning level. It is called a
shortage in red only when the combined stock falls short **and** no supplier will quote a date for
the rest — at which point there is genuinely nowhere to go.

When the combined stock genuinely is short, the line says by how much and names whoever could
factory-order the remainder soonest:

> 15 of 200 covered — 185 still to find; soonest for the rest is Mouser in 6 weeks

The per-supplier question has not gone away, it has just moved to where it belongs: the supplier
cart tiles price single-sourcing everything from one place, so *"3 lines it cannot cover alone"* is
exactly the right thing for those to say.

### Long lead times

**Lead times** answers a purchasing question rather than a pricing one: not what a part costs, but
when it can be here and from whom. Every part that was looked up appears — a whole BOM, several
BOMs, or a handful of part numbers typed into the lookup form — sorted worst first, so the report
reads as a worklist rather than a table to search.

Two rules decide who is named for each line:

1. **Stock on hand is the fastest answer there is.** A supplier holding enough counts as zero days,
   whatever the factory quotes behind it. A part in stock at one supplier beats the same part quoted
   at four weeks by another, even if the quote is cheaper.
2. **Where several suppliers can deliver in the same time, the cheapest wins.** Being equally fast,
   price is the only thing left to choose on.

Each line is then shaded by how long it takes:

| | Band | Means |
| --- | --- | --- |
| 🟩 | **In stock or under 3 weeks** | Ships now, or quoted inside three weeks |
| 🟨 | **3–8 weeks** | Plan around it |
| 🟧 | **Over 8 weeks** | The long poles |
| 🟥 | **Not available** | No supplier searched can provide the part at this time |
| ⬜ | **Unknown** | Carried, but no supplier would quote a date |

Sub-cent unit prices keep their five decimal places on a shaded row, and quantities stay integers:
in Excel a fill and a number format are one style, so each band carries a copy per format rather
than giving one up for the other.

Two of those need a word of explanation. The colours you asked for leave a gap between "in stock"
and "three weeks", so a part quoted inside a fortnight is shaded green alongside stock: the band
exists to separate "order it" from "plan around it", and a fortnight is not something to plan
around. The exact wording is still in the Availability column either way. And a part that is
carried but that nobody would date is left **unshaded** rather than binned as long: TrustedParts,
for one, reports stock without a lead time, and colouring that orange would assert a delay no
supplier ever quoted.

Alongside the winner each line carries the price, the stock, how many suppliers carry it at all,
who else could supply it and when — and, where the BOM named an approved alternate that can
actually be bought, that alternate, because it is the one thing that rescues a line that is late or
gone.

**The colours follow the parts everywhere they go.** The same four fills shade the **Parts** sheet
and the *Needs a decision* table of the summary workbook, the summary report on screen, and its
Print / PDF output — one palette, defined once, so no two views can drift on to different colours.
Each shaded table carries a key. The CSV has no colour to give, so the band travels there as a word
in an **Availability** column rather than being the one thing that export loses.

**Export Excel** writes the lead-time report as a workbook of its own: a legend sheet with the band
counts, then the table with each row shaded light green, light yellow, orange or light red. On the
command line:

```bash
python3 bom.py my-bom.csv --lead-time lead-times.xlsx
```

The same table is also a **Lead times** sheet inside the main `-o report.xlsx` workbook, so one
export carries everything.

### The alternates column

Many BOMs carry a column naming the parts engineering approved as substitutes. It is detected by its
header the same way every other column is — *Alternate Part Number*, *Alternate Parts*, *Alt P/N*,
*Alt MPN*, *Approved Alternates*, *Second Source*, *Substitute Part*, and the shorter spellings —
and can be remapped or unmapped in the browser like any other.

One cell may name several parts. Commas, semicolons, pipes and newlines separate them; a space does
not, because part numbers contain spaces, and neither does a slash — `LM358DR/NOPB` is one part
number, not two. A cell that says `N/A`, `None`, `TBD` or `-` names none. Repeats collapse.

Each named alternate is then looked up against the same suppliers as the primary, at the same
quantity — the question an alternate answers is "could this cover the same build", which is a
question about the same number of pieces — and reduced to one verdict:

| Verdict | Means |
| --- | --- |
| **available** | Found, stocked somewhere in the quantity this line needs, and not itself ending |
| **no stock** | Found, and nobody has enough today |
| **ending** | Found and stocked, but obsolete or discontinued itself — a substitution that buys you nothing |
| **no match** | No supplier carries that number at all |

Where that shows up:

- **In the comparison table** — a line whose primary is unbuyable, unstocked or ending, but whose
  BOM names a stocked alternate, is badged `alt: <part>` in the Verdict column. One with alternates
  that are all unavailable says that instead. A healthy line is not badged: it has no question to
  answer.
- **When you expand a row** — every alternate, with its verdict, lifecycle, stock, best price and
  which suppliers have it.
- **In the summary report** — an *Approved alternates* column on the parts table and on *Needs a
  decision*, and an **Alternates** sheet in the exported workbook.
- **On the DMSMS form** — an approved alternate leads the Suggested Replacement column, ahead of
  anything Nexar suggests and anything the supplier's own catalogue field says. A stocked one is
  preferred to one that is not, and an unavailable one is still named, labelled `unavailable`,
  because knowing the approved substitute is also gone is the point of the form.

Looking the column up costs one lookup per alternate per supplier. `--no-alternates` on the CLI, or
`SKIP_ALTERNATES=1` in `.env`, leaves it unqueried — the column is still read, mapped and shown,
just not priced. `MAX_ALTERNATES_PER_PART` caps how many one line may name (default 6).

### How a part number is matched

Exact, by default. A supplier's answer counts only when its manufacturer part number is the one you
asked for, after folding case and punctuation:

| Asked for | Supplier returned | |
| --- | --- | --- |
| `RC0603FR-0710KL` | `RC0603FR0710KL` | **match** — same number, punctuated differently |
| `LM358DR` | `lm358dr` | **match** — case is not identity |
| `LM358` | `LM358DR` | **rejected** — a different device in a different package |
| `LM358DR` | `LM358` | **rejected** — the suffix is part of the number |
| `LM358DR` | `LM358DR/NOPB` | **rejected** — decide for yourself whether that substitution is safe |

A rejection is reported, not hidden: the cell reads *"DigiKey has no exact match for LM358 — closest
was LM358DR (Texas Instruments)"*, so a part spelled differently is distinguishable from one nobody
carries. That matters for the last row above: strict matching can turn a legitimate part into a
miss, and the message is what makes it obvious rather than silent.

If you would rather see the nearest thing than nothing, `MPN_MATCH=relaxed` in `.env` or `--match
relaxed` on the command line restores the old behaviour, and the detail panel marks such a row
*closest match*. TrustedParts is always exact — its API is keyed on the part number.

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
they are missing. This query is never called during a BOM analysis — only the supplier search below
is.

### Nexar as a supplier

Once its credentials are set, Nexar also takes its place beside DigiKey, Mouser and TrustedParts as
a fourth column. Like TrustedParts it is an aggregator: one part number comes back with every seller
Nexar knows, each with its own stock, price ladder, minimum order and packaging. The column quotes
whichever of them is cheapest for the whole order, and expanding the row lists them all. Sellers
Nexar reports as unauthorized are kept but ranked last, the way a marketplace listing is.

Lifecycle status arrives as a specification rather than a field, so it is read from the specs and
only used when the value is actually a lifecycle status — an unfamiliar spec value is not a claim
about supply, and is left as Unknown rather than rendered as one.

A whole BOM goes out as **one batched request** (`supMultiMatch`), not one per part, which is what
makes a metered plan last. `NEXAR_BATCH_SIZE` sets how many parts ride in each request.

If `supMultiMatch` is not in your plan's schema, the client notices the rejection, drops to the
per-part `supSearchMpn` query for the rest of the run, and says nothing about it — reporting every
line as "not carried" because one query is missing would be a lie about availability. A rejection
that means something else is raised, not worked around.

To keep Nexar out of the comparison and use it for alternatives only — worth doing if you set the
credentials up for that and would rather not spend the quota on every line — set
`SKIP_NEXAR_SUPPLIER=1`, or pick suppliers explicitly with `-s` on the CLI.

`python3 bom.py --nexar-check` tests the credentials and both queries separately and prints exactly
what Nexar said about each, so "alternatives are not on this plan" reads differently from "these
credentials are wrong".

#### If Nexar refuses the credentials

```bash
python3 bom.py --nexar-check
```

That exchanges the credentials for a token on their own and prints what the endpoint actually said —
nothing else runs. A failure there is one of:

| What it says | What it means |
| --- | --- |
| `invalid_client` | The ID or secret is wrong. A secret truncated on paste looks exactly like this |
| `invalid_scope` | The application was not granted the scope being asked for. The app retries once without a scope on its own; if that fails too, set `NEXAR_SCOPE` to a scope your application does have, or `NEXAR_SCOPE=` (empty) to ask for none |
| `unauthorized_client` | The credentials are real, but the application cannot authenticate as itself. Most often it is a **Design** application, which signs a *user* in instead — a different grant and a different API. Alternatives come from the **Supply** API, so the application needs Supply access |

The scope defaults to `supply.domain`. Setting `NEXAR_SCOPE=` to an empty value is different from
leaving it out: empty means *ask for no scope at all*, which is what some Nexar applications want.

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
| `MPN_MATCH` | `exact` | `exact` accepts only the part number asked for, ignoring case and punctuation; `relaxed` accepts the closest match the search returned |
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
python3 -m unittest discover -s tests -t .   # 329 tests, no network access required
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
and exit codes, the comparison verdicts, exact part-number matching — that a suffix counts and
punctuation does not, and that a rejection names what came back — screening — that a prefix only counts at the start of
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
- **Matching is by manufacturer part number, and by nothing else.** Package, temperature grade and
  tolerance live in the suffix, so they are compared as part of the number. What this cannot do is
  know that two differently numbered parts are interchangeable — that is what **Find alternatives**
  is for, and it is a decision rather than a lookup.
- **Strict matching can produce a miss where a match exists.** A packaging or lead-free suffix
  (`/NOPB`, `-TR`) makes a different number, so it is rejected. The message names what came back, so
  the case is visible; `MPN_MATCH=relaxed` accepts it if you would rather judge each one.
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
