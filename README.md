# Drawing & BOM Studio

A desktop app for editing Microsoft **Visio drawings (`.vsdx`)** and **Excel
workbooks (`.xlsx`)** in bulk. It streamlines the repetitive edits of a drawing
release — find/replace, parts add/remove/edit, revisions, approvals, and Change
Log — makes a next-revision copy of each file, and optionally exports to PDF.

The app is a **guided 3-step wizard** — **1 Files → 2 Edit → 3 Review & Run** —
with a stepper across the top and **Back / Next** navigation. A friendly
**animated robot helper** in the corner shows what to do at each step, and waves,
thinks, types, sweats, celebrates, or looks worried as the run progresses.

The **Edit** step holds **three tabs that all run together** in a single pass:

- **Find → Replace** — text rules, each aimed at all files or specific ones.
- **Parts** — add, remove, or edit rows in the Visio parts tables and Excel
  BOMs, with item/line numbers renumbered automatically.
- **Approve & Revise** — bump revisions, append revision-table rows, sign off
  approvals (EE / ME / Production), stamp the Author box, and add Change Log
  entries.

Workflow:

1. **Step 1 — Files:** choose a single file, or switch to **Batch** and add many
   files (or a whole folder). You can mix `.vsdx` and `.xlsx`; each file's type
   is **detected automatically**. Click **Next**.
2. **Step 2 — Edit:** fill in whichever tabs you need — Find → Replace rules,
   Parts add/remove/edit, and Approve & Revise actions. Anything you stage on any
   tab is applied in the same run. Click **Next**.
3. **Step 3 — Review & Run:** check the summary and options, then click **Run**.
4. Get a **next-revision copy** of each file next to each original. Your
   originals are never modified.

### Revisions

The tool works on **copies**, never the originals. By default it names each copy
as the **next revision letter**, read from the file name:

- `Floor Plan REVA.vsdx` → `Floor Plan REVB.vsdx` (A→B→C … major changes only).
- It also updates the revision **inside the file**: the single-letter box/cell
  (e.g. `A`) that sits next to a `REV`/`Revision` label is bumped to match. The
  title block repeats on every page/sheet, so this is updated **on every page
  of a Visio drawing and every worksheet of an Excel workbook** (not just the
  first). In Excel the protected **Change Log** sheet is left alone, and stray
  single letters elsewhere (grid/zone labels) are too.
- Already at `REVZ`? That file is skipped with a warning (no next letter).
- No `REVx` in the file name? The copy falls back to `*_edited`.

Both behaviors are checkboxes you can turn off (rename only, or no revision bump
at all).

The find/replace edits the file directly (both formats are ZIPs of XML). For
Visio it touches the text in shape `<Text>` blocks; for Excel the shared-strings
table, inline strings, and drawing text boxes. Geometry, styles, formatting, and
fonts are left untouched, and a term is matched even when it's split across
formatting runs.

### Built-in help ("How to use")

Every feature below is documented inside the app: click **❓ How to use** in the
top-right of the window for a scrollable job aid covering each step. Click
**🖨 Open printable / PDF version** in that window to open a clean, printable
HTML copy you can save or print to PDF for training/reference.

---

## Requirements

- **Python 3.8+** — use the installer from [python.org](https://www.python.org/downloads/).
  On Windows/macOS the standard installer already includes the `tkinter` GUI
  library. On Linux install it with e.g. `sudo apt install python3-tk`.
- **LibreOffice** — only needed for PDF export (text replacement works without it).
  - Windows: <https://www.libreoffice.org/download/>
  - macOS: `brew install --cask libreoffice` (or the `.dmg`)
  - Linux: `sudo apt install libreoffice`
- **Pillow** *(optional)* — `pip install pillow`. If present, the buttons are
  drawn with smooth anti-aliased corners; without it they still work, just with
  plainer corners. The app already renders at your display's real DPI so text
  stays sharp either way.

No `pip install` is required — the app uses only the Python standard library
(Pillow is an optional extra for the smoothest-looking buttons).

---

## Easiest ways to launch (Windows)

You have three options, from simplest to most self-contained:

1. **Download a ready-made `.exe` (no Python needed).**
   In the GitHub repo, open the **Actions** tab → the latest **"Build Windows
   EXE"** run → under **Artifacts**, download **`DrawingBOMStudio-windows`**.
   Unzip it and double-click `DrawingBOMStudio.exe`. (PDF export still needs
   LibreOffice installed.)

2. **Double-click `run_drawing_bom_studio.bat`** (needs Python installed). It
   launches the app without touching the command line. If Python is missing it
   tells you where to get it.

3. **Build your own `.exe` once** by double-clicking **`build_exe.bat`** (needs
   Python). It installs PyInstaller and produces `dist\DrawingBOMStudio.exe`,
   which you can then move/share and run on its own.

## Run the app (GUI, any OS)

```bash
python drawing_bom_studio.py
```

A window opens:

1. **Pick your files (step 1 in the window).** Choose **Single file**, or
   **Batch (multiple files)** to process several at once (you can mix `.vsdx`
   and `.xlsx`). Use **Add files...** (select several at once) or **Add
   folder...** (adds every `.vsdx`/`.xlsx` in a folder). The chosen files appear
   in the list; **Remove selected** / **Clear** manage it.
2. **Add your rules (step 2).** Each rule is a row with a separate **Find:** box
   and **Replace with:** box, plus an **in:** dropdown. Use **+ Add another rule**
   for more terms. The **in:** dropdown targets files by **document type**
   (detected from the file name — see *Document types* below):
   - **All files** (the default) — the rule runs on every file.
   - Or untick "All files" and **tick one or more types** (BOM / System Drawing /
     Cable Drawing) to run the rule only on files of those types.
3. **Set options (step 3)** and click **Run**. Each file is saved
   as its next-revision copy (or `*_edited`) next to the original, and the output
   folder opens when it finishes. The Status box lists what happened per file.

### Document types

The tool classifies each file by its **name** so rules can target a whole type
and the change summary can be grouped:

| Type | File name contains | Example |
|------|--------------------|---------|
| **BOM** | `DOCxxxxx` | `DOC00475_…` |
| **System Drawing** | `DWGxxxxx` | `DWG01234_…` |
| **Cable Drawing** | `CBLxxxxx` | `CBL00134-01_…` |

Anything else is **Other**. The **in:** dropdown on each rule lists the types
present in your loaded files, so a rule can apply to, say, all Cable Drawings at
once.

Options:

- **Case sensitive** — match the exact capitalization (off by default).
- **Whole word only** — only match complete words (won't change `foobar` when
  finding `foo`).
- **Also export PDF (LibreOffice)** — *off by default.* Produces a PDF via
  LibreOffice. Note: LibreOffice does not evaluate Visio page-number fields, so
  a "Sheet X of Y" title block renders as "Sheet 0 of Y". For correct sheet
  numbers and best fidelity, open the edited `.vsdx` and **export the PDF from
  Visio** (File → Export → PDF).
- **Save copy as next revision (REVx → next)** — name the copy as the next
  revision letter instead of `*_edited` (on by default; see *Revisions* above).
- **...and update the REV box in the drawing** — also bump the `REVx` letter box
  inside the drawing to match the new file name (on by default).

### Excel: find & edit BOM rows

For Excel bill-of-materials sheets, you can edit a part's whole row, not just
its number. Put the part number in a **Find** box, then click
**Excel: find & edit rows...**. The tool:

1. Finds the column headed **P/N** / **Part Number** (also `P/n`, etc.) in each
   loaded `.xlsx`, and locates every row whose part number matches a Find value.
2. Lists each matched row **grouped by file** (e.g. *DOC11111…* then its rows,
   *DOC22222…* then its rows), labelling each match with the **sheet name** and
   **row** it was found on, and showing that row's other fields under their
   headers — **Manufacturer**, **Unit Cost**, **Description**, **Qty** (or
   Quantity), **Notes** (or Comments) — pre-filled with the current values.
3. Lets you edit each row **individually per file**, so the same part can get a
   different Qty (or any value) in one file than another. Click **Save edits**
   to stage them. The editor opens in its **own window** on top of the tool.
   - A **Show find value(s)** dropdown lets you filter the list to one or more
     of the find values (useful when many rules produce a lot of matches).
   - **Copy 1st down** buttons (**Manufacturer**, **Unit Cost**, **Description**,
     **Qty**) copy the first shown row's value into the rest of that **find
     value's** rows, so you can fill every instance of a part from the first
     one. With find values selected in the dropdown, the copy is limited to
     those.
   - **Reset fields** restores every field to the originally found data.
4. **Refresh** re-runs the lookup — use it after adding files or changing the
   Find value(s); values you've already typed are kept.

On **Run**, each row's staged changes are written to that exact
cell in that file. Fields you leave unchanged are not touched, numbers stay
numeric, and the P/N itself is changed by your normal Find → Replace rule.

### Visio: find & edit parts-table rows

The R/C/P/W drawing sheets carry a **parts list** at the bottom (columns
**Item · Ref/DES # · Cable Name · Description · Part Number · Manufacturer ·
Qty/Length · Unit**). This is the Visio counterpart of *Excel: find & edit
rows*, with the same window, buttons and behaviour. Put a part number in a
**Find** box, then click **Visio: find & edit rows...**. The tool:

1. Scans every loaded `.vsdx`, finds each parts table (they're stored as
   embedded Excel "Cable BOM" objects), and locates every row whose **Part
   Number** matches a Find value.
2. Lists each matched row **grouped by file and sheet** (e.g. *Sheet: R0001*),
   labelled with its **Part Number** and **row**, showing that row's other
   fields — **Ref/DES #**, **Cable Name**, **Description**, **Manufacturer**,
   **Qty/Length**, **Unit** — pre-filled with the current values.
3. Lets you edit each row **individually**, so the same part can get a
   different value on one sheet than another. The editor opens in its **own
   window** with the same helpers as the Excel one:
   - a **Show find value(s)** dropdown to filter the list;
   - **Copy 1st down** buttons (**Manufacturer**, **Description**,
     **Qty/Length**, **Unit**) that copy the first shown row's value into the
     rest of that find value's rows;
   - **Reset fields** to restore the originally found data;
   - **Refresh lookup** to re-scan after adding files or changing Find values.

The **Part Number** itself is changed by your normal **Find → Replace** rule:
when a Find value matches a row's part number, that cell is replaced with the
Replace value (the other fields are edited in the window above). A part number
written **"<P/N> or equiv."** still matches a Find value of just the part number
— and the **whole** cell (including the "or equiv.") is replaced with your
Replace value. (To keep "or equiv.", include it in the Replace box.)

On **Run**, each change is written to that cell in the embedded
worksheet **and drawn into the table's cached picture**.

> **Note on Visio embedded tables.** Visio keeps its own cached image of an
> embedded Excel object and may not repaint it on the page view until the
> object is **double-clicked once** (which reloads it from the corrected
> worksheet). The underlying data is already correct — it prints and exports
> right — so the one-time double-click is only to refresh Visio's on-screen
> preview. This applies to the parts-table edits and the add/remove below, and
> to the revision-table row.

### Parts tab: add or remove parts

The **Parts** tab adds whole rows to — or deletes whole rows from — the Visio
parts tables and Excel BOMs, and **renumbers the Item / line sequence**
automatically. It works on the same tables as *find & edit rows* above, and runs
together with everything else when you click **Run**.

The parts list is treated as the **contiguous block of rows with a Description**,
ending at the last one. An approval or revision block that sits lower in the same
sheet (after a blank row) is **never touched, shifted, or renumbered** — fixing
the earlier behaviour where a removal could distort the data below the table.

**Remove parts.** Type a part number in the **Remove** box (independent of the
Find boxes) and click **Find & choose rows to remove...**. Every parts-table /
BOM row with that part number is listed **per file and sheet**, each with a
checkbox. Tick the ones to delete and click **Save removals**. On run, each
ticked row is removed, the rows below it **shift up to close the gap**, and the
item/line numbers are renumbered.

Removals **accumulate across part numbers**: look up one number and save, then
look up another and save — the second set is *added* to the first rather than
replacing it, so you can stage several different parts for removal in one run.
(Re-open the same number to review or clear what's ticked.)

**Add parts.** Click **Add parts...**, then use the dropdowns to pick a **file
type** (BOM / Cable Drawing / System Drawing), the **file**, and the **sheet /
table**. The current parts (with their item numbers) are shown for reference;
fill in the blank row(s) at the bottom (use **+ Add row** for more) and click
**Save**. You can stage new parts on several sheets and files before running.

Each new row has a trailing **Insert as item #** box that sets *where* the part
lands: type `3` and it becomes item 3, pushing the parts at 3-and-below down;
leave it **blank** to add at the **end** of the parts list (right after the last
Description — never below an approval block). Item/line numbers are renumbered
automatically. **Adds are applied before removals.**

Remove and Add are **staged**, not applied immediately: nothing changes until
you click **Run**, and they apply alongside the Find → Replace rules and the
Approve & Revise actions in the same pass. (See the Visio embedded-table note
above about double-clicking once to refresh Visio's on-screen preview.)

### The "Change Log" sheet is protected

A sheet named **Change Log** is **never** modified by Find → Replace or by the
BOM row editor — it's left exactly as-is for traceability (even if your search
term appears in a change description there).

To add an entry to it, click **Excel: add Change Log entry...** and fill in the
columns — **Item**, **ECN #** (or **ECO#**), **ERB Approval Date**, **Change
Description**, **Change Author**. On **Run**, that row is
**appended to the Change Log sheet of every Excel file**, at the next free row.
The next row is found using the **ECN #** column (so it works even when the Item
numbers run past the last real entry). Leave **Item** blank to keep any
pre-filled item number; leave any field blank to skip it.

### Excel: Author name + date

Many BOM title blocks have an **Author:** box with the author's name in the cell
to its right and a date in the next cell. Click **Excel: set Author + date...**
and type the new name. On **Run**, for **every** Excel file the
name beside the **Author** box is set to your value, and the date beside it is
stamped with **today's date** — kept in the **same format** as the date that was
there (an Excel date serial stays a properly-formatted date).

### Visio: add a revision-table entry

Visio cover pages usually carry a **revision-history table** in a corner — the
chart with **REV / DESCRIPTION / DATE / APPROVED** columns. This is the Visio
counterpart of the Excel *add Change Log entry* feature. Click **Visio: add
revision entry...** and the tool:

1. Finds the table on each loaded `.vsdx` cover page by its **column headers**
   (a generous set of spellings: `REV`, `REVISION`, `DESCRIPTION`/`REASON FOR
   CHANGE`, `DATE`, `APPROVED`/`BY`/`ENGINEER`, `ECN`/`ECO`, …) and their
   on-page geometry, and shows you **which columns it detected**.
2. Lets you fill in **ECN #**, **Description**, **Date**, and **Approved By**
   (leave any blank to skip it) — these are the **same for every file**. You do
   **not** type the **REV**: it's filled **automatically per file** with that
   file's own **next revision letter**, so a batch of drawings at different
   revisions each gets its correct next letter (e.g. a `REVD` file gets an `E`
   row, a `REVF` file gets a `G` row).
3. On **Run**, adds the new revision row to **every** Visio file:
   an existing **blank row** in the table is filled if there is one, otherwise a
   new row is **cloned** from the last row (same columns/formatting) and placed
   just below it. Because many drawings draw the table grid as a fixed
   background image that can't be extended, the tool also **draws matching
   border lines** (the row box plus column dividers) around the appended row so
   it lines up with the rows above.

Some drawings don't store the revision table as native Visio text boxes at all —
they keep it as an **embedded Excel worksheet** (an OLE object shown in the
corner). The tool handles those too: it scans the drawing's embedded
worksheets, finds the one whose header row is a revision table (a **REV** column
plus **DESCRIPTION / DATE / APPROVED**, …), and writes the new entry straight
into that worksheet — filling the first **pre-formatted blank row** that's
already there for the next revision (or adding one if the table is full). The
Status box reports it as an *embedded sheet*. No grid lines are drawn for these,
because the worksheet supplies its own borders. Every value written into an
embedded revision table — by both the revision-entry and the **approval**
features — is formatted **Calibri, size 8, centered and middle-aligned**, so it
matches the table regardless of what format the blank cell happened to carry.

Visio shows an embedded object from a **cached picture** and only refreshes it
when you double-click the object — so editing the worksheet alone would leave
the drawing looking unchanged until you opened the table by hand. To avoid that,
the tool also patches that cached picture (the metafile Visio displays),
**drawing the new row's text straight into it** using the table's own font and
spacing. The result: the new revision row (and any approval) shows up
**automatically** when the drawing is opened — no manual refresh needed.

Because a Visio "table" is really just positioned text boxes (or an embedded
sheet), the detection is **conservative**: a real revision-history table must
have a **REV/LTR** column, which is what tells it apart from the title-block
sign-off block (DRAWN BY / CHECKED BY / ENGINEER / APPROVED BY with dates) and
from BOM tables. If it can't confidently locate the table on a file (or can't
find a safe place to add the row), that file is **left completely unchanged**
and the Status box says so — it never risks corrupting a drawing. If your
template uses unusual column labels, share a sample `.vsdx` so the detection can
be tuned to it.

### Approvals (sign off many files at once)

An approver can sign off **every** loaded file in one click, without opening
each drawing/workbook by hand.

**Visio — approve a revision.** Click **Visio: approve revision...** and type
your **name**. By default it signs off **each file's own latest revision**,
auto-detected per file — so a batch of drawings at different revisions each gets
the correct row (the dialog shows the latest letter found in each file). Leave
the **REV letter** field blank for that, or type a specific letter to force one.
On **Run**, your name is written into the **Approved** column of
that revision row, in each loaded Visio file's revision table — whether that
table is native Visio text cells or an embedded Excel worksheet. If that row had
no Approved entry yet, one is added in the Approved column (not the Date
column).

**Excel — approve by discipline.** Click **Excel: approve (EE/ME/Prod)...**,
pick the **discipline** (**EE**, **ME**, or **Production**) and type your
**name**. The dialog shows which approval boxes it found. On **Replace &
Convert**, for **every** Excel file your name is placed in the cell **beside
that discipline's label** and **today's date** in the next cell — kept in the
**same format** as the existing date — on **every sheet except the Change Log**.

Discipline labels are matched flexibly (e.g. `EE`/`Electrical`, `ME`/
`Mechanical`, `Production`/`Prod`/`MFG`). A file with no matching label is left
unchanged.

An approval is a **sign-off, not a revision change**, so it does **not** bump
the revision letter even if *Save copy as next revision* is ticked. The approved
copy is named **`<name>_approved_<today's date>`** (e.g.
`CBL00011_REVD_…_approved_2026-06-28.vsdx`).

### Output folder

By default each finished copy is written **next to its original**. To collect
everything in one place, use **Output folder → Choose folder...** in the Options
section — every result (and the change summary) is written there instead.
**Use source folder** switches back to the default.

### Reset all

The orange **Reset all** button (next to *Run*) clears the loaded
files, every find/replace rule, all staged edits (BOM rows, Change Log entry,
Author, Visio revision entry, and both approvals) and the output folder — so you
can start from scratch on a new file or batch. It asks for confirmation first.

### Change summary (for approval review)

With **Generate change summary** ticked (on by default), the tool writes a
**`Change_Summary_<timestamp>.html`** document next to the edited files (or in
your output folder) and opens it. Files are **grouped by document type** (BOM /
System Drawing / Cable Drawing), then by file, and for every file it lists each
change as **Change (where) · Before · After** — including text replacements, BOM
row edits, the Author name/date, the appended Change Log row, the revision
bump, the **Visio revision-table row** that was added, and **Visio parts-table**
edits (field changes and Part Number replacements, labelled by sheet such as
*R0001*). Cells are labelled by their column or title-block label (e.g. *P/N*,
*Unit Cost*, *Author*, *Revision*, *Manufacturer*), and dates are shown as real
dates.

**Identical changes are grouped into a single line.** A change that happens the
same way across many sheets/pages — most notably the **REV bump**, which touches
every sheet — appears once, noting whether it occurred on **all sheets** or on
specific named sheets, instead of one row per sheet. This keeps the summary fast
to review. An approver can check every change in one document without opening
each file, and can print it to PDF.

**Parts tables and BOMs are compared by row, not by position.** When a part is
added or removed and the rows below it shift up (and renumber), the summary does
**not** list every shifted item as a change — it reports only the genuine
**Part added** / **Part removed**, plus any in-place field edits. The Item /
line-number column is ignored, so renumbering alone is never flagged.

---

## Command-line use (optional / scriptable)

Passing any arguments runs in CLI mode instead of opening the window. Inputs may
be `.vsdx` or `.xlsx` (detected automatically):

```bash
# Single replacement, edited .vsdx only
python drawing_bom_studio.py drawing.vsdx --find "Old Server" --replace "New Host"

# Multiple replacements + PDF export
python drawing_bom_studio.py drawing.vsdx \
    --find "Old Server" --replace "New Host" \
    --find "2023"       --replace "2026" \
    --pdf

# Batch: every rule is applied to ALL listed files
python drawing_bom_studio.py one.vsdx two.vsdx three.vsdx \
    --find "Old Server" --replace "New Host" --pdf

# Save each copy as the next revision (REVA -> REVB) and bump the in-drawing box
python drawing_bom_studio.py "Floor Plan REVA.vsdx" \
    --find "Old Server" --replace "New Host" --bump-revision --pdf

# Choose the output name (single file only); case-sensitive, whole-word
python drawing_bom_studio.py in.vsdx -f Dev -r Prod -o out.vsdx \
    --case-sensitive --whole-word --pdf
```

| Flag | Meaning |
|------|---------|
| `inputs` | One or more `.vsdx` files (every rule applies to all of them) |
| `-f`, `--find` | Text to search for (repeatable) |
| `-r`, `--replace` | Replacement, paired with each `--find` in order |
| `-o`, `--output` | Output `.vsdx` path (single input only; default: `<name>_edited.vsdx`) |
| `--pdf` | Also export to PDF |
| `--case-sensitive` | Match capitalization exactly |
| `--whole-word` | Only match whole words |
| `--bump-revision` | Name each copy as the next revision (REVx → next) read from the file name |
| `--no-rev-text` | With `--bump-revision`, only rename the file; don't touch the REV box in the drawing |

> The CLI applies every rule to every file listed. **Per-file targeting**
> (a rule that only touches certain files) is available in the **GUI** via each
> rule's **Applies to** button.

---

## Notes & limitations

- **Format:** supports the modern **`.vsdx`** format. For legacy binary `.vsd`
  files, open them in Visio and *Save As* `.vsdx` first.
- **Split text is handled:** Visio often stores a label as several formatting
  runs (e.g. `Old ` `<cp/>` `Server`), sometimes even across paragraph markers.
  Search terms are matched across those runs, so `Old Server` is still found and
  replaced — and the replacement is collapsed onto **one line**, dropping the
  stray run/paragraph markers that used to push part of the new text onto a
  second line.
- The original file is never modified; results are written to new files
  (the next-revision copy, or `*_edited.vsdx` / `*_edited.pdf`).
