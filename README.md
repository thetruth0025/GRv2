# Visio / Excel Text Replacer

A small desktop app that finds and replaces text **everywhere it appears** in
Microsoft **Visio drawings (`.vsdx`)** and **Excel workbooks (`.xlsx`)**, makes a
next-revision copy, and optionally exports to PDF.

Workflow:

1. **Choose** a single file — or switch to **Batch** mode and add many files (or
   a whole folder). You can mix `.vsdx` and `.xlsx`; each file's type is
   **detected automatically**.
2. **Enter** one or more *Find* → *Replace with* rules. Each rule's **in:**
   dropdown lets you target **all files** or **tick several specific files**.
3. Click **Replace & Convert**.
4. Get a **next-revision copy** of each file next to each original. Your
   originals are never modified.

### Revisions

The tool works on **copies**, never the originals. By default it names each copy
as the **next revision letter**, read from the file name:

- `Floor Plan REVA.vsdx` → `Floor Plan REVB.vsdx` (A→B→C … major changes only).
- It also updates the revision **inside the file**: the single-letter box/cell
  (e.g. `A`) that sits next to a `REV`/`Revision` label is bumped to match. In
  Visio that's the text box nearest the `REV` label on the front page; in Excel
  it's the cell next to the `REV` cell. Stray single letters elsewhere (grid/zone
  labels) are left alone.
- Already at `REVZ`? That file is skipped with a warning (no next letter).
- No `REVx` in the file name? The copy falls back to `*_edited`.

Both behaviors are checkboxes you can turn off (rename only, or no revision bump
at all).

The find/replace edits the file directly (both formats are ZIPs of XML). For
Visio it touches the text in shape `<Text>` blocks; for Excel the shared-strings
table, inline strings, and drawing text boxes. Geometry, styles, formatting, and
fonts are left untouched, and a term is matched even when it's split across
formatting runs.

---

## Requirements

- **Python 3.8+** — use the installer from [python.org](https://www.python.org/downloads/).
  On Windows/macOS the standard installer already includes the `tkinter` GUI
  library. On Linux install it with e.g. `sudo apt install python3-tk`.
- **LibreOffice** — only needed for PDF export (text replacement works without it).
  - Windows: <https://www.libreoffice.org/download/>
  - macOS: `brew install --cask libreoffice` (or the `.dmg`)
  - Linux: `sudo apt install libreoffice`

No `pip install` is required — the app uses only the Python standard library.

---

## Easiest ways to launch (Windows)

You have three options, from simplest to most self-contained:

1. **Download a ready-made `.exe` (no Python needed).**
   In the GitHub repo, open the **Actions** tab → the latest **"Build Windows
   EXE"** run → under **Artifacts**, download **`VisioTextReplacer-windows`**.
   Unzip it and double-click `VisioTextReplacer.exe`. (PDF export still needs
   LibreOffice installed.)

2. **Double-click `run_visio_tool.bat`** (needs Python installed). It launches
   the app without touching the command line. If Python is missing it tells you
   where to get it.

3. **Build your own `.exe` once** by double-clicking **`build_exe.bat`** (needs
   Python). It installs PyInstaller and produces `dist\VisioTextReplacer.exe`,
   which you can then move/share and run on its own.

## Run the app (GUI, any OS)

```bash
python visio_replace_tool.py
```

A window opens:

1. **Pick your files (step 1 in the window).** Choose **Single file**, or
   **Batch (multiple files)** to process several at once (you can mix `.vsdx`
   and `.xlsx`). Use **Add files...** (select several at once) or **Add
   folder...** (adds every `.vsdx`/`.xlsx` in a folder). The chosen files appear
   in the list; **Remove selected** / **Clear** manage it.
2. **Add your rules (step 2).** Each rule is a row with a separate **Find:** box
   and **Replace with:** box, plus an **in:** dropdown. Use **+ Add another rule**
   for more terms. The **in:** dropdown is a checklist of every file you added:
   - **All files** (the default) — the rule runs on every file.
   - Or untick "All files" and **tick one or more specific files** to run the
     rule on just those.
3. **Set options (step 3)** and click **Replace & Convert**. Each file is saved
   as its next-revision copy (or `*_edited`) next to the original, and the output
   folder opens when it finishes. The Status box lists what happened per file.

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
2. Shows that row's other fields under their headers — **Manufacturer**,
   **Unit Cost**, **Description**, **Qty** (or Quantity), **Notes** (or
   Comments) — pre-filled with the current values.
3. Lets you change any of them. When you click **Save edits**, those changes are
   staged.

On **Replace & Convert**, the staged changes are written to that part number's
row in **every** Excel file that contains it (and on every matching row).
Fields you leave unchanged are not touched, numbers stay numeric, and the P/N
itself is changed by your normal Find → Replace rule.

### The "Change Log" sheet is protected

A sheet named **Change Log** is **never** modified by Find → Replace or by the
BOM row editor — it's left exactly as-is for traceability (even if your search
term appears in a change description there).

To add an entry to it, click **Excel: add Change Log entry...** and fill in the
columns — **Item**, **ECN #** (or **ECO#**), **ERB Approval Date**, **Change
Description**, **Change Author**. On **Replace & Convert**, that row is
**appended to the Change Log sheet of every Excel file**, at the next free row.
The next row is found using the **ECN #** column (so it works even when the Item
numbers run past the last real entry). Leave **Item** blank to keep any
pre-filled item number; leave any field blank to skip it.

### Excel: Author name + date

Many BOM title blocks have an **Author:** box with the author's name in the cell
to its right and a date in the next cell. Click **Excel: set Author + date...**
and type the new name. On **Replace & Convert**, for **every** Excel file the
name beside the **Author** box is set to your value, and the date beside it is
stamped with **today's date** — kept in the **same format** as the date that was
there (an Excel date serial stays a properly-formatted date).

### Change summary (for approval review)

With **Generate change summary** ticked (on by default), the tool writes a
**`Change_Summary_<timestamp>.html`** document next to the edited files and opens
it. For every file it lists each change as **Location · Before · After** —
including text replacements, BOM row edits, the Author name/date, the appended
Change Log row, and the revision bump. Cells are labelled by their column or
title-block label (e.g. *P/N*, *Unit Cost*, *Author*, *Revision*), and dates are
shown as real dates. An approver can review every change in one document without
opening each file, and can print it to PDF.

---

## Command-line use (optional / scriptable)

Passing any arguments runs in CLI mode instead of opening the window. Inputs may
be `.vsdx` or `.xlsx` (detected automatically):

```bash
# Single replacement, edited .vsdx only
python visio_replace_tool.py drawing.vsdx --find "Old Server" --replace "New Host"

# Multiple replacements + PDF export
python visio_replace_tool.py drawing.vsdx \
    --find "Old Server" --replace "New Host" \
    --find "2023"       --replace "2026" \
    --pdf

# Batch: every rule is applied to ALL listed files
python visio_replace_tool.py one.vsdx two.vsdx three.vsdx \
    --find "Old Server" --replace "New Host" --pdf

# Save each copy as the next revision (REVA -> REVB) and bump the in-drawing box
python visio_replace_tool.py "Floor Plan REVA.vsdx" \
    --find "Old Server" --replace "New Host" --bump-revision --pdf

# Choose the output name (single file only); case-sensitive, whole-word
python visio_replace_tool.py in.vsdx -f Dev -r Prod -o out.vsdx \
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
  runs (e.g. `Old ` `<cp/>` `Server`). Search terms are matched across those
  runs, so `Old Server` is still found and replaced.
- The original file is never modified; results are written to new files
  (the next-revision copy, or `*_edited.vsdx` / `*_edited.pdf`).
