# Visio Text Replacer → PDF

A small desktop app that edits Microsoft Visio drawings (`.vsdx`): find text and
replace it **everywhere it appears** in the drawing, then export the result to
PDF.

Workflow:

1. **Choose** a single Visio file (`.vsdx`) — or switch to **Batch** mode and add
   many files (or a whole folder).
2. **Enter** one or more *Find* → *Replace with* rules. Each rule can target
   **all files** or just **specific files** in the batch.
3. Click **Replace & Convert**.
4. Get a **next-revision copy** of each file (and a PDF) next to each original.
   Your originals are never modified.

### Revisions

The tool works on **copies**, never the originals. By default it names each copy
as the **next revision letter**, read from the file name:

- `Floor Plan REVA.vsdx` → `Floor Plan REVB.vsdx` (A→B→C … major changes only).
- It also updates the revision **inside the drawing**: the single-letter box
  (e.g. `A`) that sits next to a `REV` label on page 1 is bumped to match. Other
  stray single letters (grid/zone labels around the border) are left alone — the
  right box is found by its closeness to the `REV` label.
- Already at `REVZ`? That file is skipped with a warning (no next letter).
- No `REVx` in the file name? The copy falls back to `*_edited.vsdx`.

Both behaviors are checkboxes you can turn off (rename only, or no revision bump
at all).

The find/replace edits the Visio file directly (it's a ZIP of XML parts) and only
touches the visible text inside shapes — geometry, colors, themes, connectors and
formatting are left untouched. PDF export uses **LibreOffice**, which renders the
drawing faithfully.

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

1. **Pick your files (step 1 in the window).** Choose **Single file** to work on
   one drawing, or **Batch (multiple files)** to process several at once. In
   batch mode use **Add files...** (select several at once) or **Add folder...**
   (adds every `.vsdx` in a folder). The chosen files appear in the list;
   **Remove selected** / **Clear** manage it.
2. **Add your rules (step 2).** Each rule is a row with a separate **Find:** box
   and **Replace with:** box, plus an **in:** dropdown. Use **+ Add another rule**
   for more terms. The **in:** dropdown lists every file you added:
   - **All files** (the default) — the rule runs on every file.
   - Pick a **file name** to run that rule on only that file. To target a few
     specific files, add one rule per file and point each at its file.
3. **Set options (step 3)** and click **Replace & Convert**. Each file is saved
   as `*_edited.vsdx` (and `*_edited.pdf`) next to the original, and the output
   folder opens when it finishes. The Status box lists what happened per file.

Options:

- **Case sensitive** — match the exact capitalization (off by default).
- **Whole word only** — only match complete words (won't change `foobar` when
  finding `foo`).
- **Also export PDF** — produce a PDF in addition to the edited `.vsdx`.
- **Save copy as next revision (REVx → next)** — name the copy as the next
  revision letter instead of `*_edited` (on by default; see *Revisions* above).
- **...and update the REV box in the drawing** — also bump the `REVx` letter box
  inside the drawing to match the new file name (on by default).

---

## Command-line use (optional / scriptable)

Passing any arguments runs in CLI mode instead of opening the window:

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
