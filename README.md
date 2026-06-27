# Visio Text Replacer → PDF

A small desktop app that edits Microsoft Visio drawings (`.vsdx`): find text and
replace it **everywhere it appears** in the drawing, then export the result to
PDF.

Workflow:

1. **Choose** a Visio file (`.vsdx`).
2. **Enter** one or more *Find* → *Replace with* pairs.
3. Click **Replace & Convert**.
4. Get an edited `*_edited.vsdx` and a `*_edited.pdf` next to your original.

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

## Run the app (GUI)

```bash
python visio_replace_tool.py
```

A window opens. Browse to your `.vsdx`, fill in the Find/Replace boxes (use
**+ Add another pair** for multiple terms), tick the options you want, and click
**Replace & Convert**. The output folder opens automatically when it finishes.

Options:

- **Case sensitive** — match the exact capitalization (off by default).
- **Whole word only** — only match complete words (won't change `foobar` when
  finding `foo`).
- **Also export PDF** — produce a PDF in addition to the edited `.vsdx`.

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

# Choose the output name; case-sensitive, whole-word matching
python visio_replace_tool.py in.vsdx -f Dev -r Prod -o out.vsdx \
    --case-sensitive --whole-word --pdf
```

| Flag | Meaning |
|------|---------|
| `-f`, `--find` | Text to search for (repeatable) |
| `-r`, `--replace` | Replacement, paired with each `--find` in order |
| `-o`, `--output` | Output `.vsdx` path (default: `<name>_edited.vsdx`) |
| `--pdf` | Also export to PDF |
| `--case-sensitive` | Match capitalization exactly |
| `--whole-word` | Only match whole words |

---

## Notes & limitations

- **Format:** supports the modern **`.vsdx`** format. For legacy binary `.vsd`
  files, open them in Visio and *Save As* `.vsdx` first.
- **Text split across formatting runs:** a single label whose characters carry
  *different* inline formatting in the middle of your search term may not be
  matched (this is rare). Whole labels and normally-formatted text replace
  reliably.
- The original file is never modified; results are written to new files
  (`*_edited.vsdx` / `*_edited.pdf`).
