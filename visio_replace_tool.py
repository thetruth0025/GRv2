#!/usr/bin/env python3
"""
Visio Text Replacer -> PDF
==========================

A small desktop application that lets you:

  1. Pick a Visio drawing (.vsdx).
  2. Enter one or more "find" / "replace with" text pairs.
  3. Replace that text everywhere it appears in the drawing.
  4. Save the edited .vsdx and (optionally) export it to PDF.

The find/replace works directly on the .vsdx file format (a ZIP archive of
XML parts). Only the visible text inside Visio's <Text> blocks is touched, so
shape geometry, formatting, themes, connectors, etc. are left untouched.

PDF export is done with LibreOffice (it has a built-in Visio import filter),
which produces a faithful rendering of the drawing.

Run the GUI:
    python visio_replace_tool.py

Or use it from the command line:
    python visio_replace_tool.py drawing.vsdx --find "Old" --replace "New" --pdf

Requirements:
    * Python 3.8+ (tkinter is included with the standard python.org installers)
    * LibreOffice installed (only needed for PDF export)
        - Windows:  https://www.libreoffice.org/download/
        - macOS:    brew install --cask libreoffice  (or the .dmg)
        - Linux:    sudo apt install libreoffice   (or your package manager)
"""

from __future__ import annotations

import argparse
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable, List, Sequence, Tuple

# ---------------------------------------------------------------------------
# Core logic (no GUI dependencies -- safe to import and unit test)
# ---------------------------------------------------------------------------

# Visio text lives in <Text> ... </Text> blocks inside the XML parts of the
# .vsdx archive. We only edit the character data inside those blocks.
_TEXT_BLOCK_RE = re.compile(r"(<Text\b[^>]*>)(.*?)(</Text>)", re.DOTALL)
_TAG_SPLIT_RE = re.compile(r"(<[^>]*>)")


def xml_escape(text: str) -> str:
    """Escape the characters that are special inside XML character data."""
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def replace_in_xml(
    xml_text: str,
    pairs: Sequence[Tuple[str, str]],
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> Tuple[str, int]:
    """Replace text inside <Text> blocks of a single Visio XML part.

    Returns the modified XML and the number of replacements made.

    Only the text *between* tags is modified, so inline formatting markers
    such as <cp/>, <pp/> and <fld/> are preserved. Text the user types is
    XML-escaped before matching so that, e.g., searching for "A & B" matches
    the stored "A &amp; B".
    """
    flags = 0 if case_sensitive else re.IGNORECASE

    # Pre-compile a pattern + escaped replacement for each pair.
    compiled: List[Tuple[re.Pattern, str]] = []
    for find, repl in pairs:
        if not find:
            continue
        pattern = re.escape(xml_escape(find))
        if whole_word:
            pattern = r"\b" + pattern + r"\b"
        compiled.append((re.compile(pattern, flags), xml_escape(repl)))

    if not compiled:
        return xml_text, 0

    total = 0

    def process_block(match: re.Match) -> str:
        nonlocal total
        open_tag, inner, close_tag = match.group(1), match.group(2), match.group(3)
        # Split inner content into [text, tag, text, tag, ...]; text is at even
        # indices, tags (which we must not touch) at odd indices.
        parts = _TAG_SPLIT_RE.split(inner)
        for i in range(0, len(parts), 2):
            segment = parts[i]
            for pattern, repl in compiled:
                segment, n = pattern.subn(lambda _m, r=repl: r, segment)
                total += n
            parts[i] = segment
        return open_tag + "".join(parts) + close_tag

    new_text = _TEXT_BLOCK_RE.sub(process_block, xml_text)
    return new_text, total


def _is_text_part(name: str) -> bool:
    """True for archive members that can contain shape text."""
    lname = name.lower()
    return lname.startswith("visio/") and lname.endswith(".xml")


def replace_text_in_vsdx(
    in_path: str | os.PathLike,
    out_path: str | os.PathLike,
    pairs: Sequence[Tuple[str, str]],
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> dict:
    """Copy a .vsdx applying text replacements; return a report dict.

    Report: {"total": int, "by_part": {part_name: count, ...}}
    """
    in_path = Path(in_path)
    out_path = Path(out_path)

    if not zipfile.is_zipfile(in_path):
        raise ValueError(
            f"'{in_path.name}' is not a .vsdx file. This tool supports the "
            "modern .vsdx format (ZIP-based). Re-save legacy .vsd files as "
            ".vsdx from Visio first."
        )

    by_part: dict[str, int] = {}
    total = 0

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(in_path, "r") as zin, zipfile.ZipFile(
        out_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if _is_text_part(item.filename):
                try:
                    text = data.decode("utf-8")
                except UnicodeDecodeError:
                    pass  # leave non-text/binary-ish parts untouched
                else:
                    new_text, count = replace_in_xml(
                        text, pairs, case_sensitive, whole_word
                    )
                    if count:
                        by_part[item.filename] = count
                        total += count
                        data = new_text.encode("utf-8")
            # Preserve the original name; recompress with deflate.
            zout.writestr(item.filename, data)

    return {"total": total, "by_part": by_part}


def find_libreoffice() -> str | None:
    """Locate the LibreOffice/soffice executable across platforms."""
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found

    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/bin/libreoffice",
        "/usr/local/bin/soffice",
        "/snap/bin/libreoffice",
    ]
    for c in candidates:
        if Path(c).exists():
            return c
    return None


def convert_to_pdf(
    in_path: str | os.PathLike,
    out_dir: str | os.PathLike,
    soffice: str | None = None,
    timeout: int = 180,
) -> Path:
    """Convert a Visio file to PDF using LibreOffice headless mode."""
    in_path = Path(in_path)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    soffice = soffice or find_libreoffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice was not found. Install it from "
            "https://www.libreoffice.org/download/ to enable PDF export."
        )

    # A private, throwaway user profile lets us run even if a normal
    # LibreOffice window is already open.
    profile_dir = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        cmd = [
            soffice,
            f"-env:UserInstallation={Path(profile_dir).as_uri()}",
            "--headless",
            "--norestore",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(in_path),
        ]
        proc = subprocess.run(
            cmd, capture_output=True, text=True, timeout=timeout
        )
        pdf_path = out_dir / (in_path.stem + ".pdf")
        if not pdf_path.exists():
            raise RuntimeError(
                "PDF conversion failed.\n"
                f"LibreOffice stdout: {proc.stdout.strip()}\n"
                f"LibreOffice stderr: {proc.stderr.strip()}"
            )
        return pdf_path
    finally:
        shutil.rmtree(profile_dir, ignore_errors=True)


# ---------------------------------------------------------------------------
# Command-line interface
# ---------------------------------------------------------------------------

def _run_cli(argv: Sequence[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Find/replace text in a Visio .vsdx file and optionally "
        "export it to PDF.",
    )
    parser.add_argument("input", help="Path to the input .vsdx file")
    parser.add_argument(
        "-f", "--find", action="append", default=[],
        help="Text to search for (repeatable)",
    )
    parser.add_argument(
        "-r", "--replace", action="append", default=[],
        help="Replacement text, paired with --find in order (repeatable)",
    )
    parser.add_argument(
        "-o", "--output",
        help="Output .vsdx path (default: <name>_edited.vsdx beside input)",
    )
    parser.add_argument(
        "--pdf", action="store_true", help="Also export the result to PDF",
    )
    parser.add_argument(
        "--case-sensitive", action="store_true",
        help="Match case exactly (default: case-insensitive)",
    )
    parser.add_argument(
        "--whole-word", action="store_true",
        help="Only match whole words",
    )
    args = parser.parse_args(argv)

    if len(args.find) != len(args.replace):
        parser.error("each --find must be paired with a --replace")
    if not args.find:
        parser.error("provide at least one --find/--replace pair")

    in_path = Path(args.input)
    if not in_path.exists():
        parser.error(f"input file not found: {in_path}")

    out_path = (
        Path(args.output)
        if args.output
        else in_path.with_name(in_path.stem + "_edited.vsdx")
    )
    pairs = list(zip(args.find, args.replace))

    try:
        report = replace_text_in_vsdx(
            in_path, out_path, pairs,
            case_sensitive=args.case_sensitive, whole_word=args.whole_word,
        )
        print(f"Replaced {report['total']} occurrence(s); saved -> {out_path}")
        for part, n in report["by_part"].items():
            print(f"    {part}: {n}")
        if report["total"] == 0:
            print(
                "Warning: no matches found. Check spelling, or try "
                "without --case-sensitive."
            )

        if args.pdf:
            pdf = convert_to_pdf(out_path, out_path.parent)
            print(f"PDF written -> {pdf}")
    except (ValueError, RuntimeError) as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    return 0


# ---------------------------------------------------------------------------
# Tkinter GUI
# ---------------------------------------------------------------------------

def launch_gui() -> int:
    # Imported lazily so the core logic / CLI work without a display.
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    class App:
        def __init__(self, root: "tk.Tk"):
            self.root = root
            root.title("Visio Text Replacer  ->  PDF")
            root.geometry("680x640")
            root.minsize(560, 520)

            self.pair_rows: List[Tuple[tk.Entry, tk.Entry, tk.Frame]] = []

            pad = {"padx": 10, "pady": 6}

            # --- Input file -------------------------------------------------
            top = ttk.LabelFrame(root, text="1.  Visio file (.vsdx)")
            top.pack(fill="x", **pad)
            self.file_var = tk.StringVar()
            ttk.Entry(top, textvariable=self.file_var).pack(
                side="left", fill="x", expand=True, padx=8, pady=8
            )
            ttk.Button(top, text="Browse...", command=self.browse).pack(
                side="left", padx=8, pady=8
            )

            # --- Find / replace pairs --------------------------------------
            mid = ttk.LabelFrame(root, text="2.  Find  ->  Replace with")
            mid.pack(fill="x", **pad)
            self.pairs_frame = ttk.Frame(mid)
            self.pairs_frame.pack(fill="x", padx=6, pady=6)
            self._header_row()
            self.add_pair()
            ttk.Button(
                mid, text="+ Add another pair", command=self.add_pair
            ).pack(anchor="w", padx=8, pady=(0, 8))

            # --- Options ----------------------------------------------------
            opts = ttk.LabelFrame(root, text="3.  Options")
            opts.pack(fill="x", **pad)
            self.case_var = tk.BooleanVar(value=False)
            self.word_var = tk.BooleanVar(value=False)
            self.pdf_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(
                opts, text="Case sensitive", variable=self.case_var
            ).pack(side="left", padx=8, pady=8)
            ttk.Checkbutton(
                opts, text="Whole word only", variable=self.word_var
            ).pack(side="left", padx=8, pady=8)
            ttk.Checkbutton(
                opts, text="Also export PDF", variable=self.pdf_var
            ).pack(side="left", padx=8, pady=8)

            # --- Run --------------------------------------------------------
            run = ttk.Frame(root)
            run.pack(fill="x", **pad)
            self.run_btn = ttk.Button(
                run, text="Replace  &  Convert", command=self.run
            )
            self.run_btn.pack(side="left", padx=8)
            self.progress = ttk.Progressbar(run, mode="indeterminate")
            self.progress.pack(side="left", fill="x", expand=True, padx=8)

            # --- Log --------------------------------------------------------
            logf = ttk.LabelFrame(root, text="Status")
            logf.pack(fill="both", expand=True, **pad)
            self.log = scrolledtext.ScrolledText(
                logf, height=8, state="disabled", wrap="word"
            )
            self.log.pack(fill="both", expand=True, padx=6, pady=6)

            if find_libreoffice() is None:
                self._log(
                    "Note: LibreOffice was not found, so PDF export is "
                    "unavailable. Install it from libreoffice.org to enable "
                    "PDF output. Text replacement still works."
                )

        # -- pair rows ------------------------------------------------------
        def _header_row(self):
            hdr = ttk.Frame(self.pairs_frame)
            hdr.pack(fill="x")
            ttk.Label(hdr, text="Find", width=30).pack(side="left", padx=4)
            ttk.Label(hdr, text="Replace with", width=30).pack(
                side="left", padx=4
            )

        def add_pair(self):
            row = ttk.Frame(self.pairs_frame)
            row.pack(fill="x", pady=2)
            find_e = ttk.Entry(row, width=30)
            find_e.pack(side="left", fill="x", expand=True, padx=4)
            repl_e = ttk.Entry(row, width=30)
            repl_e.pack(side="left", fill="x", expand=True, padx=4)
            rm = ttk.Button(
                row, text="X", width=3,
                command=lambda: self.remove_pair(find_e, repl_e, row),
            )
            rm.pack(side="left", padx=4)
            self.pair_rows.append((find_e, repl_e, row))

        def remove_pair(self, find_e, repl_e, row):
            if len(self.pair_rows) <= 1:
                return  # keep at least one row
            row.destroy()
            self.pair_rows = [
                t for t in self.pair_rows if t[2] is not row
            ]

        def collect_pairs(self) -> List[Tuple[str, str]]:
            pairs = []
            for find_e, repl_e, _ in self.pair_rows:
                f = find_e.get()
                if f:
                    pairs.append((f, repl_e.get()))
            return pairs

        # -- helpers --------------------------------------------------------
        def browse(self):
            path = filedialog.askopenfilename(
                title="Choose a Visio file",
                filetypes=[("Visio drawing", "*.vsdx"), ("All files", "*.*")],
            )
            if path:
                self.file_var.set(path)

        def _log(self, msg: str):
            self.log.configure(state="normal")
            self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")

        def log(self, msg: str):
            # Safe to call from worker thread.
            self.root.after(0, self._log, msg)

        def _set_busy(self, busy: bool):
            self.run_btn.configure(state="disabled" if busy else "normal")
            if busy:
                self.progress.start(12)
            else:
                self.progress.stop()

        # -- run ------------------------------------------------------------
        def run(self):
            in_path = self.file_var.get().strip()
            if not in_path:
                messagebox.showwarning(
                    "No file", "Please choose a .vsdx file first."
                )
                return
            if not Path(in_path).exists():
                messagebox.showerror("Not found", f"File not found:\n{in_path}")
                return
            pairs = self.collect_pairs()
            if not pairs:
                messagebox.showwarning(
                    "Nothing to find",
                    "Enter at least one 'Find' value.",
                )
                return

            self._set_busy(True)
            threading.Thread(
                target=self._worker,
                args=(
                    in_path, pairs, self.case_var.get(),
                    self.word_var.get(), self.pdf_var.get(),
                ),
                daemon=True,
            ).start()

        def _worker(self, in_path, pairs, case_sensitive, whole_word, make_pdf):
            try:
                src = Path(in_path)
                out_vsdx = src.with_name(src.stem + "_edited.vsdx")
                self.log(f"Reading {src.name} ...")
                report = replace_text_in_vsdx(
                    src, out_vsdx, pairs,
                    case_sensitive=case_sensitive, whole_word=whole_word,
                )
                self.log(
                    f"Replaced {report['total']} occurrence(s)."
                )
                for part, n in report["by_part"].items():
                    self.log(f"    {part}: {n}")
                self.log(f"Saved edited drawing -> {out_vsdx}")

                if report["total"] == 0:
                    self.log(
                        "Warning: no matches were found. Check spelling, or "
                        "toggle 'Case sensitive'."
                    )

                pdf_path = None
                if make_pdf:
                    self.log("Converting to PDF with LibreOffice ...")
                    pdf_path = convert_to_pdf(out_vsdx, out_vsdx.parent)
                    self.log(f"PDF written -> {pdf_path}")

                self.log("Done.")
                final = str(pdf_path or out_vsdx)
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Finished",
                        f"All done!\n\nEdited file:\n{out_vsdx}"
                        + (f"\n\nPDF:\n{pdf_path}" if pdf_path else ""),
                    ),
                )
                self._reveal(final)
            except Exception as exc:  # noqa: BLE001 - surface to the user
                self.log(f"ERROR: {exc}")
                self.root.after(
                    0, lambda: messagebox.showerror("Error", str(exc))
                )
            finally:
                self.root.after(0, self._set_busy, False)

        def _reveal(self, path: str):
            """Open the folder containing the result, best-effort."""
            folder = str(Path(path).parent)
            try:
                if sys.platform.startswith("win"):
                    os.startfile(folder)  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", folder])
                else:
                    subprocess.Popen(["xdg-open", folder])
            except Exception:
                pass

    root = tk.Tk()
    App(root)
    root.mainloop()
    return 0


# ---------------------------------------------------------------------------

def main(argv: Sequence[str] | None = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    # No arguments -> launch the GUI. Any arguments -> CLI mode.
    if not argv:
        try:
            return launch_gui()
        except Exception as exc:  # pragma: no cover - environment dependent
            print(f"Could not start the GUI ({exc}).", file=sys.stderr)
            print(
                "Use the command line instead, e.g.:\n"
                "  python visio_replace_tool.py drawing.vsdx "
                '--find "Old" --replace "New" --pdf',
                file=sys.stderr,
            )
            return 1
    return _run_cli(argv)


if __name__ == "__main__":
    raise SystemExit(main())
