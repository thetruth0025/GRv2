#!/usr/bin/env python3
"""
Visio Text Replacer -> PDF
==========================

A small desktop application that lets you:

  1. Pick a single Visio drawing (.vsdx) OR a batch of them.
  2. Enter one or more "find" / "replace with" text rules.
  3. Aim each rule at *all* files or only *specific* files in the batch.
  4. Replace that text everywhere it appears in the drawing(s).
  5. Save the edited .vsdx file(s) and (optionally) export them to PDF.

The find/replace works directly on the .vsdx file format (a ZIP archive of
XML parts). Only the visible text inside Visio's <Text> blocks is touched, so
shape geometry, formatting, themes, connectors, etc. are left untouched.

PDF export is done with LibreOffice (it has a built-in Visio import filter),
which produces a faithful rendering of the drawing.

Run the GUI:
    python visio_replace_tool.py

Or use it from the command line (one or many files; rules apply to all of them):
    python visio_replace_tool.py a.vsdx b.vsdx --find "Old" --replace "New" --pdf

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
        description="Find/replace text in one or more Visio .vsdx files and "
        "optionally export them to PDF. When several files are given, every "
        "--find/--replace rule is applied to all of them (use the GUI for "
        "per-file targeting).",
    )
    parser.add_argument(
        "inputs", nargs="+", help="One or more input .vsdx files",
    )
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
        help="Output .vsdx path (only valid with a single input; default: "
        "<name>_edited.vsdx beside each input)",
    )
    parser.add_argument(
        "--pdf", action="store_true", help="Also export the result(s) to PDF",
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
    if args.output and len(args.inputs) > 1:
        parser.error("--output cannot be used with multiple input files")

    pairs = list(zip(args.find, args.replace))
    had_error = False

    for raw in args.inputs:
        in_path = Path(raw)
        if not in_path.exists():
            print(f"Error: input file not found: {in_path}", file=sys.stderr)
            had_error = True
            continue

        out_path = (
            Path(args.output)
            if args.output
            else in_path.with_name(in_path.stem + "_edited.vsdx")
        )
        try:
            report = replace_text_in_vsdx(
                in_path, out_path, pairs,
                case_sensitive=args.case_sensitive, whole_word=args.whole_word,
            )
            print(
                f"{in_path.name}: replaced {report['total']} occurrence(s) "
                f"-> {out_path}"
            )
            if report["total"] == 0:
                print(
                    "    (no matches found; check spelling or "
                    "--case-sensitive)"
                )
            if args.pdf:
                pdf = convert_to_pdf(out_path, out_path.parent)
                print(f"    PDF -> {pdf}")
        except (ValueError, RuntimeError) as exc:
            print(f"Error ({in_path.name}): {exc}", file=sys.stderr)
            had_error = True

    return 1 if had_error else 0


# ---------------------------------------------------------------------------
# Tkinter GUI
# ---------------------------------------------------------------------------

def launch_gui() -> int:
    # Imported lazily so the core logic / CLI work without a display.
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    # Sentinel meaning "this rule applies to every file (now and later)".
    ALL_FILES = "ALL"

    class App:
        def __init__(self, root: "tk.Tk"):
            self.root = root
            root.title("Visio Text Replacer  ->  PDF")
            root.geometry("780x760")
            root.minsize(640, 640)

            self.files: List[str] = []
            # Each rule: {"frame","find","repl","scope","btn"}; scope is
            # ALL_FILES or a set of file paths.
            self.rule_rows: List[dict] = []

            pad = {"padx": 10, "pady": 6}

            # --- 1. Files --------------------------------------------------
            top = ttk.LabelFrame(root, text="1.  Visio files (.vsdx)")
            top.pack(fill="x", **pad)

            mode_row = ttk.Frame(top)
            mode_row.pack(fill="x", padx=8, pady=(8, 2))
            self.mode_var = tk.StringVar(value="single")
            ttk.Radiobutton(
                mode_row, text="Single file", value="single",
                variable=self.mode_var, command=self.on_mode_change,
            ).pack(side="left", padx=(0, 12))
            ttk.Radiobutton(
                mode_row, text="Batch (multiple files)", value="batch",
                variable=self.mode_var, command=self.on_mode_change,
            ).pack(side="left")

            btn_row = ttk.Frame(top)
            btn_row.pack(fill="x", padx=8, pady=2)
            self.add_btn = ttk.Button(
                btn_row, text="Add file...", command=self.add_files
            )
            self.add_btn.pack(side="left")
            self.folder_btn = ttk.Button(
                btn_row, text="Add folder...", command=self.add_folder
            )
            self.folder_btn.pack(side="left", padx=6)
            ttk.Button(
                btn_row, text="Remove selected",
                command=self.remove_selected_files,
            ).pack(side="left", padx=6)
            ttk.Button(
                btn_row, text="Clear", command=self.clear_files
            ).pack(side="left")

            list_row = ttk.Frame(top)
            list_row.pack(fill="x", padx=8, pady=(2, 8))
            self.files_box = tk.Listbox(
                list_row, height=5, selectmode="extended",
                activestyle="none",
            )
            self.files_box.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(
                list_row, orient="vertical", command=self.files_box.yview
            )
            sb.pack(side="left", fill="y")
            self.files_box.configure(yscrollcommand=sb.set)

            # --- 2. Find / replace rules -----------------------------------
            mid = ttk.LabelFrame(
                root, text="2.  Find  ->  Replace with  (and which files)"
            )
            mid.pack(fill="x", **pad)
            self.pairs_frame = ttk.Frame(mid)
            self.pairs_frame.pack(fill="x", padx=6, pady=6)
            self._header_row()
            self.add_pair()
            ttk.Button(
                mid, text="+ Add another rule", command=self.add_pair
            ).pack(anchor="w", padx=8, pady=(0, 8))

            # --- 3. Options ------------------------------------------------
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

            # --- Run -------------------------------------------------------
            run = ttk.Frame(root)
            run.pack(fill="x", **pad)
            self.run_btn = ttk.Button(
                run, text="Replace  &  Convert", command=self.run
            )
            self.run_btn.pack(side="left", padx=8)
            self.progress = ttk.Progressbar(run, mode="indeterminate")
            self.progress.pack(side="left", fill="x", expand=True, padx=8)

            # --- Log -------------------------------------------------------
            logf = ttk.LabelFrame(root, text="Status")
            logf.pack(fill="both", expand=True, **pad)
            self.log = scrolledtext.ScrolledText(
                logf, height=8, state="disabled", wrap="word"
            )
            self.log.pack(fill="both", expand=True, padx=6, pady=6)

            self.on_mode_change()
            if find_libreoffice() is None:
                self._log(
                    "Note: LibreOffice was not found, so PDF export is "
                    "unavailable. Install it from libreoffice.org to enable "
                    "PDF output. Text replacement still works."
                )

        # -- file list ------------------------------------------------------
        def on_mode_change(self):
            single = self.mode_var.get() == "single"
            self.add_btn.configure(text="Choose file..." if single else "Add files...")
            self.folder_btn.configure(state="disabled" if single else "normal")
            if single and len(self.files) > 1:
                self.files = self.files[:1]
                self.refresh_files_box()
                self.log("Single-file mode: keeping only the first file.")

        def add_files(self):
            ft = [("Visio drawing", "*.vsdx"), ("All files", "*.*")]
            if self.mode_var.get() == "single":
                path = filedialog.askopenfilename(
                    title="Choose a Visio file", filetypes=ft
                )
                self.files = [path] if path else self.files[:0]
            else:
                paths = filedialog.askopenfilenames(
                    title="Choose Visio files", filetypes=ft
                )
                for p in paths:
                    if p not in self.files:
                        self.files.append(p)
            self.refresh_files_box()

        def add_folder(self):
            folder = filedialog.askdirectory(title="Choose a folder of .vsdx files")
            if not folder:
                return
            found = sorted(str(p) for p in Path(folder).glob("*.vsdx"))
            for p in found:
                if p not in self.files:
                    self.files.append(p)
            if not found:
                messagebox.showinfo(
                    "No files", "No .vsdx files were found in that folder."
                )
            self.refresh_files_box()

        def remove_selected_files(self):
            for i in reversed(self.files_box.curselection()):
                del self.files[i]
            self.refresh_files_box()
            self._refresh_scope_buttons()

        def clear_files(self):
            self.files = []
            self.refresh_files_box()
            self._refresh_scope_buttons()

        def refresh_files_box(self):
            self.files_box.delete(0, "end")
            for f in self.files:
                self.files_box.insert("end", Path(f).name)
            self._refresh_scope_buttons()

        # -- rule rows ------------------------------------------------------
        def _header_row(self):
            hdr = ttk.Frame(self.pairs_frame)
            hdr.pack(fill="x")
            ttk.Label(hdr, text="Find", width=24).pack(side="left", padx=4)
            ttk.Label(hdr, text="Replace with", width=24).pack(
                side="left", padx=4
            )
            ttk.Label(hdr, text="Applies to", width=16).pack(
                side="left", padx=4
            )

        def add_pair(self):
            row = ttk.Frame(self.pairs_frame)
            row.pack(fill="x", pady=2)
            find_e = ttk.Entry(row, width=24)
            find_e.pack(side="left", fill="x", expand=True, padx=4)
            repl_e = ttk.Entry(row, width=24)
            repl_e.pack(side="left", fill="x", expand=True, padx=4)
            scope_btn = ttk.Button(row, width=16)
            scope_btn.pack(side="left", padx=4)
            rule = {
                "frame": row, "find": find_e, "repl": repl_e,
                "scope": ALL_FILES, "btn": scope_btn,
            }
            scope_btn.configure(command=lambda r=rule: self.edit_scope(r))
            ttk.Button(
                row, text="X", width=3,
                command=lambda r=rule: self.remove_pair(r),
            ).pack(side="left", padx=4)
            self.rule_rows.append(rule)
            self._update_scope_button(rule)

        def remove_pair(self, rule):
            if len(self.rule_rows) <= 1:
                return  # keep at least one row
            rule["frame"].destroy()
            self.rule_rows = [r for r in self.rule_rows if r is not rule]

        def _update_scope_button(self, rule):
            scope = rule["scope"]
            if scope == ALL_FILES:
                rule["btn"].configure(text="All files")
            else:
                n = len(scope)
                rule["btn"].configure(
                    text=("No files" if n == 0 else f"{n} file(s)")
                )

        def _refresh_scope_buttons(self):
            # Drop file paths that no longer exist in the list from each scope.
            current = set(self.files)
            for rule in self.rule_rows:
                if rule["scope"] != ALL_FILES:
                    rule["scope"] = {f for f in rule["scope"] if f in current}
                self._update_scope_button(rule)

        def edit_scope(self, rule):
            if not self.files:
                messagebox.showinfo(
                    "Add files first",
                    "Add one or more Visio files before choosing which ones "
                    "this rule applies to.",
                )
                return

            win = tk.Toplevel(self.root)
            win.title("Apply this rule to...")
            win.transient(self.root)
            win.grab_set()

            all_var = tk.BooleanVar(value=(rule["scope"] == ALL_FILES))
            ttk.Checkbutton(
                win, text="All files (including any added later)",
                variable=all_var,
            ).pack(anchor="w", padx=12, pady=(12, 4))
            ttk.Label(
                win, text="...or pick specific files:"
            ).pack(anchor="w", padx=12)

            box = tk.Listbox(win, selectmode="extended", height=8, width=48)
            box.pack(fill="both", expand=True, padx=12, pady=6)
            for f in self.files:
                box.insert("end", Path(f).name)
            if rule["scope"] != ALL_FILES:
                for i, f in enumerate(self.files):
                    if f in rule["scope"]:
                        box.selection_set(i)

            def sync_state(*_):
                box.configure(state="disabled" if all_var.get() else "normal")
            all_var.trace_add("write", sync_state)
            sync_state()

            def ok():
                if all_var.get():
                    rule["scope"] = ALL_FILES
                else:
                    rule["scope"] = {
                        self.files[i] for i in box.curselection()
                    }
                self._update_scope_button(rule)
                win.destroy()

            btns = ttk.Frame(win)
            btns.pack(fill="x", padx=12, pady=(0, 12))
            ttk.Button(btns, text="OK", command=ok).pack(side="right")
            ttk.Button(
                btns, text="Cancel", command=win.destroy
            ).pack(side="right", padx=6)

        # -- helpers --------------------------------------------------------
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

        def pairs_for_file(self, path: str) -> List[Tuple[str, str]]:
            """Find/replace pairs whose scope includes this file."""
            pairs = []
            for rule in self.rule_rows:
                find = rule["find"].get()
                if not find:
                    continue
                scope = rule["scope"]
                if scope == ALL_FILES or path in scope:
                    pairs.append((find, rule["repl"].get()))
            return pairs

        # -- run ------------------------------------------------------------
        def run(self):
            if not self.files:
                messagebox.showwarning(
                    "No files", "Please add at least one .vsdx file."
                )
                return
            missing = [f for f in self.files if not Path(f).exists()]
            if missing:
                messagebox.showerror(
                    "Not found",
                    "These files no longer exist:\n" + "\n".join(missing),
                )
                return
            if not any(r["find"].get() for r in self.rule_rows):
                messagebox.showwarning(
                    "Nothing to find", "Enter at least one 'Find' value."
                )
                return

            self._set_busy(True)
            threading.Thread(
                target=self._worker,
                args=(
                    list(self.files),
                    {f: self.pairs_for_file(f) for f in self.files},
                    self.case_var.get(), self.word_var.get(),
                    self.pdf_var.get(),
                ),
                daemon=True,
            ).start()

        def _worker(self, files, pairs_by_file, case_sensitive, whole_word,
                    make_pdf):
            total_repl = 0
            done = 0
            errors = 0
            last_output = None
            try:
                for src in (Path(f) for f in files):
                    pairs = pairs_by_file.get(str(src), [])
                    if not pairs:
                        self.log(
                            f"- {src.name}: skipped (no rule targets this file)"
                        )
                        continue
                    try:
                        out_vsdx = src.with_name(src.stem + "_edited.vsdx")
                        report = replace_text_in_vsdx(
                            src, out_vsdx, pairs,
                            case_sensitive=case_sensitive,
                            whole_word=whole_word,
                        )
                        total_repl += report["total"]
                        msg = (
                            f"+ {src.name}: {report['total']} replacement(s) "
                            f"-> {out_vsdx.name}"
                        )
                        if report["total"] == 0:
                            msg += "  (no matches found)"
                        self.log(msg)
                        last_output = out_vsdx

                        if make_pdf:
                            pdf_path = convert_to_pdf(out_vsdx, out_vsdx.parent)
                            self.log(f"    PDF -> {pdf_path.name}")
                            last_output = pdf_path
                        done += 1
                    except Exception as exc:  # noqa: BLE001
                        errors += 1
                        self.log(f"! {src.name}: ERROR - {exc}")

                summary = (
                    f"Done. {done} file(s) processed, "
                    f"{total_repl} total replacement(s)"
                    + (f", {errors} error(s)" if errors else "")
                    + "."
                )
                self.log(summary)
                self.root.after(
                    0, lambda: messagebox.showinfo("Finished", summary)
                )
                if last_output is not None:
                    self._reveal(str(last_output))
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
