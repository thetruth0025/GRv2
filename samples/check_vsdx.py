#!/usr/bin/env python3
"""
Visio text checker -- double-click this file.

It asks you to pick a .vsdx and a piece of text, then reports whether that
text actually exists as editable text inside the drawing (and on which
sheets). This tells us why the replacer tool can or can't find it.
"""
import os
import re
import zipfile
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

TEXT_BLOCK = re.compile(r"<Text\b[^>]*>(.*?)</Text>", re.DOTALL)
TAGS = re.compile(r"<[^>]*>")


def main():
    root = tk.Tk()
    root.withdraw()

    path = filedialog.askopenfilename(
        title="Pick the Visio .vsdx to check",
        filetypes=[("Visio drawing", "*.vsdx"), ("All files", "*.*")],
    )
    if not path:
        return

    term = simpledialog.askstring(
        "Search text", "Text to look for:", initialvalue="SM4233"
    )
    if not term:
        return

    try:
        z = zipfile.ZipFile(path)
    except Exception as exc:  # noqa: BLE001
        messagebox.showerror("Error", f"Could not open file:\n{exc}")
        return

    raw_parts = []
    text_hits = []  # (part, snippet)
    low = term.lower()
    for n in z.namelist():
        ln = n.lower()
        if not (ln.startswith("visio/") and ln.endswith(".xml")):
            continue
        data = z.read(n)
        if term.encode("utf-8") in data:
            raw_parts.append(n)
        try:
            xml = data.decode("utf-8")
        except UnicodeDecodeError:
            continue
        for m in TEXT_BLOCK.finditer(xml):
            visible = TAGS.sub("", m.group(1))
            if low in visible.lower():
                text_hits.append((n, visible.strip()[:90]))

    lines = [
        f"File: {os.path.basename(path)}",
        f"Looking for: {term!r}",
        "",
    ]
    lines.append(
        "Appears in raw file parts: "
        + (", ".join(raw_parts) if raw_parts else "NONE")
    )
    lines.append("")
    if text_hits:
        lines.append(
            f"Found inside {len(text_hits)} text box(es) -- the replacer "
            "WILL replace these:"
        )
        for part, snip in text_hits[:12]:
            lines.append(f"  [{os.path.basename(part)}] {snip}")
    else:
        lines.append("NOT found inside any text box.")
        lines.append(
            "=> The replacer can't change it because it isn't stored as "
            "editable text"
        )
        lines.append(
            "   in THIS file (different file/revision, or the text was "
            "converted to outlines/an image)."
        )

    report = "\n".join(lines)
    messagebox.showinfo("Visio text check", report)
    print(report)


if __name__ == "__main__":
    main()
