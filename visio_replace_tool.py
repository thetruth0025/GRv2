#!/usr/bin/env python3
"""
Visio / Excel Text Replacer
===========================

A small desktop application that lets you:

  1. Pick a single Visio drawing (.vsdx) or Excel workbook (.xlsx) -- or a
     batch mixing both. The file type is detected automatically per file.
  2. Enter one or more "find" / "replace with" text rules.
  3. Aim each rule at *all* files or only *specific* files (multi-select).
  4. Replace that text everywhere it appears.
  5. Save the edited copy of each file (named as the next revision), and
     optionally export to PDF.

The find/replace works directly on the file (both .vsdx and .xlsx are ZIP
archives of XML). For Visio, only the text in <Text> shape blocks is touched;
for Excel, the shared-strings table, inline strings, and drawing text boxes.
Geometry, formatting, themes, styles, etc. are left untouched, and a search
term is matched even when it is split across formatting runs.

PDF export is done with LibreOffice. Note: LibreOffice does not evaluate
field-based values (e.g. a Visio "Sheet X of Y" page-number field), so export
from the source app for those; PDF export is off by default.

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

__version__ = "1.9 (Change Log protect + append)"

import argparse
import os
import posixpath
import re
import shutil
import string
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

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


def _splice_segments(segs: List[str], start: int, end: int, repl: str) -> None:
    """Replace full-text range [start, end) across a list of text runs.

    ``segs`` is edited in place. The replacement text lands in the first run
    the match touches; characters in later runs the match spans are removed.
    This lets a search term match even when Visio split it across formatting
    runs (e.g. "Old " <cp/> "Server").
    """
    pos = 0
    first = True
    for k in range(len(segs)):
        seg = segs[k]
        seg_start = pos
        seg_end = pos + len(seg)
        pos = seg_end
        if seg_end <= start or seg_start >= end:
            continue  # this run is entirely outside the match
        a = max(start, seg_start) - seg_start  # local removal start
        b = min(end, seg_end) - seg_start      # local removal end
        if first:
            segs[k] = seg[:a] + repl + seg[b:]
            first = False
        else:
            segs[k] = seg[:a] + seg[b:]


def _compile_pairs(
    pairs: Sequence[Tuple[str, str]],
    case_sensitive: bool,
    whole_word: bool,
    escape: bool = True,
) -> List[Tuple["re.Pattern", str]]:
    """Compile find/replace pairs into (pattern, replacement).

    With escape=True the find/replace are XML-escaped (for matching inside raw
    XML run text); with escape=False they are used as-is (for matching against
    already-decoded cell text).
    """
    flags = 0 if case_sensitive else re.IGNORECASE
    compiled: List[Tuple[re.Pattern, str]] = []
    for find, repl in pairs:
        if not find:
            continue
        f = xml_escape(find) if escape else find
        r = xml_escape(repl) if escape else repl
        pattern = re.escape(f)
        if whole_word:
            pattern = r"\b" + pattern + r"\b"
        compiled.append((re.compile(pattern, flags), r))
    return compiled


def _replace_across_runs(segs: List[str], compiled) -> int:
    """Run-aware replace over a list of text runs (edited in place).

    Matches against the concatenation of the runs so a term can span them,
    then splices replacements back. Returns the number of replacements.
    """
    total = 0
    for pattern, repl in compiled:
        full = "".join(segs)
        matches = list(pattern.finditer(full))
        if not matches:
            continue
        for m in reversed(matches):  # right-to-left keeps offsets valid
            _splice_segments(segs, m.start(), m.end(), repl)
        total += len(matches)
    return total


def replace_in_xml(
    xml_text: str,
    pairs: Sequence[Tuple[str, str]],
    case_sensitive: bool = True,
    whole_word: bool = False,
) -> Tuple[str, int]:
    """Replace text inside <Text> blocks of a single Visio XML part.

    Returns the modified XML and the number of replacements made.

    Inline formatting markers (<cp/>, <pp/>, <fld/>, ...) are preserved, and a
    search term is matched even when it spans several of them. Text the user
    types is XML-escaped before matching so that, e.g., searching for "A & B"
    matches the stored "A &amp; B".
    """
    compiled = _compile_pairs(pairs, case_sensitive, whole_word)
    if not compiled:
        return xml_text, 0

    total = 0

    def process_block(match: re.Match) -> str:
        nonlocal total
        open_tag, inner, close_tag = match.group(1), match.group(2), match.group(3)
        # Split inner into [text, tag, text, tag, ...]; text runs are at even
        # indices, tags (which we must not touch) at odd indices.
        parts = _TAG_SPLIT_RE.split(inner)
        text_indices = list(range(0, len(parts), 2))
        segs = [parts[i] for i in text_indices]

        total += _replace_across_runs(segs, compiled)

        for idx, i in enumerate(text_indices):
            parts[i] = segs[idx]
        return open_tag + "".join(parts) + close_tag

    new_text = _TEXT_BLOCK_RE.sub(process_block, xml_text)
    return new_text, total


def _is_text_part(name: str) -> bool:
    """True for archive members that can contain shape text."""
    lname = name.lower()
    return lname.startswith("visio/") and lname.endswith(".xml")


# ---------------------------------------------------------------------------
# Revision handling
# ---------------------------------------------------------------------------
#
# Drawings are versioned with a revision letter (A, B, C, ... Z), major changes
# only. It shows up two ways:
#   * In the FILE NAME, e.g. "Floor Plan REVA.vsdx".
#   * In the DRAWING, as a single letter ("A") in its own text box, sitting
#     next to a separate box labelled "REV" (a typical title block).
# We read the current letter from the file name and bump it to the next one.

# "REV" + optional space/dash/underscore + a single letter, not part of a word.
_FILENAME_REV_RE = re.compile(r"(?i)(REV)([ _-]?)([A-Za-z])(?![A-Za-z])")
_PAGE_NAME_RE = re.compile(r"visio/pages/page(\d+)\.xml$", re.IGNORECASE)


def next_revision_letter(letter: str) -> Optional[str]:
    """Return the next revision letter (A->B ... Y->Z), or None past Z."""
    idx = string.ascii_uppercase.find(letter.upper())
    if idx < 0 or idx >= 25:
        return None
    return string.ascii_uppercase[idx + 1]


def revision_output_path(in_path: str | os.PathLike):
    """Work out the next-revision copy name for a file.

    Returns (out_path, old_letter, new_letter, status) where status is:
      'ok'      -> out_path is the bumped-revision name
      'no_rev'  -> no REVx in the name; out_path falls back to *_edited.vsdx
      'at_z'    -> already at REVZ; out_path is None (caller should skip)
    """
    p = Path(in_path)
    match = list(_FILENAME_REV_RE.finditer(p.stem))
    if not match:
        fallback = p.with_name(p.stem + "_edited" + p.suffix)
        return fallback, None, None, "no_rev"

    m = match[-1]  # last occurrence is the revision marker
    old = m.group(3).upper()
    nxt = next_revision_letter(old)
    if nxt is None:
        return None, old, None, "at_z"

    new_stem = p.stem[: m.start(3)] + nxt + p.stem[m.end(3):]
    return p.with_name(new_stem + p.suffix), old, nxt, "ok"


def _first_page_part(names: Sequence[str]) -> Optional[str]:
    """The archive member for page 1 (lowest-numbered page)."""
    pages = [n for n in names if _PAGE_NAME_RE.search(n)]
    if not pages:
        return None
    return min(pages, key=lambda n: int(_PAGE_NAME_RE.search(n).group(1)))


def _cell_value(shape_el, ns: str, name: str) -> Optional[float]:
    for cell in shape_el.findall(ns + "Cell"):
        if cell.get("N") == name:
            try:
                return float(cell.get("V"))
            except (TypeError, ValueError):
                return None
    return None


def _is_rev_label(text: str) -> bool:
    """True for a revision label box: REV, Rev, REV., Revision, Revision: ..."""
    norm = text.strip().strip(".:").strip().upper()
    return norm in ("REV", "REVISION")


def _choose_rev_candidate(page_xml: str, old_letter: str) -> Optional[int]:
    """Pick which single-letter box (in document order) is THE revision.

    The revision letter sits next to a "REV"/"Revision" label in the title
    block (lower-right of the front page). We pick the matching letter box
    closest to such a label; if that can't be resolved, we fall back to the
    box nearest the lower-right corner. Returns the index among same-letter
    boxes, or None if it can't be determined.
    """
    try:
        root = ET.fromstring(page_xml)
    except ET.ParseError:
        return None
    ns = root.tag[: root.tag.index("}") + 1] if root.tag.startswith("{") else ""

    candidates = []  # (pinx, piny) for single-letter boxes == old_letter
    rev_labels = []  # (pinx, piny) for REV/Revision label boxes
    for sh in root.iter(ns + "Shape"):
        text_el = sh.find(ns + "Text")
        if text_el is None:
            continue
        norm = "".join(text_el.itertext()).strip()
        if _is_rev_label(norm):
            rev_labels.append(
                (_cell_value(sh, ns, "PinX"), _cell_value(sh, ns, "PinY"))
            )
        elif len(norm) == 1 and norm.upper() == old_letter.upper():
            candidates.append(
                (_cell_value(sh, ns, "PinX"), _cell_value(sh, ns, "PinY"))
            )

    # Primary: the matching letter box closest to a REV/Revision label.
    if rev_labels:
        best_idx, best_dist = None, None
        for i, (px, py) in enumerate(candidates):
            if px is None or py is None:
                continue
            for rx, ry in rev_labels:
                if rx is None or ry is None:
                    continue
                dist = (px - rx) ** 2 + (py - ry) ** 2
                if best_dist is None or dist < best_dist:
                    best_dist, best_idx = dist, i
        if best_idx is not None:
            return best_idx

    # Fallback: the candidate nearest the lower-right corner of the page
    # (large X, small Y in Visio coordinates), where the title block lives.
    best_idx, best_score = None, None
    for i, (px, py) in enumerate(candidates):
        if px is None or py is None:
            continue
        score = px - py
        if best_score is None or score > best_score:
            best_score, best_idx = score, i
    return best_idx


def bump_revision_in_page(page_xml: str, old_letter: str, new_letter: str):
    """Update the revision-letter box on a page. Returns (xml, status).

    status: 'updated' | 'not_found' | 'ambiguous'
    """
    old_u = old_letter.upper()
    # Single-letter <Text> boxes matching the old revision, in document order.
    cands = []
    for m in _TEXT_BLOCK_RE.finditer(page_xml):
        text_only = _TAG_SPLIT_RE.sub("", m.group(2)).strip()
        if len(text_only) == 1 and text_only.upper() == old_u:
            cands.append(m)

    if not cands:
        return page_xml, "not_found"

    index = 0
    if len(cands) > 1:
        index = _choose_rev_candidate(page_xml, old_u)
        if index is None or index >= len(cands):
            return page_xml, "ambiguous"

    m = cands[index]
    new_inner = re.sub(
        re.escape(old_letter), lambda _m: new_letter, m.group(2),
        count=1, flags=re.IGNORECASE,
    )
    new_block = m.group(1) + new_inner + m.group(3)
    return page_xml[: m.start()] + new_block + page_xml[m.end():], "updated"


def replace_text_in_vsdx(
    in_path: str | os.PathLike,
    out_path: str | os.PathLike,
    pairs: Sequence[Tuple[str, str]],
    case_sensitive: bool = True,
    whole_word: bool = False,
    revision: Optional[Tuple[str, str]] = None,
    update_drawing_rev: bool = False,
) -> dict:
    """Copy a .vsdx applying text replacements; return a report dict.

    If ``revision`` is (old_letter, new_letter) and ``update_drawing_rev`` is
    true, the single-letter revision box on page 1 is bumped to new_letter.

    Report: {"total", "by_part", "rev_drawing"} where rev_drawing is one of
    'na' (not attempted), 'updated', 'not_found', 'ambiguous'.
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
    rev_status = "na"
    do_rev = bool(revision and update_drawing_rev)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(in_path, "r") as zin:
        first_page = _first_page_part(zin.namelist()) if do_rev else None
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
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
                        if do_rev and item.filename == first_page:
                            new_text, rev_status = bump_revision_in_page(
                                new_text, revision[0], revision[1]
                            )
                        if new_text != text:
                            data = new_text.encode("utf-8")
                # Preserve the original name; recompress with deflate.
                zout.writestr(item.filename, data)

    return {"total": total, "by_part": by_part, "rev_drawing": rev_status}


# ---------------------------------------------------------------------------
# Excel workbooks (.xlsx)
# ---------------------------------------------------------------------------
#
# Excel text is not stored in the worksheet cells directly: string cells point
# at a shared-strings table (xl/sharedStrings.xml), where the text lives inside
# <t> elements. Text can also appear as inline strings in sheets and as shape
# text in drawings (<a:t>). We run-aware replace across all of those.

_XLSX_SI_BLOCK = re.compile(r"<si\b[^>]*>.*?</si>", re.DOTALL)
_XLSX_IS_BLOCK = re.compile(r"<is\b[^>]*>.*?</is>", re.DOTALL)
_XLSX_AP_BLOCK = re.compile(r"<a:p\b[^>]*>.*?</a:p>", re.DOTALL)
_XLSX_T_RUN = re.compile(r"<t\b[^>]*>(.*?)</t>", re.DOTALL)
_XLSX_AT_RUN = re.compile(r"<a:t\b[^>]*>(.*?)</a:t>", re.DOTALL)
_XLSX_CELL = re.compile(r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', re.DOTALL)
_WORKSHEET_RE = re.compile(r"xl/worksheets/sheet\d+\.xml$", re.IGNORECASE)


def _xml_unescape(text: str) -> str:
    return (text.replace("&lt;", "<").replace("&gt;", ">")
            .replace("&quot;", '"').replace("&apos;", "'")
            .replace("&amp;", "&"))


def _replace_in_blocks(xml_text, block_re, run_re, compiled) -> Tuple[str, int]:
    """Run-aware replace within each block's <t>/<a:t> runs."""
    total = 0

    def process_block(bm: re.Match) -> str:
        nonlocal total
        block = bm.group(0)
        runs = list(run_re.finditer(block))
        if not runs:
            return block
        texts = [r.group(1) for r in runs]
        n = _replace_across_runs(texts, compiled)
        if not n:
            return block
        total += n
        out, last = [], 0
        for i, r in enumerate(runs):
            out.append(block[last:r.start(1)])
            out.append(texts[i])
            last = r.end(1)
        out.append(block[last:])
        return "".join(out)

    return block_re.sub(process_block, xml_text), total


def _parse_cell_ref(ref: str) -> Optional[Tuple[int, int]]:
    """'B12' -> (column=2, row=12)."""
    m = re.match(r"([A-Z]+)(\d+)", ref)
    if not m:
        return None
    col = 0
    for ch in m.group(1):
        col = col * 26 + (ord(ch) - 64)
    return col, int(m.group(2))


def _read_shared_strings(zin: zipfile.ZipFile) -> List[str]:
    try:
        xml = zin.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    out = []
    for sim in _XLSX_SI_BLOCK.finditer(xml):
        ts = _XLSX_T_RUN.findall(sim.group(0))
        out.append(_xml_unescape("".join(ts)))
    return out


def _cell_text(attrs: str, content: str, shared: List[str]) -> Optional[str]:
    tm = re.search(r'\bt="([^"]+)"', attrs)
    t = tm.group(1) if tm else "n"
    if t == "s":
        v = re.search(r"<v>(.*?)</v>", content, re.DOTALL)
        if not v:
            return None
        try:
            idx = int(v.group(1))
        except ValueError:
            return None
        return shared[idx] if 0 <= idx < len(shared) else None
    if t == "inlineStr":
        return _xml_unescape("".join(_XLSX_T_RUN.findall(content)))
    if t == "str":
        v = re.search(r"<v>(.*?)</v>", content, re.DOTALL)
        return _xml_unescape(v.group(1)) if v else None
    if t == "n":  # plain number (no string table)
        v = re.search(r"<v>(.*?)</v>", content, re.DOTALL)
        return v.group(1) if v else None
    return None


# -- sheet name -> part, and the protected "Change Log" sheet ---------------

def _sheet_name_parts(zin: zipfile.ZipFile) -> dict:
    """Map each worksheet's normalized name -> its archive part name."""
    try:
        wb = zin.read("xl/workbook.xml").decode("utf-8", "replace")
        rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    except KeyError:
        return {}
    rid_target = {}
    for rel in re.finditer(r"<Relationship\b[^>]*>", rels):
        s = rel.group(0)
        idm = re.search(r'Id="([^"]+)"', s)
        tm = re.search(r'Target="([^"]+)"', s)
        if idm and tm:
            rid_target[idm.group(1)] = tm.group(1)
    out = {}
    for sh in re.finditer(r"<sheet\b[^>]*>", wb):
        s = sh.group(0)
        nm = re.search(r'name="([^"]+)"', s)
        rm = re.search(r'r:id="([^"]+)"', s)
        if not nm or not rm:
            continue
        target = rid_target.get(rm.group(1))
        if not target:
            continue
        part = posixpath.normpath(posixpath.join("xl", target.lstrip("/")))
        out[_norm_header(_xml_unescape(nm.group(1)))] = part
    return out


def _changelog_part(zin: zipfile.ZipFile) -> Optional[str]:
    """The worksheet part for a sheet named 'Change Log' (if any)."""
    return _sheet_name_parts(zin).get("changelog")


def _drawing_for_sheet(zin: zipfile.ZipFile, sheet_part: str) -> Optional[str]:
    """The drawing part used by a worksheet, via its rels (if any)."""
    base = posixpath.dirname(sheet_part)
    rels = f"{base}/_rels/{posixpath.basename(sheet_part)}.rels"
    try:
        xml = zin.read(rels).decode("utf-8", "replace")
    except KeyError:
        return None
    m = re.search(r'Target="([^"]*drawings/drawing\d+\.xml)"', xml)
    if not m:
        return None
    return posixpath.normpath(posixpath.join(base, m.group(1)))


def _replace_cells_in_sheet(sheet_xml, shared, plain_compiled):
    """Find/replace within a worksheet's string cells (per-cell, no aliasing).

    Only string cells (shared or inline) are touched; numbers and formulas are
    left alone. Changed cells become self-contained inline strings.
    """
    count = 0

    def repl_cell(m: re.Match) -> str:
        nonlocal count
        ref, attrs, content = m.group(1), m.group(2), m.group(3)
        tm = re.search(r'\bt="([^"]+)"', attrs)
        t = tm.group(1) if tm else "n"
        if t not in ("s", "inlineStr"):
            return m.group(0)
        text = _cell_text(attrs, content, shared)
        if text is None:
            return m.group(0)
        new_text = text
        for pat, rep in plain_compiled:
            new_text, n = pat.subn(lambda mm, r=rep: r, new_text)
            count += n
        if new_text == text:
            return m.group(0)
        sm = re.search(r'\bs="(\d+)"', attrs)
        style = f' s="{sm.group(1)}"' if sm else ""
        return _cell_content_xml(ref, style, new_text)

    return _XLSX_CELL.sub(repl_cell, sheet_xml), count


def _find_rev_cell(zin, names, shared, old_letter):
    """Find the worksheet cell holding the revision letter (next to a REV
    label). Returns ((part_name, cell_ref), status)."""
    old_u = old_letter.upper()
    best = None  # (score, part, ref)
    no_label_cands = []
    any_cands = False
    protected = _changelog_part(zin)
    for name in names:
        if not _WORKSHEET_RE.search(name.lower()) or name == protected:
            continue
        xml = zin.read(name).decode("utf-8", "replace")
        labels, cands = [], []
        for cm in _XLSX_CELL.finditer(xml):
            ref, attrs, content = cm.group(1), cm.group(2), cm.group(3)
            txt = _cell_text(attrs, content, shared)
            if txt is None:
                continue
            norm = txt.strip()
            if _is_rev_label(norm):
                labels.append(_parse_cell_ref(ref))
            elif len(norm) == 1 and norm.upper() == old_u:
                cands.append((ref, _parse_cell_ref(ref)))
        if not cands:
            continue
        any_cands = True
        if labels:
            for ref, pos in cands:
                if pos is None:
                    continue
                for lp in labels:
                    if lp is None:
                        continue
                    same_row = pos[1] == lp[1]
                    dist = abs(pos[0] - lp[0]) + (
                        0 if same_row else 1000 + abs(pos[1] - lp[1])
                    )
                    if best is None or dist < best[0]:
                        best = (dist, name, ref)
        else:
            no_label_cands.extend((name, ref) for ref, _ in cands)

    if best is not None:
        return (best[1], best[2]), "ok"
    if len(no_label_cands) == 1:
        return no_label_cands[0], "ok"
    if any_cands:
        return None, "ambiguous"
    return None, "not_found"


def _cell_content_xml(ref: str, style: str, text: str) -> str:
    """Build a <c> element: numeric values stay numbers, else inline string."""
    if re.fullmatch(r"-?\d+(\.\d+)?", text.strip()):
        return f'<c r="{ref}"{style}><v>{text.strip()}</v></c>'
    return (f'<c r="{ref}"{style} t="inlineStr"><is><t>'
            f'{xml_escape(text)}</t></is></c>')


def _set_cell_text(xml_text: str, ref: str, text: str) -> Tuple[str, bool]:
    """Set one cell's value (keeps its style; keeps numbers numeric)."""
    cell_re = re.compile(
        r'<c r="' + re.escape(ref) + r'"([^>]*?)(?:/>|>.*?</c>)', re.DOTALL
    )

    def repl(m: re.Match) -> str:
        sm = re.search(r'\bs="(\d+)"', m.group(1))
        style = f' s="{sm.group(1)}"' if sm else ""
        return _cell_content_xml(ref, style, text)

    new, n = cell_re.subn(repl, xml_text, count=1)
    return new, n > 0


def replace_text_in_xlsx(
    in_path, out_path, pairs,
    case_sensitive=True, whole_word=False,
    revision=None, update_drawing_rev=False,
    cell_edits=None,
) -> dict:
    """Find/replace (and optional revision bump + cell edits) for a workbook.

    ``cell_edits`` is an optional {worksheet_part_name: {cell_ref: new_text}}
    mapping for setting specific cells (used by the BOM row editor).
    """
    in_path = Path(in_path)
    out_path = Path(out_path)
    if not zipfile.is_zipfile(in_path):
        raise ValueError(f"'{in_path.name}' is not a valid .xlsx file.")

    # Cell text is decoded, so match with un-escaped patterns; drawing run text
    # is raw XML, so match with escaped patterns.
    cell_compiled = _compile_pairs(pairs, case_sensitive, whole_word, escape=False)
    draw_compiled = _compile_pairs(pairs, case_sensitive, whole_word, escape=True)
    cell_edits = cell_edits or {}
    by_part: dict[str, int] = {}
    total = 0
    cells_changed = 0
    rev_status = "na"
    target = None
    do_rev = bool(revision and update_drawing_rev)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(in_path, "r") as zin:
        names = zin.namelist()
        shared = _read_shared_strings(zin)
        # The "Change Log" sheet (and its drawing) are never edited by
        # find/replace -- they must stay intact for traceability.
        protected = _changelog_part(zin)
        protected_draw = (_drawing_for_sheet(zin, protected)
                          if protected else None)
        if do_rev:
            target, rev_status = _find_rev_cell(zin, names, shared, revision[0])
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                name = item.filename
                ln = name.lower()
                if ln.startswith("xl/") and ln.endswith(".xml"):
                    try:
                        text = data.decode("utf-8")
                    except UnicodeDecodeError:
                        zout.writestr(name, data)
                        continue
                    new_text = text
                    is_protected = (name == protected
                                    or name == protected_draw)
                    if cell_compiled and not is_protected:
                        if _WORKSHEET_RE.search(ln):
                            new_text, count = _replace_cells_in_sheet(
                                new_text, shared, cell_compiled
                            )
                        elif "/drawings/" in ln:
                            new_text, count = _replace_in_blocks(
                                new_text, _XLSX_AP_BLOCK, _XLSX_AT_RUN,
                                draw_compiled,
                            )
                        else:
                            count = 0
                        if count:
                            by_part[name] = count
                            total += count
                    # Explicit cell edits (BOM edits, Change Log append) apply
                    # even to the protected sheet (the append is intentional).
                    for ref, value in cell_edits.get(name, {}).items():
                        new_text = _set_or_insert_cell(new_text, ref, value)
                        cells_changed += 1
                    if target and name == target[0]:
                        new_text, changed = _set_cell_text(
                            new_text, target[1], revision[1]
                        )
                        if changed:
                            rev_status = "updated"
                    if new_text != text:
                        data = new_text.encode("utf-8")
                zout.writestr(name, data)

    return {"total": total, "by_part": by_part, "rev_drawing": rev_status,
            "cells_changed": cells_changed}


# ---------------------------------------------------------------------------
# Excel BOM rows: find a part number's row and read/edit its other columns
# ---------------------------------------------------------------------------
#
# In a bill-of-materials sheet, the part number lives in a "P/N" / "Part Number"
# column and the rest of the row holds Manufacturer, Unit Cost, etc. We locate
# the header row, map columns to those fields, and find rows whose P/N matches.

# Canonical field -> accepted header spellings (normalized: lowercase, alnum).
_BOM_FIELDS = {
    "Part Number": {"pn", "partnumber", "partno", "partnum", "part"},
    "Manufacturer": {"manufacturer", "mfg", "mfr", "manuf", "make", "vendor"},
    "Unit Cost": {"unitcost", "cost", "price", "unitprice", "ucost"},
    "Description": {"description", "desc", "descr"},
    "Qty": {"qty", "quantity", "qnty"},
    "Notes": {"notes", "note", "comments", "comment", "remarks"},
}
# Order for display (Part Number first).
BOM_FIELD_ORDER = ["Part Number", "Manufacturer", "Unit Cost",
                   "Description", "Qty", "Notes"]


def _norm_header(text: str) -> str:
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _canonical_field(text: str) -> Optional[str]:
    n = _norm_header(text)
    if not n:
        return None
    for field, variants in _BOM_FIELDS.items():
        if n in variants:
            return field
    return None


def _col_num(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _col_letters(col: int) -> str:
    s = ""
    while col > 0:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return s


def _set_or_insert_cell(xml_text: str, ref: str, text: str) -> str:
    """Set a cell's text; if the cell/row doesn't exist, insert it in order."""
    new, ok = _set_cell_text(xml_text, ref, text)
    if ok:
        return new
    pos = _parse_cell_ref(ref)
    if pos is None:
        return xml_text
    col, row = pos
    cell_xml = _cell_content_xml(ref, "", text)

    row_re = re.compile(r'(<row r="%d"[^>]*>)(.*?)(</row>)' % row, re.DOTALL)
    rm = row_re.search(xml_text)
    if rm:
        inner = rm.group(2)
        insert_at = len(inner)
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"', inner):
            if _col_num(cm.group(1)) > col:
                insert_at = cm.start()
                break
        new_inner = inner[:insert_at] + cell_xml + inner[insert_at:]
        return (xml_text[:rm.start()] + rm.group(1) + new_inner
                + rm.group(3) + xml_text[rm.end():])

    sd = re.search(r'(<sheetData[^>]*>)(.*?)(</sheetData>)', xml_text, re.DOTALL)
    if not sd:
        return xml_text
    inner = sd.group(2)
    row_xml = f'<row r="{row}">{cell_xml}</row>'
    insert_at = len(inner)
    for rm2 in re.finditer(r'<row r="(\d+)"', inner):
        if int(rm2.group(1)) > row:
            insert_at = rm2.start()
            break
    new_inner = inner[:insert_at] + row_xml + inner[insert_at:]
    return (xml_text[:sd.start()] + sd.group(1) + new_inner
            + sd.group(3) + xml_text[sd.end():])


def _read_sheet_cells(sheet_xml: str, shared: List[str]):
    """Return {(col,row): (ref, text)} for all non-empty cells in a sheet."""
    cells = {}
    for cm in _XLSX_CELL.finditer(sheet_xml):
        ref, attrs, content = cm.group(1), cm.group(2), cm.group(3)
        pos = _parse_cell_ref(ref)
        if pos is None:
            continue
        cells[pos] = (ref, _cell_text(attrs, content, shared))
    return cells


def _find_bom_header(cells):
    """Find the header row; return (row, {col: canonical_field})."""
    rows: dict = {}
    for (col, row), (ref, txt) in cells.items():
        rows.setdefault(row, {})[col] = txt
    for row in sorted(rows):
        colmap, has_pn = {}, False
        for col, txt in rows[row].items():
            if not txt:
                continue
            cf = _canonical_field(txt)
            if cf:
                colmap[col] = cf
                if cf == "Part Number":
                    has_pn = True
        if has_pn:
            return row, colmap
    return None, {}


def excel_scan_rows(path, part_numbers, case_sensitive=False):
    """Find rows whose P/N matches; return per-row field data + cell refs.

    Each match: {file, sheet, part, row, matched, fields:{field:(ref,value)}}.
    """
    matches = []
    norm_targets = [(p, p if case_sensitive else p.lower()) for p in part_numbers
                    if p]
    with zipfile.ZipFile(path) as z:
        shared = _read_shared_strings(z)
        protected = _changelog_part(z)
        for name in z.namelist():
            if not _WORKSHEET_RE.search(name.lower()) or name == protected:
                continue
            sheet_xml = z.read(name).decode("utf-8", "replace")
            cells = _read_sheet_cells(sheet_xml, shared)
            hrow, colmap = _find_bom_header(cells)
            if hrow is None:
                continue
            pn_cols = [c for c, f in colmap.items() if f == "Part Number"]
            if not pn_cols:
                continue
            pn_col = pn_cols[0]
            rows: dict = {}
            for (col, row), (ref, txt) in cells.items():
                rows.setdefault(row, {})[col] = (ref, txt)
            for row in sorted(rows):
                if row <= hrow:
                    continue
                pn_cell = rows[row].get(pn_col)
                if not pn_cell or not pn_cell[1]:
                    continue
                pnval = pn_cell[1].strip()
                cmp_val = pnval if case_sensitive else pnval.lower()
                hit = next((orig for orig, t in norm_targets if t == cmp_val),
                           None)
                if hit is None:
                    continue
                fields = {}
                for col, cf in colmap.items():
                    cell = rows[row].get(col)
                    ref = cell[0] if cell else f"{_col_letters(col)}{row}"
                    val = cell[1] if cell and cell[1] is not None else ""
                    fields[cf] = (ref, val)
                matches.append({"file": str(path), "sheet": name,
                                "part": pnval, "row": row, "matched": hit,
                                "fields": fields})
    return matches


def build_excel_cell_edits(path, bom_edits, case_sensitive=False) -> dict:
    """Map staged BOM field edits to actual cell refs for one file.

    bom_edits: {part_number: {canonical_field: new_value}}.
    Returns {worksheet_part_name: {cell_ref: new_value}} for every matching row.
    """
    edits: dict = {}
    if not bom_edits:
        return edits
    matches = excel_scan_rows(path, list(bom_edits.keys()), case_sensitive)
    for m in matches:
        for field, new_value in bom_edits.get(m["matched"], {}).items():
            cell = m["fields"].get(field)
            if cell is None:
                continue  # that column doesn't exist in this sheet
            edits.setdefault(m["sheet"], {})[cell[0]] = new_value
    return edits


# ---------------------------------------------------------------------------
# Excel Change Log: append a new row to the (protected) "Change Log" sheet
# ---------------------------------------------------------------------------

_CHANGELOG_FIELDS = {
    "Item": {"item", "itemno", "itemnumber", "no", "line"},
    "ECN #": {"ecn", "eco", "ecnno", "econo", "ecnnumber", "econumber",
              "ecn#", "eco#"},
    "ERB Approval Date": {"erbapprovaldate", "approvaldate", "erbdate",
                          "dateapproved", "approval", "erbapproval"},
    "Change Description": {"changedescription", "description", "changedesc",
                           "desc", "change"},
    "Change Author": {"changeauthor", "author", "changedby", "by",
                      "enteredby"},
}
CHANGELOG_FIELD_ORDER = ["Item", "ECN #", "ERB Approval Date",
                         "Change Description", "Change Author"]


def _canonical_changelog_field(text: str) -> Optional[str]:
    n = _norm_header(text)
    if not n:
        return None
    for field, variants in _CHANGELOG_FIELDS.items():
        if n in variants:
            return field
    return None


def changelog_info(path):
    """Inspect the Change Log sheet.

    Returns (part, header_row, {field: column}, next_row) or None if the file
    has no Change Log sheet / no recognizable header. ``next_row`` is the first
    empty row after the last entry, judged by the ECN # column.
    """
    with zipfile.ZipFile(path) as z:
        part = _changelog_part(z)
        if not part:
            return None
        try:
            xml = z.read(part).decode("utf-8", "replace")
        except KeyError:
            return None
        shared = _read_shared_strings(z)

    cells = _read_sheet_cells(xml, shared)
    rows: dict = {}
    for (col, row), (ref, txt) in cells.items():
        rows.setdefault(row, {})[col] = txt

    header_row, colmap = None, {}
    for row in sorted(rows):
        cmap, has_ecn = {}, False
        for col, txt in rows[row].items():
            if not txt:
                continue
            cf = _canonical_changelog_field(txt)
            if cf:
                cmap[col] = cf
                if cf == "ECN #":
                    has_ecn = True
        if has_ecn:
            header_row, colmap = row, cmap
            break
    if header_row is None:
        return None

    field_to_col = {f: c for c, f in colmap.items()}
    ecn_col = field_to_col.get("ECN #")
    # Next available row = after the last row that has an ECN # value.
    last = header_row
    for row in sorted(rows):
        if row <= header_row:
            continue
        cell = rows[row].get(ecn_col)
        if cell:  # non-empty ECN # marks a real entry
            last = row
    return part, header_row, field_to_col, last + 1


def build_changelog_cell_edits(path, entry) -> dict:
    """Map a new Change Log entry to cell edits on the Change Log sheet.

    entry: {canonical_field: value}. Returns {part: {ref: value}} or {}.
    """
    if not entry or not any(v.strip() for v in entry.values()):
        return {}
    info = changelog_info(path)
    if info is None:
        return {}
    part, header_row, field_to_col, next_row = info
    edits = {}
    for field, value in entry.items():
        if not value.strip():
            continue
        col = field_to_col.get(field)
        if col is None:
            continue
        ref = f"{_col_letters(col)}{next_row}"
        edits[ref] = value
    return {part: edits} if edits else {}


# ---------------------------------------------------------------------------
# Format detection + dispatch
# ---------------------------------------------------------------------------

def detect_format(path: str | os.PathLike) -> Optional[str]:
    """Return 'vsdx', 'xlsx', or None, by extension then by archive contents."""
    ext = Path(path).suffix.lower()
    if ext == ".vsdx":
        return "vsdx"
    if ext == ".xlsx":
        return "xlsx"
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if any(n.startswith("visio/") for n in names):
                return "vsdx"
            if any(n.startswith("xl/") for n in names):
                return "xlsx"
    except (zipfile.BadZipFile, OSError):
        pass
    return None


def replace_text_in_file(
    in_path, out_path, pairs,
    case_sensitive=True, whole_word=False,
    revision=None, update_drawing_rev=False,
    cell_edits=None,
) -> dict:
    """Dispatch to the Visio or Excel engine based on the file type."""
    fmt = detect_format(in_path)
    if fmt == "vsdx":
        return replace_text_in_vsdx(
            in_path, out_path, pairs, case_sensitive, whole_word,
            revision=revision, update_drawing_rev=update_drawing_rev,
        )
    if fmt == "xlsx":
        return replace_text_in_xlsx(
            in_path, out_path, pairs, case_sensitive, whole_word,
            revision=revision, update_drawing_rev=update_drawing_rev,
            cell_edits=cell_edits,
        )
    raise ValueError(
        f"Unsupported file type '{Path(in_path).suffix}'. "
        "This tool handles Visio .vsdx and Excel .xlsx files."
    )


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
            "--invisible",
            "--nodefault",
            "--norestore",
            "--nolockcheck",
            "--nofirststartwizard",
            "--nologo",
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
        "--version", action="version", version=f"%(prog)s {__version__}",
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
    parser.add_argument(
        "--bump-revision", action="store_true",
        help="Name each copy as the next revision (e.g. REVA -> REVB) read "
        "from the file name, instead of the *_edited suffix",
    )
    parser.add_argument(
        "--no-rev-text", action="store_true",
        help="With --bump-revision, do NOT change the REV letter box inside "
        "the drawing (only rename the file)",
    )
    args = parser.parse_args(argv)

    if len(args.find) != len(args.replace):
        parser.error("each --find must be paired with a --replace")
    if not args.find and not args.bump_revision:
        parser.error(
            "provide at least one --find/--replace pair (or --bump-revision)"
        )
    if args.output and len(args.inputs) > 1:
        parser.error("--output cannot be used with multiple input files")
    if args.output and args.bump_revision:
        parser.error("--output cannot be used with --bump-revision")

    pairs = list(zip(args.find, args.replace))
    update_rev_text = not args.no_rev_text
    had_error = False

    for raw in args.inputs:
        in_path = Path(raw)
        if not in_path.exists():
            print(f"Error: input file not found: {in_path}", file=sys.stderr)
            had_error = True
            continue

        revision = None
        if args.bump_revision:
            out_path, old, new, status = revision_output_path(in_path)
            if status == "at_z":
                print(
                    f"{in_path.name}: already at REVZ - skipped "
                    "(no next revision)."
                )
                continue
            if status == "no_rev":
                print(
                    f"{in_path.name}: no REVx in the name; saving as "
                    f"{Path(out_path).name}"
                )
            else:
                revision = (old, new)
        elif args.output:
            out_path = Path(args.output)
        else:
            out_path = in_path.with_name(
                in_path.stem + "_edited" + in_path.suffix
            )

        try:
            report = replace_text_in_file(
                in_path, out_path, pairs,
                case_sensitive=args.case_sensitive, whole_word=args.whole_word,
                revision=revision, update_drawing_rev=update_rev_text,
            )
            print(
                f"{in_path.name}: replaced {report['total']} occurrence(s) "
                f"-> {Path(out_path).name}"
            )
            if revision:
                rd = report["rev_drawing"]
                note = {
                    "updated": f"REV box bumped {revision[0]} -> {revision[1]}",
                    "not_found": "REV box not found in drawing (file renamed "
                    "only)",
                    "ambiguous": "couldn't identify the REV box (file renamed "
                    "only)",
                    "na": "in-drawing REV left unchanged",
                }.get(rd, rd)
                print(f"    {note}")
            if args.pdf:
                pdf = convert_to_pdf(out_path, Path(out_path).parent)
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

    class App:
        def __init__(self, root: "tk.Tk"):
            self.root = root
            root.title(f"Visio / Excel Text Replacer   [v{__version__}]")
            root.geometry("820x760")
            root.minsize(680, 640)

            self.files: List[str] = []
            # Each rule: {frame, find, repl, menubtn, menu, all_var, file_vars}.
            # all_var True => applies to every file; otherwise file_vars holds a
            # BooleanVar per file path for multi-select targeting.
            self.rule_rows: List[dict] = []
            # Staged Excel row edits: {part_number: {canonical_field: value}}.
            self.bom_edits: dict = {}
            # Staged Change Log row to append: {canonical_field: value}.
            self.changelog_entry: dict = {}

            pad = {"padx": 10, "pady": 6}

            # --- 1. Files --------------------------------------------------
            top = ttk.LabelFrame(
                root, text="1.  Files (.vsdx Visio / .xlsx Excel)"
            )
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
            rule_btns = ttk.Frame(mid)
            rule_btns.pack(fill="x", padx=8, pady=(0, 8))
            ttk.Button(
                rule_btns, text="+ Add another rule", command=self.add_pair
            ).pack(side="left")
            ttk.Button(
                rule_btns, text="Excel: find & edit rows...",
                command=self.open_bom_editor,
            ).pack(side="left", padx=8)
            ttk.Button(
                rule_btns, text="Excel: add Change Log entry...",
                command=self.open_changelog_editor,
            ).pack(side="left")
            self.bom_status = ttk.Label(rule_btns, text="")
            self.bom_status.pack(side="left", padx=8)

            # --- 3. Options ------------------------------------------------
            opts = ttk.LabelFrame(root, text="3.  Options")
            opts.pack(fill="x", **pad)
            self.case_var = tk.BooleanVar(value=False)
            self.word_var = tk.BooleanVar(value=False)
            # PDF export (LibreOffice) is off by default: export from Visio for
            # correct sheet numbers and full fidelity.
            self.pdf_var = tk.BooleanVar(value=False)
            self.rev_var = tk.BooleanVar(value=True)
            self.revtext_var = tk.BooleanVar(value=True)

            opt_row1 = ttk.Frame(opts)
            opt_row1.pack(fill="x")
            ttk.Checkbutton(
                opt_row1, text="Case sensitive", variable=self.case_var
            ).pack(side="left", padx=8, pady=6)
            ttk.Checkbutton(
                opt_row1, text="Whole word only", variable=self.word_var
            ).pack(side="left", padx=8, pady=6)
            ttk.Checkbutton(
                opt_row1, text="Also export PDF (LibreOffice)",
                variable=self.pdf_var,
            ).pack(side="left", padx=8, pady=6)

            opt_row2 = ttk.Frame(opts)
            opt_row2.pack(fill="x")
            ttk.Checkbutton(
                opt_row2,
                text="Save copy as next revision (REVx → next)",
                variable=self.rev_var, command=self._sync_rev_options,
            ).pack(side="left", padx=8, pady=6)
            self.revtext_chk = ttk.Checkbutton(
                opt_row2, text="...and update the REV box in the drawing",
                variable=self.revtext_var,
            )
            self.revtext_chk.pack(side="left", padx=8, pady=6)

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
            self.log_box = scrolledtext.ScrolledText(
                logf, height=8, state="disabled", wrap="word"
            )
            self.log_box.pack(fill="both", expand=True, padx=6, pady=6)

            self.on_mode_change()
            self._sync_rev_options()
            self._log(f"Visio Text Replacer  v{__version__}")
            if find_libreoffice() is None:
                self._log(
                    "Note: LibreOffice was not found, so PDF export is "
                    "unavailable. Install it from libreoffice.org to enable "
                    "PDF output. Text replacement still works."
                )

        def _sync_rev_options(self):
            self.revtext_chk.configure(
                state="normal" if self.rev_var.get() else "disabled"
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
            ft = [
                ("Visio / Excel", "*.vsdx *.xlsx"),
                ("Visio drawing", "*.vsdx"),
                ("Excel workbook", "*.xlsx"),
                ("All files", "*.*"),
            ]
            if self.mode_var.get() == "single":
                path = filedialog.askopenfilename(
                    title="Choose a Visio or Excel file", filetypes=ft
                )
                self.files = [path] if path else self.files[:0]
            else:
                paths = filedialog.askopenfilenames(
                    title="Choose Visio/Excel files", filetypes=ft
                )
                for p in paths:
                    if p not in self.files:
                        self.files.append(p)
            self.refresh_files_box()

        def add_folder(self):
            folder = filedialog.askdirectory(
                title="Choose a folder of .vsdx / .xlsx files"
            )
            if not folder:
                return
            found = sorted(
                str(p) for p in Path(folder).iterdir()
                if p.suffix.lower() in (".vsdx", ".xlsx")
            )
            for p in found:
                if p not in self.files:
                    self.files.append(p)
            if not found:
                messagebox.showinfo(
                    "No files",
                    "No .vsdx or .xlsx files were found in that folder.",
                )
            self.refresh_files_box()

        def remove_selected_files(self):
            for i in reversed(self.files_box.curselection()):
                del self.files[i]
            self.refresh_files_box()

        def clear_files(self):
            self.files = []
            self.refresh_files_box()

        def refresh_files_box(self):
            self.files_box.delete(0, "end")
            for f in self.files:
                self.files_box.insert("end", Path(f).name)
            self._refresh_rule_menus()

        def _refresh_rule_menus(self):
            """Rebuild every rule's multi-select file dropdown."""
            for rule in self.rule_rows:
                menu = rule["menu"]
                menu.delete(0, "end")
                fv = rule["file_vars"]
                for p in list(fv):  # drop files no longer loaded
                    if p not in self.files:
                        del fv[p]
                for f in self.files:
                    fv.setdefault(f, tk.BooleanVar(value=False))
                menu.add_checkbutton(
                    label="All files", variable=rule["all_var"],
                    command=lambda r=rule: self._on_all_toggle(r),
                )
                menu.add_separator()
                for f in self.files:
                    menu.add_checkbutton(
                        label=Path(f).name, variable=fv[f],
                        command=lambda r=rule, p=f: self._on_file_toggle(r, p),
                    )
                self._update_menu_label(rule)

        def _on_all_toggle(self, rule):
            if rule["all_var"].get():
                for v in rule["file_vars"].values():
                    v.set(False)
            self._update_menu_label(rule)

        def _on_file_toggle(self, rule, path):
            if rule["file_vars"][path].get():
                rule["all_var"].set(False)
            self._update_menu_label(rule)

        def _update_menu_label(self, rule):
            if rule["all_var"].get():
                rule["menubtn"]["text"] = "All files"
                return
            sel = [f for f in self.files
                   if rule["file_vars"].get(f) and rule["file_vars"][f].get()]
            if not sel:
                rule["menubtn"]["text"] = "(no files)"
            elif len(sel) == 1:
                rule["menubtn"]["text"] = Path(sel[0]).name[:18]
            else:
                rule["menubtn"]["text"] = f"{len(sel)} files"

        # -- rule rows ------------------------------------------------------
        def _header_row(self):
            ttk.Label(
                self.pairs_frame,
                text="Type the text to find and its replacement, then pick "
                "which file(s) the rule applies to (the 'in' menu lets you "
                "tick several).",
            ).pack(anchor="w", padx=4, pady=(0, 4))

        def add_pair(self):
            row = ttk.Frame(self.pairs_frame)
            row.pack(fill="x", pady=3)
            ttk.Label(row, text="Find:").pack(side="left", padx=(4, 2))
            find_e = ttk.Entry(row, width=16)
            find_e.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ttk.Label(row, text="Replace with:").pack(side="left", padx=(0, 2))
            repl_e = ttk.Entry(row, width=16)
            repl_e.pack(side="left", fill="x", expand=True, padx=(0, 8))
            ttk.Label(row, text="in:").pack(side="left", padx=(0, 2))
            menubtn = ttk.Menubutton(row, text="All files", width=16)
            menu = tk.Menu(menubtn, tearoff=0)
            menubtn["menu"] = menu
            menubtn.pack(side="left", padx=(0, 4))
            rule = {
                "frame": row, "find": find_e, "repl": repl_e,
                "menubtn": menubtn, "menu": menu,
                "all_var": tk.BooleanVar(value=True), "file_vars": {},
            }
            ttk.Button(
                row, text="X", width=3,
                command=lambda r=rule: self.remove_pair(r),
            ).pack(side="left", padx=4)
            self.rule_rows.append(rule)
            self._refresh_rule_menus()

        def remove_pair(self, rule):
            if len(self.rule_rows) <= 1:
                return  # keep at least one row
            rule["frame"].destroy()
            self.rule_rows = [r for r in self.rule_rows if r is not rule]

        # -- helpers --------------------------------------------------------
        def _log(self, msg: str):
            self.log_box.configure(state="normal")
            self.log_box.insert("end", msg + "\n")
            self.log_box.see("end")
            self.log_box.configure(state="disabled")

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
            """Find/replace pairs whose dropdown selection includes this file."""
            pairs = []
            for rule in self.rule_rows:
                find = rule["find"].get()
                if not find:
                    continue
                if rule["all_var"].get():
                    pairs.append((find, rule["repl"].get()))
                else:
                    v = rule["file_vars"].get(path)
                    if v is not None and v.get():
                        pairs.append((find, rule["repl"].get()))
            return pairs

        # -- Excel BOM row editor ------------------------------------------
        def _find_values(self) -> List[str]:
            seen = []
            for rule in self.rule_rows:
                f = rule["find"].get().strip()
                if f and f not in seen:
                    seen.append(f)
            return seen

        def open_bom_editor(self):
            parts = self._find_values()
            if not parts:
                messagebox.showinfo(
                    "No Find values",
                    "Type the part number(s) into the Find box(es) first.",
                )
                return
            excel_files = [f for f in self.files
                           if detect_format(f) == "xlsx"]
            if not excel_files:
                messagebox.showinfo(
                    "No Excel files", "Add at least one .xlsx file first."
                )
                return

            found: dict = {}            # part -> {field: value}
            files_with: dict = {}       # part -> set(files)
            rows_with: dict = {}        # part -> count
            cs = self.case_var.get()
            for f in excel_files:
                try:
                    matches = excel_scan_rows(f, parts, cs)
                except Exception:  # noqa: BLE001
                    matches = []
                for mt in matches:
                    p = mt["matched"]
                    files_with.setdefault(p, set()).add(f)
                    rows_with[p] = rows_with.get(p, 0) + 1
                    if p not in found:
                        found[p] = {
                            fld: val for fld, (ref, val) in mt["fields"].items()
                        }
            if not found:
                messagebox.showinfo(
                    "No rows found",
                    "None of the Find value(s) were located in a P/N / "
                    "Part Number column of the Excel files.",
                )
                return

            self._build_bom_window(found, files_with, rows_with)

        def _build_bom_window(self, found, files_with, rows_with):
            win = tk.Toplevel(self.root)
            win.title("Find & edit Excel rows")
            win.geometry("680x520")
            win.transient(self.root)
            win.grab_set()

            ttk.Label(
                win, wraplength=640, justify="left",
                text="Edit any field below to a new value. Changes are applied "
                "to that part number's row in EVERY Excel file that contains "
                "it. Fields you leave unchanged are not touched.",
            ).pack(fill="x", padx=10, pady=8)

            canvas = tk.Canvas(win, highlightthickness=0)
            sb = ttk.Scrollbar(win, orient="vertical", command=canvas.yview)
            inner = ttk.Frame(canvas)
            inner.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
            )
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=sb.set)
            canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
            sb.pack(side="left", fill="y", padx=(0, 10))

            entries: dict = {}
            for part in found:
                lf = ttk.LabelFrame(
                    inner,
                    text=f"P/N: {part}    (in {len(files_with[part])} file(s), "
                    f"{rows_with[part]} row(s))",
                )
                lf.pack(fill="x", padx=6, pady=6)
                entries[part] = {}
                for field in BOM_FIELD_ORDER:
                    if field == "Part Number":
                        continue
                    val = found[part].get(field, "")
                    rowf = ttk.Frame(lf)
                    rowf.pack(fill="x", padx=6, pady=2)
                    ttk.Label(rowf, text=field + ":", width=14).pack(side="left")
                    e = ttk.Entry(rowf)
                    e.insert(0, val)
                    e.pack(side="left", fill="x", expand=True)
                    entries[part][field] = (e, val)

            def apply():
                edits = {}
                for part, flds in entries.items():
                    changed = {f: e.get() for f, (e, orig) in flds.items()
                               if e.get() != orig}
                    if changed:
                        edits[part] = changed
                self.bom_edits = edits
                n = sum(len(v) for v in edits.values())
                self.bom_status.configure(
                    text=(f"{n} row edit(s) staged" if n else "")
                )
                self.log(
                    f"Staged {n} Excel field edit(s) across "
                    f"{len(edits)} part number(s)."
                    if n else "No Excel field edits staged."
                )
                win.destroy()

            btns = ttk.Frame(win)
            btns.pack(fill="x", padx=10, pady=10)
            ttk.Button(btns, text="Save edits", command=apply).pack(
                side="right"
            )
            ttk.Button(btns, text="Cancel", command=win.destroy).pack(
                side="right", padx=6
            )

        # -- Change Log entry ----------------------------------------------
        def open_changelog_editor(self):
            excel_files = [f for f in self.files
                           if detect_format(f) == "xlsx"]
            if not excel_files:
                messagebox.showinfo(
                    "No Excel files", "Add at least one .xlsx file first."
                )
                return
            if not any(changelog_info(f) for f in excel_files):
                messagebox.showinfo(
                    "No Change Log",
                    "None of the Excel files have a sheet named 'Change Log' "
                    "with recognizable column headers (Item, ECN #, ...).",
                )
                return

            win = tk.Toplevel(self.root)
            win.title("Add Change Log entry")
            win.transient(self.root)
            win.grab_set()
            ttk.Label(
                win, wraplength=560, justify="left",
                text="This row is appended to the 'Change Log' sheet of every "
                "Excel file (at the next free row, found via the ECN # "
                "column). Leave a field blank to skip it. Leave Item blank to "
                "keep any pre-filled item number.",
            ).pack(fill="x", padx=10, pady=8)

            form = ttk.Frame(win)
            form.pack(fill="x", padx=10, pady=4)
            entries = {}
            for field in CHANGELOG_FIELD_ORDER:
                rowf = ttk.Frame(form)
                rowf.pack(fill="x", pady=3)
                ttk.Label(rowf, text=field + ":", width=20).pack(side="left")
                e = ttk.Entry(rowf, width=44)
                e.insert(0, self.changelog_entry.get(field, ""))
                e.pack(side="left", fill="x", expand=True)
                entries[field] = e

            def apply():
                entry = {f: e.get() for f, e in entries.items()}
                if not any(v.strip() for v in entry.values()):
                    self.changelog_entry = {}
                    self.bom_status.configure(text="")
                else:
                    self.changelog_entry = entry
                    self.log(
                        "Staged a Change Log entry (ECN "
                        f"{entry.get('ECN #', '').strip() or '—'})."
                    )
                win.destroy()

            btns = ttk.Frame(win)
            btns.pack(fill="x", padx=10, pady=10)
            ttk.Button(btns, text="Save entry", command=apply).pack(
                side="right"
            )
            ttk.Button(btns, text="Cancel", command=win.destroy).pack(
                side="right", padx=6
            )

        # -- run ------------------------------------------------------------
        def run(self):
            if not self.files:
                messagebox.showwarning(
                    "No files", "Please add at least one file."
                )
                return
            missing = [f for f in self.files if not Path(f).exists()]
            if missing:
                messagebox.showerror(
                    "Not found",
                    "These files no longer exist:\n" + "\n".join(missing),
                )
                return
            bump_rev = self.rev_var.get()
            if (not bump_rev
                    and not any(r["find"].get() for r in self.rule_rows)
                    and not self.bom_edits and not self.changelog_entry):
                messagebox.showwarning(
                    "Nothing to do",
                    "Enter a 'Find' value, tick 'Save copy as next revision', "
                    "or stage an Excel row / Change Log edit.",
                )
                return

            self._set_busy(True)
            threading.Thread(
                target=self._worker,
                args=(
                    list(self.files),
                    {f: self.pairs_for_file(f) for f in self.files},
                    self.case_var.get(), self.word_var.get(),
                    self.pdf_var.get(), bump_rev, self.revtext_var.get(),
                    dict(self.bom_edits), dict(self.changelog_entry),
                ),
                daemon=True,
            ).start()

        def _worker(self, files, pairs_by_file, case_sensitive, whole_word,
                    make_pdf, bump_rev, update_rev_text, bom_edits,
                    changelog_entry):
            total_repl = 0
            done = 0
            errors = 0
            last_output = None
            try:
                for f in files:
                    src = Path(f)
                    # Key by the original string used to build the dict (path
                    # separators differ between tkinter and Path on Windows).
                    pairs = pairs_by_file.get(f, [])

                    # Excel cell edits for this file: BOM row edits plus a
                    # Change Log append (by cell ref).
                    cell_edits = None
                    if (bom_edits or changelog_entry) and \
                            detect_format(f) == "xlsx":
                        merged: dict = {}
                        try:
                            if bom_edits:
                                for part, d in build_excel_cell_edits(
                                    f, bom_edits, case_sensitive
                                ).items():
                                    merged.setdefault(part, {}).update(d)
                            if changelog_entry:
                                for part, d in build_changelog_cell_edits(
                                    f, changelog_entry
                                ).items():
                                    merged.setdefault(part, {}).update(d)
                        except Exception:  # noqa: BLE001
                            merged = {}
                        cell_edits = merged or None
                    has_edits = bool(cell_edits)

                    # Decide the copy's name and whether to bump the revision.
                    revision = None
                    if bump_rev:
                        out_vsdx, old, new, status = revision_output_path(src)
                        if status == "at_z":
                            self.log(
                                f"- {src.name}: already at REVZ - skipped "
                                "(no next revision)."
                            )
                            continue
                        if status == "no_rev":
                            self.log(
                                f"  {src.name}: no REVx in name; saving as "
                                f"{out_vsdx.name}"
                            )
                        else:
                            revision = (old, new)
                        out_vsdx = Path(out_vsdx)
                    else:
                        if not pairs and not has_edits:
                            self.log(
                                f"- {src.name}: skipped (no rule targets this "
                                "file)"
                            )
                            continue
                        out_vsdx = src.with_name(
                            src.stem + "_edited" + src.suffix
                        )

                    try:
                        report = replace_text_in_file(
                            src, out_vsdx, pairs,
                            case_sensitive=case_sensitive,
                            whole_word=whole_word,
                            revision=revision,
                            update_drawing_rev=update_rev_text,
                            cell_edits=cell_edits,
                        )
                        total_repl += report["total"]
                        msg = (
                            f"+ {src.name}: {report['total']} replacement(s) "
                            f"-> {out_vsdx.name}"
                        )
                        if report["total"] == 0 and pairs and not has_edits:
                            msg += "  (no matches found)"
                        self.log(msg)
                        if report.get("cells_changed"):
                            self.log(
                                f"    {report['cells_changed']} row cell(s) "
                                "updated"
                            )
                        if revision:
                            rd = report["rev_drawing"]
                            note = {
                                "updated":
                                    f"    REV box {revision[0]} -> "
                                    f"{revision[1]}",
                                "not_found":
                                    "    (REV box not found in drawing; file "
                                    "renamed only)",
                                "ambiguous":
                                    "    (couldn't identify the REV box; file "
                                    "renamed only)",
                                "na": None,
                            }.get(rd)
                            if note:
                                self.log(note)
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
            except Exception as exc:  # noqa: BLE001 - never kill the thread
                self.log(f"UNEXPECTED ERROR: {exc}")
                self.root.after(
                    0, lambda e=exc: messagebox.showerror("Error", str(e))
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
    root._app_ref = App(root)  # keep a reference (also handy for testing)
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
