#!/usr/bin/env python3
"""
Drawing & BOM Studio
====================

A desktop application for editing Visio cable/system drawings (.vsdx) and Excel
bills of material (.xlsx). It streamlines the repetitive edits in a drawing
release in one place:

  1. Pick a single Visio drawing (.vsdx) or Excel workbook (.xlsx) -- or a
     batch mixing both. The file type is detected automatically per file.
  2. Find & replace text with one or more rules, aimed at all files or only
     specific ones.
  3. Add, remove, or edit parts in the Visio parts tables and Excel BOMs, with
     item/line numbers renumbered automatically.
  4. Bump revisions, append revision-table rows, sign off approvals, and add
     Change Log entries.
  5. Save the edited copy of each file (named as the next revision), and
     optionally export to PDF.

The edits work directly on the file (both .vsdx and .xlsx are ZIP archives of
XML). For Visio, the text in <Text> shape blocks plus the embedded Excel parts
and revision tables (and their cached pictures); for Excel, the shared-strings
table, inline strings, and drawing text boxes. Geometry, formatting, themes,
styles, etc. are left untouched, and a search term is matched even when it is
split across formatting runs.

PDF export is done with LibreOffice. Note: LibreOffice does not evaluate
field-based values (e.g. a Visio "Sheet X of Y" page-number field), so export
from the source app for those; PDF export is off by default.

Run the GUI:
    python drawing_bom_studio.py

Or use it from the command line (one or many files; rules apply to all of them):
    python drawing_bom_studio.py a.vsdx b.vsdx --find "Old" --replace "New" --pdf

Requirements:
    * Python 3.8+ (tkinter is included with the standard python.org installers)
    * LibreOffice installed (only needed for PDF export)
        - Windows:  https://www.libreoffice.org/download/
        - macOS:    brew install --cask libreoffice  (or the .dmg)
        - Linux:    sudo apt install libreoffice   (or your package manager)
"""

from __future__ import annotations

__version__ = "2.18.0 (Drawing & BOM Studio: add/remove parts on tables & BOMs)"

import argparse
import datetime
import io
import os
import posixpath
import re
import shutil
import string
import struct
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
# Visio inline formatting markers: <cp/> character, <pp/> paragraph, <tp/> tab.
_FMT_MARKER_RE = re.compile(r"<(?:cp|pp|tp)\b", re.IGNORECASE)


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
        orig_segs = list(segs)

        n = _replace_across_runs(segs, compiled)
        if not n:
            return open_tag + inner + close_tag
        total += n

        for idx, i in enumerate(text_indices):
            parts[i] = segs[idx]

        # When a match spanned several runs, the replacement lands in the first
        # run and the later runs are consumed -- but any <cp/>/<pp/>/<tp/> marker
        # that sat between them stays. A leftover <pp/> there is a stray
        # paragraph break that pushes part of the replacement onto a second
        # line. Drop a marker whose following run was consumed by the replace
        # (had real text, now blank/whitespace); boundary markers are kept.
        consumed = [orig_segs[k].strip() != "" and segs[k].strip() == ""
                    for k in range(len(segs))]
        for t in range(1, len(parts), 2):
            if not _FMT_MARKER_RE.match(parts[t]):
                continue
            run_after = (t + 1) // 2
            if run_after < len(consumed) and consumed[run_after]:
                parts[t] = ""
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


# ---------------------------------------------------------------------------
# Visio revision-history table (the chart in a corner of the cover page)
# ---------------------------------------------------------------------------
#
# A Visio cover page usually carries a revision-history block: a grid of column
# headers (REV / DESCRIPTION / DATE / APPROVED ...) with one data row per
# revision. Unlike Excel there is no real "table" -- every cell is a separately
# positioned <Shape> carrying a <Text>. We find the header shapes by their text,
# read their X positions as columns, group the data shapes into rows by their Y,
# and add a new revision row: either by filling an existing blank row, or by
# cloning the most recent row's cells just below it. Everything is geometry- and
# guard-checked; if the table can't be confidently located the page is left
# untouched.

_REVTABLE_FIELDS = {
    "Rev": {"rev", "revision", "revno", "ltr", "letter", "rev#"},
    "ECN #": {"ecn", "eco", "ecnno", "econo", "ecnnumber", "econumber",
              "ecn#", "eco#"},
    "Description": {"description", "desc", "descr", "reasonforchange",
                    "reason", "change", "changedescription",
                    "revisiondescription", "natureofchange", "nature"},
    "Date": {"date", "approvaldate", "revdate", "dateapproved",
             "erbapprovaldate"},
    "Approved By": {"approved", "approvedby", "by", "author", "engineer",
                    "eng", "drawnby", "chk", "checked", "appd", "apprd",
                    "changeauthor"},
}
REVTABLE_FIELD_ORDER = ["Rev", "ECN #", "Description", "Date", "Approved By"]


def _canonical_revtable_field(text: str) -> Optional[str]:
    """Match a shape's text to a revision-table column (headers often carry
    extra words, e.g. 'DESCRIPTION OF CHANGE')."""
    n = _norm_header(text)
    if not n:
        return None
    for field, variants in _REVTABLE_FIELDS.items():
        for v in variants:
            if n == v or (len(v) >= 4 and n.startswith(v)):
                return field
    return None


def _cluster(values, tol):
    """Greedy 1-D clustering: sorted values within ``tol`` of the running mean
    join the same cluster. Returns a list of (mean, [members])."""
    clusters = []
    for v in sorted(values, key=lambda t: t[0]):
        if clusters and abs(v[0] - clusters[-1][0]) <= tol:
            grp = clusters[-1][1]
            grp.append(v)
            clusters[-1][0] = sum(m[0] for m in grp) / len(grp)
        else:
            clusters.append([v[0], [v]])
    return [(c[0], c[1]) for c in clusters]


def _visio_leaf_cells(page_xml):
    """Leaf text-box shapes with geometry. Returns (ns, cells, max_id) where
    each cell is {id, x, y, w, h, text}. Leaf = has its own <Text> and no child
    <Shapes>, i.e. a single grid cell (not a group)."""
    try:
        root = ET.fromstring(page_xml)
    except ET.ParseError:
        return None, [], 0
    ns = root.tag[: root.tag.index("}") + 1] if root.tag.startswith("{") else ""
    cells = []
    max_id = 0
    for sh in root.iter(ns + "Shape"):
        sid = sh.get("ID")
        if sid and sid.isdigit():
            max_id = max(max_id, int(sid))
        if sh.find(ns + "Shapes") is not None:
            continue  # a group, not a single cell
        text_el = sh.find(ns + "Text")
        if text_el is None:
            continue
        cells.append({
            "id": sid,
            "x": _cell_value(sh, ns, "PinX"),
            "y": _cell_value(sh, ns, "PinY"),
            "w": _cell_value(sh, ns, "Width"),
            "h": _cell_value(sh, ns, "Height"),
            "text": "".join(text_el.itertext()).strip(),
        })
    return ns, cells, max_id


def _detect_revtable(page_xml):
    """Locate the revision table. Returns a dict with the column->x map, the
    header Y, the data rows (each {field: cell}), the row pitch and growth
    direction, or None if no confident table is found."""
    ns, cells, max_id = _visio_leaf_cells(page_xml)
    if not cells:
        return None

    geo = [c for c in cells if c["x"] is not None and c["y"] is not None]
    if len(geo) < 4:
        return None
    heights = [c["h"] for c in geo if c["h"]]
    row_tol = (sorted(heights)[len(heights) // 2] * 0.6) if heights else 0.12
    row_tol = max(row_tol, 0.04)

    # Header candidates: leaf cells whose text names a column.
    hdr = []
    for c in geo:
        cf = _canonical_revtable_field(c["text"])
        if cf:
            hdr.append((c["y"], cf, c))
    if len(hdr) < 2:
        return None

    # The header row is the Y band holding the most DISTINCT column names.
    best = None  # (distinct_count, mean_y, members)
    for mean_y, members in _cluster([(y, (cf, c)) for y, cf, c in hdr], row_tol):
        fields = {}
        for _y, (cf, c) in members:
            fields.setdefault(cf, c)  # first cell wins per field
        if best is None or len(fields) > best[0]:
            best = (len(fields), mean_y, fields)
    if best is None or best[0] < 2:
        return None
    _n, header_y, columns = best  # columns: {field: cell}
    # A real revision-history table has a Rev/Ltr column. Requiring it rejects
    # the title-block sign-off block (DRAWN BY/CHECKED BY/... each paired with a
    # DATE), which otherwise looks like a 2-column "Date + Approved By" table.
    if "Rev" not in columns:
        return None
    col_x = {f: c["x"] for f, c in columns.items()}

    # Column tolerance from the smallest gap between adjacent columns.
    xs = sorted(col_x.values())
    gaps = [b - a for a, b in zip(xs, xs[1:]) if b - a > 1e-6]
    col_tol = (min(gaps) * 0.45) if gaps else 0.5
    header_ids = {c["id"] for c in columns.values()}

    # The full set of header-row cells of THIS table (incl. columns we don't
    # write to, like ZONE), grown contiguously outward from the detected
    # columns so stray same-row text elsewhere on the page is excluded. Used to
    # reconstruct the grid lines for an appended row.
    header_pool = [c for c in geo
                   if c["w"] and abs(c["y"] - header_y) <= row_tol]
    chosen_ids = set(header_ids)
    header_cells = [c for c in header_pool if c["id"] in chosen_ids] \
        or list(columns.values())
    gap_tol = max(col_tol * 0.5, 0.05)
    changed = True
    while changed:
        changed = False
        lo = min(c["x"] - c["w"] / 2.0 for c in header_cells)
        hi = max(c["x"] + c["w"] / 2.0 for c in header_cells)
        for c in header_pool:
            if c["id"] in chosen_ids:
                continue
            cl, cr = c["x"] - c["w"] / 2.0, c["x"] + c["w"] / 2.0
            if cr >= lo - gap_tol and cl <= hi + gap_tol:
                header_cells.append(c)
                chosen_ids.add(c["id"])
                changed = True

    # Assign every other leaf cell to its nearest column (if close enough).
    placed = []  # (field, cell)
    for c in geo:
        if c["id"] in header_ids:
            continue
        nearest = min(col_x, key=lambda f: abs(c["x"] - col_x[f]))
        if abs(c["x"] - col_x[nearest]) <= col_tol:
            placed.append((nearest, c))
    if not placed:
        return None

    # Group those cells into rows by Y.
    rows = []
    for mean_y, members in _cluster(
        [(c["y"], (f, c)) for f, c in placed], row_tol
    ):
        row = {}
        for _y, (f, c) in members:
            row.setdefault(f, c)
        rows.append({"y": mean_y, "cells": row})
    if not rows:
        return None

    # The real revision rows form a regular grid butting up against the header.
    # Other shapes on the page (part number, SIZE/REV/TITLE blocks, etc.) can
    # coincidentally line up with a column, so we don't trust every aligned row.
    # Instead we walk outward from the header in fixed-pitch steps and keep only
    # the contiguous, mostly-complete rows -- this excludes stray title-block
    # text that sits far from the table.
    ncols = len(col_x)
    # A data row must fill at least about half the columns -- enough to exclude
    # stray single-cell alignments, but lenient for tables whose rows leave some
    # columns blank (no cell shape). The contiguous walk below is what actually
    # rejects far-away spurious rows.
    need = max(2, (ncols + 1) // 2)
    qual = [r for r in rows
            if len(r["cells"]) >= need and abs(r["y"] - header_y) > 1e-6]
    if not qual:
        return None
    nearest = min(qual, key=lambda r: abs(r["y"] - header_y))
    pitch = abs(nearest["y"] - header_y)
    if pitch <= 1e-6:
        return None
    slot_tol = min(row_tol, pitch * 0.45)

    def walk(direction):
        seq, k = [], 1
        while k <= 500:
            target = header_y + direction * k * pitch
            hit = None
            for r in rows:
                if (len(r["cells"]) >= need
                        and abs(r["y"] - target) <= slot_tol
                        and (hit is None
                             or abs(r["y"] - target) < abs(hit["y"] - target))):
                    hit = r
            if hit is None:
                break
            seq.append(hit)
            k += 1
        return seq

    down, up = walk(-1), walk(1)
    data = down if len(down) >= len(up) else up
    if not data:
        return None
    grow_down = data is down  # rows extend away from the header in this dir
    # ``data`` is ordered nearest-header -> farthest; the farthest row is the
    # most recent revision, and the next new row goes one pitch past it.

    return {
        "col_x": col_x, "header_y": header_y, "rows": data,
        "pitch": pitch, "grow_down": grow_down, "max_id": max_id,
        "header_cells": header_cells,
    }


# Line weight (inches) for drawn revision-row grid lines. 0.003in is ~0.22pt --
# a fine grid line; raise/lower to taste if it doesn't match a drawing's grid.
_REV_BORDER_WEIGHT = "0.003"


def _revtable_border_shape(header_cells, new_y, row_h, new_id):
    """Build a native Visio shape that draws a revision row's grid lines (the
    outer box plus the vertical column dividers), so an appended row gets the
    same borders as the rows above -- whose grid is usually a baked-in image
    that can't be extended. Returns the shape XML, or None if it can't be sized.
    """
    edges = sorted(((c["x"] - c["w"] / 2.0, c["x"] + c["w"] / 2.0)
                    for c in header_cells if c.get("w")), key=lambda e: e[0])
    if len(edges) < 2:
        return None
    left = min(e[0] for e in edges)
    right = max(e[1] for e in edges)
    width, height = right - left, row_h
    bottom = new_y - row_h / 2.0
    if width <= 0 or height <= 0:
        return None
    divs = []
    for (l1, r1), (l2, r2) in zip(edges, edges[1:]):
        dx = (r1 + l2) / 2.0 - left
        if 1e-4 < dx < width - 1e-4:
            divs.append(dx)

    rows, ix = [], 1

    def add(tag, x, y):
        nonlocal ix
        rows.append(f"<Row T='{tag}' IX='{ix}'><Cell N='X' V='{x:.6f}'/>"
                    f"<Cell N='Y' V='{y:.6f}'/></Row>")
        ix += 1

    add("MoveTo", 0, 0)  # outer box
    for x, y in ((width, 0), (width, height), (0, height), (0, 0)):
        add("LineTo", x, y)
    for dx in divs:  # interior vertical dividers
        add("MoveTo", dx, 0)
        add("LineTo", dx, height)

    geom = ("<Section N='Geometry' IX='0'><Cell N='NoFill' V='1'/>"
            "<Cell N='NoLine' V='0'/>" + "".join(rows) + "</Section>")
    return (
        f"<Shape ID='{new_id}' Type='Shape'>"
        f"<Cell N='PinX' V='{left:.6f}'/><Cell N='PinY' V='{bottom:.6f}'/>"
        f"<Cell N='Width' V='{width:.6f}'/><Cell N='Height' V='{height:.6f}'/>"
        f"<Cell N='LocPinX' V='0' F='Width*0'/>"
        f"<Cell N='LocPinY' V='0' F='Height*0'/>"
        f"<Cell N='LineWeight' V='{_REV_BORDER_WEIGHT}'/>"
        f"<Cell N='LineColor' V='0'/>"
        f"<Cell N='LinePattern' V='1'/><Cell N='FillPattern' V='0'/>"
        f"{geom}</Shape>"
    )


def _leaf_shape_span(page_xml, sid):
    """Raw [start, end) of a leaf <Shape ID="sid">...</Shape> (or self-closing).
    Attribute quoting (single or double) is handled."""
    m = re.search(r'<Shape\b[^>]*\bID=["\']' + re.escape(str(sid))
                  + r'["\'][^>]*?(/?)>', page_xml)
    if not m:
        return None
    if m.group(1) == "/":  # self-closing <Shape .../>
        return m.start(), m.end()
    end = page_xml.find("</Shape>", m.end())
    if end < 0:
        return None
    return m.start(), end + len("</Shape>")


def _set_shape_text_raw(raw: str, text: str) -> str:
    """Set a leaf shape's text in place.

    Visio stores cell text as formatting markers (<cp/>, <pp/>, <tp/>, <fld>)
    interspersed with literal characters, and the last run often ends with a
    trailing <cp/> marker and a "\\r\\n". We replace ONLY the first non-blank
    literal run, keeping every marker and the surrounding whitespace -- so the
    new text keeps the cell's font, size AND vertical alignment (dropping the
    trailing run/newline shifts a single-line cell up or down).
    """
    esc = xml_escape(text)

    def repl(m):
        inner = m.group(2)
        # Split into tags (odd indices) and literal text (even indices).
        parts = re.split(r'(<[^>]*>)', inner)
        done = False
        for i in range(0, len(parts), 2):
            seg = parts[i]
            if seg.strip():
                lead = seg[:len(seg) - len(seg.lstrip())]
                trail = seg[len(seg.rstrip()):]
                parts[i] = lead + esc + trail
                done = True
                break
        if not done:  # blank cell: put the text after the formatting markers
            parts[-1] = esc + parts[-1]
        return m.group(1) + "".join(parts) + m.group(3)

    if _TEXT_BLOCK_RE.search(raw):
        return _TEXT_BLOCK_RE.sub(repl, raw, count=1)
    close = raw.rfind("</Shape>")
    if close < 0:
        return raw
    return raw[:close] + f"<Text>{esc}</Text>" + raw[close:]


def add_revision_entry_to_page(page_xml: str, entry: dict):
    """Add a revision row to the page's revision table. Returns (xml, status):
    'filled' (reused a blank row), 'appended' (cloned a new row), 'not_found'
    (no table) or 'no_slot' (table found but no safe place to add a row)."""
    values = {f: v for f, v in (entry or {}).items() if v and v.strip()}
    if not values:
        return page_xml, "na"
    tbl = _detect_revtable(page_xml)
    if tbl is None:
        return page_xml, "not_found"
    col_x, rows = tbl["col_x"], tbl["rows"]
    writable = [f for f in values if f in col_x]
    if not writable:
        return page_xml, "not_found"

    # 1) Reuse an existing fully-blank row if there is one.
    for row in rows:
        if row["cells"] and all(not c["text"] for c in row["cells"].values()):
            edits = []  # (start, end, new_raw)
            for f in writable:
                cell = row["cells"].get(f)
                if not cell or not cell["id"]:
                    continue
                span = _leaf_shape_span(page_xml, cell["id"])
                if span:
                    raw = page_xml[span[0]:span[1]]
                    edits.append((span[0], span[1],
                                  _set_shape_text_raw(raw, values[f])))
            if not edits:
                continue
            for start, end, new_raw in sorted(edits, reverse=True):
                page_xml = page_xml[:start] + new_raw + page_xml[end:]
            return page_xml, "filled"

    # 2) Otherwise clone the most recent row's cells just past it.
    last = rows[-1]
    new_y = last["y"] - tbl["pitch"] if tbl["grow_down"] else \
        last["y"] + tbl["pitch"]
    # PinY value (single or double quoted), captured so we can swap only it.
    piny_re = re.compile(
        r'(<Cell\b[^>]*?\bN=["\']PinY["\'][^>]*?\bV=)(["\'])[^"\']*\2')
    has_formula = re.compile(
        r'<Cell\b[^>]*?\bN=["\']PinY["\'][^>]*?\bF=["\']')
    inserts = []  # (after_index, new_raw)
    next_id = tbl["max_id"]
    last_end = 0
    for f in writable:
        cell = last["cells"].get(f)
        if not cell or not cell["id"]:
            return page_xml, "no_slot"
        span = _leaf_shape_span(page_xml, cell["id"])
        if not span:
            return page_xml, "no_slot"
        raw = page_xml[span[0]:span[1]]
        if has_formula.search(raw) or not piny_re.search(raw):
            return page_xml, "no_slot"  # PinY is a formula; can't reposition
        next_id += 1
        clone = re.sub(
            r'(<Shape\b[^>]*\bID=)(["\'])\d+\2',
            lambda m: f"{m.group(1)}{m.group(2)}{next_id}{m.group(2)}",
            raw, count=1)
        clone = re.sub(r'\s+(?:UniqueID|NameU|Name)=(["\'])[^"\']*\1', "", clone)
        clone = piny_re.sub(
            lambda m: f"{m.group(1)}{m.group(2)}{new_y:.6f}{m.group(2)}",
            clone, count=1)
        clone = _set_shape_text_raw(clone, values[f])
        inserts.append((span[1], clone))
        last_end = max(last_end, span[1])
    if not inserts:
        return page_xml, "no_slot"

    # The rows above get their borders from a baked-in grid image that can't be
    # extended, so draw native grid lines for the new row (best-effort -- if it
    # can't be built we still add the text). Place it after the row's cells so
    # the lines sit on top.
    border = _revtable_border_shape(
        tbl.get("header_cells") or [], new_y, tbl["pitch"], next_id + 1)
    if border:
        inserts.append((last_end, border))

    for after, clone in sorted(inserts, reverse=True):
        page_xml = page_xml[:after] + clone + page_xml[after:]
    return page_xml, "appended"


def add_approval_to_page(page_xml: str, rev_letter: str, name: str):
    """Write an approver's name into the 'Approved By' cell of the revision-row
    whose REV letter is ``rev_letter``. Returns (xml, status): 'approved',
    'row_not_found' (no row with that letter), 'no_column' (table has no
    Approved By column) or 'not_found' (no table)."""
    if not rev_letter or not rev_letter.strip() or not name or not name.strip():
        return page_xml, "na"
    want = rev_letter.strip().upper()
    tbl = _detect_revtable(page_xml)
    if tbl is None:
        return page_xml, "not_found"
    if "Approved By" not in tbl["col_x"]:
        return page_xml, "no_column"

    target = None
    for row in tbl["rows"]:
        rc = row["cells"].get("Rev")
        if rc and rc["text"].strip().upper() == want:
            target = row
            break
    if target is None:
        return page_xml, "row_not_found"

    # The cell already exists for that column on that row: just set its text.
    cell = target["cells"].get("Approved By")
    if cell and cell["id"]:
        span = _leaf_shape_span(page_xml, cell["id"])
        if span:
            raw = page_xml[span[0]:span[1]]
            new_raw = _set_shape_text_raw(raw, name.strip())
            return page_xml[:span[0]] + new_raw + page_xml[span[1]:], "approved"

    # No Approved By cell on that row yet -> clone one from another row that has
    # it, repositioned to this row's Y (PinX already = the Approved By column).
    donor = next((r["cells"]["Approved By"] for r in tbl["rows"]
                  if r is not target and r["cells"].get("Approved By")
                  and r["cells"]["Approved By"]["id"]), None)
    if not donor:
        return page_xml, "no_column"
    span = _leaf_shape_span(page_xml, donor["id"])
    if not span:
        return page_xml, "no_column"
    raw = page_xml[span[0]:span[1]]
    piny_re = re.compile(
        r'(<Cell\b[^>]*?\bN=["\']PinY["\'][^>]*?\bV=)(["\'])[^"\']*\2')
    if re.search(r'<Cell\b[^>]*?\bN=["\']PinY["\'][^>]*?\bF=["\']', raw) \
            or not piny_re.search(raw):
        return page_xml, "no_column"
    new_id = tbl["max_id"] + 1
    clone = re.sub(r'(<Shape\b[^>]*\bID=)(["\'])\d+\2',
                   lambda m: f"{m.group(1)}{m.group(2)}{new_id}{m.group(2)}",
                   raw, count=1)
    clone = re.sub(r'\s+(?:UniqueID|NameU|Name)=(["\'])[^"\']*\1', "", clone)
    clone = piny_re.sub(
        lambda m: f"{m.group(1)}{m.group(2)}{target['y']:.6f}{m.group(2)}",
        clone, count=1)
    clone = _set_shape_text_raw(clone, name.strip())
    return page_xml[:span[1]] + clone + page_xml[span[1]:], "approved"


_EMBED_XLSX_RE = re.compile(r"visio/embeddings/.*\.xlsx$", re.IGNORECASE)


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _xml_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;"))


def _col_index(letter: str) -> int:
    """A->1, B->2, ... AA->27."""
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _xlsx_replace_member(raw: bytes, member: str, new_text: str) -> bytes:
    """Return a copy of the .xlsx byte blob with one part's text replaced."""
    src = zipfile.ZipFile(io.BytesIO(raw))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for it in src.infolist():
            data = src.read(it.filename)
            if it.filename == member:
                data = new_text.encode("utf-8")
            zout.writestr(it, data)
    return buf.getvalue()


def _set_xlsx_cell_inline(sheet_xml: str, ref: str, value: str) -> str:
    """Set worksheet cell ``ref`` (e.g. 'B5') to ``value`` as an inline string,
    preserving the cell's existing style attribute. Inserts the cell in column
    order if its row has no such cell yet."""
    esc = _xml_escape(value)
    new_cell = (f'<c r="{ref}"{{attrs}} t="inlineStr">'
                f'<is><t xml:space="preserve">{esc}</t></is></c>')

    def _clean_attrs(attrs: str) -> str:
        # Drop any existing type attribute; keep the style (s=...) and the rest.
        return re.sub(r'\s+t="[^"]*"', "", attrs)

    m = re.search(r'<c r="%s"([^>]*?)/>' % re.escape(ref), sheet_xml)
    if m:
        return (sheet_xml[:m.start()]
                + new_cell.format(attrs=_clean_attrs(m.group(1)))
                + sheet_xml[m.end():])
    m = re.search(r'<c r="%s"([^>]*?)>.*?</c>' % re.escape(ref), sheet_xml,
                  re.S)
    if m:
        return (sheet_xml[:m.start()]
                + new_cell.format(attrs=_clean_attrs(m.group(1)))
                + sheet_xml[m.end():])

    # The cell doesn't exist yet: insert it in column order within its row.
    col = re.match(r"[A-Z]+", ref).group(0)
    rownum = re.search(r"\d+", ref).group(0)
    rm = re.search(r'(<row r="%s"[^>]*>)(.*?)(</row>)' % re.escape(rownum),
                   sheet_xml, re.S)
    if not rm:
        return sheet_xml  # row missing; give up rather than corrupt the sheet
    body = rm.group(2)
    cells = list(re.finditer(r'<c r="([A-Z]+)\d+"[^>]*?(?:/>|>.*?</c>)', body,
                             re.S))
    insert_at = len(body)
    for cm in cells:
        if _col_index(cm.group(1)) > _col_index(col):
            insert_at = cm.start()
            break
    new = new_cell.format(attrs="")
    body = body[:insert_at] + new + body[insert_at:]
    return (sheet_xml[:rm.start()] + rm.group(1) + body + rm.group(3)
            + sheet_xml[rm.end():])


def _embedded_revtable_info(xlsx_bytes: bytes):
    """Inspect an embedded .xlsx (an OLE object inside a .vsdx). If its sheet
    looks like a revision-history table (a Rev column plus >=2 more known
    columns), return a dict describing it; else None.

    dict: {sheet_member, sheet_xml, col_field (col_letter->field), rev_col,
    header_row, data_revs [(rownum, rev_value)], blank_row, last_data_row}.
    """
    try:
        z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    except (zipfile.BadZipFile, OSError):
        return None
    names = z.namelist()
    sheet_member = next(
        (n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)),
        None)
    if not sheet_member:
        return None
    sst: List[str] = []
    if "xl/sharedStrings.xml" in names:
        try:
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
        except ET.ParseError:
            sroot = None
        if sroot is not None:
            for si in sroot:
                if _localname(si.tag) == "si":
                    sst.append("".join(
                        t.text or "" for t in si.iter()
                        if _localname(t.tag) == "t"))
    try:
        sheet_xml = z.read(sheet_member).decode("utf-8", "replace")
    except (KeyError, OSError):
        return None
    try:
        wroot = ET.fromstring(sheet_xml)
    except ET.ParseError:
        return None

    rows: dict[int, list[Tuple[str, Optional[str]]]] = {}
    for c in wroot.iter():
        if _localname(c.tag) != "c":
            continue
        ref = c.get("r")
        if not ref:
            continue
        cm = re.match(r"([A-Z]+)(\d+)", ref)
        if not cm:
            continue
        col, rownum = cm.group(1), int(cm.group(2))
        raw_v = None
        for ch in c:
            ln = _localname(ch.tag)
            if ln == "v":
                raw_v = ch.text
            elif ln == "is":
                raw_v = "".join(x.text or "" for x in ch.iter()
                                if _localname(x.tag) == "t")
        val = raw_v
        if c.get("t") == "s" and raw_v is not None and raw_v.isdigit():
            idx = int(raw_v)
            val = sst[idx] if idx < len(sst) else ""
        rows.setdefault(rownum, []).append((col, val))

    # Header row: the row naming the most distinct revision columns.
    header_row = None
    col_field: dict[str, str] = {}
    best = 0
    for rn in sorted(rows):
        fields: dict[str, str] = {}
        for col, val in rows[rn]:
            cf = _canonical_revtable_field(val or "")
            if cf and col not in fields:
                fields[col] = cf
        distinct = len(set(fields.values()))
        if distinct > best:
            best, header_row, col_field = distinct, rn, fields
    # Require a Rev column plus >=2 more known columns; this rejects the
    # title-block sign-off block (Approved By + Date only) and BOM tables.
    if header_row is None or len(set(col_field.values())) < 3:
        return None
    rev_col = next((c for c, f in col_field.items() if f == "Rev"), None)
    if not rev_col:
        return None

    table_rows = sorted(
        rn for rn in rows
        if rn > header_row and any(col == rev_col for col, _ in rows[rn]))
    data_revs = []
    for rn in table_rows:
        rv = dict(rows[rn]).get(rev_col)
        data_revs.append((rn, rv))
    blank_row = next((rn for rn, rv in data_revs if not (rv or "").strip()),
                     None)
    return {
        "sheet_member": sheet_member,
        "sheet_xml": sheet_xml,
        "col_field": col_field,
        "rev_col": rev_col,
        "header_row": header_row,
        "data_revs": data_revs,
        "blank_row": blank_row,
        "last_data_row": table_rows[-1] if table_rows else header_row,
    }


def _embedded_revtable_member(zin: zipfile.ZipFile) -> Optional[str]:
    """The embedding member whose worksheet is a revision table, or None.
    Some drawings store the revision history as an embedded Excel OLE object
    rather than native Visio text cells."""
    members = sorted(n for n in zin.namelist() if _EMBED_XLSX_RE.search(n))
    for m in members:
        try:
            if _embedded_revtable_info(zin.read(m)) is not None:
                return m
        except (KeyError, OSError):
            continue
    return None


def embedded_revtable_columns(xlsx_bytes: bytes) -> Optional[List[str]]:
    info = _embedded_revtable_info(xlsx_bytes)
    if not info:
        return None
    present = set(info["col_field"].values())
    return [f for f in REVTABLE_FIELD_ORDER if f in present]


def embedded_revtable_rev_letters(xlsx_bytes: bytes) -> List[str]:
    info = _embedded_revtable_info(xlsx_bytes)
    if not info:
        return []
    return [(rv or "").strip() for _rn, rv in info["data_revs"]
            if (rv or "").strip()]


def add_revision_entry_to_embedded(xlsx_bytes: bytes, entry: dict):
    """Add a revision row to an embedded-Excel revision table. Returns
    (bytes, status): 'filled' (reused a pre-formatted blank row), 'appended'
    (added a new row), 'not_found' or 'na'."""
    values = {f: v.strip() for f, v in (entry or {}).items()
              if v and v.strip()}
    if not values:
        return xlsx_bytes, "na"
    info = _embedded_revtable_info(xlsx_bytes)
    if not info:
        return xlsx_bytes, "not_found"
    writable = {col: f for col, f in info["col_field"].items()
                if f in values}
    if not writable:
        return xlsx_bytes, "not_found"

    sheet_xml = info["sheet_xml"]
    target = info["blank_row"]
    status = "filled"
    if target is None:
        # No pre-formatted blank row left: clone the last data row's cells
        # onto a fresh row just below it so the new entry keeps the grid.
        target = info["last_data_row"] + 1
        sheet_xml = _append_embedded_row(
            sheet_xml, info["last_data_row"], target)
        status = "appended"
    for col, f in writable.items():
        sheet_xml = _set_xlsx_cell_inline(sheet_xml, f"{col}{target}",
                                          values[f])
    new_bytes = _xlsx_replace_member(xlsx_bytes, info["sheet_member"],
                                     sheet_xml)
    # The OLE object only displays the range named by <oleSize>; extend it so
    # the new row is actually inside the area Visio shows (THIS is what makes
    # the row appear in the drawing without opening the object).
    new_bytes = _extend_olesize(new_bytes, target)
    return new_bytes, status


def _extend_olesize(xlsx_bytes: bytes, end_row: int,
                    exact: bool = False) -> bytes:
    """Adjust the embedded workbook's <oleSize ref="A1:E7"/> bottom row -- the
    range the embedded object actually shows in the drawing.

    By default the range is only *grown* so it covers ``end_row`` (rows are
    never hidden by accident). With ``exact=True`` the bottom row is set to
    ``end_row`` even if that shrinks the range -- used after a removal so the
    now-blank trailing row isn't displayed. No-op if there's no oleSize."""
    try:
        z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    except (zipfile.BadZipFile, OSError):
        return xlsx_bytes
    if "xl/workbook.xml" not in z.namelist():
        return xlsx_bytes
    wb = z.read("xl/workbook.xml").decode("utf-8", "replace")
    m = re.search(r'<oleSize ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', wb)
    if not m:
        return xlsx_bytes
    cur_end = int(m.group(4))
    if cur_end == end_row or (not exact and cur_end >= end_row):
        return xlsx_bytes
    new_ref = (f'<oleSize ref="{m.group(1)}{m.group(2)}:'
               f'{m.group(3)}{end_row}"')
    wb = wb[:m.start()] + new_ref + wb[m.end():]
    return _xlsx_replace_member(xlsx_bytes, "xl/workbook.xml", wb)


def _append_embedded_row(sheet_xml: str, src_row: int, new_row: int) -> str:
    """Clone worksheet row ``src_row`` as a blank row numbered ``new_row``
    (preserving cell styles, clearing values) and extend the sheet dimension."""
    rm = re.search(r'<row r="%d"[^>]*>.*?</row>' % src_row, sheet_xml, re.S)
    if not rm:
        return sheet_xml
    block = rm.group(0)
    block = re.sub(r'(<row r=")%d(")' % src_row,
                   lambda m: f"{m.group(1)}{new_row}{m.group(2)}", block,
                   count=1)

    # Renumber each cell's row and strip its value, keeping the style attr.
    def _blank_cell(m):
        attrs = re.sub(r'\s+t="[^"]*"', "", m.group(2))
        return f'<c r="{m.group(1)}{new_row}"{attrs}/>'

    block = re.sub(r'<c r="([A-Z]+)%d"([^>]*?)(?:/>|>.*?</c>)' % src_row,
                   _blank_cell, block, flags=re.S)
    sheet_xml = sheet_xml[:rm.end()] + block + sheet_xml[rm.end():]
    # Extend <dimension ref="A1:Jn"> if the new row is past the current end.
    def _bump_dim(m):
        start, end = m.group(1), m.group(2)
        cm = re.match(r"([A-Z]+)(\d+)", end)
        if cm and int(cm.group(2)) < new_row:
            end = f"{cm.group(1)}{new_row}"
        return f'<dimension ref="{start}:{end}"'
    sheet_xml = re.sub(r'<dimension ref="([A-Z]+\d+):([A-Z]+\d+)"',
                       _bump_dim, sheet_xml, count=1)
    return sheet_xml


def add_approval_to_embedded(xlsx_bytes: bytes, rev_letter: str, name: str):
    """Write an approver's name into the Approved By cell of the embedded
    revision row whose REV letter matches. Returns (bytes, status)."""
    if not rev_letter or not rev_letter.strip() or not name or not name.strip():
        return xlsx_bytes, "na"
    info = _embedded_revtable_info(xlsx_bytes)
    if not info:
        return xlsx_bytes, "not_found"
    appr_col = next((c for c, f in info["col_field"].items()
                     if f == "Approved By"), None)
    if not appr_col:
        return xlsx_bytes, "no_column"
    want = rev_letter.strip().upper()
    target = next((rn for rn, rv in info["data_revs"]
                   if (rv or "").strip().upper() == want), None)
    if target is None:
        return xlsx_bytes, "row_not_found"
    sheet_xml = _set_xlsx_cell_inline(info["sheet_xml"],
                                      f"{appr_col}{target}", name.strip())
    new_bytes = _xlsx_replace_member(xlsx_bytes, info["sheet_member"],
                                     sheet_xml)
    # The approved row may itself be past the displayed range; include it.
    return _extend_olesize(new_bytes, target), "approved"


# --- Cached OLE presentation (EMF) -------------------------------------------
#
# Visio displays an embedded Excel object from a *cached* picture -- an EMF
# metafile linked from the embedding's .rels -- and only re-renders it when the
# object is activated (double-clicked) in Visio. So after editing the worksheet
# data we must also patch that EMF, or the drawing keeps showing the old table
# until the user opens it by hand. The EMF already draws the grid for the blank
# rows; we just inject the new row's text (EMR_EXTTEXTOUTW records), reusing the
# metafile's own measured glyph widths so spacing matches.

_EMF_EXTTEXTOUTW = 84


def _emf_text_records(emf: bytes):
    """Parse the EXTTEXTOUTW records out of an EMF. Each item:
    {off, size, bounds, refx, refy, nchars, text, dx}."""
    out = []
    off = 0
    n = len(emf)
    while off + 8 <= n:
        rtype, rsize = struct.unpack_from("<II", emf, off)
        if rsize < 8 or rsize % 4 or off + rsize > n:
            break
        if rtype == _EMF_EXTTEXTOUTW and rsize >= 76:
            bounds = struct.unpack_from("<4i", emf, off + 8)
            refx, refy = struct.unpack_from("<2i", emf, off + 36)
            nch, off_str = struct.unpack_from("<II", emf, off + 44)
            off_dx = struct.unpack_from("<I", emf, off + 72)[0]
            try:
                text = emf[off + off_str:off + off_str + nch * 2].decode(
                    "utf-16-le", "replace")
            except Exception:  # noqa: BLE001
                text = ""
            dx = (list(struct.unpack_from("<%di" % nch, emf, off + off_dx))
                  if off_dx and nch else [])
            out.append({"off": off, "size": rsize, "bounds": bounds,
                        "refx": refx, "refy": refy, "nchars": nch,
                        "text": text, "dx": dx})
        off += rsize
    return out


def _emf_char_widths(records):
    """Build a glyph-advance lookup from the metafile's own text, plus a
    sensible default, so synthesised rows are spaced like the existing ones."""
    acc: dict[str, list] = {}
    for r in records:
        for ch, d in zip(r["text"], r["dx"]):
            acc.setdefault(ch, []).append(d)
    widths = {c: round(sum(v) / len(v)) for c, v in acc.items()}
    all_dx = [d for v in acc.values() for d in v]
    default = round(sorted(all_dx)[len(all_dx) // 2]) if all_dx else 8
    return widths, default


def _emf_rows(records):
    """Cluster text records into rows by their Y, and pick out the header row
    (the one naming the most revision columns). Returns (header, data_rows)
    where header is (y, {field: x}) and data_rows is [(y, [records])]."""
    ys = sorted({r["refy"] for r in records})
    bands: list = []
    for y in ys:
        if bands and y - bands[-1][-1] <= 5:
            bands[-1].append(y)
        else:
            bands.append([y])
    rows = []
    for band in bands:
        cells = [r for r in records if r["refy"] in band]
        rows.append((min(band), cells))
    rows.sort()
    header = None
    for y, cells in rows:
        fields = {}
        for c in cells:
            f = _canonical_revtable_field(c["text"])
            if f:
                fields.setdefault(f, c["refx"])
        if header is None or len(fields) > len(header[1]):
            header = (y, fields)
    if header is None or "Rev" not in header[1]:
        return None, []
    data_rows = [(y, cells) for y, cells in rows if y > header[0]]
    return header, data_rows


def _build_emf_text_record(emf: bytes, tmpl, new_text: str, new_y: int,
                           charw, default_w) -> bytes:
    """Clone an EXTTEXTOUTW record at a new Y with new text, keeping the
    template's font/colour/graphics state (its leading 76 bytes).

    We deliberately emit NO intercharacter-spacing (Dx) array (offDx = 0). With
    no Dx, the text is laid out with the font's own glyph advances -- exactly
    what Excel uses when it draws the cell -- so the spacing matches instead of
    relying on our estimated per-glyph widths. The estimated width is still used
    only for the bounding rectangle, which is just a hint and doesn't affect how
    the characters are spaced."""
    fixed = bytearray(emf[tmpl["off"]:tmpl["off"] + 76])
    dy = new_y - tmpl["refy"]
    width = sum(charw.get(ch, default_w) for ch in new_text)  # bounds hint only
    left = tmpl["refx"]
    struct.pack_into("<4i", fixed, 8, left, tmpl["bounds"][1] + dy,
                     left + width, tmpl["bounds"][3] + dy)   # rclBounds
    struct.pack_into("<2i", fixed, 36, tmpl["refx"], new_y)  # ptlReference
    strb = new_text.encode("utf-16-le")
    if len(strb) % 4:
        strb += b"\x00" * (4 - len(strb) % 4)
    struct.pack_into("<I", fixed, 44, len(new_text))         # nChars
    struct.pack_into("<I", fixed, 48, 76)                    # offString
    struct.pack_into("<I", fixed, 72, 0)                     # offDx = 0 (none)
    rec = bytes(fixed) + strb
    return rec[:4] + struct.pack("<I", len(rec)) + rec[8:]   # nSize


def _emf_set_counts(emf: bytearray, byte_delta: int, record_delta: int):
    n_bytes, n_records = struct.unpack_from("<II", emf, 48)
    struct.pack_into("<II", emf, 48, n_bytes + byte_delta,
                     n_records + record_delta)


def _emf_revrow_growth(emf: bytes) -> float:
    """How much taller the cached picture must get to show one more revision
    row: the ratio new_height / old_height (1.0 if the new row already fits in
    the metafile's bounds). Computed without modifying the metafile."""
    try:
        records = _emf_text_records(emf)
        if not records:
            return 1.0
        header, data_rows = _emf_rows(records)
        if not data_rows:
            return 1.0
        pitch = (data_rows[1][0] - data_rows[0][0] if len(data_rows) > 1
                 else data_rows[0][0] - header[0])
        if pitch <= 0:
            return 1.0
        bounds = struct.unpack_from("<4i", emf, 8)
        old_h = bounds[3] - bounds[1]
        bottom = data_rows[-1][0] + 2 * pitch
        if bottom > bounds[3] and old_h > 0:
            return (bottom - bounds[1]) / old_h
        return 1.0
    except (struct.error, ValueError, IndexError):
        return 1.0


def add_revision_row_to_emf(emf: bytes, entry: dict) -> Optional[bytes]:
    """Inject the new revision row's text into the cached EMF so Visio shows it
    without the user activating the object. Returns patched bytes, or None if
    the metafile couldn't be patched (caller then leaves it as-is)."""
    values = {f: v.strip() for f, v in (entry or {}).items()
              if v and v.strip()}
    if not values:
        return None
    try:
        records = _emf_text_records(emf)
        if not records:
            return None
        header, data_rows = _emf_rows(records)
        if not data_rows:
            return None
        charw, default_w = _emf_char_widths(records)
        pitch = (data_rows[1][0] - data_rows[0][0] if len(data_rows) > 1
                 else data_rows[0][0] - header[0])
        if pitch <= 0:
            return None
        _last_y, last_cells = data_rows[-1]
        head_cols = header[1]
        additions = b""
        added = 0
        insert_at = 0
        for cell in last_cells:
            insert_at = max(insert_at, cell["off"] + cell["size"])
            field = min(head_cols, key=lambda f: abs(head_cols[f] - cell["refx"]))
            if field not in values:
                continue
            additions += _build_emf_text_record(
                emf, cell, values[field], cell["refy"] + pitch,
                charw, default_w)
            added += 1
        if not added:
            return None
        out = bytearray(emf[:insert_at] + additions + emf[insert_at:])
        _emf_set_counts(out, len(additions), added)
        # Grow the metafile's bounds AND frame (in step) so the new row sits in
        # view at the same scale, instead of below the visible window.
        bounds = list(struct.unpack_from("<4i", out, 8))
        old_h = bounds[3] - bounds[1]
        bottom = data_rows[-1][0] + 2 * pitch
        if bottom > bounds[3] and old_h > 0:
            ratio = (bottom - bounds[1]) / old_h
            bounds[3] = bottom
            struct.pack_into("<4i", out, 8, *bounds)
            frame = list(struct.unpack_from("<4i", out, 24))
            fh = frame[3] - frame[1]
            frame[3] = frame[1] + int(round(fh * ratio))
            struct.pack_into("<4i", out, 24, *frame)
        return bytes(out)
    except (struct.error, ValueError, IndexError):
        return None


def _ole_shape_for_embedding(zin: zipfile.ZipFile, embed_member: str):
    """Find the page part and shape ID of the OLE object that displays an
    embedded worksheet. Returns (page_part, shape_id) or (None, None)."""
    for n in zin.namelist():
        if not re.match(r"visio/pages/page\d+\.xml$", n):
            continue
        rels = posixpath.join("visio/pages", "_rels",
                              posixpath.basename(n) + ".rels")
        try:
            rtext = zin.read(rels).decode("utf-8", "replace")
        except (KeyError, OSError):
            continue
        rid = None
        for rid_m, tgt in re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"',
                                     rtext):
            resolved = posixpath.normpath(posixpath.join("visio/pages", tgt))
            if resolved == embed_member:
                rid = rid_m
                break
        if not rid:
            continue
        page_xml = zin.read(n).decode("utf-8", "replace")
        pos = page_xml.find("r:id='%s'" % rid)
        if pos < 0:
            pos = page_xml.find('r:id="%s"' % rid)
        if pos < 0:
            continue
        shapes = list(re.finditer(r"<Shape\b[^>]*\bID=['\"](\d+)['\"]",
                                   page_xml[:pos]))
        if shapes:
            return n, shapes[-1].group(1)
    return None, None


def grow_ole_shape_height(page_xml: str, shape_id: str, ratio: float) -> str:
    """Grow an OLE object shape's height (and ObjectHeight) by ``ratio``,
    extending downward (top edge fixed), so its display window shows the extra
    revision row. Returns the edited page XML."""
    if ratio <= 1.0:
        return page_xml
    m = re.search(r"<Shape\b[^>]*\bID=['\"]%s['\"].*?</Shape>"
                  % re.escape(shape_id), page_xml, re.S)
    if not m:
        return page_xml
    block = m.group(0)
    hm = re.search(r"<Cell N=['\"]Height['\"][^>]*?\bV=['\"]([\d.]+)['\"]",
                   block)
    if not hm:
        return page_xml
    height = float(hm.group(1))
    new_height = height * ratio
    d_h = new_height - height

    def _set(blk, name, value):
        return re.sub(
            r"(<Cell N=['\"]%s['\"][^>]*?\bV=['\"])[\d.]+(['\"])" % name,
            lambda mm: f"{mm.group(1)}{value:.6f}{mm.group(2)}", blk, count=1)

    block = _set(block, "Height", new_height)
    lm = re.search(r"<Cell N=['\"]LocPinY['\"][^>]*?\bV=['\"]([\d.]+)['\"]",
                   block)
    if lm:  # keep the pin at the top edge so the shape grows downward
        block = _set(block, "LocPinY", float(lm.group(1)) + d_h)
    om = re.search(r"ObjectHeight=['\"]([\d.]+)['\"]", block)
    if om:
        block = re.sub(
            r"(ObjectHeight=['\"])[\d.]+(['\"])",
            lambda mm: f"{mm.group(1)}{float(om.group(1)) * ratio:.10f}"
                       f"{mm.group(2)}", block, count=1)
    return page_xml[:m.start()] + block + page_xml[m.end():]


def approve_revision_in_emf(emf: bytes, rev_letter: str,
                            name: str) -> Optional[bytes]:
    """Write an approver's name into the cached EMF's Approved cell for the row
    with ``rev_letter``. Returns patched bytes, or None if not patchable."""
    if not rev_letter or not rev_letter.strip() or not name or not name.strip():
        return None
    try:
        records = _emf_text_records(emf)
        if not records:
            return None
        header, data_rows = _emf_rows(records)
        if not data_rows or "Approved By" not in header[1] \
                or "Rev" not in header[1]:
            return None
        header_fields = header[1]
        charw, default_w = _emf_char_widths(records)
        rev_x = header_fields["Rev"]
        appr_x = header_fields["Approved By"]
        want = rev_letter.strip().upper()

        def nearest_field(cx):
            return min(header_fields, key=lambda f: abs(header_fields[f] - cx))

        target_row = None
        for y, cells in data_rows:
            rev_cell = min(cells, key=lambda c: abs(c["refx"] - rev_x))
            if rev_cell["text"].strip().upper() == want:
                target_row = (y, cells)
                break
        if target_row is None:
            return None
        y, cells = target_row
        # The Approved cell is one whose column (by nearest header) really is
        # "Approved By" -- so a blank Approved column doesn't get mistaken for
        # the Date cell next to it.
        appr_cells = [c for c in cells
                      if nearest_field(c["refx"]) == "Approved By"]
        if appr_cells:
            appr_cell = min(appr_cells, key=lambda c: abs(c["refx"] - appr_x))
            new_rec = _build_emf_text_record(emf, appr_cell, name.strip(),
                                             appr_cell["refy"], charw,
                                             default_w)
            out = bytearray(emf[:appr_cell["off"]] + new_rec
                            + emf[appr_cell["off"] + appr_cell["size"]:])
            _emf_set_counts(out, len(new_rec) - appr_cell["size"], 0)
            return bytes(out)
        # No Approved cell drawn on that row yet: add one at the Approved
        # column's X, cloning a sibling cell's text state, inserted after the
        # row's last cell so the same font/colour is in effect.
        tmpl = dict(cells[0])
        tmpl["refx"] = appr_x
        tmpl["bounds"] = (appr_x, cells[0]["bounds"][1], appr_x,
                          cells[0]["bounds"][3])
        new_rec = _build_emf_text_record(emf, tmpl, name.strip(), y,
                                         charw, default_w)
        insert_at = max(c["off"] + c["size"] for c in cells)
        out = bytearray(emf[:insert_at] + new_rec + emf[insert_at:])
        _emf_set_counts(out, len(new_rec), 1)
        return bytes(out)
    except (struct.error, ValueError, IndexError):
        return None


def _embedded_emf_member(zin: zipfile.ZipFile,
                         xlsx_member: str) -> Optional[str]:
    """The cached-presentation EMF linked from an embedded worksheet's .rels."""
    member, _ext = _embedded_presentation(zin, xlsx_member)
    return member if (_ext == "emf") else None


def _embedded_presentation(zin: zipfile.ZipFile, xlsx_member: str):
    """The cached-presentation image an OLE object is displayed from, of ANY
    format. Returns (member, ext) -- ext lower-case without the dot -- or
    (None, None). Only EMF can be patched; other formats (PNG/BMP/WMF/...) are
    raster/opaque, so the picture can't be edited in place."""
    rels = posixpath.join(posixpath.dirname(xlsx_member), "_rels",
                          posixpath.basename(xlsx_member) + ".rels")
    names = set(zin.namelist())
    if rels not in names:
        return None, None
    try:
        text = zin.read(rels).decode("utf-8", "replace")
    except (KeyError, OSError):
        return None, None
    # Prefer an EMF (patchable); otherwise take whatever image is linked.
    targets = re.findall(
        r'Type="[^"]*/image"[^>]*Target="([^"]+)"'
        r'|Target="([^"]+)"[^>]*Type="[^"]*/image"', text, re.IGNORECASE)
    images = [a or b for a, b in targets]
    if not images:
        images = re.findall(
            r'Target="([^"]+\.(?:emf|wmf|png|bmp|jpe?g|gif|tiff?))"', text,
            re.IGNORECASE)
    if not images:
        return None, None
    images.sort(key=lambda t: 0 if t.lower().endswith(".emf") else 1)
    chosen = images[0]
    resolved = posixpath.normpath(
        posixpath.join(posixpath.dirname(xlsx_member), chosen))
    if resolved not in names:
        return None, None
    ext = chosen.rsplit(".", 1)[-1].lower() if "." in chosen else ""
    return resolved, ext


# Backwards-compatible alias (the revision-table code used the old name).
_embedded_revtable_emf_member = _embedded_emf_member


# ---------------------------------------------------------------------------
# Visio parts tables ("Cable BOM"): the find-&-edit-rows feature for .vsdx
# ---------------------------------------------------------------------------
#
# The R/C/P/W drawing sheets carry a parts list at the bottom, stored the same
# way as the revision table -- an embedded Excel "Cable BOM" worksheet shown via
# a cached EMF. We locate those worksheets, find rows by Part Number, and edit
# their cells (updating both the worksheet data and the cached picture).

_VISIO_BOM_FIELDS = {
    "Item": {"item", "itemno", "itemnumber", "item#", "no"},
    "Ref/DES #": {"refdes", "refdes#", "referencedesignator", "designator",
                  "refdesignator", "refdesno"},
    "Cable Name": {"cablename", "cable", "wirename"},
    "Description": {"description", "desc", "descr"},
    "Part Number": {"partnumber", "partno", "partnum", "pn", "part", "pno"},
    "Manufacturer": {"manufacturer", "mfg", "mfr", "manuf", "make", "vendor"},
    "Qty/Length": {"qtylength", "qty", "quantity", "length", "qtylen", "qnty"},
    "Unit": {"unit", "units", "uom"},
}
VISIO_BOM_FIELD_ORDER = ["Item", "Ref/DES #", "Cable Name", "Description",
                         "Part Number", "Manufacturer", "Qty/Length", "Unit"]
# Shown as editable (the find key Part Number and the auto Item are omitted).
VISIO_BOM_EDIT_FIELDS = ["Ref/DES #", "Cable Name", "Description",
                         "Manufacturer", "Qty/Length", "Unit"]
VISIO_BOM_COPYDOWN = ["Manufacturer", "Description", "Qty/Length", "Unit"]

# Part numbers in the parts table are sometimes written "<P/N> or equiv." -- a
# Find value that names just the part number should still match such a cell.
_OR_EQUIV_RE = re.compile(r"\s*\bor\s+equiv(?:alent)?\.?\s*$", re.IGNORECASE)


def _strip_or_equiv(text: str) -> str:
    """Drop a trailing 'or equiv.'/'or equivalent' so the bare part number is
    left for matching (e.g. 'AL10 or equiv.' -> 'AL10')."""
    return _OR_EQUIV_RE.sub("", text or "").strip()


def _canonical_bom_field(text: str) -> Optional[str]:
    n = _norm_header(text)
    if not n:
        return None
    for field, variants in _VISIO_BOM_FIELDS.items():
        if n in variants or any(len(v) >= 4 and n.startswith(v)
                                for v in variants):
            return field
    return None


def _read_embedded_sheet_rows(xlsx_bytes: bytes):
    """Parse an embedded .xlsx's first worksheet. Returns
    (sheet_member, sheet_xml, rows) where rows is {rownum: {col: value}}."""
    try:
        z = zipfile.ZipFile(io.BytesIO(xlsx_bytes))
    except (zipfile.BadZipFile, OSError):
        return None
    names = z.namelist()
    sheet_member = next(
        (n for n in names if re.match(r"xl/worksheets/sheet\d+\.xml$", n)),
        None)
    if not sheet_member:
        return None
    sst: List[str] = []
    if "xl/sharedStrings.xml" in names:
        try:
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
        except ET.ParseError:
            sroot = None
        if sroot is not None:
            for si in sroot:
                if _localname(si.tag) == "si":
                    sst.append("".join(
                        t.text or "" for t in si.iter()
                        if _localname(t.tag) == "t"))
    try:
        sheet_xml = z.read(sheet_member).decode("utf-8", "replace")
    except (KeyError, OSError):
        return None
    try:
        wroot = ET.fromstring(sheet_xml)
    except ET.ParseError:
        return None
    rows: dict[int, dict[str, str]] = {}
    for c in wroot.iter():
        if _localname(c.tag) != "c":
            continue
        ref = c.get("r")
        if not ref:
            continue
        cm = re.match(r"([A-Z]+)(\d+)", ref)
        if not cm:
            continue
        col, rownum = cm.group(1), int(cm.group(2))
        raw_v = None
        for ch in c:
            ln = _localname(ch.tag)
            if ln == "v":
                raw_v = ch.text
            elif ln == "is":
                raw_v = "".join(x.text or "" for x in ch.iter()
                                if _localname(x.tag) == "t")
        val = raw_v
        if c.get("t") == "s" and raw_v is not None and raw_v.isdigit():
            idx = int(raw_v)
            val = sst[idx] if idx < len(sst) else ""
        if val is not None:
            rows.setdefault(rownum, {})[col] = val
    return sheet_member, sheet_xml, rows


def _embedded_bom_info(xlsx_bytes: bytes):
    """If an embedded worksheet is a Cable BOM parts table (a Part Number column
    plus >=2 more known columns), return a dict describing it; else None."""
    parsed = _read_embedded_sheet_rows(xlsx_bytes)
    if not parsed:
        return None
    sheet_member, sheet_xml, rows = parsed
    header_row = None
    col_field: dict[str, str] = {}
    best = 0
    for rn in sorted(rows):
        fields: dict[str, str] = {}
        for col, val in rows[rn].items():
            cf = _canonical_bom_field(val or "")
            if cf and col not in fields:
                fields[col] = cf
        distinct = len(set(fields.values()))
        if distinct > best:
            best, header_row, col_field = distinct, rn, fields
    if header_row is None or "Part Number" not in col_field.values() \
            or len(set(col_field.values())) < 3:
        return None
    return {"sheet_member": sheet_member, "sheet_xml": sheet_xml,
            "rows": rows, "col_field": col_field, "header_row": header_row}


def _embedding_page_names(zin: zipfile.ZipFile) -> dict:
    """Map each embedding member -> the display name of the page that shows it
    (so a parts table can be labelled with its sheet, e.g. 'R0001')."""
    out: dict[str, str] = {}
    # page part -> display name, from pages.xml + its rels.
    try:
        pages_xml = zin.read("visio/pages/pages.xml").decode("utf-8", "replace")
        prels = zin.read(
            "visio/pages/_rels/pages.xml.rels").decode("utf-8", "replace")
    except (KeyError, OSError):
        return out
    relmap = dict(re.findall(r'Id="([^"]+)"[^>]*?Target="([^"]+)"', prels))
    rel_ns = ("{http://schemas.openxmlformats.org/officeDocument/2006/"
              "relationships}id")
    part_name: dict[str, str] = {}
    try:
        proot = ET.fromstring(pages_xml)
    except ET.ParseError:
        proot = None
    if proot is not None:
        for pg in proot.iter():
            if _localname(pg.tag) != "Page":
                continue
            name = pg.get("Name") or pg.get("NameU") or ""
            rid = None
            for ch in pg:
                if _localname(ch.tag) == "Rel":
                    rid = ch.get(rel_ns)
            tgt = relmap.get(rid or "")
            if tgt:
                resolved = posixpath.normpath(
                    posixpath.join("visio/pages", tgt))
                part_name[resolved] = name
    # For each page part, read its rels and attribute embeddings to that page.
    for part, name in part_name.items():
        rels = posixpath.join(posixpath.dirname(part), "_rels",
                              posixpath.basename(part) + ".rels")
        try:
            text = zin.read(rels).decode("utf-8", "replace")
        except (KeyError, OSError):
            continue
        for tgt in re.findall(r'Target="([^"]*\.xlsx)"', text, re.IGNORECASE):
            resolved = posixpath.normpath(
                posixpath.join(posixpath.dirname(part), tgt))
            out.setdefault(resolved, name)
    return out


def _embedded_bom_members(zin: zipfile.ZipFile) -> List[str]:
    """Every embedding member whose worksheet is a Cable BOM parts table."""
    out = []
    for n in sorted(zin.namelist()):
        if not _EMBED_XLSX_RE.search(n):
            continue
        try:
            if _embedded_bom_info(zin.read(n)) is not None:
                out.append(n)
        except (KeyError, OSError):
            continue
    return out


def visio_bom_scan_rows(path, part_numbers, case_sensitive=False):
    """Find parts-table rows whose Part Number matches, across every Cable BOM
    embedded in a .vsdx. Each match: {file, embed, emf, sheet_name, part, row,
    matched, fields:{field:(ref,value)}}."""
    matches = []
    targets = [(p, p if case_sensitive else p.lower())
               for p in part_numbers if p]
    if not targets:
        return matches
    try:
        zin = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return matches
    with zin:
        page_names = _embedding_page_names(zin)
        for member in _embedded_bom_members(zin):
            info = _embedded_bom_info(zin.read(member))
            if not info:
                continue
            rows, col_field = info["rows"], info["col_field"]
            hrow = info["header_row"]
            pn_col = next(c for c, f in col_field.items()
                          if f == "Part Number")
            emf = _embedded_emf_member(zin, member)
            for rn in sorted(rows):
                if rn <= hrow:
                    continue
                pnval = (rows[rn].get(pn_col) or "").strip()
                if not pnval:
                    continue
                # Match against the bare part number, ignoring a trailing
                # "or equiv." so "AL10 or equiv." still matches Find "AL10".
                base = _strip_or_equiv(pnval)
                cmp_val = base if case_sensitive else base.lower()
                hit = next((o for o, t in targets if t == cmp_val), None)
                if hit is None:
                    continue
                fields = {}
                for col, cf in col_field.items():
                    ref = f"{col}{rn}"
                    fields[cf] = (ref, (rows[rn].get(col) or ""))
                item_cell = fields.get("Item")
                matches.append({
                    "file": str(path), "embed": member, "emf": emf,
                    "sheet_name": page_names.get(member, ""),
                    "part": pnval, "row": rn, "matched": hit, "fields": fields,
                    # The Item value and the 0-based data-row ordinal uniquely
                    # identify the row in the cached picture (the Part Number
                    # is NOT always unique -- e.g. "See Specifications").
                    "item": (item_cell[1] if item_cell else ""),
                    "row_index": rn - hrow - 1})
    return matches


def build_visio_bom_edits(path, bom_edits, case_sensitive=False) -> dict:
    """Map staged Visio BOM field edits to per-embedding cell edits for one
    file. bom_edits: {part_number: {field: new_value}}. Returns
    {embed_member: {"emf": emf_member, "cells": [edit, ...]}} where each edit is
    {ref, col, row, field, old, new, pn}."""
    out: dict = {}
    if not bom_edits:
        return out
    for m in visio_bom_scan_rows(path, list(bom_edits.keys()), case_sensitive):
        for field, new_value in bom_edits.get(m["matched"], {}).items():
            cell = m["fields"].get(field)
            if cell is None:
                continue
            ref, old = cell
            if new_value == old:
                continue
            col = re.match(r"[A-Z]+", ref).group(0)
            entry = out.setdefault(m["embed"],
                                   {"emf": m["emf"], "cells": []})
            entry["cells"].append({
                "ref": ref, "col": col, "row": m["row"], "field": field,
                "old": old, "new": new_value, "pn": m["part"],
                "item": m.get("item", ""), "row_index": m.get("row_index")})
    return out


def build_visio_pn_replacements(path, pairs, case_sensitive=False) -> dict:
    """Find/Replace applied to the Part Number column of the parts tables.

    For each find->replace pair, any parts-table row whose Part Number matches
    the find value (ignoring a trailing 'or equiv.') has its **whole** Part
    Number cell replaced with the replacement value. Returns the same shape as
    build_visio_bom_edits so the two can be merged."""
    out: dict = {}
    repl = {}
    for find, replace in pairs or []:
        if find and find not in repl:
            repl[find] = replace
    if not repl:
        return out
    for m in visio_bom_scan_rows(path, list(repl.keys()), case_sensitive):
        new_value = repl.get(m["matched"])
        if new_value is None:
            continue
        cell = m["fields"].get("Part Number")
        if cell is None:
            continue
        ref, old = cell  # old is the full cell, e.g. "AL10 or equiv."
        if new_value == old:
            continue
        col = re.match(r"[A-Z]+", ref).group(0)
        entry = out.setdefault(m["embed"], {"emf": m["emf"], "cells": []})
        entry["cells"].append({
            "ref": ref, "col": col, "row": m["row"], "field": "Part Number",
            "old": old, "new": new_value, "pn": m["part"],
            "item": m.get("item", ""), "row_index": m.get("row_index")})
    return out


def _merge_bom_edit_dicts(*dicts) -> dict:
    """Merge several {embed: {emf, cells}} maps into one (concatenating cells)."""
    merged: dict = {}
    for d in dicts:
        for embed, ed in (d or {}).items():
            m = merged.setdefault(embed, {"emf": ed.get("emf"), "cells": []})
            if not m.get("emf"):
                m["emf"] = ed.get("emf")
            m["cells"].extend(ed.get("cells", []))
            # An add/remove carries the exact final display range; keep it.
            if ed.get("ole_end") is not None:
                m["ole_end"] = ed["ole_end"]
            if ed.get("add_rows"):
                m.setdefault("add_rows", []).extend(ed["add_rows"])
    return merged


def _emf_bom_header(records):
    """For a parts-table EMF: (header_fields {field: x}, header_bottom_y,
    col_tol). header_fields includes Part Number. Returns None if not a BOM."""
    header_fields: dict[str, int] = {}
    header_bottom = None
    for r in records:
        f = _canonical_bom_field(r["text"])
        if f and f not in header_fields:
            header_fields[f] = r["refx"]
            header_bottom = (r["refy"] if header_bottom is None
                             else max(header_bottom, r["refy"]))
    if "Part Number" not in header_fields or len(header_fields) < 3:
        return None
    xs = sorted(header_fields.values())
    gaps = [b - a for a, b in zip(xs, xs[1:]) if b - a > 0]
    col_tol = (min(gaps) * 0.45) if gaps else 60
    return header_fields, header_bottom, col_tol


def apply_bom_edits_to_emf(emf: bytes, cells: list) -> Optional[bytes]:
    """Patch the cached parts-table EMF for a set of cell edits. Each edit:
    {col, row, field, old, new, pn}.

    Data text is left-aligned while headers are positioned differently, so we do
    NOT map columns by header X. Instead each row is found by its (unique) Part
    Number value, and the target cell is located by matching its *old* value's
    text on that row -- which also pins the column's real X, so multi-line
    (wrapped) cells are replaced as a unit. Returns patched bytes, or None."""
    try:
        records = _emf_text_records(emf)
        if not records:
            return None
        head = _emf_bom_header(records)
        if head is None:
            return None
        header_fields, header_bottom, col_tol = head
        charw, default_w = _emf_char_widths(records)
        group_tol = max(8, col_tol * 0.5)

        data = [r for r in records if r["refy"] > header_bottom]
        if not data:
            return None
        # The left-most (Item) column: single-line, one record per row, so it
        # anchors each row. Its Y gaps give the row pitch, and its values (the
        # item numbers) identify rows even when Part Numbers repeat.
        left_x = min(r["refx"] for r in data)
        item_recs = sorted((r for r in data
                            if abs(r["refx"] - left_x) <= group_tol),
                           key=lambda r: r["refy"])
        left_ys = [r["refy"] for r in item_recs]
        gaps = [b - a for a, b in zip(left_ys, left_ys[1:]) if b - a > 0]
        window = (sorted(gaps)[len(gaps) // 2] * 0.6) if gaps else 16

        ops = []  # (start, end, new_bytes, record_delta)
        for ed in cells:
            old = (ed["old"] or "").strip()
            new = (ed["new"] or "").strip()
            field = ed["field"]
            # 1) Find the row by its Item value, falling back to the row's
            #    ordinal among data rows -- NOT by Part Number (not unique).
            anchor = None
            item = (ed.get("item") or "").strip()
            if item:
                anchor = next((r for r in item_recs
                               if r["text"].strip() == item), None)
            if anchor is None:
                idx = ed.get("row_index")
                if isinstance(idx, int) and 0 <= idx < len(item_recs):
                    anchor = item_recs[idx]
            if anchor is None:
                continue
            row_y = anchor["refy"]
            row_recs = [r for r in data if abs(r["refy"] - row_y) <= window]

            # 2) Locate the target column's real X.
            target_x = None
            if old:
                prefix = [r for r in row_recs if r["text"].strip()
                          and old.startswith(r["text"].strip())]
                if prefix and field in header_fields:
                    prefix.sort(
                        key=lambda r: abs(r["refx"] - header_fields[field]))
                if prefix:
                    target_x = prefix[0]["refx"]
            if target_x is None:
                target_x = header_fields.get(field)
            if target_x is None:
                continue

            # 3) All records of that cell (wrapped lines share the column X).
            col_recs = sorted(
                (r for r in row_recs if abs(r["refx"] - target_x) <= group_tol),
                key=lambda r: r["refy"])
            if col_recs:
                primary = col_recs[0]
                if new:
                    ops.append((primary["off"],
                                primary["off"] + primary["size"],
                                _build_emf_text_record(
                                    emf, primary, new, primary["refy"],
                                    charw, default_w), 0))
                else:
                    ops.append((primary["off"],
                                primary["off"] + primary["size"], b"", -1))
                for extra in col_recs[1:]:  # drop extra wrapped lines
                    ops.append((extra["off"], extra["off"] + extra["size"],
                                b"", -1))
            elif new:
                tmpl = dict(anchor)
                tmpl["refx"] = target_x
                tmpl["bounds"] = (target_x, anchor["bounds"][1], target_x,
                                  anchor["bounds"][3])
                new_rec = _build_emf_text_record(emf, tmpl, new, row_y,
                                                 charw, default_w)
                ins = anchor["off"] + anchor["size"]
                ops.append((ins, ins, new_rec, 1))
        if not ops:
            return None
        # Apply back-to-front. Guard against overlapping byte ranges (two edits
        # resolving to the same record would corrupt the metafile).
        out = bytearray(emf)
        byte_delta = record_delta = 0
        last_start = len(emf) + 1
        for start, end, new_bytes, rdelta in sorted(
                ops, key=lambda o: (o[0], o[1]), reverse=True):
            if end > last_start:
                continue  # overlaps an already-applied op; skip it
            out[start:end] = new_bytes
            byte_delta += len(new_bytes) - (end - start)
            record_delta += rdelta
            last_start = start
        _emf_set_counts(out, byte_delta, record_delta)
        return bytes(out)
    except (struct.error, ValueError, IndexError):
        return None


def _parts_emf_layout(emf: bytes):
    """Geometry of a parts-table EMF: (records, header_fields {field: x},
    item_recs [left-column anchors, top->bottom], pitch). None if not a BOM."""
    records = _emf_text_records(emf)
    if not records:
        return None
    head = _emf_bom_header(records)
    if head is None:
        return None
    header_fields, header_bottom, col_tol = head
    data = [r for r in records if r["refy"] > header_bottom]
    if not data:
        return None
    left_x = min(r["refx"] for r in data)
    group_tol = max(8, col_tol * 0.5)
    item_recs = sorted((r for r in data
                        if abs(r["refx"] - left_x) <= group_tol),
                       key=lambda r: r["refy"])
    if not item_recs:
        return None
    ys = [r["refy"] for r in item_recs]
    gaps = [b - a for a, b in zip(ys, ys[1:]) if b - a > 0]
    pitch = sorted(gaps)[len(gaps) // 2] if gaps else (
        item_recs[0]["refy"] - header_bottom)
    if pitch <= 0:
        return None
    return records, header_fields, header_bottom, item_recs, pitch, data, \
        group_tol


def _parts_emf_growth(emf: bytes, n_new: int) -> float:
    """new_height / old_height needed so ``n_new`` appended parts rows fit in
    the cached picture's bounds (1.0 if they already fit)."""
    try:
        lay = _parts_emf_layout(emf)
        if not lay or n_new <= 0:
            return 1.0
        _r, _hf, _hb, item_recs, pitch, _d, _gt = lay
        bounds = struct.unpack_from("<4i", emf, 8)
        old_h = bounds[3] - bounds[1]
        bottom = item_recs[-1]["refy"] + (n_new + 1) * pitch
        if bottom > bounds[3] and old_h > 0:
            return (bottom - bounds[1]) / old_h
        return 1.0
    except (struct.error, ValueError, IndexError):
        return 1.0


def append_parts_rows_to_emf(emf: bytes, rows_fields: list) -> Optional[bytes]:
    """Draw genuinely-new parts rows into the cached EMF so Visio shows them
    without the user activating the object. ``rows_fields`` is a list of
    {canonical_field: value} dicts (in final display order). Returns patched
    bytes, or None if it couldn't be patched."""
    rows_fields = [r for r in (rows_fields or []) if r]
    if not rows_fields:
        return None
    try:
        lay = _parts_emf_layout(emf)
        if not lay:
            return None
        records, header_fields, _hb, item_recs, pitch, data, group_tol = lay
        charw, default_w = _emf_char_widths(records)
        last_y = item_recs[-1]["refy"]
        last_recs = [r for r in data if abs(r["refy"] - last_y) <= pitch * 0.5]
        # A template record per column (the last row's cell nearest that X), so
        # new cells inherit the data font/style. Fall back to the item anchor.
        templates = {}
        for field, fx in header_fields.items():
            cand = sorted((r for r in last_recs),
                          key=lambda r: abs(r["refx"] - fx))
            if cand and abs(cand[0]["refx"] - fx) <= max(group_tol, 1):
                templates[field] = cand[0]
        additions = b""
        added = 0
        insert_at = max(r["off"] + r["size"] for r in records)
        for i, rowvals in enumerate(rows_fields, start=1):
            new_y = last_y + i * pitch
            for field, value in rowvals.items():
                value = (value or "").strip()
                if not value or field not in header_fields:
                    continue
                tmpl = dict(templates.get(field) or item_recs[-1])
                tmpl["refx"] = header_fields[field]
                bx = header_fields[field]
                tmpl["bounds"] = (bx, tmpl["bounds"][1], bx, tmpl["bounds"][3])
                additions += _build_emf_text_record(
                    emf, tmpl, value, new_y, charw, default_w)
                added += 1
        if not added:
            return None
        out = bytearray(emf[:insert_at] + additions + emf[insert_at:])
        _emf_set_counts(out, len(additions), added)
        # Grow bounds/frame in step so the rows show at the same scale.
        bounds = list(struct.unpack_from("<4i", out, 8))
        old_h = bounds[3] - bounds[1]
        bottom = last_y + (len(rows_fields) + 1) * pitch
        if bottom > bounds[3] and old_h > 0:
            ratio = (bottom - bounds[1]) / old_h
            bounds[3] = bottom
            struct.pack_into("<4i", out, 8, *bounds)
            frame = list(struct.unpack_from("<4i", out, 24))
            fh = frame[3] - frame[1]
            frame[3] = frame[1] + int(round(fh * ratio))
            struct.pack_into("<4i", out, 24, *frame)
        return bytes(out)
    except (struct.error, ValueError, IndexError):
        return None


def apply_bom_edits_to_embedded(xlsx_bytes: bytes, cells: list,
                                ole_end: Optional[int] = None) -> bytes:
    """Set the given cells in an embedded parts worksheet. Each edit: {ref,
    new, ...}. Returns the new .xlsx bytes (unchanged if it can't be parsed).

    If ``ole_end`` is given (a row number), the OLE display range's bottom row
    is set to exactly that -- so an add/remove that changes the row count is
    shown without a trailing blank row. Otherwise the range is only grown to
    cover the edited rows."""
    parsed = _read_embedded_sheet_rows(xlsx_bytes)
    if not parsed:
        return xlsx_bytes
    sheet_member, sheet_xml, _rows = parsed
    max_row = 0
    for ed in cells:
        sheet_xml = _set_xlsx_cell_inline(sheet_xml, ed["ref"], ed["new"])
        rm = re.search(r"\d+", ed["ref"])
        if rm:
            max_row = max(max_row, int(rm.group(0)))
    new_bytes = _xlsx_replace_member(xlsx_bytes, sheet_member, sheet_xml)
    # Make sure every edited row is inside the range the OLE object displays;
    # for add/remove, size the range exactly to the final last data row.
    if ole_end is not None:
        return _extend_olesize(new_bytes, ole_end, exact=True)
    return _extend_olesize(new_bytes, max_row)


# ---------------------------------------------------------------------------
# Add / remove parts: shift rows up (remove) or append (add), renumber the
# item/line-number column, and express the result as ordinary cell edits so the
# existing worksheet + cached-picture appliers can carry them out.
# ---------------------------------------------------------------------------

def _table_data_rownums(rows: dict, col_field: dict, header_row: int):
    """The filled data-row numbers of a parts/BOM table (rows with any value in
    a known column), in order."""
    cols = set(col_field)
    out = [rn for rn in rows
           if rn > header_row
           and any((rows[rn].get(c) or "").strip() for c in cols)]
    return sorted(out)


def _detect_seq_cols(rows: dict, data_rownums: list, cols: list):
    """Columns whose values are a consecutive integer run (1,2,3,...) -- the
    item/line-number sequences that must be renumbered. Returns {col: start}."""
    seq = {}
    for col in cols:
        vals = [(rows.get(rn, {}).get(col) or "").strip() for rn in data_rownums]
        nums = [v for v in vals if v]
        if len(nums) >= 2 and all(v.lstrip("-").isdigit() for v in nums):
            ints = [int(v) for v in nums]
            if ints == list(range(ints[0], ints[0] + len(ints))):
                seq[col] = ints[0]
    return seq


def _parts_final_end_row(rows: dict, col_field: dict, header_row: int,
                         remove_rownums: set, add_rows: list = None) -> int:
    """The bottom row a parts/BOM table will occupy after the given removals
    and additions are applied (the rows shift up / new rows append). Mirrors
    the destination-row layout in :func:`compute_parts_shift_edits`."""
    data = _table_data_rownums(rows, col_field, header_row)
    keep = [rn for rn in data if rn not in set(remove_rownums or ())]
    n = len(keep) + len(list(add_rows or ()))
    dest = list(data)
    last = data[-1] if data else header_row
    while len(dest) < n:
        last += 1
        dest.append(last)
    return dest[n - 1] if n else header_row


def compute_parts_shift_edits(rows: dict, col_field: dict, header_row: int,
                              remove_rownums: set, add_rows: list = None):
    """Compute the cell edits that remove ``remove_rownums`` and/or append
    ``add_rows`` (each a {field: value} dict) to a parts/BOM table: remaining
    rows shift up, new rows go on the end, and sequence columns are renumbered.

    Each edit: {ref, col, row, field, old, new, item, pn} -- 'item'/'pn' carry
    the row's CURRENT values so the cached-picture applier can find it."""
    remove_rownums = set(remove_rownums or ())
    add_rows = list(add_rows or ())
    cols = sorted(
        {c for rn in rows for c in rows[rn]} | set(col_field), key=_col_index)
    data_rownums = _table_data_rownums(rows, col_field, header_row)
    if not data_rownums and not add_rows:
        return []
    seq = _detect_seq_cols(rows, data_rownums, cols) or {}
    # If there's an Item field but it wasn't a clean run, still renumber it.
    item_col = next((c for c, f in col_field.items() if f == "Item"), None)
    if item_col and item_col not in seq:
        seq[item_col] = 1
    pn_col = next((c for c, f in col_field.items() if f == "Part Number"), None)

    keep = [rn for rn in data_rownums if rn not in remove_rownums]
    # Source values for each destination slot: kept rows, then the new rows.
    sources = [("row", rn) for rn in keep] + [("new", d) for d in add_rows]

    # Destination row numbers: reuse existing rows, then extend past the last.
    last = data_rownums[-1] if data_rownums else header_row
    dest_rows = list(data_rownums)
    while len(dest_rows) < len(sources):
        last += 1
        dest_rows.append(last)

    edits = []
    for i, dest_rn in enumerate(dest_rows):
        old_item = (rows.get(dest_rn, {}).get(item_col) or "") if item_col else ""
        old_pn = (rows.get(dest_rn, {}).get(pn_col) or "") if pn_col else ""
        if i < len(sources):
            kind, src = sources[i]
            for col in cols:
                if col in seq:
                    newv = str(seq[col] + i)
                elif kind == "row":
                    newv = (rows.get(src, {}).get(col) or "")
                else:  # new row keyed by column letter (or canonical field)
                    newv = src.get(col)
                    if newv is None:
                        newv = src.get(col_field.get(col, ""), "")
                oldv = (rows.get(dest_rn, {}).get(col) or "")
                if str(newv) != str(oldv):
                    edits.append({
                        "ref": f"{col}{dest_rn}", "col": col, "row": dest_rn,
                        "field": col_field.get(col, ""), "old": str(oldv),
                        "new": str(newv), "item": str(old_item),
                        "pn": str(old_pn)})
        else:  # trailing rows left over after a removal: clear them
            for col in cols:
                oldv = (rows.get(dest_rn, {}).get(col) or "")
                if str(oldv) != "":
                    edits.append({
                        "ref": f"{col}{dest_rn}", "col": col, "row": dest_rn,
                        "field": col_field.get(col, ""), "old": str(oldv),
                        "new": "", "item": str(old_item), "pn": str(old_pn)})
    return edits


def _parts_add_field_rows(rows: dict, col_field: dict, header_row: int,
                          remove_rownums: set, add_rows: list):
    """The genuinely-new rows expressed as {canonical_field: final_value} dicts
    (with sequence columns already renumbered) -- for drawing into the EMF."""
    add_rows = list(add_rows or ())
    if not add_rows:
        return []
    cols = sorted(
        {c for rn in rows for c in rows[rn]} | set(col_field), key=_col_index)
    data_rownums = _table_data_rownums(rows, col_field, header_row)
    keep = [rn for rn in data_rownums if rn not in set(remove_rownums or ())]
    seq = _detect_seq_cols(rows, data_rownums, cols) or {}
    item_col = next((c for c, f in col_field.items() if f == "Item"), None)
    if item_col and item_col not in seq:
        seq[item_col] = 1
    out = []
    for j, src in enumerate(add_rows):
        fr = {}
        for col in cols:
            field = col_field.get(col, "")
            if col in seq:
                val = str(seq[col] + len(keep) + j)
            else:
                val = src.get(col)
                if val is None:
                    val = src.get(field, "")
            if field and val:
                fr[field] = str(val)
        out.append(fr)
    return out


def build_visio_parts_ops(path, removals: dict, additions: dict,
                          case_sensitive=False) -> dict:
    """Turn staged removals/additions into per-embedding cell edits for a .vsdx.

    removals:  {embed_member: set(rownums)}
    additions: {embed_member: [ {field: value}, ... ]}
    Returns {embed_member: {"emf", "cells", "ole_end", "add_rows"}}."""
    out: dict = {}
    if not removals and not additions:
        return out
    try:
        zin = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return out
    with zin:
        members = set(removals or ()) | set(additions or ())
        for member in members:
            try:
                info = _embedded_bom_info(zin.read(member))
            except (KeyError, OSError):
                info = None
            if not info:
                continue
            rm = set((removals or {}).get(member, ()))
            ad = list((additions or {}).get(member, ()))
            edits = compute_parts_shift_edits(
                info["rows"], info["col_field"], info["header_row"], rm, ad)
            if edits:
                out[member] = {
                    "emf": _embedded_emf_member(zin, member),
                    "cells": edits,
                    "ole_end": _parts_final_end_row(
                        info["rows"], info["col_field"],
                        info["header_row"], rm, ad),
                    "add_rows": _parts_add_field_rows(
                        info["rows"], info["col_field"],
                        info["header_row"], rm, ad)}
    return out


def visio_parts_sheets(path):
    """[(embed_member, sheet_name)] for every parts table in a .vsdx."""
    out = []
    try:
        with zipfile.ZipFile(path) as z:
            names = _embedding_page_names(z)
            for m in _embedded_bom_members(z):
                out.append((m, names.get(m, "")))
    except (zipfile.BadZipFile, OSError):
        return []
    out.sort(key=lambda t: _natural_key(t[1]))
    return out


def excel_bom_sheets(path):
    """[(worksheet_part, sheet_name)] for every BOM table in an .xlsx."""
    out = []
    try:
        with zipfile.ZipFile(path) as z:
            shared = _read_shared_strings(z)
            protected = _changelog_part(z)
            names = _xlsx_sheet_display(z)
            for n in z.namelist():
                if not _WORKSHEET_RE.search(n.lower()) or n == protected:
                    continue
                if _excel_bom_table(z, n, shared):
                    out.append((n, names.get(n, "")))
    except (zipfile.BadZipFile, OSError):
        return []
    return out


def parts_table_view(path, fmt, sheet_id):
    """For displaying a table in the add editor: (col_order [col letters],
    headers {col: header text}, rows [{col: value}, ...]) or None."""
    try:
        zin = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return None
    with zin:
        if fmt == "vsdx":
            info = _embedded_bom_info(zin.read(sheet_id))
            if not info:
                return None
            rows_map, col_field, hrow = (info["rows"], info["col_field"],
                                         info["header_row"])
        else:
            shared = _read_shared_strings(zin)
            tbl = _excel_bom_table(zin, sheet_id, shared)
            if not tbl:
                return None
            rows_map, col_field, hrow = tbl
    col_order = sorted(col_field, key=_col_index)
    headers = {c: (rows_map.get(hrow, {}).get(c) or col_field[c])
               for c in col_order}
    data_rownums = _table_data_rownums(rows_map, col_field, hrow)
    rows = [{c: (rows_map.get(rn, {}).get(c) or "") for c in col_order}
            for rn in data_rownums]
    return col_order, headers, rows


def _excel_bom_table(zin: zipfile.ZipFile, sheet_part: str, shared):
    """For an Excel BOM worksheet: (rows {rownum:{col_letter:value}},
    col_field {col_letter: field}, header_row) or None."""
    try:
        sheet_xml = zin.read(sheet_part).decode("utf-8", "replace")
    except (KeyError, OSError):
        return None
    cells = _read_sheet_cells(sheet_xml, shared)
    hrow, colmap = _find_bom_header(cells)
    if hrow is None:
        return None
    rows: dict = {}
    for (colnum, row), (_ref, txt) in cells.items():
        rows.setdefault(row, {})[_col_letters(colnum)] = txt
    col_field = {_col_letters(c): f for c, f in colmap.items()}
    return rows, col_field, hrow


def build_excel_parts_ops(path, removals: dict, additions: dict) -> dict:
    """Turn staged removals/additions into per-sheet cell edits for an .xlsx.

    removals:  {sheet_part: set(rownums)}
    additions: {sheet_part: [ {col_letter|field: value}, ... ]}
    Returns {sheet_part: {cell_ref: new_value}} (the cell_edits format)."""
    out: dict = {}
    if not removals and not additions:
        return out
    try:
        zin = zipfile.ZipFile(path)
    except (zipfile.BadZipFile, OSError):
        return out
    with zin:
        shared = _read_shared_strings(zin)
        for sheet in set(removals or ()) | set(additions or ()):
            tbl = _excel_bom_table(zin, sheet, shared)
            if not tbl:
                continue
            rows, col_field, hrow = tbl
            edits = compute_parts_shift_edits(
                rows, col_field, hrow,
                set((removals or {}).get(sheet, ())),
                list((additions or {}).get(sheet, ())))
            if edits:
                out[sheet] = {e["ref"]: e["new"] for e in edits}
    return out


def _revtable_page_part(zin: zipfile.ZipFile) -> Optional[str]:
    """The page archive member that actually holds the revision table.

    Visio page part numbers (page1.xml, page2.xml, ...) follow creation order,
    NOT the display order -- so the cover/title page that carries the revision
    table is often not page1.xml. We scan the pages (lowest number first) and
    return the first one where a table is detected, instead of assuming page 1.
    """
    parts = [n for n in zin.namelist() if _PAGE_NAME_RE.search(n)]
    parts.sort(key=lambda n: int(_PAGE_NAME_RE.search(n).group(1)))
    for n in parts:
        try:
            xml = zin.read(n).decode("utf-8", "replace")
        except KeyError:
            continue
        if _detect_revtable(xml) is not None:
            return n
    return None


def _revtable_for_path(path):
    """Detect the revision table on whichever page holds it. Returns the
    _detect_revtable dict, or None."""
    try:
        with zipfile.ZipFile(path) as z:
            page = _revtable_page_part(z)
            if not page:
                return None
            xml = z.read(page).decode("utf-8", "replace")
    except (zipfile.BadZipFile, OSError, KeyError):
        return None
    return _detect_revtable(xml)


def vsdx_revtable_columns(path) -> Optional[List[str]]:
    """The revision-table column names detected on a .vsdx (in canonical order),
    or None if no table is found. Used to preview what the 'add revision entry'
    feature will write to."""
    tbl = _revtable_for_path(path)
    if tbl:
        return [f for f in REVTABLE_FIELD_ORDER if f in tbl["col_x"]]
    # Fall back to a revision table stored as an embedded Excel OLE object.
    try:
        with zipfile.ZipFile(path) as z:
            m = _embedded_revtable_member(z)
            if m:
                return embedded_revtable_columns(z.read(m))
    except (zipfile.BadZipFile, OSError, KeyError):
        pass
    return None


def vsdx_revtable_rev_letters(path) -> List[str]:
    """The REV letters present in the revision table (e.g. ['A','B','C','D']),
    for the approval dialog. [] if no table/letters."""
    tbl = _revtable_for_path(path)
    if tbl:
        out = []
        for row in tbl["rows"]:
            rc = row["cells"].get("Rev")
            if rc and rc["text"].strip():
                out.append(rc["text"].strip())
        return out
    try:
        with zipfile.ZipFile(path) as z:
            m = _embedded_revtable_member(z)
            if m:
                return embedded_revtable_rev_letters(z.read(m))
    except (zipfile.BadZipFile, OSError, KeyError):
        pass
    return []


def vsdx_latest_revision_letter(path) -> str:
    """The most recently added revision letter in a file's revision table --
    the last filled row -- used to auto-target the approval. '' if none."""
    letters = vsdx_revtable_rev_letters(path)
    return letters[-1] if letters else ""


def replace_text_in_vsdx(
    in_path: str | os.PathLike,
    out_path: str | os.PathLike,
    pairs: Sequence[Tuple[str, str]],
    case_sensitive: bool = True,
    whole_word: bool = False,
    revision: Optional[Tuple[str, str]] = None,
    update_drawing_rev: bool = False,
    rev_entry: Optional[dict] = None,
    approval: Optional[dict] = None,
    bom_cell_edits: Optional[dict] = None,
) -> dict:
    """Copy a .vsdx applying text replacements; return a report dict.

    If ``revision`` is (old_letter, new_letter) and ``update_drawing_rev`` is
    true, the single-letter revision box is bumped to new_letter on *every*
    page that has one (the title block repeats on each sheet). If ``rev_entry``
    is given, a new row is added to the cover page's revision table. If
    ``approval`` is {'rev': letter, 'name': name}, that name is written into the
    Approved By cell of the matching revision row.

    Report: {"total", "by_part", "rev_drawing", "rev_sheets", "rev_table",
    "approval"}. rev_drawing is one of 'na'/'updated'/'not_found'/'ambiguous';
    rev_sheets is the number of pages whose REV box was bumped; rev_table is one
    of 'na'/'filled'/'appended'/'not_found'/'no_slot'; approval is one of
    'na'/'approved'/'row_not_found'/'no_column'/'not_found'.
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
    rev_sheets = 0
    table_status = "na"
    approval_status = "na"
    do_rev = bool(revision and update_drawing_rev)
    do_table = bool(rev_entry and any(
        v and v.strip() for v in rev_entry.values()))
    do_approval = bool(approval and approval.get("rev") and
                       approval.get("name"))

    bom_cell_edits = bom_cell_edits or {}
    bom_cells = 0
    # Reverse map: cached-EMF member -> the op (cells + any new rows to draw).
    bom_emf_map = {d["emf"]: d for d in bom_cell_edits.values()
                   if d.get("emf")}

    out_path.parent.mkdir(parents=True, exist_ok=True)
    page_count = 0
    table_page = None
    table_embed = None
    table_embed_emf = None
    rev_picture = "na"  # na/native/patched/patch_failed/unsupported/no_picture
    with zipfile.ZipFile(in_path, "r") as zin:
        page_count = sum(1 for n in zin.namelist()
                         if _PAGE_NAME_RE.search(n))
        # The revision table / approval go on whichever page actually has the
        # table (often not page1.xml -- part numbers follow creation order).
        if do_table or do_approval:
            table_page = _revtable_page_part(zin)
            if table_page:
                rev_picture = "native"
            # Some drawings keep the revision history as an embedded Excel OLE
            # object instead of native Visio text cells -- handle that too, and
            # patch its cached presentation (EMF) so the change shows without
            # the user having to open the object in Visio.
            else:
                table_embed = _embedded_revtable_member(zin)
                if table_embed:
                    table_embed_emf = _embedded_revtable_emf_member(
                        zin, table_embed)
                    if not table_embed_emf:
                        # There's a presentation, but not an EMF we can edit.
                        pres, ext = _embedded_presentation(zin, table_embed)
                        rev_picture = ("unsupported" if pres else "no_picture")
        # If adding the row pushes it below the OLE object's visible window,
        # grow that object's shape so Visio shows the new row in the page view
        # (otherwise it's only visible after opening the object).
        ole_grow_ratio = 1.0
        ole_grow_page = None
        ole_grow_shape = None
        if do_table and table_embed_emf:
            ole_grow_ratio = _emf_revrow_growth(zin.read(table_embed_emf))
            if ole_grow_ratio > 1.001:
                ole_grow_page, ole_grow_shape = _ole_shape_for_embedding(
                    zin, table_embed)
        # Parts tables that gained rows must also grow their OLE shape so the
        # new rows show in the page view. {page_part: [(shape_id, ratio), ...]}.
        parts_grow: dict = {}
        for embed, op in bom_cell_edits.items():
            if not op.get("add_rows") or not op.get("emf"):
                continue
            try:
                ratio = _parts_emf_growth(zin.read(op["emf"]),
                                          len(op["add_rows"]))
            except (KeyError, OSError):
                ratio = 1.0
            if ratio > 1.001:
                pg, sh = _ole_shape_for_embedding(zin, embed)
                if pg and sh:
                    parts_grow.setdefault(pg, []).append((sh, ratio))
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
                        # The REV box sits in the title block on every sheet,
                        # so bump it on each page, not just the cover.
                        if do_rev and _PAGE_NAME_RE.search(item.filename):
                            new_text, st = bump_revision_in_page(
                                new_text, revision[0], revision[1]
                            )
                            if st == "updated":
                                rev_sheets += 1
                                rev_status = "updated"
                            elif rev_status == "na":
                                rev_status = st
                        if do_table and item.filename == table_page:
                            new_text, table_status = add_revision_entry_to_page(
                                new_text, rev_entry
                            )
                        if do_approval and item.filename == table_page:
                            new_text, approval_status = add_approval_to_page(
                                new_text, approval["rev"], approval["name"]
                            )
                        # Grow the OLE object's shape if the new revision row
                        # falls below its visible window.
                        if (ole_grow_shape and ole_grow_page
                                and item.filename == ole_grow_page):
                            new_text = grow_ole_shape_height(
                                new_text, ole_grow_shape, ole_grow_ratio)
                        # Grow any parts-table OLE shapes that gained rows.
                        for sh, ratio in parts_grow.get(item.filename, ()):
                            new_text = grow_ole_shape_height(
                                new_text, sh, ratio)
                        if new_text != text:
                            data = new_text.encode("utf-8")
                elif table_embed and item.filename == table_embed:
                    # The revision table is an embedded Excel OLE object; edit
                    # its worksheet bytes directly.
                    if do_table:
                        data, table_status = add_revision_entry_to_embedded(
                            data, rev_entry)
                    if do_approval:
                        data, approval_status = add_approval_to_embedded(
                            data, approval["rev"], approval["name"])
                elif table_embed_emf and item.filename == table_embed_emf:
                    # Patch the cached EMF presentation to match the edit, so
                    # Visio shows the new row/approval without re-activating.
                    if do_table:
                        patched = add_revision_row_to_emf(data, rev_entry)
                        rev_picture = ("patched" if patched is not None
                                       else "patch_failed")
                        if patched is not None:
                            data = patched
                    if do_approval:
                        patched = approve_revision_in_emf(
                            data, approval["rev"], approval["name"])
                        if patched is not None:
                            data = patched
                            if do_table is False:
                                rev_picture = "patched"
                elif item.filename in bom_cell_edits:
                    # A parts-table (Cable BOM) edit: update the worksheet data.
                    cells = bom_cell_edits[item.filename]["cells"]
                    data = apply_bom_edits_to_embedded(
                        data, cells,
                        bom_cell_edits[item.filename].get("ole_end"))
                    bom_cells += len(cells)
                elif item.filename in bom_emf_map:
                    # Patch the parts table's cached EMF to match the edits,
                    # then draw any genuinely-new rows below the last one.
                    op = bom_emf_map[item.filename]
                    patched = apply_bom_edits_to_emf(data, op["cells"])
                    if patched is not None:
                        data = patched
                    if op.get("add_rows"):
                        grown = append_parts_rows_to_emf(data, op["add_rows"])
                        if grown is not None:
                            data = grown
                # Preserve the original name; recompress with deflate.
                zout.writestr(item.filename, data)

    if table_page:
        table_loc = Path(table_page).name
    elif table_embed:
        table_loc = f"embedded sheet {Path(table_embed).name}"
    else:
        table_loc = None
    return {"total": total, "by_part": by_part, "rev_drawing": rev_status,
            "rev_sheets": rev_sheets, "rev_table": table_status,
            "approval": approval_status, "page_count": page_count,
            "rev_table_page": table_loc, "bom_cells": bom_cells,
            "rev_picture": rev_picture}


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
# Matches a cell either self-closing (<c r=".." .../>) or with content
# (<c r="..">...</c>). group(3) is None for a self-closing (empty) cell.
_XLSX_CELL = re.compile(
    r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.DOTALL
)
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
    if content is None:  # self-closing (empty) cell
        return None
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

def _resolve_part(base: str, target: str) -> str:
    """Resolve an OPC relationship Target to an archive part name.

    Targets may be relative to ``base`` (the owning part's folder) or absolute
    from the package root (a leading '/', as some writers emit)."""
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base, target))


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
        part = _resolve_part("xl", target)
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
    return _resolve_part(base, m.group(1))


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


def _find_rev_cells(zin, names, shared, old_letter):
    """Find the revision-letter cell on EVERY worksheet that has one (the box
    next to a REV/Revision label in that sheet's title block).

    Returns (targets, status) where ``targets`` is a list of (part_name,
    cell_ref) -- one per sheet. A title block is repeated on each sheet, so the
    revision letter must be bumped on all of them, not just the first.
    """
    old_u = old_letter.upper()
    targets = []           # (part, ref) chosen via a REV label on that sheet
    no_label_cands = []    # (part, ref) candidates on sheets with no REV label
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
            # Pick the candidate closest to a REV label on THIS sheet.
            best = None  # (dist, ref)
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
                        best = (dist, ref)
            if best is not None:
                targets.append((name, best[1]))
        else:
            no_label_cands.extend((name, ref) for ref, _ in cands)

    if targets:
        return targets, "ok"
    # No REV label anywhere: only safe if there's exactly one stray candidate.
    if len(no_label_cands) == 1:
        return [no_label_cands[0]], "ok"
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
    targets = None
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
            targets, rev_status = _find_rev_cells(zin, names, shared,
                                                  revision[0])
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
                    if targets:
                        for tpart, tref in targets:
                            if tpart != name:
                                continue
                            new_text, changed = _set_cell_text(
                                new_text, tref, revision[1]
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

    Each match: {file, sheet, sheet_name, part, row, matched,
    fields:{field:(ref,value)}}, where ``sheet`` is the worksheet part and
    ``sheet_name`` is its display name (the tab label).
    """
    matches = []
    norm_targets = [(p, p if case_sensitive else p.lower()) for p in part_numbers
                    if p]
    with zipfile.ZipFile(path) as z:
        shared = _read_shared_strings(z)
        protected = _changelog_part(z)
        sheet_names = _xlsx_sheet_display(z)
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
                                "sheet_name": sheet_names.get(name, ""),
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
    "Item": {"item", "itemno", "itemnumber"},
    "ECN #": {"ecn", "eco", "ecnno", "econo", "ecnnumber", "econumber",
              "ecn#", "eco#"},
    "ERB Approval Date": {"erbapprovaldate", "approvaldate", "erbdate",
                          "dateapproved", "erbapproval"},
    "Change Description": {"changedescription", "changedesc"},
    "Change Author": {"changeauthor", "author", "changedby", "by",
                      "enteredby"},
}
CHANGELOG_FIELD_ORDER = ["Item", "ECN #", "ERB Approval Date",
                         "Change Description", "Change Author"]


def _canonical_changelog_field(text: str) -> Optional[str]:
    """Match a header to a Change Log field (real headers carry extra text,
    e.g. 'Change Description (Include Line #s)')."""
    n = _norm_header(text)
    if not n:
        return None
    for field, variants in _CHANGELOG_FIELDS.items():
        for v in variants:
            if n == v or (len(v) >= 5 and n.startswith(v)):
                return field
    return None


# A real ECN/ECO value (used to find the ECN column when it has no header).
_ECN_VALUE_RE = re.compile(r"(?i)^\s*(ECN|ECO)\b")


def changelog_info(path):
    """Inspect the Change Log sheet.

    Returns (part, header_row, {field: column}, next_row) or None. The ECN #
    column is found by header if present, else by its ECN/ECO data values (some
    sheets leave it unlabelled). ``next_row`` is the row after the last entry
    that has an ECN # value -- robust to Item numbers pre-filled past the data.
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

    # Headers can be split across the first couple of (often merged) rows.
    field_to_col: dict = {}
    header_row = 0
    for (col, row), (ref, txt) in cells.items():
        if row > 8 or not txt:
            continue
        cf = _canonical_changelog_field(txt)
        if cf and cf not in field_to_col:
            field_to_col[cf] = col
            header_row = max(header_row, row)

    if not any(f in field_to_col for f in
               ("Change Description", "Change Author", "ERB Approval Date")):
        return None

    ecn_col = field_to_col.get("ECN #")
    if ecn_col is None:
        counts: dict = {}
        for (col, row), (ref, txt) in cells.items():
            if row > header_row and txt and _ECN_VALUE_RE.match(str(txt)):
                counts[col] = counts.get(col, 0) + 1
        if counts:
            ecn_col = max(counts, key=counts.get)
        elif "Item" in field_to_col:
            ecn_col = field_to_col["Item"] + 1  # column right of Item
    if ecn_col is None:
        return None
    field_to_col["ECN #"] = ecn_col

    last = header_row
    for (col, row), (ref, txt) in cells.items():
        if col == ecn_col and row > header_row and txt and str(txt).strip():
            last = max(last, row)
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
# Excel "Author" box: replace the name beside it and stamp today's date
# ---------------------------------------------------------------------------
#
# A title block has a cell labelled "Author:" with the name in the next column
# and the date in the column after that. We replace the name with a designated
# one and set the date to the day the script runs, keeping the date's format.

def _excel_serial(d: datetime.date) -> int:
    """Excel 1900-system serial number for a date."""
    return (d - datetime.date(1899, 12, 30)).days


def _run_date_value(date_cell, run_date: datetime.date) -> str:
    """The value to write for the date, matching the existing cell's format.

    A numeric (serial) date stays a serial (the cell keeps its date style); a
    text date is re-rendered in the same textual pattern.
    """
    txt = date_cell[1] if date_cell else None
    if txt is not None and str(txt).strip():
        s = str(txt).strip()
        try:
            float(s)
            return str(_excel_serial(run_date))  # serial -> serial
        except ValueError:
            pass
        if re.match(r"^\d{4}-\d{1,2}-\d{1,2}$", s):
            return run_date.strftime("%Y-%m-%d")
        if re.match(r"^\d{1,2}/\d{1,2}/\d{4}$", s):
            return run_date.strftime("%m/%d/%Y")
        if re.match(r"^\d{1,2}/\d{1,2}/\d{2}$", s):
            return run_date.strftime("%m/%d/%y")
        if re.match(r"^\d{1,2}-\d{1,2}-\d{4}$", s):
            return run_date.strftime("%m-%d-%Y")
        return run_date.strftime("%Y-%m-%d")
    return str(_excel_serial(run_date))  # empty -> serial (keeps any date style)


def build_author_edits(path, new_name, run_date=None) -> dict:
    """Replace the name beside the 'Author' box and stamp the run date.

    Returns {worksheet_part: {name_ref: new_name, date_ref: date_value}} or {}.
    """
    if not new_name or not new_name.strip():
        return {}
    run_date = run_date or datetime.date.today()
    edits: dict = {}
    with zipfile.ZipFile(path) as z:
        protected = _changelog_part(z)
        shared = _read_shared_strings(z)
        for name in z.namelist():
            # Update the Author box on EVERY worksheet except the Change Log.
            if not _WORKSHEET_RE.search(name.lower()) or name == protected:
                continue
            cells = _read_sheet_cells(
                z.read(name).decode("utf-8", "replace"), shared
            )
            for (col, row), (ref, txt) in cells.items():
                if txt and _norm_header(txt) == "author":
                    name_ref = f"{_col_letters(col + 1)}{row}"
                    date_ref = f"{_col_letters(col + 2)}{row}"
                    date_val = _run_date_value(
                        cells.get((col + 2, row)), run_date
                    )
                    sheet = edits.setdefault(name, {})
                    sheet[name_ref] = new_name
                    sheet[date_ref] = date_val
    return edits


# Discipline approval boxes: a cell labelled with the discipline (EE / ME /
# Production) has the approver's name in the next cell and a date after that.
APPROVAL_DISCIPLINES = ["EE", "ME", "Production"]
_APPROVAL_LABELS = {
    "EE": {"ee", "eeapproval", "electrical", "electricalengineer",
           "electricalengineering", "elecengineer", "eeengineer"},
    "ME": {"me", "meapproval", "mechanical", "mechanicalengineer",
           "mechanicalengineering", "mecheng", "mechengineer"},
    "Production": {"production", "prod", "manufacturing", "mfg",
                   "productionapproval", "prodapproval", "mfgapproval"},
}


def build_approval_edits(path, discipline, new_name, run_date=None) -> dict:
    """Write an approver's name (and today's date) beside the EE/ME/Production
    label on every worksheet except the Change Log.

    The label cell's text must match the chosen ``discipline``; the name goes in
    the next column and the date in the one after, matching the existing date's
    format. Returns {worksheet_part: {name_ref: name, date_ref: date}} or {}.
    """
    if not new_name or not new_name.strip() or not discipline:
        return {}
    variants = _APPROVAL_LABELS.get(discipline)
    if not variants:
        return {}
    run_date = run_date or datetime.date.today()
    edits: dict = {}
    with zipfile.ZipFile(path) as z:
        protected = _changelog_part(z)
        shared = _read_shared_strings(z)
        for name in z.namelist():
            if not _WORKSHEET_RE.search(name.lower()) or name == protected:
                continue
            cells = _read_sheet_cells(
                z.read(name).decode("utf-8", "replace"), shared
            )
            for (col, row), (ref, txt) in cells.items():
                if txt and _norm_header(txt) in variants:
                    name_ref = f"{_col_letters(col + 1)}{row}"
                    date_ref = f"{_col_letters(col + 2)}{row}"
                    date_val = _run_date_value(
                        cells.get((col + 2, row)), run_date
                    )
                    sheet = edits.setdefault(name, {})
                    sheet[name_ref] = new_name
                    sheet[date_ref] = date_val
    return edits


def xlsx_approval_disciplines(path) -> List[str]:
    """Which of EE/ME/Production have a label cell present in the workbook
    (outside the Change Log). For the approval dialog."""
    found = []
    for disc in APPROVAL_DISCIPLINES:
        if build_approval_edits(path, disc, "x"):
            found.append(disc)
    return found


# ---------------------------------------------------------------------------
# File category (by file name) + format detection + dispatch
# ---------------------------------------------------------------------------

# Document type is read from the file name: DOCxxxxx = BOM, DWGxxxxx = System
# Drawing, CBLxxxxx = Cable Drawing.
_FILE_CATEGORIES = [
    ("BOM", re.compile(r"DOC\d", re.IGNORECASE)),
    ("System Drawing", re.compile(r"DWG\d", re.IGNORECASE)),
    ("Cable Drawing", re.compile(r"CBL\d", re.IGNORECASE)),
]
CATEGORY_ORDER = ["BOM", "System Drawing", "Cable Drawing", "Other"]


def file_category(path: str | os.PathLike) -> str:
    """Classify a file by its name: 'BOM', 'System Drawing', 'Cable Drawing',
    or 'Other'."""
    name = Path(path).name
    for cat, pat in _FILE_CATEGORIES:
        if pat.search(name):
            return cat
    return "Other"


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
    cell_edits=None, rev_entry=None, approval=None,
    bom_cell_edits=None,
) -> dict:
    """Dispatch to the Visio or Excel engine based on the file type."""
    fmt = detect_format(in_path)
    if fmt == "vsdx":
        return replace_text_in_vsdx(
            in_path, out_path, pairs, case_sensitive, whole_word,
            revision=revision, update_drawing_rev=update_drawing_rev,
            rev_entry=rev_entry, approval=approval,
            bom_cell_edits=bom_cell_edits,
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


# ---------------------------------------------------------------------------
# Change summary: diff the original vs the edited copy and report before/after
# ---------------------------------------------------------------------------

def _xlsx_sheet_display(zin: zipfile.ZipFile) -> dict:
    """{worksheet_part: display sheet name}."""
    try:
        wb = zin.read("xl/workbook.xml").decode("utf-8", "replace")
        rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
    except KeyError:
        return {}
    rid_target = {}
    for rel in re.finditer(r"<Relationship\b[^>]*>", rels):
        idm = re.search(r'Id="([^"]+)"', rel.group(0))
        tm = re.search(r'Target="([^"]+)"', rel.group(0))
        if idm and tm:
            rid_target[idm.group(1)] = tm.group(1)
    out = {}
    for sh in re.finditer(r"<sheet\b[^>]*>", wb):
        nm = re.search(r'name="([^"]+)"', sh.group(0))
        rm = re.search(r'r:id="([^"]+)"', sh.group(0))
        if nm and rm and rm.group(1) in rid_target:
            part = _resolve_part("xl", rid_target[rm.group(1)])
            out[part] = _xml_unescape(nm.group(1))
    return out


def _label_row(cells) -> dict:
    """Best-effort {col: header text} from the most header-like top row."""
    rows: dict = {}
    for (col, row), (ref, txt) in cells.items():
        rows.setdefault(row, {})[col] = txt
    best_row, best_score = None, 0
    for row, colvals in rows.items():
        if row > 10:
            continue
        score = sum(
            1 for txt in colvals.values()
            if txt and (_canonical_field(txt) or _canonical_changelog_field(txt))
        )
        if score > best_score:
            best_score, best_row = score, row
    if best_row is None or best_score < 1:
        return {}
    return {col: txt for col, txt in rows[best_row].items() if txt}


_BUILTIN_DATE_FMT = (set(range(14, 23)) | set(range(45, 48)) | {27, 28, 29,
                     30, 31, 32, 33, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56,
                     57, 58})


def _xlsx_date_styles(zin: zipfile.ZipFile) -> set:
    """Style indices (cellXfs) that render as a date."""
    try:
        styles = zin.read("xl/styles.xml").decode("utf-8", "replace")
    except KeyError:
        return set()
    date_fmt_ids = set(_BUILTIN_DATE_FMT)
    for m in re.finditer(r'<numFmt numFmtId="(\d+)" formatCode="([^"]*)"',
                         styles):
        core = re.sub(r'"[^"]*"', "", m.group(2).lower())
        if re.search(r"[dmy]", core) and not re.search(r"[#0]",
                                                       core.replace("e", "")):
            date_fmt_ids.add(int(m.group(1)))
    out = set()
    cx = re.search(r"<cellXfs[^>]*>(.*?)</cellXfs>", styles, re.DOTALL)
    if cx:
        for i, xf in enumerate(
            re.findall(r"<xf\b[^>]*?(?:/>|>.*?</xf>)", cx.group(1), re.DOTALL)
        ):
            fm = re.search(r'numFmtId="(\d+)"', xf)
            if fm and int(fm.group(1)) in date_fmt_ids:
                out.add(i)
    return out


def _sheet_cell_styles(sheet_xml: str) -> dict:
    """{cell_ref: style index} for cells that declare a style."""
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*?\bs="(\d+)"', sheet_xml):
        out[m.group(1)] = int(m.group(2))
    return out


def _serial_to_date(value: str) -> str:
    """Excel date serial -> 'YYYY-MM-DD' (passes other text through)."""
    try:
        n = float(value)
    except (TypeError, ValueError):
        return value
    try:
        d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(n))
        return d.strftime("%Y-%m-%d")
    except (OverflowError, ValueError):
        return value


def _cell_label(rows_out: dict, col: int, row: int, headers: dict) -> str:
    """Label a cell: the 'Key:' to its left for a key/value row (e.g. Author:,
    Revision:), else the column header for a table row."""
    rowcells = sorted(
        (c, txt) for c, txt in rows_out.get(row, {}).items()
        if txt and str(txt).strip()
    )
    if rowcells and str(rowcells[0][1]).rstrip().endswith(":"):
        label = str(rowcells[0][1]).rstrip().rstrip(":")
        for c, txt in rowcells:
            if c < col and str(txt).rstrip().endswith(":"):
                label = str(txt).rstrip().rstrip(":")
        return label.strip()
    return headers.get(col, "")


def _diff_xlsx(in_path, out_path) -> list:
    changes = []
    with zipfile.ZipFile(in_path) as zi, zipfile.ZipFile(out_path) as zo:
        shi, sho = _read_shared_strings(zi), _read_shared_strings(zo)
        names = _xlsx_sheet_display(zi)
        date_styles_i = _xlsx_date_styles(zi)
        date_styles_o = _xlsx_date_styles(zo)
        for part, sheet_name in names.items():
            try:
                xml_i = zi.read(part).decode("utf-8", "replace")
                xml_o = zo.read(part).decode("utf-8", "replace")
            except KeyError:
                continue
            ci = _read_sheet_cells(xml_i, shi)
            co = _read_sheet_cells(xml_o, sho)
            style_i = _sheet_cell_styles(xml_i)
            style_o = _sheet_cell_styles(xml_o)
            headers = _label_row(co or ci)
            rows_out: dict = {}
            for (col, row), (ref, txt) in co.items():
                rows_out.setdefault(row, {})[col] = txt
            for key in sorted(set(ci) | set(co), key=lambda k: (k[1], k[0])):
                bef = (ci.get(key) or (None, None))[1] or ""
                aft = (co.get(key) or (None, None))[1] or ""
                if str(bef) == str(aft):
                    continue
                ref = (co.get(key) or ci.get(key))[0]
                is_date = (style_o.get(ref) in date_styles_o
                           or style_i.get(ref) in date_styles_i)
                if is_date:
                    bef, aft = _serial_to_date(str(bef)), _serial_to_date(str(aft))
                changes.append({
                    "location": f"{sheet_name}!{ref}",
                    "field": _cell_label(rows_out, key[0], key[1], headers),
                    "before": str(bef), "after": str(aft),
                })
    return changes


def _visio_page_names(zin: zipfile.ZipFile) -> dict:
    """{page part: display page name}, best effort."""
    try:
        pages = zin.read("visio/pages/pages.xml").decode("utf-8", "replace")
        rels = zin.read(
            "visio/pages/_rels/pages.xml.rels"
        ).decode("utf-8", "replace")
    except KeyError:
        return {}
    rid_target = {}
    for rel in re.finditer(r"<Relationship\b[^>]*>", rels):
        idm = re.search(r'Id="([^"]+)"', rel.group(0))
        tm = re.search(r'Target="([^"]+)"', rel.group(0))
        if idm and tm:
            rid_target[idm.group(1)] = tm.group(1)
    out = {}
    for pm in re.finditer(r"<Page\b[^>]*>.*?</Page>|<Page\b[^>]*/>", pages,
                          re.DOTALL):
        block = pm.group(0)
        nm = re.search(r'\bName="([^"]+)"', block)
        rm = re.search(r'r:id="([^"]+)"', block)
        if nm and rm and rm.group(1) in rid_target:
            part = _resolve_part("visio/pages", rid_target[rm.group(1)])
            out[part] = _xml_unescape(nm.group(1))
    return out


def _visio_shape_texts(xml_text: str) -> dict:
    """{shape ID: visible text} for every shape that carries its own <Text>.

    Keying by shape ID (rather than by document position) lets the diff line up
    the same shape before/after even when other shapes were inserted -- e.g.
    when a revision row is added -- so it doesn't report phantom changes.
    """
    try:
        root = ET.fromstring(xml_text)
    except ET.ParseError:
        return {}
    ns = root.tag[: root.tag.index("}") + 1] if root.tag.startswith("{") else ""
    out = {}
    for sh in root.iter(ns + "Shape"):
        sid = sh.get("ID")
        text_el = sh.find(ns + "Text")
        if sid is None or text_el is None:
            continue
        out[sid] = "".join(text_el.itertext()).strip()
    return out


def _diff_vsdx_embeddings(zi: zipfile.ZipFile, zo: zipfile.ZipFile) -> list:
    """Diff the embedded Excel objects (the revision table and the parts tables)
    so their cell changes show in the summary -- they live in embedded
    worksheets, not in the page shapes the text diff looks at."""
    changes = []
    page_names = _embedding_page_names(zi)
    out_names = set(zo.namelist())
    for member in zi.namelist():
        if not _EMBED_XLSX_RE.search(member) or member not in out_names:
            continue
        try:
            bi, bo = zi.read(member), zo.read(member)
        except KeyError:
            continue
        if bi == bo:
            continue  # unchanged -- skip before parsing
        pi = _read_embedded_sheet_rows(bi)
        po = _read_embedded_sheet_rows(bo)
        if not pi or not po:
            continue
        rows_i, rows_o = pi[2], po[2]
        bom = _embedded_bom_info(bo)
        rev = None if bom else _embedded_revtable_info(bo)
        info = bom or rev
        col_field = info["col_field"] if info else {}
        hrow = info["header_row"] if info else 0
        if bom:
            location = page_names.get(member) or "Parts table"
        elif rev:
            location = "Revision table"
        else:
            location = page_names.get(member) or Path(member).stem
        cells = set()
        for rn, cols in rows_i.items():
            cells.update((c, rn) for c in cols)
        for rn, cols in rows_o.items():
            cells.update((c, rn) for c in cols)
        for col, rn in sorted(cells, key=lambda k: (k[1], k[0])):
            if rn <= hrow:
                continue  # header / title rows
            bef = (rows_i.get(rn) or {}).get(col) or ""
            aft = (rows_o.get(rn) or {}).get(col) or ""
            if str(bef) == str(aft):
                continue
            changes.append({
                "location": location, "field": col_field.get(col, ""),
                "before": str(bef), "after": str(aft),
            })
    return changes


def _diff_vsdx(in_path, out_path) -> list:
    changes = []
    with zipfile.ZipFile(in_path) as zi, zipfile.ZipFile(out_path) as zo:
        page_names = _visio_page_names(zi)
        parts = [n for n in zi.namelist()
                 if re.match(r"visio/pages/page\d+\.xml$", n)]
        for part in sorted(parts,
                           key=lambda n: int(re.search(r"(\d+)", n).group())):
            try:
                ti = _visio_shape_texts(zi.read(part).decode("utf-8", "replace"))
                to = _visio_shape_texts(zo.read(part).decode("utf-8", "replace"))
            except KeyError:
                continue
            label = page_names.get(part) or Path(part).stem
            # Edited or removed shapes (matched by ID), then newly added shapes.
            for sid, bef in ti.items():
                aft = to.get(sid, "")
                if bef != aft and (bef or aft):
                    changes.append({
                        "location": label, "field": "",
                        "before": bef, "after": aft,
                    })
            for sid, aft in to.items():
                if sid not in ti and aft:
                    changes.append({
                        "location": label, "field": "",
                        "before": "", "after": aft,
                    })
        # The revision table and parts tables are embedded Excel objects.
        changes.extend(_diff_vsdx_embeddings(zi, zo))
    return changes


def diff_files(in_path, out_path) -> list:
    """Return a list of {location, field, before, after} changes."""
    fmt = detect_format(in_path)
    try:
        if fmt == "xlsx":
            return _diff_xlsx(in_path, out_path)
        if fmt == "vsdx":
            return _diff_vsdx(in_path, out_path)
    except Exception:  # noqa: BLE001 - a summary should never block the run
        return []
    return []


def _natural_key(s: str):
    """Sort key that orders embedded numbers numerically, so sheet names like
    R0001, R0002, R0010 come out in order (not R0001, R0010, R0002)."""
    return [int(t) if t.isdigit() else t.lower()
            for t in re.split(r"(\d+)", s or "")]


def generate_change_summary(records, summary_path, run_dt=None) -> Path:
    """Write an HTML before/after change-review document.

    records: list of {input, output, revision, changes:[...]}.
    """
    import html as _html

    run_dt = run_dt or datetime.datetime.now()
    summary_path = Path(summary_path)
    total = sum(len(r["changes"]) for r in records)

    def render_file(r) -> str:
        rev = ""
        if r.get("revision"):
            rev = (f' &nbsp;|&nbsp; revision <b>{_html.escape(r["revision"][0])}'
                   f' &rarr; {_html.escape(r["revision"][1])}</b>')
        head = (f'<h3>{_html.escape(r["input"])}</h3>'
                f'<div class="sub">saved as <b>{_html.escape(r["output"])}</b>'
                f'{rev}</div>')
        changes = r["changes"]
        if not changes:
            return head + ('<p class="none">No content changes '
                           '(copied as-is).</p>')
        noun = "sheet" if any("!" in c["location"] for c in changes) else "page"

        # Collapse identical changes (same field + before + after) into one
        # row, collecting every place they happened, so a change repeated on
        # many sheets (e.g. the REV bump) is a single line, not one per sheet.
        groups: dict = {}
        order: list = []
        all_sheets: list = []
        for c in changes:
            loc = c["location"]
            sheet = loc.split("!", 1)[0] if "!" in loc else loc
            if sheet and sheet not in all_sheets:
                all_sheets.append(sheet)
            k = (c.get("field", "") or "", c["before"], c["after"])
            g = groups.get(k)
            if g is None:
                g = {"field": k[0], "before": c["before"],
                     "after": c["after"], "locs": []}
                groups[k] = g
                order.append(k)
            if loc not in g["locs"]:
                g["locs"].append(loc)

        def where(locs) -> str:
            if len(locs) == 1:
                return _html.escape(locs[0])
            sheets: list = []
            for loc in locs:
                s = loc.split("!", 1)[0] if "!" in loc else loc
                if s not in sheets:
                    sheets.append(s)
            sheets.sort(key=_natural_key)  # list pages in order
            if len(sheets) == 1:
                return (f'{_html.escape(sheets[0])} '
                        f'<span class="dim">({len(locs)} cells)</span>')
            if len(all_sheets) > 1 and set(sheets) == set(all_sheets):
                return f'<b>all {len(sheets)} {noun}s</b>'
            shown = ", ".join(_html.escape(s) for s in sheets[:12])
            if len(sheets) > 12:
                shown += f' <span class="dim">+{len(sheets) - 12} more</span>'
            return f'{len(sheets)} {noun}s: {shown}'

        trs = []
        for k in order:
            g = groups[k]
            loc_html = ""
            if g["field"]:
                loc_html += (f'<span class="field">'
                             f'{_html.escape(g["field"])}</span><br>')
            loc_html += f'<span class="where">{where(g["locs"])}</span>'
            trs.append(
                f'<tr><td class="loc">{loc_html}</td>'
                f'<td class="before">{_html.escape(g["before"]) or "&nbsp;"}</td>'
                f'<td class="after">{_html.escape(g["after"]) or "&nbsp;"}</td>'
                f'</tr>'
            )
        return (head + '<table><thead><tr><th>Change (where)</th>'
                '<th>Before</th><th>After</th></tr></thead><tbody>'
                + "".join(trs) + '</tbody></table>')

    # Group files by document type (BOM / System Drawing / Cable Drawing).
    by_cat: dict = {}
    for r in records:
        by_cat.setdefault(file_category(r["input"]), []).append(r)

    rows_html = []
    for cat in CATEGORY_ORDER:
        recs = by_cat.get(cat)
        if not recs:
            continue
        nchg = sum(len(r["changes"]) for r in recs)
        rows_html.append(
            f'<h2>{_html.escape(cat)} '
            f'<span class="count">({len(recs)} file(s), '
            f'{nchg} change(s))</span></h2>'
        )
        rows_html.extend(render_file(r) for r in recs)

    doc = f"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<title>Change Summary</title>
<style>
 body {{ font-family: Segoe UI, Arial, sans-serif; margin: 28px; color:#1a1a1a; }}
 h1 {{ margin-bottom: 4px; }}
 .meta {{ color:#555; margin-bottom: 22px; }}
 h2 {{ margin: 30px 0 6px; padding: 6px 10px; background:#243b53;
       color:#fff; border-radius:4px; }}
 h2 .count {{ font-weight: normal; font-size: 0.8em; opacity: 0.85; }}
 h3 {{ margin: 18px 0 2px; padding-top: 8px; border-top: 1px solid #ddd; }}
 .sub {{ color:#555; margin-bottom: 8px; font-size: 0.95em; }}
 table {{ border-collapse: collapse; width: 100%; margin: 8px 0 4px; }}
 th, td {{ border: 1px solid #ccc; padding: 6px 9px; text-align: left;
          vertical-align: top; font-size: 0.92em; }}
 th {{ background:#f3f3f3; }}
 td.loc {{ font-family: Segoe UI, Arial; max-width: 40%; }}
 .field {{ color:#0066aa; font-weight:600; font-size:0.92em; }}
 .where {{ color:#333; font-size:0.9em; }}
 .dim {{ color:#888; }}
 td.before {{ background:#fff3f3; }}
 td.after {{ background:#f1fbf1; }}
 .none {{ color:#777; font-style: italic; }}
 @media print {{ h2 {{ page-break-before: auto; }} }}
</style></head><body>
<h1>Change Summary</h1>
<div class="meta">Generated {_html.escape(run_dt.strftime("%Y-%m-%d %H:%M"))}
 &nbsp;|&nbsp; {len(records)} file(s) &nbsp;|&nbsp; {total} change(s).
 Before/after values for approver review.</div>
{''.join(rows_html)}
</body></html>"""

    summary_path.write_text(doc, encoding="utf-8")
    return summary_path


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
# Job aid / help content (rendered in-app and as printable HTML)
# ---------------------------------------------------------------------------
#
# Each section: (heading, [lines]). A line starting with "* " is a bullet;
# **bold** marks emphasis.

HELP_TITLE = "Drawing & BOM Studio — How to Use"

HELP_SECTIONS = [
    ("What this tool does", [
        "**Drawing & BOM Studio** edits Microsoft **Visio drawings (.vsdx)** "
        "and **Excel workbooks (.xlsx)** in bulk: it finds and replaces text "
        "across many files at once, saves each result as a new "
        "**next-revision** copy (your originals are never changed), and writes "
        "a **change summary** you can hand to an approver.",
        "The window has three tabs that all run together: **Find → Replace** "
        "(text rules), **Parts** (add / remove / edit rows in the Visio parts "
        "tables and Excel BOMs), and **Approve & Revise** (revision rows, "
        "approvals, Change Log, Author).",
        "It also has per-format helpers: edit a part's whole Excel BOM row, "
        "append a Change Log entry, stamp the Author box, add a row to a Visio "
        "revision table, and **approve** drawings/workbooks in bulk.",
        "* **Approvals** let one approver sign off **every** loaded file at "
        "once — by REV letter in Visio drawings, or by discipline (EE / ME / "
        "Production) in Excel workbooks — without opening each file by hand.",
    ]),
    ("Before you start", [
        "* Your originals are never modified — every result is a separate copy.",
        "* PDF export needs **LibreOffice** installed (optional; off by "
        "default). For correct Visio 'Sheet X of Y' numbers, export the PDF "
        "from Visio instead.",
        "* Only the modern **.vsdx** and **.xlsx** formats are supported. "
        "Re-save old .vsd / .xls files in the new format first.",
    ]),
    ("Document types (from the file name)", [
        "Files are classified by their name so rules can target a whole type:",
        "* **BOM** — name contains DOCxxxxx (e.g. DOC00475)",
        "* **System Drawing** — name contains DWGxxxxx",
        "* **Cable Drawing** — name contains CBLxxxxx",
        "Anything else is **Other**.",
    ]),
    ("Step 1 — Add your files", [
        "* Pick **Single file**, or **Batch** to process several at once "
        "(you can mix Visio and Excel).",
        "* Use **Add files...** to pick several, or **Add folder...** to add "
        "every .vsdx/.xlsx in a folder.",
        "* **Remove selected** / **Clear** manage the list.",
    ]),
    ("Step 2 — Add find / replace rules", [
        "* Each rule has a **Find:** box and a **Replace with:** box. Click "
        "**+ Add another rule** for more.",
        "* The **in:** dropdown chooses which document type(s) the rule "
        "applies to — tick **All files**, or any of BOM / System Drawing / "
        "Cable Drawing.",
        "* Matching is case-insensitive unless you tick **Case sensitive**, "
        "and it finds text even when Visio/Excel split it into pieces.",
    ]),
    ("Step 2 (Excel only) — Edit BOM rows", [
        "Put a part number in a **Find** box, then click **Excel: find & edit "
        "rows...**.",
        "* It finds the **P/N** / **Part Number** column and lists every "
        "matching row **grouped by file**, showing the **sheet name** and row "
        "it was found on.",
        "* Edit **Manufacturer, Unit Cost, Description, Qty, Notes** for each "
        "row — you can set a different value per file.",
        "* The **Show find value(s)** dropdown filters the list to one or more "
        "find values (handy when many rules produce a lot of matches).",
        "* The **Copy 1st down** buttons (Manufacturer / Unit Cost / "
        "Description / Qty) copy the first shown row's value into the rest of "
        "that **find value's** rows — fill all instances of a part from the "
        "first one.",
        "* **Reset fields** puts every field back to the originally found "
        "data. **Refresh lookup** re-scans; **Save edits** stages the changes. "
        "Numbers stay numeric.",
    ]),
    ("Step 2 (Excel only) — Change Log entry", [
        "Click **Excel: add Change Log entry...** and fill in **Item, ECN #** "
        "(or ECO#), **ERB Approval Date, Change Description, Change Author**.",
        "* The row is appended to the **Change Log** sheet of every Excel "
        "file, at the next free row (found via the ECN # column).",
        "* The Change Log sheet is otherwise **protected** — Find/Replace and "
        "BOM edits never touch it, for traceability.",
    ]),
    ("Step 2 (Excel only) — Author + date", [
        "Click **Excel: set Author + date...** and type a name. On every "
        "Excel sheet (except the Change Log), the name beside the **Author** "
        "box is replaced and the date beside it is set to **today** (kept in "
        "the same format).",
    ]),
    ("Step 2 (Excel only) — Approve (EE / ME / Production)", [
        "Click **Excel: approve (EE/ME/Prod)...** to sign off as an approver. "
        "Pick the **discipline** (EE, ME, or Production) and type your name.",
        "* The dialog shows which **approval boxes were found** in your files.",
        "* On run, your name goes in the cell **beside that discipline's "
        "label** and **today's date** in the next cell (kept in the existing "
        "date's format), on **every sheet except the Change Log**, in every "
        "loaded Excel file.",
    ]),
    ("Step 2 (Visio only) — Edit parts-table rows", [
        "Put a part number in a **Find** box, then click **Visio: find & edit "
        "rows...** to edit the parts list at the bottom of the R/C/P/W sheets "
        "(Item / Ref/DES # / Cable Name / Description / Part Number / "
        "Manufacturer / Qty/Length / Unit).",
        "* Rows are matched by **Part Number** and listed **per file and "
        "sheet** (e.g. Sheet: R0001), each field pre-filled — edit any of them.",
        "* Same helpers as the Excel editor: a **Show find value(s)** filter, "
        "**Copy 1st down** buttons (Manufacturer / Description / Qty/Length / "
        "Unit), **Reset fields**, and **Refresh lookup**. Opens in its own "
        "window.",
        "* The **Part Number** itself is changed by your normal **Find -> "
        "Replace** rule: a row whose part number matches the Find value has "
        "that cell set to the Replace value. A cell written **'<P/N> or "
        "equiv.'** still matches the bare part number, and the whole cell is "
        "replaced (put 'or equiv.' in the Replace box to keep it).",
        "* On run, each change is written to the embedded worksheet **and "
        "drawn into the table's cached picture**, so it shows when the drawing "
        "opens — no need to double-click the table in Visio.",
    ]),
    ("Parts tab — Remove parts", [
        "On the **Parts** tab, type a part number in the **Remove** box and "
        "click **Find & choose rows to remove...**. Every parts-table / BOM "
        "row with that part number is listed **per file and sheet**, each with "
        "a checkbox.",
        "* Tick the rows to delete and click **Save removals**. They are "
        "staged, not applied yet — the **Run** button does the work.",
        "* On run, each ticked row is removed, the rows below it **shift up**, "
        "and any **item / line-number** sequence is renumbered (1, 2, 3, …). "
        "Works on both Excel BOMs and the embedded Visio parts tables, "
        "including the cached picture and the table's display range.",
    ]),
    ("Parts tab — Add parts", [
        "On the **Parts** tab, click **Add parts...**. Pick a **file type** "
        "(BOM / Cable Drawing / System Drawing), then the **file** and the "
        "**sheet / table**. The current parts are shown for reference.",
        "* Fill in the blank row(s) at the bottom — one row per new part. Use "
        "**+ Add row** for more. Leave the item/line number blank: it is "
        "**assigned automatically** when you run.",
        "* Click **Save** to stage the new parts (you can add to several "
        "sheets / files before running). The **Run** button appends them to "
        "the worksheet, renumbers the sequence, grows the table's display "
        "range, and draws the new rows into the cached picture so they show "
        "when the drawing opens.",
        "* **Remove and Add run together**, alongside any Find → Replace rules "
        "and Approve & Revise actions — fill in whichever tabs you need and "
        "they are all applied in one run.",
    ]),
    ("Step 2 (Visio only) — Add a revision entry", [
        "Click **Visio: add revision entry...** to add a row to the "
        "**revision-history table** on a drawing's cover page (the chart in a "
        "corner with REV / DESCRIPTION / DATE / APPROVED columns).",
        "* The dialog shows the **columns it detected**. Fill in **ECN #, "
        "Description, Date, Approved By** (skip any) — the same for every file.",
        "* The **REV** column is filled **automatically per file** with that "
        "file's own next revision letter, so a whole batch each gets its "
        "correct next letter (you don't type it).",
        "* On run, **every** Visio file gets the new row: an existing **blank "
        "row** is filled, otherwise a new row is **cloned** below the last one, "
        "with **matching border lines** drawn around it.",
        "* If no table is confidently found on a file, that file is **left "
        "unchanged** (the Status box says so) — it never risks the drawing.",
    ]),
    ("Step 2 (Visio only) — Approve a revision", [
        "Click **Visio: approve revision...** to sign off a revision without "
        "opening the drawing. Just type your **name**.",
        "* By default it signs off **each file's own latest revision**, "
        "auto-detected per file (the dialog shows the latest letter found in "
        "each). Leave **REV letter** blank for that, or type a letter to force "
        "a specific one.",
        "* On run, your name goes into the **Approved** column of that revision "
        "row, in every loaded Visio file's revision table (native cells or an "
        "embedded Excel table).",
        "* An approval **does not bump the revision**; the copy is named "
        "**<name>_approved_<today's date>** so it's clearly a sign-off.",
    ]),
    ("Step 3 — Options", [
        "* **Case sensitive** / **Whole word only** — control matching.",
        "* **Also export PDF (LibreOffice)** — off by default.",
        "* **Save copy as next revision (REVx -> next)** — name each copy "
        "as the next letter (REVA -> REVB) and bump the REV box inside the "
        "file. On by default. (Ignored for a file you're approving — those are "
        "named *_approved_<date> and keep their revision.)",
        "* **Generate change summary** — write a before/after review document.",
        "* **Output folder** — click **Choose folder...** to send every "
        "finished file (and the change summary) to one folder; **Use source "
        "folder** puts each copy next to its original (the default).",
    ]),
    ("Run it", [
        "Click **Run**. Everything you staged on the **Find → Replace**, "
        "**Parts**, and **Approve & Revise** tabs is applied in one pass. Each "
        "file is saved as its next-revision copy (or *_edited, or "
        "*_approved_<date> for an approval) — next to the original, or in your "
        "chosen **output folder**. The **Status** box reports what happened per "
        "file, and the **change summary** opens when done.",
        "* **Reset all** (the orange button) clears the loaded files, every "
        "rule, all staged edits and the output folder — use it to start fresh "
        "on a new file or batch.",
    ]),
    ("The change summary", [
        "An HTML document grouped by **document type**, then **file**. For each "
        "file it lists changes as **Change (where) / Before / After** — text "
        "replacements, BOM edits, the Author name+date, the appended Change Log "
        "row, and the revision bump. Print it to PDF for sign-off.",
        "* **Identical changes are grouped into one line.** A change repeated "
        "across sheets/pages (like the REV bump) shows once, noting **all "
        "sheets** or the **specific sheets** it happened on — so the summary is "
        "quick to review instead of one line per sheet.",
    ]),
    ("Revisions", [
        "* The revision letter is read from the **file name** (e.g. REVA).",
        "* A->B->C ... major changes only. Already at REVZ? That file is "
        "skipped with a warning. No REVx in the name? The copy is named "
        "*_edited instead.",
        "* The matching letter **inside** the file (next to a REV / Revision "
        "label) is bumped too — on **every page** of a Visio drawing and "
        "**every worksheet** of an Excel workbook, since the title block "
        "repeats on each sheet.",
    ]),
    ("Appearance", [
        "* Use the **Dark / Light** button (top-right) to switch themes.",
    ]),
    ("Tips & troubleshooting", [
        "* Nothing replaced? Check spelling, the **in:** type, and **Case "
        "sensitive**.",
        "* 'Sheet 0 of N' in a PDF is a LibreOffice limitation — export the "
        "PDF from Visio for correct sheet numbers.",
        "* The title bar shows the version; keep it up to date.",
    ]),
]


def _help_lines_to_html(lines) -> str:
    import html as _html
    out, in_list = [], False

    def fmt(s):
        s = _html.escape(s)
        while "**" in s:
            s = s.replace("**", "<b>", 1)
            if "**" in s:
                s = s.replace("**", "</b>", 1)
            else:
                s += "</b>"
        return s

    for ln in lines:
        if ln.startswith("* "):
            if not in_list:
                out.append("<ul>")
                in_list = True
            out.append(f"<li>{fmt(ln[2:])}</li>")
        else:
            if in_list:
                out.append("</ul>")
                in_list = False
            out.append(f"<p>{fmt(ln)}</p>")
    if in_list:
        out.append("</ul>")
    return "".join(out)


def help_to_html(path) -> "Path":
    """Write the job aid as a printable HTML document."""
    import html as _html
    path = Path(path)
    body = [f"<h1>{_html.escape(HELP_TITLE)}</h1>"]
    for heading, lines in HELP_SECTIONS:
        body.append(f"<h2>{_html.escape(heading)}</h2>")
        body.append(_help_lines_to_html(lines))
    doc = f"""<!DOCTYPE html><html lang="en"><head><meta charset="utf-8">
<title>{_html.escape(HELP_TITLE)}</title><style>
 body {{ font-family: Segoe UI, Arial, sans-serif; max-width: 820px;
        margin: 30px auto; color:#1f2933; line-height:1.5; padding:0 16px; }}
 h1 {{ color:#1d4ed8; }}
 h2 {{ color:#243b53; margin-top:26px; border-bottom:2px solid #e2e8f0;
       padding-bottom:4px; }}
 ul {{ margin:6px 0 6px 4px; }} li {{ margin:3px 0; }}
 b {{ color:#111; }}
 @media print {{ h2 {{ page-break-inside: avoid; }} }}
</style></head><body>{''.join(body)}
<p style="color:#888;margin-top:30px;font-size:0.85em">Drawing &amp; BOM
 Studio v{__version__}</p></body></html>"""
    path.write_text(doc, encoding="utf-8")
    return path


# ---------------------------------------------------------------------------
# Tkinter GUI
# ---------------------------------------------------------------------------

def _enable_high_dpi() -> None:
    """Tell Windows this process renders at the screen's real DPI, so text and
    shapes aren't bitmap-stretched (which looks blurry). Must run before the
    first Tk window is created. A no-op off Windows."""
    if sys.platform != "win32":
        return
    try:
        import ctypes
    except Exception:  # noqa: BLE001
        return
    # Try newest -> oldest: per-monitor-v2 (Win10 1703+), then per-monitor,
    # then plain system-DPI aware. Any one of these stops the blurry stretch.
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(-4)  # PMv2
        return
    except Exception:  # noqa: BLE001
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)  # per-monitor
        return
    except Exception:  # noqa: BLE001
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()  # system DPI aware
    except Exception:  # noqa: BLE001
        pass


def launch_gui() -> int:
    # Imported lazily so the core logic / CLI work without a display.
    import threading
    import tkinter as tk
    import tkinter.font as tkfont
    from tkinter import filedialog, messagebox, scrolledtext, ttk

    # Crisp anti-aliased buttons when Pillow is present; plain canvas otherwise.
    try:
        from PIL import Image, ImageDraw, ImageTk
        _HAVE_PIL = True
    except Exception:  # noqa: BLE001
        _HAVE_PIL = False

    _enable_high_dpi()

    LIGHT = {
        "bg": "#eef1f6", "card": "#ffffff", "accent": "#2563eb",
        "accent_dk": "#1d4ed8", "text": "#1f2933", "muted": "#5b6b7b",
        "border": "#cbd5e1", "field": "#ffffff", "field_fg": "#1f2933",
        "btn": "#dbe3ee", "btn_fg": "#1f2933", "btn_hover": "#c7d3e3",
        "btn_off": "#cdd6e2", "btn_off_fg": "#8a97a8",
        "green": "#16a34a", "green_dk": "#15803d",
        "orange": "#ea580c", "orange_dk": "#c2410c",
        "banner": "#2563eb", "banner_sub": "#cfe0ff", "sel": "#2563eb",
    }
    DARK = {
        "bg": "#1f2430", "card": "#272d3a", "accent": "#3b82f6",
        "accent_dk": "#2563eb", "text": "#e6e9ef", "muted": "#9aa6b2",
        "border": "#3a4150", "field": "#2b3240", "field_fg": "#e6e9ef",
        "btn": "#3a4150", "btn_fg": "#e6e9ef", "btn_hover": "#49525f",
        "btn_off": "#333a47", "btn_off_fg": "#6b7585",
        "green": "#22c55e", "green_dk": "#16a34a",
        "orange": "#fb923c", "orange_dk": "#f97316",
        "banner": "#172554", "banner_sub": "#9db8ff", "sel": "#3b82f6",
    }

    class RoundedButton(tk.Canvas):
        """A flat button with rounded corners drawn on a canvas."""

        def __init__(self, parent, text="", command=None, kind="normal",
                     palette=LIGHT, radius=11, padx=16, pady=8, font=None,
                     bg_key="bg"):
            self.palette = palette
            self.kind = kind
            self.command = command
            self.radius = radius
            self.bg_key = bg_key
            self._text = text
            self._enabled = True
            self._hover = False
            self.font = font or tkfont.Font(
                family="Segoe UI", size=10,
                weight=("bold" if kind in ("accent", "green", "orange")
                        else "normal"),
            )
            w = self.font.measure(text) + padx * 2
            h = self.font.metrics("linespace") + pady * 2
            super().__init__(parent, width=w, height=h, highlightthickness=0,
                             bd=0, bg=palette[bg_key])
            self.bind("<Enter>", lambda e: self._set_hover(True))
            self.bind("<Leave>", lambda e: self._set_hover(False))
            self.bind("<Button-1>", self._click)
            self.bind("<Configure>", lambda e: self._draw())
            self._draw()

        def _fill_fg(self):
            p = self.palette
            if not self._enabled:
                return p["btn_off"], p["btn_off_fg"]
            if self.kind == "accent":
                return (p["accent_dk"] if self._hover else p["accent"]), "#ffffff"
            if self.kind == "green":
                return (p["green_dk"] if self._hover else p["green"]), "#ffffff"
            if self.kind == "orange":
                return (p["orange_dk"] if self._hover
                        else p["orange"]), "#ffffff"
            return (p["btn_hover"] if self._hover else p["btn"]), p["btn_fg"]

        def _draw(self):
            self.delete("all")
            w, h, r = int(self["width"]), int(self["height"]), self.radius
            fill, fg = self._fill_fg()
            bg = self.palette[self.bg_key]
            self.configure(bg=bg)
            if _HAVE_PIL and w > 1 and h > 1:
                self._draw_pil(w, h, r, fill, bg)
            else:
                # Fallback: a smoothed canvas polygon (no anti-aliasing, but at
                # the screen's real DPI it's acceptable once high-DPI is on).
                pts = [r, 0, w - r, 0, w, 0, w, r, w, h - r, w, h,
                       w - r, h, r, h, 0, h, 0, h - r, 0, r, 0, 0]
                self.create_polygon(pts, smooth=True, splinesteps=16,
                                     fill=fill, outline=fill)
            self.create_text(w // 2, h // 2 + 1, text=self._text, fill=fg,
                             font=self.font)

        def _draw_pil(self, w, h, r, fill, bg):
            """Render the rounded rectangle super-sampled, then downscale so the
            corners are smoothly anti-aliased."""
            ss = 4  # super-sample factor
            img = Image.new("RGB", (w * ss, h * ss), bg)
            draw = ImageDraw.Draw(img)
            draw.rounded_rectangle(
                (0, 0, w * ss - 1, h * ss - 1), radius=max(0, r) * ss,
                fill=fill)
            img = img.resize((w, h), Image.LANCZOS)
            self._img = ImageTk.PhotoImage(img)  # keep a ref (GC guard)
            self.create_image(0, 0, anchor="nw", image=self._img)

        def _set_hover(self, v):
            if self._enabled:
                self._hover = v
                self.configure(cursor="hand2" if v else "")
                self._draw()

        def _click(self, _e):
            if self._enabled and self.command:
                self.command()

        def set_enabled(self, v):
            self._enabled = bool(v)
            self._hover = False
            self._draw()

        def set_palette(self, palette):
            self.palette = palette
            self._draw()

        def configure_text(self, text):
            self._text = text
            self._draw()

    class App:
        def __init__(self, root: "tk.Tk"):
            self.root = root
            root.title(f"Drawing & BOM Studio   [v{__version__}]")
            # The window is sized in pixels, so grow it to match the display's
            # DPI (otherwise it's tiny and cramped on high-DPI screens now that
            # the app renders at the real resolution) -- but never larger than
            # the screen, so the window (and its scrollbar) always fit. The form
            # itself scrolls, so a small window is fine.
            try:
                scale = max(1.0, root.winfo_fpixels("1i") / 96.0)
            except Exception:  # noqa: BLE001
                scale = 1.0
            sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
            w = min(int(980 * scale), sw - 60)
            h = min(int(820 * scale), sh - 100)
            root.geometry(f"{w}x{h}")
            root.minsize(min(int(720 * scale), sw - 60),
                         min(int(420 * scale), sh - 100))
            self.dark = False
            self.palette = LIGHT
            self._rbuttons: list = []      # rounded buttons to re-theme
            self._setup_style()

            # Title banner (with a Light/Dark toggle)
            self._banner = tk.Frame(root, bg=self.palette["banner"])
            self._banner.pack(fill="x")
            self._banner_title = tk.Label(
                self._banner, text="Drawing & BOM Studio",
                bg=self.palette["banner"], fg="#ffffff",
                font=("Segoe UI", 15, "bold"),
            )
            self._banner_title.pack(side="left", padx=16, pady=10)
            self._banner_sub = tk.Label(
                self._banner, text=f"v{__version__}",
                bg=self.palette["banner"], fg=self.palette["banner_sub"],
                font=("Segoe UI", 9),
            )
            self._banner_sub.pack(side="left", pady=(14, 0))
            self.theme_btn = RoundedButton(
                self._banner, text="🌙  Dark", command=self.toggle_theme,
                kind="normal", palette=self.palette, radius=14,
                bg_key="banner", padx=12, pady=6,
            )
            self.theme_btn.pack(side="right", padx=(0, 14), pady=8)
            self._rbuttons.append(self.theme_btn)
            self.help_btn = RoundedButton(
                self._banner, text="❓  How to use", command=self.show_help,
                kind="accent", palette=self.palette, radius=14,
                bg_key="banner", padx=12, pady=6,
            )
            self.help_btn.pack(side="right", padx=(0, 8), pady=8)
            self._rbuttons.append(self.help_btn)

            self.files: List[str] = []
            # Each rule: {frame, find, repl, menubtn, menu, all_var, cat_vars}.
            # all_var True => applies to every file; otherwise cat_vars holds a
            # BooleanVar per document type (BOM/System Drawing/Cable Drawing).
            self.rule_rows: List[dict] = []
            # Staged Excel row edits: {part_number: {canonical_field: value}}.
            self.bom_edits: dict = {}
            # Staged Visio parts-table edits, computed per file/embedding:
            #   {file: {embed_member: {"emf": emf_member, "cells": [edit, ...]}}}
            self.visio_bom_edits: dict = {}
            # Staged part removals: {file: [{"embed"/"sheet", "rows":[..], ...}]}
            self.remove_parts: dict = {}
            # Staged part additions: {file: [{"embed"/"sheet", "rows":[{..}], ..}]}
            self.add_parts: dict = {}
            # Staged Change Log row to append: {canonical_field: value}.
            self.changelog_entry: dict = {}
            # New name for the "Author" box (date is stamped automatically).
            self.author_name: str = ""
            # Staged Visio revision-table row: {canonical_field: value}.
            self.visio_rev_entry: dict = {}
            # Staged Visio approval: {"rev": letter, "name": approver}.
            self.visio_approval: dict = {}
            # Staged Excel approval: {"discipline": EE/ME/Production, "name":..}.
            self.excel_approval: dict = {}
            # Optional folder for all finished files ("" = beside each source).
            self.out_dir: str = ""

            pad = {"padx": 10, "pady": 6}

            # --- scrollable body: the whole form lives in a canvas so it can be
            # scrolled when the window is shorter than the content (otherwise
            # the Run button at the bottom is unreachable on small screens).
            self._scroll_outer = tk.Frame(root, bg=self.palette["bg"])
            self._scroll_outer.pack(fill="both", expand=True)
            self._scroll_canvas = tk.Canvas(
                self._scroll_outer, highlightthickness=0, bd=0,
                bg=self.palette["bg"])
            _vbar = ttk.Scrollbar(
                self._scroll_outer, orient="vertical",
                command=self._scroll_canvas.yview)
            self._scroll_canvas.configure(yscrollcommand=_vbar.set)
            _vbar.pack(side="right", fill="y")
            self._scroll_canvas.pack(side="left", fill="both", expand=True)
            self.body = ttk.Frame(self._scroll_canvas)
            _body_id = self._scroll_canvas.create_window(
                (0, 0), window=self.body, anchor="nw")

            def _fit_scroll(_e=None):
                self._scroll_canvas.configure(
                    scrollregion=self._scroll_canvas.bbox("all"))

            def _fit_width(e):
                self._scroll_canvas.itemconfigure(_body_id, width=e.width)

            self.body.bind("<Configure>", _fit_scroll)
            self._scroll_canvas.bind("<Configure>", _fit_width)

            def _on_wheel(e):
                # Scroll the form; let the log box / file list keep their own
                # wheel, and ignore wheel events that belong to a dialog.
                try:
                    if e.widget.winfo_toplevel() is not self.root:
                        return
                    if e.widget.winfo_class() in ("Text", "Listbox",
                                                  "TCombobox"):
                        return
                except Exception:  # noqa: BLE001
                    return
                if getattr(e, "num", None) == 4:
                    n = -1
                elif getattr(e, "num", None) == 5:
                    n = 1
                else:
                    n = -1 if getattr(e, "delta", 0) > 0 else 1
                self._scroll_canvas.yview_scroll(n * 2, "units")

            root.bind_all("<MouseWheel>", _on_wheel)   # Windows / macOS
            root.bind_all("<Button-4>", _on_wheel)     # Linux wheel up
            root.bind_all("<Button-5>", _on_wheel)     # Linux wheel down

            # --- 1. Files --------------------------------------------------
            top = ttk.LabelFrame(
                self.body, text="1.  Files (.vsdx Visio / .xlsx Excel)"
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
            self.add_btn = self._rbtn(btn_row, "Add file...", self.add_files)
            self.folder_btn = self._rbtn(
                btn_row, "Add folder...", self.add_folder)
            _rm_btn = self._rbtn(
                btn_row, "Remove selected", self.remove_selected_files)
            _clr_btn = self._rbtn(btn_row, "Clear", self.clear_files)
            self._make_flow(
                btn_row, [self.add_btn, self.folder_btn, _rm_btn, _clr_btn])

            list_row = ttk.Frame(top)
            list_row.pack(fill="x", padx=8, pady=(2, 8))
            self.files_box = tk.Listbox(
                list_row, height=5, selectmode="extended",
                activestyle="none", bg=self.palette["field"],
                fg=self.palette["field_fg"], font=("Segoe UI", 10),
                relief="solid", borderwidth=1,
                highlightthickness=0, selectbackground=self.palette["sel"],
                selectforeground="#ffffff",
            )
            self.files_box.pack(side="left", fill="both", expand=True)
            sb = ttk.Scrollbar(
                list_row, orient="vertical", command=self.files_box.yview
            )
            sb.pack(side="left", fill="y")
            self.files_box.configure(yscrollcommand=sb.set)

            # --- 2. Feature tabs -------------------------------------------
            self.nb = ttk.Notebook(self.body)
            self.nb.pack(fill="x", **pad)

            # Tab 1 -- Find & Replace (the find->replace rules only).
            tab_fr = ttk.Frame(self.nb)
            self.nb.add(tab_fr, text="  Find → Replace  ")
            ttk.Label(
                tab_fr, wraplength=900, justify="left",
                text="Type the text to find and its replacement, then choose "
                "which document type(s) it applies to via 'in'.",
            ).pack(fill="x", padx=8, pady=(8, 2))
            self.pairs_frame = ttk.Frame(tab_fr)
            self.pairs_frame.pack(fill="x", padx=6, pady=6)
            self._header_row()
            self.add_pair()
            rule_btns = ttk.Frame(tab_fr)
            rule_btns.pack(fill="x", padx=8, pady=(0, 8))
            self._rbtn(
                rule_btns, "+ Add another rule", self.add_pair
            ).pack(side="left")

            # Tab 2 -- Parts (find & edit, remove, add).
            tab_parts = ttk.Frame(self.nb)
            self.nb.add(tab_parts, text="  Parts  ")
            fe = ttk.LabelFrame(tab_parts, text="Find & edit existing rows")
            fe.pack(fill="x", padx=8, pady=(8, 4))
            fe_btns = ttk.Frame(fe)
            fe_btns.pack(fill="x", padx=6, pady=6)
            self._make_flow(fe_btns, [
                self._rbtn(fe_btns, "Excel: find & edit rows...",
                           self.open_bom_editor),
                self._rbtn(fe_btns, "Visio: find & edit rows...",
                           self.open_visio_bom_editor),
            ])
            self.bom_status = ttk.Label(fe, text="")
            self.bom_status.pack(anchor="w", padx=10, pady=(0, 6))

            rmf = ttk.LabelFrame(tab_parts, text="Remove a part from sheets")
            rmf.pack(fill="x", padx=8, pady=4)
            rm_in = ttk.Frame(rmf)
            rm_in.pack(fill="x", padx=8, pady=(8, 2))
            ttk.Label(rm_in, text="Part number:").pack(side="left")
            self.remove_pn_var = tk.StringVar()
            ttk.Entry(rm_in, textvariable=self.remove_pn_var, width=28).pack(
                side="left", padx=6)
            self._rbtn(rm_in, "Find & choose rows to remove...",
                       self.open_remove_parts, kind="orange").pack(side="left")
            self.remove_status = ttk.Label(rmf, text="")
            self.remove_status.pack(anchor="w", padx=10, pady=(0, 6))

            adf = ttk.LabelFrame(tab_parts, text="Add new parts to a sheet")
            adf.pack(fill="x", padx=8, pady=(4, 8))
            ad_row = ttk.Frame(adf)
            ad_row.pack(fill="x", padx=8, pady=8)
            self._rbtn(ad_row, "Add parts...", self.open_add_parts,
                       kind="accent").pack(side="left")
            self.add_status = ttk.Label(adf, text="")
            self.add_status.pack(anchor="w", padx=10, pady=(0, 6))

            # Tab 3 -- Approve & Revise.
            tab_appr = ttk.Frame(self.nb)
            self.nb.add(tab_appr, text="  Approve & Revise  ")
            appr_btns = ttk.Frame(tab_appr)
            appr_btns.pack(fill="x", padx=8, pady=(10, 4))
            self._make_flow(appr_btns, [
                self._rbtn(appr_btns, "Visio: approve revision...",
                           self.open_visio_approval, kind="green"),
                self._rbtn(appr_btns, "Excel: approve (EE/ME/Prod)...",
                           self.open_excel_approval, kind="green"),
            ])
            rev_btns = ttk.Frame(tab_appr)
            rev_btns.pack(fill="x", padx=8, pady=(0, 10))
            self._make_flow(rev_btns, [
                self._rbtn(rev_btns, "Visio: add revision entry...",
                           self.open_visio_rev_editor),
                self._rbtn(rev_btns, "Excel: add Change Log entry...",
                           self.open_changelog_editor),
                self._rbtn(rev_btns, "Excel: set Author + date...",
                           self.open_author_editor),
            ])

            # --- 3. Options ------------------------------------------------
            opts = ttk.LabelFrame(self.body, text="3.  Options")
            opts.pack(fill="x", **pad)
            self.case_var = tk.BooleanVar(value=False)
            self.word_var = tk.BooleanVar(value=False)
            # PDF export (LibreOffice) is off by default: export from Visio for
            # correct sheet numbers and full fidelity.
            self.pdf_var = tk.BooleanVar(value=False)
            self.rev_var = tk.BooleanVar(value=True)
            self.revtext_var = tk.BooleanVar(value=True)
            self.summary_var = tk.BooleanVar(value=True)

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

            opt_row3 = ttk.Frame(opts)
            opt_row3.pack(fill="x")
            ttk.Checkbutton(
                opt_row3,
                text="Generate change summary (before/after, for approval "
                "review)",
                variable=self.summary_var,
            ).pack(side="left", padx=8, pady=(0, 6))

            # Output folder: where all finished files (and the summary) land.
            opt_row4 = ttk.Frame(opts)
            opt_row4.pack(fill="x", pady=(2, 6))
            ttk.Label(opt_row4, text="Output folder:").pack(
                side="left", padx=(8, 4)
            )
            self._rbtn(
                opt_row4, "Choose folder...", self._choose_out_dir, radius=10,
            ).pack(side="left")
            self._rbtn(
                opt_row4, "Use source folder", self._clear_out_dir, radius=10,
            ).pack(side="left", padx=6)
            self.out_dir_lbl = ttk.Label(
                opt_row4, text="(same folder as each source file)",
                foreground=self.palette["muted"],
            )
            self.out_dir_lbl.pack(side="left", padx=6)

            # --- Run -------------------------------------------------------
            run = ttk.Frame(self.body)
            run.pack(fill="x", **pad)
            self.run_btn = self._rbtn(
                run, "Run", self.run, kind="accent",
                radius=13, padx=22, pady=10,
            )
            self.run_btn.pack(side="left", padx=8)
            self._rbtn(
                run, "Reset all", self.reset_all, kind="orange",
                radius=13, padx=16, pady=10,
            ).pack(side="right", padx=8)
            self.progress = ttk.Progressbar(
                run, mode="indeterminate", style="Horizontal.TProgressbar"
            )
            self.progress.pack(side="left", fill="x", expand=True, padx=8)

            # --- Log -------------------------------------------------------
            logf = ttk.LabelFrame(self.body, text="Status")
            logf.pack(fill="both", expand=True, **pad)
            self.log_box = scrolledtext.ScrolledText(
                logf, height=8, state="disabled", wrap="word",
                bg=self.palette["field"], fg=self.palette["field_fg"],
                font=("Consolas", 9), relief="flat", borderwidth=0,
                highlightthickness=0,
            )
            self.log_box.pack(fill="both", expand=True, padx=6, pady=6)

            self.on_mode_change()
            self._sync_rev_options()
            self._log(f"Drawing & BOM Studio  v{__version__}")
            if find_libreoffice() is None:
                self._log(
                    "Note: LibreOffice was not found, so PDF export is "
                    "unavailable. Install it from libreoffice.org to enable "
                    "PDF output. Text replacement still works."
                )

        def _setup_style(self):
            """Apply the current palette to the ttk widgets."""
            c = self.palette
            base = ("Segoe UI", 10)
            style = ttk.Style()
            try:
                style.theme_use("clam")
            except tk.TclError:
                pass
            self.root.configure(bg=c["bg"])
            style.configure(".", font=base, background=c["bg"],
                            foreground=c["text"])
            style.configure("TFrame", background=c["bg"])
            style.configure("TLabel", background=c["bg"], foreground=c["text"])
            for w in ("TCheckbutton", "TRadiobutton"):
                style.configure(w, background=c["bg"], foreground=c["text"])
                style.map(w, background=[("active", c["bg"])],
                          foreground=[("disabled", c["muted"])])
            style.configure("TLabelframe", background=c["bg"],
                            bordercolor=c["border"], relief="solid",
                            borderwidth=1)
            style.configure("TLabelframe.Label", background=c["bg"],
                            foreground=c["accent"],
                            font=("Segoe UI", 10, "bold"))
            style.configure("TNotebook", background=c["bg"],
                            bordercolor=c["border"], tabmargins=(4, 4, 4, 0))
            style.configure("TNotebook.Tab", background=c["btn"],
                            foreground=c["text"], padding=(16, 7),
                            font=("Segoe UI", 10, "bold"))
            style.map("TNotebook.Tab",
                      background=[("selected", c["card"]),
                                  ("active", c["btn_hover"])],
                      foreground=[("selected", c["accent"])])
            style.configure("TEntry", fieldbackground=c["field"],
                            foreground=c["field_fg"], bordercolor=c["border"],
                            insertcolor=c["field_fg"], padding=4)
            style.map("TEntry", fieldbackground=[("disabled", c["bg"])])
            style.configure("TMenubutton", background=c["field"],
                            foreground=c["field_fg"], padding=(8, 4),
                            arrowcolor=c["text"], relief="solid",
                            borderwidth=1, bordercolor=c["border"])
            style.map("TMenubutton", background=[("active", c["btn_hover"])])
            style.configure("TCombobox", fieldbackground=c["field"],
                            foreground=c["field_fg"])
            style.configure("Horizontal.TProgressbar",
                            background=c["accent"], troughcolor=c["card"],
                            bordercolor=c["bg"])
            style.configure("Vertical.TScrollbar", background=c["btn"],
                            troughcolor=c["bg"], bordercolor=c["bg"],
                            arrowcolor=c["text"])

        def _rbtn(self, parent, text, command, kind="normal", **kw):
            b = RoundedButton(parent, text=text, command=command, kind=kind,
                              palette=self.palette, **kw)
            self._rbuttons.append(b)
            return b

        def _make_flow(self, fr, widgets, pad=4):
            """Lay buttons out in a frame, wrapping to new rows when the frame
            is too narrow -- so every button (and its full text) stays visible
            instead of being clipped at the window edge. Uses absolute placement
            (not grid, which would align columns across the wrapped rows)."""
            fr.pack_propagate(False)  # we set the frame's height ourselves
            state = {"w": 0}

            def reflow(e=None):
                # Use the event's width: winfo_width() can still report the old
                # value while a resize is being processed.
                w = (e.width if (e is not None and getattr(e, "width", 0) > 1)
                     else fr.winfo_width())
                if w <= 1 or w == state["w"]:
                    return
                state["w"] = w
                x = pad
                y = pad
                row_h = 0
                for b in widgets:
                    if not b.winfo_exists():
                        continue
                    bw, bh = b.winfo_reqwidth(), b.winfo_reqheight()
                    if x > pad and x + bw + pad > w:
                        x = pad
                        y += row_h + pad
                        row_h = 0
                    b.place(x=x, y=y)
                    x += bw + pad
                    row_h = max(row_h, bh)
                fr.configure(height=y + row_h + pad)

            fr.bind("<Configure>", reflow)
            self.root.after(0, reflow)

        def toggle_theme(self):
            self.dark = not self.dark
            self.palette = DARK if self.dark else LIGHT
            self.theme_btn.configure_text(
                "☀  Light" if self.dark else "🌙  Dark"
            )
            self._setup_style()
            self._recolor()

        def _recolor(self):
            c = self.palette
            self._banner.configure(bg=c["banner"])
            self._banner_title.configure(bg=c["banner"])
            self._banner_sub.configure(bg=c["banner"], fg=c["banner_sub"])
            self.files_box.configure(
                bg=c["field"], fg=c["field_fg"], selectbackground=c["sel"]
            )
            self.log_box.configure(bg=c["field"], fg=c["field_fg"])
            self.out_dir_lbl.configure(foreground=c["muted"])
            self._scroll_outer.configure(bg=c["bg"])
            self._scroll_canvas.configure(bg=c["bg"])
            self._rbuttons = [b for b in self._rbuttons if b.winfo_exists()]
            for b in self._rbuttons:
                b.set_palette(c)
            self._refresh_rule_menus()  # recolors the dropdown menus

        def _sync_rev_options(self):
            self.revtext_chk.configure(
                state="normal" if self.rev_var.get() else "disabled"
            )

        # -- file list ------------------------------------------------------
        def on_mode_change(self):
            single = self.mode_var.get() == "single"
            self.add_btn.configure_text(
                "Choose file..." if single else "Add files..."
            )
            self.folder_btn.set_enabled(not single)
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

        def _present_categories(self) -> List[str]:
            present = {file_category(f) for f in self.files}
            return [c for c in CATEGORY_ORDER if c in present]

        def _refresh_rule_menus(self):
            """Rebuild every rule's multi-select dropdown of file TYPES."""
            cats = self._present_categories()
            pal = self.palette
            for rule in self.rule_rows:
                menu = rule["menu"]
                menu.configure(
                    bg=pal["field"], fg=pal["field_fg"],
                    activebackground=pal["accent"],
                    activeforeground="#ffffff",
                    selectcolor=pal["accent"], borderwidth=0,
                )
                menu.delete(0, "end")
                cv = rule["cat_vars"]
                for c in list(cv):  # drop categories no longer present
                    if c not in cats:
                        del cv[c]
                for c in cats:
                    cv.setdefault(c, tk.BooleanVar(value=False))
                menu.add_checkbutton(
                    label="All files", variable=rule["all_var"],
                    command=lambda r=rule: self._on_all_toggle(r),
                )
                menu.add_separator()
                for c in cats:
                    menu.add_checkbutton(
                        label=c, variable=cv[c],
                        command=lambda r=rule: self._on_cat_toggle(r),
                    )
                self._update_menu_label(rule)

        def _on_all_toggle(self, rule):
            if rule["all_var"].get():
                for v in rule["cat_vars"].values():
                    v.set(False)
            self._update_menu_label(rule)

        def _on_cat_toggle(self, rule):
            if any(v.get() for v in rule["cat_vars"].values()):
                rule["all_var"].set(False)
            self._update_menu_label(rule)

        def _update_menu_label(self, rule):
            if rule["all_var"].get():
                rule["menubtn"]["text"] = "All files"
                return
            sel = [c for c, v in rule["cat_vars"].items() if v.get()]
            if not sel:
                rule["menubtn"]["text"] = "(none)"
            elif len(sel) == 1:
                rule["menubtn"]["text"] = sel[0]
            else:
                rule["menubtn"]["text"] = f"{len(sel)} types"

        # -- rule rows ------------------------------------------------------
        def _header_row(self):
            ttk.Label(
                self.pairs_frame, wraplength=820, justify="left",
                text="Type the text to find and its replacement, then choose "
                "which document type(s) it applies to via 'in' "
                "(BOM=DOC, System Drawing=DWG, Cable Drawing=CBL).",
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
            menubtn = ttk.Menubutton(row, text="All files", width=15)
            menu = tk.Menu(menubtn, tearoff=0)
            menubtn["menu"] = menu
            menubtn.pack(side="left", padx=(0, 4))
            rule = {
                "frame": row, "find": find_e, "repl": repl_e,
                "menubtn": menubtn, "menu": menu,
                "all_var": tk.BooleanVar(value=True), "cat_vars": {},
            }
            self._rbtn(
                row, "✕", lambda r=rule: self.remove_pair(r),
                padx=10, pady=6,
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
            self.run_btn.set_enabled(not busy)
            if busy:
                self.progress.start(12)
            else:
                self.progress.stop()

        def pairs_for_file(self, path: str) -> List[Tuple[str, str]]:
            """Find/replace pairs whose dropdown selection includes this file's
            document type."""
            cat = file_category(path)
            pairs = []
            for rule in self.rule_rows:
                find = rule["find"].get()
                if not find:
                    continue
                if rule["all_var"].get():
                    pairs.append((find, rule["repl"].get()))
                else:
                    v = rule["cat_vars"].get(cat)
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
            if not self._find_values():
                messagebox.showinfo(
                    "No Find values",
                    "Type the part number(s) into the Find box(es) first.",
                )
                return
            if not any(detect_format(f) == "xlsx" for f in self.files):
                messagebox.showinfo(
                    "No Excel files", "Add at least one .xlsx file first."
                )
                return
            self._build_bom_window()

        def _build_bom_window(self):
            pal = self.palette
            win = tk.Toplevel(self.root)
            win.title("Find & edit Excel rows")
            win.geometry("840x640")
            win.minsize(660, 480)
            win.transient(self.root)
            win.grab_set()
            win.configure(bg=pal["bg"])

            # Persistent state, kept across filtering / refresh.
            #   all  : every scanned match (excel_scan_rows dict + 'file')
            #   orig : (file, sheet, ref) -> original value
            #   cur  : (file, sheet, ref) -> current (edited) value
            #   rows : currently displayed entries [{key, field, entry}]
            state = {"all": [], "orig": {}, "cur": {}, "rows": []}
            filter_vars: dict = {}  # find value -> BooleanVar (dropdown)

            def key(mt, ref):
                return (mt["file"], mt["sheet"], ref)

            def capture():
                for r in state["rows"]:
                    try:
                        state["cur"][r["key"]] = r["entry"].get()
                    except tk.TclError:
                        pass

            def shown_finds():
                sel = [fv for fv, v in filter_vars.items() if v.get()]
                return set(sel) if sel else set(filter_vars)  # none = all

            ttk.Label(
                win, wraplength=810, justify="left",
                text="Each matched row is shown per file and sheet. Use the "
                "dropdown to show only certain find values. The 'Copy 1st down' "
                "buttons copy the first shown row's value into the rest of that "
                "find value's rows. Edit any field, then Save.",
            ).pack(side="top", fill="x", padx=10, pady=(8, 4))

            # --- filter dropdown -----------------------------------------
            frow = ttk.Frame(win)
            frow.pack(side="top", fill="x", padx=10, pady=(0, 2))
            ttk.Label(frow, text="Show find value(s):").pack(side="left")
            filt_btn = ttk.Menubutton(frow, text="All", width=22)
            filt_menu = tk.Menu(filt_btn, tearoff=0)
            filt_btn["menu"] = filt_menu
            filt_btn.pack(side="left", padx=(4, 0))

            # --- copy-down buttons ---------------------------------------
            crow = ttk.Frame(win)
            crow.pack(side="top", fill="x", padx=10, pady=(0, 4))
            ttk.Label(
                crow, text="Copy 1st down (per find value):"
            ).pack(side="left")
            for _fld in ("Manufacturer", "Unit Cost", "Description", "Qty"):
                self._rbtn(
                    crow, _fld, lambda fld=_fld: copy_down(fld),
                    kind="green", radius=9, padx=10, pady=5,
                ).pack(side="left", padx=3)

            bottom = ttk.Frame(win)
            bottom.pack(side="bottom", fill="x", padx=10, pady=10)

            body = ttk.Frame(win)
            body.pack(side="top", fill="both", expand=True)
            canvas = tk.Canvas(body, highlightthickness=0, bg=pal["bg"])
            sb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
            inner = ttk.Frame(canvas)
            inner.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
            )
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=sb.set)
            canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
            sb.pack(side="left", fill="y", padx=(0, 10))

            def update_filter_label():
                sel = [fv for fv, v in filter_vars.items() if v.get()]
                if not sel:
                    filt_btn["text"] = "All"
                elif len(sel) == 1:
                    filt_btn["text"] = sel[0]
                else:
                    filt_btn["text"] = f"{len(sel)} selected"

            def render():
                for w in inner.winfo_children():
                    w.destroy()
                state["rows"] = []
                finds = shown_finds()
                matches = [mt for mt in state["all"]
                           if mt["matched"] in finds]
                if not matches:
                    ttk.Label(
                        inner,
                        text="No matching rows for the current selection.",
                    ).pack(padx=10, pady=12)
                    canvas.update_idletasks()
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    return
                by_file: dict = {}
                for mt in matches:
                    by_file.setdefault(mt["file"], []).append(mt)
                for f, fmatches in by_file.items():
                    ttk.Label(
                        inner, text="📄  " + Path(f).name,
                        font=("Segoe UI", 10, "bold"),
                        foreground=pal["accent"],
                    ).pack(anchor="w", padx=6, pady=(10, 0))
                    by_sheet: dict = {}
                    for mt in fmatches:
                        by_sheet.setdefault(
                            (mt["sheet"], mt.get("sheet_name") or ""), []
                        ).append(mt)
                    for (sp, sn), smatches in by_sheet.items():
                        if sn:
                            ttk.Label(
                                inner, text="     Sheet:  " + sn,
                                font=("Segoe UI", 9, "bold"),
                                foreground=pal["muted"],
                            ).pack(anchor="w", padx=12, pady=(4, 0))
                        for mt in smatches:
                            lf = ttk.LabelFrame(
                                inner,
                                text=f"P/N: {mt['part']}    (row {mt['row']})",
                            )
                            lf.pack(fill="x", padx=12, pady=4)
                            for field in BOM_FIELD_ORDER:
                                if field == "Part Number":
                                    continue
                                cell = mt["fields"].get(field)
                                if cell is None:
                                    continue
                                ref, val = cell
                                k = key(mt, ref)
                                rowf = ttk.Frame(lf)
                                rowf.pack(fill="x", padx=6, pady=2)
                                ttk.Label(
                                    rowf, text=field + ":", width=14
                                ).pack(side="left")
                                e = ttk.Entry(rowf)
                                e.insert(0, state["cur"].get(k, val))
                                e.pack(side="left", fill="x", expand=True)
                                state["rows"].append(
                                    {"key": k, "field": field, "entry": e}
                                )
                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))

            def filter_changed():
                capture()
                update_filter_label()
                render()

            def show_all():
                for v in filter_vars.values():
                    v.set(False)
                filter_changed()

            def rebuild_filter_menu():
                finds = sorted({mt["matched"] for mt in state["all"]})
                for fv in list(filter_vars):
                    if fv not in finds:
                        del filter_vars[fv]
                for fv in finds:
                    filter_vars.setdefault(fv, tk.BooleanVar(value=False))
                filt_menu.configure(
                    bg=pal["field"], fg=pal["field_fg"],
                    activebackground=pal["accent"],
                    activeforeground="#ffffff",
                    selectcolor=pal["accent"], borderwidth=0,
                )
                filt_menu.delete(0, "end")
                filt_menu.add_command(label="(show all)", command=show_all)
                filt_menu.add_separator()
                for fv in finds:
                    filt_menu.add_checkbutton(
                        label=fv, variable=filter_vars[fv],
                        command=filter_changed,
                    )
                update_filter_label()

            def copy_down(field):
                capture()
                n = 0
                for fv in shown_finds():
                    rows_fv = [mt for mt in state["all"]
                               if mt["matched"] == fv
                               and mt["fields"].get(field)]
                    if len(rows_fv) < 2:
                        continue
                    src = key(rows_fv[0], rows_fv[0]["fields"][field][0])
                    val = state["cur"].get(src, "")
                    for mt in rows_fv[1:]:
                        k = key(mt, mt["fields"][field][0])
                        if state["cur"].get(k) != val:
                            state["cur"][k] = val
                            n += 1
                render()
                self.log(
                    f"Copied the first {field} value down to {n} row(s)."
                    if n else f"Nothing to copy for {field}."
                )

            def reset_fields():
                state["cur"] = dict(state["orig"])
                render()
                self.log("Reset all BOM fields to the originally found data.")

            def scan():
                capture()
                parts = self._find_values()
                cs = self.case_var.get()
                state["all"] = []
                for f in [x for x in self.files
                          if detect_format(x) == "xlsx"]:
                    try:
                        state["all"].extend(excel_scan_rows(f, parts, cs))
                    except Exception:  # noqa: BLE001
                        pass
                for mt in state["all"]:
                    for _fld, (ref, val) in mt["fields"].items():
                        k = key(mt, ref)
                        state["orig"].setdefault(k, val)
                        state["cur"].setdefault(k, state["orig"][k])
                rebuild_filter_menu()
                render()

            def apply():
                capture()
                edits: dict = {}  # {file: {sheet: {ref: value}}}
                n = 0
                for mt in state["all"]:
                    for _fld, (ref, _v) in mt["fields"].items():
                        k = key(mt, ref)
                        cur, orig = state["cur"].get(k), state["orig"].get(k)
                        if cur is not None and cur != orig:
                            edits.setdefault(mt["file"], {}).setdefault(
                                mt["sheet"], {}
                            )[ref] = cur
                            n += 1
                self.bom_edits = edits
                self.bom_status.configure(
                    text=(f"{n} cell edit(s) staged" if n else "")
                )
                self.log(
                    f"Staged {n} Excel cell edit(s) across {len(edits)} file(s)."
                    if n else "No Excel cell edits staged."
                )
                win.destroy()

            scan()

            self._rbtn(bottom, "Save edits", apply, kind="accent").pack(
                side="right"
            )
            self._rbtn(bottom, "Cancel", win.destroy).pack(
                side="right", padx=6
            )
            self._rbtn(bottom, "Refresh lookup", scan).pack(side="left")
            self._rbtn(
                bottom, "Reset fields", reset_fields, kind="orange",
            ).pack(side="left", padx=6)

        # -- Visio parts-table (Cable BOM) row editor ----------------------
        def open_visio_bom_editor(self):
            if not self._find_values():
                messagebox.showinfo(
                    "No Find values",
                    "Type the part number(s) into the Find box(es) first.",
                )
                return
            if not any(detect_format(f) == "vsdx" for f in self.files):
                messagebox.showinfo(
                    "No Visio files", "Add at least one .vsdx file first."
                )
                return
            self._build_visio_bom_window()

        def _build_visio_bom_window(self):
            pal = self.palette
            win = tk.Toplevel(self.root)
            win.title("Find & edit Visio parts-table rows")
            win.geometry("860x640")
            win.minsize(680, 480)
            win.transient(self.root)
            win.grab_set()
            win.configure(bg=pal["bg"])

            # Persistent state, kept across filtering / refresh (mirrors the
            # Excel editor). Keys are (file, embed, ref).
            state = {"all": [], "orig": {}, "cur": {}, "rows": []}
            filter_vars: dict = {}

            def key(mt, ref):
                return (mt["file"], mt["embed"], ref)

            def capture():
                for r in state["rows"]:
                    try:
                        state["cur"][r["key"]] = r["entry"].get()
                    except tk.TclError:
                        pass

            def shown_finds():
                sel = [fv for fv, v in filter_vars.items() if v.get()]
                return set(sel) if sel else set(filter_vars)

            ttk.Label(
                win, wraplength=830, justify="left",
                text="The parts list at the bottom of each R/C/P/W sheet is an "
                "embedded table. Rows are matched by Part Number and shown per "
                "file and sheet. Use the dropdown to show only certain find "
                "values. The 'Copy 1st down' buttons copy the first shown row's "
                "value into the rest of that find value's rows. Edit any field, "
                "then Save.",
            ).pack(side="top", fill="x", padx=10, pady=(8, 4))

            frow = ttk.Frame(win)
            frow.pack(side="top", fill="x", padx=10, pady=(0, 2))
            ttk.Label(frow, text="Show find value(s):").pack(side="left")
            filt_btn = ttk.Menubutton(frow, text="All", width=22)
            filt_menu = tk.Menu(filt_btn, tearoff=0)
            filt_btn["menu"] = filt_menu
            filt_btn.pack(side="left", padx=(4, 0))

            crow = ttk.Frame(win)
            crow.pack(side="top", fill="x", padx=10, pady=(0, 4))
            ttk.Label(
                crow, text="Copy 1st down (per find value):"
            ).pack(side="left")
            for _fld in VISIO_BOM_COPYDOWN:
                self._rbtn(
                    crow, _fld, lambda fld=_fld: copy_down(fld),
                    kind="green", radius=9, padx=10, pady=5,
                ).pack(side="left", padx=3)

            bottom = ttk.Frame(win)
            bottom.pack(side="bottom", fill="x", padx=10, pady=10)

            body = ttk.Frame(win)
            body.pack(side="top", fill="both", expand=True)
            canvas = tk.Canvas(body, highlightthickness=0, bg=pal["bg"])
            sb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
            inner = ttk.Frame(canvas)
            inner.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
            )
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=sb.set)
            canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
            sb.pack(side="left", fill="y", padx=(0, 10))

            def update_filter_label():
                sel = [fv for fv, v in filter_vars.items() if v.get()]
                if not sel:
                    filt_btn["text"] = "All"
                elif len(sel) == 1:
                    filt_btn["text"] = sel[0]
                else:
                    filt_btn["text"] = f"{len(sel)} selected"

            def render():
                for w in inner.winfo_children():
                    w.destroy()
                state["rows"] = []
                finds = shown_finds()
                matches = [mt for mt in state["all"] if mt["matched"] in finds]
                if not matches:
                    ttk.Label(
                        inner,
                        text="No matching rows for the current selection.",
                    ).pack(padx=10, pady=12)
                    canvas.update_idletasks()
                    canvas.configure(scrollregion=canvas.bbox("all"))
                    return
                by_file: dict = {}
                for mt in matches:
                    by_file.setdefault(mt["file"], []).append(mt)
                for f, fmatches in by_file.items():
                    ttk.Label(
                        inner, text="📄  " + Path(f).name,
                        font=("Segoe UI", 10, "bold"),
                        foreground=pal["accent"],
                    ).pack(anchor="w", padx=6, pady=(10, 0))
                    by_sheet: dict = {}
                    for mt in fmatches:
                        by_sheet.setdefault(
                            mt.get("sheet_name") or "", []).append(mt)
                    for sn, smatches in by_sheet.items():
                        if sn:
                            ttk.Label(
                                inner, text="     Sheet:  " + sn,
                                font=("Segoe UI", 9, "bold"),
                                foreground=pal["muted"],
                            ).pack(anchor="w", padx=12, pady=(4, 0))
                        for mt in smatches:
                            lf = ttk.LabelFrame(
                                inner,
                                text=f"P/N: {mt['part']}    (row {mt['row']})",
                            )
                            lf.pack(fill="x", padx=12, pady=4)
                            for field in VISIO_BOM_EDIT_FIELDS:
                                cell = mt["fields"].get(field)
                                if cell is None:
                                    continue
                                ref, val = cell
                                k = key(mt, ref)
                                rowf = ttk.Frame(lf)
                                rowf.pack(fill="x", padx=6, pady=2)
                                ttk.Label(
                                    rowf, text=field + ":", width=14
                                ).pack(side="left")
                                e = ttk.Entry(rowf)
                                e.insert(0, state["cur"].get(k, val))
                                e.pack(side="left", fill="x", expand=True)
                                state["rows"].append(
                                    {"key": k, "field": field, "entry": e}
                                )
                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))

            def filter_changed():
                capture()
                update_filter_label()
                render()

            def show_all():
                for v in filter_vars.values():
                    v.set(False)
                filter_changed()

            def rebuild_filter_menu():
                finds = sorted({mt["matched"] for mt in state["all"]})
                for fv in list(filter_vars):
                    if fv not in finds:
                        del filter_vars[fv]
                for fv in finds:
                    filter_vars.setdefault(fv, tk.BooleanVar(value=False))
                filt_menu.configure(
                    bg=pal["field"], fg=pal["field_fg"],
                    activebackground=pal["accent"],
                    activeforeground="#ffffff",
                    selectcolor=pal["accent"], borderwidth=0,
                )
                filt_menu.delete(0, "end")
                filt_menu.add_command(label="(show all)", command=show_all)
                filt_menu.add_separator()
                for fv in finds:
                    filt_menu.add_checkbutton(
                        label=fv, variable=filter_vars[fv],
                        command=filter_changed,
                    )
                update_filter_label()

            def copy_down(field):
                capture()
                n = 0
                for fv in shown_finds():
                    rows_fv = [mt for mt in state["all"]
                               if mt["matched"] == fv
                               and mt["fields"].get(field)]
                    if len(rows_fv) < 2:
                        continue
                    src = key(rows_fv[0], rows_fv[0]["fields"][field][0])
                    val = state["cur"].get(src, "")
                    for mt in rows_fv[1:]:
                        k = key(mt, mt["fields"][field][0])
                        if state["cur"].get(k) != val:
                            state["cur"][k] = val
                            n += 1
                render()
                self.log(
                    f"Copied the first {field} value down to {n} row(s)."
                    if n else f"Nothing to copy for {field}."
                )

            def reset_fields():
                state["cur"] = dict(state["orig"])
                render()
                self.log("Reset all parts-table fields to the found data.")

            def scan():
                capture()
                parts = self._find_values()
                cs = self.case_var.get()
                state["all"] = []
                for f in [x for x in self.files
                          if detect_format(x) == "vsdx"]:
                    try:
                        state["all"].extend(visio_bom_scan_rows(f, parts, cs))
                    except Exception:  # noqa: BLE001
                        pass
                for mt in state["all"]:
                    for _fld, (ref, val) in mt["fields"].items():
                        k = key(mt, ref)
                        state["orig"].setdefault(k, val)
                        state["cur"].setdefault(k, state["orig"][k])
                rebuild_filter_menu()
                render()

            def apply():
                capture()
                # {file: {embed: {"emf": emf, "cells": [edit, ...]}}}
                edits: dict = {}
                n = 0
                for mt in state["all"]:
                    for field, (ref, _v) in mt["fields"].items():
                        k = key(mt, ref)
                        cur, orig = state["cur"].get(k), state["orig"].get(k)
                        if cur is None or cur == orig:
                            continue
                        col = re.match(r"[A-Z]+", ref).group(0)
                        entry = edits.setdefault(mt["file"], {}).setdefault(
                            mt["embed"], {"emf": mt["emf"], "cells": []})
                        entry["cells"].append({
                            "ref": ref, "col": col, "row": mt["row"],
                            "field": field, "old": orig, "new": cur,
                            "pn": mt["part"], "item": mt.get("item", ""),
                            "row_index": mt.get("row_index")})
                        n += 1
                self.visio_bom_edits = edits
                self.bom_status.configure(
                    text=(f"{n} Visio cell edit(s) staged" if n else "")
                )
                self.log(
                    f"Staged {n} Visio parts-table edit(s) across "
                    f"{len(edits)} file(s)." if n
                    else "No Visio parts-table edits staged."
                )
                win.destroy()

            scan()

            self._rbtn(bottom, "Save edits", apply, kind="accent").pack(
                side="right"
            )
            self._rbtn(bottom, "Cancel", win.destroy).pack(
                side="right", padx=6
            )
            self._rbtn(bottom, "Refresh lookup", scan).pack(side="left")
            self._rbtn(
                bottom, "Reset fields", reset_fields, kind="orange",
            ).pack(side="left", padx=6)

        # -- Remove parts --------------------------------------------------
        def _scan_parts(self, pn):
            """Find rows with part number ``pn`` across all loaded files.
            Each: {file, fmt, sheet, sheet_name, row, part, item, desc}."""
            cs = self.case_var.get()
            out = []
            for f in self.files:
                fmt = detect_format(f)
                try:
                    if fmt == "xlsx":
                        for m in excel_scan_rows(f, [pn], cs):
                            out.append({
                                "file": f, "fmt": "xlsx", "sheet": m["sheet"],
                                "sheet_name": m.get("sheet_name") or "",
                                "row": m["row"], "part": m["part"], "item": "",
                                "desc": (m["fields"].get("Description")
                                         or ("", ""))[1]})
                    elif fmt == "vsdx":
                        for m in visio_bom_scan_rows(f, [pn], cs):
                            out.append({
                                "file": f, "fmt": "vsdx", "sheet": m["embed"],
                                "sheet_name": m.get("sheet_name") or "",
                                "row": m["row"], "part": m["part"],
                                "item": m.get("item", ""),
                                "desc": (m["fields"].get("Description")
                                         or ("", ""))[1]})
                except Exception:  # noqa: BLE001
                    pass
            return out

        def open_remove_parts(self):
            pn = self.remove_pn_var.get().strip()
            if not pn:
                messagebox.showinfo(
                    "Remove parts", "Type the part number to remove first.")
                return
            matches = self._scan_parts(pn)
            if not matches:
                messagebox.showinfo(
                    "Remove parts",
                    f"No parts-table rows found with part number '{pn}'.")
                return

            pal = self.palette
            win = tk.Toplevel(self.root)
            win.title(f"Remove part '{pn}'")
            win.geometry("820x600")
            win.transient(self.root)
            win.grab_set()
            win.configure(bg=pal["bg"])
            ttk.Label(
                win, wraplength=790, justify="left",
                text=f"Rows containing part number '{pn}', grouped by file and "
                "sheet. Tick the ones to remove. On run, each ticked row is "
                "deleted, the rows below it shift up, and the item/line numbers "
                "are renumbered.",
            ).pack(fill="x", padx=10, pady=(10, 4))

            body = ttk.Frame(win)
            body.pack(fill="both", expand=True)
            canvas = tk.Canvas(body, highlightthickness=0, bg=pal["bg"])
            sb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
            inner = ttk.Frame(canvas)
            inner.bind("<Configure>", lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=inner, anchor="nw")
            canvas.configure(yscrollcommand=sb.set)
            canvas.pack(side="left", fill="both", expand=True, padx=(10, 0))
            sb.pack(side="left", fill="y", padx=(0, 10))

            already = self.remove_parts
            state = []  # {file, fmt, sheet, row, var}
            by_file: dict = {}
            for m in matches:
                by_file.setdefault(m["file"], []).append(m)
            for f, fmatches in by_file.items():
                ttk.Label(
                    inner, text="📄  " + Path(f).name,
                    font=("Segoe UI", 10, "bold"), foreground=pal["accent"],
                ).pack(anchor="w", padx=6, pady=(10, 0))
                by_sheet: dict = {}
                for m in fmatches:
                    by_sheet.setdefault(
                        (m["sheet"], m["sheet_name"]), []).append(m)
                for (sheet, sname), smatches in by_sheet.items():
                    if sname:
                        ttk.Label(
                            inner, text="     Sheet:  " + sname,
                            font=("Segoe UI", 9, "bold"),
                            foreground=pal["muted"],
                        ).pack(anchor="w", padx=12, pady=(4, 0))
                    for m in smatches:
                        pre = (m["row"] in
                               already.get(f, {}).get(sheet, set()))
                        var = tk.BooleanVar(value=pre)
                        label = (
                            (f"Item {m['item']}  ·  " if m["item"] else "")
                            + f"P/N {m['part']}"
                            + (f"  ·  {m['desc'][:54]}" if m["desc"] else "")
                            + f"   (row {m['row']})")
                        ttk.Checkbutton(
                            inner, text=label, variable=var,
                        ).pack(anchor="w", padx=22, pady=1)
                        state.append({"file": f, "sheet": sheet,
                                      "row": m["row"], "var": var})

            def apply():
                removals: dict = {}
                n = 0
                for s in state:
                    if s["var"].get():
                        removals.setdefault(s["file"], {}).setdefault(
                            s["sheet"], set()).add(s["row"])
                        n += 1
                self.remove_parts = removals
                self.remove_status.configure(
                    text=(f"{n} part row(s) staged for removal" if n else ""))
                self.log(f"Staged {n} part removal(s)." if n
                         else "No part removals staged.")
                win.destroy()

            bottom = ttk.Frame(win)
            bottom.pack(side="bottom", fill="x", padx=10, pady=10)
            self._rbtn(bottom, "Save removals", apply, kind="orange").pack(
                side="right")
            self._rbtn(bottom, "Cancel", win.destroy).pack(
                side="right", padx=6)

        # -- Add parts -----------------------------------------------------
        def open_add_parts(self):
            if not self.files:
                messagebox.showinfo("Add parts", "Add files first.")
                return
            pal = self.palette
            win = tk.Toplevel(self.root)
            win.title("Add new parts")
            win.geometry("900x640")
            win.minsize(700, 480)
            win.transient(self.root)
            win.grab_set()
            win.configure(bg=pal["bg"])

            ttk.Label(
                win, wraplength=870, justify="left",
                text="Pick a file type, then the file and the sheet/table to "
                "add to. The current parts are shown for reference; fill in the "
                "blank row(s) at the bottom to add new parts. Item/line numbers "
                "are assigned automatically on run.",
            ).pack(fill="x", padx=10, pady=(10, 4))

            ctl = ttk.Frame(win)
            ctl.pack(fill="x", padx=10, pady=4)
            ttk.Label(ctl, text="File type:").pack(side="left")
            type_var = tk.StringVar()
            type_cb = ttk.Combobox(ctl, textvariable=type_var, width=16,
                                   state="readonly")
            type_cb.pack(side="left", padx=(4, 10))
            ttk.Label(ctl, text="File:").pack(side="left")
            file_var = tk.StringVar()
            file_cb = ttk.Combobox(ctl, textvariable=file_var, width=34,
                                   state="readonly")
            file_cb.pack(side="left", padx=(4, 10))
            ttk.Label(ctl, text="Sheet:").pack(side="left")
            sheet_var = tk.StringVar()
            sheet_cb = ttk.Combobox(ctl, textvariable=sheet_var, width=18,
                                    state="readonly")
            sheet_cb.pack(side="left", padx=(4, 0))

            body = ttk.Frame(win)
            body.pack(fill="both", expand=True, padx=10, pady=6)
            canvas = tk.Canvas(body, highlightthickness=0, bg=pal["bg"])
            vsb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
            hsb = ttk.Scrollbar(body, orient="horizontal",
                                command=canvas.xview)
            grid = ttk.Frame(canvas)
            grid.bind("<Configure>", lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")))
            canvas.create_window((0, 0), window=grid, anchor="nw")
            canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            vsb.pack(side="right", fill="y")
            hsb.pack(side="bottom", fill="x")
            canvas.pack(side="left", fill="both", expand=True)

            # mutable view state
            st = {"file": None, "fmt": None, "sheet": None, "cols": [],
                  "headers": {}, "new_entries": [], "existing_rows": []}

            def file_map():
                return {Path(f).name: f for f in self.files}

            def sheets_for(f):
                fmt = detect_format(f)
                if fmt == "vsdx":
                    return [(sid, nm or sid.split("/")[-1])
                            for sid, nm in visio_parts_sheets(f)]
                if fmt == "xlsx":
                    return [(sid, nm or sid.split("/")[-1])
                            for sid, nm in excel_bom_sheets(f)]
                return []

            def add_new_row(values=None):
                r = len(st["existing_rows"]) + 1 + len(st["new_entries"])
                ents = {}
                for ci, col in enumerate(st["cols"]):
                    e = ttk.Entry(grid, width=18)
                    if values and col in values:
                        e.insert(0, values[col])
                    e.grid(row=r, column=ci, sticky="we", padx=1, pady=1)
                    ents[col] = e
                st["new_entries"].append(ents)
                canvas.update_idletasks()
                canvas.configure(scrollregion=canvas.bbox("all"))

            def commit_current():
                """Capture the current sheet's filled new rows into add_parts."""
                f, sid = st.get("file"), st.get("sheet")
                if not (f and sid):
                    return
                new_rows = []
                for ents in st["new_entries"]:
                    vals = {c: e.get().strip() for c, e in ents.items()
                            if e.get().strip()}
                    if vals:
                        new_rows.append(vals)
                self.add_parts.setdefault(f, {})
                if new_rows:
                    self.add_parts[f][sid] = new_rows
                else:
                    self.add_parts[f].pop(sid, None)
                if not self.add_parts[f]:
                    self.add_parts.pop(f, None)

            def load_sheet(*_):
                commit_current()
                for w in grid.winfo_children():
                    w.destroy()
                st["new_entries"] = []
                st["existing_rows"] = []
                st["file"] = None
                st["sheet"] = None
                f = file_map().get(file_var.get())
                smap = {nm: sid for sid, nm in sheets_for(f)} if f else {}
                sid = smap.get(sheet_var.get())
                if not f or not sid:
                    return
                view = parts_table_view(f, detect_format(f), sid)
                if not view:
                    ttk.Label(grid, text="(couldn't read this table)").grid(
                        row=0, column=0, padx=6, pady=6)
                    return
                cols, headers, rows = view
                st.update({"file": f, "fmt": detect_format(f), "sheet": sid,
                           "cols": cols, "headers": headers,
                           "existing_rows": rows})
                for ci, col in enumerate(cols):
                    ttk.Label(grid, text=headers.get(col, col),
                              font=("Segoe UI", 9, "bold"),
                              foreground=pal["accent"]).grid(
                        row=0, column=ci, sticky="w", padx=2, pady=2)
                for ri, row in enumerate(rows, start=1):
                    for ci, col in enumerate(cols):
                        ttk.Label(grid, text=str(row.get(col, ""))[:30],
                                  foreground=pal["muted"]).grid(
                            row=ri, column=ci, sticky="w", padx=2, pady=1)
                # restore any staged new rows, else one blank row
                staged = self.add_parts.get(f, {}).get(sid, [])
                if staged:
                    for vals in staged:
                        add_new_row(vals)
                else:
                    add_new_row()

            def on_type(*_):
                f_of_type = [Path(f).name for f in self.files
                             if file_category(f) == type_var.get()]
                file_cb["values"] = f_of_type
                file_var.set(f_of_type[0] if f_of_type else "")
                on_file()

            def on_file(*_):
                f = file_map().get(file_var.get())
                names = [nm for _sid, nm in sheets_for(f)] if f else []
                sheet_cb["values"] = names
                sheet_var.set(names[0] if names else "")
                load_sheet()

            type_cb.bind("<<ComboboxSelected>>", on_type)
            file_cb.bind("<<ComboboxSelected>>", on_file)
            sheet_cb.bind("<<ComboboxSelected>>", load_sheet)

            cats = [c for c in CATEGORY_ORDER
                    if any(file_category(f) == c for f in self.files)]
            type_cb["values"] = cats
            if cats:
                type_var.set(cats[0])
                on_type()

            def save():
                commit_current()
                total = sum(len(v) for d in self.add_parts.values()
                            for v in d.values())
                self.add_status.configure(
                    text=(f"{total} new part(s) staged" if total else ""))
                self.log(f"Staged {total} new part(s) to add." if total
                         else "No new parts staged.")
                win.destroy()

            bottom = ttk.Frame(win)
            bottom.pack(side="bottom", fill="x", padx=10, pady=10)
            self._rbtn(bottom, "Save", save, kind="accent").pack(side="right")
            self._rbtn(bottom, "Cancel", win.destroy).pack(
                side="right", padx=6)
            self._rbtn(bottom, "+ Add row", lambda: add_new_row()).pack(
                side="left")

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
            win.configure(bg=self.palette["bg"])
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
            self._rbtn(btns, "Save entry", apply, kind="accent").pack(
                side="right"
            )
            self._rbtn(btns, "Cancel", win.destroy).pack(
                side="right", padx=6
            )

        # -- Author box ----------------------------------------------------
        def open_author_editor(self):
            from tkinter import simpledialog
            excel_files = [f for f in self.files
                           if detect_format(f) == "xlsx"]
            if not excel_files:
                messagebox.showinfo(
                    "No Excel files", "Add at least one .xlsx file first."
                )
                return
            if not any(build_author_edits(f, "x") for f in excel_files):
                messagebox.showinfo(
                    "No Author box",
                    "None of the Excel files have an 'Author' cell with a name "
                    "cell to its right.",
                )
                return
            name = simpledialog.askstring(
                "Author name",
                "Replace the name beside the 'Author' box with:\n"
                "(the date beside it is set to today automatically, in the "
                "same format)",
                initialvalue=self.author_name,
                parent=self.root,
            )
            if name is None:
                return
            self.author_name = name.strip()
            if self.author_name:
                self.log(
                    f"Author will be set to '{self.author_name}' "
                    "(date = today)."
                )
            else:
                self.log("Author update cleared.")

        # -- Visio revision table ------------------------------------------
        def open_visio_rev_editor(self):
            vsdx_files = [f for f in self.files
                          if detect_format(f) == "vsdx"]
            if not vsdx_files:
                messagebox.showinfo(
                    "No Visio files", "Add at least one .vsdx file first."
                )
                return
            # Report which revision-table columns were detected per file.
            detected = {}
            for f in vsdx_files:
                cols = vsdx_revtable_columns(f)
                if cols:
                    detected[Path(f).name] = cols
            if not detected:
                messagebox.showwarning(
                    "No revision table found",
                    "Couldn't confidently locate a revision-history table on "
                    "the cover page of any loaded Visio file.\n\n"
                    "The table is found by its column headers (REV, "
                    "DESCRIPTION, DATE, APPROVED, ECN, ...). If your drawing "
                    "uses different labels, share a sample .vsdx so it can be "
                    "tuned. You can still stage an entry below; files without "
                    "a detected table are left unchanged.",
                )

            win = tk.Toplevel(self.root)
            win.title("Add Visio revision entry")
            win.transient(self.root)
            win.grab_set()
            win.configure(bg=self.palette["bg"])
            note = ("The new row is added to every loaded Visio file's "
                    "revision table. The **REV** column is filled automatically "
                    "with each file's own next revision letter — so a batch of "
                    "files each gets its correct next letter. The other fields "
                    "below are the same for every file (leave any blank).")
            if detected:
                found = "; ".join(
                    f"{n}: {', '.join(cols)}" for n, cols in detected.items()
                )
                info = note + f"\n\nDetected columns — {found}"
            else:
                info = (note + "\n\nNo revision table was detected yet; any "
                        "file where the table is found at run time will get the "
                        "row, others are left unchanged.")
            ttk.Label(
                win, wraplength=580, justify="left",
                text=info.replace("**", ""),
            ).pack(fill="x", padx=10, pady=8)

            form = ttk.Frame(win)
            form.pack(fill="x", padx=10, pady=4)
            # "Rev" is filled per-file at run time, so it isn't entered here.
            entries = {}
            for field in REVTABLE_FIELD_ORDER:
                if field == "Rev":
                    continue
                rowf = ttk.Frame(form)
                rowf.pack(fill="x", pady=3)
                ttk.Label(rowf, text=field + ":", width=14).pack(side="left")
                e = ttk.Entry(rowf, width=46)
                e.insert(0, self.visio_rev_entry.get(field, ""))
                e.pack(side="left", fill="x", expand=True)
                entries[field] = e

            def apply():
                entry = {f: e.get() for f, e in entries.items()}
                if not any(v.strip() for v in entry.values()):
                    self.visio_rev_entry = {}
                    self.log("Visio revision entry cleared.")
                else:
                    self.visio_rev_entry = entry
                    self.log(
                        "Staged a Visio revision-table entry "
                        "(REV auto-filled per file)."
                    )
                win.destroy()

            btns = ttk.Frame(win)
            btns.pack(fill="x", padx=10, pady=10)
            self._rbtn(btns, "Save entry", apply, kind="accent").pack(
                side="right"
            )
            self._rbtn(btns, "Cancel", win.destroy).pack(
                side="right", padx=6
            )

        # -- Visio approval (sign off an existing revision row) ------------
        def open_visio_approval(self):
            vsdx_files = [f for f in self.files if detect_format(f) == "vsdx"]
            if not vsdx_files:
                messagebox.showinfo(
                    "No Visio files", "Add at least one .vsdx file first."
                )
                return
            # Detect each file's latest (most recent) revision letter.
            latest = {Path(f).name: vsdx_latest_revision_letter(f)
                      for f in vsdx_files}
            found = "; ".join(f"{n}: {l}" for n, l in latest.items() if l)
            avail = (f"Latest revision per file — {found}."
                     if found else "No revision letters were detected; you can "
                     "still enter one to try.")
            win = tk.Toplevel(self.root)
            win.title("Approve a Visio revision")
            win.transient(self.root)
            win.grab_set()
            win.configure(bg=self.palette["bg"])
            ttk.Label(
                win, wraplength=580, justify="left",
                text="Sign off a revision: your name is written into the "
                "Approved column of a revision row, in the revision table of "
                "every loaded Visio file. By default it signs off each file's "
                "OWN latest revision (auto-detected per file) — so a batch at "
                "different revisions each gets the right row. Leave the REV "
                "field blank for that, or type a specific letter to force one.\n\n"
                + avail,
            ).pack(fill="x", padx=10, pady=8)

            form = ttk.Frame(win)
            form.pack(fill="x", padx=10, pady=4)
            r2 = ttk.Frame(form)
            r2.pack(fill="x", pady=3)
            ttk.Label(r2, text="Approver name:", width=20).pack(side="left")
            name_e = ttk.Entry(r2, width=38)
            name_e.insert(0, self.visio_approval.get("name", ""))
            name_e.pack(side="left", fill="x", expand=True)
            r1 = ttk.Frame(form)
            r1.pack(fill="x", pady=3)
            ttk.Label(r1, text="REV letter (blank = latest):",
                      width=20).pack(side="left")
            rev_e = ttk.Entry(r1, width=10)
            rev_e.insert(0, self.visio_approval.get("rev", ""))
            rev_e.pack(side="left")

            def apply():
                rev = rev_e.get().strip()
                name = name_e.get().strip()
                if not name:
                    self.visio_approval = {}
                    self.log("Visio approval cleared.")
                else:
                    self.visio_approval = {"rev": rev, "name": name}
                    self.log(
                        "Staged Visio approval by " + name + " on "
                        + (f"REV {rev}." if rev
                           else "each file's latest revision.")
                    )
                win.destroy()

            btns = ttk.Frame(win)
            btns.pack(fill="x", padx=10, pady=10)
            self._rbtn(btns, "Save", apply, kind="green").pack(side="right")
            self._rbtn(btns, "Cancel", win.destroy).pack(side="right", padx=6)

        # -- Excel approval (EE / ME / Production sign off) ----------------
        def open_excel_approval(self):
            excel_files = [f for f in self.files if detect_format(f) == "xlsx"]
            if not excel_files:
                messagebox.showinfo(
                    "No Excel files", "Add at least one .xlsx file first."
                )
                return
            found = sorted({d for f in excel_files
                            for d in xlsx_approval_disciplines(f)},
                           key=APPROVAL_DISCIPLINES.index)
            win = tk.Toplevel(self.root)
            win.title("Approve Excel files")
            win.transient(self.root)
            win.grab_set()
            win.configure(bg=self.palette["bg"])
            avail = (f"Approval boxes found: {', '.join(found)}."
                     if found else "No EE/ME/Production approval boxes were "
                     "detected; nothing will change unless one is present.")
            ttk.Label(
                win, wraplength=560, justify="left",
                text="Sign off as approver: your name goes in the cell beside "
                "the discipline's label and today's date in the next cell "
                "(same format as the existing date), on every sheet except the "
                "Change Log, in every loaded Excel file.\n\n" + avail,
            ).pack(fill="x", padx=10, pady=8)

            form = ttk.Frame(win)
            form.pack(fill="x", padx=10, pady=4)
            r1 = ttk.Frame(form)
            r1.pack(fill="x", pady=3)
            ttk.Label(r1, text="Discipline:", width=14).pack(side="left")
            disc_var = tk.StringVar(
                value=self.excel_approval.get("discipline",
                                              APPROVAL_DISCIPLINES[0]))
            for d in APPROVAL_DISCIPLINES:
                ttk.Radiobutton(r1, text=d, value=d,
                                variable=disc_var).pack(side="left", padx=6)
            r2 = ttk.Frame(form)
            r2.pack(fill="x", pady=3)
            ttk.Label(r2, text="Approver name:", width=14).pack(side="left")
            name_e = ttk.Entry(r2, width=40)
            name_e.insert(0, self.excel_approval.get("name", ""))
            name_e.pack(side="left", fill="x", expand=True)

            def apply():
                name = name_e.get().strip()
                disc = disc_var.get()
                if not name:
                    self.excel_approval = {}
                    self.log("Excel approval cleared.")
                else:
                    self.excel_approval = {"discipline": disc, "name": name}
                    self.log(
                        f"Staged Excel approval: {disc} by {name} "
                        "(date = today)."
                    )
                win.destroy()

            btns = ttk.Frame(win)
            btns.pack(fill="x", padx=10, pady=10)
            self._rbtn(btns, "Save", apply, kind="green").pack(side="right")
            self._rbtn(btns, "Cancel", win.destroy).pack(side="right", padx=6)

        # -- output folder --------------------------------------------------
        def _choose_out_dir(self):
            d = filedialog.askdirectory(
                title="Choose a folder for all finished files",
                initialdir=self.out_dir or None,
            )
            if d:
                self.out_dir = d
                self.out_dir_lbl.configure(text=d)
                self.log(f"Output folder set to: {d}")

        def _clear_out_dir(self):
            self.out_dir = ""
            self.out_dir_lbl.configure(text="(same folder as each source file)")

        # -- reset everything ----------------------------------------------
        def reset_all(self):
            """Clear loaded files, all find/replace rules, every staged edit
            and the output folder, so a fresh file/batch starts from scratch."""
            if not messagebox.askyesno(
                "Reset everything?",
                "This clears the loaded files, all find/replace rules, every "
                "staged edit (BOM rows, Change Log, Author, revision entry, "
                "approvals) and the output folder.\n\nStart from scratch?",
            ):
                return
            # Files.
            self.files = []
            self.refresh_files_box()
            # Rules: drop them all, then re-add a single empty one.
            for r in list(self.rule_rows):
                r["frame"].destroy()
            self.rule_rows = []
            self._rbuttons = [b for b in self._rbuttons if b.winfo_exists()]
            self.add_pair()
            # Every staged action.
            self.bom_edits = {}
            self.visio_bom_edits = {}
            self.remove_parts = {}
            self.add_parts = {}
            self.changelog_entry = {}
            self.author_name = ""
            self.visio_rev_entry = {}
            self.visio_approval = {}
            self.excel_approval = {}
            self.bom_status.configure(text="")
            self.remove_status.configure(text="")
            self.add_status.configure(text="")
            self._clear_out_dir()
            self.log(
                "Reset: cleared files, rules, all staged edits and the output "
                "folder. Ready for a new file or batch."
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
                    and not self.bom_edits and not self.changelog_entry
                    and not self.author_name and not self.visio_rev_entry
                    and not self.visio_approval and not self.excel_approval
                    and not self.visio_bom_edits and not self.remove_parts
                    and not self.add_parts):
                messagebox.showwarning(
                    "Nothing to do",
                    "Enter a 'Find' value, tick 'Save copy as next revision', "
                    "or stage an Excel row / Change Log / Author / Visio "
                    "revision / approval edit.",
                )
                return

            if self.out_dir:
                try:
                    Path(self.out_dir).mkdir(parents=True, exist_ok=True)
                except OSError as exc:
                    messagebox.showerror(
                        "Output folder",
                        f"Couldn't use that output folder:\n{exc}",
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
                    self.author_name, self.summary_var.get(),
                    dict(self.visio_rev_entry), dict(self.visio_approval),
                    dict(self.excel_approval), self.out_dir,
                    dict(self.visio_bom_edits),
                    {f: dict(d) for f, d in self.remove_parts.items()},
                    {f: dict(d) for f, d in self.add_parts.items()},
                ),
                daemon=True,
            ).start()

        def _worker(self, files, pairs_by_file, case_sensitive, whole_word,
                    make_pdf, bump_rev, update_rev_text, bom_edits,
                    changelog_entry, author_name, make_summary,
                    visio_rev_entry, visio_approval, excel_approval, out_dir,
                    visio_bom_edits=None, remove_parts=None, add_parts=None):
            visio_bom_edits = visio_bom_edits or {}
            remove_parts = remove_parts or {}
            add_parts = add_parts or {}
            total_repl = 0
            done = 0
            errors = 0
            last_output = None
            summary_records = []
            try:
                for f in files:
                    src = Path(f)
                    # Key by the original string used to build the dict (path
                    # separators differ between tkinter and Path on Windows).
                    pairs = pairs_by_file.get(f, [])

                    # Excel cell edits for this file: BOM row edits, a Change
                    # Log append, the Author name/date, and an EE/ME/Production
                    # approval name+date (all by cell ref).
                    cell_edits = None
                    excel_approved = False
                    if ((bom_edits or changelog_entry or author_name
                         or excel_approval or remove_parts.get(f)
                         or add_parts.get(f))
                            and detect_format(f) == "xlsx"):
                        merged: dict = {}
                        try:
                            # BOM edits are per-file: {file: {sheet: {ref: val}}}
                            for sheet, cells in bom_edits.get(f, {}).items():
                                merged.setdefault(sheet, {}).update(cells)
                            if changelog_entry:
                                for part, d in build_changelog_cell_edits(
                                    f, changelog_entry
                                ).items():
                                    merged.setdefault(part, {}).update(d)
                            if author_name:
                                for part, d in build_author_edits(
                                    f, author_name
                                ).items():
                                    merged.setdefault(part, {}).update(d)
                            if excel_approval:
                                ae = build_approval_edits(
                                    f, excel_approval.get("discipline"),
                                    excel_approval.get("name"),
                                )
                                if ae:
                                    excel_approved = True
                                for part, d in ae.items():
                                    merged.setdefault(part, {}).update(d)
                            # Add / remove parts (shift + renumber) -> cells.
                            if remove_parts.get(f) or add_parts.get(f):
                                for sheet, d in build_excel_parts_ops(
                                    f, remove_parts.get(f, {}),
                                    add_parts.get(f, {}),
                                ).items():
                                    merged.setdefault(sheet, {}).update(d)
                        except Exception:  # noqa: BLE001
                            merged = {}
                        cell_edits = merged or None

                    # A Visio revision-table row / approval apply only to .vsdx.
                    is_vsdx = detect_format(f) == "vsdx"
                    staged_rev_entry = (visio_rev_entry
                                        if (visio_rev_entry and is_vsdx)
                                        else None)
                    # Visio approval: by default sign off THIS file's own
                    # latest revision (auto-detected); a typed REV overrides it.
                    appr = None
                    if visio_approval and is_vsdx and visio_approval.get("name"):
                        rev = (visio_approval.get("rev") or "").strip()
                        if not rev:
                            rev = vsdx_latest_revision_letter(f)
                        if rev:
                            appr = {"rev": rev,
                                    "name": visio_approval["name"]}
                    # Parts-table edits = the staged field edits from the
                    # editor PLUS Part Number replacements driven by the
                    # Find->Replace rules (the embedded tables aren't reached by
                    # the generic text engine).
                    vbom = None
                    if is_vsdx:
                        staged_vbom = (visio_bom_edits.get(f)
                                       if visio_bom_edits else None)
                        pn_repl = {}
                        if pairs:
                            try:
                                pn_repl = build_visio_pn_replacements(
                                    f, pairs, case_sensitive)
                            except Exception:  # noqa: BLE001
                                pn_repl = {}
                        parts_ops = {}
                        if remove_parts.get(f) or add_parts.get(f):
                            try:
                                parts_ops = build_visio_parts_ops(
                                    f, remove_parts.get(f, {}),
                                    add_parts.get(f, {}), case_sensitive)
                            except Exception:  # noqa: BLE001
                                parts_ops = {}
                        merged = _merge_bom_edit_dicts(
                            staged_vbom, pn_repl, parts_ops)
                        vbom = merged or None
                    has_edits = bool(cell_edits or staged_rev_entry or appr
                                     or vbom)
                    # An approval signs off an existing revision -- it must NOT
                    # bump the revision; the copy is named "*_approved_<date>".
                    is_approving = bool(appr) or excel_approved

                    # Decide the copy's name and whether to bump the revision.
                    revision = None
                    if is_approving:
                        stamp = datetime.date.today().strftime("%Y-%m-%d")
                        out_vsdx = src.with_name(
                            f"{src.stem}_approved_{stamp}{src.suffix}"
                        )
                    elif bump_rev:
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

                    # Redirect every output to the chosen folder, if any.
                    if out_dir:
                        out_vsdx = Path(out_dir) / out_vsdx.name

                    # Build this file's revision-table row: the REV column uses
                    # THIS file's own next revision letter (so each file in a
                    # batch gets its correct next letter); the other fields are
                    # shared across all files.
                    rev_entry = None
                    if staged_rev_entry:
                        rev_entry = {k: v for k, v in staged_rev_entry.items()
                                     if k != "Rev"}
                        next_letter = (revision[1] if revision
                                       else revision_output_path(src)[2])
                        if next_letter:
                            rev_entry["Rev"] = next_letter

                    try:
                        report = replace_text_in_file(
                            src, out_vsdx, pairs,
                            case_sensitive=case_sensitive,
                            whole_word=whole_word,
                            revision=revision,
                            update_drawing_rev=update_rev_text,
                            cell_edits=cell_edits,
                            rev_entry=rev_entry,
                            approval=appr,
                            bom_cell_edits=vbom,
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
                        if report.get("bom_cells"):
                            self.log(
                                f"    {report['bom_cells']} parts-table cell(s) "
                                "updated"
                            )
                        if revision:
                            rd = report["rev_drawing"]
                            sheets = report.get("rev_sheets", 0)
                            on = (f" on {sheets} sheet(s)"
                                  if sheets > 1 else "")
                            note = {
                                "updated":
                                    f"    REV box {revision[0]} -> "
                                    f"{revision[1]}{on}",
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
                        if rev_entry:
                            rl = rev_entry.get("Rev", "")
                            tp = report.get("rev_table_page") or "?"
                            added = (f"    revision table (on {tp}): row "
                                     + (f"REV {rl} " if rl else "")
                                     + "added")
                            npg = report.get("page_count", 0)
                            tnote = {
                                "filled": added,
                                "appended": added,
                                "not_found":
                                    f"    (NO revision table detected on any of "
                                    f"this file's {npg} page(s); left "
                                    "unchanged)",
                                "no_slot":
                                    f"    (revision table found on {tp} but no "
                                    "safe place to add a row; left unchanged)",
                                "na": None,
                            }.get(report.get("rev_table"))
                            if tnote:
                                self.log(tnote)
                            # Tell the user whether the displayed picture (the
                            # embedded OLE object) was updated, since that's the
                            # part that has to refresh for the row to be seen.
                            pic = {
                                "patched":
                                    "    revision picture updated (shows in "
                                    "Visio without opening the object)",
                                "patch_failed":
                                    "    (couldn't update the revision picture; "
                                    "open the table once in Visio to refresh "
                                    "it)",
                                "unsupported":
                                    "    (the revision table's picture isn't an "
                                    "EMF, so it can't be auto-updated; open the "
                                    "table once in Visio to refresh it)",
                                "no_picture":
                                    "    (no cached picture found for the "
                                    "revision table; open it once in Visio to "
                                    "refresh it)",
                            }.get(report.get("rev_picture"))
                            if pic:
                                self.log(pic)
                        if appr:
                            anote = {
                                "approved":
                                    f"    revision {appr['rev']}: approved by "
                                    f"{appr['name']}",
                                "row_not_found":
                                    f"    (no revision row '{appr['rev']}' "
                                    "found; approval skipped)",
                                "no_column":
                                    "    (revision table has no Approved By "
                                    "column; approval skipped)",
                                "not_found":
                                    "    (revision table not found; approval "
                                    "skipped)",
                                "na": None,
                            }.get(report.get("approval"))
                            if anote:
                                self.log(anote)
                        last_output = out_vsdx

                        if make_summary:
                            summary_records.append({
                                "input": src.name, "output": out_vsdx.name,
                                "revision": revision,
                                "changes": diff_files(src, out_vsdx),
                            })

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

                summary_path = None
                if make_summary and summary_records:
                    try:
                        folder = (Path(out_dir) if out_dir
                                  else Path(str(last_output)).parent
                                  if last_output else Path(files[0]).parent)
                        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
                        summary_path = generate_change_summary(
                            summary_records,
                            folder / f"Change_Summary_{stamp}.html",
                        )
                        self.log(f"Change summary -> {summary_path.name}")
                    except Exception as exc:  # noqa: BLE001
                        self.log(f"(could not write change summary: {exc})")

                self.root.after(
                    0, lambda: messagebox.showinfo("Finished", summary)
                )
                if summary_path is not None:
                    self._open_path(str(summary_path))
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
            self._open_path(str(Path(path).parent))

        def show_help(self):
            """Open the in-app job aid: a scrollable How-to window."""
            c = self.palette
            win = tk.Toplevel(self.root)
            win.title(HELP_TITLE)
            win.geometry("780x620")
            win.transient(self.root)
            win.configure(bg=c["bg"])

            bottom = ttk.Frame(win)
            bottom.pack(side="bottom", fill="x", padx=12, pady=10)

            txt = scrolledtext.ScrolledText(
                win, wrap="word", relief="flat", borderwidth=0,
                bg=c["card"], fg=c["text"], padx=18, pady=14,
                insertbackground=c["text"], font=("Segoe UI", 10),
                cursor="arrow", spacing1=2, spacing3=4,
            )
            txt.pack(side="top", fill="both", expand=True, padx=12, pady=(12, 0))

            txt.tag_configure(
                "h1", font=("Segoe UI", 16, "bold"), foreground=c["accent"],
                spacing1=4, spacing3=10,
            )
            txt.tag_configure(
                "h2", font=("Segoe UI", 12, "bold"), foreground=c["text"],
                spacing1=14, spacing3=6,
            )
            txt.tag_configure("body", font=("Segoe UI", 10), lmargin1=4,
                              lmargin2=4, spacing3=4)
            txt.tag_configure("bullet", font=("Segoe UI", 10), lmargin1=18,
                              lmargin2=32, spacing3=3)
            txt.tag_configure("b", font=("Segoe UI", 10, "bold"))
            txt.tag_configure("bb", font=("Segoe UI", 10, "bold"),
                              foreground=c["accent"])

            def emit(text, base):
                # Render inline **bold** spans within a line.
                bold = base == "bullet" and "bb" or "b"
                parts = text.split("**")
                for i, seg in enumerate(parts):
                    if not seg:
                        continue
                    tags = (base, bold) if i % 2 else (base,)
                    txt.insert("end", seg, tags)
                txt.insert("end", "\n")

            txt.insert("end", HELP_TITLE + "\n", ("h1",))
            for heading, lines in HELP_SECTIONS:
                txt.insert("end", heading + "\n", ("h2",))
                for ln in lines:
                    if ln.startswith("* "):
                        txt.insert("end", "•  ", ("bullet",))
                        emit(ln[2:], "bullet")
                    else:
                        emit(ln, "body")
            txt.configure(state="disabled")

            def open_printable():
                import tempfile
                tmp = Path(tempfile.gettempdir()) / "Visio_Tool_How_To.html"
                try:
                    help_to_html(tmp)
                    self._open_path(str(tmp))
                except Exception as exc:  # noqa: BLE001
                    messagebox.showerror(
                        "Help", f"Could not open printable version:\n{exc}",
                        parent=win,
                    )

            self._rbtn(
                bottom, "🖨  Open printable / PDF version", open_printable,
                kind="normal", radius=12,
            ).pack(side="left")
            self._rbtn(
                bottom, "Close", win.destroy, kind="accent", radius=12,
            ).pack(side="right")

        def _open_path(self, path: str):
            """Open a file or folder with the OS default app, best-effort."""
            try:
                if sys.platform.startswith("win"):
                    os.startfile(path)  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", path])
                else:
                    subprocess.Popen(["xdg-open", path])
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
                "  python drawing_bom_studio.py drawing.vsdx "
                '--find "Old" --replace "New" --pdf',
                file=sys.stderr,
            )
            return 1
    return _run_cli(argv)


if __name__ == "__main__":
    raise SystemExit(main())
