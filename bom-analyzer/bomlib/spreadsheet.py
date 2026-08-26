"""Read the file formats BOMs actually arrive in.

CSV/TSV exported from an EDA tool, or the .xlsx a purchasing department sends.
The stdlib does the container work: `csv` for delimited text, `zipfile` plus
`xml.etree` for the OOXML package.
"""

import csv
import io
import re
import zipfile
import xml.etree.ElementTree as ET

# ── Cleaning cells ─────────────────────────────────────────────────────────

# Characters that take up no space on screen and so cannot be seen or deleted
# in a spreadsheet, but stop a part number matching anything. They arrive by
# the usual routes: a part number copied out of a PDF datasheet or a web page,
# an ERP export, a file that has been through a translation tool.
#
# Ordinary spaces are not in here — str.strip() already removes those, along
# with tabs, non-breaking spaces and the rest of the Unicode space characters.
# These are the ones it leaves behind, because Unicode does not classify them
# as whitespace at all.
_INVISIBLE = re.compile(
    '['
    '\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f-\u009f'  # control codes
    '\u00ad'                                                     # soft hyphen
    '\u200b-\u200f'                                             # zero-width, bidi marks
    '\u2028\u2029'                                               # line, paragraph separators
    '\u202a-\u202e'                                             # bidi embedding
    '\u2060-\u2064'                                             # word joiner, invisible operators
    '\ufeff'                                                     # byte-order mark
    '\ufff9-\ufffb'                                             # interlinear annotation
    ']'
)


def clean_cell(value):
    """What a spreadsheet cell actually meant, with the invisible parts gone.

    Leading and trailing space of every Unicode kind is removed, invisible
    characters are dropped wherever they sit, and a run of spaces inside the
    value collapses to one. Interior single spaces are kept: a few real part
    numbers contain one, and guessing which is which would break more than it
    fixed.
    """
    if value is None:
        return ''
    text = _INVISIBLE.sub('', str(value))
    # Runs of horizontal space collapse; a newline inside a quoted CSV field is
    # part of the value and stays, and str.strip() then takes whitespace of
    # every Unicode kind off both ends, \xa0 included.
    return re.sub(r'[^\S\n\r]+', ' ', text).strip()


# ── Delimited text ─────────────────────────────────────────────────────────


def count_outside_quotes(line, delimiter):
    count = 0
    in_quotes = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if in_quotes and i + 1 < len(line) and line[i + 1] == '"':
                i += 1
            else:
                in_quotes = not in_quotes
        elif ch == delimiter and not in_quotes:
            count += 1
        i += 1
    return count


def detect_delimiter(text):
    """Favour the delimiter that appears often AND consistently per line.

    csv.Sniffer guesses badly on BOMs, where description fields are full of
    commas and semicolons regardless of the real separator.
    """
    sample = text[:8192].splitlines()[:20]
    best = ','
    best_score = float('-inf')
    for delimiter in (',', '\t', ';', '|'):
        counts = [count_outside_quotes(line, delimiter) for line in sample if line.strip()]
        if not counts:
            continue
        total = sum(counts)
        if total == 0:
            continue
        mean = total / len(counts)
        variance = sum((c - mean) ** 2 for c in counts) / len(counts)
        score = mean - variance
        if score > best_score:
            best_score = score
            best = delimiter
    return best


def parse_delimited(text, delimiter=None):
    if text and text[0] == '﻿':
        text = text[1:]
    sep = delimiter or detect_delimiter(text)
    reader = csv.reader(io.StringIO(text, newline=''), delimiter=sep, quotechar='"')
    return [[clean_cell(cell) for cell in row] for row in reader]


# ── XLSX ───────────────────────────────────────────────────────────────────

MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _tag(element):
    """Local tag name, ignoring whichever namespace the writer used."""
    return element.tag.rsplit('}', 1)[-1]


def _shared_strings(archive):
    try:
        xml = archive.read('xl/sharedStrings.xml')
    except KeyError:
        return []
    root = ET.fromstring(xml)
    strings = []
    for si in root:
        # A shared string can be split across several runs (<r><t>…</t></r>).
        strings.append(''.join(node.text or '' for node in si.iter() if _tag(node) == 't'))
    return strings


def _first_sheet_path(archive, sheet=0):
    """Path to a worksheet inside the package, by position or by tab name.

    Defaults to the first sheet, which is what reading someone's uploaded BOM
    wants; the name form is what lets a caller reach a later tab of a workbook
    this project wrote.
    """
    names = set(archive.namelist())
    try:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        rels = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
    except (KeyError, ET.ParseError):
        workbook = rels = None

    if workbook is not None and rels is not None:
        entries = [n for n in workbook.iter() if _tag(n) == 'sheet']
        if isinstance(sheet, int):
            chosen = entries[sheet] if -len(entries) <= sheet < len(entries) else None
        else:
            chosen = next((n for n in entries if n.get('name') == sheet), None)
            if chosen is None:
                raise ValueError('No sheet named %r in this workbook' % sheet)
        if chosen is not None:
            rel_id = chosen.get('{%s}id' % REL_NS)
            if rel_id:
                for relationship in rels:
                    if relationship.get('Id') == rel_id:
                        target = relationship.get('Target', '').lstrip('/')
                        target = re.sub(r'^xl/', '', target)
                        path = 'xl/' + target
                        if path in names:
                            return path

    for name in archive.namelist():
        if re.fullmatch(r'xl/worksheets/sheet[^/]*\.xml', name):
            return name
    return None


def _column_index(ref):
    match = re.match(r'^([A-Za-z]+)', ref or '')
    if not match:
        return None
    index = 0
    for char in match.group(1).upper():
        index = index * 26 + (ord(char) - 64)
    return index - 1


def parse_xlsx(data, sheet=0):
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        sheet_path = _first_sheet_path(archive, sheet)
        if not sheet_path:
            raise ValueError('No worksheet found inside the .xlsx file')
        shared = _shared_strings(archive)
        root = ET.fromstring(archive.read(sheet_path))

    rows = []
    for row_node in (n for n in root.iter() if _tag(n) == 'row'):
        cells = []
        for cell in (c for c in row_node if _tag(c) == 'c'):
            cell_type = cell.get('t', 'n')
            index = _column_index(cell.get('r', '')) if cell.get('r') else len(cells)

            if cell_type == 'inlineStr':
                value = ''.join(n.text or '' for n in cell.iter() if _tag(n) == 't')
            else:
                v_node = next((n for n in cell if _tag(n) == 'v'), None)
                raw = v_node.text or '' if v_node is not None else ''
                if cell_type == 's':
                    try:
                        value = shared[int(raw)]
                    except (ValueError, IndexError):
                        value = ''
                elif cell_type == 'b':
                    value = 'TRUE' if raw == '1' else 'FALSE'
                else:
                    value = raw

            if index is None or index < 0:
                continue
            while len(cells) < index:
                cells.append('')
            if index < len(cells):
                cells[index] = clean_cell(value)
            else:
                cells.append(clean_cell(value))
        rows.append(cells)
    return rows


# ── Header detection and column mapping ────────────────────────────────────

# Each field lists header aliases in descending confidence. Matching is done on
# a squashed form so "Mfr. Part #" and "mfr_part_no" both land on `mpn`.
FIELD_ALIASES = {
    'mpn': [
        'manufacturerpartnumber', 'mfrpartnumber', 'mfgpartnumber', 'manufacturerpartno',
        'mfrpartno', 'mfgpartno', 'manufacturerpart', 'mfrpart', 'mfgpart', 'mpn',
        'manufacturernumber', 'partnumber', 'partno', 'partnum', 'part', 'componentpartnumber',
        'vendorpartnumber', 'orderingcode', 'ordernumber',
    ],
    'quantity': [
        'quantity', 'qty', 'qtyper', 'quantityper', 'qtyperboard', 'quantityperboard',
        'count', 'amount', 'qtyrequired',
    ],
    'reference': [
        'referencedesignator', 'referencedesignators', 'reference', 'references', 'refdes',
        'refdesignator', 'designator', 'designators', 'ref',
    ],
    'manufacturer': [
        'manufacturer', 'manufacturername', 'mfr', 'mfg', 'mfrname', 'mfgname', 'brand',
        'vendor', 'supplier', 'make',
    ],
    'description': [
        'description', 'desc', 'partdescription', 'componentdescription', 'comment', 'value',
        'name', 'partname',
    ],
    'footprint': ['footprint', 'package', 'packagetype', 'casecode', 'case'],
}

FIELD_ORDER = ['mpn', 'quantity', 'reference', 'manufacturer', 'description', 'footprint']


def squash(text):
    return re.sub(r'[^a-z0-9]', '', str(text or '').lower())


def score_header(header, field):
    key = squash(header)
    if not key:
        return 0
    for i, alias in enumerate(FIELD_ALIASES.get(field, [])):
        # Later aliases are weaker matches, so their score decays.
        weight = 100 - i * 2
        if key == alias:
            return weight
        if key.startswith(alias) or key.endswith(alias):
            return int(round(weight * 0.8))
        if alias in key:
            return int(round(weight * 0.6))
    return 0


def find_header_row(rows):
    """BOM exports often carry a title block or blank rows above the real
    header, so the header row is found by scoring rather than assumed."""
    limit = min(len(rows), 25)
    best_index = 0
    best_score = float('-inf')

    for i in range(limit):
        row = rows[i] or []
        filled = sum(1 for cell in row if str(cell or '').strip())
        if filled < 2:
            continue
        score = 0
        for field in FIELD_ORDER:
            score += max([score_header(cell, field) for cell in row] or [0])
        # A header row is text, not data; numeric cells argue against it.
        numeric = sum(1 for cell in row if re.fullmatch(r'-?\d+(\.\d+)?', str(cell or '').strip()))
        score -= numeric * 15
        score += filled
        if score > best_score:
            best_score = score
            best_index = i
    return best_index if best_score > 40 else 0


def map_columns(headers):
    mapping = {}
    taken = set()

    # Resolve the strongest header/field pairs first so a single "Part Number"
    # column is not claimed by a weaker field.
    pairs = []
    for field in FIELD_ORDER:
        for index, header in enumerate(headers):
            score = score_header(header, field)
            if score > 0:
                pairs.append((score, field, index))
    pairs.sort(key=lambda p: -p[0])

    for _, field, index in pairs:
        if field in mapping or index in taken:
            continue
        mapping[field] = index
        taken.add(index)
    return mapping


def line_from_row(row, mapping, row_index):
    def cell(field):
        index = mapping.get(field)
        if index is None or index >= len(row):
            return ''
        # The readers clean the grid already; this covers a grid handed back by
        # a client when a column mapping is corrected.
        return clean_cell(row[index])

    quantity_raw = cell('quantity')
    digits = re.sub(r'[^0-9-]', '', quantity_raw)
    try:
        quantity = int(digits)
    except ValueError:
        quantity = 0

    return {
        'row': row_index + 1,
        'mpn': cell('mpn'),
        'quantity': quantity if quantity > 0 else 1,
        'quantityRaw': quantity_raw or None,
        'reference': cell('reference') or None,
        'manufacturer': cell('manufacturer') or None,
        'description': cell('description') or None,
        'footprint': cell('footprint') or None,
    }


def is_zip(data):
    return len(data) > 4 and data[:4] == b'PK\x03\x04'


def parse_workbook(data, filename=None):
    name = str(filename or '').lower()
    if isinstance(data, bytes) and (name.endswith(('.xlsx', '.xlsm')) or is_zip(data)):
        return parse_xlsx(data)
    text = data.decode('utf-8', errors='replace') if isinstance(data, bytes) else str(data)
    delimiter = '\t' if name.endswith('.tsv') else None
    return parse_delimited(text, delimiter)


def extract_bom(rows):
    """Turn a raw grid into BOM lines plus the mapping decisions, which the UI
    shows so the user can correct a bad guess instead of re-exporting."""
    grid = [row for row in (rows or []) if isinstance(row, list)]
    if not grid:
        return {'headerRow': 0, 'headers': [], 'mapping': {}, 'lines': [], 'skipped': 0, 'totalRows': 0}

    header_row = find_header_row(grid)
    headers = [str(cell or '').strip() for cell in (grid[header_row] if header_row < len(grid) else [])]
    mapping = map_columns(headers)
    lines = []
    skipped = 0

    for i in range(header_row + 1, len(grid)):
        row = grid[i]
        if not row or all(not str(cell or '').strip() for cell in row):
            continue
        line = line_from_row(row, mapping, i)
        if not line['mpn']:
            skipped += 1
            continue
        lines.append(line)

    return {
        'headerRow': header_row,
        'headers': headers,
        'mapping': mapping,
        'lines': lines,
        'skipped': skipped,
        'totalRows': len(grid),
    }
