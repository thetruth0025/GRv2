"""Build a DMSMS case form from an analyzed BOM.

DMSMS — Diminishing Manufacturing Sources and Material Shortages — is the
process for handling parts whose supply is ending. The case form is what an
obsolescence analyst opens per program: the parts at risk, what the suppliers
currently say about them, and empty columns for the decisions that are the
analyst's to make.

The layout follows the case data elements described in the DoD DMSMS guidebook
(SD-22): identification, status and its provenance, exposure, and resolution.
Everything this app can answer is filled in; everything it cannot is left blank
rather than guessed, because a case form is evidence for a purchasing decision
and an invented CAGE code or last-time-buy date is worse than an empty cell.
"""

import datetime

from .normalize import Lifecycle
from .xlsx_writer import (
    Cell,
    STYLE_BAD,
    STYLE_BOLD,
    STYLE_DEFAULT,
    STYLE_HEADER,
    STYLE_INT,
    STYLE_LABEL,
    STYLE_MONEY,
    STYLE_MONEY_FINE,
    STYLE_MUTED,
    STYLE_SECTION,
    STYLE_SUBTITLE,
    STYLE_TITLE,
    STYLE_WARN,
    column_letter,
    write_xlsx,
)

# Every status that belongs on a DMSMS case form. The first three are the ones
# that mean the part is gone or going; the last two are still buyable but on
# the way out, and a last-time-buy window is often what actually bites.
DMSMS_STATUSES = (
    Lifecycle.OBSOLETE,
    Lifecycle.DISCONTINUED,
    Lifecycle.END_OF_LIFE,
    Lifecycle.LAST_TIME_BUY,
    Lifecycle.NRND,
)

# Ticked by default when the form is opened: no longer available at all.
DEFAULT_SELECTED_STATUSES = (
    Lifecycle.OBSOLETE,
    Lifecycle.DISCONTINUED,
    Lifecycle.END_OF_LIFE,
)

CASE_COLUMNS = [
    'Item',
    'Manufacturer Part Number',
    'Manufacturer',
    'CAGE Code',
    'Description',
    'Reference Designators',
    'Next Higher Assembly',
    'Qty per Assembly',
    'Lifecycle Status',
    'Status Source',
    'Date Obtained',
    'Distributor Stock',
    'Preferred Source',
    'Unit Price',
    'Extended Price',
    'Suggested Replacement',
    'Suggested Risk',
    'Lifetime Buy Qty',
    'Last Time Buy Date',
    'Resolution Option',
    'Disposition / Remarks',
]

CASE_WIDTHS = [
    6, 26, 20, 11, 34, 20, 20, 14, 22, 18, 13, 15, 16, 12, 14, 24, 14, 15, 18, 22, 34,
]

# Columns the analyst fills in. Named here so the header can mark them and
# nothing downstream mistakes an empty cell for missing data.
ANALYST_COLUMNS = (
    'CAGE Code', 'Lifetime Buy Qty', 'Last Time Buy Date',
    'Resolution Option', 'Disposition / Remarks',
)

# The choices SD-22 lays out, in the order it escalates them. Offered as a note
# on the form rather than enforced, because picking one is the analyst's call.
RESOLUTION_OPTIONS = [
    'Existing stock',
    'Reclamation',
    'Alternate or substitute part',
    'Aftermarket source',
    'Lifetime / bridge buy',
    'Emulation',
    'Redesign — next higher assembly',
]

FORM_WIDTH = len(CASE_COLUMNS)


def qualifies(status):
    """True when a lifecycle status belongs on a DMSMS form at all."""
    return status in DMSMS_STATUSES


def default_selected(status):
    return status in DEFAULT_SELECTED_STATUSES


def candidate_rows(result):
    """Every analyzed line whose worst status across suppliers is at risk."""
    return [row for row in result.get('rows') or []
            if qualifies((row.get('comparison') or {}).get('lifecycle'))]


def status_sources(row):
    """Which suppliers reported the status the form is citing.

    The comparison keeps the worst status across suppliers, so a part DigiKey
    calls Obsolete is never shown as Active because Mouser has not caught up.
    Naming the supplier is what makes the entry checkable.
    """
    worst = (row.get('comparison') or {}).get('lifecycle')
    names = []
    for offer in (row.get('offers') or {}).values():
        if not offer or not offer.get('found'):
            continue
        if offer.get('lifecycle') == worst and offer.get('supplier') not in names:
            names.append(offer.get('supplier'))
    return names


def total_stock(row):
    """Distributor stock across every supplier that carries the part."""
    total = None
    for offer in (row.get('offers') or {}).values():
        if not offer or not offer.get('found'):
            continue
        held = offer.get('totalStock')
        if held is None:
            held = offer.get('stock')
        if isinstance(held, (int, float)):
            total = (total or 0) + held
    return total


def preferred_offer(row):
    """The offer the comparison recommends, which is what the form prices."""
    name = (row.get('comparison') or {}).get('recommendedSupplier')
    for offer in (row.get('offers') or {}).values():
        if offer and offer.get('found') and offer.get('supplier') == name:
            return offer
    return None


def suggested_replacement(row):
    """What to use instead, if anything has said so.

    An alternative found through Nexar wins over a supplier's own suggestion:
    it was looked up deliberately for this part, whereas the supplier field is
    whatever happened to be attached to the catalogue entry.
    """
    named = row.get('suggestedReplacement')
    if named:
        return str(named)

    for offer in (row.get('offers') or {}).values():
        if offer and offer.get('suggestedReplacement'):
            return offer['suggestedReplacement']
    return None


def suggest_risk(row):
    """A starting point, not a determination — hence "suggested" on the form.

    Built only from what the suppliers actually reported: how dead the part is,
    and whether anyone still has enough of it to cover the build.
    """
    comparison = row.get('comparison') or {}
    status = comparison.get('lifecycle')
    needed = row.get('quantity') or 0
    stock = total_stock(row)
    covered = stock is not None and needed and stock >= needed

    if status in (Lifecycle.OBSOLETE, Lifecycle.DISCONTINUED, Lifecycle.END_OF_LIFE):
        return 'High' if not covered else 'Medium'
    if status == Lifecycle.LAST_TIME_BUY:
        return 'Medium' if covered else 'High'
    if status == Lifecycle.NRND:
        return 'Low' if covered else 'Medium'
    return 'Low'


RISK_STYLE = {'High': STYLE_BAD, 'Medium': STYLE_WARN, 'Low': STYLE_DEFAULT}


def build_case_rows(rows, styled=False, obtained=None):
    """The case table: one row per part, analyst columns left empty."""
    def cell(value, style=STYLE_DEFAULT):
        return Cell(value, style) if styled else value

    obtained = obtained or datetime.date.today().isoformat()
    table = [[cell(name, STYLE_HEADER) for name in CASE_COLUMNS]]

    for index, row in enumerate(rows):
        comparison = row.get('comparison') or {}
        offer = preferred_offer(row)
        status = comparison.get('lifecycle')
        risk = suggest_risk(row)

        table.append([
            cell(index + 1, STYLE_INT),
            cell(row.get('mpn')),
            cell(row.get('manufacturer') or (offer or {}).get('manufacturer')),
            cell(None),                                   # CAGE Code — analyst
            cell(row.get('description')),
            cell(row.get('reference')),
            cell(row.get('assembly')),
            cell(row.get('quantity'), STYLE_INT),
            cell(status, STYLE_BAD if comparison.get('lifecycleSeverity') == 'bad' else STYLE_WARN),
            cell(' / '.join(status_sources(row)) or None),
            cell(obtained),
            cell(total_stock(row), STYLE_INT),
            cell(comparison.get('recommendedSupplier')),
            cell((offer or {}).get('unitPrice'), STYLE_MONEY_FINE),
            cell((offer or {}).get('extendedPrice'), STYLE_MONEY),
            cell(suggested_replacement(row)),
            cell(risk, RISK_STYLE.get(risk, STYLE_DEFAULT)),
            cell(None),                                   # Lifetime Buy Qty — analyst
            cell(None),                                   # Last Time Buy Date — analyst
            cell(None),                                   # Resolution Option — analyst
            cell(None),                                   # Disposition — analyst
        ])
    return table


def _pad(row, width=FORM_WIDTH):
    return row + [Cell('')] * max(0, width - len(row))


def build_form_rows(rows, meta=None):
    """Title block, the case table, and the notes that make it auditable."""
    meta = meta or {}
    obtained = meta.get('obtained') or datetime.date.today().isoformat()
    program = meta.get('program') or 'Unnamed program'

    counts = {}
    for row in rows:
        status = (row.get('comparison') or {}).get('lifecycle')
        counts[status] = counts.get(status, 0) + 1

    high = sum(1 for row in rows if suggest_risk(row) == 'High')

    table = [
        _pad([Cell('DMSMS Case Form', STYLE_TITLE)]),
        _pad([Cell('%s — %d part%s at risk, %d suggested high risk'
                   % (program, len(rows), '' if len(rows) == 1 else 's', high),
                   STYLE_SUBTITLE)]),
        _pad([]),
        _pad([Cell('Case details', STYLE_SECTION)]),
    ]

    details = [
        ('Program / platform', meta.get('program')),
        ('DMSMS case number', meta.get('caseNumber')),
        ('Prepared by', meta.get('preparedBy')),
        ('Organization', meta.get('organization')),
        ('Contract number', meta.get('contract')),
        ('CAGE code', meta.get('cage')),
        ('Date prepared', meta.get('date') or obtained),
        ('Status obtained', obtained),
        ('Source of status', meta.get('sources') or 'DigiKey, Mouser, TrustedParts APIs'),
        ('Bills of materials', meta.get('scope')),
    ]
    for label, value in details:
        table.append(_pad([Cell(label, STYLE_LABEL), Cell(value if value else None)]))

    if meta.get('notes'):
        table.append(_pad([Cell('Notes', STYLE_LABEL), Cell(meta['notes'])]))

    table.append(_pad([]))
    table.append(_pad([Cell('Parts at risk', STYLE_SECTION)]))
    for status in DMSMS_STATUSES:
        if counts.get(status):
            table.append(_pad([Cell(status, STYLE_LABEL), Cell(counts[status], STYLE_INT)]))

    table.append(_pad([]))
    table.append(_pad([Cell('Case records', STYLE_SECTION)]))
    header_index = len(table)
    table.extend(build_case_rows(rows, styled=True, obtained=obtained))

    table.append(_pad([]))
    table.append(_pad([Cell(
        'Blank columns are for the analyst: %s.' % ', '.join(ANALYST_COLUMNS), STYLE_MUTED)]))
    table.append(_pad([Cell(
        'Resolution options: %s.' % '; '.join(RESOLUTION_OPTIONS), STYLE_MUTED)]))
    table.append(_pad([Cell(
        'Suggested Risk is derived from the reported status and whether distributor stock '
        'covers the quantity per assembly. It is a starting point for the case, not a '
        'determination.', STYLE_MUTED)]))
    table.append(_pad([Cell(
        'Lifecycle status is whichever supplier reports the worst standing for the part, '
        'named in Status Source. Confirm against the manufacturer before acting.', STYLE_MUTED)]))

    return table, header_index


def write_form(target, rows, meta=None):
    """Write the case form workbook. Accepts a path or a file-like object."""
    table, header_index = build_form_rows(rows, meta)

    sheets = [{
        'name': 'DMSMS Case Form',
        'rows': table,
        'widths': CASE_WIDTHS,
        # The title block is not a header row, so neither a frozen pane nor a
        # filter on row 1 would mean anything.
        'freeze': 0,
        'autofilter': False,
        'merges': ['A1:%s1' % column_letter(min(FORM_WIDTH, 8) - 1),
                   'A2:%s2' % column_letter(min(FORM_WIDTH, 8) - 1)],
        'heights': {0: 26, 1: 18},
    }]
    return write_xlsx(target, sheets, freeze_rows=0, autofilter=False), header_index
